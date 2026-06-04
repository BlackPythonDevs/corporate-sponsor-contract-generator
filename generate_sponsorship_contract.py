#!/usr/bin/env -S uv run --script
# /// script
# requires-python = ">=3.9"
# dependencies = [
#     "python-docx>=1.1.0",
#     "click>=8.1.0",
#     "PyYAML>=6.0",
# ]
# ///
"""
Corporate Sponsorship Contract Generator
=========================================

Builds a GNOME Foundation / Black Python Devs (BPD) corporate sponsorship
agreement and writes it to a formatted .docx file.

GNOME Foundation (the 501(c)(3) fiscal sponsor) and Black Python Devs (the
fiscal-sponsored community org) are the default parties, but every field —
including the nonprofit and community blocks — can be supplied through a YAML
config file. Anything not provided via YAML or command-line flags is prompted
for interactively (unless --non-interactive is set).

Configurable fields:

  * Sponsor legal name            (e.g. "JetBrains s.r.o.")
  * Sponsor entity description    (e.g. "a private limited company")
  * Sponsor principal place of business (short, for the intro paragraph)
  * Payment amount                (default $5,000)
  * Payment due window (days)     (default 30)
  * Effective date                (default: today)
  * Notice block: POC name, title, email, and mailing address
  * Output filename
  * nonprofit / community party blocks (YAML only)

Usage:
    # Fully interactive (original behavior):
    uv run generate_sponsorship_contract.py

    # Everything from a YAML file:
    uv run generate_sponsorship_contract.py --config sponsor.yaml

    # YAML plus a CLI override, no prompting:
    uv run generate_sponsorship_contract.py -c sponsor.yaml --amount 7500 --non-interactive

    # Write a starter config you can fill in:
    uv run generate_sponsorship_contract.py --dump-config > sponsor.yaml

uv reads the inline script metadata above and installs the dependencies into an
ephemeral, isolated environment automatically — no manual pip step or venv.
"""

import io
import sys
import datetime
import urllib.request
from pathlib import Path

try:
    import click
    import yaml
    from docx import Document
    from docx.shared import Pt, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
except ImportError:
    sys.exit(
        "Missing dependency. Install with:\n"
        "    pip install python-docx click PyYAML\n"
        "(or run via uv, which installs them automatically)"
    )


# --------------------------------------------------------------------------- #
# Default parties (the nonprofit fiscal sponsor and its sponsored community org)
# Each can be overridden wholesale via the YAML config's `nonprofit:` /
# `community:` keys.
# --------------------------------------------------------------------------- #
DEFAULT_NONPROFIT = {
    "name": "GNOME Foundation",
    "short": "GNOME",
    "place": "Orinda, California",
    "notice": [
        "#117 21c Orinda Way",
        "Orinda",
        "CA 94563",
        "USA",
    ],
    "email": "president@gnome.org",
    "attn": "Allan Day/President",
}
DEFAULT_COMMUNITY = {
    "name": "Black Python Devs",
    "short": "BPD",
}

# Generated contracts land here by default (gitignored). A filename with its
# own directory component, or an absolute path, bypasses this.
OUTPUT_DIR = "output"

# Width logos are scaled to in Exhibit B (inches). python-docx keeps aspect ratio.
LOGO_WIDTH_IN = 2.0

# Sponsor-field defaults (None == required, must be supplied or prompted for).
FIELD_DEFAULTS = {
    "sponsor_name": None,
    "sponsor_entity": "a private limited company",
    "sponsor_place": None,
    "amount": 5000,
    "payment_days": "30",
    "effective": None,  # None -> today
    "poc_name": None,
    "poc_title": "",
    "poc_email": None,
    "notice_address": None,
    "logos": [],  # list of logo images embedded in Exhibit B (paths or URLs)
    "out_name": None,  # None -> derived from sponsor_name
}


# --------------------------------------------------------------------------- #
# Interactive prompts
# --------------------------------------------------------------------------- #
def ask(prompt, default=None):
    """Prompt for a single line, returning the default on empty input."""
    suffix = f" [{default}]" if default not in (None, "") else ""
    while True:
        value = input(f"{prompt}{suffix}: ").strip()
        if value:
            return value
        if default is not None:
            return default
        print("  (required) please enter a value.")


def ask_multiline(prompt):
    """Prompt for several lines; blank line ends input."""
    print(f"{prompt} (enter one line at a time; blank line to finish)")
    lines = []
    while True:
        line = input("  > ").rstrip()
        if not line:
            if lines:
                break
            print("  (required) please enter at least one line.")
            continue
        lines.append(line)
    return lines


# --------------------------------------------------------------------------- #
# Normalization helpers (applied to YAML, CLI, and interactive values alike)
# --------------------------------------------------------------------------- #
def normalize_amount(raw):
    """Coerce a currency amount to a formatted string like '5,000'."""
    cleaned = str(raw).replace("$", "").replace(",", "").strip()
    value = int(round(float(cleaned)))
    return f"{value:,}"


def normalize_date(raw):
    """Coerce an effective date to 'Month D, YYYY'; pass through if unparseable."""
    if isinstance(raw, (datetime.date, datetime.datetime)):
        d = raw if isinstance(raw, datetime.date) else raw.date()
        return d.strftime("%B %d, %Y").replace(" 0", " ")
    text = str(raw).strip()
    for fmt in ("%Y-%m-%d", "%B %d, %Y", "%B %d %Y", "%m/%d/%Y"):
        try:
            d = datetime.datetime.strptime(text, fmt).date()
            return d.strftime("%B %d, %Y").replace(" 0", " ")
        except ValueError:
            continue
    return text  # Couldn't parse — trust the free-form text.


def today_str():
    today = datetime.date.today()
    return (
        today.strftime("%B %-d, %Y")
        if sys.platform != "win32"
        else today.strftime("%B %d, %Y")
    )


def normalize_logos(raw, base_dir=None):
    """
    Coerce the `logos` config into a list of {label, source, width} dicts.

    Each input entry may be:
      * a plain string  -> a local path or an http(s) URL
      * a mapping with `path`/`url`/`source` (one required) plus optional
        `label` (caption shown above the image) and `width` (inches).

    Local relative paths are resolved against `base_dir` (the config file's
    directory) first, then the current working directory.
    """
    if raw in (None, ""):
        return []
    if isinstance(raw, (str, dict)):
        raw = [raw]

    resolved = []
    for entry in raw:
        if isinstance(entry, str):
            entry = {"source": entry}
        elif isinstance(entry, dict):
            entry = dict(entry)
            entry["source"] = entry.get("source") or entry.get("path") or entry.get("url")
        else:
            raise click.ClickException(f"Invalid logo entry: {entry!r}")

        source = entry.get("source")
        if not source:
            raise click.ClickException(f"Logo entry missing a path/url: {entry!r}")

        is_url = str(source).lower().startswith(("http://", "https://"))
        if not is_url:
            p = Path(source).expanduser()
            candidates = [p]
            if not p.is_absolute():
                if base_dir:
                    candidates.append(Path(base_dir) / p)
                candidates.append(Path.cwd() / p)
            found = next((c for c in candidates if c.exists()), None)
            if found is None:
                raise click.ClickException(
                    f"Logo file not found: {source} "
                    f"(looked in {', '.join(str(c) for c in candidates)})"
                )
            source = str(found)

        resolved.append({
            "label": entry.get("label", ""),
            "source": source,
            "is_url": is_url,
            "width": float(entry["width"]) if entry.get("width") else LOGO_WIDTH_IN,
        })
    return resolved


def load_logo_stream(logo):
    """Return a file path or a BytesIO (for URLs) suitable for add_picture."""
    if logo["is_url"]:
        with urllib.request.urlopen(logo["source"]) as resp:  # noqa: S310 (config-supplied)
            return io.BytesIO(resp.read())
    return logo["source"]


def derive_out_name(sponsor_name):
    return (
        "Corporate Sponsorship Agreement_BPD_"
        + sponsor_name.split()[0].replace("/", "-")
        + ".docx"
    )


def ask_amount(prompt, default):
    """Prompt for a currency amount, returning a formatted string like '5,000'."""
    while True:
        raw = ask(prompt, str(default))
        try:
            return normalize_amount(raw)
        except ValueError:
            print("  please enter a number, e.g. 5000")


def ask_date(prompt):
    """Prompt for an effective date, default today, return 'Month D, YYYY'."""
    raw = ask(prompt + " (e.g. 'March 13, 2026' or YYYY-MM-DD)", today_str())
    return normalize_date(raw)


# --------------------------------------------------------------------------- #
# Config resolution: YAML -> CLI overrides -> interactive fallback
# --------------------------------------------------------------------------- #
def load_yaml_config(path):
    """Read a YAML config file into a dict (empty file -> {})."""
    with open(path, "r", encoding="utf-8") as fh:
        data = yaml.safe_load(fh) or {}
    if not isinstance(data, dict):
        raise click.ClickException(
            f"Config file {path} must contain a YAML mapping at the top level."
        )
    return data


def resolve_inputs(cfg, cli_overrides, interactive, base_dir=None):
    """
    Merge YAML config (cfg) and CLI overrides, prompting for anything still
    missing when `interactive` is True. Returns the data dict consumed by
    build_document plus the resolved party blocks.
    """
    # Party blocks: start from defaults, overlay YAML keys.
    nonprofit = {**DEFAULT_NONPROFIT, **(cfg.get("nonprofit") or {})}
    community = {**DEFAULT_COMMUNITY, **(cfg.get("community") or {})}

    # Sponsor fields: YAML value, then CLI override (if given).
    merged = {}
    for key in FIELD_DEFAULTS:
        merged[key] = cli_overrides.get(key, cfg.get(key))

    def need(key):
        return merged.get(key) in (None, "")

    interactive_banner_shown = False

    def banner():
        nonlocal interactive_banner_shown
        if not interactive_banner_shown:
            print("=" * 64)
            print("  Corporate Sponsorship Contract Generator")
            print(f"  ({nonprofit['name']} / {community['name']})")
            print("=" * 64)
            print()
            interactive_banner_shown = True

    def require(key, prompt_fn):
        if need(key):
            if not interactive:
                raise click.ClickException(
                    f"Missing required field '{key}'. Provide it via --config, "
                    f"a CLI flag, or drop --non-interactive to be prompted."
                )
            banner()
            merged[key] = prompt_fn()

    # Required text fields.
    require("sponsor_name", lambda: ask("Sponsor legal name (e.g. 'JetBrains s.r.o.')"))
    require("sponsor_place", lambda: ask("Sponsor principal place of business (city, country)"))
    require("poc_name", lambda: ask("Sponsor point of contact (name)"))
    require("poc_email", lambda: ask("Sponsor email"))
    require("notice_address", lambda: ask_multiline("Sponsor mailing address"))

    # Optional / defaulted fields.
    if need("sponsor_entity"):
        merged["sponsor_entity"] = FIELD_DEFAULTS["sponsor_entity"]
    if need("payment_days"):
        merged["payment_days"] = FIELD_DEFAULTS["payment_days"]
    if merged.get("poc_title") is None:
        merged["poc_title"] = FIELD_DEFAULTS["poc_title"]

    # Amount: normalize whatever we have, else default.
    if need("amount"):
        merged["amount"] = normalize_amount(FIELD_DEFAULTS["amount"])
    else:
        merged["amount"] = normalize_amount(merged["amount"])

    # Effective date: normalize whatever we have, else today.
    if need("effective"):
        merged["effective"] = today_str()
    else:
        merged["effective"] = normalize_date(merged["effective"])

    # payment_days as plain string (YAML may give an int).
    merged["payment_days"] = str(merged["payment_days"]).strip()

    # notice_address: allow a multi-line string in YAML, normalize to list.
    addr = merged["notice_address"]
    if isinstance(addr, str):
        merged["notice_address"] = [ln.strip() for ln in addr.splitlines() if ln.strip()]

    # logos: normalize to a list of resolved {label, source, ...} dicts.
    merged["logos"] = normalize_logos(merged.get("logos"), base_dir=base_dir)

    # Output filename: derive default, ensure .docx extension.
    if need("out_name"):
        merged["out_name"] = derive_out_name(merged["sponsor_name"])
    if not merged["out_name"].lower().endswith(".docx"):
        merged["out_name"] += ".docx"

    return merged, nonprofit, community


def dump_starter_config():
    """Return a commented YAML starter config covering every field."""
    return """\
# Corporate Sponsorship Contract — configuration
# Any field omitted here is prompted for interactively (unless --non-interactive).

# --- Sponsor (the paying company) ---
sponsor_name: "ACME s.r.o."                  # required
sponsor_entity: "a private limited company"
sponsor_place: "Prague, Czech Republic"      # required

# --- Payment ---
amount: 5000                                 # USD; formatted automatically
payment_days: 30
effective: 2026-03-13                        # YYYY-MM-DD, or "March 13, 2026"

# --- Sponsor notice / point of contact ---
poc_name: "Jane Doe"                         # required
poc_title: "Head of Open Source"
poc_email: "jane@acme.example"               # required
notice_address:                              # required; one list item per line
  - "123 Main Street"
  - "Suite 400"
  - "Prague 110 00"
  - "Czech Republic"

# --- Logos (embedded in Exhibit B "Licensed Marks") ---
# Each entry is a local path or an http(s) URL. Relative paths are resolved
# against this config file's directory, then the current directory. Use the
# long form to add a caption or custom width (inches).
logos:
  - logos/gnome.png
  - logos/black-python-devs.png
  - path: logos/acme.png
    label: "ACME, Inc."
    width: 2.0
  # - https://example.com/sponsor-logo.png

# --- Output ---
out_name: ""                                 # blank -> derived from sponsor_name

# --- Default parties (override only if needed) ---
nonprofit:
  name: "GNOME Foundation"
  short: "GNOME"
  place: "Orinda, California"
  notice:
    - "#117 21c Orinda Way"
    - "Orinda"
    - "CA 94563"
    - "USA"
  email: "president@gnome.org"
  attn: "Allan Day/President"
community:
  name: "Black Python Devs"
  short: "BPD"
"""


# --------------------------------------------------------------------------- #
# Document building helpers
# --------------------------------------------------------------------------- #
def setup_styles(doc):
    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(11)
    pf = normal.paragraph_format
    pf.space_after = Pt(8)
    pf.line_spacing = 1.15


def add_title(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(13)
    return p


def add_body(doc, text, justify=True):
    p = doc.add_paragraph(text)
    if justify:
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    return p


def add_heading_num(doc, number, heading, body=""):
    """A top-level numbered section: bold '<n>. <heading>.' then body text."""
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0)
    run = p.add_run(f"{number}. ")
    run.bold = True
    run = p.add_run(heading + ".")
    run.bold = True
    if body:
        p.add_run(" " + body)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    return p


def add_sub(doc, label, body, level=1, bold_label=True):
    """A sub-item with hanging-style label (a., 1., etc.)."""
    indent = Inches(0.5 * level)
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = indent
    if label:
        run = p.add_run(label + " ")
        run.bold = bold_label
    if body:
        p.add_run(body)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    return p


def add_signature_block(doc, party_label):
    doc.add_paragraph(party_label).runs[0].bold = True
    for field in ("Signature", "Printed Name", "Title", "Date"):
        doc.add_paragraph("___________________________________")
        doc.add_paragraph(field)


# --------------------------------------------------------------------------- #
# Document assembly
# --------------------------------------------------------------------------- #
def build_document(data, nonprofit, community):
    np = nonprofit
    cm = community
    doc = Document()
    setup_styles(doc)

    add_title(doc, "CORPORATE SPONSORSHIP AGREEMENT")
    doc.add_paragraph()

    add_body(
        doc,
        f"This Corporate Sponsorship Agreement (the “Agreement”), effective as of "
        f"{data['effective']} (“Effective Date”), is between {np['name']}, a nonprofit "
        f"organization with a principal place of business at {np['place']} (“{np['short']}”) "
        f"and {data['sponsor_name']}, {data['sponsor_entity']} with a principal place of business "
        f"in {data['sponsor_place']} (“Company”). {np['short']} and/or Company may be "
        f"referred to as a “Party” or collectively as the “Parties.”",
    )

    add_body(
        doc,
        f"WHEREAS, {np['short']} is a 501(c)(3) tax-exempt, nonprofit organization that is formed "
        "for the charitable purposes of (i) facilitating the creation of GNOME, a free, "
        "technologically innovative open source computing platform, (ii) making that platform "
        "available for the benefit of the general public, including underserved or disadvantaged "
        "groups, and (iii) disseminating research and knowledge of GNOME and providing individuals "
        "with educational forums to develop their software development capabilities;",
    )
    add_body(
        doc,
        f"WHEREAS, {np['short']} is the fiscal sponsor of {cm['name']} (“{cm['short']}”), "
        "a community organization;",
    )
    add_body(
        doc,
        f"WHEREAS, Company wishes to support {np['short']}’s mission by providing a corporate "
        f"sponsorship payment to support the work of {cm['short']} (the “Contribution”);",
    )
    add_body(
        doc,
        f"WHEREAS, in connection with the Payment, {cm['short']} shall publicly recognize "
        "Company’s contribution in accordance with the terms and conditions of this Agreement.",
    )
    add_body(
        doc,
        "NOW THEREFORE, for the mutual agreements and promises set forth herein and for other "
        "valuable consideration, the receipt and sufficiency of which is hereby acknowledged, the "
        "Parties agree as follows:",
    )

    # 1. Responsibilities of the Parties
    add_heading_num(doc, "1", "Responsibilities of the Parties")
    add_sub(doc, "a.", "Responsibilities of Company.", level=1)
    add_sub(
        doc,
        "1.",
        f"Payment. Company shall deliver a payment of ${data['amount']} to {np['short']}, within "
        f"{data['payment_days']} days of receipt of a properly submitted invoice from {np['short']}. "
        "Company shall provide necessary vendor registration information (if applicable).",
        level=2,
    )
    add_sub(doc, "b.", f"Responsibilities of {np['short']}.", level=1)
    add_sub(
        doc,
        "1.",
        f"Use of Funds. {np['short']} shall use the Contribution to support the work of {cm['short']}.",
        level=2,
    )
    add_sub(
        doc,
        "2.",
        f"Promotional Activities. {cm['short']} agrees to use reasonable efforts to provide "
        "promotional recognition of Company’s Contribution as mutually agreed upon, and as "
        "described more fully in Exhibit A, attached hereto.",
        level=2,
    )
    add_sub(
        doc,
        "3.",
        f"Notwithstanding the foregoing, {np['short']} is not obligated to provide any services to "
        "Company, and no part of the Contribution received shall be deemed to be in exchange for "
        "services of any kind.",
        level=2,
    )
    add_sub(doc, "c.", "Mutual Responsibilities.", level=1)
    add_sub(
        doc,
        "",
        "Both parties agree to maintain open communication regarding:",
        level=2,
        bold_label=False,
    )
    add_sub(doc, "1.", "The approval and timing of all promotional content.", level=3)
    add_sub(doc, "2.", "Impact reporting and metrics sharing.", level=3)
    add_sub(doc, "3.", "Any changes to agreed-upon promotional activities.", level=3)

    # 2. Term and Termination
    add_heading_num(doc, "2", "Term and Termination")
    add_sub(
        doc,
        "a.",
        f"Term. This Agreement shall commence as of the Effective Date, and shall expire one year "
        f"from the date the Contribution is received by {np['short']}, unless terminated earlier "
        "pursuant to section 2b of the Agreement (the “Term”).",
        level=1,
    )
    add_sub(
        doc,
        "b.",
        "Termination. Before expiration of the Term, either Party may terminate this Agreement "
        "upon: (i) any material breach of the Agreement by the other Party, if such breach is not "
        "remedied to the reasonable satisfaction of the nonbreaching Party within twenty (20) days "
        "after written notice; (ii) twenty (20) days written notice to the other Party whenever the "
        "notifying Party in its sole discretion determines that the continuation of the Agreement "
        "will damage its reputation or good will; or (iii) written notice in the event one Party "
        "(a) becomes or is declared insolvent or bankrupt or is subject to the appointment of a "
        "trustee or receiver or any equivalent thereof, (b) is the subject of any proceeding "
        "related to its liquidation or insolvency (whether voluntary or involuntary) which is not "
        "dismissed within ninety (90) days, (c) makes an assignment for the benefit of creditors, "
        "(d) is subject to any sale, lease or other transfer of all or substantially all of its "
        "assets to any entity; or (e) is subject to a change of control (whether by merger, stock "
        "transfer or otherwise), except in the case of an initial public offering.",
        level=1,
    )

    # 3. Intellectual Property License
    add_heading_num(doc, "3", "Intellectual Property License")
    add_sub(
        doc,
        "a.",
        f"License of {np['short']} Marks. {np['short']} hereby grants to Company a limited, "
        "non-exclusive, nontransferable, royalty-free right and license to use, reproduce and "
        f"publicly display {np['short']}’s name and logo and {cm['short']}’s name and logo "
        f"(the “{np['short']} and {cm['short']} Marks”), as set forth on Exhibit B, solely "
        "in order to carry out Company’s responsibilities under this Agreement. "
        f"{np['short']} represents and warrants that, as of the date hereof, it has the power and "
        f"authority to license the {np['short']} and {cm['short']} Marks pursuant to the terms and "
        f"conditions of this Agreement. All uses of the {np['short']} and {cm['short']} Marks shall "
        f"inure to the benefit of {np['short']}.",
        level=1,
    )
    add_sub(
        doc,
        "b.",
        "License of Company Marks. Company hereby grants to "
        f"{np['short']} a limited, non-exclusive, nontransferable, royalty-free right and license "
        "to use, reproduce and publicly display Company’s name and logo (“Company Marks,” "
        f"and collectively with the {np['short']} and {cm['short']} Marks, the “Licensed Marks”), "
        "as set forth on Exhibit B, for purposes of providing recognition and acknowledgment of "
        "Company, as described in Section 1(b) of the Agreement. Company represents and warrants to "
        f"{np['short']} that, as of the date hereof, it has the power and authority to license the "
        "Company Marks pursuant to the terms and conditions of this Agreement. All uses of the "
        "Company Marks shall inure to the benefit of Company.",
        level=1,
    )

    # 4. Approvals and Quality Control
    add_heading_num(doc, "4", "Approvals and Quality Control")
    add_sub(
        doc,
        "a.",
        "Approvals. Each Party shall have the right to monitor the quality of the other Party’s "
        "use of its Licensed Marks. In order to afford the Licensing Party the opportunity to "
        "monitor use of its Licensed Marks, and ensure that its Licensed Marks are used in a "
        "positive manner consistent with its brand, the other Party shall submit to the Licensing "
        "Party for advanced written approval samples of any content or materials in which any of "
        "the Licensed Marks will appear. Each Party agrees to review such submissions within five "
        "(5) business days of its receipt, and will not unreasonably withhold its consent. Any "
        "references to a Party’s Licensed Marks shall contain the appropriate trademark notice "
        "provided from time to time by the Licensing Party.",
        level=1,
    )
    add_sub(
        doc,
        "b.",
        "Use of Marks. Each Party agrees not to use the Licensed Marks in a manner that is "
        "defamatory, misleading, libelous, obscene, or otherwise potentially damaging to the "
        "reputation or goodwill of the other Party.",
        level=1,
    )

    # 5. Confidentiality
    add_heading_num(
        doc,
        "5",
        "Confidentiality",
        "Each Party acknowledges that certain information received under this Agreement, including, "
        "but not limited to, non-public information concerning the disclosing Party’s business, "
        "methods, programs, activities, services, donors, members, users, consumers, or finances "
        "(“Confidential Information”), is confidential and proprietary. Each Party agrees to "
        "hold any and all Confidential Information of the disclosing Party in strict confidence and "
        "not disclose such Confidential Information to any third party, except as required by law or "
        "as authorized by the disclosing Party. Confidential Information does not include any "
        "information that: (i) is or becomes publicly available without breach of this Agreement; "
        "(ii) was known to the Party receiving such information at the time of its receipt from the "
        "disclosing Party; (iii) is rightfully received from a third party who did not acquire or "
        "disclose such information by a wrongful or tortious act; (iv) is independently developed by "
        "the receiving Party without reference to any Confidential Information; or (v) is inherently "
        "non-distinctive or non-proprietary in nature. This confidentiality provision shall survive "
        "termination or expiration of the Agreement.",
    )

    # 6. Representations & Warranties
    add_heading_num(doc, "6", "Representations & Warranties")
    add_sub(
        doc,
        "a.",
        "Each Party represents and warrants that it is the owner of its Licensed Marks and holds a "
        "current trademark registration or otherwise possesses all rights to utilize its Licensed "
        "Marks without approval of any third party, and otherwise has the ability to perform its "
        "obligations hereunder.",
        level=1,
        bold_label=False,
    )
    add_sub(
        doc,
        "b.",
        "Each Party represents and warrants that it has the right, power and authority to enter "
        "into this Agreement, grant the rights and benefits herein described and satisfy the "
        "obligations hereunder.",
        level=1,
        bold_label=False,
    )

    # 7. No Endorsement
    add_heading_num(
        doc,
        "7",
        "No Endorsement",
        f"The Promotional Activities shall in no way suggest endorsement by {np['short']} or of "
        f"{cm['short']} of any product or service of Company.",
    )

    # 8. Waiver
    add_heading_num(
        doc,
        "8",
        "Waiver",
        "The failure of any Party at any time to require performance of any provision of or resort "
        "to any remedy provided under this Agreement shall in no way affect the right of that Party "
        "to require performance or resort to a remedy at any time thereafter, nor shall the waiver "
        "by any Party of a breach be deemed to be a waiver of any subsequent breach. A waiver shall "
        "not be effective unless it is in writing and signed by the Party against whom the waiver "
        "is being enforced.",
    )

    # 9. Severability
    add_heading_num(
        doc,
        "9",
        "Severability",
        "If any term or provision of this Agreement is determined to be invalid, illegal, or "
        "incapable of being enforced by any law, regulation, or public policy, all other terms and "
        "provisions of this Agreement shall nevertheless remain in full force and effect.",
    )

    # 10. Survival
    add_heading_num(
        doc,
        "10",
        "Survival",
        "In addition to those provisions of the Agreement that, by their express terms, survive the "
        "termination or expiration of this Agreement, any provision which would, by its nature, "
        "survive the termination or expiration of this Agreement, shall do so.",
    )

    # 11. Assignment
    add_heading_num(
        doc,
        "11",
        "Assignment",
        f"Neither Company nor {np['short']} shall assign this Agreement, or any rights, interests, "
        "or obligations hereunder, without the prior written consent of the other Party.",
    )

    # 12. Notices
    add_heading_num(
        doc,
        "12",
        "Notices",
        "All notices required or permitted to be given under this Agreement shall be deemed to have "
        "been duly given: (a) when delivered in person, (b) upon confirmation of receipt when "
        "transmitted by e-mail, (c) upon receipt after dispatch by registered or certified mail, "
        "postage prepaid, or (d) on the next business day if sent by national overnight courier "
        "(with confirmation of delivery). All such notices shall be delivered to the addressees "
        "listed below (or at such other address as may be furnished in writing to the notifying "
        "Party):",
    )

    doc.add_paragraph()
    p = doc.add_paragraph()
    p.add_run(f"If to {np['name'].upper()}:").bold = True
    for line in np["notice"]:
        doc.add_paragraph(line).paragraph_format.space_after = Pt(0)
    doc.add_paragraph(f"Email: {np['email']}").paragraph_format.space_after = Pt(0)
    doc.add_paragraph(f"Attn: {np['attn']}")

    doc.add_paragraph()
    p = doc.add_paragraph()
    p.add_run(f"If to {data['sponsor_name']}:").bold = True
    for line in data["notice_address"]:
        doc.add_paragraph(line).paragraph_format.space_after = Pt(0)
    doc.add_paragraph(f"Email: {data['poc_email']}").paragraph_format.space_after = Pt(0)
    attn = data["poc_name"]
    if data["poc_title"]:
        attn = f"{data['poc_name']}/{data['poc_title']}"
    doc.add_paragraph(f"Attn: {attn}")

    # 13. Governing Law
    add_heading_num(
        doc,
        "13",
        "Governing Law",
        "This Agreement shall be construed and governed by the laws of the state of California, "
        "United States, without regard to conflicts of laws principles, and any dispute pertaining "
        "to this Agreement shall be brought exclusively in a court of competent jurisdiction within "
        "State of California, County of Contra Costa.",
    )

    # 14. Entire Agreement
    add_heading_num(
        doc,
        "14",
        "Entire Agreement",
        "This Agreement shall constitute the complete understanding of the Parties and shall "
        "supersede all prior written or oral agreements or understandings. This Agreement shall not "
        "be modified or amended except in writing signed by the authorized representatives of each "
        "Party.",
    )

    # 15. Counterparts
    add_heading_num(
        doc,
        "15",
        "Counterparts",
        "The Agreement may be executed in one or more counterparts, each of which will be deemed an "
        "original, but together will constitute one and the same instrument. A facsimile (or other "
        "copy made by reliable mechanical means) of a signed Agreement may be relied upon as an "
        "original.",
    )

    p = doc.add_paragraph("[Signatures on Following Page]")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Signature page
    doc.add_page_break()
    add_body(
        doc,
        "In WITNESS WHEREOF, the Parties have executed this Agreement by their duly authorized "
        "representatives on the dates indicated below.",
        justify=False,
    )
    doc.add_paragraph()
    add_signature_block(
        doc, f"{np['name'].upper()}\n(on its own behalf and on behalf of {cm['name']})"
    )
    doc.add_paragraph()
    add_signature_block(doc, data["sponsor_name"])

    # Exhibit A
    doc.add_page_break()
    p = add_title(doc, "EXHIBIT A")
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run("Promotional Activities").bold = True
    add_body(
        doc,
        f"{cm['short']} agrees to provide promotional recognition as mutually agreed upon, which "
        "may include:",
        justify=False,
    )
    for item in (
        f"Social media announcements across {cm['short']}'s platform on Mastodon and LinkedIn",
        "Newsletter mentions and features",
        "Blog post highlighting the partnership",
        f"Recognition at {cm['short']} events and activities",
        "Advertisement of promotional activities as discussed and agreed upon by the Parties.",
    ):
        doc.add_paragraph(item, style="List Bullet")

    # Exhibit B
    doc.add_page_break()
    add_title(doc, "EXHIBIT B")
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run("Licensed Marks").bold = True

    logos = data.get("logos") or []
    if logos:
        doc.add_paragraph()
        for logo in logos:
            if logo.get("label"):
                cap = doc.add_paragraph()
                cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                cap.add_run(logo["label"]).bold = True
            pic = doc.add_paragraph()
            pic.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pic.add_run().add_picture(
                load_logo_stream(logo), width=Inches(logo["width"])
            )
            doc.add_paragraph()

    return doc


# --------------------------------------------------------------------------- #
# Command-line interface
# --------------------------------------------------------------------------- #
@click.command(context_settings={"help_option_names": ["-h", "--help"]})
@click.option(
    "-c", "--config", "config_path",
    type=click.Path(exists=True, dir_okay=False, path_type=Path),
    help="YAML file supplying any or all fields.",
)
@click.option("--sponsor-name", help="Sponsor legal name, e.g. 'JetBrains s.r.o.'")
@click.option("--sponsor-entity", help="Sponsor entity description.")
@click.option("--sponsor-place", help="Sponsor principal place of business (city, country).")
@click.option("--amount", help="Payment amount in USD, e.g. 5000.")
@click.option("--payment-days", help="Payment due window in days.")
@click.option("--effective", help="Effective date (YYYY-MM-DD or 'March 13, 2026').")
@click.option("--poc-name", help="Sponsor point of contact name.")
@click.option("--poc-title", help="Sponsor point of contact title.")
@click.option("--poc-email", help="Sponsor email.")
@click.option(
    "--notice-address", "notice_address", multiple=True,
    help="One line of the sponsor mailing address; repeat for multiple lines.",
)
@click.option(
    "--logo", "logos", multiple=True,
    help="Path or URL to a logo for Exhibit B; repeat for multiple logos.",
)
@click.option("-o", "--output", "out_name", help="Output .docx filename.")
@click.option(
    "--non-interactive", is_flag=True,
    help="Never prompt; error if a required field is missing.",
)
@click.option(
    "--dump-config", is_flag=True,
    help="Print a starter YAML config to stdout and exit.",
)
def main(config_path, non_interactive, dump_config, notice_address, logos, **cli_fields):
    """Generate a corporate sponsorship agreement (.docx).

    Fields are resolved in order: YAML --config, then CLI flags, then
    interactive prompts for anything still missing.
    """
    if dump_config:
        click.echo(dump_starter_config(), nl=False)
        return

    cfg = load_yaml_config(config_path) if config_path else {}
    base_dir = config_path.parent if config_path else None

    # Collect only the CLI flags the user actually supplied (non-None / non-empty).
    cli_overrides = {k: v for k, v in cli_fields.items() if v not in (None, "")}
    if notice_address:
        cli_overrides["notice_address"] = list(notice_address)
    if logos:
        cli_overrides["logos"] = list(logos)

    data, nonprofit, community = resolve_inputs(
        cfg, cli_overrides, interactive=not non_interactive, base_dir=base_dir
    )

    doc = build_document(data, nonprofit, community)

    out_path = Path(data["out_name"]).expanduser()
    # Bare filenames go into the gitignored output/ dir; explicit paths are kept.
    if not out_path.is_absolute() and out_path.parent == Path("."):
        out_path = Path.cwd() / OUTPUT_DIR / out_path
    elif not out_path.is_absolute():
        out_path = Path.cwd() / out_path
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))

    click.echo()
    click.echo("=" * 64)
    click.echo(f"  Contract created: {out_path}")
    click.echo("=" * 64)
    click.echo(f"  Sponsor : {data['sponsor_name']}")
    click.echo(f"  Amount  : ${data['amount']}")
    click.echo(f"  Date    : {data['effective']}")
    click.echo("  Review and fill Exhibit B (Licensed Marks) before sending.")


if __name__ == "__main__":
    main()
