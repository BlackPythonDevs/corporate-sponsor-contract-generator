# Corporate Sponsorship Contract Creator

Generate a formatted corporate sponsorship agreement as a Microsoft Word
(`.docx`) document from a single Python script.

The generator builds a complete, lawyer-style sponsorship agreement — intro
recitals, fifteen numbered sections (responsibilities, term & termination, IP
license, confidentiality, governing law, notices, and more), a signature page,
and two exhibits (Promotional Activities and Licensed Marks). It ships with the
[GNOME Foundation](https://www.gnome.org/) (the 501(c)(3) fiscal sponsor) and
[Black Python Devs](https://blackpythondevs.com/) (the sponsored community org)
as the default parties, but **every field is configurable** — including the
nonprofit and community blocks — via a YAML config, command-line flags, or
interactive prompts.

> ⚠️ This tool produces a draft document for convenience. It is not legal
> advice. Have any generated agreement reviewed before signing.

## Requirements

- Python 3.9+
- [`uv`](https://docs.astral.sh/uv/) (recommended) — runs the script and
  installs dependencies automatically from the inline script metadata.

The script declares its own dependencies (`python-docx`, `click`, `PyYAML`), so
with `uv` there's no manual `pip install` or virtualenv step.

## Usage

### Fully interactive

Prompts for every required field:

```bash
uv run generate_sponsorship_contract.py
```

### From a YAML config

```bash
uv run generate_sponsorship_contract.py --config sponsor.yaml
```

### YAML plus a CLI override, no prompting

```bash
uv run generate_sponsorship_contract.py -c sponsor.yaml --amount 7500 --non-interactive
```

### Write a starter config to fill in

```bash
uv run generate_sponsorship_contract.py --dump-config > sponsor.yaml
```

Fields are resolved in order: **YAML `--config` → CLI flags → interactive
prompts** for anything still missing. Pass `--non-interactive` to error out
instead of prompting when a required field is absent.

## Configurable fields

| Field | Flag | Default |
| --- | --- | --- |
| Sponsor legal name | `--sponsor-name` | _required_ |
| Sponsor entity description | `--sponsor-entity` | `a private limited company` |
| Sponsor principal place of business | `--sponsor-place` | _required_ |
| Payment amount (USD) | `--amount` | `5000` |
| Payment due window (days) | `--payment-days` | `30` |
| Effective date | `--effective` | today |
| Point of contact name | `--poc-name` | _required_ |
| Point of contact title | `--poc-title` | _(blank)_ |
| Point of contact email | `--poc-email` | _required_ |
| Notice mailing address | `--notice-address` (repeatable) | _required_ |
| Logo for Exhibit B | `--logo` (repeatable) | _(none)_ |
| Output filename | `-o` / `--output` | derived from sponsor name |

The `nonprofit:` and `community:` party blocks can only be overridden via the
YAML config. Dates accept `YYYY-MM-DD` or free-form text like
`March 13, 2026`; amounts are reformatted automatically (e.g. `5000` →
`$5,000`).

### Logos

The `logos` config (or `--logo` flag) embeds images in Exhibit B
("Licensed Marks"). Each entry is a local path or an `http(s)` URL. Relative
paths resolve against the config file's directory first, then the current
directory. The long form supports a caption (`label`) and a custom `width` in
inches:

```yaml
logos:
  - logos/gnome.png
  - logos/black-python-devs.png
  - path: logos/acme.png
    label: "ACME, Inc."
    width: 2.0
```

## Output

Bare filenames are written to the gitignored `output/` directory. Filenames
with a directory component, or absolute paths, are written as-is. If no output
name is given, one is derived from the sponsor name, e.g.
`Corporate Sponsorship Agreement_BPD_ACME.docx`.

## License

[MIT](LICENSE)
