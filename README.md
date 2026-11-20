# JE Validator

A Claude skill (and standalone script) that validates a double-entry journal-entry Excel export before you import it into accounting software — mapping each transaction's account name against a chart of accounts, confirming debits/credits balance to the cent per entry, and flagging anything that would trip up a native importer.

Ships with a bundled manufacturing-company chart of accounts as a default, a demo file so you can try it with zero setup, and an 8-case regression suite. See `SKILL.md` for the full feature list and validation rules.

## Usage

You can use this three ways — none of them require Claude at all unless you want one of the AI-assisted options.

### Option 0: Plain Python (no AI product needed)

```bash
git clone https://github.com/ingenaitive/je-validator-skill.git
cd je-validator-skill
python3 -m venv venv
./venv/bin/pip install -r requirements.txt

# Try the bundled demo (zero setup)
./venv/bin/python scripts/validate_je.py

# Validate your own file
./venv/bin/python scripts/validate_je.py path/to/your_journal_entries.xlsx

# Validate against your own chart of accounts instead of the bundled manufacturing one
./venv/bin/python scripts/validate_je.py path/to/your_journal_entries.xlsx --coa path/to/your_coa.xlsx
```

### Option 1: Claude Code

Clone this repo, then copy the skill folder (this one — the folder containing `SKILL.md`) into either:

- **Project-scoped** (shareable via your own repo): `<your-project>/.claude/skills/je-validator/`
- **Personal** (available across all your projects): `~/.claude/skills/je-validator/`

No registration step — Claude Code auto-discovers skills in those locations. If it doesn't pick it up right away in an existing session, run `/reload-skills`. Then just ask Claude Code to validate a journal-entry file and it'll invoke this skill.

### Option 2: Claude Desktop app / claude.ai (web)

claude.ai supports uploading custom Skills in this same format:

1. Zip this skill folder (`SKILL.md` must sit at the top level inside the zip).
2. In claude.ai: **Customize → Skills → + → Create skill → Upload**, and select the zip.
3. Enable **Settings → Capabilities → "Create and edit files"** (code execution) — required for the skill to actually run.
4. Available on Free/Pro/Max/Ultra/Team/Enterprise plans (Team/Enterprise admins can disable user-uploaded skills org-wide).

Note: in claude.ai the skill runs in a sandboxed environment, so attach your Excel file to the chat rather than pointing the script at a local file path.

## Output

Each run writes:
- `validation_report.csv` — one row per journal entry (`SourceFile, EntryId, Status, ErrorCount, Errors`)
- `je_import_ready.csv` — only entries that passed, reformatted in QBO's exact native-importer header format (`Journal No., Journal Date, Account Name, Description, Debits, Credits`)

## Roadmap

Built for QuickBooks Online first; designed to add other double-entry platforms (e.g. Bigcapital) as additional output-format targets within this same skill, rather than as separate repos/skills.
