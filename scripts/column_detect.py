"""Header-name heuristics for mapping arbitrary spreadsheet columns to canonical
fields. Never guesses silently on a required field -- callers must check
`missing` and `ambiguous` and surface them to the user instead of proceeding.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field


def _norm(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", str(s).lower()).strip()


# Each canonical field maps to a list of substrings; a header matches a field
# if the normalized header contains (or equals) one of these substrings.
JE_FIELD_PATTERNS = {
    "entry_id": ["je no", "je #", "journal no", "journal #", "entry no",
                 "entry #", "transaction id", "txn id", "entry id", "doc number", "doc no"],
    "date": ["date"],
    "account": ["account name", "gl account", "account", "acct"],
    "debit": ["debit", "dr amt", "dr amount", " dr", "debits"],
    "credit": ["credit", "cr amt", "cr amount", " cr", "credits"],
    "amount": ["amount", "amt"],
    "description": ["description", "memo", "memo description", "particulars", "line description"],
}

COA_FIELD_PATTERNS = {
    "account_name": ["fully qualified name", "full name", "account name", "name"],
    "account_type": ["account type", "type", "category"],
}


@dataclass
class DetectionResult:
    mapping: dict = field(default_factory=dict)       # canonical field -> original header
    missing: list = field(default_factory=list)       # required fields with no candidate
    ambiguous: dict = field(default_factory=dict)      # field -> list of candidate headers
    notes: list = field(default_factory=list)          # human-readable explanation lines


def _match_fields(headers, patterns):
    """Return {field: [candidate headers]} for every field with >=1 candidate."""
    candidates = {field_name: [] for field_name in patterns}
    normed = {h: _norm(h) for h in headers}
    for field_name, needles in patterns.items():
        for header, norm_header in normed.items():
            if any(norm_header == n or n in norm_header for n in needles):
                candidates[field_name].append(header)
    return candidates


def detect_je_columns(headers: list[str]) -> DetectionResult:
    result = DetectionResult()
    candidates = _match_fields(headers, JE_FIELD_PATTERNS)

    # debit/credit vs single signed amount are mutually exclusive strategies.
    has_debit_credit = candidates["debit"] and candidates["credit"]
    has_amount = candidates["amount"] and not has_debit_credit

    required = ["account"]
    if has_debit_credit:
        required += ["debit", "credit"]
    elif has_amount:
        required += ["amount"]
    else:
        result.missing.append("debit/credit (or a single signed amount column)")

    required.append("entry_id") if candidates["entry_id"] else result.notes.append(
        "No journal/entry-id column found -- will fall back to grouping by "
        "contiguous row blocks separated by blank rows."
    )

    for f in required:
        opts = candidates.get(f, [])
        if not opts:
            result.missing.append(f)
        elif len(opts) == 1:
            result.mapping[f] = opts[0]
        else:
            result.ambiguous[f] = opts

    # optional fields: take unambiguous single matches, otherwise leave unmapped
    for f in ("entry_id", "date", "description"):
        opts = candidates.get(f, [])
        if len(opts) == 1:
            result.mapping[f] = opts[0]
        elif len(opts) > 1:
            result.ambiguous[f] = opts

    if has_amount:
        result.mapping["amount"] = candidates["amount"][0] if len(candidates["amount"]) == 1 else None
        if len(candidates["amount"]) > 1:
            result.ambiguous["amount"] = candidates["amount"]

    return result


def detect_coa_columns(headers: list[str]) -> DetectionResult:
    result = DetectionResult()
    candidates = _match_fields(headers, COA_FIELD_PATTERNS)

    name_opts = candidates.get("account_name", [])
    if not name_opts:
        result.missing.append("account_name")
    elif len(name_opts) == 1:
        result.mapping["account_name"] = name_opts[0]
    else:
        # Prefer an exact "fully qualified name" header over a plain "name" if both exist.
        fqn_like = [h for h in name_opts if "full" in _norm(h)]
        if fqn_like:
            result.mapping["account_name"] = fqn_like[0]
        else:
            result.ambiguous["account_name"] = name_opts

    type_opts = candidates.get("account_type", [])
    if len(type_opts) == 1:
        result.mapping["account_type"] = type_opts[0]
    elif len(type_opts) > 1:
        result.ambiguous["account_type"] = type_opts

    return result
