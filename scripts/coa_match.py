"""Chart-of-accounts loading and account-name matching, in tiers:
exact FullyQualifiedName match -> leaf-name-ambiguous -> fuzzy suggestion
(never auto-accepted) -> phantom (no match at all).
"""
from __future__ import annotations

import difflib
from collections import defaultdict

FUZZY_CUTOFF = 0.82  # below this, don't even suggest -- just report phantom.


class ChartOfAccounts:
    def __init__(self, rows: list[dict]):
        """rows: list of {"FullyQualifiedName": ..., "Name": ..., "AccountType": ...}"""
        self.fqn_set = set()
        self.leaf_index: dict[str, list[str]] = defaultdict(list)
        for row in rows:
            fqn = (row.get("FullyQualifiedName") or "").strip()
            if not fqn:
                continue
            self.fqn_set.add(fqn)
            leaf = fqn.split(":")[-1].strip()
            self.leaf_index[leaf].append(fqn)

    def match(self, account_name: str) -> dict:
        """Returns one of:
        {"status": "exact", "fqn": ...}
        {"status": "ambiguous_leaf", "candidates": [...]}
        {"status": "fuzzy_suggestion", "candidates": [(fqn, score), ...]}
        {"status": "phantom"}
        """
        name = (account_name or "").strip()
        if not name:
            return {"status": "phantom"}

        if name in self.fqn_set:
            return {"status": "exact", "fqn": name}

        leaf_matches = self.leaf_index.get(name, [])
        if leaf_matches:
            return {"status": "ambiguous_leaf", "candidates": leaf_matches}

        # Fuzzy: compare against all known FullyQualifiedNames and leaf names.
        pool = list(self.fqn_set)
        scored = []
        for candidate in pool:
            score = difflib.SequenceMatcher(None, name.lower(), candidate.lower()).ratio()
            leaf = candidate.split(":")[-1]
            leaf_score = difflib.SequenceMatcher(None, name.lower(), leaf.lower()).ratio()
            best = max(score, leaf_score)
            if best >= FUZZY_CUTOFF:
                scored.append((candidate, best))
        if scored:
            scored.sort(key=lambda t: -t[1])
            return {"status": "fuzzy_suggestion", "candidates": scored[:3]}

        return {"status": "phantom"}


def load_coa_rows_from_dicts(records: list[dict], mapping: dict) -> list[dict]:
    """Normalize arbitrary COA rows (per detected column mapping) into the
    {FullyQualifiedName, Name, AccountType} shape ChartOfAccounts expects."""
    name_col = mapping.get("account_name")
    type_col = mapping.get("account_type")
    out = []
    for rec in records:
        fqn = str(rec.get(name_col, "")).strip()
        if not fqn or fqn.lower() == "nan":
            continue
        out.append({
            "FullyQualifiedName": fqn,
            "Name": fqn.split(":")[-1].strip(),
            "AccountType": str(rec.get(type_col, "")).strip() if type_col else "",
        })
    return out
