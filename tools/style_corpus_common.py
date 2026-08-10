"""
style_corpus_common.py
-----------------------
Shared logic for the cross-agent style-learning pipeline (agent-commons issue #3):
redaction classification and recipient-tier mapping, used by both
sent_corpus_pull.py (bulk historical corpus) and draft_final_diff_capture.py
(ongoing draft/final diff capture). Factored out here once a second script
needed the same logic, rather than duplicating it a second time -- a single
"keep in sync" comment across two copies is a drift risk; a shared import
is not.

Nothing in this module touches Outlook COM or the filesystem -- it's pure
classification/mapping logic, safe to unit-test without any real mailbox
access.
"""

import re

# ---------------------------------------------------------------------------
# Redaction
# ---------------------------------------------------------------------------

REDACTION_PATTERNS = {
    "health": [
        r"\bsick(?:ness)?\b", r"\bunwell\b", r"\bpoorly\b", r"\bsigned off\b",
        r"\bfit note\b", r"\bsick note\b", r"\bGP appointment\b",
        r"\bdoctor'?s? appointment\b", r"\bhospital\b", r"\bdiagnos(?:is|ed)\b",
        r"\bsurgery\b", r"\boperation\b", r"\bmedical (?:condition|appointment|leave)\b",
        r"\bmental health\b", r"\btherapy\b", r"\bcounsell?ing\b",
        r"\boccupational health\b", r"\bstress leave\b", r"\banxiety\b",
        r"\bdepression\b", r"\bmedication\b", r"\btreatment\b", r"\bcancer\b",
        r"\bchemotherapy\b", r"\bdisabilit(?:y|ies)\b", r"\blong[- ]term sick\b",
    ],
    "bereavement": [
        r"\bbereavement\b", r"\bcompassionate leave\b", r"\bfuneral\b",
        r"\bpassed away\b", r"\bpassing of\b", r"\bcondolences?\b", r"\bloss of\b",
        r"\bsadly died\b", r"\bdeath of\b", r"\bmemorial service\b", r"\bwake\b",
        r"\bhospice\b", r"\bpalliative\b", r"\bterminally ill\b", r"\bsadly passed\b",
    ],
    "hr_case": [
        r"\bdisciplinary\b", r"\bgrievance\b", r"\binvestigation\b",
        r"\bsafeguarding\b", r"\bHR case\b", r"\bconfidential HR\b",
        r"\bcapability process\b", r"\bperformance improvement plan\b", r"\bPIP\b",
        r"\bwhistleblow(?:er|ing)?\b", r"\bsuspend(?:ed|sion)\b",
        r"\bformal warning\b", r"\bmisconduct\b", r"\btribunal\b",
        r"\bwithout prejudice\b", r"\bsettlement agreement\b", r"\bmediation\b",
    ],
    "absence": [
        r"\breturn to work\b", r"\bphased return\b", r"\blong[- ]term absence\b",
        r"\babsence review\b", r"\bwelfare meeting\b", r"\bwelfare check\b",
        r"\boccupational health referral\b", r"\bfit for work\b",
    ],
}

_COMPILED = {
    cat: [re.compile(p, re.IGNORECASE) for p in pats]
    for cat, pats in REDACTION_PATTERNS.items()
}


def classify(text):
    """Return the sorted set of category names whose pattern matched anywhere in text."""
    hits = set()
    for cat, patterns in _COMPILED.items():
        for pat in patterns:
            if pat.search(text):
                hits.add(cat)
                break
    return sorted(hits)


def is_sensitive(subject, body):
    return classify(f"{subject}\n{body}")


# ---------------------------------------------------------------------------
# Named-entity co-occurrence (informational only -- audit-ledger enrichment,
# never a redaction gate on its own)
# ---------------------------------------------------------------------------

KNOWN_NAMES = {
    'Athena Artuso', 'Marie Cooksey', 'Sarah Rowles', 'Simon Burford',
    'Asta Palmer', 'James Salas Guillen', "Michael O'Sullivan",
    'Anna Carter-Windle', 'Anthony Kong', 'Beth Gray', 'Christopher Sanders',
    'David Johnson', 'Emma Fitz-Gibbon', 'Henry Acheampong', 'Iyanuloluwa Akinsanya',
    'Julie Hickman', 'Marie King', 'Michelle Williams', 'Nathan Kirwan',
    'Susan Pratt', 'Anne Mortimer', 'Nicholas Chandler', 'Steve McBrearty',
}
_GENERIC_NAME_RE = re.compile(r"\b[A-Z][a-z]+ [A-Z][a-z]+\b")


def mentions_named_person(text):
    for name in KNOWN_NAMES:
        if name in text:
            return True
    return bool(_GENERIC_NAME_RE.search(text))


# ---------------------------------------------------------------------------
# Recipient tier -- confirmed live 10 Aug 2026 (agent-commons corpus/sent-items).
# "Marie" resolved empirically from real 'to'-field frequency in Kevin's Sent
# Mail: Marie Cooksey 118x vs Marie King 3x -- not guessed.
# ---------------------------------------------------------------------------

DIRECT_REPORTS = ["Michael O'Sullivan", "James Salas Guillen", "Asta Palmer"]
SENIOR_MANAGEMENT = ["Simon Burford", "Marie Cooksey"]


def recipient_tier(to_field):
    """Mixed-audience 'to' fields (both a direct-report and senior-management
    name present) resolve to 'other' -- don't guess which register dominates.
    No reliable 'peer' name-list exists yet, so genuine peers currently land
    in 'other' too, same as truly unmapped recipients."""
    to = to_field or ""
    has_dr = any(name in to for name in DIRECT_REPORTS)
    has_sm = any(name in to for name in SENIOR_MANAGEMENT)
    if has_dr and has_sm:
        return "other"
    if has_sm:
        return "senior-management"
    if has_dr:
        return "direct-report"
    return "other"


# ---------------------------------------------------------------------------
# Outlook COM item-Class constant -- confirmed live 10 Aug 2026: Sent Items
# and Drafts both hold non-mail items (meeting requests/responses/cancellations,
# Class 53/54/55/56/57) that lack mail-style Body/To. Filter explicitly rather
# than relying on exception shape (see agent-commons memory/index.json entry
# 2026-08-10-outlook-com-sent-items-folder-contains-non-mail-items...).
# ---------------------------------------------------------------------------

OL_MAIL_CLASS = 43  # olMail
