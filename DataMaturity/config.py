# ──────────────────────────────────────────────────────────────
#  DATA MATURITY CONFIGURATION
# ──────────────────────────────────────────────────────────────


# ══════════════════════════════════════════════════════════════
# 🎨 THEME COLORS (Used in Visualizations + Reports)
# ══════════════════════════════════════════════════════════════

UNIQU_PURPLE   = "#6A0DAD"
UNIQU_MAGENTA  = "#FF00FF"
UNIQU_LAVENDER = "#E6E6FA"

UNIQU_LIGHT_BG = "#F8F5FF"
UNIQU_TEXT     = "#2E2E2E"
UNIQU_GREY     = "#B0B0B0"


# ══════════════════════════════════════════════════════════════
# 📊 MATURITY RATING SCALE
# ══════════════════════════════════════════════════════════════

RATING_LABELS = [
    "Adhoc",
    "Repeatable",
    "Defined",
    "Managed",
    "Optimised",
]

RATING_TO_SCORE = {
    "Adhoc":       1,
    "Repeatable":  2,
    "Defined":     3,
    "Managed":     4,
    "Optimised":   5,
}


# ══════════════════════════════════════════════════════════════
# 🔗 DQ → MATURITY MAPPING
# ══════════════════════════════════════════════════════════════
# Used in helpers.py → dq_score_to_maturity_level()

DQ_MATURITY_MAP = [
    (95, "Optimised"),
    (80, "Managed"),
    (60, "Defined"),
    (40, "Repeatable"),
    (0,  "Adhoc"),
]


# ══════════════════════════════════════════════════════════════
# 🏢 DEFAULT MASTER DATA OBJECTS
# ══════════════════════════════════════════════════════════════

DEFAULT_MASTER_OBJECTS = [
    "Customer",
    "Vendor",
    "Material",
    "Finance",
    "Employee",
]


# ══════════════════════════════════════════════════════════════
# 🧭 MATURITY DIMENSIONS (DAMA-Aligned)
# ══════════════════════════════════════════════════════════════

MATURITY_DIMS = [
    "Data Governance",
    "Data Quality",
    "Data Architecture",
    "Data Security",
    "Data Operations",
]


# ══════════════════════════════════════════════════════════════
# ❓ QUESTION BANK
# ══════════════════════════════════════════════════════════════
# Structure required by helpers.py → build_question_df()

QUESTION_BANK = {

    # ──────────────────────────────────────────────────────────
    # DATA GOVERNANCE
    # ──────────────────────────────────────────────────────────
    "Data Governance": [
        {
            "id": "DG1",
            "section": "Framework",
            "question": "Is there a formal data governance framework?",
            "weight": 1,
        },
        {
            "id": "DG2",
            "section": "Ownership",
            "question": "Are data owners and stewards defined?",
            "weight": 1,
        },
        {
            "id": "DG3",
            "section": "Policies",
            "question": "Are governance policies documented and enforced?",
            "weight": 1,
        },
    ],


    # ──────────────────────────────────────────────────────────
    # DATA QUALITY
    # ──────────────────────────────────────────────────────────
    "Data Quality": [
        {
            "id": "DQ1",
            "section": "Standards",
            "question": "Are data quality standards defined?",
            "weight": 1,
        },
        {
            "id": "DQ2",
            "section": "Monitoring",
            "question": "Is data quality monitored continuously?",
            "weight": 1,
        },
        {
            "id": "DQ3",
            "section": "Remediation",
            "question": "Are issue remediation workflows defined?",
            "weight": 1,
        },
    ],


    # ──────────────────────────────────────────────────────────
    # DATA ARCHITECTURE
    # ──────────────────────────────────────────────────────────
    "Data Architecture": [
        {
            "id": "DA1",
            "section": "Design",
            "question": "Is enterprise data architecture defined?",
            "weight": 1,
        },
        {
            "id": "DA2",
            "section": "Integration",
            "question": "Are integration standards documented?",
            "weight": 1,
        },
        {
            "id": "DA3",
            "section": "Metadata",
            "question": "Is metadata centrally managed?",
            "weight": 1,
        },
    ],


    # ──────────────────────────────────────────────────────────
    # DATA SECURITY
    # ──────────────────────────────────────────────────────────
    "Data Security": [
        {
            "id": "DS1",
            "section": "Access",
            "question": "Are access controls role-based?",
            "weight": 1,
        },
        {
            "id": "DS2",
            "section": "Encryption",
            "question": "Is sensitive data encrypted?",
            "weight": 1,
        },
        {
            "id": "DS3",
            "section": "Compliance",
            "question": "Are compliance audits conducted?",
            "weight": 1,
        },
    ],


    # ──────────────────────────────────────────────────────────
    # DATA OPERATIONS
    # ──────────────────────────────────────────────────────────
    "Data Operations": [
        {
            "id": "DO1",
            "section": "Monitoring",
            "question": "Are data pipelines monitored?",
            "weight": 1,
        },
        {
            "id": "DO2",
            "section": "Incident Mgmt",
            "question": "Are incident response processes defined?",
            "weight": 1,
        },
        {
            "id": "DO3",
            "section": "Automation",
            "question": "Are operations automated where possible?",
            "weight": 1,
        },
    ],
}
