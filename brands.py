"""
Brand registry: maps program-file filenames to brand prefixes, and holds
per-brand product-name mappings and bottle counts.

Detection is by filename keyword (case-insensitive). If a file doesn't match
any known brand, it's flagged so the user can correct it rather than guessed.
"""
import re

# filename keyword -> (prefix, display name, bottles_per_case default map)
# Order matters: more specific keywords first.
BRAND_RULES = [
    ("clr",            "LR", "Roederer Collection"),
    ("collection",     "LR", "Roederer Collection"),
    ("roederer_estate","RE", "Roederer Estate"),
    ("estate",         "RE", "Roederer Estate"),
    ("ott",            "OT", "Domaine Ott"),
    ("dominus",        "DE", "Dominus"),
    ("mdm",            "MU", "Marques de Murrieta"),
    ("murrieta",       "MU", "Marques de Murrieta"),
    ("scharffenberger","SC", "Scharffenberger"),
    ("livio",          "LF", "Livio Felluga"),
    ("felluga",        "LF", "Livio Felluga"),
    ("bon_vivant",     "BV", "Bon Vivant"),
    ("carpe_diem",     "CD", "Carpe Diem"),
    ("cdb",            "CB", "Castiglion del Bosco"),
    ("castiglion",     "CB", "Castiglion del Bosco"),
    ("clos_jordanne",  "CJ", "Clos Jordanne"),
    ("delas",          "DL", "Delas"),
    ("domaine_anderson","DA","Domaine Anderson"),
    ("duas_quintas",   "DQ", "Duas Quintas"),
    ("jackson",        "JT", "Jackson-Triggs"),
    ("loudenne",       "LD", "Loudenne"),
    ("meerlust",       "ML", "Meerlust"),
    ("merry_edwards",  "ME", "Merry Edwards"),
    ("moueix",         "MX", "Moueix"),
    ("pez",            "PZ", "Pez"),
    ("pio_cesare",     "PC", "Pio Cesare"),
    ("querciabella",   "QB", "Querciabella"),
    ("ramos_pinto",    "RP", "Ramos Pinto"),
    ("regnard",        "RG", "Regnard"),
    ("schlumberger",   "SL", "Schlumberger"),
]

DEFAULT_BOTTLES = {"750ML": 6, "375ML": 12, "1.5L": 3, "1750": 3, "3L": 1}


def detect_brand_from_filename(filename):
    """Return (prefix, display_name) or (None, None) if no match."""
    f = filename.lower()
    for kw, prefix, name in BRAND_RULES:
        if kw in f:
            return prefix, name
    return None, None


def bottles_for(size_label):
    s = str(size_label or "")
    if "375" in s: return 12
    if "1.5" in s or "1750" in s or "175" in s: return 3
    if "3l" in s.lower(): return 1
    return 6


# ---- Sub-brand routing: split one prefix into display groups by MMD ----
# Returns a sub-brand label for a given (prefix, mmd). If a prefix has no
# sub-brand rules, the brand's normal display name is used.
def sub_brand(prefix, mmd):
    m = (mmd or "").upper()
    if prefix == "LR":
        # Cristal: LR1711xxx, LR1911xxx
        if m.startswith("LR1711") or m.startswith("LR1911"):
            return "LR Cristal"
        # Vintage line: Brut Vintage 16, Blanc de Blancs, Rose, Nature
        if m.startswith(("LR1401", "LR1411", "LR1451", "LR1501", "LR1601")):
            return "LR Vintage"
        # Everything else under LR = Collection (245/246/247/244/243 + variants)
        return "Roederer Collection"
    return None  # no split for other brands


# ---- Per-brand product-name maps (MMD prefix -> friendly product name) ----
# Where a product isn't listed, the app falls back to the bill-back description.
PRODUCT_NAMES = {
    "LR": {  # Roederer Collection
        "LR1184525": "Collection 245 750ML",
        "LR1184525G": "Collection 245 G 750ML",
        "LR1184615": "Collection 246 375ML",
        "LR1184615G": "Collection 246 G 375ML",
        "LR1184715": "Collection 247 375ML",
        "LR1184425": "Collection 244 750ML",
        "LR1184325": "Collection 243 750ML",
        "LR1401625": "Brut Vintage 16",
        "LR1411825": 'Brut "Nature" 18',
        "LR1451825": 'Brut "Nature" Rose 18',
        "LR1501725": "Blanc de Blancs 17",
        "LR1601725": "Brut Rose 17",
    },
    # Other brands: add maps as validated. Until then, products show the
    # bill-back description text as their label (the chosen placeholder).
}


def product_name(prefix, mmd, bb_description):
    """Friendly name if mapped; otherwise the bill-back description text."""
    names = PRODUCT_NAMES.get(prefix, {})
    if mmd in names:
        return names[mmd]
    # try the base (strip trailing G/D variant) for a softer match
    if mmd:
        base = mmd.rstrip("GD")
        for k, v in names.items():
            if k.rstrip("GD") == base:
                return v
    # placeholder = bill-back description text
    return (bb_description or mmd or "Unknown").strip()
