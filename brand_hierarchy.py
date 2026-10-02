"""
Complete brand hierarchy with description-based matching.
Maps product descriptions to brand and sub-group labels.
"""
import re

BRAND_KEYWORDS = [
    # (keyword_pattern, display_brand, sub_group_or_None)
    # Ordered by specificity (longest/most specific first)
    # Domaine Ott — must come before generic wine-term keywords (e.g. BLANC DE BL)
    ("DOM OTT BY OTT", "Domaine Ott", "By.Ott"),
    ("DOM OTT CHT ROMASSAN", "Domaine Ott", "Ott Crus"),
    ("DOM OTT CHT DE SELLE", "Domaine Ott", "Ott Crus"),
    ("DOM OTT ROMASSAN", "Domaine Ott", "Ott Crus"),
    ("DOM OTT SELLE", "Domaine Ott", "Ott Crus"),
    ("DOM OTT MIREILLE", "Domaine Ott", "Ott Crus"),
    ("DOM OTT ETOILE", "Domaine Ott", "Ott Crus"),
    ("DOM OTT", "Domaine Ott", "Ott Crus"),
    ("BY OTT", "Domaine Ott", "By.Ott"),
    ("LOUIS ROEDERER BRUT CRISTAL", "Roederer Cristal", None),
    ("BRUT CRIS", "Roederer Cristal", None),
    ("CRISTAL ROSE", "Roederer Cristal", None),
    ("CRISTAL", "Roederer Cristal", None),
    ("ROEDERER DEMI SEC", "Roederer VRBB", None),
    ("BRUT COL 245", "Roederer Collection", None),
    ("BRUT COL 246", "Roederer Collection", None),
    ("BRUT COL 247", "Roederer Collection", None),
    ("BRUT COL 244", "Roederer Collection", None),
    ("BRUT COL 243", "Roederer Collection", None),
    ("CARTE BLANCHE", "Roederer Collection", None),
    ("BRUT PREMIER", "Roederer Collection", None),
    ("COLLECTION", "Roederer Collection", None),
    # Bon Vivant must precede generic VRBB wine-type words ("BON VIVANT BRUT ROSE"
    # must not be caught by the generic "BRUT ROSE" -> VRBB rule).
    ("BON VIVANT", "Bon Vivant", None),
    # Estate must come BEFORE the generic VRBB wine-type keywords below,
    # because "ROEDERER EST BRUT ROSE" / "...BLANC" are Estate products,
    # not VRBB. Brand-specific prefixes win over generic wine-type words.
    ("L'ERMITAGE", "Roederer Estate", None),
    ("ROEDERER EST", "Roederer Estate", None),
    # Plain "LOUIS ROEDERER BRUT <2-digit year>" -> VRBB, for ANY vintage
    # (e.g. BRUT 15, BRUT 16, BRUT 17...). Uses a regex so new vintages never
    # need a code change. Placed before the literal generic BRUT keywords.
    # Does NOT catch BRUT COL/CRIS/CRISTAL/NATURE/ROSE (handled by their own rules).
    (re.compile(r"LOUIS ROEDERER BRUT \d{2}\b"), "Roederer VRBB", None),
    ("BRUT 16", "Roederer VRBB", None),
    ("BRUT VINTAGE", "Roederer VRBB", None),
    ("BL DE BL", "Roederer VRBB", None),
    ("BLANC DE BL", "Roederer VRBB", None),
    ("BLANC DE BLANC", "Roederer VRBB", None),
    ("BRUT NAT", "Roederer VRBB", None),
    ("BRUT NATURE", "Roederer VRBB", None),
    ("BRUT ROSE", "Roederer VRBB", None),
    ("BY OTT", "Domaine Ott", "By.Ott"),
    ("CHT ROMASSAN", "Domaine Ott", "Ott Crus"),
    ("CHT DE SELLE", "Domaine Ott", "Ott Crus"),
    ("SELLE", "Domaine Ott", "Ott Crus"),
    ("ROMASSAN", "Domaine Ott", "Ott Crus"),
    ("MIREILLE", "Domaine Ott", "Ott Crus"),
    ("ETOILE", "Domaine Ott", "Ott Crus"),
    ("DOM OTT", "Domaine Ott", "Ott Crus"),
    ("SCHARFFENBERGER", "Scharffenberger", None),
    ("DELAS", "Delas", None),
    ("FELLUGA", "Livio Felluga", None),
    ("PIO CESARE", "Pio Cesare", None),
    ("QUERCIABELLA MONGRANA", "Querciabella", "Mongrana"),
    ("QUERCIABELLA", "Querciabella", "Querciabella"),
    ("DOMINUS NAPANOOK", "Dominus", None),
    ("NAPANOOK", "Dominus", None),
    ("DOMINUS OTHELLO", "Dominus", None),
    ("OTHELLO", "Dominus", None),
    ("DOMINUS", "Dominus", None),
    ("MURRIETA", "Marqués de Murrieta", None),
    ("DALMAU", "Marqués de Murrieta", None),
    ("CAPELLANIA", "Marqués de Murrieta", None),
    ("PAZO BARRANTES", "Marqués de Murrieta", None),
    ("CASTIGLION", "Castiglion del Bosco", None),
    ("MEERLUST", "Meerlust", None),
    ("REGNARD", "Regnard", None),
    ("LADOUCETTE", "De Ladoucette", None),
    ("MARC BREDIF", "De Ladoucette", None),
    ("LA POUSSIE", "De Ladoucette", None),
    ("COMTE LAFOND", "De Ladoucette", None),
    ("CLOS JORDANNE", "Clos Jordanne", None),
    ("RAMOS PINTO", "Ramos Pinto", None),
    ("INNISKILLIN", "Inniskillin Icewine", None),
    ("JACKSON TRIGGS", "Jackson Triggs Niagara", None),
    ("SCHLUMBERGER", "Schlumberger", None),
    ("FLEUR DU CAP", "Distell", None),
    ("LOUDENNE", "Château Loudenne", None),
    ("CH DE PEZ", "CLR Bordeaux", None),
    ("CARPE DIEM", "Carpe Diem", None),
    ("DUAS QUINTAS", "Duas Quintas", None),
    ("ESPERTO", "Esperto", None),
    ("MERRY EDWARDS", "Merry Edwards", None),
    ("DOMAINE ANDERSON", "Domaine Anderson", None),
    ("CH DE SALES", "Ets. J-P Moueix", None),
    ("CH PEYMOUTON", "Ets. J-P Moueix", None),
    ("PEYMOUTON", "Ets. J-P Moueix", None),
]

def get_brand_and_subgroup(mmm_id, description):
    """Match description against keywords to return (brand_name, sub_group).

    A keyword may be a plain string (substring match) or a compiled regex
    (pattern search), allowing rules like "any LOUIS ROEDERER BRUT vintage".
    """
    if not description:
        return ("⚠ Unmapped", None)
    desc_upper = str(description).upper()
    for keyword, brand, sub_group in BRAND_KEYWORDS:
        if hasattr(keyword, "search"):          # compiled regex
            if keyword.search(desc_upper):
                return (brand, sub_group)
        elif keyword in desc_upper:             # literal substring
            return (brand, sub_group)
    return ("⚠ Unmapped", None)


def build_material_to_mmm(deal_file_path):
    """Placeholder - not used for bill-backs since SGWS material # don't crosswalk."""
    return {}
