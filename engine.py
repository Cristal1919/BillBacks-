"""
Reconciliation engine wrapper for the local web app.
Wraps the validated da_audit logic and returns:
  - summary dict (cases, spent, allowance, over_under, gp_pct)
  - global_df, not_in_program_df, overspend_df  (pandas DataFrames for on-screen display)
  - workbook_bytes  (the Global + Exceptions .xlsx, in memory)

The matching methodology is unchanged from the validated tool:
  closest-PTA tier per premise, positive-only overage, weighted averages.
"""
import io
from collections import defaultdict
import pandas as pd

import da_audit as A  # the validated engine (copied into this folder)


def _bucket(da):
    if not da or da <= 20: return "Regular"
    if da <= 40: return "Deep Deal"
    return "Aggressive"


def _group_for(desc, size, group_rules):
    """Map a bill-back line to a product group using size hints.
    group_rules: list of (label, [substrings]) — first match wins.
    Falls back to the brand's default group."""
    d = (desc or "").upper(); s = str(size or "")
    for label, needles in group_rules:
        if any(n in s or n in d for n in needles):
            return label
    return group_rules[-1][0]  # default = last rule


def _deal_for(deals_by_mmd, mmd, band):
    """Find the deal row for THIS product only.
    Exact MMD first; then the same base MMD with G/D variant collapsed
    (e.g. LR1184525G <-> LR1184525). NEVER a broad prefix match across
    different products/sizes — better to return None (blank price) than to
    grab another product's NCP and fabricate a price."""
    if not mmd:
        return None
    c = list(deals_by_mmd.get(mmd, []))
    if not c:
        base = mmd.rstrip("GD")
        # only collapse the exact G/D variant of the SAME base, not a prefix scan
        for k, v in deals_by_mmd.items():
            if k.rstrip("GD") == base:
                c.extend(v)
    act = [x for x in c if x.get("is_active")
           and isinstance(x.get("net_case_price"), (int, float)) and x["net_case_price"] > 0]
    if not act:
        return None
    return min(act, key=lambda x: abs(x["net_case_price"] - (band or 0)))


def run_multi(bb_inputs, vip_xlsx, deal_xlsx, program_files, selected_brands):
    """Multi-brand entry point.
    bb_inputs:       list of (pdf_path, region)
    program_files:   dict {prefix: program_xlsx_path}
    selected_brands: list of (prefix, display_name)
    One combined deal file + one combined VIP file cover all brands.
    Returns a stacked result: per-brand blocks + grand total + per-brand coverage.
    """
    import brands as BR
    brand_results = []
    grand = {"cases": 0, "spent": 0, "allowance": 0, "gn": 0, "gd": 0}
    coverage_rows = []

    for prefix, display in selected_brands:
        prog_path = program_files.get(prefix)
        if not prog_path:
            coverage_rows.append({"Brand": display, "Status": "NO PROGRAM FILE — skipped",
                                  "Parsed $": 0, "Audited $": 0, "Unmatched $": 0})
            continue

        # Determine sub-brands for this prefix (e.g. LR -> Collection/Vintage/Cristal).
        # If none, the single display name is used as one group.
        probe = BR.sub_brand(prefix, "")  # None if prefix has no split rule
        if probe is None:
            sub_targets = [(display, None)]   # (label, filter) ; None filter = all
        else:
            # discover which sub-brand labels actually occur for this prefix
            sub_targets = [(lbl, lbl) for lbl in _discover_sub_brands(
                bb_inputs, deal_xlsx, prefix)]
            if not sub_targets:
                sub_targets = [(display, None)]

        for label, sub_filter in sub_targets:
            try:
                r = _process_one_brand(bb_inputs, vip_xlsx, deal_xlsx, prog_path,
                                       prefix, label, sub_filter=sub_filter)
            except Exception as e:
                coverage_rows.append({"Brand": label, "Status": f"ERROR: {e}",
                                      "Parsed $": 0, "Audited $": 0, "Unmatched $": 0})
                continue
            if r["summary"]["spent"] == 0 and len(r["global_df"]) == 0:
                continue  # this sub-brand had no lines; skip quietly
            brand_results.append(r)
            s = r["summary"]
            grand["cases"] += s["cases"]; grand["spent"] += s["spent"]
            grand["allowance"] += s["allowance"]; grand["gn"] += r["gn"]; grand["gd"] += r["gd"]
            coverage_rows.append({
                "Brand": label, "Status": "OK" if r["unmatched_total"] < 1 else "PARTIAL — see exceptions",
                "Parsed $": round(r["parsed_total"], 0),
                "Audited $": round(s["spent"], 0),
                "Unmatched $": round(r["unmatched_total"], 0),
            })

    grand_summary = {
        "cases": grand["cases"], "spent": grand["spent"], "allowance": grand["allowance"],
        "over_under": grand["spent"] - grand["allowance"],
        "gp_pct": grand["gn"] / grand["gd"] if grand["gd"] else 0,
    }
    workbook_bytes = _build_multi_workbook(brand_results, grand_summary, coverage_rows)
    return {
        "brand_results": brand_results, "grand_summary": grand_summary,
        "coverage_rows": coverage_rows, "workbook_bytes": workbook_bytes,
    }


def run_reconciliation(bb_inputs, vip_xlsx, deal_xlsx, prog_xlsx,
                       brand_prefix, bottles_per_case):
    """Single-brand entry point (kept for backward compatibility)."""
    # 1. Parse the four inputs
    program = A.parse_program_file(prog_xlsx, brand_prefix=brand_prefix)
    vip_rows = A.parse_vip_file(vip_xlsx, brand_prefix=brand_prefix)
    material_to_mmd = A.build_material_to_mmd_mapping(deal_xlsx, brand_prefix=brand_prefix)
    ca_deals = A.extract_ca_deals(deal_xlsx, brand_prefix=brand_prefix)

    deals_by_mmd = defaultdict(list)
    for d in ca_deals:
        if d.get("mmd_id"):
            deals_by_mmd[d["mmd_id"]].append(d)

    # 2. Parse bill-back. Each PDF is parsed ONCE with its user-assigned region.
    bb_lines = []
    for pdf_path, region in bb_inputs:
        try:
            bb_lines += A.parse_billback_pdf(
                pdf_path, region=region, brand_prefix=brand_prefix,
                material_to_mmd=material_to_mmd)
        except Exception:
            pass
    # de-dup identical lines that matched under both region passes
    seen = set(); uniq = []
    for b in bb_lines:
        k = (b["region"], b["mat"], b["band"], b["cases"], b["da"], b["desc"])
        if k not in seen:
            seen.add(k); uniq.append(b)
    bb_lines = uniq

    # 3. Run the audit (closest-PTA, premise split, positive-only overage)
    audit = A.run_audit(bb_lines, vip_rows, program, material_to_mmd)

    # 4. Build group rules from the brand's bottle map (keys are size labels)
    group_rules = []
    for size_label in bottles_per_case:
        needles = [size_label.replace("ML", ""), size_label]
        group_rules.append((f"{brand_prefix} {size_label}", needles))
    if not group_rules:
        group_rules = [(f"{brand_prefix} (all)", [""])]

    def bpc(group_label):
        for size_label, n in bottles_per_case.items():
            if size_label in group_label:
                return n
        return 6

    # 5. Aggregate group × bucket
    G = defaultdict(lambda: defaultdict(lambda: {
        "cs": 0, "spent": 0, "allow": 0, "ncp_w": 0, "prof_w": 0,
        "li_w": 0, "spa_w": 0, "ds_w": 0}))
    for a in audit:
        bb = a["bb"]
        if bb["cases"] <= 0:
            continue
        g = _group_for(bb["desc"], bb["size"], group_rules)
        b = _bucket(bb["da"]); cs = bb["cases"]
        on_cs = a["cases_by_premise"].get("ON", 0); off_cs = a["cases_by_premise"].get("OFF", 0)
        on_da = a["authorized_da_by_premise"].get("ON", 0); off_da = a["authorized_da_by_premise"].get("OFF", 0)
        tp = on_cs + off_cs
        arate = (on_da * on_cs + off_da * off_cs) / tp if tp > 0 else 0
        dl = _deal_for(deals_by_mmd, a.get("mmd"), bb["band"])
        ncp = dl["net_case_price"] if dl else 0
        spa = (dl.get("spa") or 0) if dl else 0
        ds = (dl.get("discount_support") or 0) if dl else 0
        li = (dl.get("adjusted_laid_in") or 0) if dl else 0
        profcs = ncp + spa + ds - li
        s = G[g][b]
        s["cs"] += cs; s["spent"] += bb["tot"]; s["allow"] += arate * cs
        s["ncp_w"] += ncp * cs; s["prof_w"] += profcs * cs
        s["li_w"] += li * cs; s["spa_w"] += spa * cs; s["ds_w"] += ds * cs

    # 6. Build the Global DataFrame
    rows = []
    TC = TS = TA = TGN = TGD = 0
    for g in [r[0] for r in group_rules if r[0] in G]:
        for b in ["Regular", "Deep Deal", "Aggressive"]:
            s = G[g].get(b)
            if not s or s["cs"] == 0:
                continue
            cs = s["cs"]; ncp = s["ncp_w"] / cs if cs else 0
            rows.append({
                "Group": g, "Deal Level": b,
                "Cases": round(cs, 2),
                "Sell Price Case": round(ncp, 2),
                "Sell Price Bottle": round(ncp / bpc(g), 2),
                "Dollar Profit Case": round(s["prof_w"] / cs, 2) if cs else 0,
                "Total Dist. Revenue": round(s["ncp_w"], 0),
                "Dist. Dollar Profit": round(s["prof_w"], 0),
                "DA Spent": round(s["spent"], 0),
                "Program Allowance": round(s["allow"], 0),
                "Over/Under": round(s["spent"] - s["allow"], 0),
                "GP %": round(s["prof_w"] / s["ncp_w"], 4) if s["ncp_w"] else 0,
            })
            TC += cs; TS += s["spent"]; TA += s["allow"]; TGN += s["prof_w"]; TGD += s["ncp_w"]
    global_df = pd.DataFrame(rows)

    # 7. Exceptions
    def allowed_rate(a):
        on_cs = a["cases_by_premise"].get("ON", 0); off_cs = a["cases_by_premise"].get("OFF", 0)
        on_da = a["authorized_da_by_premise"].get("ON", 0); off_da = a["authorized_da_by_premise"].get("OFF", 0)
        tp = on_cs + off_cs
        return (on_da * on_cs + off_da * off_cs) / tp if tp > 0 else 0

    def tier_str(a):
        on = a["matched_tier_by_premise"].get("ON"); off = a["matched_tier_by_premise"].get("OFF")
        parts = []
        if a["cases_by_premise"].get("OFF", 0) > 0 and off: parts.append(f"{off['tier']} (OFF)")
        if a["cases_by_premise"].get("ON", 0) > 0 and on: parts.append(f"{on['tier']} (ON)")
        return ", ".join(parts) if parts else "—"

    nip = [a for a in audit if a.get("not_in_program") and a["bb"]["cases"] > 0]
    over = sorted([a for a in audit if a["overage_total"] > 0.5 and not a.get("not_in_program")],
                  key=lambda a: -a["overage_total"])

    # "No tier matched": MMD IS in program, has cases, but no premise got a
    # real authorized DA (allowance rate is 0). These would otherwise silently
    # count 100% of their DA as overspend. Flag them as their own category.
    def has_real_allowance(a):
        return allowed_rate(a) > 0
    no_tier = [a for a in audit
               if not a.get("not_in_program") and a["bb"]["cases"] > 0
               and not has_real_allowance(a)]

    not_in_program_df = pd.DataFrame([{
        "Region": a["bb"]["region"], "MMD": a.get("mmd"), "Description": a["bb"]["desc"],
        "Band": a["bb"]["band"], "Cases": round(a["bb"]["cases"], 2),
        "BB DA/Cs": a["bb"]["da"], "DA Charged $": round(a["bb"]["tot"], 0),
        "Flag": "MMD not in program file",
    } for a in nip])

    no_tier_df = pd.DataFrame([{
        "Region": a["bb"]["region"], "MMD": a.get("mmd"), "Description": a["bb"]["desc"],
        "Band": a["bb"]["band"], "Cases": round(a["bb"]["cases"], 2),
        "BB DA/Cs": a["bb"]["da"], "DA Charged $": round(a["bb"]["tot"], 0),
        "Flag": "No program tier matched at this price/premise — allowance unknown",
    } for a in no_tier])

    overspend_df = pd.DataFrame([{
        "Region": a["bb"]["region"], "MMD": a.get("mmd"), "Description": a["bb"]["desc"],
        "Band": a["bb"]["band"], "Cases": round(a["bb"]["cases"], 2),
        "BB DA/Cs": a["bb"]["da"], "Allowed/Cs": round(allowed_rate(a), 2),
        "Over $": round(a["overage_total"], 0), "Matched Program Tier": tier_str(a),
    } for a in over])

    summary = {
        "cases": TC, "spent": TS, "allowance": TA,
        "over_under": TS - TA, "gp_pct": TGN / TGD if TGD else 0,
    }

    # 8. Build the downloadable workbook (Global + Exceptions)
    workbook_bytes = _build_workbook(global_df, not_in_program_df, overspend_df, summary)

    return {
        "summary": summary, "global_df": global_df,
        "not_in_program_df": not_in_program_df, "overspend_df": overspend_df,
        "workbook_bytes": workbook_bytes,
    }


def _discover_sub_brands(bb_inputs, deal_xlsx, prefix):
    """Find which sub-brand labels actually occur in the bill-back for this prefix."""
    import brands as BR
    material_to_mmd = A.build_material_to_mmd_mapping(deal_xlsx, brand_prefix=prefix)
    labels = []
    seen = set()
    for pdf_path, region in bb_inputs:
        try:
            lines = A.parse_billback_pdf(pdf_path, region=region, brand_prefix=prefix,
                                         material_to_mmd=material_to_mmd)
        except Exception:
            continue
        for b in lines:
            if b["cases"] <= 0:
                continue
            mmd = material_to_mmd.get(b["mat"]) or ""
            lbl = BR.sub_brand(prefix, mmd)
            if lbl and lbl not in seen:
                seen.add(lbl); labels.append(lbl)
    # stable order: Collection, Vintage, Cristal, then any others
    order = {"Roederer Collection": 0, "LR Vintage": 1, "LR Cristal": 2}
    return sorted(labels, key=lambda x: order.get(x, 99))


def _process_one_brand(bb_inputs, vip_xlsx, deal_xlsx, prog_xlsx, prefix, display, sub_filter=None):
    """Process a single brand for the multi-brand report.
    Groups by INDIVIDUAL PRODUCT (MMD), labels via brands.product_name
    (friendly name or bill-back description placeholder), and tracks coverage."""
    import brands as BR
    program = A.parse_program_file(prog_xlsx, brand_prefix=prefix)
    vip_rows = A.parse_vip_file(vip_xlsx, brand_prefix=prefix)
    material_to_mmd = A.build_material_to_mmd_mapping(deal_xlsx, brand_prefix=prefix)
    ca_deals = A.extract_ca_deals(deal_xlsx, brand_prefix=prefix)
    deals_by_mmd = defaultdict(list)
    for d in ca_deals:
        if d.get("mmd_id"):
            deals_by_mmd[d["mmd_id"]].append(d)

    bb_lines = []
    for pdf_path, region in bb_inputs:
        try:
            bb_lines += A.parse_billback_pdf(
                pdf_path, region=region, brand_prefix=prefix,
                material_to_mmd=material_to_mmd)
        except Exception:
            pass
    seen = set(); uniq = []
    for b in bb_lines:
        k = (b["region"], b["mat"], b["band"], b["cases"], b["da"], b["desc"])
        if k not in seen:
            seen.add(k); uniq.append(b)
    bb_lines = uniq

    parsed_total = 0
    audit = A.run_audit(bb_lines, vip_rows, program, material_to_mmd)

    # Aggregate by INDIVIDUAL PRODUCT (mmd) x bucket
    P = defaultdict(lambda: defaultdict(lambda: {
        "cs": 0, "spent": 0, "allow": 0, "ncp_w": 0, "prof_w": 0,
        "li_w": 0, "spa_w": 0, "ds_w": 0, "label": "", "size": ""}))
    audited_total = 0
    # filter audit to this sub-brand if requested
    if sub_filter is not None:
        audit = [a for a in audit if BR.sub_brand(prefix, a.get("mmd") or "") == sub_filter]
    parsed_total = sum(a["bb"]["tot"] for a in audit if a["bb"]["cases"] > 0)
    for a in audit:
        bb = a["bb"]
        if bb["cases"] <= 0:
            continue
        mmd = a.get("mmd") or bb["mat"]
        label = BR.product_name(prefix, a.get("mmd"), bb["desc"])
        b = _bucket(bb["da"]); cs = bb["cases"]
        on_cs = a["cases_by_premise"].get("ON", 0); off_cs = a["cases_by_premise"].get("OFF", 0)
        on_da = a["authorized_da_by_premise"].get("ON", 0); off_da = a["authorized_da_by_premise"].get("OFF", 0)
        tp = on_cs + off_cs
        arate = (on_da * on_cs + off_da * off_cs) / tp if tp > 0 else 0
        dl = _deal_for(deals_by_mmd, a.get("mmd"), bb["band"])
        ncp = dl["net_case_price"] if dl else 0
        spa = (dl.get("spa") or 0) if dl else 0
        ds = (dl.get("discount_support") or 0) if dl else 0
        li = (dl.get("adjusted_laid_in") or 0) if dl else 0
        profcs = ncp + spa + ds - li
        s = P[label][b]
        s["label"] = label; s["size"] = bb["size"]
        s["cs"] += cs; s["spent"] += bb["tot"]; s["allow"] += arate * cs
        s["ncp_w"] += ncp * cs; s["prof_w"] += profcs * cs
        s["li_w"] += li * cs; s["spa_w"] += spa * cs; s["ds_w"] += ds * cs
        audited_total += bb["tot"]

    rows = []
    TC = TS = TA = GN = GD = 0
    for label in sorted(P.keys()):
        for b in ["Regular", "Deep Deal", "Aggressive"]:
            s = P[label].get(b)
            if not s or s["cs"] == 0:
                continue
            cs = s["cs"]; ncp = s["ncp_w"] / cs if cs else 0
            bpc = BR.bottles_for(s["size"])
            rows.append({
                "Product": label, "Deal Level": b,
                "Cases": round(cs, 2),
                "Sell Price Case": round(ncp, 2),
                "Sell Price Bottle": round(ncp / bpc, 2) if bpc else 0,
                "Dollar Profit Case": round(s["prof_w"] / cs, 2) if cs else 0,
                "DA Spent": round(s["spent"], 0),
                "Program Allowance": round(s["allow"], 0),
                "Over/Under": round(s["spent"] - s["allow"], 0),
                "GP %": round(s["prof_w"] / s["ncp_w"], 4) if s["ncp_w"] else 0,
            })
            TC += cs; TS += s["spent"]; TA += s["allow"]; GN += s["prof_w"]; GD += s["ncp_w"]
    global_df = pd.DataFrame(rows)

    def allowed_rate(a):
        on = a["cases_by_premise"].get("ON", 0); off = a["cases_by_premise"].get("OFF", 0)
        ond = a["authorized_da_by_premise"].get("ON", 0); offd = a["authorized_da_by_premise"].get("OFF", 0)
        tp = on + off
        return (ond * on + offd * off) / tp if tp > 0 else 0

    def tier_str(a):
        on = a["matched_tier_by_premise"].get("ON"); off = a["matched_tier_by_premise"].get("OFF")
        parts = []
        if a["cases_by_premise"].get("OFF", 0) > 0 and off: parts.append(f"{off['tier']} (OFF)")
        if a["cases_by_premise"].get("ON", 0) > 0 and on: parts.append(f"{on['tier']} (ON)")
        return ", ".join(parts) if parts else "—"

    over = sorted([a for a in audit if a["overage_total"] > 0.5 and not a.get("not_in_program") and a["bb"]["cases"] > 0],
                  key=lambda a: -a["overage_total"])
    overspend_df = pd.DataFrame([{
        "Region": a["bb"]["region"], "MMD": a.get("mmd"), "Description": a["bb"]["desc"],
        "Band": a["bb"]["band"], "Cases": round(a["bb"]["cases"], 2),
        "BB DA/Cs": a["bb"]["da"], "Allowed/Cs": round(allowed_rate(a), 2),
        "Over $": round(a["overage_total"], 0), "Matched Program Tier": tier_str(a),
    } for a in over])

    # Not-in-program and no-tier-matched (allowance could not be determined)
    nip = [a for a in audit if a.get("not_in_program") and a["bb"]["cases"] > 0]
    no_tier = [a for a in audit
               if not a.get("not_in_program") and a["bb"]["cases"] > 0
               and not (allowed_rate(a) > 0)]
    not_in_program_df = pd.DataFrame([{
        "Region": a["bb"]["region"], "MMD": a.get("mmd"), "Description": a["bb"]["desc"],
        "Band": a["bb"]["band"], "Cases": round(a["bb"]["cases"], 2),
        "BB DA/Cs": a["bb"]["da"], "DA Charged $": round(a["bb"]["tot"], 0),
        "Flag": "MMD not in program file",
    } for a in nip])
    no_tier_df = pd.DataFrame([{
        "Region": a["bb"]["region"], "MMD": a.get("mmd"), "Description": a["bb"]["desc"],
        "Band": a["bb"]["band"], "Cases": round(a["bb"]["cases"], 2),
        "BB DA/Cs": a["bb"]["da"], "DA Charged $": round(a["bb"]["tot"], 0),
        "Flag": "No program tier matched at this price/premise — allowance unknown",
    } for a in no_tier])

    summary = {"cases": TC, "spent": TS, "allowance": TA,
               "over_under": TS - TA, "gp_pct": GN / GD if GD else 0}
    return {
        "prefix": prefix, "display": display, "summary": summary,
        "global_df": global_df, "overspend_df": overspend_df,
        "not_in_program_df": not_in_program_df, "no_tier_df": no_tier_df,
        "gn": GN, "gd": GD,
        "parsed_total": parsed_total,
        "unmatched_total": max(0, parsed_total - audited_total),
    }


def _build_workbook(global_df, nip_df, over_df, summary):
    """Write the same clean, no-color Global + Exceptions workbook to bytes."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    thin = Side(style="thin", color="000000")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    DOL = '"$"#,##0'; DOL2 = '"$"#,##0.00'; NUM = "#,##0.00"; PCT = "0.0%"

    def C(ws, r, c, v, bold=False, sz=11, fmt=None, align="center"):
        cell = ws.cell(r, c, v)
        cell.font = Font(bold=bold, size=sz, name="Calibri")
        cell.alignment = Alignment(horizontal=align, vertical="center")
        if fmt: cell.number_format = fmt
        cell.border = box
        return cell

    # ---- Global sheet ----
    ws = wb.active; ws.title = "Global"; ws.sheet_view.showGridLines = False
    ws.cell(1, 1, "DA Reconciliation — Program Analysis").font = Font(bold=True, size=14, name="Calibri")
    card = [("Total Cases", summary["cases"], NUM), ("DA Spent (Billed)", summary["spent"], DOL),
            ("Program Allowance", summary["allowance"], DOL), ("Over / Under", summary["over_under"], DOL),
            ("Dist GP %", summary["gp_pct"], PCT)]
    for i, (lab, val, fmt) in enumerate(card):
        a_ = ws.cell(3, 1 + i, lab); a_.font = Font(bold=True, size=10, name="Calibri")
        a_.alignment = Alignment(horizontal="center", wrap_text=True); a_.border = box
        v_ = ws.cell(4, 1 + i, val); v_.font = Font(bold=True, size=13, name="Calibri")
        v_.alignment = Alignment(horizontal="center"); v_.number_format = fmt; v_.border = box

    HDRS = ["Deal Level", "Sell Price Case", "Sell Price Bottle", "Dollar Profit Case",
            "Cases", "Total Dist. Revenue", "Dist. Dollar Profit", "DA Spent",
            "Program Allowance", "Over/Under", "GP %"]
    FMT = [None, DOL2, DOL2, DOL2, NUM, DOL, DOL, DOL, DOL, DOL, PCT]
    colmap = ["Deal Level", "Sell Price Case", "Sell Price Bottle", "Dollar Profit Case",
              "Cases", "Total Dist. Revenue", "Dist. Dollar Profit", "DA Spent",
              "Program Allowance", "Over/Under", "GP %"]

    r = 6
    for g in global_df["Group"].unique():
        ws.cell(r, 1, g).font = Font(bold=True, size=12, name="Calibri")
        ws.cell(r, 1).border = box; ws.cell(r, 1).alignment = Alignment(horizontal="left")
        for i, h in enumerate(HDRS[1:], 1):
            C(ws, r, 1 + i, h, bold=True)
        r += 1
        sub = global_df[global_df["Group"] == g]
        tot = {"Cases": 0, "Total Dist. Revenue": 0, "Dist. Dollar Profit": 0,
               "DA Spent": 0, "Program Allowance": 0, "Over/Under": 0}
        gn = gd = 0
        for _, row in sub.iterrows():
            ws.cell(r, 1, row["Deal Level"]).font = Font(size=11, name="Calibri")
            ws.cell(r, 1).border = box; ws.cell(r, 1).alignment = Alignment(horizontal="left")
            for i, key in enumerate(colmap[1:], 1):
                C(ws, r, 1 + i, row[key], fmt=FMT[i])
            for k in tot: tot[k] += row[k]
            gd += row["Total Dist. Revenue"]; gn += row["Dist. Dollar Profit"]
            r += 1
        tl = ws.cell(r, 1, "Total"); tl.font = Font(bold=True, size=11, name="Calibri")
        tl.border = box; tl.alignment = Alignment(horizontal="left")
        C(ws, r, 5, round(tot["Cases"], 2), bold=True, fmt=NUM)
        C(ws, r, 6, tot["Total Dist. Revenue"], bold=True, fmt=DOL)
        C(ws, r, 7, tot["Dist. Dollar Profit"], bold=True, fmt=DOL)
        C(ws, r, 8, tot["DA Spent"], bold=True, fmt=DOL)
        C(ws, r, 9, tot["Program Allowance"], bold=True, fmt=DOL)
        C(ws, r, 10, tot["Over/Under"], bold=True, fmt=DOL)
        C(ws, r, 11, (gn / gd if gd else 0), bold=True, fmt=PCT)
        for i in (2, 3, 4): C(ws, r, i, "")
        r += 2

    gl = ws.cell(r, 1, "GRAND TOTAL"); gl.font = Font(bold=True, size=12, name="Calibri")
    gl.border = box; gl.alignment = Alignment(horizontal="left")
    C(ws, r, 5, round(summary["cases"], 2), bold=True, sz=12, fmt=NUM)
    C(ws, r, 8, summary["spent"], bold=True, sz=12, fmt=DOL)
    C(ws, r, 9, summary["allowance"], bold=True, sz=12, fmt=DOL)
    C(ws, r, 10, summary["over_under"], bold=True, sz=12, fmt=DOL)
    C(ws, r, 11, summary["gp_pct"], bold=True, sz=12, fmt=PCT)
    for i in (2, 3, 4, 6, 7): C(ws, r, i, "")
    ws.column_dimensions["A"].width = 20
    for c in range(2, 12): ws.column_dimensions[get_column_letter(c)].width = 15

    # ---- Exceptions sheet ----
    ex = wb.create_sheet("Exceptions"); ex.sheet_view.showGridLines = False
    ex.cell(1, 1, "Exceptions — Where We Overspend vs Program").font = Font(bold=True, size=13, name="Calibri")

    r = 3
    ex.cell(r, 1, "NOT IN PROGRAM — no authorized tier exists").font = Font(bold=True, size=12, name="Calibri")
    r += 1
    nip_cols = ["Region", "MMD", "Description", "Band", "Cases", "BB DA/Cs", "Over $"]
    for i, h in enumerate(nip_cols): C(ex, r, 1 + i, h, bold=True)
    r += 1
    nip_tot = 0
    for _, row in nip_df.iterrows():
        for i, k in enumerate(nip_cols):
            fmt = DOL if k == "Over $" else (DOL if k == "Band" else (DOL2 if k == "BB DA/Cs" else (NUM if k == "Cases" else None)))
            C(ex, r, 1 + i, row[k], fmt=fmt, align="left" if k in ("Description",) else "center")
        nip_tot += row["Over $"]; r += 1
    C(ex, r, 1, "Subtotal", bold=True, align="left")
    C(ex, r, 7, nip_tot, bold=True, fmt=DOL)
    r += 3

    ex.cell(r, 1, "OVERSPEND — BB DA exceeds allowance (ranked)").font = Font(bold=True, size=12, name="Calibri")
    r += 1
    ov_cols = ["Region", "MMD", "Description", "Band", "Cases", "BB DA/Cs", "Allowed/Cs", "Over $", "Matched Program Tier"]
    for i, h in enumerate(ov_cols): C(ex, r, 1 + i, h, bold=True)
    r += 1
    ov_tot = 0
    for _, row in over_df.iterrows():
        for i, k in enumerate(ov_cols):
            fmt = DOL if k in ("Over $", "Band") else (DOL2 if k in ("BB DA/Cs", "Allowed/Cs") else (NUM if k == "Cases" else None))
            C(ex, r, 1 + i, row[k], fmt=fmt, align="left" if k in ("Description", "Matched Program Tier") else "center")
        ov_tot += row["Over $"]; r += 1
    C(ex, r, 1, "Subtotal", bold=True, align="left")
    C(ex, r, 8, ov_tot, bold=True, fmt=DOL)
    r += 1
    C(ex, r, 1, "TOTAL EXCEPTIONS", bold=True, sz=12, align="left")
    C(ex, r, 8, nip_tot + ov_tot, bold=True, sz=12, fmt=DOL)

    ex.column_dimensions["A"].width = 9
    ex.column_dimensions["B"].width = 13
    ex.column_dimensions["C"].width = 30
    for c in range(4, 9): ex.column_dimensions[get_column_letter(c)].width = 12
    ex.column_dimensions["I"].width = 28

    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.getvalue()


def _build_multi_workbook(brand_results, grand_summary, coverage_rows):
    """Stacked multi-brand workbook: Global (Brand -> Product -> Deal Level),
    Exceptions (per brand), and a Coverage sheet."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    thin = Side(style="thin", color="000000")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    DOL = '"$"#,##0'; DOL2 = '"$"#,##0.00'; NUM = "#,##0.00"; PCT = "0.0%"

    def C(ws, r, c, v, bold=False, sz=11, fmt=None, align="center"):
        cell = ws.cell(r, c, v)
        cell.font = Font(bold=bold, size=sz, name="Calibri")
        cell.alignment = Alignment(horizontal=align, vertical="center")
        if fmt: cell.number_format = fmt
        cell.border = box
        return cell

    # ---- Global ----
    ws = wb.active; ws.title = "Global"; ws.sheet_view.showGridLines = False
    ws.cell(1, 1, "DA Reconciliation — Multi-Brand Program Analysis").font = Font(bold=True, size=14, name="Calibri")
    card = [("Total Cases", grand_summary["cases"], NUM), ("DA Spent", grand_summary["spent"], DOL),
            ("Program Allowance", grand_summary["allowance"], DOL),
            ("Over / Under", grand_summary["over_under"], DOL), ("Dist GP %", grand_summary["gp_pct"], PCT)]
    for i, (lab, val, fmt) in enumerate(card):
        a_ = ws.cell(3, 1 + i, lab); a_.font = Font(bold=True, size=10, name="Calibri")
        a_.alignment = Alignment(horizontal="center", wrap_text=True); a_.border = box
        v_ = ws.cell(4, 1 + i, val); v_.font = Font(bold=True, size=13, name="Calibri")
        v_.alignment = Alignment(horizontal="center"); v_.number_format = fmt; v_.border = box

    HDRS = ["Product / Deal Level", "Cases", "Sell Price Case", "Sell Price Bottle",
            "Dollar Profit Case", "DA Spent", "Program Allowance", "Over/Under", "GP %"]
    keys = ["Cases", "Sell Price Case", "Sell Price Bottle", "Dollar Profit Case",
            "DA Spent", "Program Allowance", "Over/Under", "GP %"]
    FMT = [NUM, DOL2, DOL2, DOL2, DOL, DOL, DOL, PCT]

    from openpyxl.styles import PatternFill
    NAVY = PatternFill("solid", start_color="1F4E78", end_color="1F4E78")
    r = 6
    for br in brand_results:
        # Brand banner
        bc = ws.cell(r, 1, f"{br['display']}  ({br['prefix']})")
        bc.font = Font(bold=True, size=13, color="FFFFFF", name="Calibri")
        for c in range(1, 10):
            ws.cell(r, c).fill = NAVY; ws.cell(r, c).border = box
        r += 1
        for i, h in enumerate(HDRS):
            C(ws, r, 1 + i, h, bold=True)
        r += 1
        gdf = br["global_df"]
        cur_product = None
        for _, row in gdf.iterrows():
            if row["Product"] != cur_product:
                cur_product = row["Product"]
                pc = ws.cell(r, 1, cur_product); pc.font = Font(bold=True, size=11, name="Calibri")
                pc.alignment = Alignment(horizontal="left"); pc.border = box
                for c in range(2, 10): C(ws, r, c, "")
                r += 1
            ws.cell(r, 1, "   " + row["Deal Level"]).font = Font(size=11, name="Calibri")
            ws.cell(r, 1).alignment = Alignment(horizontal="left"); ws.cell(r, 1).border = box
            for i, k in enumerate(keys):
                C(ws, r, 2 + i, row[k], fmt=FMT[i])
            r += 1
        # brand subtotal
        s = br["summary"]
        tl = ws.cell(r, 1, f"{br['display']} TOTAL"); tl.font = Font(bold=True, size=11, name="Calibri")
        tl.alignment = Alignment(horizontal="left"); tl.border = box
        C(ws, r, 2, round(s["cases"], 2), bold=True, fmt=NUM)
        for c in (3, 4, 5): C(ws, r, c, "")
        C(ws, r, 6, s["spent"], bold=True, fmt=DOL)
        C(ws, r, 7, s["allowance"], bold=True, fmt=DOL)
        C(ws, r, 8, s["over_under"], bold=True, fmt=DOL)
        C(ws, r, 9, s["gp_pct"], bold=True, fmt=PCT)
        r += 2

    # grand total
    gl = ws.cell(r, 1, "GRAND TOTAL (all brands)"); gl.font = Font(bold=True, size=12, name="Calibri")
    gl.alignment = Alignment(horizontal="left"); gl.border = box
    C(ws, r, 2, round(grand_summary["cases"], 2), bold=True, sz=12, fmt=NUM)
    for c in (3, 4, 5): C(ws, r, c, "")
    C(ws, r, 6, grand_summary["spent"], bold=True, sz=12, fmt=DOL)
    C(ws, r, 7, grand_summary["allowance"], bold=True, sz=12, fmt=DOL)
    C(ws, r, 8, grand_summary["over_under"], bold=True, sz=12, fmt=DOL)
    C(ws, r, 9, grand_summary["gp_pct"], bold=True, sz=12, fmt=PCT)

    ws.column_dimensions["A"].width = 28
    for c in range(2, 10): ws.column_dimensions[get_column_letter(c)].width = 15

    # ---- Coverage sheet ----
    cov = wb.create_sheet("Coverage")
    cov.sheet_view.showGridLines = False
    cov.cell(1, 1, "Coverage Check — did every bill-back dollar get audited?").font = Font(bold=True, size=13, name="Calibri")
    ch = ["Brand", "Status", "Parsed $", "Audited $", "Unmatched $"]
    for i, h in enumerate(ch): C(cov, 3, 1 + i, h, bold=True)
    rr = 4
    for row in coverage_rows:
        C(cov, rr, 1, row["Brand"], align="left")
        C(cov, rr, 2, row["Status"], align="left")
        C(cov, rr, 3, row["Parsed $"], fmt=DOL)
        C(cov, rr, 4, row["Audited $"], fmt=DOL)
        C(cov, rr, 5, row["Unmatched $"], fmt=DOL)
        rr += 1
    cov.column_dimensions["A"].width = 24
    cov.column_dimensions["B"].width = 30
    for c in (3, 4, 5): cov.column_dimensions[get_column_letter(c)].width = 14

    # ---- Exceptions sheet (overspend, all brands stacked) ----
    ex = wb.create_sheet("Exceptions"); ex.sheet_view.showGridLines = False
    ex.cell(1, 1, "Overspend Exceptions — biggest first, per brand").font = Font(bold=True, size=13, name="Calibri")
    r = 3
    ocols = ["Region", "MMD", "Description", "Band", "Cases", "BB DA/Cs", "Allowed/Cs", "Over $", "Matched Program Tier"]
    for br in brand_results:
        odf = br["overspend_df"]
        if odf is None or len(odf) == 0:
            continue
        bc = ex.cell(r, 1, f"{br['display']} ({br['prefix']})")
        bc.font = Font(bold=True, size=12, name="Calibri"); r += 1
        for i, h in enumerate(ocols): C(ex, r, 1 + i, h, bold=True)
        r += 1
        for _, row in odf.iterrows():
            for i, k in enumerate(ocols):
                fmt = DOL if k in ("Over $", "Band") else (DOL2 if k in ("BB DA/Cs", "Allowed/Cs") else (NUM if k == "Cases" else None))
                C(ex, r, 1 + i, row[k], fmt=fmt, align="left" if k in ("Description", "Matched Program Tier") else "center")
            r += 1
        r += 1

    # ---- Needs Review: no tier matched + not in program ----
    flagcols = ["Brand", "Region", "MMD", "Description", "Band", "Cases", "BB DA/Cs", "DA Charged $", "Flag"]
    review_rows = []
    for br in brand_results:
        for key in ("no_tier_df", "not_in_program_df"):
            d = br.get(key)
            if d is None or len(d) == 0:
                continue
            for _, row in d.iterrows():
                rr = {"Brand": br["display"]}
                rr.update({c: row[c] for c in flagcols if c in row})
                review_rows.append(rr)
    if review_rows:
        r += 1
        h = ex.cell(r, 1, "NEEDS REVIEW — DA charged but allowance could not be determined")
        h.font = Font(bold=True, size=12, name="Calibri"); r += 1
        sub = ex.cell(r, 1, "(These are NOT counted as confirmed overspend. Allowance unknown — verify the program tier or product mapping.)")
        sub.font = Font(italic=True, size=9, name="Calibri"); r += 1
        for i, hh in enumerate(flagcols): C(ex, r, 1 + i, hh, bold=True)
        r += 1
        for row in review_rows:
            for i, k in enumerate(flagcols):
                v = row.get(k, "")
                fmt = DOL if k in ("DA Charged $", "Band") else (DOL2 if k == "BB DA/Cs" else (NUM if k == "Cases" else None))
                C(ex, r, 1 + i, v, fmt=fmt, align="left" if k in ("Description", "Flag", "Brand") else "center")
            r += 1

    ex.column_dimensions["A"].width = 14
    ex.column_dimensions["B"].width = 9
    ex.column_dimensions["C"].width = 13
    ex.column_dimensions["D"].width = 30
    for c in range(5, 9): ex.column_dimensions[get_column_letter(c)].width = 12
    ex.column_dimensions["I"].width = 42

    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.getvalue()


# ============ SPEND SUMMARY (bill-back only) ============
def _read_bb_header(pdf_path):
    """Pull invoice date and market from a bill-back PDF header."""
    import pdfplumber, re
    from datetime import datetime
    date_val = None; market = None
    try:
        with pdfplumber.open(pdf_path) as pdf:
            txt = pdf.pages[0].extract_text() or ""
    except Exception:
        return None, None
    m = re.search(r"Invoice date:\s*([0-9]{1,2}/[0-9]{1,2}/[0-9]{4})", txt)
    if m:
        try: date_val = datetime.strptime(m.group(1), "%m/%d/%Y").date()
        except Exception: date_val = None
    # market from "State : SGWS / North California" or "SGWS South California"
    ms = re.search(r"State\s*:\s*(.+)", txt)
    if ms:
        raw = ms.group(1).split("\n")[0].strip()
        up = raw.upper()
        if "NORTH" in up: market = "NCA"
        elif "SOUTH" in up: market = "SCA"
        else: market = raw
    return date_val, market


def spend_summary(bb_files, deal_xlsx, selected_prefixes, date_from=None, date_to=None):
    """Bill-back-only per-product spend.
    bb_files: list of pdf paths. Auto-reads invoice date + market from each.
    Filters to date range. Returns per-product totals across all prefixes."""
    import brands as BR
    import os
    from collections import defaultdict
    material_to_mmd = {}
    if deal_xlsx:
        for p in selected_prefixes:
            try:
                material_to_mmd.update(A.build_material_to_mmd_mapping(deal_xlsx, brand_prefix=p))
            except Exception:
                pass

    file_meta = []   # (path, date, market)
    for path in bb_files:
        d, mk = _read_bb_header(path)
        file_meta.append((path, d, mk))

    agg = defaultdict(lambda: {"cs": 0, "da_w": 0, "price_w": 0, "spent": 0,
                               "prefix": "", "product": ""})
    used_dates = []
    for path, d, mk in file_meta:
        if date_from and d and d < date_from: continue
        if date_to and d and d > date_to: continue
        if d: used_dates.append(d)
        for prefix in selected_prefixes:
            try:
                lines = A.parse_billback_pdf(path, region=(mk or "NCA"),
                                             brand_prefix=prefix,
                                             material_to_mmd=material_to_mmd)
            except Exception:
                continue
            for b in lines:
                if b["cases"] <= 0:
                    continue
                mmd = b.get("mmd") or material_to_mmd.get(b["mat"])
                label = BR.product_name(prefix, mmd, b["desc"])
                key = (prefix, label)
                s = agg[key]
                s["prefix"] = prefix; s["product"] = label
                s["cs"] += b["cases"]
                s["da_w"] += b["da"] * b["cases"]
                s["price_w"] += (b["band"] or 0) * b["cases"]
                s["spent"] += b["tot"]

    rows = []
    for (prefix, label), s in sorted(agg.items(), key=lambda kv: -kv[1]["spent"]):
        cs = s["cs"]
        rows.append({
            "Product": label,
            "Total Cases": round(cs, 2),
            "Total DA Spent": round(s["spent"], 0),
            "Avg DA / Case": round(s["da_w"] / cs, 2) if cs else 0,
            "Avg Price / Case": round(s["price_w"] / cs, 2) if cs else 0,
        })
    df = pd.DataFrame(rows)
    meta = {
        "files": [(os.path.basename(p), str(d) if d else "?", mk or "?") for p, d, mk in file_meta],
        "date_min": min(used_dates) if used_dates else None,
        "date_max": max(used_dates) if used_dates else None,
        "total_spent": round(sum(s["spent"] for s in agg.values()), 0),
        "total_cases": round(sum(s["cs"] for s in agg.values()), 2),
    }
    return df, meta


# ============ FORMAT B: Monthly List of Discount Participation ============
def parse_participation_pdf(path):
    """Parse a 'Monthly List of Discount Participation' PDF (line-item format).

    Market-agnostic: works across all SGWS markets (Texas, California, etc.).
    The line layout is invariant across markets (same SAP program ZMRDARP0_V2);
    only cosmetic details differ. We anchor on structure, not market text.

    Returns (lines, month_date, market, stated_totals) where stated_totals also
    carries the file's own grand totals so callers can verify the parse to the
    penny against the PDF's printed 'TOTAL PARTICIPATION' line.
    """
    import pdfplumber, re
    from datetime import date
    with pdfplumber.open(path) as pdf:
        full = "\n".join((p.extract_text() or "") for p in pdf.pages)

    # ---- Invoice date: "Ending MM/DD/YYYY" ----
    md = re.search(r"Ending\s+(\d{2}/\d{2}/\d{4})", full)
    month_date = None
    if md:
        mm, dd, yy = md.group(1).split("/")
        month_date = date(int(yy), int(mm), int(dd))

    # ---- Market: fully generic. Read whatever follows "SGWS". ----
    # Handles "SGWS of Texas", "SGWS Florida", "SGWS of New York",
    # "North California", "SGWS California South", and the older
    # "State : SGWS / North California".
    market = "?"
    m = re.search(r"(North|South)\s+California", full, re.I) or \
        re.search(r"California\s+(North|South)", full, re.I)
    if m:
        market = "NCA" if m.group(1).upper() == "NORTH" else "SCA"
    else:
        m = re.search(r"SGWS\s+(?:of\s+)?([A-Za-z][A-Za-z ]*?)\s+(?:Date|Page|Monthly|\n|$)", full)
        if m:
            market = m.group(1).strip()
        else:
            m = re.search(r"State\s*:\s*SGWS\s*/?\s*([^\n]+)", full)
            if m:
                raw = m.group(1).strip().upper()
                market = "NCA" if "NORTH" in raw else ("SCA" if "SOUTH" in raw else m.group(1).strip())

    # ---- Line items ----
    # Anchored on invariant structure:
    #   material(7) group(NNNNN(WORD)) hierarchy(digits) DESC SIZE BPC
    #   price "Less than" price cases part-level extended  [trailing program cols ignored]
    # Notes handled across markets:
    #   - an optional 4-digit group-number column may precede the NNNNN(WORD) group (Florida)
    #   - size token may be jammed onto the description with no space (\s* before size)
    #   - "Less than" may have no space before a >=1000 price (Less than1,194.00)
    #   - values may be negative with a trailing '-' (18.00-)
    #   - trailing program/recovery columns vary (2000001014 / X200000101 PRIORMONTH / SW1) -> ignored
    rx = re.compile(
        r"^(\d{7})\s+(?:\d{3,4}\s+)?\d+\([A-Z]+\)\s+\d+\s+(.+?)\s*(\d+(?:\.\d+)?(?:ML|LT))\s+(\d+)\s*"
        r"([\d,]+\.\d{2})\s+Less than\s*([\d,]+\.\d{2})\s+([\d,]+\.\d{2}-?)\s+"
        r"([\d,]+\.\d{2})\s+([\d,]+\.\d{2}-?)(?:\s|$)")
    lines = []
    for ln in full.split("\n"):
        m = rx.match(ln)
        if not m:
            continue
        mat, desc, size, bpc, casedisc, lessthan, cases_p, partlvl, ext = m.groups()
        cases = float(cases_p.replace(",", "").rstrip("-")) * (-1 if cases_p.endswith("-") else 1)
        ext_amt = float(ext.replace(",", "").rstrip("-")) * (-1 if ext.endswith("-") else 1)
        per_case = float(partlvl.replace(",", ""))
        lines.append({"mat": mat, "desc": desc.strip(), "size": size,
                      "cases": cases, "price": float(lessthan.replace(",", "")),
                      "da_per_case": per_case, "ext": ext_amt})

    # ---- Stated answer key ----
    # Per-group totals (the boxed "TOTALS BY MATERIAL GROUP" section)
    stated = {}
    for m in re.finditer(r"\|([A-Z ]+?)\s+(\d{3,6})\s+([\d,]+\.\d{2}-?)\s+([\d,]+\.\d{2}-?)\s*\|", full):
        val = float(m.group(4).replace(",", "").rstrip("-")) * (-1 if m.group(4).endswith("-") else 1)
        stated[m.group(1).strip()] = val
    # File grand totals from "TOTAL PARTICIPATION: <cases> <ext>" — the definitive check
    gt = re.search(r"TOTAL PARTICIPATION:\s+([\d,]+\.\d{2}-?)\s+([\d,]+\.\d{2}-?)", full)
    if gt:
        gc = float(gt.group(1).replace(",", "").rstrip("-")) * (-1 if gt.group(1).endswith("-") else 1)
        ge = float(gt.group(2).replace(",", "").rstrip("-")) * (-1 if gt.group(2).endswith("-") else 1)
        stated["__GRAND_CASES__"] = gc
        stated["__GRAND_EXT__"] = ge
    return lines, month_date, market, stated


# Material-group name -> brand prefix (from the PDF's TOTALS BY MATERIAL GROUP)
_GROUP_TO_PREFIX = {
    "LOUIS ROEDERER": "LR", "ROEDERER ESTATE": "RE", "SCHARFFENBERGER": "SC",
    "DOMAINES OTT": "OT", "BY OTT": "OT", "MARQUES DE MURR": "MU", "NAPANOOK": "DE",
    "OTHELLO": "DE", "DOMINUS": "DE", "LIVIO FELLUGA": "LF", "DALMAU": "MU",
    "CAPELLANIA": "MU",
}


def spend_summary_participation(pdf_paths, date_from=None, date_to=None,
                                selected_prefixes=None):
    """Bill-back spend summary from Format B participation PDFs.
    Groups by brand and sub-group (per brand_hierarchy.py), per product,
    with brand totals. Uses description matching to map to brands.
    Verifies each file against its stated group totals."""
    import brands as BR
    from brand_hierarchy import get_brand_and_subgroup
    from collections import defaultdict
    from datetime import date as _date
    import os

    files_meta = []
    # brand_label -> sub_group -> product_label -> agg
    agg = defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: {"cs": 0, "da_w": 0, "price_w": 0, "ext": 0})))
    used_dates = []
    unmapped = []
    stated_groups_per_file = []  # per-file PDF material-group answer keys (in-range only)

    for path in pdf_paths:
        lines, mdate, market, stated = parse_participation_pdf(path)
        parsed_tot = sum(l["ext"] for l in lines)
        # Verify against the PDF's own printed grand total ("TOTAL PARTICIPATION").
        # Fall back to summing the per-group answer key if the grand line is absent.
        if "__GRAND_EXT__" in stated:
            stated_tot = stated["__GRAND_EXT__"]
        else:
            stated_tot = sum(v for k, v in stated.items() if not k.startswith("__"))
        files_meta.append((os.path.basename(path), str(mdate) if mdate else "?", market,
                           parsed_tot, stated_tot))
        if date_from and mdate and mdate < date_from:
            continue
        if date_to and mdate and mdate > date_to:
            continue
        if mdate:
            used_dates.append(mdate)
        stated_groups_per_file.append({k: v for k, v in stated.items() if not k.startswith("__")})
        for l in lines:
            # Use the hierarchy to map description -> brand/sub-group
            # (MMD ID is used as a fallback, but descriptions are the primary key)
            brand_label, sub_label = get_brand_and_subgroup(l["mat"], l["desc"])

            if brand_label.startswith("⚠"):
                unmapped.append({"mat": l["mat"], "desc": l["desc"], "cases": l["cases"], "ext": l["ext"]})
                continue

            # Treat sub-groups (By.Ott, Ott Crus, Mongrana, etc.) as their own
            # top-level brands. Brands without a sub-group keep their own name.
            effective_brand = sub_label or brand_label

            plabel = f"{l['desc']} {l['size']}"
            s = agg[effective_brand][""][plabel]
            s["cs"] += l["cases"]
            s["da_w"] += l["da_per_case"] * l["cases"]
            s["price_w"] += l["price"] * l["cases"]
            s["ext"] += l["ext"]

    # build rows: brand -> products -> brand total (sub-groups now ARE brands)
    rows = []
    for blabel in agg.keys():
        brand_subs = agg[blabel]
        brand_cs = brand_ext = 0
        sub_rows = []

        for sub_label in sorted(brand_subs.keys()):
            products = brand_subs[sub_label]
            sub_cs = sub_ext = 0
            prods = []
            for plabel, s in sorted(products.items(), key=lambda kv: -kv[1]["ext"]):
                cs = s["cs"]
                prods.append({
                    "Product": plabel, "Cases": round(cs, 2),
                    "Total DA Spent": round(s["ext"], 2),
                    "Avg DA / Case": round(s["da_w"] / cs, 2) if cs else 0,
                    "Avg Price / Case": round(s["price_w"] / cs, 2) if cs else 0,
                })
                sub_cs += cs; sub_ext += s["ext"]
            sub_rows.append({"sub_label": sub_label or blabel, "products": prods,
                            "sub_cases": round(sub_cs, 2), "sub_total": round(sub_ext, 2)})
            brand_cs += sub_cs; brand_ext += sub_ext

        rows.append({"brand": blabel, "sub_groups": sub_rows,
                     "brand_cases": round(brand_cs, 2), "brand_total": round(brand_ext, 2)})

    # Sort brands by total cases sold, descending
    rows.sort(key=lambda r: -r["brand_cases"])

    # ---- Group-level accuracy check ----
    # Beyond the file-level total check, verify each app-brand's parsed total
    # against the PDF's printed "TOTALS BY MATERIAL GROUP" answer key.
    # The PDF groups don't map 1:1 to app brands (e.g. the single PDF group
    # "LOUIS ROEDERER" covers Collection + VRBB + Cristal), so we map each
    # app brand to the PDF group name(s) it draws from, then compare sums.
    # APP_BRAND -> set of PDF material-group name prefixes that feed it.
    _BRAND_TO_PDFGROUPS = {
        "Roederer Estate": ["ROEDERER ESTATE"],
        "Roederer Collection": ["LOUIS ROEDERER"],
        "Roederer VRBB": ["LOUIS ROEDERER"],
        "Roederer Cristal": ["LOUIS ROEDERER"],
        "By.Ott": ["BY OTT"],
        "Ott Crus": ["DOMAINES OTT"],
        "De Ladoucette": ["LADOUCETTE", "COMTE LAFOND", "LES DEUX TOURS", "MARC BREDIF", "LA POUSSIE"],
        "Marqués de Murrieta": ["MARQUES DE MURR", "DALMAU", "CAPELLANIA", "PAZO BARRANTES"],
        "Dominus": ["NAPANOOK", "OTHELLO", "DOMINUS"],
        "Bon Vivant": ["BON VIVANT"],
        "Inniskillin Icewine": ["INNISKILLIN"],
        "Jackson Triggs Niagara": ["JACKSON TRIGGS"],
        "Distell": ["FLEUR DU CAP", "DISTELL"],
        "Ets. J-P Moueix": ["CHATEAU DE SALE", "CHATEAU PEYMOUT", "CHATEAU PEYMOU"],
        "Château Loudenne": ["CHATEAU LOUDENN"],
        "CLR Bordeaux": ["CHATEAU DE PEZ"],
        "Castiglion del Bosco": ["CASTIGLION DEL"],
        "Querciabella": ["QUERCIABELLA"],
        "Mongrana": ["QUERCIABELLA"],
        "Merry Edwards": ["MERRY EDWARDS"],
        "Domaine Anderson": ["DOMAINE ANDERSO"],
    }
    # Sum the stated per-group answer key across all in-range files.
    stated_groups_all = {}
    for sg in stated_groups_per_file:
        for gname, gval in sg.items():
            if gname.startswith("__"):
                continue
            stated_groups_all[gname] = stated_groups_all.get(gname, 0) + gval

    def _stated_for_prefixes(prefixes):
        """Sum stated group totals whose names match any of the prefixes."""
        total = 0.0
        matched_any = False
        for gname, gval in stated_groups_all.items():
            for p in prefixes:
                if gname.startswith(p) or p.startswith(gname):
                    total += gval
                    matched_any = True
                    break
        return (total if matched_any else None)

    # Auto-detect which PDF groups are shared by 2+ app brands. Any such group
    # must be reconciled as a COMBINED unit (sum of all app brands that draw
    # from it), because the PDF prints a single total for the whole group.
    # Examples: "LOUIS ROEDERER" feeds Collection+VRBB+Cristal; "QUERCIABELLA"
    # feeds both Querciabella and Mongrana. This is derived from the mapping
    # above, so future splits are handled with no code change.
    _group_to_brands = {}
    for brand_name, prefixes in _BRAND_TO_PDFGROUPS.items():
        for p in prefixes:
            _group_to_brands.setdefault(p, []).append(brand_name)
    _SHARED_GROUP_BRANDS = {gp: bl for gp, bl in _group_to_brands.items() if len(bl) > 1}

    group_check = []
    handled_brands = set()

    # First, the shared-group reconciliations (sum the sharing app brands).
    for gprefix, brand_list in _SHARED_GROUP_BRANDS.items():
        parsed_val = sum(r["brand_total"] for r in rows if r["brand"] in brand_list)
        stated_val = _stated_for_prefixes([gprefix])
        handled_brands.update(brand_list)
        if stated_val is None:
            continue
        gap = round(parsed_val - stated_val, 2)
        group_check.append({"group": gprefix, "parsed": round(parsed_val, 2),
                            "stated": round(stated_val, 2), "gap": gap, "ok": abs(gap) < 0.50})

    # Then each remaining app brand vs the sum of all PDF groups that feed it.
    for r in rows:
        if r["brand"] in handled_brands:
            continue
        prefixes = _BRAND_TO_PDFGROUPS.get(r["brand"])
        if not prefixes:
            continue
        stated_val = _stated_for_prefixes(prefixes)
        if stated_val is None:
            continue
        gap = round(r["brand_total"] - stated_val, 2)
        group_check.append({"group": r["brand"], "parsed": round(r["brand_total"], 2),
                            "stated": round(stated_val, 2), "gap": gap, "ok": abs(gap) < 0.50})

    meta = {
        "files": files_meta,
        "date_min": min(used_dates) if used_dates else None,
        "date_max": max(used_dates) if used_dates else None,
        "grand_total": round(sum(b["brand_total"] for b in rows), 2),
        "grand_cases": round(sum(b["brand_cases"] for b in rows), 2),
        "unmapped_lines": unmapped,  # Flag unmapped for user review
        "group_check": group_check,  # Per-PDF-group accuracy reconciliation
    }
    return rows, meta


# ============================================================
# MARKET-AWARE RECONCILIATION  (catalog-driven, CA + TX)
# ============================================================
# Resolves material->MMD via the master SGWS item catalog (all markets),
# groups each line by the existing brand_hierarchy rules (by description),
# and audits against the right authorization source per market:
#   CA  -> program file (closest-PTA tiers)
#   TX  -> deal file's month-active tiers (deal file IS the authorization)
# Premise split comes from the depletion/VIP report (CA or TX layout).
def run_market_reconciliation(bb_inputs, catalog_xlsx, vip_xlsx,
                              authorization_xlsx, market_kind,
                              recon_month="April", deal_xlsx=None):
    """
    bb_inputs:          list of (pdf_path, region)
    catalog_xlsx:       master SGWS item catalog (Item#->MMD, all markets)
    vip_xlsx:           depletion/VIP report (CA or TX layout)
    authorization_xlsx: CA program file OR TX deal file
    market_kind:        "CA" or "TX" (selects authorization handling)
    Returns the same shape as run_multi so app.py can render it unchanged.
    """
    import market_inputs as MI
    import audit_unified as AU
    import brand_hierarchy as BH
    import da_audit as A
    from collections import defaultdict

    crosswalk = MI.build_master_crosswalk(catalog_xlsx)

    # Parse all bill-back lines (all brands, all regions)
    bb_lines = []
    for pdf_path, region in bb_inputs:
        bb_lines.extend(A.parse_billback_pdf(pdf_path, region=region))

    # Read depletion/VIP (premise split)
    vip_rows = MI.read_vip_any(vip_xlsx) if vip_xlsx else []

    # Build tiers_by_mmd from the authorization/deal file(s). Accept a single
    # path or a list of paths (multiple program files merged). Auto-detect each
    # file's format (TX deal-style vs CA program-style).
    auth_list = authorization_xlsx if isinstance(authorization_xlsx, (list, tuple)) else [authorization_xlsx]
    tiers_by_mmd = {}
    for auth_path in auth_list:
        if not auth_path:
            continue
        if _detect_auth_format(auth_path) == "TX":
            part = MI.read_deal_tiers_tx(auth_path, month=recon_month, crosswalk=crosswalk)
        else:
            part = _ca_program_to_tiers(auth_path)
        for mmd, tiers in part.items():
            tiers_by_mmd.setdefault(mmd, []).extend(tiers)

    audit = AU.run_audit_unified(bb_lines, vip_rows, tiers_by_mmd, crosswalk,
                                 default_premise="OFF")

    # Read deal economics (for Sell Price / GP%) if a deal file is available.
    deals_by_mmd = {}
    if deal_xlsx:
        try:
            deals_by_mmd = _deals_by_mmd_any(deal_xlsx, crosswalk)
        except Exception:
            deals_by_mmd = {}

    import brands as BR

    # Which MMDs does the uploaded program cover? Used to scope the report to
    # the brands you actually uploaded a program for (e.g. upload Murrieta →
    # see Murrieta), instead of every brand on the bill-back.
    program_mmds = set(tiers_by_mmd.keys())

    # Group each audited line by the existing hierarchy (by description),
    # then build the SAME global_df / overspend_df shape the page already renders.
    brand_lines = defaultdict(list)
    for a in audit:
        # Only include lines whose product is covered by an uploaded program.
        if a.get("mmd") not in program_mmds:
            continue
        brand, sub = BH.get_brand_and_subgroup(a.get("mmd"), a["bb"]["desc"])
        brand_lines[brand].append(a)

    brand_results = []
    coverage_rows = []
    grand = {"cases": 0.0, "spent": 0.0, "allowance": 0.0, "gn": 0.0, "gd": 0.0}

    for brand in sorted(brand_lines, key=lambda b: -sum(x["bb"]["tot"] for x in brand_lines[b])):
        block = _build_brand_block(brand, brand_lines[brand], deals_by_mmd)
        brand_results.append(block)
        s = block["summary"]
        grand["cases"] += s["cases"]; grand["spent"] += s["spent"]
        grand["allowance"] += s["allowance"]; grand["gn"] += block["gn"]; grand["gd"] += block["gd"]
        coverage_rows.append({
            "Brand": block["display"], "Status": "OK",
            "Parsed $": round(s["spent"], 0),
            "Audited $": round(s["spent"], 0),
            "Unmatched $": 0,
        })

    grand_summary = {
        "cases": grand["cases"], "spent": grand["spent"],
        "allowance": grand["allowance"],
        "over_under": grand["spent"] - grand["allowance"],
        "gp_pct": grand["gn"] / grand["gd"] if grand["gd"] else 0,
    }
    return {
        "brand_results": brand_results, "grand_summary": grand_summary,
        "coverage_rows": coverage_rows, "workbook_bytes": None,
    }


def _build_brand_block(display, lines, deals_by_mmd):
    """Build the original per-brand block: global_df (Product x Deal Level with
    Sell Price / Profit / GP%), overspend_df, not_in_program_df. Same shape the
    Reconciliation page and Excel export already expect."""
    import pandas as pd
    import brands as BR

    def deal_level(da_per_cs):
        if da_per_cs <= 20: return "Regular"
        if da_per_cs <= 40: return "Deep Deal"
        return "Aggressive"

    def allowed_rate(a):
        on = a["cases_by_premise"].get("ON", 0); off = a["cases_by_premise"].get("OFF", 0)
        ond = a["authorized_da_by_premise"].get("ON", 0); offd = a["authorized_da_by_premise"].get("OFF", 0)
        tp = on + off
        return (ond * on + offd * off) / tp if tp > 0 else 0

    P = defaultdict(lambda: defaultdict(lambda: {
        "label": "", "size": "", "cs": 0.0, "spent": 0.0, "allow": 0.0,
        "ncp_w": 0.0, "prof_w": 0.0}))

    for a in lines:
        bb = a["bb"]
        if bb["cases"] <= 0 or bb["tot"] <= 0:
            continue
        label = bb["desc"]
        b = deal_level(bb["da"])
        cs = bb["cases"]
        arate = allowed_rate(a)
        dl = _deal_for(deals_by_mmd, a.get("mmd"), bb["band"]) if deals_by_mmd else None
        ncp = dl["net_case_price"] if dl else 0
        spa = (dl.get("spa") or 0) if dl else 0
        ds = (dl.get("discount_support") or 0) if dl else 0
        li = (dl.get("adjusted_laid_in") or 0) if dl else 0
        profcs = ncp + spa + ds - li
        s = P[label][b]
        s["label"] = label; s["size"] = bb["size"]
        s["cs"] += cs; s["spent"] += bb["tot"]; s["allow"] += arate * cs
        s["ncp_w"] += ncp * cs; s["prof_w"] += profcs * cs

    rows = []
    TC = TS = TA = GN = GD = 0
    for label in sorted(P.keys()):
        for b in ["Regular", "Deep Deal", "Aggressive"]:
            s = P[label].get(b)
            if not s or s["cs"] == 0:
                continue
            cs = s["cs"]; ncp = s["ncp_w"] / cs if cs else 0
            bpc = BR.bottles_for(s["size"])
            rows.append({
                "Product": label, "Deal Level": b,
                "Cases": round(cs, 2),
                "Sell Price Case": round(ncp, 2),
                "Sell Price Bottle": round(ncp / bpc, 2) if bpc else 0,
                "Dollar Profit Case": round(s["prof_w"] / cs, 2) if cs else 0,
                "DA Spent": round(s["spent"], 0),
                "Program Allowance": round(s["allow"], 0),
                "Over/Under": round(s["spent"] - s["allow"], 0),
                "GP %": round(s["prof_w"] / s["ncp_w"], 4) if s["ncp_w"] else 0,
            })
            TC += cs; TS += s["spent"]; TA += s["allow"]; GN += s["prof_w"]; GD += s["ncp_w"]
    if rows:
        rows.append({
            "Product": f"— {display} subtotal —", "Deal Level": "",
            "Cases": round(TC, 2),
            "Sell Price Case": "", "Sell Price Bottle": "", "Dollar Profit Case": "",
            "DA Spent": round(TS, 0),
            "Program Allowance": round(TA, 0),
            "Over/Under": round(TS - TA, 0),
            "GP %": round(GN / GD, 4) if GD else 0,
        })
    global_df = pd.DataFrame(rows)

    def tier_str(a):
        on = a["matched_tier_by_premise"].get("ON"); off = a["matched_tier_by_premise"].get("OFF")
        parts = []
        if a["cases_by_premise"].get("OFF", 0) > 0 and off: parts.append(f"{off.get('tier','deal')} (OFF)")
        if a["cases_by_premise"].get("ON", 0) > 0 and on: parts.append(f"{on.get('tier','deal')} (ON)")
        return ", ".join(parts) if parts else "—"

    over = sorted([a for a in lines if a["overage_total"] > 0.5 and not a.get("not_in_program") and a["bb"]["cases"] > 0],
                  key=lambda a: -a["overage_total"])
    overspend_df = pd.DataFrame([{
        "Region": a["bb"]["region"], "MMD": a.get("mmd"), "Description": a["bb"]["desc"],
        "Band": a["bb"]["band"], "Cases": round(a["bb"]["cases"], 2),
        "BB DA/Cs": a["bb"]["da"], "Allowed/Cs": round(allowed_rate(a), 2),
        "Over $": round(a["overage_total"], 0), "Matched Program Tier": tier_str(a),
    } for a in over])

    nip = [a for a in lines if a.get("not_in_program") and a["bb"]["cases"] > 0]
    not_in_program_df = pd.DataFrame([{
        "Region": a["bb"]["region"], "MMD": a.get("mmd"), "Description": a["bb"]["desc"],
        "Band": a["bb"]["band"], "Cases": round(a["bb"]["cases"], 2),
        "BB DA/Cs": a["bb"]["da"], "DA Charged $": round(a["bb"]["tot"], 0),
        "Flag": a.get("note", "MMD not in program file"),
    } for a in nip])

    summary = {"cases": TC, "spent": TS, "allowance": TA,
               "over_under": TS - TA, "gp_pct": GN / GD if GD else 0}
    return {
        "prefix": "", "display": display, "summary": summary,
        "global_df": global_df, "overspend_df": overspend_df,
        "not_in_program_df": not_in_program_df, "no_tier_df": pd.DataFrame(),
        "gn": GN, "gd": GD,
        "parsed_total": TS, "unmatched_total": 0,
    }


def _deals_by_mmd_any(deal_xlsx, crosswalk):
    """Best-effort deal economics keyed by MMD, for the GP% columns. Tolerant of
    missing columns — returns {} silently if the file doesn't carry deal economics."""
    import da_audit as A
    from collections import defaultdict
    try:
        deals = A.extract_ca_deals(deal_xlsx, brand_prefix='') or []
    except Exception:
        return {}
    by_mmd = defaultdict(list)
    for d in deals:
        mmd = d.get("mmd_id")
        if mmd:
            by_mmd[mmd].append(d)
    return dict(by_mmd)


def _ca_program_to_tiers(program_xlsx):
    """Read a CA program file into {mmd: [{premise, pta, da}]} using da_audit's
    program parser + tier reader, so the unified audit can consume it.

    No brand-prefix filter: the parser defaults to 'RE', which would discard
    every block in an LR/OT/etc. program file. We want all blocks regardless of
    brand, since the catalog already establishes each line's identity.
    """
    import da_audit as A
    program = A.parse_program_file(program_xlsx, brand_prefix='')
    out = {}
    for mmd in program["mmd_to_blocks"].keys():
        tiers = []
        for prem in ("ON", "OFF"):
            for t in A.get_program_tiers(program, mmd, prem):
                tiers.append({"premise": prem, "pta": t["pta"], "da": t["da"]})
        if tiers:
            out[mmd] = tiers
    return out


def _detect_auth_format(path):
    """Return 'TX' if the file looks like a TX deal file (month columns + Premise),
    else 'CA' (program file)."""
    from openpyxl import load_workbook
    import warnings
    warnings.filterwarnings('ignore')
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb.active
    months = {'January','February','March','April','May','June','July',
              'August','September','October','November','December'}
    found_months = False
    found_premise = False
    for r in range(1, min(16, ws.max_row) + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if v is None:
                continue
            v = str(v).strip()
            if v in months:
                found_months = True
            if v in ('Premise', 'OnOff Premises'):
                found_premise = True
        if found_months and found_premise:
            break
    wb.close()
    return "TX" if (found_months and found_premise) else "CA"
