import datetime as dt
from pathlib import Path
from io import BytesIO
import base64
import html
import pandas as pd
import xlsxwriter

# Pandas is used for all source reading, transformation, aggregation and report data.
# XlsxWriter is used only by pandas ExcelWriter to create the final XLSX file.

BSNL_LOGO_B64 = """iVBORw0KGgoAAAANSUhEUgAAAUAAAACQCAMAAABOB0IDAAAC/VBMVEVHcEz7ax75ayH7ax77ax4HA437ax4HA40HA435ayH7ax77ax77ax4Eazj7ax77ax4HA40HA40HA437ax4HA43yaywHA43yayz7ax77ax77ax4HA40EazgEazgHA40HA43yayzyaywHA43yayzyaywHA40EazgHA43yayzyayz///8EazgHA43yayz///8EazgEazgEazgEazjyaywEazjyayzyayz////zbC7///////////8EazgEazgEazj///////////////////////8Eazj///+HbzcHA43yaywEazj////7ax7ybS/xbzLxcDPwdDrwdTvxbjHvekPugU7ybC3ugE3vfUjug1Hwdj3xcTXybC7xcTbuglDxcjbweEDvfknthlbwdz/thVXve0XweEHvfEfsjGDweULxczjuf0vth1jthFTsjmL0poHuhFPsil31pIDzp4PtiFryrY3uf0ztiVvybjD0o3zzqIbzdDnwdj7rkGXwcznsi1/1oHjrkWfyq4rsi17ve0b0fkbxcjjyr5HyspT2mGz2oXryuZ/1nXPsilz22s7tiFnwdz713NHys5byro/ysJLskmnzrIzrlGz1n3fytZrzxrLsj2T228/11sj1nnXyqojqlm/23tPywKrzvaTzy7nzsZP0zr70ybX12Mvzu6L0oXvyuJ31m3H0pH/veUPyxK/th1n2mW7yqYjwwq3ytJj0wqzzr5D51cXzekLyv6f1zbzql3Hytpvz0MH10sLufkrxyLbyuqH++vjyq4n3q4f1xa/01MX86eD97eb++PXrmXT85Nrqmnbwi1z0tZj1mnD32szwvab60r/98evqmHP5yLH72sr73tD4wKYTbjv6zbnpm3j74dX0kGL+9fH5tZYybTn0iFb4uZv3sZBhimExeUucnnz4vKDhbjFhbDXEcDbBqIlhlnHCiFucbTMgdkZadkSVf1CbrpLjtJnqoH47gVXBwavxmXCAmXXfg1PXsJRLh17d0MCpdD15g1fQ4deoEmT0AAAASHRSTlMAEDC/gEDvv4BAz2CfQCDfEO9gj8/vnxBQr3DfgL8wIIC/j59gUF1wzyDvEK/fEM+f7yBQ34+vmHBoMIAwj3C/3yDPQFCvr/s8OcONAAAbuklE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"""


# ---------------------------------------------------------------------------
# External master files
# ---------------------------------------------------------------------------
# Keep these two files in the GitHub repository root. They can be replaced
# whenever OLT/BBC/Manager/Area/Target information changes; no Python changes
# are required.
REPO_ROOT = Path(__file__).resolve().parent
OLT_MASTER_FILE = REPO_ROOT / "OLT_BBC_MAP.xlsx"
BBC_MASTER_FILE = REPO_ROOT / "BBC_Master.xlsx"

def _pick_col(df, aliases, required=True):
    lookup = {str(c).strip().upper().replace("_"," "): c for c in df.columns}
    for alias in aliases:
        key = alias.upper().replace("_"," ")
        if key in lookup:
            return lookup[key]
    if required:
        raise ValueError(f"Required column not found. Expected one of: {aliases}. Found: {list(df.columns)}")
    return None

def _read_master(path):
    if not path.exists():
        raise FileNotFoundError(
            f"Master file '{path.name}' is missing. Upload it to the GitHub repository "
            f"beside app.py/report_processor.py and redeploy Streamlit."
        )
    return pd.read_excel(path, sheet_name=0)

def _load_masters():
    olt = _read_master(OLT_MASTER_FILE)
    bbc = _read_master(BBC_MASTER_FILE)

    olt_ip_col = _pick_col(olt, ["OLT IP", "OLT_IP", "OLT", "IP"])
    olt_bbc_col = _pick_col(olt, ["BBC Name", "BBC", "BBM Name", "Employee", "BBC_NAME"])
    olt = olt[[olt_ip_col, olt_bbc_col]].rename(columns={olt_ip_col:"OLT IP", olt_bbc_col:"BBC Name"})
    olt["OLT IP"] = olt["OLT IP"].fillna("").astype(str).str.strip()
    olt["BBC Name"] = olt["BBC Name"].map(normalize_bbc_name)

    bbc_name_col = _pick_col(bbc, ["BBC Name", "BBC", "BBM Name", "Employee"])
    manager_col = _pick_col(bbc, ["DE / Manager", "DE/Manager", "Manager", "MT", "DE", "AGM/ Manager(MT)"])
    area_col = _pick_col(bbc, ["Area / TIP", "Area/TIP", "Area", "TIP"])
    target_col = _pick_col(bbc, ["Monthly Target", "BBC Target", "BBCTarget", "Target"])
    display_col = _pick_col(bbc, ["Display Name", "Employee Display Name", "Name"], required=False)
    manager_target_col = _pick_col(bbc, ["Manager Target", "DE Target", "MT Target"], required=False)
    order_col = _pick_col(bbc, ["S.No", "SNO", "Order"], required=False)

    bbc = bbc.rename(columns={
        bbc_name_col:"BBC Name", manager_col:"Manager", area_col:"Area", target_col:"BBCTarget"
    })
    if display_col:
        bbc = bbc.rename(columns={display_col:"Display Name"})
    else:
        bbc["Display Name"] = bbc["BBC Name"]
    if manager_target_col:
        bbc = bbc.rename(columns={manager_target_col:"ManagerTarget"})
    else:
        bbc["ManagerTarget"] = pd.to_numeric(bbc["BBCTarget"], errors="coerce").fillna(0)
    if order_col:
        bbc = bbc.rename(columns={order_col:"Order"})
    else:
        bbc["Order"] = range(1, len(bbc)+1)

    bbc["BBC Name"] = bbc["BBC Name"].map(normalize_bbc_name)
    bbc["Manager"] = bbc["Manager"].fillna("UNMAPPED").astype(str).str.strip()
    bbc["Area"] = bbc["Area"].fillna("UNMAPPED").astype(str).str.strip()
    bbc["BBCTarget"] = pd.to_numeric(bbc["BBCTarget"], errors="coerce").fillna(0).astype(int)
    bbc["ManagerTarget"] = pd.to_numeric(bbc["ManagerTarget"], errors="coerce").fillna(0).astype(int)
    bbc["Display Name"] = bbc["Display Name"].fillna(bbc["BBC Name"]).astype(str).str.strip()
    bbc = bbc.drop_duplicates("BBC Name", keep="last").sort_values("Order")
    olt_map = dict(zip(olt["OLT IP"], olt["BBC Name"]))
    bbc_info = {
        r["BBC Name"]: (r["Manager"], int(r["ManagerTarget"]), r["Area"], int(r["BBCTarget"]), r["Display Name"])
        for _, r in bbc.iterrows()
    }
    bbc_order = bbc["BBC Name"].tolist()
    return olt_map, bbc_info, bbc_order, olt, bbc

def normalize_bbc_name(v):
    return " ".join(str(v if pd.notna(v) else "").split())

CONN_NPC = "NPC"

CONN_RECON = "RECONNECTION"
CONN_CLSNP = "DUE TO NON PAYMENT (CLSNP)"
CONN_CLSVO = "VOLUNTAORY DISCONNECTION (CLSVO)"
CONN_OTHER = "OTHER"


def _parse_source_date(series):
    # Handles Excel serials, datetime objects and strings such as 16-AUG-2026.
    if pd.api.types.is_datetime64_any_dtype(series):
        return pd.to_datetime(series, errors="coerce").dt.normalize()
    out = pd.to_datetime(series, errors="coerce", dayfirst=True)
    return out.dt.normalize()


def _read_source(path):
    # The source export has column headings on Excel row 3 => header=2.
    df = pd.read_excel(path, sheet_name="Sheet0", header=2, engine="calamine")
    df = df.dropna(how="all").copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _classify(df, olt_map):
    required = ["BBC Name", "CLSR", "Ont Acquisition Type", "Disconnection reason",
                "Completion_Date", "Maintenance Franchisee", "OLT IP", "Order Id"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {missing}")

    df = df.copy()
    df["OLT IP"] = df["OLT IP"].fillna("").astype(str).str.strip()
    df["Order Id"] = df["Order Id"].fillna("").astype(str).str.strip()
    df["CLSR"] = df["CLSR"].fillna("").astype(str).str.strip().str.upper()
    df["Maintenance Franchisee"] = df["Maintenance Franchisee"].fillna("").astype(str).str.strip()

    # OLT IP -> canonical BBC mapping, with source BBC as fallback.
    df["BBC Name"] = df["OLT IP"].map(olt_map).fillna(df["BBC Name"].map(normalize_bbc_name))
    df["BBC Name"] = df["BBC Name"].map(normalize_bbc_name)

    oid = df["Order Id"].str[:5].str.upper()
    df["Connection Type"] = CONN_OTHER
    df.loc[(df["CLSR"] == "ACTIVE") & (oid == "BFBNC"), "Connection Type"] = CONN_NPC
    df.loc[(df["CLSR"] == "ACTIVE") & (oid != "BFBNC"), "Connection Type"] = CONN_RECON
    df.loc[(df["CLSR"] == "CLSD") & (oid == "BFBDI"), "Connection Type"] = CONN_CLSNP
    df.loc[(df["CLSR"] == "CLSV") & (oid == "BFBDI"), "Connection Type"] = CONN_CLSVO

    df["DATE"] = _parse_source_date(df["Completion_Date"])
    return df


def _aggregate(df):
    # Vectorized flags avoid row-by-row loops.
    g = df.assign(
        NPC=(df["Connection Type"] == CONN_NPC).astype(int),
        RECON=(df["Connection Type"] == CONN_RECON).astype(int),
        CLSVO=(df["Connection Type"] == CONN_CLSVO).astype(int),
        CLSNP=(df["Connection Type"] == CONN_CLSNP).astype(int),
    )
    g["PROVISION"] = g["NPC"] + g["RECON"]
    return g


def _bbc_report(df, report_date, olt_map, bbc_info, bbc_order):
    rows=[]
    for bbc in bbc_order:
        mgr, mgr_target, area, bbc_target, display_name = bbc_info[bbc]
        x=df[df["BBC Name"]==bbc]
        npc=int((x["Connection Type"]==CONN_NPC).sum())
        recon=int((x["Connection Type"]==CONN_RECON).sum())
        clsvo=int((x["Connection Type"]==CONN_CLSVO).sum())
        clsnp=int((x["Connection Type"]==CONN_CLSNP).sum())
        today=int(((x["Connection Type"].isin([CONN_NPC,CONN_RECON])) & (x["DATE"]==report_date)).sum())
        cum=npc+recon; disc=clsvo+clsnp; net=cum-disc
        olt_count=sum(1 for v in olt_map.values() if v==bbc)
        rows.append({
            "S.No":len(rows)+1, "DE / AGM ":mgr, "BBM NAME":display_name,
            "AREA":area, "Exclusive/Non Exclusive":"EXCLUSIVE" if len(rows)==0 else "",
            "No. Of OLTEs Mapped":olt_count, "Monthly Target":bbc_target,
            f"Daily Provision{report_date:%d-%m-%Y}":today,
            "Cumulative Achievement":cum, "% of Achievement":(cum/bbc_target if bbc_target else 0),
            "CLSVO":clsvo, "CLSNP":clsnp, "Disconnections":disc, "NET":net,
            "NPC":npc, "RECONNECTIONS":recon
        })
    return pd.DataFrame(rows)

def _style_table(ws, df, workbook, startrow=0, header_format=None, widths=None):
    """Apply the same visual language used by the HTML dashboard table."""
    if df.empty:
        return
    cols = list(df.columns)
    for j, col in enumerate(cols):
        if widths and col in widths:
            width = widths[col]
        else:
            sample = df[col].head(300).fillna("").astype(str)
            width = min(max(10, max([len(str(col))] + [len(str(x)) for x in sample]) + 2), 34)
        ws.set_column(j, j, width)
    if header_format:
        for j, col in enumerate(cols):
            ws.write(startrow, j, col, header_format)


def _write_xlsx(df, bbc_report, output_xlsx, report_date, stats,
                franchise_report, manager_report, olt_master, bbc_master):
    """Create Excel output with HTML-matching styling and all columns retained."""
    out = Path(output_xlsx)
    out.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(
        out, engine="xlsxwriter",
        datetime_format="dd-mmm-yyyy", date_format="dd-mmm-yyyy"
    ) as writer:
        wb = writer.book

        # HTML palette from the supplied dashboard:
        # body #F2F4F8, header #373C6F, blue #1F4EBA,
        # negative #FDE8E8/#C80000, positive #E8F8E8/#0A8A0A.
        body_bg, table_header, blue = "#F2F4F8", "#373C6F", "#1F4EBA"
        white, text, muted, border = "#FFFFFF", "#222222", "#666666", "#DDDDDD"
        stripe, neg_bg, neg_fg = "#F7F9FC", "#FDE8E8", "#C80000"
        pos_bg, pos_fg = "#E8F8E8", "#0A8A0A"

        title = wb.add_format({
            "bold": True, "font_size": 22, "font_color": white,
            "bg_color": table_header, "align": "center", "valign": "vcenter"
        })
        sub = wb.add_format({
            "italic": True, "font_color": muted, "bg_color": body_bg,
            "align": "center", "valign": "vcenter"
        })
        header = wb.add_format({
            "bold": True, "font_color": white, "bg_color": table_header,
            "border": 1, "border_color": border,
            "align": "center", "valign": "vcenter", "text_wrap": True
        })
        total_fmt = wb.add_format({
            "bold": True, "font_color": white, "bg_color": table_header,
            "border": 1, "border_color": border,
            "align": "center", "valign": "vcenter"
        })
        cell = wb.add_format({
            "font_color": text, "bg_color": white, "border": 1,
            "border_color": border, "align": "center", "valign": "vcenter"
        })
        cell_alt = wb.add_format({
            "font_color": text, "bg_color": stripe, "border": 1,
            "border_color": border, "align": "center", "valign": "vcenter"
        })
        pct = wb.add_format({
            "num_format": "0.0%", "border": 1, "border_color": border,
            "align": "center", "valign": "vcenter"
        })
        pct_alt = wb.add_format({
            "num_format": "0.0%", "bg_color": stripe, "border": 1,
            "border_color": border, "align": "center", "valign": "vcenter"
        })
        neg = wb.add_format({
            "font_color": neg_fg, "bg_color": neg_bg, "bold": True,
            "border": 1, "border_color": border, "align": "center", "valign": "vcenter"
        })
        pos = wb.add_format({
            "font_color": pos_fg, "bg_color": pos_bg, "bold": True,
            "border": 1, "border_color": border, "align": "center", "valign": "vcenter"
        })
        wrap = wb.add_format({"text_wrap": True, "valign": "top",
                              "font_color": text, "bg_color": white})
        card_value = wb.add_format({
            "bold": True, "font_size": 20, "font_color": blue,
            "bg_color": white, "border": 1, "border_color": "#E0E0E0",
            "align": "center", "valign": "vcenter"
        })
        card_label = wb.add_format({
            "font_size": 9, "font_color": muted, "bg_color": white,
            "border": 1, "border_color": "#E0E0E0",
            "align": "center", "valign": "vcenter"
        })
        card_neg = wb.add_format({
            "bold": True, "font_size": 20, "font_color": neg_fg,
            "bg_color": neg_bg, "border": 1, "border_color": "#E0E0E0",
            "align": "center", "valign": "vcenter"
        })

        # -------------------- DATA --------------------
        data_out = df.copy()
        if "DATE" in data_out.columns:
            data_out["DATE"] = pd.to_datetime(data_out["DATE"], errors="coerce")
        data_out.to_excel(writer, sheet_name="Data", index=False)
        ws = writer.sheets["Data"]
        ws.hide_gridlines(2)
        ws.freeze_panes(1, 0)
        ws.set_tab_color(blue)
        _style_table(ws, data_out, wb, 0, header,
                     {"BBC Name": 30, "Maintenance Franchisee": 28, "Connection Type": 22})
        if len(data_out):
            ws.add_table(0, 0, len(data_out), len(data_out.columns)-1,
                         {"name": "tbl_FTTH_Data", "style": "Table Style Medium 2",
                          "columns": [{"header": c} for c in data_out.columns]})

        # -------------------- MAIN HTML-STYLE TABLE --------------------
        bbc_report = bbc_report.copy()
        cols = list(bbc_report.columns)
        bbc_report.to_excel(writer, sheet_name="FTTHDashboard", index=False, startrow=6)

        ws = writer.sheets["FTTHDashboard"]
        ws.hide_gridlines(2)
        ws.freeze_panes(7, 0)
        ws.set_tab_color(blue)
        last_letter = xlsxwriter.utility.xl_col_to_name(max(0, len(cols)-1))

        ws.merge_range(f"A1:{last_letter}2", "FTTH WARANGAL OA DASHBOARD", title)
        ws.merge_range(f"A3:{last_letter}3",
                       f"Daily provisions dashboard | As on {report_date:%d-%m-%Y}", sub)
        ws.set_row(0, 28)
        ws.set_row(1, 28)
        ws.set_row(6, 34)

        widths = {}
        for c in cols:
            cu = str(c).upper()
            if "MANAGER" in cu:
                widths[c] = 22
            elif c in ("BBM NAME", "BBM Name", "BBC Name", "Display Name"):
                widths[c] = 34
            elif c in ("AREA", "Area"):
                widths[c] = 16
            elif "%" in str(c):
                widths[c] = 15
            elif "DAILY" in cu:
                widths[c] = 18
            else:
                widths[c] = min(max(12, len(str(c))+3), 20)
        _style_table(ws, bbc_report, wb, 6, header, widths)

        # Total OA row. No columns are removed, including NPC and RECONNECTIONS.
        total = {}
        for c in cols:
            cu = str(c).upper()
            if c == "S.No":
                total[c] = "Total OA"
            elif "MANAGER" in cu:
                total[c] = "WGL"
            elif any(k in cu for k in ("BBM NAME", "BBC NAME", "DISPLAY NAME", "AREA",
                                       "EXCLUSIVE/NON EXCLUSIVE")):
                total[c] = ""
            elif c == "% of Achievement":
                target = pd.to_numeric(bbc_report.get("Monthly Target", 0), errors="coerce").fillna(0).sum()
                ach = pd.to_numeric(bbc_report.get("Cumulative Achievement", 0), errors="coerce").fillna(0).sum()
                total[c] = ach / target if target else 0
            else:
                total[c] = pd.to_numeric(bbc_report[c], errors="coerce").fillna(0).sum()

        for j, c in enumerate(cols):
            val = total.get(c, "")
            if c == "% of Achievement":
                total_pct = wb.add_format({
                    "bold": True, "font_color": white, "bg_color": table_header,
                    "border": 1, "border_color": border, "num_format": "0.0%",
                    "align": "center", "valign": "vcenter"
                })
                ws.write_number(7, j, float(val), total_pct)
            else:
                ws.write(7, j, "" if pd.isna(val) else val, total_fmt)

        net_col = next((i for i,c in enumerate(cols) if str(c).upper()=="NET"), None)
        pct_col = next((i for i,c in enumerate(cols) if str(c).upper()=="% OF ACHIEVEMENT"), None)

        for i, (_, r) in enumerate(bbc_report.iterrows(), start=8):
            alt = ((i-8) % 2 == 1)
            for j, c in enumerate(cols):
                val = "" if pd.isna(r[c]) else r[c]
                if j == net_col and val != "":
                    try:
                        fval = float(val)
                        ws.write(i, j, fval, neg if fval < 0 else pos)
                    except (TypeError, ValueError):
                        ws.write(i, j, val, cell_alt if alt else cell)
                elif j == pct_col and val != "":
                    ws.write(i, j, val, pct_alt if alt else pct)
                else:
                    ws.write(i, j, val, cell_alt if alt else cell)

        ws.set_row(7, 24)
        ws.autofilter(6, 0, 7+len(bbc_report), len(cols)-1)

        if net_col is not None and len(bbc_report):
            ws.conditional_format(8, net_col, 7+len(bbc_report), net_col,
                                  {"type": "cell", "criteria": "<", "value": 0, "format": neg})
            ws.conditional_format(8, net_col, 7+len(bbc_report), net_col,
                                  {"type": "cell", "criteria": ">=", "value": 0, "format": pos})

        # -------------------- OTHER REPORTS --------------------
        for sname, dd in [("Manager_Report", manager_report),
                          ("Franchisee_Report", franchise_report)]:
            dd.to_excel(writer, sheet_name=sname, index=False)
            ws = writer.sheets[sname]
            ws.hide_gridlines(2)
            ws.freeze_panes(1, 0)
            ws.set_tab_color(blue)
            _style_table(ws, dd, wb, 0, header,
                         {"Name": 30, "Manager": 24, "Area": 18})
            if len(dd):
                ws.add_table(0, 0, len(dd), len(dd.columns)-1,
                             {"name": "tbl_"+sname, "style": "Table Style Medium 2",
                              "columns": [{"header": c} for c in dd.columns]})

        # -------------------- MASTER FILES --------------------
        for sname, dd in [("OLT_BBC_Map", olt_master), ("BBC_Master", bbc_master)]:
            dd.to_excel(writer, sheet_name=sname, index=False)
            ws = writer.sheets[sname]
            ws.hide_gridlines(2)
            ws.freeze_panes(1, 0)
            ws.set_tab_color("#4472C4")
            _style_table(ws, dd, wb, 0, header)
            if len(dd):
                ws.add_table(0, 0, len(dd), len(dd.columns)-1,
                             {"name": "tbl_"+sname, "style": "Table Style Medium 2",
                              "columns": [{"header": c} for c in dd.columns]})

        # -------------------- CHART SHEETS --------------------
        cr = wb.add_worksheet("Charts_BBC")
        cr.hide_gridlines(2)
        cr.set_column("A:A", 2)
        cr.set_column("B:Q", 13)
        cr.merge_range("B2:Q3", "BBC / EMPLOYEE PERFORMANCE", title)
        n = len(bbc_report)
        if n:
            chart = wb.add_chart({"type":"bar"})
            chart.add_series({"name":"Cumulative", "categories":["FTTHDashboard",8,2,7+n,2],
                              "values":["FTTHDashboard",8,8,7+n,8], "data_labels":{"value":True}})
            chart.set_title({"name":"Cumulative Achievement by BBC / Employee"})
            chart.set_x_axis({"name":"Connections"})
            chart.set_y_axis({"name":"BBC / Employee"})
            chart.set_legend({"none":True})
            chart.set_size({"width":1050,"height":620})
            cr.insert_chart("B5", chart)

            chart2 = wb.add_chart({"type":"bar"})
            chart2.add_series({"name":"NET", "categories":["FTTHDashboard",8,2,7+n,2],
                               "values":["FTTHDashboard",8,13,7+n,13],
                               "data_labels":{"value":True}})
            chart2.set_title({"name":"NET by BBC / Employee"})
            chart2.set_x_axis({"name":"NET Connections"})
            chart2.set_y_axis({"name":"BBC / Employee"})
            chart2.set_legend({"none":True})
            chart2.set_size({"width":1050,"height":620})
            cr.insert_chart("B38", chart2)

        cr2 = wb.add_worksheet("Charts_Operations")
        cr2.hide_gridlines(2)
        cr2.set_column("A:A", 2)
        cr2.set_column("B:Q", 13)
        cr2.merge_range("B2:Q3", "OPERATIONS & TARGET ANALYSIS", title)
        if n:
            ch = wb.add_chart({"type":"column"})
            ch.add_series({"name":"Target", "categories":["FTTHDashboard",8,2,7+n,2],
                           "values":["FTTHDashboard",8,6,7+n,6]})
            ch.add_series({"name":"Achieved", "categories":["FTTHDashboard",8,2,7+n,2],
                           "values":["FTTHDashboard",8,8,7+n,8]})
            ch.set_title({"name":"Monthly Target vs Cumulative Achievement"})
            ch.set_y_axis({"name":"Connections"})
            ch.set_size({"width":1050,"height":600})
            cr2.insert_chart("B5", ch)

            ch2 = wb.add_chart({"type":"column","subtype":"stacked"})
            ch2.add_series({"name":"CLSVO", "categories":["FTTHDashboard",8,2,7+n,2],
                            "values":["FTTHDashboard",8,10,7+n,10]})
            ch2.add_series({"name":"CLSNP", "categories":["FTTHDashboard",8,2,7+n,2],
                            "values":["FTTHDashboard",8,11,7+n,11]})
            ch2.set_title({"name":"Disconnections by BBC / Employee"})
            ch2.set_y_axis({"name":"Connections"})
            ch2.set_size({"width":1050,"height":600})
            cr2.insert_chart("B36", ch2)

        # -------------------- KPI DASHBOARD --------------------
        dash = wb.add_worksheet("Dashboard")
        dash.hide_gridlines(2)
        dash.set_column("A:A", 3)
        dash.set_column("B:Y", 12)
        dash.merge_range("B2:Y4", "FTTH WARANGAL OA – EXECUTIVE DASHBOARD", title)

        kpis = [
            ("MONTHLY TARGET", stats["target"]),
            ("CUMULATIVE", stats["cum"]),
            ("NPC", stats["npc"]),
            ("RECONNECTIONS", stats["reconnections"]),
            ("DISCONNECTIONS", stats["disc"]),
            ("ACHIEVEMENT", f'{stats["pct"]:.1f}%'),
            ("NET", stats["net"])
        ]
        for i, (lab, val) in enumerate(kpis):
            c = 2 + i*3
            fmt = card_neg if lab == "NET" and float(stats["net"]) < 0 else card_value
            dash.merge_range(5, c, 6, c+2, str(val), fmt)
            dash.merge_range(7, c, 7, c+2, lab, card_label)

        dash.write("B10",
                   "All BBC / employee rows and ALL columns are retained. "
                   "NET uses the same red/green highlighting as the HTML dashboard.",
                   wrap)
        dash.write_url("B12", "internal:'Charts_BBC'!B5",
                       string="Open BBC / Employee Charts")
        dash.write_url("B14", "internal:'Charts_Operations'!B5",
                       string="Open Operations Charts")
        dash.write_url("B16", "internal:'Manager_Report'!A1",
                       string="Open Manager / MT Report")
        dash.write_url("B18", "internal:'Franchisee_Report'!A1",
                       string="Open Maintenance Franchisee Report")

def run_report(input_file, output_xlsx, output_html):
    src=Path(input_file)
    olt_map, bbc_info, bbc_order, olt_master, bbc_master = _load_masters()
    raw=_read_source(src)
    df=_aggregate(_classify(raw, olt_map))
    report_date=pd.Timestamp(dt.date.today()-dt.timedelta(days=1))
    live_date = dt.date.today()
    bbc=_bbc_report(df, report_date, olt_map, bbc_info, bbc_order)

    unmapped_olts=sorted(set(df.loc[~df["OLT IP"].isin(olt_map),"OLT IP"]) - {""})
    unmapped_names=sorted(set(df.loc[~df["BBC Name"].isin(bbc_info),"BBC Name"]) - {""})

    bdf=df.copy()
    bdf["Manager"]=bdf["BBC Name"].map({k:v[0] for k,v in bbc_info.items()}).fillna("UNMAPPED")
    bdf["Area"]=bdf["BBC Name"].map({k:v[2] for k,v in bbc_info.items()}).fillna("UNMAPPED")

    manager_rows=[]
    for mgr,g in bdf.groupby("Manager",sort=False):
        manager_rows.append({
            "Manager":mgr,
            "NPC":int((g["Connection Type"]==CONN_NPC).sum()),
            "Reconnections":int((g["Connection Type"]==CONN_RECON).sum()),
            "CLSVO":int((g["Connection Type"]==CONN_CLSVO).sum()),
            "CLSNP":int((g["Connection Type"]==CONN_CLSNP).sum()),
            "Cumulative":int(g["Connection Type"].isin([CONN_NPC,CONN_RECON]).sum()),
            "NET":int(g["Connection Type"].isin([CONN_NPC,CONN_RECON]).sum()-g["Connection Type"].isin([CONN_CLSVO,CONN_CLSNP]).sum())
        })
    manager_report=pd.DataFrame(manager_rows)

    franchise_rows=[]
    for mf,g in df.groupby("Maintenance Franchisee",dropna=False,sort=False):
        mf=str(mf) if pd.notna(mf) else "(Blank)"
        franchise_rows.append({
            "Maintenance Franchisee":mf,
            "NPC":int((g["Connection Type"]==CONN_NPC).sum()),
            "Reconnections":int((g["Connection Type"]==CONN_RECON).sum()),
            "CLSVO":int((g["Connection Type"]==CONN_CLSVO).sum()),
            "CLSNP":int((g["Connection Type"]==CONN_CLSNP).sum()),
            "NET":int(g["Connection Type"].isin([CONN_NPC,CONN_RECON]).sum()-g["Connection Type"].isin([CONN_CLSVO,CONN_CLSNP]).sum()),
            "Total Orders":len(g)
        })
    franchise_report=pd.DataFrame(franchise_rows).sort_values("NET",ascending=False) if franchise_rows else pd.DataFrame()

    stats={
        "target":int(bbc["Monthly Target"].sum()),
        "cum":int(bbc["Cumulative Achievement"].sum()),
        "today":int(bbc[f"Daily Provision{report_date:%d-%m-%Y}"].sum()),
        "disc":int(bbc["Disconnections"].sum()),
        "net":int(bbc["NET"].sum()),
        "npc":int(bbc["NPC"].sum()),
        "reconnections":int(bbc["RECONNECTIONS"].sum())
    }
    stats["pct"]=(stats["cum"]/stats["target"]*100) if stats["target"] else 0

    _write_xlsx(df, bbc, output_xlsx, report_date, stats, franchise_report, manager_report, olt_master, bbc_master)

    # -----------------------------------------------------------------------
    # HTML DASHBOARD
    # Based on the supplied reference text document.
    # IMPORTANT: NPC and RECONNECTIONS remain in the Python report/table.
    # -----------------------------------------------------------------------
    rows = bbc.copy()

    # Keep the existing Python columns, including NPC and RECONNECTIONS.
    # The HTML table is generated from the dataframe so no report data is lost.
    html_columns = [
        "S.No", "AGM/ Manager(MT)", "BBM NAME", "AREA",
        "OLTEs Mapped", "Monthly Target",
        f"Daily Provision{report_date:%d-%m-%Y}",
        "Cumulative Achievement", "% of Achievement",
        "NPC", "RECONNECTIONS", "CLSVO", "CLSNP",
        "Disconnections", "NET"
    ]
    html_columns = [c for c in html_columns if c in rows.columns]
    table_rows = rows[html_columns].copy()

    # Build the table explicitly so NET can receive the reference red/green
    # formatting while NPC and RECONNECTIONS remain visible columns.
    def _fmt_html_value(value, col):
        if pd.isna(value):
            return ""
        if col == "% of Achievement":
            try:
                return f"{float(value) * 100:.2f}%"
            except (TypeError, ValueError):
                return str(value)
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        return str(value)

    header_html = "".join(f"<th>{html.escape(str(c))}</th>" for c in html_columns)
    body_html_parts = []

    for _, r in table_rows.iterrows():
        cells = []
        for col in html_columns:
            value = _fmt_html_value(r[col], col)
            if col == "NET":
                try:
                    net_value = float(r[col])
                    css = "neg" if net_value < 0 else "pos"
                    cells.append(f"<td class='{css}'>{html.escape(value)}</td>")
                except (TypeError, ValueError):
                    cells.append(f"<td>{html.escape(value)}</td>")
            else:
                cells.append(f"<td>{html.escape(value)}</td>")
        body_html_parts.append("<tr>" + "".join(cells) + "</tr>")

    table_body = "\n".join(body_html_parts)

    # Chart data: reference dashboard uses BBC-wise Target vs NET and
    # BBC-wise % Achievement.
    labels_js = ",".join(
        "'" + str(v).replace("\\", "\\\\").replace("'", "\\'") + "'"
        for v in rows["BBM NAME"].fillna("").astype(str)
    )
    target_js = ",".join(
        str(int(v)) if pd.notna(v) else "0"
        for v in pd.to_numeric(rows["Monthly Target"], errors="coerce").fillna(0)
    )
    net_js = ",".join(
        str(int(v)) if pd.notna(v) else "0"
        for v in pd.to_numeric(rows["NET"], errors="coerce").fillna(0)
    )
    pct_js = ",".join(
        f"{float(v) * 100:.2f}" if pd.notna(v) else "0"
        for v in pd.to_numeric(rows["% of Achievement"], errors="coerce").fillna(0)
    )

    kpi_items = [
        ("Monthly Target", f'{stats["target"]:,}'),
        (f'Provisions on {report_date:%d-%b-%Y}', f'{stats["today"]:,}'),
        ("Cumulative Achievement", f'{stats["cum"]:,}'),
        ("% Achievement", f'{stats["pct"]:.2f}%'),
        ("Total Disconnections", f'{stats["disc"]:,}'),
        ("NET", f'{stats["net"]:+,}'),
        ("RECONNECTIONS", f'{stats["reconnections"]:,}')
    ]
    kpi_html = "".join(
        f'<div class="kpi {"neg" if lab == "NET" and stats["net"] < 0 else ""}">'
        f'<div class="v">{value}</div><div class="l">{html.escape(lab)}</div></div>'
        for lab, value in kpi_items
    )


    def _esc_svg(value):
        return (html.escape(str(value))
                .replace("'", "&#39;"))

    def _nice_max(value, minimum=1):
        value = float(value or 0)
        if value <= 0:
            return minimum
        raw = value * 1.15
        step = 1
        if raw >= 10000:
            step = 5000
        elif raw >= 1000:
            step = 500
        elif raw >= 100:
            step = 50
        elif raw >= 10:
            step = 5
        return max(minimum, int((raw + step - 1) // step) * step)

    def _fmt_num(value):
        try:
            return f"{float(value):,.0f}"
        except Exception:
            return "0"

    chart_w, chart_h = 760, 320
    plot_left, plot_right = 58, 730
    plot_top, plot_bottom = 38, 255
    plot_h = plot_bottom - plot_top

    # Chart 1: Monthly Target vs NET (inline SVG; works offline and in Streamlit).
    target_vals = pd.to_numeric(rows["Monthly Target"], errors="coerce").fillna(0).astype(float).tolist()
    net_vals = pd.to_numeric(rows["NET"], errors="coerce").fillna(0).astype(float).tolist()
    chart_labels = rows["BBM NAME"].fillna("").astype(str).tolist()
    max_abs = max([abs(v) for v in target_vals + net_vals] + [1])
    y_max = _nice_max(max_abs)
    y_min = min(0, min(net_vals + [0]))
    y_min = -max(y_max * 0.15, abs(y_min) * 1.15) if y_min < 0 else 0
    y_range = y_max - y_min

    def _y(v):
        return plot_bottom - ((float(v) - y_min) / y_range) * plot_h

    svg1 = [f'<svg viewBox="0 0 {chart_w} {chart_h}" class="chart-svg" role="img" aria-label="BBC Wise Monthly Target versus NET">']
    svg1.append('<rect x="0" y="0" width="760" height="320" fill="#ffffff"/>')
    svg1.append('<text x="380" y="22" text-anchor="middle" font-size="15" font-weight="700" fill="#222">BBC Wise: Monthly Target vs NET</text>')

    for tick in range(5):
        val = y_min + (y_max - y_min) * tick / 4
        yy = _y(val)
        svg1.append(f'<line x1="{plot_left}" y1="{yy:.1f}" x2="{plot_right}" y2="{yy:.1f}" stroke="#e5e7eb"/>')
        svg1.append(f'<text x="50" y="{yy+4:.1f}" text-anchor="end" font-size="10" fill="#666">{_fmt_num(val)}</text>')

    svg1.append(f'<line x1="{plot_left}" y1="{plot_top}" x2="{plot_left}" y2="{plot_bottom}" stroke="#888"/>')
    svg1.append(f'<line x1="{plot_left}" y1="{_y(0):.1f}" x2="{plot_right}" y2="{_y(0):.1f}" stroke="#888"/>')

    n = max(len(chart_labels), 1)
    group_w = (plot_right - plot_left) / n
    bar_w = max(4, min(22, group_w * 0.30))
    for i, label in enumerate(chart_labels):
        cx = plot_left + group_w * (i + 0.5)
        vals = [(target_vals[i], "#1f4eba"), (net_vals[i], "#c80000")]
        for j, (val, fill) in enumerate(vals):
            x = cx + (j - 0.5) * bar_w
            y0, y1 = _y(0), _y(val)
            top = min(y0, y1)
            height = max(1, abs(y1 - y0))
            svg1.append(f'<rect x="{x:.1f}" y="{top:.1f}" width="{bar_w:.1f}" height="{height:.1f}" rx="2" fill="{fill}"/>')
        if n <= 12:
            short = label if len(label) <= 16 else label[:14] + "…"
            svg1.append(f'<text x="{cx:.1f}" y="275" text-anchor="middle" font-size="10" fill="#444" transform="rotate(-25 {cx:.1f} 275)">{_esc_svg(short)}</text>')

    svg1.append('<rect x="520" y="292" width="12" height="12" fill="#1f4eba"/><text x="538" y="302" font-size="11" fill="#444">Monthly Target</text>')
    svg1.append('<rect x="640" y="292" width="12" height="12" fill="#c80000"/><text x="658" y="302" font-size="11" fill="#444">NET</text>')
    svg1.append('</svg>')
    chart1_svg = "".join(svg1)

    # Chart 2: % Achievement as an inline SVG line chart.
    pct_vals = pd.to_numeric(rows["% of Achievement"], errors="coerce").fillna(0).astype(float).mul(100).tolist()
    pct_max = max([100] + pct_vals)
    pct_ymax = max(100, ((pct_max * 1.15 + 9) // 10) * 10)
    def _yp(v):
        return plot_bottom - (float(v) / pct_ymax) * plot_h

    svg2 = [f'<svg viewBox="0 0 {chart_w} {chart_h}" class="chart-svg" role="img" aria-label="BBC Wise percentage of achievement">']
    svg2.append('<rect x="0" y="0" width="760" height="320" fill="#ffffff"/>')
    svg2.append('<text x="380" y="22" text-anchor="middle" font-size="15" font-weight="700" fill="#222">BBC Wise: % of Achievement</text>')
    for tick in range(6):
        val = pct_ymax * tick / 5
        yy = _yp(val)
        svg2.append(f'<line x1="{plot_left}" y1="{yy:.1f}" x2="{plot_right}" y2="{yy:.1f}" stroke="#e5e7eb"/>')
        svg2.append(f'<text x="50" y="{yy+4:.1f}" text-anchor="end" font-size="10" fill="#666">{val:.0f}%</text>')
    svg2.append(f'<line x1="{plot_left}" y1="{plot_top}" x2="{plot_left}" y2="{plot_bottom}" stroke="#888"/>')
    svg2.append(f'<line x1="{plot_left}" y1="{plot_bottom}" x2="{plot_right}" y2="{plot_bottom}" stroke="#888"/>')

    if pct_vals:
        points = []
        for i, val in enumerate(pct_vals):
            x = plot_left + ((plot_right - plot_left) * (i / max(len(pct_vals)-1, 1)))
            y = _yp(val)
            points.append((x, y))
        if len(points) > 1:
            point_str = " ".join(f"{x:.1f},{y:.1f}" for x, y in points)
            svg2.append(f'<polyline points="{point_str}" fill="none" stroke="#c80000" stroke-width="3"/>')
        for i, (x, y) in enumerate(points):
            val = pct_vals[i]
            svg2.append(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="4" fill="#c80000"/>')
            if len(points) <= 12:
                svg2.append(f'<text x="{x:.1f}" y="{max(34, y-8):.1f}" text-anchor="middle" font-size="10" fill="#333">{val:.1f}%</text>')
                short = chart_labels[i] if len(chart_labels[i]) <= 16 else chart_labels[i][:14] + "…"
                svg2.append(f'<text x="{x:.1f}" y="275" text-anchor="middle" font-size="10" fill="#444" transform="rotate(-25 {x:.1f} 275)">{_esc_svg(short)}</text>')
    svg2.append('</svg>')
    chart2_svg = "".join(svg2)

    html_doc = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>WARANGAL OA FTTH Dashboard</title>
<style>
body{{font-family:Segoe UI,Arial,sans-serif;background:#f2f4f8;margin:0;color:#222;}}
.banner{{background:linear-gradient(90deg,#373c6f,#1f4eba);color:#fff;padding:22px 30px;
display:flex;align-items:center;gap:18px;}}
.logo-img{{height:58px;width:auto;object-fit:contain;display:block;}}
.banner h1{{margin:0;font-size:26px;letter-spacing:1px;}}
.banner .sub{{opacity:.85;font-size:13px;margin-top:4px;}}
.kpis{{display:flex;gap:16px;padding:20px 30px 0 30px;flex-wrap:wrap;}}
.kpi{{background:#fff;border-radius:10px;box-shadow:0 2px 6px rgba(0,0,0,.08);
padding:16px 22px;min-width:150px;}}
.kpi .v{{font-size:26px;font-weight:700;color:#1f4eba;}}
.kpi .l{{font-size:12px;color:#666;text-transform:uppercase;letter-spacing:.5px;}}
.kpi.neg .v{{color:#c80000;}}
.wrap{{padding:24px 30px;}}
.table-wrap{{overflow-x:auto;border-radius:10px;}}
table{{border-collapse:collapse;width:100%;min-width:1250px;background:#fff;
box-shadow:0 2px 6px rgba(0,0,0,.06);}}
th,td{{border:1px solid #ddd;padding:8px 10px;text-align:center;font-size:13px;}}
th{{background:#373c6f;color:#fff;white-space:nowrap;}}
tr:nth-child(even){{background:#f7f9fc;}}
td.neg{{color:#c80000;font-weight:700;background:#fde8e8;}}
td.pos{{color:#0a8a0a;font-weight:700;background:#e8f8e8;}}
.note{{font-size:12px;color:#888;margin-top:10px;font-style:italic;}}
button.dl{{margin-top:16px;background:#1f4eba;color:#fff;border:0;padding:10px 18px;
border-radius:6px;cursor:pointer;font-size:14px;}}
button.dl:hover{{background:#173a8f;}}
.charts{{display:flex;gap:20px;flex-wrap:wrap;margin-top:24px;}}
.chartbox{{background:#fff;border-radius:10px;box-shadow:0 2px 6px rgba(0,0,0,.08);
padding:16px;flex:1;min-width:420px;overflow:hidden;}}
.chart-svg{{display:block;width:100%;height:auto;min-height:300px;}}
@media(max-width:900px){{
.kpis{{display:grid;grid-template-columns:repeat(2,1fr);}}
.kpi{{min-width:0;}}
.chartbox{{min-width:0;}}
.banner{{padding:18px;}}
.wrap{{padding:18px;}}
}}
</style>
</head>
<body>
<div class="banner">
  <img class="logo-img" src="data:image/png;base64,{BSNL_LOGO_B64}" alt="BSNL Logo">
  <div>
    <h1>FTTH WARANGAL DASHBOARD</h1>
    <div class="sub">
      BBM Wise Provisioning Report of WGL OA as on {report_date:%d-%b-%Y}
      &middot; Reported by VAMSHI KRISHNA ADEPU on {live_date:%d-%b-%Y}
    </div>
  </div>
</div>

<div class="kpis">{kpi_html}</div>

<div class="wrap">
  <div class="table-wrap">
    <table id="dashboardTable">
      <thead><tr>{header_html}</tr></thead>
      <tbody>{table_body}</tbody>
    </table>
  </div>

  <div class="note">
    NPC and RECONNECTIONS are retained as separate columns. Cumulative Achievement
    equals NPC + RECONNECTIONS, while NET equals Cumulative Achievement minus
    CLSVO and CLSNP.
  </div>

  <button class="dl" onclick="downloadCsv()">Download table as CSV</button>

  <div class="charts">
    <div class="chartbox">{chart1_svg}</div>
    <div class="chartbox">{chart2_svg}</div>
  </div>
</div>

<script>
function downloadCsv(){{
  const rows=[...document.querySelectorAll('#dashboardTable tr')].map(r=>
    [...r.children].map(c=>'"'+c.innerText.replace(/"/g,'""')+'"').join(',')
  );
  const blob=new Blob([rows.join('\\n')],{{type:'text/csv;charset=utf-8;'}});
  const a=document.createElement('a');
  a.href=URL.createObjectURL(blob);
  a.download='FTTH_Warangal_Dashboard.csv';
  a.click();
  URL.revokeObjectURL(a.href);
}}
</script>
</body>
</html>"""

    # Write the HTML file before returning its path.
    # Streamlit reads html_path immediately after run_report() returns.
    output_html_path = Path(output_html)
    output_html_path.parent.mkdir(parents=True, exist_ok=True)
    output_html_path.write_text(html_doc, encoding="utf-8")

    return Path(output_xlsx), output_html_path, {
        "rows_processed":len(df),"unmapped_olt_ips":unmapped_olts,
        "unmapped_bbc_names":unmapped_names,
        "connection_types":df["Connection Type"].value_counts().to_dict()
    }



# =============================================================================
# THREE-INPUT DAILY PROVISIONING REPORT
# =============================================================================
# The original run_report() and its HTML dashboard are retained for backward
# compatibility.  Streamlit can call run_daily_report() when the user uploads:
#   1) OLT Wise Provisioning workbook
#   2) BU Connection Status / UC Approved workbook
#   3) Pending OBS workbook
#
# OLT_BBC_MAP.xlsx and BBC_Master.xlsx remain the authoritative mapping/target
# masters and must be present beside this module, as in the original dashboard.

def _read_uploaded_excel(path, required_columns, preferred_header=2):
    """Read a BSNL export by detecting its header row (usually Excel row 3)."""
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"Uploaded input file not found: {path}")
    last_error = None
    for header in (preferred_header, 0, 1, 3, 4):
        try:
            frame = pd.read_excel(path, sheet_name=0, header=header)
            frame = frame.dropna(how="all").copy()
            frame.columns = [str(c).strip() for c in frame.columns]
            normalized = {str(c).strip().casefold() for c in frame.columns}
            if all(str(c).strip().casefold() in normalized for c in required_columns):
                return frame
        except Exception as exc:
            last_error = exc
    raise ValueError(
        f"Could not identify the header row in '{path.name}'. "
        f"Expected columns: {required_columns}. Last read error: {last_error}"
    )


def _clean_text(series):
    return series.fillna("").astype(str).str.strip()


def _map_franchise_to_bbm(provision_df):
    """Create a reliable franchise-code -> BBM map from the provisioning export."""
    mapping = {}
    if "Maintenance Franchisee" not in provision_df.columns or "BBC Name" not in provision_df.columns:
        return mapping
    temp = provision_df[["Maintenance Franchisee", "BBC Name"]].copy()
    temp["Maintenance Franchisee"] = _clean_text(temp["Maintenance Franchisee"])
    temp["BBC Name"] = temp["BBC Name"].map(normalize_bbc_name)
    temp = temp[(temp["Maintenance Franchisee"] != "") & (temp["BBC Name"] != "")]
    for code, group in temp.groupby("Maintenance Franchisee"):
        names = group["BBC Name"].value_counts()
        if len(names):
            mapping[code] = names.index[0]
    return mapping


def _resolve_bbm_from_olt_or_franchise(frame, olt_map, franchise_map,
                                        olt_col, franchise_cols=()):
    """Resolve BBM first from OLT IP, then from maintenance/franchise code."""
    olt_values = _clean_text(frame[olt_col]) if olt_col and olt_col in frame.columns else pd.Series("", index=frame.index)
    resolved = olt_values.map(olt_map)
    for col in franchise_cols:
        if col in frame.columns:
            codes = _clean_text(frame[col])
            fallback = codes.map(franchise_map)
            resolved = resolved.fillna(fallback)
    return resolved.map(normalize_bbc_name)


def _export_timestamp(path):
    """Read 'Export Time:' or 'Last Update Time:' metadata from the first rows."""
    try:
        raw = pd.read_excel(path, sheet_name=0, header=None, nrows=3)
        for value in raw.astype(object).values.ravel():
            if isinstance(value, str) and ("Export Time:" in value or "Last Update Time:" in value):
                candidate = value.split(":", 1)[1].strip()
                parsed = pd.to_datetime(candidate, errors="coerce", dayfirst=True)
                if pd.notna(parsed):
                    return parsed
    except Exception:
        pass
    return None


def _daily_report_frame(provisioning_file, uc_file, pending_file, report_date=None):
    """Build the 27-column BBM-wise report data using the three uploaded files."""
    olt_map, bbc_info, bbc_order, olt_master, bbc_master = _load_masters()

    provision = _read_uploaded_excel(
        provisioning_file,
        ["OLT IP", "Order Id", "CLSR", "Completion_Date", "Maintenance Franchisee", "BBC Name"],
    )
    uc = _read_uploaded_excel(
        uc_file,
        ["OLT IP", "BBC approved", "Comments", "UC Type"],
    )
    pending = _read_uploaded_excel(
        pending_file,
        ["Maintenance franchise Code", "Status", "Pending Task"],
    )

    # Normalize provisioning rows using the same classification rules as the
    # existing dashboard. This preserves its NPC / RECON / CLSNP / CLSVO logic.
    classified = _aggregate(_classify(provision, olt_map))
    classified["DATE"] = _parse_source_date(classified["Completion_Date"])

    # The provisioning export is the source of the reporting cut-off.  Its
    # export timestamp normally represents the morning after the reporting day.
    if report_date is None:
        export_time = _export_timestamp(provisioning_file)
        if export_time is not None:
            report_date = (export_time.normalize() - pd.Timedelta(days=1))
        else:
            report_date = pd.Timestamp(dt.date.today() - dt.timedelta(days=1))
    report_date = pd.Timestamp(report_date).normalize()
    month_start = report_date.replace(day=1)
    month_rows = classified[
        (classified["DATE"] >= month_start) & (classified["DATE"] <= report_date)
    ].copy()

    franchise_map = _map_franchise_to_bbm(classified)
    # Also allow a franchise-service-code link from UC export to map records
    # where OLT IP is missing or absent from the OLT master.
    uc_bbm = _resolve_bbm_from_olt_or_franchise(
        uc, olt_map, franchise_map, "OLT IP", ("Franchise Service Code",)
    )
    uc = uc.copy()
    uc["BBM_KEY"] = uc_bbm
    uc["UC_APPROVED_FLAG"] = _clean_text(uc["BBC approved"]).str.upper().eq("C")
    uc["UC_REJECTED_FLAG"] = _clean_text(uc["Comments"]).str.upper().str.contains(
        "BBC NOT APPROVED", na=False
    )

    # Pending OBS has no OLT IP in the supplied export. Map through its
    # Maintenance franchise Code, using the provisioning export as crosswalk.
    pending = pending.copy()
    pending["BBM_KEY"] = _resolve_bbm_from_olt_or_franchise(
        pending, olt_map, franchise_map, None,
        ("Maintenance franchise Code", "Sales Franchisee", "Work Franchisee")
    )

    # Counts from the source exports.
    olt_counts = {}
    if "OLT IP" in olt_master.columns and "BBC Name" in olt_master.columns:
        tmp = olt_master.copy()
        tmp["OLT IP"] = _clean_text(tmp["OLT IP"])
        tmp["BBC Name"] = tmp["BBC Name"].map(normalize_bbc_name)
        olt_counts = tmp[tmp["OLT IP"] != ""].groupby("BBC Name")["OLT IP"].nunique().to_dict()

    uc_group = uc.groupby("BBM_KEY", dropna=False).agg(
        UC_REQUIRED=("BBM_KEY", "size"),
        UC_APPROVED=("UC_APPROVED_FLAG", "sum"),
        UC_REJECTED=("UC_REJECTED_FLAG", "sum"),
    ) if len(uc) else pd.DataFrame()

    pending_group = pending.groupby("BBM_KEY", dropna=False).size().to_dict()

    # New-project connections can only be identified where the source explicitly
    # identifies a school / junior-college customer or plan. Other categories
    # are not present as explicit fields in this export and are left blank.
    school_mask = pd.Series(False, index=month_rows.index)
    for col in ("Customer Name", "plan name", "User Type"):
        if col in month_rows.columns:
            school_mask |= _clean_text(month_rows[col]).str.upper().str.contains(
                r"SCHOOL|JR\\.?\\s*COLL|JUNIOR COLLEGE", regex=True, na=False
            )
    new_month = month_rows[month_rows["Connection Type"].isin([CONN_NPC, CONN_RECON])]

    rows = []
    previous_day = report_date - pd.Timedelta(days=1)
    previous_day_2 = report_date - pd.Timedelta(days=2)
    for index, bbc in enumerate(bbc_order, start=1):
        manager, manager_target, area, target, display_name = bbc_info[bbc]
        all_bbm = classified[classified["BBC Name"] == bbc]
        mtd_bbm = month_rows[month_rows["BBC Name"] == bbc]
        provision_bbm = mtd_bbm[mtd_bbm["Connection Type"].isin([CONN_NPC, CONN_RECON])]
        clsvo = int((mtd_bbm["Connection Type"] == CONN_CLSVO).sum())
        clsnp = int((mtd_bbm["Connection Type"] == CONN_CLSNP).sum())
        gross = int(len(provision_bbm))
        disconnections = clsvo + clsnp
        net = gross - disconnections
        day_1 = int((provision_bbm["DATE"] == previous_day).sum())
        day_2 = int((provision_bbm["DATE"] == report_date).sum())

        u = uc_group.loc[bbc] if len(uc_group) and bbc in uc_group.index else None
        uc_required = int(u["UC_REQUIRED"]) if u is not None else 0
        uc_approved = int(u["UC_APPROVED"]) if u is not None else 0
        uc_rejected = int(u["UC_REJECTED"]) if u is not None else 0
        uc_pending = max(0, uc_required - uc_approved - uc_rejected)

        pending_count = int(pending_group.get(bbc, 0))
        school_count = int(
            ((new_month["BBC Name"] == bbc) & school_mask.loc[new_month.index]).sum()
        )

        rows.append({
            "S.No": index,
            "AGM/ Manager(MT)": manager,
            "BBM NAME": display_name,
            "AREA": area,
            "Exclusive/NonExclusive": "EXCLUSIVE" if index == 1 else "",
            f"OLT COUNT AS ON {report_date:%d.%m.%Y}": int(olt_counts.get(bbc, 0)),
            f"TARGET {report_date:%b-%Y}".upper(): int(target),
            f"DAY PROVISION {previous_day:%d.%m.%Y}": day_1,
            f"DAY PROVISION {report_date:%d.%m.%Y}": day_2,
            "GROSS PROVISION": gross,
            "% ACH": gross / target if target else 0,
            "CLSVO": clsvo,
            "CLSNP": clsnp,
            "TOTAL DISCONNECTIONS": disconnections,
            "NET": net,
            f"PENDING OBS AS ON {report_date:%d.%m.%Y}": pending_count,
            "COMPLETED": None,  # Not included in a current-pending-only export.
            "TO BE COMPLETED": pending_count,
            "TO BE CANCELLED": None,  # No explicit cancellation disposition in source.
            "CIRCLE PROJECT CONX (SCHOOLS/JR.COLG)": school_count,
            "NEW CONN BY OA TEAM": (int(gross)-int(school_count)-int(CONN_RECON)),  # Source has no reliable OA-team flag.
            "LL-FTTH CONVERSION": None,  # Source has no explicit LL conversion field.
            "RECONX FROM INACTIVE (LL/FTTH)": None,  # Source does not identify prior LL/FTTH status.
            "TOTAL UC REQUIRED CONX": uc_required,
            "UC APPROVED BY BBM": uc_approved,
            "UC REJECTD BY BBM": uc_rejected,
            "UC PENDING": uc_pending,
        })

    report = pd.DataFrame(rows)
    return report, {
        "report_date": report_date,
        "classified": classified,
        "uc": uc,
        "pending": pending,
        "unmapped_uc": sorted(set(uc.loc[uc["BBM_KEY"].eq(""), "OLT IP"].astype(str)) - {""}),
        "unmapped_pending": sorted(set(pending.loc[pending["BBM_KEY"].eq(""), "Maintenance franchise Code"].astype(str)) - {""}),
        "provision_export_time": _export_timestamp(provisioning_file),
        "uc_export_time": _export_timestamp(uc_file),
        "pending_export_time": _export_timestamp(pending_file),
        "master_bbc": bbc_master,
    }


def _write_daily_template_report(report, output_xlsx, report_date, raw_sheets=None):
    """Write the reference 27-column layout with grouped/merged headings."""
    import xlsxwriter

    output_xlsx = Path(output_xlsx)
    output_xlsx.parent.mkdir(parents=True, exist_ok=True)
    cols = list(report.columns)

    with pd.ExcelWriter(output_xlsx, engine="xlsxwriter") as writer:
        wb = writer.book
        ws = wb.add_worksheet("DAILY REPORT")
        writer.sheets["DAILY REPORT"] = ws
        ws.hide_gridlines(2)
        ws.freeze_panes(3, 5)
        ws.set_landscape()
        ws.fit_to_pages(1, 0)
        ws.repeat_rows(0, 2)

        navy = "#1F4E78"
        blue = "#D9EAF7"
        pale = "#EAF2F8"
        white = "#FFFFFF"
        border = "#7F8C8D"
        green = "#E2F0D9"
        red = "#FCE4D6"

        title_fmt = wb.add_format({
            "bold": True, "font_size": 16, "font_color": white,
            "bg_color": navy, "align": "center", "valign": "vcenter"
        })
        group_fmt = wb.add_format({
            "bold": True, "font_color": white, "bg_color": navy,
            "border": 1, "border_color": border, "align": "center",
            "valign": "vcenter", "text_wrap": True
        })
        sub_fmt = wb.add_format({
            "bold": True, "font_color": "#000000", "bg_color": blue,
            "border": 1, "border_color": border, "align": "center",
            "valign": "vcenter", "text_wrap": True
        })
        cell_fmt = wb.add_format({
            "border": 1, "border_color": border, "align": "center",
            "valign": "vcenter"
        })
        alt_fmt = wb.add_format({
            "border": 1, "border_color": border, "align": "center",
            "valign": "vcenter", "bg_color": "#F7FBFD"
        })
        total_fmt = wb.add_format({
            "bold": True, "font_color": white, "bg_color": navy,
            "border": 1, "border_color": border, "align": "center",
            "valign": "vcenter"
        })
        pct_fmt = wb.add_format({
            "num_format": "0.0%", "border": 1, "border_color": border,
            "align": "center"
        })
        total_pct_fmt = wb.add_format({
            "bold": True, "font_color": white, "bg_color": navy,
            "num_format": "0.0%", "border": 1, "border_color": border,
            "align": "center"
        })
        neg_fmt = wb.add_format({
            "bold": True, "font_color": "#C00000", "bg_color": red,
            "border": 1, "border_color": border, "align": "center"
        })
        pos_fmt = wb.add_format({
            "bold": True, "font_color": "#006100", "bg_color": green,
            "border": 1, "border_color": border, "align": "center"
        })

        last_col = len(cols) - 1
        ws.merge_range(0, 0, 0, last_col,
                       f"BBM WISE DAILY PROVISIONING REPORT OF WARANGAL OA AS ON {report_date:%d.%m.%Y}",
                       title_fmt)
        ws.set_row(0, 30)

        # Row 2: grouped headings, matching the supplied reference workbook.
        fixed_merges = [(0, 0), (1, 1), (2, 2), (3, 3), (4, 4),
                        (5, 5), (6, 6), (9, 9), (10, 10), (11, 11),
                        (12, 12), (13, 13), (14, 14)]
        for c in fixed_merges:
            ws.merge_range(1, c[0], 2, c[1], cols[c[0]], group_fmt)
        ws.merge_range(1, 7, 1, 8, "DAY PROVISION", group_fmt)
        ws.merge_range(1, 15, 1, 18, "PENDING OBS", group_fmt)
        ws.merge_range(1, 19, 1, 22, "NEW PROVISIONS", group_fmt)
        ws.merge_range(1, 23, 1, 26, "UC STATUS (ALL TYPES)", group_fmt)

        # Row 3: subheaders for grouped columns.
        for j in range(7, 9):
            ws.write(2, j, cols[j].replace("DAY PROVISION ", ""), sub_fmt)
        for j in range(15, 27):
            sub_label = cols[j]
            if j == 15:
                sub_label = sub_label.replace("PENDING OBS ", "")
            ws.write(2, j, sub_label, sub_fmt)

        ws.set_row(1, 28)
        ws.set_row(2, 42)

        # Column widths adapted to the supplied reference layout.
        widths = [7, 20, 31, 16, 17, 15, 14, 13, 13, 14, 10, 10, 10,
                  14, 10, 14, 12, 15, 15, 18, 16, 16, 18, 15, 15, 15, 12]
        for j, width in enumerate(widths):
            ws.set_column(j, j, width)

        # Main rows.
        for i, (_, row) in enumerate(report.iterrows(), start=3):
            fmt = alt_fmt if (i - 3) % 2 else cell_fmt
            for j, col in enumerate(cols):
                value = row[col]
                if pd.isna(value):
                    ws.write_blank(i, j, None, fmt)
                elif col == "% ACH":
                    ws.write_number(i, j, float(value), pct_fmt)
                elif col == "NET":
                    try:
                        ws.write_number(i, j, float(value), neg_fmt if float(value) < 0 else pos_fmt)
                    except Exception:
                        ws.write(i, j, value, fmt)
                else:
                    ws.write(i, j, value, fmt)

        # Merge manager cells vertically for contiguous BBM rows.
        start = 3
        while start < 3 + len(report):
            manager = report.iloc[start - 3]["AGM/ Manager(MT)"]
            end = start
            while end + 1 < 3 + len(report) and report.iloc[end + 1 - 3]["AGM/ Manager(MT)"] == manager:
                end += 1
            if end > start:
                ws.merge_range(start, 1, end, 1, manager, cell_fmt)
            start = end + 1

        # Total row, including weighted achievement.
        total_row = 3 + len(report)
        ws.merge_range(total_row, 0, total_row, 4, "TOTAL", total_fmt)
        for j, col in enumerate(cols):
            if j <= 4:
                continue
            if col == "% ACH":
                target_col = next(c for c in cols if c.startswith("TARGET "))
                total_target = pd.to_numeric(report[target_col], errors="coerce").fillna(0).sum()
                total_gross = pd.to_numeric(report["GROSS PROVISION"], errors="coerce").fillna(0).sum()
                ws.write_number(total_row, j, total_gross / total_target if total_target else 0, total_pct_fmt)
            elif pd.api.types.is_numeric_dtype(report[col]):
                val = pd.to_numeric(report[col], errors="coerce").sum(min_count=1)
                if pd.isna(val):
                    ws.write_blank(total_row, j, None, total_fmt)
                else:
                    ws.write_number(total_row, j, float(val), total_fmt)
            else:
                ws.write(total_row, j, "", total_fmt)
        ws.set_row(total_row, 25)
        ws.autofilter(2, 0, total_row - 1, last_col)

        # Additional raw-data and audit sheets make reconciliation possible.
        if raw_sheets:
            for sheet_name, frame in raw_sheets.items():
                safe_name = sheet_name[:31]
                frame.to_excel(writer, sheet_name=safe_name, index=False)
                raw_ws = writer.sheets[safe_name]
                raw_ws.freeze_panes(1, 0)
                raw_ws.hide_gridlines(2)
                raw_ws.autofilter(0, 0, max(0, len(frame)), max(0, len(frame.columns)-1))
                for j, col in enumerate(frame.columns):
                    raw_ws.set_column(j, j, min(max(12, len(str(col)) + 2), 30))


def run_daily_report(provisioning_file, uc_file, pending_file,
                     output_xlsx, output_html=None, report_date=None):
    """
    Generate the 27-column reference-format report from three Excel uploads.

    Args:
        provisioning_file: OLT Wise Provisioning workbook.
        uc_file: BU Connection Status / UC status workbook.
        pending_file: Pending OBS workbook.
        output_xlsx: destination XLSX path.
        output_html: optional HTML summary path.
        report_date: optional reporting cut-off; defaults to one day before
                     the provisioning export timestamp.
    Returns:
        (xlsx_path, html_path_or_none, metadata)
    """
    report, context = _daily_report_frame(
        provisioning_file, uc_file, pending_file, report_date=report_date
    )
    as_on = context["report_date"]
    raw_sheets = {
        "Provisioning_Input": context["classified"],
        "UC_Input": context["uc"],
        "Pending_OBS_Input": context["pending"],
    }
    audit_rows = [
        ["Reporting date", as_on.strftime("%d-%m-%Y")],
        ["Provisioning export time", str(context["provision_export_time"] or "Not found")],
        ["UC export time", str(context["uc_export_time"] or "Not found")],
        ["Pending OBS export time", str(context["pending_export_time"] or "Not found")],
        ["Unmapped UC OLT IPs", ", ".join(context["unmapped_uc"]) or "None"],
        ["Unmapped Pending franchise codes", ", ".join(context["unmapped_pending"]) or "None"],
        ["Important note", "COMPLETED and TO BE CANCELLED are blank because the Pending OBS export contains only current pending records and no completion/cancellation history."],
        ["Important note", "NEW CONN BY OA TEAM, LL-FTTH CONVERSION and RECONX FROM INACTIVE are blank because the uploaded provisioning export does not contain explicit reliable category fields."],
        ["UC interpretation", "BBC approved = C counted as approved; Comments containing BBC NOT APPROVED counted as rejected; balance counted as pending."],
    ]
    export_dates = [context["provision_export_time"], context["uc_export_time"], context["pending_export_time"]]
    valid_export_dates = [pd.Timestamp(x).normalize() for x in export_dates if x is not None]
    if valid_export_dates and (max(valid_export_dates) - min(valid_export_dates)).days > 1:
        audit_rows.append([
            "DATA DATE WARNING",
            "The three uploaded exports have different export dates. Verify that the inputs belong to the same reporting cycle before using the consolidated totals."
        ])
    raw_sheets["Audit_Notes"] = pd.DataFrame(audit_rows, columns=["Item", "Value"])
    _write_daily_template_report(report, output_xlsx, as_on, raw_sheets=raw_sheets)

    html_path = None
    if output_html:
        html_path = Path(output_html)
        html_path.parent.mkdir(parents=True, exist_ok=True)
        display = report.copy()
        display = display.to_html(index=False, border=0, classes="report", na_rep="")
        html_path.write_text(
            "<!doctype html><html><head><meta charset='utf-8'>"
            "<title>Warangal OA Daily Provisioning Report</title>"
            "<style>body{font-family:Arial;margin:24px;color:#222}"
            "h2{background:#1F4E78;color:white;padding:14px;text-align:center}"
            "table{border-collapse:collapse;width:100%;font-size:12px}"
            "th{background:#1F4E78;color:white}td,th{border:1px solid #bbb;padding:6px;text-align:center}"
            "tr:nth-child(even){background:#f4f8fb}</style></head><body>"
            f"<h2>BBM WISE DAILY PROVISIONING REPORT OF WARANGAL OA AS ON {as_on:%d.%m.%Y}</h2>"
            + display + "</body></html>",
            encoding="utf-8"
        )

    metadata = {
        "report_date": as_on.strftime("%Y-%m-%d"),
        "bbm_count": int(len(report)),
        "gross_provision": int(pd.to_numeric(report["GROSS PROVISION"], errors="coerce").sum()),
        "net": int(pd.to_numeric(report["NET"], errors="coerce").sum()),
        "unmapped_uc_count": len(context["unmapped_uc"]),
        "unmapped_pending_count": len(context["unmapped_pending"]),
    }
    return Path(output_xlsx), html_path, metadata
