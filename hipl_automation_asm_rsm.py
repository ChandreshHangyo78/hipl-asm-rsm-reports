# streamlit_asm_rsm_final_v5.py
import streamlit as st
import pandas as pd
import numpy as np
import io, os, base64, datetime, requests
from msal import ConfidentialClientApplication

# -----------------------------
# CONFIG
# -----------------------------
CLIENT_ID = st.secrets["CLIENT_ID"]
CLIENT_SECRET = st.secrets["CLIENT_SECRET"]
TENANT_ID = st.secrets["TENANT_ID"]
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://graph.microsoft.com/.default"]
SENDER_EMAIL = st.secrets["SENDER_EMAIL"]
GRAPH_API_ENDPOINT = f"https://graph.microsoft.com/v1.0/users/{SENDER_EMAIL}/sendMail"

LOGO_PATH = "hangyo_logo.png"

EMAIL_MAP = {
    "VIJAYA SEKARAN D": {"to": "vijay.sekaran@hangyo.in", "cc": []},
    "PRABHAKARAN D": {"to": "prabhakaran.d@hangyo.in", "cc": []},
    "SASIKUMAR M": {"to": "sasikumar.m@hangyo.in", "cc": []},
    "SRINIVASAN RAVIKUMAR": {"to": "s.ravikumar@hangyo.in", "cc": []},
    "SAKTHIVEL R": {"to": "sakthivel.r@hangyo.in", "cc": []},
    "RAJENDRA H G": {"to": "rajendrahg@hangyo.in", "cc": []},
    "MALLIKARJUNA N": {"to": "mallikarjun.n@hangyo.in", "cc": []},
    "VINAY KUMAR T S": {"to": "vinaykumar@hangyo.in", "cc": []},
    "VEERESH M": {"to": "veereshm254@gmail.com", "cc": []},
    "MANSURALI KHAN KATTIMANI": {"to": "nk8709468562@gmail.com", "cc": []},
    "GARE MOHAN": {"to": "mohan.g@hangyo.in", "cc": []},
    "RAAVI VENKATA RAVI": {"to": "venkata.ravi@hangyo.in", "cc": []},
    "ETTA HARSHAVARDHAN REDDY": {"to": "harshavardhan.reddy@hangyo.in", "cc": []},
    "BEVARA VAMSI SAI KRISHNA": {"to": "sai.krishna@hangyo.in", "cc": []},
    "SARDAR GURUPREET SINGH": {"to": "gurupreet.singh@hangyo.in", "cc": []},
    "DOTI RAMESH": {"to": "ramesh.doti@hangyo.in", "cc": []},
    "RAMCHANDRA KULKARNI": {"to": "ramchandra.k@hangyo.in", "cc": []},
    "AMRESH ANANT PADTE": {"to": "amresh.p@hangyo.in", "cc": []},
    "VIVEK VAMAN KINI": {"to": "vivekkini@hangyo.in", "cc": []},
    "DERICK JASWIN BASTEV FERNANDES": {"to": "derickfernandes@hangyo.in", "cc": []},
    "SUNEEL RAMESH NAIK": {"to": "sunirnaik777@gmail.com", "cc": []},
    "PRAVEEN ACHARYA": {"to": "praveenacharya@hangyo.in", "cc": []},
    "V SHRIDHARA BHAT": {"to": "sridharbhat@hangyo.in", "cc": []},
    "ARJUN RAO": {"to": "arjun.rao@hangyo.in", "cc": []},
    "YASH MONAPPA KULAL": {"to": "yash.kulal@hangyo.in", "cc": []},
    "SHARAD EKNATH SARDAR": {"to": "sharad.sardar@hangyo.in", "cc": []},
    "ANIL BABURAO JADHAV": {"to": "anil.jadhav@hangyo.in", "cc": []},
    "SARFARAZ HIRAPURE": {"to": "sarfaraz.hirapure@hangyo.in", "cc": []},
    "K RAKESH KAMATH": {"to": "rakeshkamath@hangyo.in", "cc": []},
    "VENKATESH SHIVABODHA PATTAR": {"to": "venkatesh.pattar@hangyo.in", "cc": []},
    "MOHHAMMED IBRAHIM": {"to": "m.ibrahim@hangyo.in", "cc": []},
    "RIZWAN M D": {"to": "rizwan.md@hangyo.in", "cc": []},
    "KRISHNA M A": {"to": "krishna@hangyo.in", "cc": []},
    "MOHAMMED AHMED MOHAMMED ABDUL SHUKOOR": {"to": "m.ahmed@hangyo.in", "cc": []},
    "RAVI KUMAR T": {"to": "ravikumar.t@hangyo.in", "cc": []},
    "TANAY SANJAY KULKARNI": {"to": "tanay.kulkarni@hangyo.in", "cc": []},
    "KADAM DNYANESHWAR TANAJI": {"to": "dnyaneshwar.kadam@hangyo.in", "cc": []},
    "PRAVEEN ANTHONY": {"to": "praveenanthony@hangyo.in", "cc": []},
    "PRADEESH C": {"to": "pradeesh.c@hangyo.in", "cc": []},
    "CHANDA KARUNAKAR": {"to": "c.karunakar@hangyo.in", "cc": []},
    "V MADHUSUDHANA RAO": {"to": "madhusudhana.rao@hangyo.in", "cc": []},
    "MUNAGALA VENKATAIAH": {"to": "mungalavenkataiah@hangyo.in", "cc": []},
    "VAKA RAJANIKANTH": {"to": "rajanikanth.v@hangyo.in", "cc": []},
    "BANDLA RUPESH": {"to": "b.rupesh@hangyo.in", "cc": []},
    "TUMMALA VEMAIAH": {"to": "tummala.vemaiah@hangyo.in", "cc": []},
    "GIRISH ANAND KAMAT": {"to": "girishkamathangy@gmail.com", "cc": []},
    "TANGATURI VENKATA MASTHAN": {"to": "tv.masthan@hangyo.in", "cc": []},
    "SARAVANAN A": {"to": "saravananhrks@gmail.com", "cc": []},
    "SUNIL BABU N": {"to": "sunilbabu038@gmail.com", "cc": []},
    "PANTARANGAM RAJASEKHAR REDDY": {"to": "rajasekharrreddy@gmail.com", "cc": []},
    "BIRARI KISHOR BHALCHANDRA": {"to": "kishorbirari@hangyo.in", "cc": []},
    "KISHOR NATHU THAKARE": {"to": "kt.kishor22@gmail.com", "cc": []},
    "YOGESHKUMAR KALAIMANI": {"to": "yogeshkumar.k@hangyo.in", "cc": []},
    "PUSARLA VENKATA RAJU": {"to": "venkata.raju@hangyo.in", "cc": []},
    "JAI PRAKASH HARIWAN": {"to": "jaiprakashhariwan@hangyo.in", "cc": []},
    "JISHNU PRASAD": {"to": "jishnu.prasad@hangyo.in", "cc": []},
    "NOYAL SEBASTIAN": {"to": "noyal.sebastian@hangyo.in", "cc": []},
    "RAJESH KUMAR M": {"to": "rajeshkumar.m@hangyo.in", "cc": []},
    "RAVIKIRAN": {"to": "ravikiran@hangyo.in", "cc": []},
    "MUJEER AHMED KHANBU": {"to": "mujeer.khanbu@hangyo.in", "cc": []},
    "SHIJI KHAN B": {"to": "shijikhan487@gmail.com", "cc": []},
    "THANGAVEL M": {"to": "gemgoldrm@yahoo.co.in", "cc": []},
    "SARAVANAKUMAR M": {"to": "saravanakumar.m@hangyo.in", "cc": []},
    "KASTURI BASKAR": {"to": "kasturi.baskar@gmail.com", "cc": []},
    "AKHILRAJ K": {"to": "akhilrajk2507@gmail.com", "cc": []},
    "KISHOR KASHIRAM JADHAV": {"to": "jadhavkishork777@gmail.com", "cc": []},
    "VUGGINA RAJESH": {"to": "rajesh.vuggina@gmail.com", "cc": []},
    "Vacant Kolhapur": {"to": "venkatesh.pattar@hangyo.in", "cc": []},
    "VACANT Konkan Maharashtra": {"to": "venkatesh.pattar@hangyo.in", "cc": []},
    "Prashanth Jois": {"to": "prashanth.jois@hangyo.in", "cc": []},
    "SANTHOSH KUMAR K G": {"to": "santhosh.kumar@hangyo.in", "cc": []},
    "SHRIDHAR SHETTI": {"to": "shridharshetti@hangyo.in", "cc": []},
    "SHOPAN BABU T": {"to": "shopan.babu@hangyo.in", "cc": []},
    "SANTOSH SHASHIKANT KALE": {"to": "sudhirkapote@hangyo.in", "cc": []},
}


# -----------------------------
# HELPERS
# -----------------------------
def detect_asm_col(df: pd.DataFrame):
    for c in df.columns:
        lc = str(c).lower()
        if "asm" in lc or "rsm" in lc:
            return c
    return None

def get_access_token():
    app = ConfidentialClientApplication(CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET)
    result = app.acquire_token_for_client(scopes=SCOPE)
    return result.get("access_token") if result else None

def send_mail_via_graph(access_token, to_email, cc_list, subject, html_body, attachment_bytes, filename):
    content_b64 = base64.b64encode(attachment_bytes).decode("utf-8")
    message = {
        "message": {
            "subject": subject,
            "body": {"contentType": "HTML", "content": html_body},
            "toRecipients": [{"emailAddress": {"address": to_email}}],
            "ccRecipients": [{"emailAddress": {"address": cc}} for cc in cc_list],
            "attachments": [{
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": filename,
                "contentBytes": content_b64,
                "contentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            }]
        }
    }
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
    resp = requests.post(GRAPH_API_ENDPOINT, headers=headers, json=message, timeout=120)
    return resp.status_code, resp.text

def safe_float(x):
    try:
        if pd.isna(x): return None
        f = float(x)
        if np.isinf(f) or np.isnan(f): return None
        return f
    except:
        return None

def to_base64_img(path):
    if not os.path.exists(path): return ""
    with open(path, "rb") as f: return base64.b64encode(f.read()).decode("utf-8")

def find_col(df_or_index, candidates):
    cols = list(df_or_index)
    # exact match first
    for cand in candidates:
        for c in cols:
            if str(c).strip().lower() == cand.strip().lower():
                return c
    # contains fallback
    for cand in candidates:
        low = cand.strip().lower()
        for c in cols:
            if low in str(c).strip().lower():
                return c
    return None

def looks_like_date_series(s: pd.Series):
    name = str(s.name).lower()
    if "date" in name: 
        return True
    if pd.api.types.is_datetime64_any_dtype(s):
        return True
    try:
        pd.to_datetime(s.dropna().astype(str).head(3), errors="raise")
        return True
    except:
        return False

def is_percent_col(colname: str):
    n = str(colname).lower()
    if "crr" in n or "rrr" in n:
        return False
    tokens = ["%", "ach", "fill rate", "fill_rate", "fillrate", "fill-rate", "trgt vs ach", "achievement"]
    return any(t in n for t in tokens)

def to_fraction_for_excel(raw):
    """
    Convert any raw percent-like number to a fraction for Excel's 0.00% format:
    - 0.72 -> 0.72 (means 72%)  [already fraction]
    - 95.53 -> 0.9553
    - 126   -> 1.26
    """
    v = safe_float(raw)
    if v is None: 
        return None
    if abs(v) <= 1.5:   # already fraction
        return v
    return v / 100.0

def percent_str_for_email(raw, decimals=2):
    """
    Render a correct percent string for the email:
    - 0.72 -> "72.00%"
    - 95.53 -> "95.53%"
    - 126   -> "126.00%"
    """
    v = safe_float(raw)
    if v is None: 
        return f"{0:.{decimals}f}%"
        # return "0.00%"
    if abs(v) <= 1.5:
        v = v * 100
    return f"{v:.{decimals}f}%"

# -----------------------------
# BUILD ATTACHMENT
# -----------------------------
def build_attachment_bytes_for_asm(asm_name, sheet_dfs, asm_col_map):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        wb = writer.book

        # ASM sheet formats (dark borders)
        border_dark = 2
        title_fmt = wb.add_format({"bold": True,"font_color": "white","bg_color": "#d63384",
                                   "align": "center","valign": "vcenter","font_name": "Cambria","font_size": 14,"border": border_dark})
        header_fmt = wb.add_format({"bold": True,"bg_color": "#ffd166","font_color": "#d63384",
                                    "align": "center","valign": "vcenter","border": border_dark,"font_name": "Cambria"})
        text_fmt = wb.add_format({"align": "center","valign": "vcenter","border": border_dark,"font_name": "Cambria"})
        num_fmt = wb.add_format({"num_format": "#,##0.00","align": "center","valign": "vcenter","border": border_dark,"font_name": "Cambria"})
        int_fmt = wb.add_format({"num_format": "#,##0","align": "center","valign": "vcenter","border": border_dark,"font_name": "Cambria"})
        pct_fmt = wb.add_format({"num_format": "0.00%","align": "center","valign": "vcenter","border": border_dark,"font_name": "Cambria"})
        date_fmt = wb.add_format({"num_format": "yyyy-mm-dd","align": "center","valign": "vcenter","border": border_dark,"font_name": "Cambria"})
        hl_text_fmt = wb.add_format({"align":"center","valign":"vcenter","border":border_dark,"font_name":"Cambria",
                                     "bg_color":"#fff19c","bold":True})
        hl_num_fmt = wb.add_format({"num_format":"#,##0.00","align":"center","valign":"vcenter","border":border_dark,
                                    "font_name":"Cambria","bg_color":"#fff19c","bold":True})
        hl_pct_fmt = wb.add_format({"num_format":"0.00%","align":"center","valign":"vcenter","border":border_dark,
                                    "font_name":"Cambria","bg_color":"#fff19c","bold":True})

        # Data-sheet formats (light borders + colored headers)
        border = 1
        ds_title_fmt = wb.add_format({"bold": True,"font_color": "white","bg_color": "#d63384",
                                      "align": "center","valign": "vcenter","font_name": "Cambria","font_size": 12})
        ds_header_fmt = wb.add_format({"bold": True,"bg_color": "#ffd166","font_color": "#d63384",
                                       "align": "center","valign": "vcenter","border": border,"font_name": "Cambria"})
        ds_text_fmt = wb.add_format({"align": "center","valign": "vcenter","border": border,"font_name": "Cambria"})
        ds_num_fmt = wb.add_format({"num_format": "#,##0.00","align": "center","valign": "vcenter","border": border,"font_name": "Cambria"})
        ds_int_fmt = wb.add_format({"num_format": "#,##0","align": "center","valign": "vcenter","border": border,"font_name": "Cambria"})
        ds_pct_fmt = wb.add_format({"num_format": "0.00%","align": "center","valign": "vcenter","border": border,"font_name": "Cambria"})
        ds_date_fmt = wb.add_format({"num_format": "yyyy-mm-dd","align": "center","valign": "vcenter","border": border,"font_name": "Cambria"})

        # ----------------- ASM Summary (perfectly formatted) -----------------
        if "ASM" in sheet_dfs and not sheet_dfs["ASM"].empty:
            df_asm = sheet_dfs["ASM"].copy()
            asm_col = asm_col_map.get("ASM")
            if asm_col and asm_col in df_asm.columns:
                df_filtered = df_asm[df_asm[asm_col].astype(str).str.upper() == asm_name.upper()]
            else:
                df_filtered = df_asm
            if not df_filtered.empty:
                row = df_filtered.iloc[0]
                ws = wb.add_worksheet("ASM Summary")
                ws.merge_range(0,0,0,1,f"ASM Summary - {asm_name}", title_fmt)
                ws.write(1,0,"Metric", header_fmt); ws.write(1,1,"Value", header_fmt)
                r = 2

                # helpers
                def get_fill_rate_fraction():
                    cb = find_col(row.index, ["Secondary Billing"])
                    ck = find_col(row.index, ["Secondary Booking"])
                    bill = safe_float(row[cb]) if cb else None
                    book = safe_float(row[ck]) if ck else None
                    return (bill/book) if (book and book != 0) else 0

                # Date as metric label & primary day sale as value
                date_col = find_col(row.index, ["Date"])
                date_val = ""
                if date_col:
                    try:
                        date_val = pd.to_datetime(row[date_col]).date()
                    except:
                        try:
                            date_val = pd.to_datetime(str(row[date_col]).split()[0]).date()
                        except:
                            date_val = str(row[date_col])
                ws.write(r,0,str(date_val), text_fmt)

                # Try best-effort to fetch the "Primary sale for the date"
                day_sale_cols = [
                    "Primary Day Sale","Primary Daily Sale","Primary Sale (Date)","Primary for Date",
                    "Primary Date Sale","Primary Today's Sale","Primary Sale"
                ]
                day_col = find_col(row.index, day_sale_cols)
                v_day = safe_float(row[day_col]) if day_col else None
                if v_day is None:
                    ws.write(r,1,"", text_fmt)
                else:
                    ws.write_number(r,1,round(v_day,2), num_fmt)
                r += 1

                # Hub Name (if available)
                hub_col = find_col(row.index, ["Hub Name","Hub","State","Region","Branch","Cluster","Territory"])
                if hub_col:
                    ws.write(r,0,"Hub Name", text_fmt)
                    ws.write(r,1,str(row[hub_col]), text_fmt)
                    r += 1

                items = [
                    ("Primary Target","Primary Target","num"),
                    ("Primary MTD","Primary MTD","num"),
                    ("ASM GOLY","Primary Trgt vs Ach %","pct"),
                    ("LYTD Primary","LYTD Primary","num"),
                    ("Primary CRR","Primary CRR","num"),
                    ("Primary RRR","Primary RRR","num"),
                    ("Monthly End Projection","Primary Projected LE","num","HL"),
                    ("Projected Primary Achievement%","Primary_Ach%","pct","HL"),
                    ("Secondary Target","Secondary Target","num"),
                    ("Secondary Booking","Secondary Booking","num"),
                    ("Secondary Billing","Secondary Billing","num"),
                    ("Secondary Fill Rate","Secondary Fill Rate","pct_recompute"),
                    ("Secondary CRR","Secondary CRR","num"),
                    ("Secondary RRR","Secondary RRR","num"),
                    ("Monthly End Projection","Secondary Projected LE","num","HL"),
                    ("Projected Secondary Achievement%","Secondary_Ach%","pct","HL"),
                    ("Total DF O/L","Total DF O/L","int"),
                    ("Zero DF O/L","Zero DF O/L","int"),
                    ("New Outlets","New Outlets","int"),
                ]

                for label, source_col, typ, *flags in items:
                    c = find_col(row.index,[source_col])
                    ws_label_fmt = hl_text_fmt if ("HL" in flags) else text_fmt
                    ws_val_num_fmt = hl_num_fmt if ("HL" in flags) else num_fmt
                    ws_val_pct_fmt = hl_pct_fmt if ("HL" in flags) else pct_fmt

                    ws.write(r,0,label,ws_label_fmt)
                    if not c:
                        ws.write(r,1,"",text_fmt)
                    else:
                        raw = row[c]
                        v = safe_float(raw)
                        if typ == "pct_recompute":
                            frac = get_fill_rate_fraction()
                            ws.write_number(r,1,frac, pct_fmt)
                        elif typ == "pct":
                            frac = to_fraction_for_excel(raw)
                            if frac is None:
                                ws.write(r,1,str(raw),text_fmt)
                            else:
                                ws.write_number(r,1,frac, ws_val_pct_fmt)
                        elif typ == "int":
                            v = 0 if v is None else v
                            ws.write_number(r,1,int(round(v)), int_fmt)
                        elif typ == "num":
                            if v is None:
                                ws.write(r,1,str(raw),text_fmt)
                            else:
                                ws.write_number(r,1,round(v,2),ws_val_num_fmt)
                        else:
                            ws.write(r,1,str(raw),text_fmt)
                    r += 1
                    if label in ("Projected Primary Achievement%","Projected Secondary Achievement%"):
                        r += 1  # spacer rows

                ws.set_column(0,0,45); ws.set_column(1,1,28)

        # ----------------- Generic writer for SO-PSR / Distributor -----------------
        def write_df_sheet(sheet_name, df, title, asm_colname=None):
            if df is None or df.empty: return
            if asm_colname and asm_colname in df.columns:
                df = df[df[asm_colname].astype(str).str.upper() == asm_name.upper()]
            if df.empty: return

            # fix Distributor Code to text
            code_col = find_col(df.columns, ["Distributor Code"])
            if code_col:
                df[code_col] = df[code_col].astype(str)

            ws = wb.add_worksheet(sheet_name)
            ws.merge_range(0,0,0,max(0,len(df.columns)-1), title, ds_title_fmt)

            # detect date columns
            date_cols = set()
            for col in df.columns:
                if looks_like_date_series(df[col]):
                    try:
                        df[col] = pd.to_datetime(df[col], errors="coerce").dt.date
                        date_cols.add(col)
                    except:
                        pass

            # headers
            for ci, col in enumerate(df.columns):
                ws.write(1, ci, str(col), ds_header_fmt)
                ws.set_column(ci, ci, 18)

            # rows
            for ri in range(len(df)):
                for ci, col in enumerate(df.columns):
                    val = df.iloc[ri, ci]
                    if pd.isna(val):
                        ws.write(ri+2, ci, "", ds_text_fmt); 
                        continue

                    # --- Force Distributor Code to TEXT ---
                    if str(col).strip().lower() in ["distributor code", "distributor_code"]:
                        ws.write_string(ri+2, ci, str(val), ds_text_fmt)
                        continue

                    # date columns → write date-only
                    if col in date_cols:
                        try:
                            d = pd.to_datetime(val).date()
                            ws.write_datetime(ri+2, ci, datetime.datetime(d.year, d.month, d.day), ds_date_fmt)
                        except:
                            ws.write(ri+2, ci, str(val), ds_text_fmt)
                        continue

                    # percent columns → correct handling
                    if is_percent_col(col):
                        frac = to_fraction_for_excel(val)
                        if frac is not None:
                            ws.write_number(ri+2, ci, frac, ds_pct_fmt)
                        else:
                            ws.write(ri+2, ci, "", ds_text_fmt)
                        continue

                    # integers heuristic
                    if any(k in str(col).lower() for k in ["count","qty","quantity","no. of","df o/l","outlets"]):
                        v = safe_float(val) or 0
                        ws.write_number(ri+2, ci, int(round(v)), ds_int_fmt)
                        continue

                    # normal numerics with 2 decimals
                    v = safe_float(val)
                    if v is None:
                        ws.write(ri+2, ci, str(val), ds_text_fmt)
                    else:
                        ws.write_number(ri+2, ci, round(v,2), ds_num_fmt)

        if "SO-PSR" in sheet_dfs and not sheet_dfs["SO-PSR"].empty:
            asm_colname = asm_col_map.get("SO-PSR") or detect_asm_col(sheet_dfs["SO-PSR"])
            write_df_sheet("SO-PSR", sheet_dfs["SO-PSR"].copy(), f"SO-PSR - {asm_name}", asm_colname)

        if "Distributor Wise Summary" in sheet_dfs and not sheet_dfs["Distributor Wise Summary"].empty:
            df_dist = sheet_dfs["Distributor Wise Summary"].copy()
            asm_colname = asm_col_map.get("Distributor Wise Summary") or detect_asm_col(df_dist)
            write_df_sheet("Distributor Wise Summary", df_dist, f"Distributor Wise Summary - {asm_name}", asm_colname)

    return output.getvalue()

# -----------------------------
# EMAIL HTML builder
# -----------------------------
def build_email_html(asm_name, logo_b64, asm_row):
    def g(keys):
        c = find_col(asm_row.index, keys)
        return safe_float(asm_row[c]) if c else None

    nf = lambda v: f"{(0 if v is None else v):,.2f}"

    # recompute Secondary Fill Rate for email display (keep your logic)
    sec_billing = g(['Secondary Billing']) or 0
    sec_booking = g(['Secondary Booking']) or 0
    sec_fill_pct = (sec_billing/sec_booking*100) if sec_booking else 0

    primary_html = f"""
    <div style="flex:1;padding:14px;background:#fff3cd;border:2px solid #d63384;border-radius:8px;">
      <h4 style="margin:0;color:#d63384;font-family:Cambria;">Primary</h4>
      <p><b>Target:</b> {nf(g(['Primary Target']))}</p>
      <p><b>MTD:</b> {nf(g(['Primary MTD']))}</p>
      <p><b>ASM GOLY:</b> {percent_str_for_email(g(['Primary Trgt vs Ach %']))}</p>
      <p><b>Monthly End Projection:</b> {nf(g(['Primary Projected LE']))}</p>
      <p><b>Projected Primary Achievement%:</b> {percent_str_for_email(g(['Primary_Ach%']))}</p>
    </div>"""

    secondary_html = f"""
    <div style="flex:1;padding:14px;background:#ffd6e7;border:2px solid #d63384;border-radius:8px;">
      <h4 style="margin:0;color:#ff6600;font-family:Cambria;">Secondary</h4>
      <p><b>Target:</b> {nf(g(['Secondary Target']))}</p>
      <p><b>Booking:</b> {nf(g(['Secondary Booking']))}</p>
      <p><b>Billing:</b> {nf(g(['Secondary Billing']))}</p>
      <p><b>Fill Rate:</b> {sec_fill_pct:.2f}%</p>
      <p><b>Monthly End Projection:</b> {nf(g(['Secondary Projected LE']))}</p>
      <p><b>Projected Secondary Achievement%:</b> {percent_str_for_email(g(['Secondary_Ach%']))}</p>
    </div>"""

    outlet_html = f"""
    <div style="flex:1;padding:14px;background:#e6f7ff;border:2px solid #3399ff;border-radius:8px;">
      <h4 style="margin:0;color:#0066cc;font-family:Cambria;">Outlets</h4>
      <p><b>Total DF O/L:</b> {int(g(['Total DF O/L']) or 0)}</p>
      <p><b>Zero DF O/L:</b> {int(g(['Zero DF O/L']) or 0)}</p>
      <p><b>New Outlets:</b> {int(g(['New Outlets']) or 0)}</p>
    </div>"""

    proj = g(['Secondary Projected LE']) or 0
    target = g(['Secondary Target']) or 0
    perc = (proj/target*100) if target else 0
    if perc >= 100:
        insight_msg = f"Excellent! You are projected to exceed your target ({nf(proj)} vs {nf(target)}). Keep pushing!"
    elif perc >= 90:
        insight_msg = f"You are close to target ({perc:.0f}%). A little more effort can help you cross it."
    elif perc >= 70:
        insight_msg = f"You are at {perc:.0f}% of target. Please motivate your distributors and team to reach closer."
    else:
        insight_msg = f"⚠️ You are only at {perc:.0f}% of target. Please review distributors and boost sales to avoid missing the target."

    today = datetime.date.today().strftime("%d-%b-%Y")
    return f"""
    <html><body style="font-family:Cambria;">
    <h2 style="color:#d63384;">HIPL — Daily MIS ({today})</h2>
    <p>Dear <b>{asm_name}</b>,</p>
    <div style='display:flex;gap:12px;'>{primary_html}{secondary_html}{outlet_html}</div>
    <div style="margin-top:14px;padding:12px;background:#e9f7ef;border:1px solid #2e8b57;border-radius:6px;font-family:Cambria;color:#123;">
      <p><b>Insight:</b> {insight_msg}</p>
    </div>
    <p style="margin-top:12px;">📎 Detailed MIS is attached.</p>
    <p>Thank you.<br>If any queries, please reach out to your respective MIS executive.<br><br><b>Best Regards,<br>MIS Team</b></p>
    </body></html>"""

# -----------------------------
# STREAMLIT UI
# -----------------------------
st.set_page_config(page_title="HIPL MIS Automation - Sampath MIS Executive", layout="wide")
st.title("📧 HIPL — ASM/RSM MIS Email Automation (Primary & Secondary Stats & Projection)")

uploaded_file = st.file_uploader("📁 Upload MIS Excel File (.xlsx)", type=["xlsx"])
if not uploaded_file:
    st.stop()

xls = pd.ExcelFile(uploaded_file)
needed = ["ASM","SO-PSR","Distributor Wise Summary"]
dfs = {s: pd.read_excel(xls, s) for s in needed if s in xls.sheet_names}
if "ASM" not in dfs:
    st.error("ASM sheet not found.")
    st.stop()

asm_col = detect_asm_col(dfs["ASM"])
asms = sorted(dfs["ASM"][asm_col].dropna().astype(str).unique()) if asm_col else []
sel_asm = st.selectbox("Select ASM/RSM", asms)

def build_all(asm_name):
    asm_row = dfs["ASM"].loc[dfs["ASM"][asm_col].astype(str).str.upper()==asm_name.upper()].iloc[0]
    html_preview = build_email_html(asm_name, to_base64_img(LOGO_PATH), asm_row)
    attach_bytes = build_attachment_bytes_for_asm(
        asm_name,
        dfs,
        {
            "ASM": asm_col,
            "SO-PSR": detect_asm_col(dfs.get("SO-PSR", pd.DataFrame())) if "SO-PSR" in dfs else None,
            "Distributor Wise Summary": detect_asm_col(dfs.get("Distributor Wise Summary", pd.DataFrame())) if "Distributor Wise Summary" in dfs else None
        }
    )
    return html_preview, attach_bytes

if sel_asm:
    html_preview, xbytes = build_all(sel_asm)
    st.components.v1.html(html_preview, height=900, scrolling=True)

    st.download_button("🔽 Download Attachment", xbytes, f"MIS_{sel_asm}.xlsx")

    if st.button("📤 Send Email"):
        to_email = EMAIL_MAP.get(sel_asm,{}).get("to")
        if not to_email:
            st.error("No email mapping for selected ASM/RSM.")
        else:
            token = get_access_token()
            status, resp = send_mail_via_graph(
                token, to_email, EMAIL_MAP.get(sel_asm,{}).get("cc",[]),
                f"HIPL Daily MIS — {sel_asm}", html_preview, xbytes, f"MIS_{sel_asm}.xlsx"
            )
            if status in (200,202): st.success("✅ Sent")
            else: st.error(f"❌ Failed: {resp[:200]}")

# -----------------------------
# Bulk Send with Progress Bar
# -----------------------------
st.markdown("---")
if st.button("📤 Bulk Send to All ASMs"):
    if not asms:
        st.error("No ASM/RSM values found in ASM sheet.")
    else:
        token = get_access_token()
        prog = st.progress(0)
        log = []
        for i, asm in enumerate(asms, start=1):
            try:
                html_body, attach_bytes = build_all(asm)
                to_email = EMAIL_MAP.get(asm,{}).get("to")
                if not to_email:
                    log.append(f"{asm}: ❌ No email mapping")
                else:
                    status, resp = send_mail_via_graph(
                        token, to_email, EMAIL_MAP.get(asm,{}).get("cc",[]),
                        f"HIPL Daily MIS — {asm}", html_body, attach_bytes, f"MIS_{asm}.xlsx"
                    )
                    if status in (200,202): log.append(f"{asm}: ✅ Sent")
                    else: log.append(f"{asm}: ❌ Failed ({resp[:100]})")
            except Exception as e:
                log.append(f"{asm}: ❌ Error {e}")
            prog.progress(i/len(asms))
        st.text_area("Bulk Send Log", "\n".join(log), height=280)
