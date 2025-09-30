# app.py
# Streamlit app to transform Odoo Excel -> ACT DBF, optionally append to last DBF
# Rules included:
# - 705000 -> 700000; 700000 aggregated per invoice; 178000 kept as-is
# - 451*** consolidated to one 451000 (non-zero VAT), DOCTYPE=3, OPCODE=FIXED, VATTAX=0
# - Add 0% VAT line when base 0% is detected from "Taxe d'origine" (label contains "0%"):
#     DOCTYPE=4, OPCODE=FIXED, VATCODE=211100, AMOUNTEUR=0, VATBASE=base0
# - VATCODE mapping: 21% -> 211400; 0% -> 211100 (editable in sidebar)
# - On 400000: VATBASE = base(HT) positive; VATTAX = - AMOUNTEUR(451000) (or 0 if none)
# - On 700000 line: VATIMPUT only here (211400 if base21>0 else 211100 if only 0%)
# - BOOKYEAR='J'; PERIOD from invoice date; AMOUNT=0
# - CURRCODE='0'; CURRAMOUNT=0; CUREURBASE=0; CURRATE=0
# - DOCORDER/COMMENTEXT/MATCHNO/OLDDATE empty
# - Clean strings 'nan'/'an' -> empty; numeric NaN -> blank; ACCOUNTGL strip '.0'
# - Writes DBF (dBase III, cp1252) with fixed schema

import io
import re
import struct
from datetime import datetime
from collections import defaultdict

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Odoo → ACT DBF (Monthly)", page_icon="📦", layout="wide")

# ----------------------------- UI / Sidebar -----------------------------------
st.title("📦 Odoo → ACT DBF (Monthly)")

with st.sidebar:
    st.header("VAT code mapping")
    vat_map_21 = st.text_input("Map 21% →", "211400")
    vat_map_00 = st.text_input("Map 0%  →", "211100")
    st.caption("Adjust these to your internal codes if needed.")

    st.header("Options")
    opt_append = st.checkbox("Append to a previous DBF (upload below)", value=False)
    opt_keep_other_70x = st.checkbox("Keep other 70x lines (besides 700000/178000)", value=True)

    st.header("Sanity checks")
    opt_check_balance = st.checkbox("Show unbalanced invoices", value=True)

col1, col2 = st.columns(2)
with col1:
    xls_file = st.file_uploader("Upload monthly Odoo Excel", type=["xlsx", "xls"])
with col2:
    prev_dbf_file = st.file_uploader("Upload previous DBF (optional, required if 'Append' is on)", type=["dbf"])

# ---------------------------- Schema + Helpers ---------------------------------
SCHEMA = [
    ("DOCTYPE", "C", 1, 0),
    ("DBKCODE", "C", 6, 0),
    ("DBKTYPE", "C", 1, 0),
    ("DOCNUMBER", "C", 8, 0),
    ("DOCORDER", "C", 3, 0),
    ("OPCODE", "C", 5, 0),
    ("ACCOUNTGL", "C", 8, 0),
    ("ACCOUNTRP", "C", 10, 0),
    ("BOOKYEAR", "C", 1, 0),
    ("PERIOD", "C", 2, 0),
    ("DATE", "D", 8, 0),
    ("DATEDOC", "D", 8, 0),
    ("DUEDATE", "D", 8, 0),
    ("COMMENT", "C", 40, 0),
    ("COMMENTEXT", "C", 35, 0),
    ("AMOUNT", "N", 17, 3),
    ("AMOUNTEUR", "N", 17, 3),
    ("VATBASE", "N", 17, 3),
    ("VATCODE", "C", 6, 0),
    ("CURRAMOUNT", "N", 17, 3),
    ("CURRCODE", "C", 3, 0),
    ("CUREURBASE", "N", 17, 3),
    ("VATTAX", "N", 17, 3),
    ("VATIMPUT", "C", 6, 0),
    ("CURRATE", "N", 12, 5),
    ("REMINDLEV", "N", 1, 0),
    ("MATCHNO", "C", 8, 0),
    ("OLDDATE", "D", 8, 0),
    ("ISMATCHED", "L", 1, 0),
    ("ISLOCKED", "L", 1, 0),
    ("ISIMPORTED", "L", 1, 0),
    ("ISPOSITIVE", "L", 1, 0),
    ("ISTEMP", "L", 1, 0),
    ("MEMOTYPE", "C", 1, 0),
    ("ISDOC", "L", 1, 0),
    ("DOCSTATUS", "C", 1, 0),
    ("DICFROM", "C", 16, 0),
]

def yyyymmdd(x):
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return ""
    try:
        dt = pd.to_datetime(x, dayfirst=True, errors="coerce")
        if pd.isna(dt):
            s = re.sub(r"[^0-9]", "", str(x))
            return s[:8]
        return dt.strftime("%Y%m%d")
    except Exception:
        s = re.sub(r"[^0-9]", "", str(x))
        return s[:8]

def to_number(x):
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return 0.0
    try:
        return float(x)
    except Exception:
        try:
            return float(str(x).replace(",", ".").replace(" ", ""))
        except Exception:
            return 0.0

def partner_code(name, maxlen=10):
    if name is None:
        return ""
    s = re.sub(r"[^A-Z0-9]", "", str(name).upper())
    return s[:maxlen]

def vat_code_from_label(label):
    if not isinstance(label, str):
        return ""
    m = re.search(r"(\d{1,2})\s*%", label)
    return m.group(1).zfill(2) if m else ""

def extract_account_code(val):
    if val is None:
        return ""
    s = str(val).strip()
    s = re.sub(r"\.0+$", "", s)  # strip trailing .0
    m = re.match(r"^(\d{3,8})", s)
    if m:
        return m.group(1)
    m2 = re.search(r"(\d{3,8})", s)
    return m2.group(1) if m2 else ""

def norm_docnum(val):
    if val is None:
        return ""
    try:
        f = float(val)
        if f.is_integer():
            return str(int(f))
    except Exception:
        pass
    s = str(val)
    return re.sub(r"\.0+$", "", s)

def is_zero_label(label):
    return isinstance(label, str) and "0%" in label.replace(" ", "")

def is_21_label(label):
    return isinstance(label, str) and "21%" in label.replace(" ", "")

def clean_str_nan(v):
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s.lower() in {"nan", "(nan)", "an"} else s

# ------------------------ DBF read/write (dBase III) ---------------------------
def write_dbf_bytes(records, schema=SCHEMA, encoding="cp1252"):
    version = 0x03
    today = datetime.today()
    header_len = 32 + 32 * len(schema) + 1
    record_len = 1 + sum(f[2] for f in schema)
    buf = io.BytesIO()
    # header
    buf.write(struct.pack("<BBBBIHH20x",
        version, today.year - 1900, today.month, today.day,
        len(records), header_len, record_len
    ))
    # field descriptors
    for name, ftype, flen, fdec in schema:
        name_b = name.encode(encoding)[:11]
        name_b = name_b + b"\x00" * (11 - len(name_b))
        buf.write(struct.pack("<11sc4xBB14x", name_b, ftype.encode("ascii"), flen, fdec))
    buf.write(b"\x0D")
    # records
    for rec in records:
        buf.write(b" ")
        for name, ftype, flen, fdec in schema:
            val = rec.get(name, None)
            if ftype == "C":
                s = "" if val is None else str(val)
                b = s.encode(encoding, errors="ignore")[:flen]
                buf.write(b + b" " * (flen - len(b)))
            elif ftype == "N":
                if val is None or (isinstance(val, float) and np.isnan(val)):
                    s = "".rjust(flen, " ")
                else:
                    try:
                        num = float(val)
                    except Exception:
                        num = 0.0
                    s = f"{num:.{fdec}f}" if fdec > 0 else f"{int(round(num))}"
                    if len(s) > flen:
                        s = "*" * flen
                    s = s.rjust(flen, " ")
                buf.write(s.encode("ascii", errors="ignore"))
            elif ftype == "D":
                s = "" if val is None else str(val)
                s = re.sub(r"[^0-9]", "", s)[:8].ljust(8, " ")
                buf.write(s.encode("ascii"))
            elif ftype == "L":
                c = b"T" if str(val).upper() in ("T","TRUE","Y","1") else (b"F" if str(val).upper() in ("F","FALSE","N","0") else b"?")
                buf.write(c)
            else:
                buf.write(b" " * flen)
    buf.write(b"\x1A")
    return buf.getvalue()

def read_dbf(file_bytes, encoding="cp1252"):
    data = file_bytes if isinstance(file_bytes, (bytes, bytearray)) else file_bytes.read()
    if len(data) < 32:
        return [], pd.DataFrame()
    n_records = struct.unpack("<I", data[4:8])[0]
    header_len = struct.unpack("<H", data[8:10])[0]
    record_len = struct.unpack("<H", data[10:12])[0]
    # fields
    fields = []
    pos = 32
    while pos < header_len - 1:
        fd = data[pos:pos+32]
        if fd[0] == 0x0D:
            break
        name = fd[0:11].split(b"\x00", 1)[0].decode(encoding, errors="ignore").strip()
        ftype = chr(fd[11])
        flen = fd[16]
        fdec = fd[17]
        fields.append((name, ftype, flen, fdec))
        pos += 32
    # records
    recs = []
    rec_start = header_len
    for i in range(n_records):
        start = rec_start + i * record_len
        rec = data[start:start+record_len]
        if len(rec) < record_len or rec[0:1] == b"*":
            continue
        offset = 1
        vals = {}
        for (name, ftype, flen, fdec) in fields:
            raw = rec[offset:offset+flen]
            offset += flen
            if ftype == "C":
                vals[name] = raw.decode("cp1252", errors="ignore").rstrip()
            elif ftype == "N":
                s = raw.decode("ascii", errors="ignore").strip()
                if s == "":
                    vals[name] = None
                else:
                    try:
                        vals[name] = float(s) if fdec > 0 else int(s)
                    except Exception:
                        try:
                            vals[name] = float(s)
                        except Exception:
                            vals[name] = s
            elif ftype == "D":
                vals[name] = raw.decode("ascii", errors="ignore").strip()
            elif ftype == "L":
                c = raw[:1]
                vals[name] = True if c in b"YyTt" else (False if c in b"NnFf" else None)
            else:
                vals[name] = raw
        recs.append(vals)
    return fields, pd.DataFrame(recs)

# -------------------------- Transformation logic -------------------------------
EXPECTED_COLS = {
    "num": ["numéro", "numero", "n°", "facture", "invoice", "num"],
    "acc": ["compte", "account", "general account", "gl"],
    "partner": ["partenaire", "client", "customer", "partner"],
    "date": ["date"],
    "inv_date": ["date de facturation", "date facture", "invoice date"],
    "due_date": ["date d'échéance", "date echeance", "due date"],
    "label": ["libellé", "libelle", "label", "description"],
    "debit": ["débit", "debit"],
    "credit": ["crédit", "credit"],
    "tax": ["taxe d'origine", "taxe", "tax", "vat"]
}

def find_col(cols, names):
    low = {c.lower(): c for c in cols}
    for n in names:
        if n in low:
            return low[n]
    # fuzzy fallback
    for c in cols:
        cl = c.lower()
        for n in names:
            if n in cl:
                return c
    return None

def transform_excel(xdf: pd.DataFrame, keep_other_70x=True, map21="211400", map00="211100"):
    # locate columns
    cols = list(xdf.columns)
    CNUM   = find_col(cols, EXPECTED_COLS["num"])
    CACC   = find_col(cols, EXPECTED_COLS["acc"])
    CPART  = find_col(cols, EXPECTED_COLS["partner"])
    CDATE  = find_col(cols, EXPECTED_COLS["date"])
    CINVD  = find_col(cols, EXPECTED_COLS["inv_date"]) or CDATE
    CDUED  = find_col(cols, EXPECTED_COLS["due_date"])
    CLABEL = find_col(cols, EXPECTED_COLS["label"])
    CDEBIT = find_col(cols, EXPECTED_COLS["debit"])
    CCRED  = find_col(cols, EXPECTED_COLS["credit"])
    CTAX   = find_col(cols, EXPECTED_COLS["tax"])

    missing = [n for n,v in {"Numéro":CNUM,"Compte":CACC,"Partenaire":CPART,"Date":CDATE,"Débit":CDEBIT,"Crédit":CCRED,"Libellé":CLABEL}.items() if v is None]
    if missing:
        raise ValueError(f"Missing required columns in Excel: {', '.join(missing)}")

    records = []
    per_inv_sum = defaultdict(float)

    for inv, g in xdf.groupby(CNUM, dropna=True):
        inv_str = norm_docnum(inv)[:8]
        partner_name = str(g[CPART].dropna().iloc[0]) if not g[CPART].dropna().empty else ""
        partner_ref  = partner_code(partner_name, 10)
        date_   = yyyymmdd(g[CDATE].dropna().iloc[0]) if not g[CDATE].dropna().empty else ""
        invdate = yyyymmdd(g[CINVD].dropna().iloc[0]) if not g[CINVD].dropna().empty else date_
        duedate = yyyymmdd(g[CDUED].dropna().iloc[0]) if CDUED and not g[CDUED].dropna().empty else ""
        period  = invdate[4:6] if invdate else ""

        rows_400, rows_700000, rows_178000, rows_other70, rows_vat, rows_misc = [], [], [], [], [], []
        vat_codes_seen = []

        for _, row in g.iterrows():
            acct = extract_account_code(row[CACC])
            if acct == "705000":
                acct = "700000"
            debit  = to_number(row[CDEBIT]) if CDEBIT else 0.0
            credit = to_number(row[CCRED]) if CCRED else 0.0
            amount = round(debit if debit else -credit, 3)
            label  = str(row[CLABEL]) if not pd.isna(row[CLABEL]) else ""
            vlabel = str(row[CTAX]) if CTAX and not pd.isna(row[CTAX]) else ""
            vcode  = vat_code_from_label(vlabel)
            if vcode:
                vat_codes_seen.append(vcode)

            if acct.startswith("400"):
                rows_400.append((acct, amount, label))
            elif acct == "700000":
                rows_700000.append((acct, amount, label, vlabel))
            elif acct == "178000":
                rows_178000.append((acct, amount, label, vlabel))
            elif acct.startswith("700"):
                rows_other70.append((acct, amount, label, vlabel))
            elif acct.startswith("451"):
                rows_vat.append((acct, amount, label))
            else:
                rows_misc.append((acct, amount, label))

        def base_for_rate(rows, pred):
            tot = 0.0
            for (_a, amt, _l, vlabel) in rows:
                if pred(vlabel):
                    tot += abs(amt)
            return round(tot, 3)

        base0  = base_for_rate(rows_700000+rows_other70+rows_178000, is_zero_label)
        base21 = base_for_rate(rows_700000+rows_other70+rows_178000, is_21_label)
        base_all = round(base0 + base21, 3)

        vat_amount_signed = round(sum(amt for (_a, amt, _l) in rows_vat), 3) if rows_vat else 0.0

        def base_record(acct, amount, label, doctype, opcode=""):
            return {
                "DOCTYPE": doctype,
                "DBKCODE": "VEN",
                "DBKTYPE": "2",
                "DOCNUMBER": inv_str,
                "DOCORDER": "",                       # empty
                "OPCODE": opcode,
                "ACCOUNTGL": (acct or "")[:8],
                "ACCOUNTRP": partner_ref,
                "BOOKYEAR": "J",
                "PERIOD": period,
                "DATE": date_,
                "DATEDOC": invdate,
                "DUEDATE": duedate,
                "COMMENT": (partner_name[:40] if partner_name else label[:40]),
                "COMMENTEXT": "",                    # empty
                "AMOUNT": 0.0,                       # force 0
                "AMOUNTEUR": amount,                 # balance here
                "VATBASE": 0.0,                      # set on 400000/TVA lines
                "VATCODE": "",
                "CURRAMOUNT": 0.0,
                "CURRCODE": "0",
                "CUREURBASE": 0.0,
                "VATTAX": 0.0,                       # set on 400000 only for non-zero VAT
                "VATIMPUT": "",                      # set on 700000 only
                "CURRATE": 0.0,
                "REMINDLEV": 0,
                "MATCHNO": "",
                "OLDDATE": "",
                "ISMATCHED": "F",
                "ISLOCKED": "F",
                "ISIMPORTED": "F",
                "ISPOSITIVE": "T",
                "ISTEMP": "F",
                "MEMOTYPE": "",
                "ISDOC": "?",
                "DOCSTATUS": "",
                "DICFROM": "",
            }

        # 400xxx lines
        for acct, amount, label in rows_400:
            rec = base_record(acct, amount, label, doctype="1")
            rec["VATBASE"] = base_all
            rec["VATTAX"]  = -vat_amount_signed if vat_amount_signed else 0.0
            records.append(rec)
            per_inv_sum[inv_str] += amount

        # 700000 consolidated
        if rows_700000:
            total_700000 = round(sum(amt for (_a, amt, _l, _v) in rows_700000), 3)
            label700 = rows_700000[0][2] if rows_700000[0][2] else "VENTES 700000"
            rec = base_record("700000", total_700000, label700, doctype="3")
            # VATIMPUT only here
            if base21 > 0:
                rec["VATIMPUT"] = map21
            elif base0 > 0:
                rec["VATIMPUT"] = map00
            records.append(rec)
            per_inv_sum[inv_str] += total_700000

        # 178000 kept as-is
        for acct, amount, label, _v in rows_178000:
            rec = base_record(acct, amount, label, doctype="3")
            records.append(rec)
            per_inv_sum[inv_str] += amount

        # other 70x kept (optional)
        if keep_other_70x:
            for acct, amount, label, _v in rows_other70:
                rec = base_record(acct, amount, label, doctype="3")
                records.append(rec)
                per_inv_sum[inv_str] += amount

        # misc
        for acct, amount, label in rows_misc:
            rec = base_record(acct, amount, label, doctype="3")
            records.append(rec)
            per_inv_sum[inv_str] += amount

        # VAT lines
        if base0 > 0:
            rec0 = base_record("", 0.0, "TVA 0%", doctype="4", opcode="FIXED")
            rec0["VATBASE"] = base0
            rec0["VATCODE"] = map00
            records.append(rec0)
        if vat_amount_signed:
            rec21 = base_record("451000", vat_amount_signed, "TVA", doctype="3", opcode="FIXED")
            rec21["VATBASE"] = base21 if base21 > 0 else base_all
            rec21["VATTAX"] = 0.0
            # map 21% if seen, else leave as empty if not detectable
            rec21["VATCODE"] = map21 if base21 > 0 else ""
            records.append(rec21)
            per_inv_sum[inv_str] += vat_amount_signed

    # cleanup: strings 'nan'/'an' -> empty; ACCOUNTGL remove '.0'
    for rec in records:
        for k in list(rec.keys()):
            v = rec[k]
            if isinstance(v, str):
                rec[k] = clean_str_nan(v)
        # ensure ACCOUNTGL text without decimals/spaces
        ag = rec.get("ACCOUNTGL", "")
        ag = re.sub(r"\.0+$", "", ag or "").replace(" ", "")[:8]
        rec["ACCOUNTGL"] = ag

    # detect unbalanced
    unbalanced = {k: v for k, v in per_inv_sum.items() if round(v, 2) != 0}
    return records, unbalanced

# ------------------------------- Main action -----------------------------------
if xls_file is not None:
    try:
        xdf = pd.read_excel(xls_file)
        st.success(f"Excel loaded: {xdf.shape[0]} rows, {xdf.shape[1]} columns")
        st.dataframe(xdf.head(20))
    except Exception as e:
        st.error(f"Failed to read Excel: {e}")
        st.stop()

    if opt_append and prev_dbf_file is None:
        st.warning("Upload a previous DBF to append, or uncheck 'Append'.")
        st.stop()

    if st.button("🚀 Generate DBF"):
        try:
            recs, unbalanced = transform_excel(
                xdf,
                keep_other_70x=opt_keep_other_70x,
                map21=vat_map_21.strip(),
                map00=vat_map_00.strip(),
            )

            # If appending to previous DBF
            if opt_append and prev_dbf_file is not None:
                fields_prev, df_prev = read_dbf(prev_dbf_file.read())
                # Convert recs to DataFrame in schema order
                df_new = pd.DataFrame([{k: r.get(k, None) for (k,_,_,_) in SCHEMA} for r in recs])
                df_out = pd.concat([df_prev, df_new], ignore_index=True)
                # Recreate records to write (keep schema)
                recs_to_write = df_out.to_dict(orient="records")
            else:
                recs_to_write = recs

            dbf_bytes = write_dbf_bytes(recs_to_write, SCHEMA)

            st.success("DBF generated successfully.")
            st.download_button("⬇️ Download DBF", dbf_bytes, file_name="export_ACT_from_Odoo.dbf", mime="application/octet-stream")

            # Also export a CSV/Excel preview
            df_preview = pd.DataFrame([{k: r.get(k, None) for (k,_,_,_) in SCHEMA} for r in recs_to_write])
            csv_bytes = df_preview.to_csv(index=False).encode("utf-8-sig")
            xlsx_buf = io.BytesIO()
            with pd.ExcelWriter(xlsx_buf, engine="openpyxl") as w:
                df_preview.to_excel(w, index=False, sheet_name="ACT", freeze_panes=(1,0))
            st.download_button("⬇️ Download Preview (CSV)", csv_bytes, file_name="preview_ACT.csv", mime="text/csv")
            st.download_button("⬇️ Download Preview (XLSX)", xlsx_buf.getvalue(), file_name="preview_ACT.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            if opt_check_balance:
                if unbalanced:
                    st.warning(f"Unbalanced invoices (sum of AMOUNTEUR ≠ 0): {len(unbalanced)}")
                    st.dataframe(pd.DataFrame([{"DOCNUMBER":k, "SUM_AMOUNTEUR":round(v,2)} for k,v in unbalanced.items()]))
                else:
                    st.info("All invoices are balanced (sum AMOUNTEUR = 0).")

        except Exception as e:
            st.error(f"Generation failed: {e}")
else:
    st.info("Upload your monthly Odoo Excel to begin.")
