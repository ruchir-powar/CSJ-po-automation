import re
from io import BytesIO

import pandas as pd
import streamlit as st

st.set_page_config(page_title="PO Automation", layout="wide")

# =========================
# Excel I/O (robust helpers)
# =========================
def safe_read_excel(uploaded_file):
    """Read .xlsx/.xls with helpful fallbacks and clear errors."""
    name = getattr(uploaded_file, "name", "") or ""
    ext = name.lower().split(".")[-1]

    try:
        if ext in ("xlsx", "xlsm", "xltx", "xltm", ""):
            return pd.read_excel(uploaded_file, engine="openpyxl")
        elif ext == "xls":
            return pd.read_excel(uploaded_file, engine="xlrd")  # requires xlrd==1.2.0
        else:
            st.error(f"Unsupported file type: .{ext}. Please upload a .xlsx file.")
            st.stop()
    except ImportError as e:
        msg = str(e)
        if "openpyxl" in msg:
            st.error(
                "Missing dependency **openpyxl** for reading .xlsx. "
                "Add it to requirements.txt and redeploy."
            )
        elif "xlrd" in msg:
            st.error(
                "Missing dependency **xlrd==1.2.0** for reading legacy .xls. "
                "Either convert to .xlsx or add xlrd==1.2.0 to requirements."
            )
        else:
            st.error(f"Excel read error: {e}")
        st.stop()


def safe_excel_writer(buffer: BytesIO):
    """Return a working ExcelWriter using openpyxl or xlsxwriter."""
    try:
        import openpyxl  # noqa: F401
        return pd.ExcelWriter(buffer, engine="openpyxl")
    except ImportError:
        try:
            import xlsxwriter  # noqa: F401
            return pd.ExcelWriter(buffer, engine="xlsxwriter")
        except ImportError:
            st.error("Missing Excel writer engine. Install **openpyxl** or **xlsxwriter**.")
            st.stop()


# =====================
# Columns / Field order
# =====================
OUTPUT_COLUMNS = [
    "SrNo",
    "StyleCode",
    "Unnamed: 2",                 # blank header in export
    "ItemSize",
    "OrderQty",
    "OrderItemPcs",
    "Metal",
    "Tone",
    "ItemPoNo",
    "ItemRefNo",
    "StockType",
    "MakeType",
    "CustomerProductionInstruction",
    "SpecialRemarks",
    "DesignProductionInstruction",
    "StampInstruction",
    "OrderGroup",
    "Certificate",
    "SKUNo",
    "Basestoneminwt",
    "Basestonemaxwt",
    "Basemetalminwt",
    "Basemetalmaxwt",
    "Productiondeliverydate",
    "Expecteddeliverydate",
    "Unnamed: 25",                # blank header in export
    "SetPrice",
    "StoneQuality",               # must remain blank
]


# =========
# Utilities
# =========
def detect_header_row(df: pd.DataFrame) -> int:
    """Find the row that actually contains SR NO. (and usually STYLE NO. etc.)."""
    max_check = min(10, len(df))
    for i in range(max_check):
        row = df.iloc[i]
        vals_upper = [str(x).strip().upper() for x in row.values]
        if "SR NO." in vals_upper:
            return i
    return 0


def normalize_input_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Standardize column names so the rest of the code always finds:
    - 'DESIGN CODE'
    - 'Purity / Color'
    Works for:
    - Files with DESIGN CODE already
    - Files with STYLE NO. instead
    - Files with PURITY + GOLD COLOR instead of Purity / Color
    """
    upper_map = {c.upper(): c for c in df.columns}

    # --- DESIGN CODE alias ---
    if "DESIGN CODE" not in df.columns:
        for cand in ["DESIGN CODE", "STYLE NO.", "STYLE NO", "STYLE", "STYLE CODE", "STYLE CODE."]:
            if cand in upper_map:
                df = df.rename(columns={upper_map[cand]: "DESIGN CODE"})
                break

    # --- Purity / Color alias or construction ---
    if "Purity / Color" not in df.columns:
        if "PURITY / COLOR" in upper_map:
            df = df.rename(columns={upper_map["PURITY / COLOR"]: "Purity / Color"})
        elif "PURITY" in upper_map and "GOLD COLOR" in upper_map:
            p_col = upper_map["PURITY"]
            gc_col = upper_map["GOLD COLOR"]
            df["Purity / Color"] = (
                df[p_col].astype(str).str.strip()
                + " / "
                + df[gc_col].astype(str).str.strip()
            )

    return df


def clean_input_df(df: pd.DataFrame) -> pd.DataFrame:
    """Use detected header row as header and normalize column names."""
    header_idx = detect_header_row(df)
    header = df.iloc[header_idx]
    new_cols = [(str(header[c]) if not pd.isna(header[c]) else str(c)) for c in df.columns]
    data = df.iloc[header_idx + 1 :].copy()
    data.columns = new_cols
    data = data.reset_index(drop=True)
    data = normalize_input_columns(data)
    return data


def make_unique_columns(columns):
    """Fix duplicate column names for preview so st.dataframe doesn't crash."""
    new_cols, counts = [], {}
    for c in columns:
        if c not in counts:
            counts[c] = 0
            new_cols.append(str(c))
        else:
            counts[c] += 1
            new_cols.append(f"{c}_{counts[c]}")
    return new_cols


def transform_stylecode(design_code: str):
    """
    Map input DESIGN CODE to Order Import StyleCode (hyphens/prefix rules).
    Skip codes not wanted in output (e.g., ERA1386).
    """
    if not isinstance(design_code, str):
        design_code = str(design_code or "")
    code = design_code.strip().upper()

    # Skip list (observed in sample comparison)
    if code in {"ERA1386"}:
        return None

    # 1) Trailing 'A' → '-A'
    if code.endswith("A") and "-" not in code:
        return code[:-1] + "-A"

    # 2) Special double-letter prefixes
    if code.startswith("JL") and code[2:].startswith("ERB"):
        return "JL-" + code[2:]
    if code.startswith("JW") and code[2:].startswith("ERB"):
        return "JW-" + code[2:]

    # 3) Generic J prefix
    if code.startswith("J") and not code.startswith("JL") and not code.startswith("JW"):
        if "-" not in code:
            return "J-" + code[1:]

    # 4) L/T/W prefixes before EAR/ERA/ERB
    if code.startswith("L") and code[1:4] == "EAR" and "-" not in code:
        return "L-" + code[1:]
    if code.startswith("T") and code[1:4] == "ERA" and "-" not in code:
        return "T-" + code[1:]
    if code.startswith("W") and code[1:3] == "ER" and "-" not in code:
        return "W-" + code[1:]

    # Default
    return code


def parse_metal_tone(purity_color: str):
    """
    '18KT / Yellow' -> ('GA18', 'Y', 18)
    '14KT / Yellow' -> ('GA14', 'Y', 14)
    Fallback: GA18, tone '', 18
    """
    if not isinstance(purity_color, str):
        purity_color = str(purity_color or "")
    s = purity_color.upper()

    m = re.search(r"(\d+)\s*KT", s)
    kt = int(m.group(1)) if m else 18

    if "YELLOW" in s or "YG" in s:
        tone = "Y"
    elif "WHITE" in s or "WG" in s:
        tone = "W"
    elif "ROSE" in s or "RG" in s:
        tone = "R"
    else:
        tone = ""

    metal = f"GA{kt}"
    return metal, tone, kt


# ==========================
# Core transform (no sizes)
# ==========================
def transform_to_order_import(
    clean_df: pd.DataFrame,
    order_group: str,
    cust_instr_template: str,
    remark_prefix: str,
) -> pd.DataFrame:
    """
    Uses only: SR NO., DESIGN CODE, CATEGORY, Purity / Color,
               DIA PCS, DIA WT, DIA QUALITY, REMARK, ORDER PCS

    Rules:
    - Ignore ring sizes entirely.
    - ORDER PCS = N → create N rows; each row has OrderQty=1 and OrderItemPcs=1.
    - SpecialRemarks format:
        NO 2 TONE RHODIUM ON METAL PART, [REMARK text if any][, 18/14kt gilit text if applicable]
    - StoneQuality must be blank.
    """
    if "ORDER PCS" not in clean_df.columns:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    rows = []

    for _, row in clean_df.iterrows():
        design_code = row.get("DESIGN CODE", "")
        style_code = transform_stylecode(design_code)
        if not style_code:
            continue  # skip unwanted codes

        purity_str = str(row.get("Purity / Color", "") or "")
        purity_upper = purity_str.upper()
        metal, tone, kt = parse_metal_tone(purity_str)

        dia_quality = str(row.get("DIA QUALITY", "") or "").strip()  # not written to StoneQuality
        remark_input = str(row.get("REMARK", "") or "").strip()

        # SpecialRemarks: prefix, remark, then gilit
        if "18KT" in purity_upper and "YELLOW" in purity_upper:
            gilit_text = "18kt Gilit for yellow gold"
        elif "14KT" in purity_upper and "YELLOW" in purity_upper:
            gilit_text = "14kt Gilit for yellow gold"
        else:
            gilit_text = ""

        parts = [remark_prefix]
        if remark_input:
            parts.append(remark_input)
        if gilit_text and gilit_text.strip().lower() not in remark_input.strip().lower():
            parts.append(gilit_text)
        special_remarks = ", ".join(parts)

        # Customer Production Instruction (keep template behavior)
        if "{quality}" in cust_instr_template:
            customer_instr = cust_instr_template.format(quality=dia_quality)
        else:
            customer_instr = cust_instr_template

        qty = row.get("ORDER PCS")
        if pd.isna(qty):
            continue
        try:
            qty_val = float(qty)
        except Exception:
            continue
        if qty_val <= 0:
            continue

        for _ in range(int(qty_val)):
            rows.append(
                {
                    "SrNo": None,
                    "StyleCode": style_code,
                    "Unnamed: 2": None,          # blanks required by template
                    "ItemSize": None,             # size ignored
                    "OrderQty": 1,
                    "OrderItemPcs": 1,
                    "Metal": metal,
                    "Tone": tone,
                    "ItemPoNo": None,
                    "ItemRefNo": None,
                    "StockType": None,
                    "MakeType": None,
                    "CustomerProductionInstruction": customer_instr,
                    "SpecialRemarks": special_remarks,
                    "DesignProductionInstruction": None,
                    "StampInstruction": f"{kt}KT & DIA WT",
                    "OrderGroup": order_group,
                    "Certificate": None,
                    "SKUNo": None,
                    "Basestoneminwt": None,
                    "Basestonemaxwt": None,
                    "Basemetalminwt": None,
                    "Basemetalmaxwt": None,
                    "Productiondeliverydate": None,
                    "Expecteddeliverydate": None,
                    "Unnamed: 25": None,          # blanks required by template
                    "SetPrice": None,
                    "StoneQuality": None,          # ALWAYS blank
                }
            )

    out_df = pd.DataFrame(rows)
    if not out_df.empty:
        out_df["SrNo"] = range(1, len(out_df) + 1)

    # force required columns (order + presence)
    for col in OUTPUT_COLUMNS:
        if col not in out_df.columns:
            out_df[col] = None

    return out_df[OUTPUT_COLUMNS]


# ============
# Streamlit UI
# ============
st.title("PO Automation – Order Import Sheet Generator")

uploaded_file = st.file_uploader(
    "Upload RAW order sheet (e.g., CSJ UPLOAD FILE- 17-11-2025.xlsx)", type=["xlsx", "xls"]
)

if uploaded_file is not None:
    raw_df = safe_read_excel(uploaded_file)
    clean_df = clean_input_df(raw_df)

    # Preview only the columns we actually use
    preview = clean_df.copy()
    if preview.columns.duplicated().any():
        preview.columns = make_unique_columns(preview.columns)

    important_cols = [
        "SR NO.",
        "DESIGN CODE",
        "CATEGORY",
        "Purity / Color",
        "DIA PCS",
        "DIA WT",
        "DIA QUALITY",
        "REMARK",
        "ORDER PCS",
    ]
    show_cols = [c for c in important_cols if c in preview.columns]

    st.subheader("Cleaned Input Preview (used columns)")
    if show_cols:
        st.dataframe(preview[show_cols].head(30))
    else:
        st.warning("Could not find expected columns in the uploaded file.")

    if "DESIGN CODE" not in clean_df.columns:
        st.error(
            "This does not look like the RAW order sheet "
            "(no 'DESIGN CODE' / 'STYLE NO.' column after cleaning). "
            "Please check the file format."
        )
    else:
        default_order_group = uploaded_file.name.rsplit(".", 1)[0]
        order_group = st.text_input("Order Group / PO No", value=default_order_group)

        st.sidebar.header("Instruction Templates")
        cust_instr_template = st.sidebar.text_area(
            "Customer Production Instruction Template",
            value="IGI CERTI-CO-BRANDING( {quality}), HALLMARK-BIS, REQ. AJ-STYLE NO.ON IGI CERTI",
            height=100,
        )
        remark_prefix = st.sidebar.text_area(
            "Special Remarks Base Prefix",
            value="NO 2 TONE RHODIUM ON METAL PART",
            height=80,
        )

        if st.button("Generate Order Import Sheet"):
            # ---- split input into 18KT / 14KT / Others based on Purity / Color ----
            purity_series = clean_df.get("Purity / Color", "").astype(str).str.upper()

            mask_18 = purity_series.str.contains("18KT", na=False)
            mask_14 = purity_series.str.contains("14KT", na=False)

            df_18 = clean_df[mask_18].copy()
            df_14 = clean_df[mask_14].copy()
            df_other = clean_df[~(mask_18 | mask_14)].copy()

            result_18 = transform_to_order_import(df_18, order_group, cust_instr_template, remark_prefix)
            result_14 = transform_to_order_import(df_14, order_group, cust_instr_template, remark_prefix)
            result_other = transform_to_order_import(df_other, order_group, cust_instr_template, remark_prefix)

            if result_18.empty and result_14.empty and result_other.empty:
                st.error("No rows generated. Check that 'ORDER PCS' has values > 0.")
            else:
                st.subheader("Generated Order Import Sheets (Preview)")
                if not result_18.empty:
                    st.write("**18KT Sheet Preview**")
                    st.dataframe(result_18.head(30))
                if not result_14.empty:
                    st.write("**14KT Sheet Preview**")
                    st.dataframe(result_14.head(30))
                if not result_other.empty:
                    st.write("**Others Sheet Preview**")
                    st.dataframe(result_other.head(30))

                buffer = BytesIO()
                with safe_excel_writer(buffer) as writer:
                    if not result_18.empty:
                        export_18 = result_18.copy().rename(
                            columns={"Unnamed: 2": "", "Unnamed: 25": ""}
                        )
                        export_18.to_excel(writer, index=False, sheet_name="18KT")

                    if not result_14.empty:
                        export_14 = result_14.copy().rename(
                            columns={"Unnamed: 2": "", "Unnamed: 25": ""}
                        )
                        export_14.to_excel(writer, index=False, sheet_name="14KT")

                    if not result_other.empty:
                        export_other = result_other.copy().rename(
                            columns={"Unnamed: 2": "", "Unnamed: 25": ""}
                        )
                        export_other.to_excel(writer, index=False, sheet_name="Others")

                buffer.seek(0)

                st.download_button(
                    label="📥 Download Order Import Workbook (.xlsx)",
                    data=buffer.getvalue(),
                    file_name=f"{order_group}_Order_Import_Split_18_14.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
else:
    st.info("Upload the RAW order Excel (not the already formatted Order Import file) to start.")
