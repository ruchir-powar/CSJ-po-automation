import re
from io import BytesIO

import pandas as pd
import streamlit as st

st.set_page_config(page_title="PO Automation", layout="wide")

# ---------- CONSTANTS ---------- #
OUTPUT_COLUMNS = [
    "SrNo",
    "StyleCode",
    "Unnamed: 2",
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
    "Unnamed: 25",
    "SetPrice",
    "StoneQuality",
]


# ---------- HELPERS ---------- #
def detect_header_row(df: pd.DataFrame) -> int:
    """Find the row that actually contains SR NO. + DESIGN CODE."""
    max_check = min(10, len(df))
    for i in range(max_check):
        row = df.iloc[i]
        vals_upper = [str(x).strip().upper() for x in row.values]
        if "SR NO." in vals_upper and "DESIGN CODE" in vals_upper:
            return i
    return 0


def clean_input_df(df: pd.DataFrame) -> pd.DataFrame:
    """Use detected header row as header."""
    header_idx = detect_header_row(df)
    header = df.iloc[header_idx]

    new_cols = []
    for c in df.columns:
        val = header[c]
        new_cols.append(str(val) if not pd.isna(val) else str(c))

    data = df.iloc[header_idx + 1 :].copy()
    data.columns = new_cols
    return data.reset_index(drop=True)


def make_unique_columns(columns):
    """Fix duplicate column names for preview so st.dataframe doesn't crash."""
    new_cols = []
    counts = {}
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
    Convert DESIGN CODE from input into StyleCode used in Order Import.
    Matches the pattern seen in your sample.
    Returns None if the design should be skipped (e.g. ERA1386).
    """
    if not isinstance(design_code, str):
        design_code = str(design_code or "")
    code = design_code.strip().upper()

    # Designs present in input but not used in expected output
    BLACKLIST = {"ERA1386"}
    if code in BLACKLIST:
        return None

    # Case 1: codes that end with A -> put -A at the end
    #  EAR7757A -> EAR7757-A,  ERC4215A -> ERC4215-A
    if code.endswith("A") and "-" not in code:
        return code[:-1] + "-A"

    # Case 2: special double-letter prefixes
    #  JLERB3500 -> JL-ERB3500
    #  JWERB3809 -> JW-ERB3809
    if code.startswith("JL") and code[2:].startswith("ERB"):
        return "JL-" + code[2:]
    if code.startswith("JW") and code[2:].startswith("ERB"):
        return "JW-" + code[2:]

    # Case 3: generic J prefix
    #  JEPEB0270 -> J-EPEB0270
    #  JERB3397  -> J-ERB3397
    #  JPEB0753  -> J-PEB0753
    #  JXEPEB0317 -> J-XEPEB0317
    if code.startswith("J") and not code.startswith("JL") and not code.startswith("JW"):
        if "-" not in code:
            return "J-" + code[1:]

    # Case 4: L/T/W prefixes before EAR/ERA/ERB
    #  LEAR10711 -> L-EAR10711
    #  TERA0923  -> T-ERA0923
    #  WERA0351  -> W-ERA0351
    #  WERB3000  -> W-ERB3000
    if code.startswith("L") and code[1:4] == "EAR" and "-" not in code:
        return "L-" + code[1:]
    if code.startswith("T") and code[1:4] == "ERA" and "-" not in code:
        return "T-" + code[1:]
    if code.startswith("W") and code[1:3] == "ER" and "-" not in code:
        return "W-" + code[1:]

    # Default: keep as-is
    return code


def parse_metal_tone(purity_color: str):
    """
    Convert '18KT / Yellow' -> ('GA18', 'Y', 18)
    Convert '14KT / Yellow' -> ('GA14', 'Y', 14)
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


# ---------- CORE TRANSFORM (NO RING SIZE) ---------- #
def transform_to_order_import(
    clean_df: pd.DataFrame,
    order_group: str,
    cust_instr_template: str,
    remark_prefix: str,
) -> pd.DataFrame:
    """
    Conversion logic WITHOUT ring sizes.

    Uses only:
    SR NO., DESIGN CODE, CATEGORY, Purity / Color,
    DIA PCS, DIA WT, DIA QUALITY, REMARK, ORDER PCS.

    For each row:
    - Read ORDER PCS = N
    - Create N rows with OrderQty = 1 and OrderItemPcs = 1
    - SpecialRemarks format:
      NO 2 TONE RHODIUM ON METAL PART, [REMARK text if any][, 18/14kt gilit text if applicable]
    """

    rows = []

    if "ORDER PCS" not in clean_df.columns:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    for _, row in clean_df.iterrows():
        design_code = row.get("DESIGN CODE", "")
        style_code = transform_stylecode(design_code)
        if not style_code:
            continue  # e.g. ERA1386 skipped

        purity = row.get("Purity / Color", "")
        purity_str = str(purity or "")
        purity_upper = purity_str.upper()

        metal, tone, kt = parse_metal_tone(purity_str)
        dia_quality = str(row.get("DIA QUALITY", "") or "").strip()
        remark_input = str(row.get("REMARK", "") or "").strip()

        # ---------- SpecialRemarks: PREFIX, REMARK, then GILIT ---------- #
        # Decide gilit phrase from purity
        if "18KT" in purity_upper and "YELLOW" in purity_upper:
            gilit_text = "18kt Gilit for yellow gold"
        elif "14KT" in purity_upper and "YELLOW" in purity_upper:
            gilit_text = "14kt Gilit for yellow gold"
        else:
            gilit_text = ""

        parts = [remark_prefix]  # always start with base

        # 1) add REMARK column text if present
        if remark_input:
            parts.append(remark_input)

        # 2) add gilit text if applicable and not already inside remark
        if gilit_text and gilit_text.strip().lower() not in remark_input.strip().lower():
            parts.append(gilit_text)

        special_remarks = ", ".join(parts)
        # ---------------------------------------------------------------- #

        # Customer Production Instruction
        if "{quality}" in cust_instr_template:
            customer_instr = cust_instr_template.format(quality=dia_quality)
        else:
            customer_instr = cust_instr_template

        # explode ORDER PCS into N rows with qty=1
        qty = row.get("ORDER PCS")
        if pd.isna(qty):
            continue
        try:
            qty_val = float(qty)
        except Exception:
            continue
        if qty_val <= 0:
            continue

        count = int(qty_val)
        for _ in range(count):
            out_row = {
                "SrNo": None,
                "StyleCode": style_code,
                "Unnamed: 2": None,
                "ItemSize": None,        # ring size ignored
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
                "Unnamed: 25": None,
                "SetPrice": None,
                "StoneQuality": None,    # ALWAYS BLANK
            }
            rows.append(out_row)

    out_df = pd.DataFrame(rows)

    if not out_df.empty:
        out_df["SrNo"] = range(1, len(out_df) + 1)

    for col in OUTPUT_COLUMNS:
        if col not in out_df.columns:
            out_df[col] = None

    return out_df[OUTPUT_COLUMNS]


# ---------- STREAMLIT UI ---------- #
st.title("PO Automation – Order Import Sheet Generator")

uploaded_file = st.file_uploader(
    "Upload RAW order sheet (like JUN-D-AJ4125.xlsx)",
    type=["xlsx", "xls"],
)

if uploaded_file is not None:
    raw_df = pd.read_excel(uploaded_file)
    clean_df = clean_input_df(raw_df)

    # safe preview (only important columns, unique col names)
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
    st.dataframe(preview[show_cols].head(30))

    if "DESIGN CODE" not in clean_df.columns:
        st.error(
            "This does not look like the RAW order sheet "
            "(no 'DESIGN CODE' column after cleaning).\n"
            "You may have uploaded the already formatted Order Import file."
        )
    else:
        default_order_group = uploaded_file.name.rsplit(".", 1)[0]
        order_group = st.text_input("Order Group / PO No", value=default_order_group)

        st.sidebar.header("Instruction Templates")
        cust_instr_template = st.sidebar.text_area(
            "Customer Production Instruction Template",
            value="IGI CERTI-CO-BRANDING( {quality}), HALLMARK-BIS, "
                  "REQ. AJ-STYLE NO.ON IGI CERTI",
            height=100,
        )
        remark_prefix = st.sidebar.text_area(
            "Special Remarks Base Prefix",
            value="NO 2 TONE RHODIUM ON METAL PART",
            height=80,
        )

        if st.button("Generate Order Import Sheet"):
            result_df = transform_to_order_import(
                clean_df, order_group, cust_instr_template, remark_prefix
            )

            if result_df.empty:
                st.error(
                    "No rows were generated.\n"
                    "Check that 'ORDER PCS' has values > 0."
                )
            else:
                st.subheader("Generated Order Import Sheet (Preview)")
                st.write(f"Total rows: {len(result_df)}")
                st.dataframe(result_df.head(50))

                buffer = BytesIO()
                with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                    # For Excel export, make the two unnamed columns truly blank headers
                    export_df = result_df.copy().rename(
                        columns={
                            "Unnamed: 2": "",
                            "Unnamed: 25": "",
                        }
                    )
                    export_df.to_excel(writer, index=False, sheet_name="Order Import")
                buffer.seek(0)

                st.download_button(
                    label="📥 Download Order Import Sheet (.xlsx)",
                    data=buffer.getvalue(),
                    file_name=f"{order_group}_Order_Import.xlsx",
                    mime=(
                        "application/"
                        "vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    ),
                )
else:
    st.info("Upload the RAW order Excel (not the already formatted Order Import file) to start.")
