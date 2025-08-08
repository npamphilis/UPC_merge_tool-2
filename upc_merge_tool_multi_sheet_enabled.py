
import streamlit as st
import pandas as pd
from io import BytesIO
import re

st.set_page_config(page_title="🔄 UPC Merge Tool (All Sheets)", layout="wide")
st.title("🔄 UPC Merge Tool (Reads All Sheets + Size & Count Parsing)")

st.markdown("""
This tool:
- Reads **all sheets** from a cleaned UPC Excel file
- Auto-maps columns like UPC, description, brand, and category
- Fixes UPC formatting issues (e.g., scientific notation, trailing decimals)
- Extracts product size and count fields from description:
  - `itemCountValue`, `itemCountMeasure`
  - `sizeValue`, `sizeMeasure`
- Merges with your Partner Dashboard file
""")

# Aliases for auto-mapping
UPC_ALIASES = ['barcode', 'upc']
DESC_ALIASES = ['description', 'name', 'product / fido id', 'product name', 'product description']
BRAND_ALIASES = ['brand']
CAT1_ALIASES = ['department', 'category 1', 'category_1']
CAT2_ALIASES = ['category', 'category 2', 'category_2']
CAT3_ALIASES = ['segment', 'category 3', 'category_3']

def detect_column(columns, aliases):
    normalized_cols = {col.lower().strip(): col for col in columns}
    for alias in aliases:
        if alias in normalized_cols:
            return normalized_cols[alias]
    return None

def extract_size_components(desc):
    desc = desc.lower()
    size_match = re.search(r'(\d+(\.\d+)?)\s?(oz|fl oz|l|ml|gallon|gal)', desc)
    count_match = re.search(r'(\d+)\s?ct', desc)

    size_value = size_match.group(1) if size_match else None
    size_measure = size_match.group(3).upper() if size_match else None
    count_value = count_match.group(1) if count_match else None
    count_measure = 'CT' if count_match else None

    if size_measure:
        if size_measure in ['FL OZ', 'OZ']:
            size_measure = 'OZ'
        elif size_measure in ['GAL', 'GALLON']:
            size_measure = 'GALLON'
        elif size_measure == 'L':
            size_measure = 'L'
        elif size_measure == 'ML':
            size_measure = 'ML'

    return pd.Series({
        'sizeValue': size_value,
        'sizeMeasure': size_measure,
        'itemCountValue': count_value,
        'itemCountMeasure': count_measure
    })

# Upload cleaned UPC file (multi-sheet support)
upc_file = st.file_uploader("📤 Upload Cleaned UPC Excel File (multi-sheet)", type=["xlsx"])
partner_file = st.file_uploader("📤 Upload Partner Product File", type=["xlsx"])

if upc_file and partner_file:
    # Load all sheets and concatenate
    all_sheets = pd.read_excel(upc_file, sheet_name=None)
    upc_df = pd.concat(all_sheets.values(), ignore_index=True)

    partner_df = pd.read_excel(partner_file)
    columns = upc_df.columns.tolist()

    # Auto-map
    upc_col = detect_column(columns, UPC_ALIASES)
    desc_col = detect_column(columns, DESC_ALIASES)
    brand_col = detect_column(columns, BRAND_ALIASES)
    dept_col = detect_column(columns, CAT1_ALIASES)
    cat2_col = detect_column(columns, CAT2_ALIASES)
    cat3_col = detect_column(columns, CAT3_ALIASES)

    st.markdown("#### Auto-Mapping Summary")
    st.write(f"🔑 UPC: `{upc_col}`")
    st.write(f"📝 Description: `{desc_col}`")
    st.write(f"🏷️ Brand: `{brand_col}`")
    st.write(f"📦 Dept: `{dept_col}`, 📁 Cat2: `{cat2_col}`, 📂 Seg: `{cat3_col}`")

    if not upc_col or not desc_col:
        st.error("❌ Cannot continue. Must detect both UPC and Description columns.")
    else:
        if st.button("🚀 Merge & Extract"):
            # Clean UPCs
            upc_df[upc_col] = (
                upc_df[upc_col]
                .astype(str)
                .str.replace(r'\.0$', '', regex=True)
                .str.extract(r'(\d+)', expand=False)
                .fillna('')
                .str.zfill(12)
            )
            partner_df['barcode'] = partner_df['barcode'].astype(str).str.extract(r'(\d+)', expand=False).fillna('').str.zfill(12)

            # Identify new UPCs
            existing_barcodes = set(partner_df['barcode'])
            upc_df['STATUS'] = upc_df[upc_col].apply(lambda x: 'Existing' if x in existing_barcodes else 'New')
            new_upcs_df = upc_df[upc_df['STATUS'] == 'New'].copy()

            # Extract sizes and counts
            parsed_fields = new_upcs_df[desc_col].fillna('').apply(extract_size_components)
            new_upcs_df = pd.concat([new_upcs_df, parsed_fields], axis=1)

            # Build partner-ready rows
            new_rows = pd.DataFrame({
                'barcode': new_upcs_df[upc_col],
                'bh2Brand': new_upcs_df[brand_col].str.upper() if brand_col else "N/A",
                'name': new_upcs_df[desc_col],
                'description': new_upcs_df[desc_col],
                'ch1Department': new_upcs_df[dept_col].str.upper() if dept_col else "N/A",
                'ch2Category': new_upcs_df[cat2_col].str.upper() if cat2_col else "N/A",
                'ch3Segment': new_upcs_df[cat3_col].str.upper() if cat3_col else "N/A",
                'itemCountValue': new_upcs_df['itemCountValue'],
                'itemCountMeasure': new_upcs_df['itemCountMeasure'],
                'sizeValue': new_upcs_df['sizeValue'],
                'sizeMeasure': new_upcs_df['sizeMeasure'],
                'partnerProduct': 'Y',
                'awardPoints': 'N'
            })

            for col in partner_df.columns:
                if col not in new_rows.columns:
                    new_rows[col] = None
            new_rows = new_rows[partner_df.columns]

            merged_df = pd.concat([partner_df, new_rows], ignore_index=True)

            output = BytesIO()
            merged_df.to_excel(output, index=False, engine='openpyxl')
            output.seek(0)

            st.download_button(
                label="📥 Download Final Merged File",
                data=output,
                file_name="merged_all_sheets_upcs.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
