
import streamlit as st
import pandas as pd
from io import BytesIO
import re

st.set_page_config(page_title="🔄 UPC Merge Tool (Description Mapping)", layout="wide")
st.title("🔄 UPC Merge Tool (with Manual Description Mapping)")

st.markdown("""
✅ Supports:
- Excel files with headers starting on row 3
- Multi-sheet Excel files
- Auto-maps:
  - `gtin`, `GTIN` → `barcode`
  - `brand` → `bh2Brand`
  - `product_type` → category breakdown
- Uses `title` as description **if available**, otherwise uses `description`, or lets you pick if neither exists
- Extracts product size and count info
- Merges new UPCs into your Partner Dashboard file
""")

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

# Upload files
upc_file = st.file_uploader("📤 Upload Cleaned UPC Excel File", type=["xlsx"])
partner_file = st.file_uploader("📤 Upload Partner Product File", type=["xlsx"])

if upc_file and partner_file:
    all_sheets = pd.read_excel(upc_file, sheet_name=None, header=2)
    upc_df = pd.concat(all_sheets.values(), ignore_index=True)
    partner_df = pd.read_excel(partner_file)

    upc_df.columns = [col.lower().strip() for col in upc_df.columns]
    columns = upc_df.columns

    # Auto-detect standard columns
    desc_col = 'title' if 'title' in columns else ('description' if 'description' in columns else None)
    upc_col = 'gtin' if 'gtin' in columns else ('barcode' if 'barcode' in columns else None)
    brand_col = 'brand' if 'brand' in columns else None
    product_type_col = 'product_type' if 'product_type' in columns else None

    # If no title or description, ask user to pick a column
    if not desc_col:
        desc_col = st.selectbox("⚠️ No `title` or `description` column found. Please select which column to use as the product description:", options=columns)

    st.markdown("#### Auto-Mapped Fields")
    st.write(f"📝 Description Source: `{desc_col}`")
    st.write(f"🔑 Barcode Source: `{upc_col}`")
    st.write(f"🏷️ Brand: `{brand_col}`")
    st.write(f"📦 Product Type → `{product_type_col}`")

    if not upc_col or not desc_col:
        st.error("❌ Cannot proceed. Must detect or select a UPC column and a Description column.")
    else:
        if st.button("🚀 Merge & Extract"):
            upc_df[upc_col] = (
                upc_df[upc_col]
                .astype(str)
                .str.replace(r'\.0$', '', regex=True)
                .str.extract(r'(\d+)', expand=False)
                .fillna('')
                .str.zfill(12)
            )
            partner_df['barcode'] = partner_df['barcode'].astype(str).str.extract(r'(\d+)', expand=False).fillna('').str.zfill(12)

            existing_barcodes = set(partner_df['barcode'])
            upc_df['STATUS'] = upc_df[upc_col].apply(lambda x: 'Existing' if x in existing_barcodes else 'New')
            new_upcs_df = upc_df[upc_df['STATUS'] == 'New'].copy()

            parsed_fields = new_upcs_df[desc_col].fillna('').apply(extract_size_components)
            new_upcs_df = pd.concat([new_upcs_df, parsed_fields], axis=1)

            # Handle category breakdown
            if product_type_col:
                cat_split = new_upcs_df[product_type_col].fillna('').str.split('>', expand=True)
                new_upcs_df['ch1Department'] = cat_split[0].str.strip().fillna("N/A") if 0 in cat_split else "N/A"
                new_upcs_df['ch2Category'] = cat_split[1].str.strip().fillna("N/A") if 1 in cat_split else "N/A"
                new_upcs_df['ch3Segment'] = cat_split[2].str.strip().fillna("N/A") if 2 in cat_split else "N/A"
            else:
                new_upcs_df['ch1Department'] = "N/A"
                new_upcs_df['ch2Category'] = "N/A"
                new_upcs_df['ch3Segment'] = "N/A"

            new_rows = pd.DataFrame({
                'barcode': new_upcs_df[upc_col],
                'bh2Brand': new_upcs_df[brand_col].str.upper() if brand_col else "N/A",
                'name': new_upcs_df[desc_col],
                'description': new_upcs_df[desc_col],
                'ch1Department': new_upcs_df['ch1Department'],
                'ch2Category': new_upcs_df['ch2Category'],
                'ch3Segment': new_upcs_df['ch3Segment'],
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
                file_name="merged_manual_description_upcs.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
