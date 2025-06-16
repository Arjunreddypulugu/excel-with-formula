import streamlit as st
import pandas as pd
from collections import defaultdict
from io import BytesIO
import difflib
import sqlalchemy

REQUIRED_COLUMNS = [
    'serial', 'total qty', 'spare qty', 'item no.', 'description', 'unit price ($)'
]

def find_best_column_matches(df_columns):
    normalized = {col.lower().strip(): col for col in df_columns if isinstance(col, str)}
    matches = {}
    for target in REQUIRED_COLUMNS:
        close_matches = difflib.get_close_matches(target, normalized.keys(), n=1, cutoff=0.6)
        if close_matches:
            matches[target] = normalized[close_matches[0]]
        else:
            raise ValueError(f"Could not find a match for required column: '{target}'")
    return matches

def get_ami_data():
    conn_str = st.secrets["mssql"]["connection_string"]
    engine = sqlalchemy.create_engine(conn_str)
    query = "SELECT SerialNumber, Model, EquipmentType FROM EquipmentDB"
    return pd.read_sql(query, con=engine)

def process_single_sheet(input_df, ami_df):
    col_map = find_best_column_matches(input_df.columns)
    input_df = input_df.rename(columns={v: k.title() for k, v in col_map.items()})
    input_df = input_df[[k.title() for k in REQUIRED_COLUMNS]]

    ami_df.dropna(subset=['SerialNumber'], inplace=True)

    serial_to_model = {}
    serial_to_type = {}

    for _, row in ami_df.iterrows():
        serial = row['SerialNumber']
        model = row['Model'] if pd.notna(row['Model']) else "MODEL MISSING"
        equip_type = row['EquipmentType'] if pd.notna(row['EquipmentType']) else "TYPE MISSING"
        serial_to_model[serial] = model
        serial_to_type[serial] = equip_type

    model_to_type = {}
    for serial, model in serial_to_model.items():
        if model not in model_to_type:
            model_to_type[model] = serial_to_type.get(serial, "TYPE MISSING")

    model_spares = defaultdict(list)
    last_serial = None
    current_model = None
    current_type = None

    for _, row in input_df.iterrows():
        serial = row['Serial']
        if serial != last_serial:
            last_serial = serial
            current_model = serial_to_model.get(serial, "MODEL MISSING")
            current_type = serial_to_type.get(serial, "TYPE MISSING")
            continue

        item_no = row['Item No.']
        description = row['Description']
        unit_price = row['Unit Price ($)']
        total_qty = pd.to_numeric(row['Total Qty'], errors='coerce') or 0
        spare_qty = pd.to_numeric(row['Spare Qty'], errors='coerce') or 0

        if pd.notna(item_no) and str(item_no).strip().upper() != 'TBD' and pd.notna(description):
            model_spares[(current_type, current_model)].append({
                'Item no.': item_no,
                'Description': description,
                'Unit Price ($)': unit_price,
                'Total qty': total_qty,
                'Spare qty': spare_qty
            })

    output_rows = []
    grouped_models = sorted(model_spares.keys(), key=lambda x: (x[0], x[1]))

    for equip_type, model in grouped_models:
        output_rows.append([equip_type, model, '', '', '', '', ''])
        parts = model_spares[(equip_type, model)]
        grouped_parts = defaultdict(lambda: {
            "Item no.": None,
            "Total qty": 0,
            "Spare qty": 0,
            "Unit Price ($)": None,
            "Description": ""
        })

        for part in parts:
            item_no = part['Item no.']
            grouped_parts[item_no]["Description"] = part['Description']
            grouped_parts[item_no]["Unit Price ($)"] = part['Unit Price ($)']
            grouped_parts[item_no]["Total qty"] += part['Total qty']
            grouped_parts[item_no]["Spare qty"] += part['Spare qty']

        for item_no in sorted(grouped_parts.keys(), key=lambda x: grouped_parts[x]["Description"]):
            data = grouped_parts[item_no]
            output_rows.append([
                '', '', data['Total qty'], data['Spare qty'],
                item_no, data['Description'], data['Unit Price ($)']
            ])

    return pd.DataFrame(output_rows, columns=[
        'Equipment Type', 'Model', 'Total qty', 'Spare qty', 'Item no.', 'Description', 'Unit Price ($)'
    ])

def process_excel(uploaded_file):
    input_excel = pd.ExcelFile(uploaded_file)
    ami_df = get_ami_data()

    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name in input_excel.sheet_names:
            input_df = input_excel.parse(sheet_name)
            try:
                processed_df = process_single_sheet(input_df, ami_df)
                processed_df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
            except Exception as e:
                error_df = pd.DataFrame({"Error": [f"Could not process sheet '{sheet_name}': {str(e)}"]})
                error_df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
    output.seek(0)
    return output

# ===== Streamlit Interface =====
st.title("🔧 Excel Re-organizer")
st.markdown("Connected to SQL Server")

uploaded_file = st.file_uploader("Upload the input Excel file", type=["xlsx"])

if uploaded_file:
    with st.spinner("Processing all sheets with SQL data..."):
        output_excel = process_excel(uploaded_file)

    st.success("✅ File processed successfully!")
    st.download_button(
        label="📥 Download Output Excel",
        data=output_excel,
        file_name="formatted_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
