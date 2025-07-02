import streamlit as st
import pandas as pd
from collections import defaultdict
from io import BytesIO
import difflib
import sqlalchemy
import math

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

    # Aggregate by Item no. only
    part_spares = defaultdict(lambda: {
        "Item no.": None,
        "Total qty": 0,
        "Unit Price ($)": None,
        "Description": "",
        "Models": set(),
        "Serials": set(),
        "Max Spare Per Machine": 0,
        "Equipment Types": set()
    })
    per_machine_spares = defaultdict(lambda: defaultdict(float))  # [item_no][serial] = spare_qty

    last_serial = None
    skip_next = False
    for _, row in input_df.iterrows():
        serial = row['Serial']
        if serial != last_serial:
            last_serial = serial
            skip_next = True
        if skip_next:
            skip_next = False
            continue

        model = serial_to_model.get(serial, "MODEL MISSING")
        equip_type = serial_to_type.get(serial, "TYPE MISSING")
        item_no = row['Item No.']
        description = row['Description']
        unit_price = row['Unit Price ($)']
        total_qty = pd.to_numeric(row['Total Qty'], errors='coerce') or 0
        spare_qty = pd.to_numeric(row['Spare Qty'], errors='coerce') or 0

        if pd.isna(item_no) or pd.isna(description):
            continue
        if str(item_no).strip().upper() == 'TBD':
            continue

        part_data = part_spares[item_no]
        part_data["Item no."] = item_no
        part_data["Description"] = description
        part_data["Unit Price ($)"] = unit_price
        part_data["Total qty"] += total_qty
        part_data["Models"].add(model)
        part_data["Serials"].add(serial)
        part_data["Equipment Types"].add(equip_type)
        per_machine_spares[item_no][serial] += spare_qty

    for item_no, part_data in part_spares.items():
        max_spare = 0
        for serial, qty in per_machine_spares[item_no].items():
            max_spare = max(max_spare, qty)
        part_data["Max Spare Per Machine"] = max_spare

    output_rows = []
    for item_no, data in sorted(part_spares.items(), key=lambda x: x[1]["Description"]):
        machine_count = len(data["Serials"])
        if machine_count < 5:
            scale_factor = 1.0
        elif machine_count < 10:
            scale_factor = 1.25
        elif machine_count < 15:
            scale_factor = 1.5
        elif machine_count < 20:
            scale_factor = 1.75
        elif machine_count < 25:
            scale_factor = 2.0
        else:
            scale_factor = 2.0
        final_spare_qty = math.ceil(data["Max Spare Per Machine"] * scale_factor)
        output_rows.append([
            data['Total qty'],
            final_spare_qty,
            item_no,
            data['Description'],
            data['Unit Price ($)'],
            ', '.join(sorted(data['Models'])),
            ', '.join(sorted(data['Equipment Types']))
        ])
    return pd.DataFrame(output_rows, columns=[
        'Total qty', 'Spare qty', 'Item no.', 'Description', 'Unit Price ($)', 'Models', 'Equipment Types'
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
st.title("🔧 Spare Parts Packager")
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
