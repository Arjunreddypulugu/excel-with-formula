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

    # Group parts by equipment type only
    type_spares = defaultdict(lambda: defaultdict(lambda: {
        "Item no.": None,
        "Total qty": 0,
        "Spare qty": 0,
        "Unit Price ($)": None,
        "Description": "",
        "Models": set(),  # Track which models use this part
        "Machine Count": 0,  # Track number of machines using this part
        "Max Spare Per Machine": 0  # Track highest spare quantity per machine
    }))

    # First pass: collect data and find max spare per machine
    current_serial = None
    current_model = None
    current_type = None
    machine_spares = defaultdict(int)  # Track spare quantities per machine for current part

    for _, row in input_df.iterrows():
        serial = row['Serial']
        if serial != current_serial:
            # Reset machine_spares when moving to a new machine
            machine_spares.clear()
            current_serial = serial
            current_model = serial_to_model.get(serial, "MODEL MISSING")
            current_type = serial_to_type.get(serial, "TYPE MISSING")
            continue

        item_no = row['Item No.']
        description = row['Description']
        unit_price = row['Unit Price ($)']
        total_qty = pd.to_numeric(row['Total Qty'], errors='coerce') or 0
        spare_qty = pd.to_numeric(row['Spare Qty'], errors='coerce') or 0

        if pd.notna(item_no) and str(item_no).strip().upper() != 'TBD' and pd.notna(description):
            part_data = type_spares[current_type][item_no]
            part_data["Item no."] = item_no
            part_data["Description"] = description
            part_data["Unit Price ($)"] = unit_price
            part_data["Total qty"] += total_qty
            part_data["Models"].add(current_model)
            part_data["Machine Count"] += 1
            
            # Track spare quantity for current machine
            machine_spares[item_no] += spare_qty
            # Update max spare per machine if current machine has more
            part_data["Max Spare Per Machine"] = max(part_data["Max Spare Per Machine"], machine_spares[item_no])

    output_rows = []
    # Sort equipment types
    for equip_type in sorted(type_spares.keys()):
        output_rows.append([equip_type, '', '', '', '', '', ''])
        # Get parts for this equipment type and sort by description
        parts = type_spares[equip_type]
        sorted_parts = sorted(parts.items(), key=lambda x: x[1]["Description"])
        
        for item_no, data in sorted_parts:
            # Calculate scaling factor based on number of machines
            machine_count = data["Machine Count"]
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
                scale_factor = 2.0  # Default to 2.0 for 25 or more machines

            # Calculate final spare quantity based on max spare per machine
            final_spare_qty = data["Max Spare Per Machine"] * scale_factor

            output_rows.append([
                '',  # Empty equipment type for sub-rows
                data['Total qty'],
                final_spare_qty,
                item_no,
                data['Description'],
                data['Unit Price ($)'],
                ', '.join(sorted(data['Models']))  # Join models with commas
            ])

    return pd.DataFrame(output_rows, columns=[
        'Equipment Type', 'Total qty', 'Spare qty', 'Item no.', 'Description', 'Unit Price ($)', 'Models'
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
