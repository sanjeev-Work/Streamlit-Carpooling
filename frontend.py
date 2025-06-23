import streamlit as st
import pandas as pd
from pathlib import Path
from main import process_carpool_assignment, get_hotel_list

st.set_page_config(page_title="Carpool Assignment Tool", layout="centered")

st.title("Carpool Assignment Tool")
st.write("Upload a travel spreadsheet and configure carpool assignment options. "
         "Select the ride type, adjust time grouping windows, and optionally merge hotel groups. "
         "Then click **Run** to generate the carpool assignments.")

# 1. File upload
uploaded_file = st.file_uploader("📄 Upload Excel Spreadsheet", type=["xlsx", "xlsm"])
if uploaded_file:
    # Parse the file (lightly) to get unique hotel names for UI. We use Parser.list_hotels via main module.
    try:
        hotel_names = get_hotel_list(uploaded_file)
    except Exception as e:
        st.error(f"Error reading file: {e}")
        st.stop()
    
    # 2. UI elements for configuration
    ride_type = st.selectbox("Ride type", ["cabpool", "team trip", "go-live"], index=0)
    
    # Two sliders for grouping windows (in minutes)
    col1, col2 = st.columns(2)
    with col1:
        hotel_window = st.slider("Ride to Hotel grouping window (minutes)", min_value=0, max_value=180, value=30, step=5)
    with col2:
        airport_window = st.slider("Ride to Airport grouping window (minutes)", min_value=0, max_value=180, value=60, step=5)

    
    # Number of hotel merge groups
    num_groups = st.number_input("Number of hotel merge groups", min_value=0, max_value=10, value=0, step=1,
                                 help="How many groups of hotels to treat as one. For example, enter 1 to merge some hotels together.")
    hotel_groups = []  # list to collect selected hotels for each group
    if num_groups:
        st.info("Select hotels to merge in each group (each hotel should belong to at most one group).")
    for i in range(int(num_groups)):
        group_selection = st.multiselect(f"Merge Group {i+1}", options=sorted(hotel_names),
                                         help="Choose hotels to treat as a single group for carpooling")
        hotel_groups.append(group_selection)
    
    # 3. Run button to trigger processing
    run_clicked = st.button("🚀 Run")
    if run_clicked:
        # Write the uploaded file to a temporary location for processing
        base = Path(uploaded_file.name).stem
        output_filename = f'output_{base}.xlsx'
        try:
            result_path, df_preview = process_carpool_assignment(uploaded_file, ride_type, hotel_window, airport_window, hotel_groups, output_filename)
        except Exception as e:
            st.error(f"Failed to generate assignments: {e}")
        else:
            st.success(f"Carpool assignments generated! Download the file below.")
            # 4. Download link for output file
            with open(result_path, "rb") as f:
                out_bytes = f.read()
            st.download_button("💾 Download results", data=out_bytes, file_name=output_filename, 
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            # 5. If result is not too large, display preview table
            if df_preview is not None:
                st.subheader("Assignment Preview")
                st.dataframe(df_preview)
