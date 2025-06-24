import os, tempfile
import pandas as pd
from io import BytesIO
from pathlib import Path

# Import the provided backend modules
import givetochat1_6 as backend
from givetochat1_6 import Parser  # Parser and Person classes
import cabpool, team_trip_carpooling, go_live_carpooling

def get_hotel_list(uploaded_file) -> list:
    """Extract unique hotel names from the uploaded Excel file."""
    # Read the file into a DataFrame and parse via Parser to ensure consistency
    # We use BytesIO to avoid saving to disk here
    data_bytes = uploaded_file.read()
    excel_buffer = BytesIO(data_bytes)
    people, _, _ = Parser.process_excel(excel_buffer)
    hotels = Parser.list_hotels(people)  # returns a set of hotel names:contentReference[oaicite:12]{index=12}
    return list(hotels)

def process_carpool_assignment(uploaded_file, ride_type: str, hotel_window_min: int, airport_window_min: int, hotel_groups: list[list[str]], output_filename: str):
    """
    Process the uploaded Excel file according to the selected ride_type,
    time windows, and hotel merge groups. Writes output to an Excel file 
    (output_filename) and returns the file path and a preview dataframe (or None).
    """
    # Save uploaded file to a temporary path for use by backend functions
    tmpdir = tempfile.gettempdir()
    base = Path(uploaded_file.name).stem
    input_path = os.path.join(tmpdir, uploaded_file.name)
    output_name = f"output_{base}.xlsx"
    with open(input_path, 'wb') as f:
        f.write(uploaded_file.getbuffer())
    output_path = os.path.join(tmpdir, output_name)
    
    # Prepare to override Parser behavior (merge hotels and adjust clustering thresholds)
    # 1. Monkey-patch Parser.process_excel to merge hotels after parsing
    orig_process_excel = Parser.process_excel
    def patched_process_excel(input_file):
        people, h_personal, a_personal = orig_process_excel(input_file)
        if hotel_groups:
            # Merge each group of hotels specified by user
            for group in hotel_groups:
                if group:  # only if the group list is not empty
                    Parser.merge_hotels(people, [group])  # merge this list of hotel names:contentReference[oaicite:13]{index=13}
        # Recompute personal ride lists after merging (just in case, though merge doesn't affect flags)
        h_personal = [p for p in people if p.personal.get("Hotel", False)]
        a_personal = [p for p in people if p.personal.get("Airport", False)]
        return people, h_personal, a_personal
    Parser.process_excel = patched_process_excel
    
    # 2. Monkey-patch Parser._cluster_by_time to use custom thresholds
    orig_cluster_func = Parser._cluster_by_time
    def patched_cluster(travelers, attribute, threshold_hours=1.0):
        # Use our custom thresholds for arrival and return time clustering
        if attribute == "arrival_dt":
            # Slider in minutes -> hours
            return orig_cluster_func(travelers, attribute, threshold_hours=hotel_window_min / 60.0)
        elif attribute == "return_dt":
            return orig_cluster_func(travelers, attribute, threshold_hours=airport_window_min / 60.0)
        else:
            # For any other usage, fall back to the original
            return orig_cluster_func(travelers, attribute, threshold_hours=threshold_hours)
    Parser._cluster_by_time = patched_cluster
    
    try:
        # 3. Call the appropriate backend function based on ride_type
        if ride_type == "Go-Live":
            go_live_carpooling.generate_carpool_assignments(input_path, output_path)
        elif ride_type == "Team Trip":
            team_trip_carpooling.write_carpool_excel(input_path, output_path)
        elif ride_type == "Cabpool":
            cabpool.write_cab_excel(input_path, output_path)
        else:
            raise ValueError(f"Unknown ride type: {ride_type}")
    finally:
        # Restore original Parser functions to avoid side effects
        Parser.process_excel = orig_process_excel
        Parser._cluster_by_time = orig_cluster_func
    
    # 4. Load output Excel to a dataframe for preview (if small)
    df_preview = None
    try:
        df_out = pd.read_excel(output_path, sheet_name=0)  # read first sheet of output
        if len(df_out) <= 100:
            df_preview = df_out
    except Exception as e:
        # If output file read fails (unexpected format), just ignore preview
        df_preview = None
    
    return output_path, df_preview
