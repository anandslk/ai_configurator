import pandas as pd
import os

# === Config ===
input_folder = r"C:\Users\anand.kumar\Documents\ruleset/files"
output_folder = "generated_rules"
os.makedirs(output_folder, exist_ok=True)

prefix = "KEY-GR"
group_name = "ball_disc_gate_material"

# === Process each Excel file in the input folder ===
for excel_file in os.listdir(input_folder):
    if not excel_file.lower().endswith('.xlsx'):
        continue  # Skip non-Excel files

    excel_path = os.path.join(input_folder, excel_file)
    input_filename = os.path.splitext(os.path.basename(excel_file))[0]
    
    print(f"Processing file: {excel_file}")

    xls = pd.ExcelFile(excel_path)
    sheet_names = xls.sheet_names
    all_rules = []

    for sheet_name in sheet_names:
        # Read full sheet without headers
        raw_df = pd.read_excel(xls, sheet_name=sheet_name, header=None)

        # Try to detect header row
        header_row_idx = None
        for idx, row in raw_df.iterrows():
            non_empty_cells = row.fillna('').astype(str).str.strip()
            if (non_empty_cells != '').sum() >= 2:
                header_row_idx = idx
                break

        if header_row_idx is None:
            print(f"⚠️ Warning: Could not find header in {sheet_name} of file {excel_file}. Skipping this sheet.")
            continue

        # Read data with header
        df = pd.read_excel(xls, sheet_name=sheet_name, header=header_row_idx)
        df.columns = df.columns.astype(str).str.strip()

        # Identify valve/size column
        valve_col = None
        for col in df.columns:
            if "size" in col.lower() or "valve" in col.lower():
                valve_col = col
                break
        if not valve_col:
            print(f"⚠️ Warning: 'Size' column not found in sheet {sheet_name} of {excel_file}. Skipping.")
            continue

        valve_sizes = df[valve_col].astype(str).str.strip()

        for col in df.columns:
            if col == valve_col:
                continue

            col_values = df[col].astype(str).str.upper()
            excluded_valves = valve_sizes[col_values == 'N']

            if not excluded_valves.empty:
                valve_entries = ", ".join([
                    f"'{prefix}'.'valve_size'.'{v.zfill(4)}'" for v in excluded_valves
                ])
                rule = f"AnyTrue({valve_entries}) Excludes AnyTrue('{prefix}'.'{group_name}'.'{col}')"
                all_rules.append(rule)

    # Output rules to a text file per Excel file
    if all_rules:
        output_file = os.path.join(output_folder, f"{input_filename}_rules.txt")
        with open(output_file, "w") as f:
            for rule in all_rules:
                f.write(rule + ";\n")
        print(f"✅ Rule text generated for {excel_file} -> {output_file}")
    else:
        print(f"❌ No rules generated for {excel_file}.")

print("🎉 Completed rule generation for all Valve Size Excel files.") 