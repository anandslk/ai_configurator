import os
import uuid
import zipfile
from fastapi import FastAPI, Query
from fastapi.responses import FileResponse
import pandas as pd
import shutil
from .extractTables import extract_tables_with_formatting

app = FastAPI()

@app.get("/extract")
def extract_from_path(
    excel_path: str = Query(..., description="Full path to Excel file"),
    start_sheet: str = Query(...),
    end_sheet: str = Query(...)
):
    if not os.path.exists(excel_path):
        return {"error": f"❌ File not found: {excel_path}"}

    session_id = uuid.uuid4().hex
    output_dir = f"output_{session_id}"
    os.makedirs(output_dir, exist_ok=True)

    extract_tables_with_formatting(excel_path, output_dir, start_sheet, end_sheet)

    zip_path = f"tables_{session_id}.zip"
    with zipfile.ZipFile(zip_path, "w") as zipf:
        for file in os.listdir(output_dir):
            zipf.write(os.path.join(output_dir, file), arcname=file)

    # Clean up extracted xlsx files (optional)
    for file in os.listdir(output_dir):
        os.remove(os.path.join(output_dir, file))
    os.rmdir(output_dir)

    return FileResponse(zip_path, filename="extracted_tables.zip", media_type="application/zip")

OUTPUT_FOLDER = "generated_rules"
ZIP_NAME = "generated_rules.zip"
PREFIX = "KEY-GR"
GROUP_NAME = "ball_disc_gate_material"

# os.makedirs(INPUT_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.get("/generate-rules")
def generate_rules(excel_files_path: str = Query(..., description="Full path to Excel file"),):
    # Clear previous outputs
    if os.path.exists(OUTPUT_FOLDER):
        shutil.rmtree(OUTPUT_FOLDER)
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    rule_files_created = []

    for excel_file in os.listdir(excel_files_path):
        if not excel_file.lower().endswith(".xlsx"):
            continue

        excel_path = os.path.join(excel_files_path, excel_file)
        input_filename = os.path.splitext(os.path.basename(excel_file))[0]
        print(f"Processing file: {excel_file}")

        try:
            xls = pd.ExcelFile(excel_path)
            sheet_names = xls.sheet_names
        except Exception as e:
            print(f"❌ Could not open {excel_file}: {e}")
            continue

        all_rules = []

        for sheet_name in sheet_names:
            try:
                raw_df = pd.read_excel(xls, sheet_name=sheet_name, header=None)

                # Detect header row
                header_row_idx = None
                for idx, row in raw_df.iterrows():
                    non_empty_cells = row.fillna("").astype(str).str.strip()
                    if (non_empty_cells != "").sum() >= 2:
                        header_row_idx = idx
                        break

                if header_row_idx is None:
                    print(f"⚠️ Header not found in {sheet_name} of {excel_file}")
                    continue

                df = pd.read_excel(xls, sheet_name=sheet_name, header=header_row_idx)
                df.columns = df.columns.astype(str).str.strip()

                # Match column like: "Valve Size", "Drilling / Schedule", etc.
                valve_col = next((
                    col for col in df.columns
                    if any(key in col.lower() for key in ["size", "valve", "drilling", "schedule"])
                ), None)

                if not valve_col:
                    print(f"⚠️ 'Valve Size' or similar column not found in {sheet_name} of {excel_file}")
                    continue

                valve_sizes = df[valve_col].astype(str).str.strip()

                for col in df.columns:
                    if col == valve_col:
                        continue

                    col_values = df[col].astype(str).str.upper()
                    excluded_valves = valve_sizes[col_values == "N"]

                    if not excluded_valves.empty:
                        valve_entries = ", ".join([
                            f"'{PREFIX}'.'valve_size'.'{v.zfill(4)}'" for v in excluded_valves
                        ])
                        rule = f"AnyTrue({valve_entries}) Excludes AnyTrue('{PREFIX}'.'{GROUP_NAME}'.'{col}')"
                        all_rules.append(rule)

            except Exception as e:
                print(f"⚠️ Error processing {sheet_name} in {excel_file}: {e}")
                continue

        # Write rules
        if all_rules:
            output_file = os.path.join(OUTPUT_FOLDER, f"{input_filename}_rules.txt")
            with open(output_file, "w") as f:
                for rule in all_rules:
                    f.write(rule + ";\n")
            rule_files_created.append(output_file)
            print(f"✅ Rules written to: {output_file}")
        else:
            print(f"❌ No rules generated for {excel_file}")

    # Create ZIP
    if not rule_files_created:
        return {"message": "No rules generated from any file."}

    with zipfile.ZipFile(ZIP_NAME, "w") as zipf:
        for file_path in rule_files_created:
            zipf.write(file_path, os.path.basename(file_path))

    return FileResponse(ZIP_NAME, media_type="application/zip", filename=ZIP_NAME)