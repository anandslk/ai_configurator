import os
import uuid
import zipfile
from fastapi import FastAPI, Query
from fastapi.responses import FileResponse
import pandas as pd
import shutil
from .utils.extractTables import extract_tables_with_formatting
from .utils.generateRules import find_header_row, generate_rules_from_sheet, clean_dataframe, process_sheet

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
    dataest = os.path.exists(zip_path)

    return FileResponse(zip_path, filename="extracted_tables.zip", media_type="application/zip")

# Configuration
OUTPUT_FOLDER = "generated_rules"
ZIP_NAME = "generated_rules.zip"
PREFIX = "KEY-GR"
VALVE_SIZE_GROUP = "valve_size"

# os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.get("/generate-rules")
def generate_rules(excel_files_path: str = Query(..., description="Full path to directory containing Excel files")):
    # Setup output directory
    if os.path.exists(OUTPUT_FOLDER):
        shutil.rmtree(OUTPUT_FOLDER)
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    
    rule_files_created = []
    processed_count = 0
    
    # Process each Excel file
    for excel_file in os.listdir(excel_files_path):
        if not excel_file.lower().endswith((".xlsx", ".xls")):
            continue
        
        excel_path = os.path.join(excel_files_path, excel_file)
        input_filename = os.path.splitext(excel_file)[0]
        print(f"\nProcessing file: {excel_file}")
        
        try:
            xls = pd.ExcelFile(excel_path)
            sheet_names = xls.sheet_names
        except Exception as e:
            print(f"❌ Could not open {excel_file}: {e}")
            continue
        
        all_rules = []
        
        # Process each sheet
        for sheet_name in sheet_names:
            print(f"  - Sheet: {sheet_name}")
            try:
                # Read raw data without headers
                raw_df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
                
                # Skip empty sheets
                if raw_df.empty:
                    print("    ⚠️ Empty sheet, skipping")
                    continue
                
                # Process sheet
                sheet_rules = process_sheet(raw_df, sheet_name)
                all_rules.extend(sheet_rules)
                print(f"    ✅ Generated {len(sheet_rules)} rules")
                
            except Exception as e:
                print(f"    ⚠️ Error processing sheet: {str(e)}")
                continue
        
        # Save rules for this file
        if all_rules:
            output_file = os.path.join(OUTPUT_FOLDER, f"{input_filename}_rules.txt")
            with open(output_file, "w", encoding="utf-8") as f:
                f.write("\n".join([rule + ";" for rule in all_rules]))
            rule_files_created.append(output_file)
            processed_count += 1
            print(f"💾 Saved {len(all_rules)} rules to {output_file}")
        else:
            print(f"❌ No rules generated for {excel_file}")
    
    # Create ZIP archive
    if not rule_files_created:
        return {"message": "No rules generated from any file."}
    
    with zipfile.ZipFile(ZIP_NAME, "w") as zipf:
        for file_path in rule_files_created:
            zipf.write(file_path, os.path.basename(file_path))
    
    return FileResponse(
        ZIP_NAME,
        media_type="application/zip",
        filename=ZIP_NAME,
        headers={"X-Files-Processed": str(processed_count)}
    )