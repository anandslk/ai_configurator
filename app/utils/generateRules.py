import os
import shutil
import zipfile
import re
import pandas as pd
import numpy as np
from fastapi import FastAPI, Query
from fastapi.responses import FileResponse
from collections import Counter


# Configuration
OUTPUT_FOLDER = "generated_rules"
ZIP_NAME = "generated_rules.zip"
PREFIX = "KEY-GR"
VALVE_SIZE_GROUP = "valve_size"

def clean_identifier(s):
    """Clean and normalize identifiers for rules"""
    s = str(s).strip()
    s = re.sub(r'[^\w\s]', '', s)  # Remove special characters
    s = re.sub(r'\s+', '_', s)      # Replace spaces with underscores
    s = s.lower()                   # Convert to lowercase
    return s

def extract_size_value(value):
    """Extract size value from string"""
    # First try to find alphanumeric size identifiers
    match = re.search(r'([a-z]\d+)', str(value), re.IGNORECASE)
    if match:
        return match.group(1).lower()
    
    # Then try any numbers
    match = re.search(r'(\d{2,})', str(value))
    if match:
        return match.group(1)
    
    return str(value)

def find_valve_size_column(df):
    """Identify valve size column using heuristics"""
    # Look for columns with size identifiers
    for col in df.columns:
        # Check for P1, P2, etc.
        if any(re.search(r'p\d+', str(v), re.IGNORECASE) for v in df[col].head(5)):
            return col
    
    # Look for columns with size numbers
    for col in df.columns:
        if any(re.search(r'\b\d{2,}\b', str(v)) for v in df[col].head(5)):
            return col
    
    # Fallback to first column
    return df.columns[0]

def find_attribute_columns(df, valve_col):
    """Identify attribute columns based on content"""
    attribute_cols = []
    for col in df.columns:
        if col == valve_col:
            continue
        # Check for Y/N values
        if any(str(v).upper() in ['Y', 'N'] for v in df[col].head(10)):
            attribute_cols.append(col)
    return attribute_cols

def find_header_row(raw_df):
    """Dynamically locate the header row"""
    # Calculate row complexity scores
    row_scores = []
    for idx, row in raw_df.iterrows():
        # Skip empty rows
        if row.isna().all():
            continue
            
        # Count non-empty cells
        non_empty = row.notna().sum()
        
        # Count unique values
        unique_vals = row.dropna().astype(str).str.strip().nunique()
        
        # Calculate string complexity
        str_complexity = row.dropna().astype(str).str.len().sum() / (non_empty or 1)
        
        # Composite score
        score = (non_empty * 0.4) + (unique_vals * 0.4) + (str_complexity * 0.2)
        row_scores.append((idx, score))
    
    # Sort by score descending
    if row_scores:
        row_scores.sort(key=lambda x: x[1], reverse=True)
        return row_scores[0][0]
    
    return 0

def clean_dataframe(df):
    """Clean and normalize dataframe"""
    # Clean column names
    df.columns = [clean_identifier(col) for col in df.columns]
    
    # Remove completely empty rows and columns
    df = df.dropna(how='all').reset_index(drop=True)
    df = df.dropna(axis=1, how='all')
    
    # Convert all values to string and strip whitespace
    df = df.map(lambda x: str(x).strip() if pd.notna(x) else x)
    
    return df

def generate_rules_from_sheet(df, sheet_name):
    """Generate exclusion rules from a DataFrame"""
    # Find valve size column
    valve_col = find_valve_size_column(df)
    if not valve_col:
        print(f"⚠️ Valve size column not found in sheet: {sheet_name}")
        return []
    
    # Identify attribute columns
    attribute_cols = find_attribute_columns(df, valve_col)
    if not attribute_cols:
        print(f"⚠️ No attribute columns found in sheet: {sheet_name}")
        return []
    
    # Clean valve size values
    valve_sizes = df[valve_col].astype(str).str.strip()
    valve_sizes = valve_sizes[valve_sizes != ""]  # Remove empty values
    
    rules = []
    
    for col in attribute_cols:
        # Clean attribute name
        attr_name = clean_identifier(col)
        
        # Clean and standardize values
        col_values = df[col].astype(str).str.strip().str.upper()
        
        # Find exclusions - only process 'N' values
        exclusion_mask = col_values == "N"
        excluded_sizes = valve_sizes[exclusion_mask].unique()
        
        # Create rule if we have exclusions
        if len(excluded_sizes) > 0:
            # Process each excluded size
            size_entries = []
            for size_val in excluded_sizes:
                if size_val:
                    # Extract size value
                    clean_size = extract_size_value(size_val)
                    size_entries.append(
                        f"'{PREFIX}'.'{VALVE_SIZE_GROUP}'.'{clean_size}'"
                    )
            
            if size_entries:
                # Generate rule
                size_entries_str = ", ".join(size_entries)
                rule = (
                    f"AnyTrue({size_entries_str}) "
                    f"Excludes AnyTrue('{PREFIX}'.'{attr_name}'.'{attr_name}')"
                )
                rules.append(rule)
    
    return rules

def process_sheet(sheet_df, sheet_name):
    """Process a sheet with intelligent header detection"""
    # Try to find header row
    header_row_idx = find_header_row(sheet_df)
    
    # Create a copy to avoid SettingWithCopyWarning
    df = sheet_df.copy()
    
    # Check if we have a valid header row
    if header_row_idx > 0:
        # Set header and remove the header row
        df.columns = df.iloc[header_row_idx]
        df = df.drop(header_row_idx).reset_index(drop=True)
    
    # Clean the dataframe
    df = clean_dataframe(df)
    
    # Generate rules
    return generate_rules_from_sheet(df, sheet_name)

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

