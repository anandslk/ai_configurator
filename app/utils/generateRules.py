import re
import pandas as pd
from fastapi import FastAPI

# Configuration
OUTPUT_FOLDER = "generated_rules"
ZIP_NAME = "generated_rules.zip"
PREFIX = "KEY-GR"

def clean_identifier(s):
    """Clean and normalize identifiers for rules"""
    s = str(s).strip()
    s = re.sub(r'[^\w\s]', '', s)  # Remove special characters
    s = re.sub(r'\s+', '_', s)      # Replace spaces with underscores
    return s

def find_key_column(df):
    """Dynamically identify the key column (valve identifiers)"""
    # Calculate uniqueness scores
    uniqueness = {col: df[col].nunique() for col in df.columns}
    
    # Calculate value length scores
    length_scores = {}
    for col in df.columns:
        avg_len = df[col].astype(str).str.len().mean()
        length_scores[col] = avg_len
    
    # Find best candidate
    best_col = None
    best_score = -1
    
    for col in df.columns:
        # Skip columns with mostly empty values
        empty_ratio = df[col].isna().mean()
        if empty_ratio > 0.5:
            continue
            
        # Calculate composite score
        uniqueness_score = uniqueness[col] / len(df)
        length_score = 1 - (length_scores[col] / (length_scores[col] + 10))
        score = uniqueness_score * 0.7 + length_score * 0.3
        
        if score > best_score:
            best_score = score
            best_col = col
    
    return best_col

def find_header_row(raw_df):
    """Dynamically locate the header row"""
    # Calculate row complexity scores
    row_scores = []
    for idx, row in raw_df.iterrows():
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
    row_scores.sort(key=lambda x: x[1], reverse=True)
    
    # Return best candidate
    return row_scores[0][0] if row_scores else 0

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
    key_col = find_key_column(df)
    if not key_col:
        print(f"⚠️ Key column not found in sheet: {sheet_name}")
        return []
    
    # Clean key values
    key_values = df[key_col].astype(str).str.strip()
    key_values = key_values[key_values != ""]  # Remove empty values
    
    rules = []
    
    for col in df.columns:
        if col == key_col or df[col].isna().all():
            continue  # Skip key column and empty columns
        
        # Clean and standardize values
        col_values = df[col].astype(str).str.strip().str.upper()
        
        # Find exclusions - only process 'N' values
        exclusion_mask = col_values == "N"
        excluded_keys = key_values[exclusion_mask].unique()
        
        # Create rule if we have exclusions
        if len(excluded_keys) > 0:
            # Clean identifiers
            safe_keys = [clean_identifier(v) for v in excluded_keys if v]
            safe_col = clean_identifier(col)
            
            # Generate rule with dynamic group names
            key_entries = ", ".join(
                [f"'{PREFIX}'.'{clean_identifier(key_col)}'.'{v}'" for v in safe_keys]
            )
            rule = (
                f"AnyTrue({key_entries}) "
                f"Excludes AnyTrue('{PREFIX}'.'{clean_identifier(col)}'.'{safe_col}')"
            )
            rules.append(rule)
    
    return rules

