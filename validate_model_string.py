import pandas as pd
import re
import sys
import io
import json

# Ensure UTF-8 output in Windows terminal
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def main():
    if len(sys.argv) < 2:
        print(json.dumps({
            "valid": False,
            "message": "No model string provided."
        }))
        return

    model_string = sys.argv[1]
    file_path = r"C:\Users\veerabhadra.ronad\Downloads\Input_folder\KEY-GR_PM.xlsx"
    sheet_name = "Product Matrix"

    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)

        # Regex pattern to extract model string components
        pattern = r'GR-(\d{4})(\w{2})(\w{2})(\d{2})(\d{2})-(\w{2})(\w{3})(\w{2})(\w{2})(\d{2})(\w)(\w)-?(\w{3})?'
        match = re.match(pattern, model_string)

        if not match:
            print(json.dumps({
                "valid": False,
                "message": "Model string format is incorrect."
            }))
            return

        (
            size, end_connection, drilling_schedule, face_to_face, pressure_rating,
            body, disc, stem, seat, gasket, operator, actuation, special
        ) = match.groups()

        # Map components to expected column names
        code_map = {
            "size": size,
            "end_connection": end_connection,
            "drilling/schedule": drilling_schedule,
            "face_to_face": face_to_face,
            "pressure_rating": pressure_rating,
            "body_material": body,
            "disc_material": disc,
            "stem_material": stem,
            "seat_material": seat,
            "sealing": gasket,
            "operator_mounting_type": operator,
            "actuation_type": actuation,
        }

        if special:
            code_map["special_features"] = special

        # Validate each component
        for col_key, code_value in code_map.items():
            matching_columns = [col for col in df.columns if col_key.lower() in col.lower()]
            if not matching_columns:
                print(json.dumps({
                    "valid": False,
                    "message": f"Missing column in Excel: {col_key}"
                }))
                return

            valid_values = df[matching_columns[0]].astype(str).str.strip().unique()
            if code_value.strip() not in valid_values:
                print(json.dumps({
                    "valid": False,
                    "message": f"Invalid code '{code_value}' for '{col_key}'."
                }))
                return

        # If all checks pass
        print(json.dumps({
            "valid": True,
            "message": f"Model string '{model_string}' is valid."
        }))

    except Exception as e:
        print(json.dumps({
            "valid": False,
            "message": f"Error reading Excel: {str(e)}"
        }))

if __name__ == "__main__":
    main()
