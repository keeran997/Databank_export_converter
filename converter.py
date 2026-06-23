import pandas as pd
import json
from openpyxl import load_workbook

def load_files(input_path, mapping_path, template_path):
    df = pd.read_csv(input_path)

    with open(mapping_path, "r", encoding="utf-8") as file:
        mapping = json.load(file)

    wb = load_workbook(template_path)

    return df, mapping, wb

def get_assessment_type(row, mapping):
    as_col = mapping["PreOp or PostOp"]
    status = str(row[as_col]).strip().lower()

    if status == "preop":
        return "preop"
    elif status == "postop":
        return "postop"
    else:
        raise ValueError(f"Unknown assessment type: {status}")

def get_target_cell(field, assessment_type):
    return field["targets"][assessment_type]

def populate_sheet(row, mapping, sheet):
    assessment_type = get_assessment_type(row, mapping)

    for field in mapping["fields"]:
        source_col = field["source"]
        value = row.get(source_col, None)
        target_cell = get_target_cell(field, assessment_type)
        sheet[target_cell] = value

def validate_columns(df, mapping):
    required_col = {mapping["PreOp or PostOp"]}

    for field in mapping["fields"]:
        required_col.add(field["source"])

    missing = [c for c in required_col if c not in df.columns]

    return missing

def convert(input_path, mapping_path, template_path, output_path, progress_callback=None, cancel_flag=None):
    """
        Convert one CSV file

        Returns True if workbook was converted and saved successfully, false if cancelled.
    """
    df, mapping, wb = load_files(input_path, mapping_path, template_path)

    if df.empty:
        raise ValueError("The CSV file is empty.")

    if len(df) > 1:
        raise ValueError(
            "The selected CSV contains more than one record.\n\n"
            "Please export and convert one record at a time."
        )

    missing = validate_columns(df, mapping)

    if missing:
        formatted_col = "\n".join(
            f"• {column}" for column in missing
        )

        raise ValueError(
            "The CSV file is missing the following columns:\n\n"
            f"{formatted_col}"
        )

    sheet = wb.active
    total = len(df)

    for i, (_, row) in enumerate(df.iterrows(), start=1):
        if cancel_flag and cancel_flag():
            return False

        populate_sheet(row, mapping, sheet)

        if progress_callback:
            progress = int((i / total) * 90)
            progress_callback(progress)

    if cancel_flag and cancel_flag():
        return False

    wb.save(output_path)

    return True