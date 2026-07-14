import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.exceptions import CellCoordinatesException


PREOP_TARGET_COLUMN = "PreOp Target (Previous)"
POSTOP_TARGET_COLUMN = "PostOp Target (Current)"
STATUS_COLUMN = "PreOp or PostOp"


def load_mapping(mapping_path):
    """Load and validate the Excel mapping file."""

    try:
        mapping_df = pd.read_excel(
            mapping_path,
            dtype=str,
        )

    except FileNotFoundError as error:
        raise FileNotFoundError(
            f"The mapping Excel file could not be found:\n{mapping_path}"
        ) from error

    except Exception as error:
        raise ValueError(
            f"The mapping Excel file could not be opened:\n{error}"
        ) from error

    # Remove accidental spaces from column headings.
    mapping_df.columns = [
        str(column).strip()
        for column in mapping_df.columns
    ]

    required_columns = {
        "Export",
        PREOP_TARGET_COLUMN,
        POSTOP_TARGET_COLUMN,
    }

    missing_columns = required_columns - set(mapping_df.columns)

    if missing_columns:
        formatted_columns = "\n".join(
            f"• {column}"
            for column in sorted(missing_columns)
        )

        raise ValueError(
            "The mapping Excel file is missing the following "
            f"required column(s):\n\n{formatted_columns}"
        )

    # Remove completely empty spreadsheet rows.
    mapping_df = mapping_df.dropna(how="all")

    return mapping_df


def get_assessment_status(input_row):
    """Read whether the REDCap export is PreOp or PostOp."""

    if STATUS_COLUMN not in input_row.index:
        raise ValueError(
            "The REDCap export does not contain the "
            f"'{STATUS_COLUMN}' column."
        )

    status = input_row[STATUS_COLUMN]

    if pd.isna(status):
        raise ValueError(
            f"The '{STATUS_COLUMN}' field is blank."
        )

    normalised_status = (
        str(status)
        .strip()
        .lower()
        .replace(" ", "")
        .replace("-", "")
    )

    if normalised_status == "preop":
        return "preop"

    if normalised_status == "postop":
        return "postop"

    raise ValueError(
        f"The '{STATUS_COLUMN}' field contains an "
        f"unrecognised value: {status!r}.\n\n"
        "Expected PreOp or PostOp."
    )


def is_blank_mapping_value(value):
    """
    Return True when a mapping cell should be ignored.

    This handles:
    - Empty Excel cells
    - Blank text
    - A dash used as a placeholder
    """

    if pd.isna(value):
        return True

    text = str(value).strip()

    return text in {"", "-"}


def get_usable_mapping_rows(mapping_df, target_column):
    """Return source-column and target-cell pairs for this assessment."""

    usable_rows = []

    for _, mapping_row in mapping_df.iterrows():
        source_value = mapping_row["Export"]
        target_value = mapping_row[target_column]

        if is_blank_mapping_value(source_value):
            continue

        if is_blank_mapping_value(target_value):
            continue

        source_column = str(source_value).strip()
        target_cell = str(target_value).strip()

        usable_rows.append(
            (source_column, target_cell)
        )

    return usable_rows


def validate_columns(input_df, usable_mapping_rows):
    """Return mapped REDCap columns missing from the export."""

    required_columns = [
        source_column
        for source_column, _ in usable_mapping_rows
    ]

    missing_columns = [
        column
        for column in required_columns
        if column not in input_df.columns
    ]

    return sorted(set(missing_columns))

def is_invalid_merged_target(worksheet, target_cell):
    """
    Return True if the target is inside a merged range but is
    not the top-left cell of that range.
    """

    for merged_range in worksheet.merged_cells.ranges:
        if target_cell in merged_range:
            top_left_cell = worksheet.cell(
                row=merged_range.min_row,
                column=merged_range.min_col,
            ).coordinate

            return target_cell != top_left_cell

    return False

def validate_target_cells(worksheet, usable_mapping_rows):
    """Check that every mapping target is a valid Excel cell."""

    invalid_targets = []

    for source_column, target_cell in usable_mapping_rows:
        try:
            worksheet[target_cell]

            if is_invalid_merged_target(
                    worksheet,
                    target_cell,
            ):
                invalid_targets.append(
                    f"{source_column} → {target_cell} "
                    "(not the top-left cell of a merged range)"
                )

        except (ValueError, CellCoordinatesException):
            invalid_targets.append(
                f"{source_column} → {target_cell}"
            )

    if invalid_targets:
        formatted_targets = "\n".join(
            f"• {target}"
            for target in invalid_targets
        )

        raise ValueError(
            "The mapping file contains invalid target "
            f"cell references:\n\n{formatted_targets}"
        )

def convert(
    input_path,
    mapping_path,
    template_path,
    output_path,
    progress_callback=None,
    cancel_flag=None,
):
    """Convert one REDCap CSV record into the Excel template."""

    # --------------------------------------------------
    # 1. Read the REDCap export
    # --------------------------------------------------

    try:
        input_df = pd.read_csv(
            input_path,
            dtype=object,
        )

    except FileNotFoundError as error:
        raise FileNotFoundError(
            f"The input file could not be found:\n{input_path}"
        ) from error

    except Exception as error:
        raise ValueError(
            f"The REDCap export could not be opened:\n{error}"
        ) from error

    if input_df.empty:
        raise ValueError(
            "The selected REDCap export contains no records."
        )

    input_df.columns = [
        str(column).strip()
        for column in input_df.columns
    ]

    if len(input_df) > 1:
        raise ValueError(
            "The selected REDCap export contains more than one "
            "record.\n\n"
            "This version of the converter processes one patient "
            "record at a time."
        )

    input_row = input_df.iloc[0]

    if progress_callback:
        progress_callback(10)

    # --------------------------------------------------
    # 2. Determine PreOp or PostOp
    # --------------------------------------------------

    assessment_status = get_assessment_status(input_row)

    if assessment_status == "preop":
        target_column = PREOP_TARGET_COLUMN
    else:
        target_column = POSTOP_TARGET_COLUMN

    if progress_callback:
        progress_callback(20)

    # --------------------------------------------------
    # 3. Load the mapping
    # --------------------------------------------------

    mapping_df = load_mapping(mapping_path)

    usable_mapping_rows = get_usable_mapping_rows(
        mapping_df,
        target_column,
    )

    if not usable_mapping_rows:
        raise ValueError(
            f"No {assessment_status.title()} targets were "
            "found in the mapping file."
        )

    missing_columns = validate_columns(
        input_df,
        usable_mapping_rows,
    )

    if missing_columns:
        formatted_columns = "\n".join(
            f"• {column}"
            for column in missing_columns
        )

        raise ValueError(
            "The REDCap export is missing the following "
            f"required column(s):\n\n{formatted_columns}"
        )

    if progress_callback:
        progress_callback(30)

    # --------------------------------------------------
    # 4. Open the physical-exam template
    # --------------------------------------------------

    try:
        workbook = load_workbook(template_path)

    except FileNotFoundError as error:
        raise FileNotFoundError(
            "The Physical Exam template file could not be "
            f"found:\n{template_path}"
        ) from error

    except Exception as error:
        raise ValueError(
            "The Physical Exam template file could not be "
            f"opened:\n{error}"
        ) from error

    worksheet = workbook.active

    validate_target_cells(
        worksheet,
        usable_mapping_rows,
    )

    # --------------------------------------------------
    # 5. Populate the template
    # --------------------------------------------------

    total_fields = len(usable_mapping_rows)

    for index, (
        source_column,
        target_cell,
    ) in enumerate(usable_mapping_rows, start=1):

        if cancel_flag and cancel_flag():
            raise InterruptedError(
                "The conversion was cancelled."
            )

        value = input_row[source_column]

        if not pd.isna(value):
            if not (
                isinstance(value, str)
                and value.strip() == ""
            ):
                worksheet[target_cell] = value

        if progress_callback:
            completed_fraction = index / total_fields
            progress = 30 + int(
                completed_fraction * 60
            )
            progress_callback(progress)

    # Check once more before saving.
    if cancel_flag and cancel_flag():
        raise InterruptedError(
            "The conversion was cancelled."
        )

    # --------------------------------------------------
    # 6. Save the completed workbook
    # --------------------------------------------------

    try:
        workbook.save(output_path)

    except PermissionError as error:
        raise PermissionError(
            "The output file could not be saved.\n\n"
            "Make sure the output file is not already open "
            "in Excel."
        ) from error

    except Exception as error:
        raise ValueError(
            f"The converted workbook could not be saved:\n{error}"
        ) from error

    if progress_callback:
        progress_callback(100)

    return {
        "output_path": output_path,
        "assessment_status": assessment_status,
    }