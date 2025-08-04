import openpyxl
from openpyxl.utils import get_column_letter
from random import shuffle
import os
from pathlib import Path

def ensure_output_folder():
    output_dir = Path('output')
    output_dir.mkdir(exist_ok=True)
    return output_dir

def validate_file(file_name):
    input_dir = Path('input')
    file_path = input_dir / file_name
    if not file_path.exists():
        raise FileNotFoundError(f"File {file_name} not found in input folder")
    if not file_name.endswith('.xlsx'):
        raise ValueError("File must be an .xlsx Excel file")
    return file_path

def get_range_input(prompt, max_value):
    while True:
        try:
            start, end = map(int, input(prompt).split('-'))
            if start < 1 or end < start or (max_value and end > max_value):
                print(f"Invalid range. Must be between 1 and {max_value if max_value else 'end'}.")
                continue
            return start, end  # 1-based indexing for openpyxl
        except ValueError:
            print("Invalid input. Enter range as 'start-end' (e.g., 1-10).")

def copy_cell_style(source_cell, target_cell):
    """Copy cell formatting from source to target without deepcopy."""
    # Copy font properties
    if source_cell.font:
        target_cell.font = openpyxl.styles.Font(
            name=source_cell.font.name,
            size=source_cell.font.size,
            bold=source_cell.font.bold,
            italic=source_cell.font.italic,
            vertAlign=source_cell.font.vertAlign,
            underline=source_cell.font.underline,
            strike=source_cell.font.strike,
            color=source_cell.font.color
        )
    else:
        target_cell.font = None

    # Copy border properties
    if source_cell.border:
        target_cell.border = openpyxl.styles.Border(
            left=source_cell.border.left,
            right=source_cell.border.right,
            top=source_cell.border.top,
            bottom=source_cell.border.bottom,
            diagonal=source_cell.border.diagonal,
            diagonal_direction=source_cell.border.diagonal_direction,
            outline=source_cell.border.outline
        )
    else:
        target_cell.border = None

    # Copy fill properties
    if source_cell.fill:
        target_cell.fill = openpyxl.styles.PatternFill(
            fill_type=source_cell.fill.fill_type,
            start_color=source_cell.fill.start_color,
            end_color=source_cell.fill.end_color
        )
    else:
        target_cell.fill = None

    # Copy number format
    target_cell.number_format = source_cell.number_format

    # Copy protection properties
    if source_cell.protection:
        target_cell.protection = openpyxl.styles.Protection(
            locked=source_cell.protection.locked,
            hidden=source_cell.protection.hidden
        )
    else:
        target_cell.protection = None

    # Copy alignment properties
    if source_cell.alignment:
        target_cell.alignment = openpyxl.styles.Alignment(
            horizontal=source_cell.alignment.horizontal,
            vertical=source_cell.alignment.vertical,
            text_rotation=source_cell.alignment.text_rotation,
            wrap_text=source_cell.alignment.wrap_text,
            shrink_to_fit=source_cell.alignment.shrink_to_fit,
            indent=source_cell.alignment.indent
        )
    else:
        target_cell.alignment = None


def shuffle_excel_rows():
    try:
        # Ensure input and output directories exist
        Path('input').mkdir(exist_ok=True)
        output_dir = ensure_output_folder()

        # Get file name
        file_name = input("Enter the Excel file name (e.g., data.xlsx): ").strip()
        file_path = validate_file(file_name)

        # Load workbook with openpyxl
        wb = openpyxl.load_workbook(file_path)
        ws = wb.active  # Use the active sheet
        max_rows, max_cols = ws.max_row, ws.max_column

        # Get row range
        row_prompt = f"Enter row range to shuffle (1-{max_rows}, e.g., 1-10): "
        start_row, end_row = get_range_input(row_prompt, max_rows)

        # Get column range
        col_prompt = f"Enter column range to shuffle (1-{max_cols}, e.g., 1-5): "
        start_col, end_col = get_range_input(col_prompt, max_cols)

        # Read data and styles from the specified range
        data = []
        for row in range(start_row, end_row + 1):
            row_data = []
            for col in range(start_col, end_col + 1):
                cell = ws.cell(row=row, column=col)
                row_data.append((cell.value, cell))  # Store value and cell object for styles
            data.append(row_data)

        # Shuffle the rows
        shuffle(data)

        # Write shuffled data back to the same range, preserving styles
        for row_idx, row_data in enumerate(data, start=start_row):
            for col_idx, (value, source_cell) in enumerate(row_data, start=start_col):
                target_cell = ws.cell(row=row_idx, column=col_idx)
                target_cell.value = value
                copy_cell_style(source_cell, target_cell)

        # Set row heights to None to allow Excel to auto-fit
        for row_idx in range(start_row, end_row + 1):
            ws.row_dimensions[row_idx].height = None

        # Save to output folder with '_shuffled' suffix
        output_file = output_dir / f"{file_name.rsplit('.', 1)[0]}_shuffled.xlsx"
        wb.save(output_file)
        print(f"Shuffled file saved as {output_file}")

    except Exception as e:
        print(f"Error: {str(e)}")


if __name__ == "__main__":
    shuffle_excel_rows()
