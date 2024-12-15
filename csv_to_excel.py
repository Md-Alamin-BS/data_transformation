import os
import pandas as pd
from datetime import datetime

date_suffix = datetime.now().strftime('%d-%m-%Y')

# File paths
input_file = 'excel_files/input_file.csv'
output_file = f'excel_files/file_output_{date_suffix}.xlsx'

def read_input_file(file_path):
    """Dynamically read the input file based on its extension."""
    file_extension = os.path.splitext(file_path)[1].lower()

    if file_extension == '.csv':
        return pd.read_csv(file_path)  # Read CSV file
    elif file_extension in ['.xls', '.xlsx']:
        return pd.read_excel(file_path, sheet_name=0)  # Read the first sheet of Excel file
    else:
        raise ValueError(f"Unsupported file type: {file_extension}")

def process_excel(input_file, output_file):
    # Dynamically read the input file
    before_data = read_input_file(input_file)

    # Extract all unique tags dynamically from the column names
    tags = sorted({col.split()[0] for col in before_data.columns if any(metric in col.lower() for metric in ['precision', 'recall', 'f1']) and 'avg' not in col.lower()})
    avg_tags = sorted({col.split()[0] for col in before_data.columns if 'avg' in col.lower()})

    metrics = ['Precision', 'Recall', 'F1']

    # Prepare the after sorting format
    data = []

    def add_data(tag, metric, before_data):
        """Helper function to add data for a specific tag and metric."""
        before_col = f"{tag} {metric.lower()}"
        after_col = f"{tag} {metric.lower()}"

        # Append the values to the data structure
        before_value = before_data[before_col].iloc[0] if before_col in before_data.columns else ''
        after_value = before_data[after_col].iloc[1] if after_col in before_data.columns else ''

        return before_value, after_value

    # Process normal tags
    for tag in tags:
        row = {'Tag': tag}
        for metric in metrics:
            before_value, after_value = add_data(tag, metric, before_data)
            row[(metric, 'Before')] = before_value
            row[(metric, 'After')] = after_value
            row[(metric, 'RegX')] = ''  # Add blank RegX column
        data.append(row)

    # Add a blank row before avg tags
    data.append({'Tag': '', ('Precision', 'Before'): '', ('Precision', 'After'): '', ('Precision', 'RegX'): '',
                 ('Recall', 'Before'): '', ('Recall', 'After'): '', ('Recall', 'RegX'): '',
                 ('F1', 'Before'): '', ('F1', 'After'): '', ('F1', 'RegX'): ''})

    # Process avg tags
    for tag in avg_tags:
        row = {'Tag': tag}
        for metric in metrics:
            if metric.lower() in tag.lower():  # Only add data for relevant metric
                before_col = [col for col in before_data.columns if tag in col and metric.lower() in col.lower()]
                if before_col:
                    before_value = before_data[before_col[0]].iloc[0]
                    after_value = before_data[before_col[0]].iloc[1]
                else:
                    before_value, after_value = '', ''
                row[(metric, 'Before')] = before_value
                row[(metric, 'After')] = after_value
            else:
                row[(metric, 'Before')] = ''
                row[(metric, 'After')] = ''
            row[(metric, 'RegX')] = ''  # Add blank RegX column
        data.append(row)

    # Create a DataFrame in the desired format with multi-level columns
    df = pd.DataFrame(data)

    # Separate 'Tag' column and create MultiIndex for the remaining columns
    tag_column = df.pop('Tag')
    df.columns = pd.MultiIndex.from_tuples(df.columns)

    # Reset the MultiIndex columns to a single level
    df.columns = [' '.join(col).strip() for col in df.columns.values]

    # Save the output to an Excel file
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Sheet1', index=False)
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']

        # Define formatting styles
        header_format = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'align': 'center', 'valign': 'vcenter', 'font_size': 12.5, 'font_name': 'Aptos'})
        subheader_format = workbook.add_format({'bold': True, 'bg_color': '#B4C6E7', 'align': 'center', 'valign': 'vcenter', 'font_size': 12.5, 'font_name': 'Aptos'})
        data_format = workbook.add_format({'align': 'left', 'valign': 'vcenter', 'font_size': 11, 'font_name': 'Aptos'})
        light_red = workbook.add_format({'bg_color': '#FFC7CE'})
        light_green = workbook.add_format({'bg_color': '#C6EFCE'})

        # Write the 'Tag' column header manually
        worksheet.write(0, 0, 'Tag', header_format)
        worksheet.write(1, 0, '', subheader_format)  # Make the cell below 'Tag' header empty

        # Write the multi-level headers for the remaining columns
        header_top = [col.split()[0] for col in df.columns]
        header_sub = [col.split()[1] if len(col.split()) > 1 else '' for col in df.columns]

        current_top = None
        start_col = 1

        for col_num, (top, sub) in enumerate(zip(header_top, header_sub), start=1):
            if top != current_top:
                if current_top is not None:
                    worksheet.merge_range(0, start_col, 0, col_num - 1, current_top, header_format)
                current_top = top
                start_col = col_num

            worksheet.write(1, col_num, sub, subheader_format)

        # Merge the final group
        if current_top is not None:
            worksheet.merge_range(0, start_col, 0, len(header_top), current_top, header_format)

        # Write the data starting from row 2
        worksheet.write_column(2, 0, tag_column, data_format)
        for row_num, row_data in enumerate(df.values, start=2):
            for col_num, cell_value in enumerate(row_data, start=1):
                worksheet.write(row_num, col_num, cell_value, data_format)

        # Apply conditional formatting
        for col in range(1, len(df.columns), 3):
            before_col = col
            after_col = col + 1
            for row in range(2, len(df) + 2):
                worksheet.conditional_format(row, before_col, row, before_col, {
                    'type': 'cell',
                    'criteria': '>',
                    'value': f'=${chr(65 + after_col)}{row + 1}',
                    'format': light_red
                })
                worksheet.conditional_format(row, after_col, row, after_col, {
                    'type': 'cell',
                    'criteria': '>',
                    'value': f'=${chr(65 + before_col)}{row + 1}',
                    'format': light_green
                })

        # Adjust column widths
        worksheet.set_column(0, 0, 15)  # Tag column
        worksheet.set_column(1, len(df.columns), 12)  # Data columns

    print(f"File successfully created at: {output_file}")

# Process the data
process_excel(input_file, output_file)
