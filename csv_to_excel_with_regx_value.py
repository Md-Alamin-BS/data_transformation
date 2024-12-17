import os
import pandas as pd
from datetime import datetime

date_suffix = datetime.now().strftime('%d-%m-%Y')

# File paths
input_file = 'excel_files/input_file.csv'
output_file = f'excel_files/file_output_{date_suffix}_with_regx_value.xlsx'

# Hard-coded RegX values for each metric and tag
precision_regx_values = {
    "boulders": 0, "dheqfailure": 0, "dircontrol": 0, "harddrilling": 0.5, "highrop": 0,
    "holecleaning": 1, "lostcirculation": 0.778, "lowrop": 0.727, "packoff": 0.8,
    "shallowgas": 1, "shallowwater": 1, "stuckpipe": 0.75, "surfeqfailure": 0.6, "tighthole": 0.385,
    "wait": 1, "wellborebreathing": 1, "wellborestability": 1, "wellcontrol": 1,
    "avg_precision_per_tag": 0.64, "avg_precision_per_ddr": 0.824
}

recall_regx_values = {
    "boulders": 1, "dheqfailure": 0, "dircontrol": 1, "harddrilling": 1, "highrop": 1,
    "holecleaning": 0, "lostcirculation": 0.93, "lowrop": 0.8, "packoff": 1,
    "shallowgas": 1, "shallowwater": 1, "stuckpipe": 0.75, "surfeqfailure": 0.643, "tighthole": 0.769,
    "wait": 1, "wellborebreathing": 1, "wellborestability": 1, "wellcontrol": 0.857,
    "avg_recall_per_tag": 0.82, "avg_recall_per_ddr": 0.82
}

f1_regx_values = {
    "boulders": 0, "dheqfailure": 0, "dircontrol": 0, "harddrilling": 0.667, "highrop": 0,
    "holecleaning": 0, "lostcirculation": 0.848, "lowrop": 0.762, "packoff": 0.889,
    "shallowgas": 1, "shallowwater": 1, "stuckpipe": 0.75, "surfeqfailure": 0.621, "tighthole": 0.513,
    "wait": 1, "wellborebreathing": 1, "wellborestability": 1, "wellcontrol": 0.923,
    "avg_f1_per_tag": 0.609, "avg_f1_per_ddr": 0.819
}


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

            # Add RegX values from the dictionaries
            if metric == 'Precision':
                row[(metric, 'RegX')] = precision_regx_values.get(tag, '')
            elif metric == 'Recall':
                row[(metric, 'RegX')] = recall_regx_values.get(tag, '')
            elif metric == 'F1':
                row[(metric, 'RegX')] = f1_regx_values.get(tag, '')

        data.append(row)

    # Add a blank row before avg tags
    data.append({
        'Tag': '', 
        ('Precision', 'Before'): '', ('Precision', 'After'): '', ('Precision', 'RegX'): '',
        ('Recall', 'Before'): '', ('Recall', 'After'): '', ('Recall', 'RegX'): '',
        ('F1', 'Before'): '', ('F1', 'After'): '', ('F1', 'RegX'): ''
    })

    # Process avg tags
    for tag in avg_tags:
        row = {'Tag': tag}
        for metric in metrics:
            # Find the appropriate columns
            relevant_col = [col for col in before_data.columns if tag in col and metric.lower() in col.lower()]
            if relevant_col:
                before_value = before_data[relevant_col[0]].iloc[0]
                after_value = before_data[relevant_col[0]].iloc[1]
            else:
                before_value, after_value = '', ''

            row[(metric, 'Before')] = before_value
            row[(metric, 'After')] = after_value

            # Add RegX values for avg tags
            if metric == 'Precision':
                row[(metric, 'RegX')] = precision_regx_values.get(tag, '')
            elif metric == 'Recall':
                row[(metric, 'RegX')] = recall_regx_values.get(tag, '')
            elif metric == 'F1':
                row[(metric, 'RegX')] = f1_regx_values.get(tag, '')

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
        light_blue = workbook.add_format({'bg_color': '#ADD8E6'})

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

        # Write the data starting from row 2 (0-based)
        worksheet.write_column(2, 0, tag_column, data_format)
        for row_num, row_data in enumerate(df.values, start=2):
            for col_num, cell_value in enumerate(row_data, start=1):
                worksheet.write(row_num, col_num, cell_value, data_format)

        # Apply conditional formatting for Before/After (mutually exclusive)
        # Each metric has 3 columns: Before, After, RegX in that order.
        # The code increments by 3 columns for each metric set.
        # For example, if the first metric (Precision) starts at column 1:
        # Before=1, After=2, RegX=3, next metric would start at column 4.
        for col in range(1, len(df.columns)+1, 3):
            before_col = col
            after_col = col + 1
            regx_col = col + 2

            # Conditional formatting: If Before > After -> Before cell red
            # If After > Before -> After cell green
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

            # Conditional formatting for RegX:
            # Highlight RegX (light blue) if RegX > Before AND RegX > After.
            # We'll use a formula-based condition.
            # Assuming the first data row is Excel row 3, the formula:
            # =AND($D3>$B3,$D3>$C3) for example if D is RegX, B and C are Before/After.
            # We'll apply the formatting from row 2 (python indexing) which is Excel row 3 downwards.
            first_data_row_excel = 3
            range_start_row = 2  # Python index for first data row
            range_end_row = len(df) + 1  # Python end row
            before_letter = chr(65 + before_col)
            after_letter = chr(65 + after_col)
            regx_letter = chr(65 + regx_col)

            # Apply conditional formatting for RegX column
            worksheet.conditional_format(range_start_row, regx_col, range_end_row, regx_col, {
            'type': 'formula',
            'criteria': f'=AND(${regx_letter}{first_data_row_excel}>0, ${regx_letter}{first_data_row_excel}>${before_letter}{first_data_row_excel}, ${regx_letter}{first_data_row_excel}>${after_letter}{first_data_row_excel})',
            'format': light_blue
})

            

        # Adjust column widths
        worksheet.set_column(0, 0, 15)  # Tag column
        worksheet.set_column(1, len(df.columns), 12)  # Data columns

    print(f"File successfully created at: {output_file}")

# Process the data
process_excel(input_file, output_file)
