from flask import Flask, request, render_template, redirect, url_for, send_file
from werkzeug.utils import secure_filename
import os
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Font, Border, Alignment, Protection, NamedStyle, PatternFill
from openpyxl.cell.cell import MergedCell

app = Flask(__name__)

# Folder to store uploaded files
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Ensure the upload folder exists
# if not os.path.exists(UPLOAD_FOLDER):
#     os.makedirs(UPLOAD_FOLDER)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Get form data
        data_file = request.files['data-file']
        report_date = request.form['report-date']
        file_name = request.form['file-name']

        # Save the uploaded file
        filename = secure_filename(data_file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        data_file.save(file_path)

        # Process the data
        members = read_process_data(file_path)  # Update to return members

        read_process_data(os.path.join(app.config['UPLOAD_FOLDER'], filename))
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], f'{file_name}.xlsx')
        generate_wb(members, output_path, report_date)

        # Set a flash message to indicate success
        # flash('File has been processed successfully.')

        # return redirect(url_for('index'))
     # Send the generated file as a response
        return send_file(output_path, as_attachment=True, attachment_filename=f'{file_name}.xlsx')

    return render_template('index.html')

# 1. Import Libraries
def copy_styles(source_cell, target_cell):
    if source_cell.has_style:
        target_cell.font = Font(
            name=source_cell.font.name,
            size=source_cell.font.size,
            bold=source_cell.font.bold,
            italic=source_cell.font.italic,
            vertAlign=source_cell.font.vertAlign,
            underline=source_cell.font.underline,
            strike=source_cell.font.strike,
            color=source_cell.font.color
        )
        target_cell.border = Border(
            left=source_cell.border.left,
            right=source_cell.border.right,
            top=source_cell.border.top,
            bottom=source_cell.border.bottom,
            diagonal=source_cell.border.diagonal,
            diagonal_direction=source_cell.border.diagonal_direction,
            outline=source_cell.border.outline,
            vertical=source_cell.border.vertical,
            horizontal=source_cell.border.horizontal
        )
        target_cell.fill = PatternFill(
            fill_type=source_cell.fill.fill_type,
            start_color=source_cell.fill.start_color,
            end_color=source_cell.fill.end_color
        )
        target_cell.alignment = Alignment(
            horizontal=source_cell.alignment.horizontal,
            vertical=source_cell.alignment.vertical,
            text_rotation=source_cell.alignment.text_rotation,
            wrap_text=source_cell.alignment.wrap_text,
            shrink_to_fit=source_cell.alignment.shrink_to_fit,
            indent=source_cell.alignment.indent
        )
        target_cell.protection = Protection(
            locked=source_cell.protection.locked,
            hidden=source_cell.protection.hidden
        )
        target_cell.number_format = source_cell.number_format

def format_excel(activesheet, destination):
    for row in activesheet.iter_rows():
        for cell in row:
            new_cell = destination[cell.coordinate]
            new_cell.value = cell.value
            copy_styles(cell, new_cell)
            
    # Copy column widths
    for col in activesheet.column_dimensions:
        destination.column_dimensions[col].width = activesheet.column_dimensions[col].width

    # Copy row heights
    for row in activesheet.row_dimensions:
        destination.row_dimensions[row].height = activesheet.row_dimensions[row].height
    
    for row in activesheet.iter_rows():
        for cell in row:
            new_cell = destination[cell.coordinate]
            new_cell.value = cell.value
            copy_styles(cell, new_cell)
            
    # Copy merged cells
    for merged_cell in activesheet.merged_cells.ranges:
        destination.merge_cells(str(merged_cell))
    
    for row in activesheet.iter_rows():
        for cell in row:
            new_cell = destination[cell.coordinate]
            if not isinstance(cell, MergedCell):
                new_cell.value = cell.value
            copy_styles(cell, new_cell)
    
    # Copy images
    for image in activesheet._images:
        img = Image(image.ref)
        destination.add_image(img, image.anchor)

def read_process_data(data_file):
    # 4. Read Data
    excel_file = pd.ExcelFile(data_file)

    df = {}
    last_no = 0

    for sheet_name in excel_file.sheet_names:
        frame = pd.read_excel(excel_file, sheet_name, header=None, usecols="C, M, O, Q, S, U, W, Y, AA, AC, AE, AG", skiprows=12, nrows=162,
            names=range(1, 13))  # Assuming 12 columns: C, M, O, Q, S, U, W, Y, AA, AC, AE, AG

        df[sheet_name] = frame
        last_no = sheet_name

    # print(df)

    # 5. Transfer & Group Data 
    members = {}
    for x in range(0, 146):
        curName = df[last_no].loc[x, 1]
        if pd.notna(curName):
            members[curName] = []

    for key in df.keys():
        names = [df[last_no].loc[x, 1] for x in range(0, 146, 10)]
        mo_scores = []
        others_scores = []
        member_sum_mo = 0
        member_sum_oth = 0

        for x in range(0, 146):
            m_val = df[key].loc[x, 2]
            o_val = df[key].loc[x, 3]
            q_val = df[key].loc[x, 4]
            s_val = df[key].loc[x, 5]
            u_val = df[key].loc[x, 6]
            w_val = df[key].loc[x, 7]
            y_val = df[key].loc[x, 8]
            aa_val = df[key].loc[x, 9]
            ac_val = df[key].loc[x, 10]
            ae_val = df[key].loc[x, 11]
            ag_val = df[key].loc[x, 12]

            if x % 10 == 0 and x != 0 or x == 145:
                mo_scores.append(member_sum_mo)
                others_scores.append(member_sum_oth)
                member_sum_mo = 0
                member_sum_oth = 0
            else:
                member_sum_mo += pd.notna(m_val) + pd.notna(o_val)
                member_sum_oth += pd.notna(q_val) + pd.notna(s_val) + pd.notna(u_val) + pd.notna(w_val) + pd.notna(y_val) + \
                    pd.notna(aa_val) + pd.notna(ac_val) + pd.notna(ae_val) + pd.notna(ag_val)

        for i in range(0, len(names), 1):
            curName = names[i]
            if (pd.notna(curName)):
                members[curName].append({key: {
                    'mo': mo_scores[i] * 1.0,
                    'others': others_scores[i] * 1.75
                    }
                })
    return members

def generate_wb(fulldict, output_path, report_date):
    template_file = 'Template.xlsx'
    template_wb = load_workbook(template_file)
    template_sheet = template_wb['Template']
    
    new_wb = Workbook()
    new_wb.remove(new_wb.active)
    
    allNames = list(fulldict.keys())
        
    for name in allNames:
        new_sheet = new_wb.create_sheet(title=name)
        format_excel(template_sheet, new_sheet)
        
        values_mo = []
        values_others = []

        for x in range(0, 31):
            values_mo.append(fulldict[name][x]['{}'.format(x + 1)]['mo'])
            values_others.append(fulldict[name][x]['{}'.format(x + 1)]['others'])
            
        new_sheet['D5'] = name
        new_sheet['K6'] = report_date
        
        for x in range(0, 31):
            new_sheet['G{}'.format(11 + x)] = values_mo[x]
            new_sheet['H{}'.format(11 + x)] = values_others[x]
    
    # new_wb.save('{}.xlsx'.format(newfile))
    new_wb.save(output_path)
    print('A new Excel document has been saved')

if __name__ == '__main__':
    app.run(debug=True)
