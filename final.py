import pandas as pd
import random
import openpyxl
from openpyxl.drawing.image import Image as ExcelImage
from docx import Document
from docx.shared import Inches, Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT
from PIL import Image as PILImage
import io
import roman
from docx.shared import Pt

def load_question_bank(file_name):
    return pd.read_excel(file_name, sheet_name='Question Bank')

def select_random_question(unit_questions, remaining_marks):
    available_questions = unit_questions[unit_questions['Marks'] <= remaining_marks]
    return None if available_questions.empty else available_questions.sample(1)

def select_questions_for_unit(questions_pool, unit_marks):
    selected_questions = []
    total_marks_selected = 0
    retries = 0
    max_retries = 50

    while total_marks_selected < unit_marks and retries < max_retries:
        remaining = unit_marks - total_marks_selected
        selected_question = select_random_question(questions_pool, remaining)

        if selected_question is not None:
            marks = selected_question['Marks'].values[0]
            selected_questions.append(selected_question)
            total_marks_selected += marks
        else:
            retries += 1
            if retries >= max_retries:
                print(f"Backtracking failed after {max_retries} retries for this unit.")
                return None
            selected_questions = []
            total_marks_selected = 0

    return (pd.concat(selected_questions), total_marks_selected) if total_marks_selected == unit_marks else None

def generate_question_paper(file_name, unitwise_marks, easy_range, medium_range, max_retries=50):
    question_bank = load_question_bank(file_name)
    last_generated_paper = None
    retries = 0

    while retries < max_retries:
        final_selected_questions = []
        unitwise_total_marks = {}
        difficulty_marks = {'Easy': 0, 'Medium': 0, 'Hard': 0}

        for unit, unit_total_marks in unitwise_marks.items():
            unit_questions = question_bank[question_bank['Unit_No'] == unit]
            result = select_questions_for_unit(unit_questions, unit_total_marks)

            if result is None:
                print(f"Error: Could not meet the total marks requirement for Unit {unit}.")
                return None

            selected_unit_questions, total_marks_selected = result
            final_selected_questions.append(selected_unit_questions)
            unitwise_total_marks[unit] = total_marks_selected

            for _, row in selected_unit_questions.iterrows():
                difficulty_marks[row['Diff_Level']] += row['Marks']

        total_marks = sum(unitwise_marks.values())
        easy_marks = difficulty_marks['Easy']
        medium_marks = difficulty_marks['Medium']

        if (easy_range[0] <= easy_marks <= easy_range[1] and 
            medium_range[0] <= medium_marks <= medium_range[1]):
            print(f"\nValid paper generated after {retries + 1} tries.")
            break
        else:
            print(f"Retry {retries + 1}: Difficulty level not met. Easy: {easy_marks}, Medium: {medium_marks}")
            retries += 1
            last_generated_paper = pd.concat(final_selected_questions)

    return last_generated_paper if retries == max_retries else pd.concat(final_selected_questions)

def convert_excel_to_word_with_images_and_equations(ws, row_number, paragraph):
    headers = {cell.value: idx for idx, cell in enumerate(ws[1])}
    
    if 'Question' not in headers:
        print("Error: 'Question' column is not found.")
        return

    row = ws[row_number]
    question_text = row[headers['Question']].value if row[headers['Question']].value else ""

    while '{{' in question_text and '}}' in question_text:
        start_idx = question_text.index('{{')
        end_idx = question_text.index('}}') + 2
        placeholder = question_text[start_idx:end_idx]
        
        before_placeholder = question_text[:start_idx]
        if before_placeholder:
            run = paragraph.add_run(before_placeholder)
            run.font.size = Pt(11)  # Set font size for text

        image_col_name = placeholder.strip('{}')
        
        if image_col_name in headers:
            for image in ws._images:
                if (image.anchor._from.row == row_number - 1 and 
                    image.anchor._from.col == headers[image_col_name]):
                    img_path = image.ref
                    pil_image = PILImage.open(img_path)
                    image_stream = io.BytesIO()
                    pil_image.save(image_stream, format='PNG')
                    image_stream.seek(0)
                    run = paragraph.add_run()
                    run.add_picture(image_stream, width=Inches(2.5))  # Adjusted image width
                    
        question_text = question_text[end_idx:]
    
    if question_text:
        run = paragraph.add_run(question_text)
        run.font.size = Pt(11)  # Set font size for remaining text

def replace_placeholders_in_template(doc, question_bank, font_size):
    sheet_name = "Question Paper - General Inform"
    general_info = pd.read_excel(question_bank, sheet_name=sheet_name, header=None)
    placeholders = dict(zip(general_info[0], general_info[1]))

    for para in doc.paragraphs:
        for key, value in placeholders.items():
            placeholder = f"{{{{{key}}}}}"
            if placeholder in para.text:
                # Replace the placeholder text
                para.text = para.text.replace(placeholder, str(value))
                
                # Set the font size for the entire paragraph
                for run in para.runs:
                    if key == 'Total Time' or key == 'Total Credits' or key == 'General Instructions':
                        run.font.size = Pt(12)
                    else:
                        run.font.size = Pt(font_size)
                        run.font.bold = True
                    run.font.name = 'Times New Roman'
                

def apply_table_styles(table):
    # Apply table-wide styles
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.LEFT

    # Style the header row
    for cell in table.rows[0].cells:
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = cell.paragraphs[0].runs[0]
        run.font.bold = True
        run.font.size = Pt(11)

    # Style data rows
    for row in table.rows[1:]:
        for i, cell in enumerate(row.cells):
            # Center align Q.No., CO, and Marks columns
            if i in [0, 2, 3]:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            # Left align Question column
            else:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT

            # Set font size
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(11)

def get_column_widths(table):
    """
    Get the widths of columns in a table.
    
    Args:
        table: A docx table object
        
    Returns:
        list: List of column widths in inches
    """
    widths = []
    for cell in table.rows[0].cells:
        if cell.width:
            widths.append(cell.width)
        else:
            # Default width if not specified
            widths.append(Inches(1))
    return widths

def set_column_widths(table, widths):
    """
    Set the widths of columns in a table.
    
    Args:
        table: A docx table object
        widths (list): List of column widths to set
    """
    for i, width in enumerate(widths):
        for row in table.rows:
            row.cells[i].width = width+1

def create_word_document_with_images(final_questions, excel_file, word_file, template_file):
    wb = openpyxl.load_workbook(excel_file)
    ws = wb['Question Bank']
    doc = Document(template_file)

    # Replace placeholders
    replace_placeholders_in_template(doc, excel_file, 16)

    # Store the widths from the third table before clearing placeholders
    template_table_widths = None
    if len(doc.tables) >= 4:  # Check if there's a fourth table (index 3)
        template_table_widths = get_column_widths(doc.tables[3])
        print(f"Retrieved template table widths: {template_table_widths}")
    else:
        print("Warning: Third table not found in template. Using default widths.")
        template_table_widths = [Inches(0.6), Inches(4.7), Inches(0.6), Inches(0.6)]

    # Clear {{Question Here}} placeholder
    for para in doc.paragraphs:
        if "{{Question Here}}" in para.text:
            para.text = ""

    # Add section for each unit
    for unit in final_questions['Unit_No'].unique():
        # Add unit header with proper formatting
        unit_header = doc.add_paragraph()
        unit_header.alignment = WD_ALIGN_PARAGRAPH.LEFT
        unit_run = unit_header.add_run(f"Q. {unit}: Solve the Following")
        unit_run.font.bold = True
        unit_run.font.size = Pt(12)
        unit_header.paragraph_format.space_after = Pt(12)

        # Create and configure table
        table = doc.add_table(rows=1, cols=4)
        table.allow_autofit = False  # Disable autofit to maintain fixed widths
        
        # Apply the stored widths
        set_column_widths(table, template_table_widths)

        # Add header row
        header_cells = table.rows[0].cells
        headers = ['Q. No.', 'Question', 'CO', 'Marks']
        for i, header in enumerate(headers):
            header_cells[i].text = header

        # Add questions
        subquestion = 0
        unit_questions = final_questions[final_questions['Unit_No'] == unit]
        for idx, question_row in unit_questions.iterrows():
            cells = table.add_row().cells

            # Set Q. No. (A, B, C, etc.)
            cells[0].text = chr(65 + subquestion % 26)
            subquestion += 1

            # Add question content
            question_para = cells[1].paragraphs[0]
            convert_excel_to_word_with_images_and_equations(ws, question_row.name + 2, question_para)

            # Set CO and Marks
            cells[2].text = str(question_row['CO'])
            cells[3].text = str(int(question_row['Marks']))

            # Reapply widths to ensure consistency after adding content
            set_column_widths(table, template_table_widths)

        # Apply table styles
        apply_table_styles(table)
        
        # Add spacing after table
        doc.add_paragraph()

    # Add end of paper
    end_para = doc.add_paragraph()
    end_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    end_run = end_para.add_run("*** End of Paper ***")
    end_run.font.bold = True
    end_run.font.size = Pt(12)
    doc.tables[2]._element.getparent().remove(doc.tables[2]._element)
    
    # Save document
    doc.save(word_file)
    print(f'Question paper has been successfully generated: {word_file}')