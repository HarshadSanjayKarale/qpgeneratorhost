from flask import Flask, request, send_file
from flask_cors import CORS
import os
import tempfile
import pandas as pd
from werkzeug.utils import secure_filename
from asgiref.wsgi import WsgiToAsgi

# Initialize Flask app and apply CORS
app = Flask(__name__)
CORS(app)

# Configure upload folder
UPLOAD_FOLDER = tempfile.gettempdir()
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route('/generate', methods=['POST'])
def generate_question_paper():
    if 'excel_file' not in request.files:
        return 'No file part', 400
    
    excel_file = request.files['excel_file']
    word_file = request.form['word_file']

    if excel_file.filename == '':
        return 'No selected file', 400

    if excel_file:
        # Save uploaded Excel file to temp directory
        excel_filename = secure_filename(excel_file.filename)
        excel_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_filename)
        excel_file.save(excel_path)

        # Set output Word file path
        word_filename = secure_filename(f"{word_file}.docx")
        word_path = os.path.join(app.config['UPLOAD_FOLDER'], word_filename)

        # Template for the Word file
        template_file = 'Pimpri Chinchwad Education Trust2.docx'

        # Initialize variables for unitwise marks and total marks
        unitwise_marks = {}
        total_marks = 0

        # Read unitwise marks from the specified Excel sheet
        general_info = pd.read_excel(excel_path, sheet_name='Question Paper - General Inform', header=None, skiprows=13)

        count = 0
        condition = 4
        for row in general_info.itertuples(index=False):
            unit_str = row[0]
            marks = int(row[1])
            if "Total Credits" in unit_str:
                condition = marks * 2
            if "Unit" in unit_str:
                unit_number = int(unit_str.split()[1])  
                
                if marks is not None and marks > 0:
                    unitwise_marks[unit_number] = marks 
                    total_marks += marks 
                    count += 1  

                if count >= condition:
                    break

        # Difficulty level percentages
        easy_percent = 40
        medium_percent = 40
        
        # Calculate mark ranges for easy and medium difficulty questions
        easy_range = (
            int(total_marks * (easy_percent - 5) / 100),
            int(total_marks * (easy_percent + 5) / 100)
        )
        medium_range = (
            int(total_marks * (medium_percent - 5) / 100),
            int(total_marks * (medium_percent + 5) / 100)
        )

        # Import and use generate_question_paper and create_word_document_with_images functions
        from final import generate_question_paper, create_word_document_with_images

        # Generate questions and create Word document
        final_questions = generate_question_paper(excel_path, unitwise_marks, easy_range, medium_range)
        if final_questions is not None:
            create_word_document_with_images(final_questions, excel_path, word_path, template_file)
            return send_file(word_path, as_attachment=True)
        else:
            return 'Could not generate the question paper.', 400

    return 'Error processing the file', 400

# Wrap the Flask app with WsgiToAsgi for ASGI compatibility
asgi_app = WsgiToAsgi(app)

if __name__ == '__main__':
    app.run(debug=True)
