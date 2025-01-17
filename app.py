from flask import Flask, request, send_file
from flask_cors import CORS
import os
import tempfile
import pandas as pd
from werkzeug.utils import secure_filename
from werkzeug.middleware.proxy_fix import ProxyFix
from starlette.middleware.wsgi import WSGIMiddleware
import zipfile
import io

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}})
app.wsgi_app = ProxyFix(app.wsgi_app)

UPLOAD_FOLDER = os.getenv('UPLOAD_FOLDER', tempfile.gettempdir())
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in {'xlsx', 'xls'}

@app.route('/', methods=['GET'])
def home():
    return "Welcome to the Question Paper Generator API!", 200

@app.route('/generate', methods=['POST'])
def generate_question_paper_app():
    try:
        if 'excel_file' not in request.files:
            return {'error': 'No file part'}, 400
        
        excel_file = request.files['excel_file']
        word_file = request.form.get('word_file')
        set_number = request.form.get('set_number')
        
        print(f"\n******************************************Selected Set Number: {set_number}**********************************\n")

        if not word_file:
            return {'error': 'No word file name provided'}, 400

        if excel_file.filename == '':
            return {'error': 'No selected file'}, 400

        if not allowed_file(excel_file.filename):
            return {'error': 'Invalid file type'}, 400

        excel_filename = secure_filename(excel_file.filename)
        excel_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_filename)
        excel_file.save(excel_path)

        word_filename = secure_filename(f"{word_file}.docx")
        word_filename_master = secure_filename(f"{word_file}_master.docx")
        word_path = os.path.join(app.config['UPLOAD_FOLDER'], word_filename)
        word_path_master = os.path.join(app.config['UPLOAD_FOLDER'], word_filename_master)

        template_file = os.path.join(os.path.dirname(__file__), 'Pimpri Chinchwad Education Trust2.docx')
        template_file_master = os.path.join(os.path.dirname(__file__), 'Pimpri Chinchwad Education Trust.docx')

        if not os.path.exists(template_file) or not os.path.exists(template_file_master):
            return {'error': 'Template file not found'}, 500

        try:
            general_info = pd.read_excel(excel_path, sheet_name='Question Paper - General Inform', header=None, skiprows=11)
        except Exception as e:
            return {'error': f'Error reading Excel file: {str(e)}'}, 400

        unitwise_marks = {}
        total_marks = 0
        count = 0
        condition = 0
        theory_percentage = None

        for row in general_info.itertuples(index=False):
            try:
                unit_str = str(row[0])
                marks = int(row[1]) if pd.notna(row[1]) else 0

                if "theory" in unit_str.lower().strip():
                    theory_percentage = marks
                if "Total Credits" in unit_str:
                    condition = marks * 2
                if "Unit" in unit_str:
                    unit_number = int(unit_str.split()[1])
                    
                    if marks > 0:
                        unitwise_marks[unit_number] = marks
                        total_marks += marks
                        count += 1

                    if count >= condition:
                        break
            except (ValueError, IndexError) as e:
                continue

        easy_percent = 40
        medium_percent = 40
        
        easy_range = (
            int(total_marks * (easy_percent - 5) / 100),
            int(total_marks * (easy_percent + 5) / 100)
        )
        medium_range = (
            int(total_marks * (medium_percent - 5) / 100),
            int(total_marks * (medium_percent + 5) / 100)
        )

        try:
            from final import generate_question_paper, create_word_document_with_images, create_word_document_master_with_images
            final_questions = generate_question_paper(excel_path, unitwise_marks, easy_range, medium_range, theory_percentage)
            
            if final_questions is None:
                return {'error': 'Could not generate questions'}, 400

            final_questions = final_questions.sort_values(by=['Unit_No', 'Marks'], ascending=[True, True])
            
            # Generate both files
            create_word_document_with_images(final_questions, excel_path, word_path, template_file, set_number)
            create_word_document_master_with_images(final_questions, excel_path, word_path_master, template_file_master, set_number)

            # Create a ZIP file in memory
            memory_file = io.BytesIO()
            with zipfile.ZipFile(memory_file, 'w') as zf:
                # Add both files to the ZIP
                zf.write(word_path, word_filename)
                zf.write(word_path_master, word_filename_master)
            
            # Seek to the beginning of the stream
            memory_file.seek(0)

            # Create the response
            response = send_file(
                memory_file,
                mimetype='application/zip',
                as_attachment=True,
                download_name=f'QuestionPaper_Set{set_number}.zip'
            )
            
            response.headers['Access-Control-Allow-Origin'] = '*'
            return response

        except Exception as e:
            return {'error': f'Error generating question paper: {str(e)}'}, 500

    except Exception as e:
        return {'error': f'Unexpected error: {str(e)}'}, 500

    finally:
        # Clean up temporary files
        try:
            if 'excel_path' in locals():
                os.remove(excel_path)
            if 'word_path' in locals() and os.path.exists(word_path):
                os.remove(word_path)
            if 'word_path_master' in locals() and os.path.exists(word_path_master):
                os.remove(word_path_master)
        except Exception as e:
            print(f"Error cleaning up files: {str(e)}")
            pass

asgi_app = WSGIMiddleware(app)

def create_app():
    return asgi_app

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8000))
    app.run(host='0.0.0.0', port=port, debug=True)