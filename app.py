from flask import Flask, request, send_file, make_response, Blueprint
from flask_cors import CORS
import os
import tempfile
import pandas as pd
from werkzeug.utils import secure_filename
from werkzeug.middleware.proxy_fix import ProxyFix
from starlette.middleware.wsgi import WSGIMiddleware
import zipfile
import io
from flask import Flask, request, jsonify
from dotenv import load_dotenv
import random
import string
from datetime import datetime, timedelta
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from pymongo import MongoClient
import jwt
from werkzeug.security import generate_password_hash, check_password_hash
from functools import wraps
import json
from bson import json_util


# Load environment variables
load_dotenv()

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}})
app.wsgi_app = ProxyFix(app.wsgi_app)

UPLOAD_FOLDER = os.getenv('UPLOAD_FOLDER', tempfile.gettempdir())
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
SECRET_KEY = os.getenv('JWT_SECRET_KEY', 'RAILWAY_TOKEN')  # Add this to your .env file

TOKEN_EXPIRATION = 1 

import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def generate_token(user_id):
    """Generate a JWT token with proper error handling"""
    try:
        # Validate inputs
        if not user_id:
            logger.error("Invalid user_id provided for token generation")
            return None
            
        # Ensure SECRET_KEY is available
        if not SECRET_KEY:
            logger.error("JWT_SECRET_KEY not properly configured")
            return None
            
        payload = {
            'exp': datetime.utcnow() + timedelta(minutes=TOKEN_EXPIRATION),
            'iat': datetime.utcnow(),
            'sub': user_id
        }
        
        token = jwt.encode(
            payload,
            SECRET_KEY,
            algorithm='HS256'
        )
        
        # Log successful token generation
        logger.info(f"Successfully generated token for user: {user_id}")
        return token
        
    except Exception as e:
        logger.error(f"Token generation failed: {str(e)}")
        return None
    

def verify_token(token):
    """Verify the JWT token"""
    try:
        payload = jwt.decode(token, SECRET_KEY, algorithms=['HS256'])
        return payload['sub']
    except jwt.ExpiredSignatureError:
        return None
    except jwt.InvalidTokenError:
        return None

def verify_token(token):
    """Verify the JWT token"""
    try:
        payload = jwt.decode(token, SECRET_KEY, algorithms=['HS256'])
        return payload['sub']
    except jwt.ExpiredSignatureError:
        return None
    except jwt.InvalidTokenError:
        return None

def token_required(f):
    """Decorator to protect routes with JWT token verification"""
    @wraps(f)
    def decorated(*args, **kwargs):
        token = None
        
        # Get token from header
        if 'Authorization' in request.headers:
            auth_header = request.headers['Authorization']
            try:
                token = auth_header.split(" ")[1]  # Bearer <token>
            except IndexError:
                return jsonify({'message': 'Invalid token format'}), 401
        
        if not token:
            return jsonify({'message': 'Token is missing'}), 401
        
        # Verify token
        user_id = verify_token(token)
        if not user_id:
            return jsonify({'message': 'Token is invalid or expired'}), 401
        
        
        return f(*args, **kwargs)
    
    return decorated



def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in {'xlsx', 'xls'}

@app.route('/', methods=['GET'])
def home():
    return "Welcome to the Question Paper Generator API!", 200

@app.route('/generate', methods=['POST'])
@token_required
def generate_question_paper_app():
    response_status = None
    error_message = None
    try:
        if 'excel_file' not in request.files:
            response_status = 400
            error_message = 'No file part'
            return {'error': error_message}, response_status
        
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
            response_status = 200
            log_api_request(
                endpoint='/generate',
                request_data={'set_number': set_number, 'filename': excel_filename},
                response_status=response_status
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

# Define and register the auth blueprint
auth_bp = Blueprint('auth', __name__)
app.register_blueprint(auth_bp)

asgi_app = WSGIMiddleware(app)



# Store OTPs with expiration
otp_store = {}

def generate_otp():
    """Generate a 6-digit OTP"""
    return ''.join(random.choices(string.digits, k=6))

def send_otp_email(otp):
    """Send OTP via email"""
    sender_email = os.getenv('EMAIL_ADDRESS')
    sender_password = os.getenv('EMAIL_PASSWORD')
    receiver_email = "harshad.karale22@pccoepune.org"

    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = "Login OTP Verification"

    body = f"""
    Dear User,

    Your OTP for login verification is: {otp}

    This OTP will expire in 5 minutes.

    Best regards,
    PCCOE Exam System
    """

    message.attach(MIMEText(body, "plain"))

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender_email, sender_password)
            server.send_message(message)
        return True
    except Exception as e:
        print(f"Error sending email: {e}")
        return False

def store_otp(username, otp):
    """Store OTP with 5-minute expiration"""
    otp_store[username] = {
        'otp': otp,
        'expires_at': datetime.now() + timedelta(minutes=5)
    }

def verify_otp(username, otp):
    """Verify OTP and check expiration"""
    if username in otp_store:
        stored_data = otp_store[username]
        if datetime.now() <= stored_data['expires_at'] and stored_data['otp'] == otp:
            del otp_store[username]  # Remove used OTP
            return True
    return False

@app.route('/api/login', methods=['POST'])
def login():
    response_status = None
    error_message = None
    print("Given call to /api/login")
    data = request.get_json()
    username = data.get('username')
    password = data.get('password')
    
    valid_username = os.getenv('LOGIN_USERNAME')
    valid_password = os.getenv('LOGIN_PASSWORD')
    print(username, password)
    print(valid_username, valid_password)
    
    if username == valid_username and password == valid_password:
        otp = generate_otp()
        if send_otp_email(otp):
            store_otp(username, otp)
            response_status = 200
            log_api_request(
                endpoint='/api/login',
                request_data={'username': username},
                response_status=response_status
            )
            return jsonify({
                'status': 'success',
                'message': 'OTP has been sent to your email'
            })
        else:
            response_status = 500
            error_message = 'Failed to send OTP email'
            log_api_request(
                endpoint='/api/login',
                request_data={'username': username},
                response_status=response_status,
                error=error_message
            )
            return jsonify({
                'status': 'error',
                'message': 'Failed to send OTP email'
            }), 500
    
    return jsonify({
        'status': 'error',
        'message': 'Invalid credentials'
    }), 401

@app.route('/api/verify-otp', methods=['POST'])
def verify():
    try:
        data = request.get_json()
        username = data.get('username')
        otp = data.get('otp')
        
        if not username or not otp:
            return jsonify({
                'status': 'error',
                'message': 'Username and OTP are required'
            }), 400
        
        if verify_otp(username, otp):
            token = generate_token(username)
            logger.info(f"Generated token status: {'Success' if token else 'Failed'}")
            
            if not token:
                logger.error("Token generation failed")
                return jsonify({
                    'status': 'error',
                    'message': 'Failed to generate authentication token'
                }), 500
                
            return jsonify({
                'status': 'success',
                'message': 'OTP verified successfully',
                'token': token,
                'expires_in': TOKEN_EXPIRATION * 60
            })
        
        return jsonify({
            'status': 'error',
            'message': 'Invalid or expired OTP'
        }), 401
        
    except Exception as e:
        logger.error(f"OTP verification failed: {str(e)}")
        return jsonify({
            'status': 'error',
            'message': 'Internal server error during authentication'
        }), 500


# MongoDB Atlas connection setup
MONGO_URI = os.getenv('MONGO_URI')  # Add this to your .env file
client = MongoClient(MONGO_URI)
db = client['api_logs']  # Database name
logs_collection = db['endpoint_logs']  # Collection name

def log_api_request(endpoint, request_data=None, response_status=None, error=None):
    """
    Log API requests to MongoDB Atlas
    
    Parameters:
    - endpoint: str - The API endpoint that was accessed
    - request_data: dict - Request data (optional)
    - response_status: int - HTTP response status code
    - error: str - Error message if any
    """
    try:
        log_entry = {
            'timestamp': datetime.utcnow(),
            'endpoint': endpoint,
            'method': request.method,
            'ip_address': request.remote_addr,
            'user_agent': request.headers.get('User-Agent'),
            'request_data': request_data,
            'response_status': response_status,
            'error': error
        }
        
        logs_collection.insert_one(log_entry)
        
    except Exception as e:
        print(f"Error logging to MongoDB: {str(e)}")

@app.route('/api/logs', methods=['GET'])
def get_logs():
    try:
        # Fetch logs from MongoDB, sorted by timestamp in descending order
        logs = logs_collection.find().sort('timestamp', -1).limit(100)
        
        # Convert MongoDB cursor to list and handle ObjectId serialization
        logs_list = json.loads(json_util.dumps(logs))
        
        return jsonify({
            'status': 'success',
            'data': logs_list
        }), 200
        
    except Exception as e:
        return jsonify({
            'status': 'error',
            'message': str(e)
        }), 500
    
@app.route('/api/logout', methods=['POST'])
def logout():
    try:
        # Get token from header
        auth_header = request.headers.get('Authorization')
        
        if not auth_header:
            return jsonify({
                'status': 'error',
                'message': 'No authorization token provided'
            }), 401

        # Extract token
        try:
            token = auth_header.split(" ")[1]
        except IndexError:
            return jsonify({
                'status': 'error',
                'message': 'Invalid token format'
            }), 401

        # Verify token
        user_id = verify_token(token)
        if not user_id:
            return jsonify({
                'status': 'error',
                'message': 'Invalid or expired token'
            }), 401

        # Log successful logout
        log_api_request(
            endpoint='/api/logout',
            request_data={'user_id': user_id},
            response_status=200
        )
        
        return jsonify({
            'status': 'success',
            'message': 'Successfully logged out'
        }), 200
        
    except Exception as e:
        # Log error
        log_api_request(
            endpoint='/api/logout',
            response_status=500,
            error=str(e)
        )
        
        return jsonify({
            'status': 'error',
            'message': f'Logout failed: {str(e)}'
        }), 500


def create_app():
    return asgi_app


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8000))
    app.run(host='0.0.0.0', port=port, debug=True)