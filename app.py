from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import pandas as pd
import os
from werkzeug.utils import secure_filename
import datetime

# Import your biometric processing function
# Ensure 'biometric_processor.py' is in the same directory as 'app.py'
from biometric_processor import process_biometric_data_for_excel_dashboard

app = Flask(__name__)

# --- New Additions for File Upload ---
UPLOAD_FOLDER = 'uploads' # Define the folder to save uploads

# Create the uploads directory if it doesn't exist
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
# --- End of New Additions for File Upload ---


CORS(app, resources={
    r"/api/*": {
        "origins": [
            "http://localhost:3000",
            "http://127.0.0.1:3000",
            "https://biometrics-dashboard-iq8y63fym-amaan-rathores-projects.vercel.app", 
            "https://biometrics-dashboard.vercel.app", # Exact Vercel URL
            "https://biometrics-dashboard-git-main-amaan-rathores-projects.vercel.app"
        ],
        "methods": ["GET", "POST", "PUT", "DELETE", "OPTIONS"],  # Include OPTIONS for preflight
        "allow_headers": ["Content-Type", "Authorization", "Accept", "Origin", "X-Requested-With"],
        "supports_credentials": False,
        "expose_headers": []  # Optional, if frontend needs specific headers
    }
})

# --- Helper function to find the latest processed Excel dashboard ---
def get_latest_processed_excel_path():
    """Finds the most recently created interactive Excel dashboard."""
    files = [f for f in os.listdir(app.config['UPLOAD_FOLDER']) if f.startswith('interactive_attendance_charts_') and f.endswith('.xlsx')]
    if not files:
        return None
    
    # Sort by modification time, newest first
    files.sort(key=lambda x: os.path.getmtime(os.path.join(app.config['UPLOAD_FOLDER'], x)), reverse=True)
    return os.path.join(app.config['UPLOAD_FOLDER'], files[0])

# --- Function to load biometric data from the latest processed Excel file ---
def load_biometric_data_from_latest_excel():
    """Load biometric data from the 'Main_Data' sheet of the latest processed Excel file."""
    file_path = get_latest_processed_excel_path()
    
    if file_path is None:
        print(f"Warning: No processed Excel dashboard found in {app.config['UPLOAD_FOLDER']}. Returning empty DataFrame.")
        expected_columns = [
            'Employee_ID', 'Employee_Name', 'Date', 'Check_In', 'Check_Out',
            'Working_Hours', 'Late_Minutes', 'Status', 'Late_Flag', 'Is_Late'
        ]
        
        # Apply expected columns, truncating or padding as needed
        df.columns = expected_columns[:len(df.columns)]
        if len(df.columns) < len(expected_columns):
            # Pad with 'Extra_X' if fewer columns than expected
            df.columns = list(df.columns) + [f'Extra_{i}' for i in range(len(df.columns), len(expected_columns))]
        elif len(df.columns) > len(expected_columns):
            # Truncate if more columns than expected
            df = df.iloc[:, :len(expected_columns)]
            df.columns = expected_columns

    print("Attempting to load biometric data from latest Excel:", file_path)
    try:
        # Load the 'Main_Data' sheet, headers are at row 3 (index 2)
        df = pd.read_excel(file_path, sheet_name='Main_Data', header=2) 
        
        # Define expected headers from your biometric_processor.py output
        expected_headers = ['Employee_ID', 'Employee_Name', 'Date', 'Check_In', 'Check_Out',
                            'Working_Hours', 'Late_Minutes', 'Status', 'Late_Flag', 'Is_Late']
        
        # Rename columns to match expected headers if there's a mismatch
        if len(df.columns) >= len(expected_headers):
            df.columns = expected_headers[:len(df.columns)]
        else: # Handle cases where fewer columns are loaded than expected
            current_cols = list(df.columns)
            new_cols = current_cols + [h for h in expected_headers if h not in current_cols]
            df = df.reindex(columns=new_cols, fill_value='N/A')
            df.columns = expected_headers # Ensure final columns are exactly as expected

        print("Loaded Processed Data (first 5 rows):")
        print(df.head().to_string())
        print("Columns after loading processed Excel sheet:", df.columns.tolist())
        
        # Ensure correct types and handle potential NaNs as before
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
        df['Date'] = df['Date'].dt.strftime('%Y-%m-%d').fillna('N/A')
        df['Employee_ID'] = df['Employee_ID'].astype(str).str.strip()
        df = df.fillna('N/A').infer_objects(copy=False)
        
        print("Final Processed Data for API (first 5 rows):")
        print(df.head().to_string())
        return df
        
    except Exception as e:
        print(f"Error loading processed Excel file '{file_path}': {e}")
        import traceback
        traceback.print_exc()
        expected_columns = [
            'Employee_ID', 'Employee_Name', 'Date', 'Check_In', 'Check_Out',
            'Working_Hours', 'Late_Minutes', 'Status', 'Late_Flag', 'Is_Late'
        ]
        return pd.DataFrame(columns=expected_columns)


@app.route('/')
def home():
    """Home route with API information"""
    return jsonify({
        'message': 'Welcome to the Biometric Search API',
        'endpoints': {
            '/api/employees': 'GET - Get all employees from latest processed data',
            '/api/search': 'GET - Search records from latest processed data with query parameters',
            '/api/upload': 'POST - Upload and process biometric raw files to generate new Excel dashboard',
            '/api/download-latest-dashboard': 'GET - Download the latest generated interactive Excel dashboard'
        },
        'status': 'active'
    })

@app.route('/api/health', methods=['GET'])
def health_check():
    """Health check endpoint"""
    current_df = load_biometric_data_from_latest_excel() # Load data for health check
    latest_excel = get_latest_processed_excel_path()
    return jsonify({
        'status': 'healthy',
        'data_loaded_from_latest_excel': not current_df.empty,
        'total_records_in_latest_excel': len(current_df) if not current_df.empty else 0,
        'latest_excel_dashboard': os.path.basename(latest_excel) if latest_excel else 'None found'
    })

@app.route('/api/employees', methods=['GET'])
def get_employees():
    """Get all unique employees from the latest processed data."""
    print("GET /api/employees endpoint hit")
    current_df = load_biometric_data_from_latest_excel() # Load data dynamically
    if current_df.empty:
        print("No data available in DataFrame for employees.")
        return jsonify({'employees': [], 'message': 'No processed data available. Please upload and process new files.'}), 200
    
    try:
        employees = current_df[['Employee_ID', 'Employee_Name']].drop_duplicates().to_dict(orient='records')
        print(f"Returning {len(employees)} unique employees from latest data.")
        return jsonify({'employees': employees})
    except Exception as e:
        print(f"Error getting employees: {e}")
        return jsonify({'error': 'Failed to retrieve employees', 'message': str(e)}), 500

@app.route('/api/search', methods=['GET'])
def search_records():
    """Search records based on employee ID and date range from the latest processed data."""
    print("GET /api/search endpoint hit")
    current_df = load_biometric_data_from_latest_excel() # Load data dynamically
    if current_df.empty:
        print("No data available in DataFrame for search.")
        return jsonify({
            'records': [],
            'message': 'No processed data available. Please upload and process new files before searching.'
        }), 200
    
    try:
        # Get query parameters
        employee_id = request.args.get('employee_id', '').strip()
        from_date = request.args.get('from_date', '').strip()
        to_date = request.args.get('to_date', '').strip()
        
        print(f"Search query: Employee_ID='{employee_id}', From_Date='{from_date}', To_Date='{to_date}'")
        
        # Start with all records
        result = current_df.copy()
        
        # Filter by employee ID if provided
        if employee_id:
            result = result[result['Employee_ID'].str.lower() == employee_id.lower()]
        
        # Filter by date range if provided
        if from_date and to_date:
            result = result[(result['Date'] >= from_date) & (result['Date'] <= to_date)]
        elif from_date:
            result = result[result['Date'] >= from_date]
        elif to_date:
            result = result[result['Date'] <= to_date]
        
        # Check if any records found
        if result.empty:
            print(f"No records found for Employee_ID: '{employee_id}', From_Date: '{from_date}', To_Date: '{to_date}'")
            return jsonify({
                'records': [],
                'message': f"No records found for Employee_ID: {employee_id}, From_Date: {from_date}, To_Date: {to_date}"
            })
        
        # Return filtered results
        display_columns = ['Employee_ID', 'Employee_Name', 'Date', 'Check_In', 'Check_Out',
                           'Working_Hours', 'Late_Minutes', 'Status', 'Late_Flag', 'Is_Late']
        
        # Ensure only columns that exist in the dataframe are selected
        actual_display_columns = [col for col in display_columns if col in result.columns]
        
        return jsonify({
            'records': result[actual_display_columns].to_dict(orient='records'),
            'total_records': len(result)
        })
        
    except Exception as e:
        print(f"Error searching records: {e}")
        return jsonify({'error': 'Failed to search records', 'message': str(e)}), 500

# --- MODIFIED: /api/upload ROUTE to accept two files ---
@app.route('/api/upload', methods=['POST'])
def upload_files_and_process():
    print("POST /api/upload endpoint hit")

    # Check if both files are present in the request
    if 'employee_file' not in request.files or 'attendance_file' not in request.files:
        print("Missing one or both file parts ('employee_file' or 'attendance_file')")
        return jsonify({"error": "Both 'Employee Data (Binary)' and 'Attendance Data (.dat/.txt)' files are required."}), 400

    employee_file = request.files['employee_file']
    attendance_file = request.files['attendance_file']

    # Check if files were actually selected (filename is not empty)
    if employee_file.filename == '' or attendance_file.filename == '':
        print("One or both files have no selected filename.")
        return jsonify({"error": "One or both selected files have no filename."}), 400

    # Basic file type validation for attendance file
    if not (attendance_file.filename.lower().endswith('.txt') or attendance_file.filename.lower().endswith('.dat')):
        print(f"Invalid attendance file type: {attendance_file.filename}")
        return jsonify({"error": "Invalid attendance file type. Please upload a .txt or .dat file for attendance."}), 400
    
    # For the binary employee file, we generally don't check extension as it might not have one,
    # or it could be a custom binary format. Trust the user's selection here.

    raw_employee_file_path = None # Initialize to None for cleanup in finally block
    raw_attendance_file_path = None

    try:
        # Secure filenames and save raw files temporarily
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        
        emp_filename = secure_filename(employee_file.filename)
        att_filename = secure_filename(attendance_file.filename)

        raw_employee_file_path = os.path.join(app.config['UPLOAD_FOLDER'], f"emp_raw_{timestamp}_{emp_filename}")
        raw_attendance_file_path = os.path.join(app.config['UPLOAD_FOLDER'], f"att_raw_{timestamp}_{att_filename}")
        
        employee_file.save(raw_employee_file_path)
        attendance_file.save(raw_attendance_file_path)
        print(f"Raw employee file saved to: {raw_employee_file_path}")
        print(f"Raw attendance file saved to: {raw_attendance_file_path}")

        # Define the output path for the *interactive* Excel file
        output_excel_filename = f"interactive_attendance_charts_{timestamp}.xlsx"
        output_excel_path = os.path.join(app.config['UPLOAD_FOLDER'], output_excel_filename)
        
        # Call the biometric processing function from your biometric_processor.py
        success = process_biometric_data_for_excel_dashboard(
            raw_employee_file_path, 
            raw_attendance_file_path, 
            output_excel_path
        )
        
        if success:
            return jsonify({
                "message": "Files uploaded and interactive dashboard created successfully!", 
                "dashboard_file": output_excel_filename,
                "download_url": f"/api/download-latest-dashboard" # Provide a generic download URL
            }), 200
        else:
            # If processing failed, but files were saved, ensure the output Excel is removed if incomplete
            if os.path.exists(output_excel_path):
                os.remove(output_excel_path)
            return jsonify({"error": "Failed to create interactive dashboard. Check server logs for details."}), 500

    except Exception as e:
        print(f"Error during file upload or processing: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({"error": f"An unexpected error occurred during processing: {str(e)}"}), 500
    finally:
        # Ensure raw files are cleaned up even on unexpected errors
        if raw_employee_file_path and os.path.exists(raw_employee_file_path):
            os.remove(raw_employee_file_path)
        if raw_attendance_file_path and os.path.exists(raw_attendance_file_path):
            os.remove(raw_attendance_file_path)

# MODIFIED: Route to allow downloading the *latest* generated interactive Excel file
@app.route('/api/download-latest-dashboard', methods=['GET'])
def download_latest_dashboard():
    latest_excel_path = get_latest_processed_excel_path()
    if latest_excel_path and os.path.exists(latest_excel_path):
        return send_from_directory(app.config['UPLOAD_FOLDER'], os.path.basename(latest_excel_path), as_attachment=True)
    return jsonify({"error": "No interactive Excel dashboard found to download. Please upload files first."}), 404


@app.errorhandler(404)
def not_found(error):
    return jsonify({'error': 'Endpoint not found'}), 404

@app.errorhandler(500)
def internal_error(error):
    return jsonify({'error': 'Internal server error'}), 500



if __name__ == "__main__":
    port = int(os.getenv("PORT", 10000))
    debug_mode = os.getenv("FLASK_ENV", "production") == "development"
    app.run(host="0.0.0.0", port=port, debug=debug_mode)
