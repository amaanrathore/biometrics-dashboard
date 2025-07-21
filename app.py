from flask import Flask, request, jsonify
from flask_cors import CORS
import pandas as pd
import os

app = Flask(__name__)


CORS(app, resources={
    r"/api/*": {
        "origins": [
            "http://localhost:3000",
            "http://127.0.0.1:3000",
            # Add your exact Vercel production domain here when you have one:
            "https://biometrics-dashboard.vercel.app",
            # Add the exact preview domain that was blocked:
            "https://biometrics-dashboard-1q8y63fym-amaan-rathores-projects.vercel.app",
            # If you have other specific Vercel preview domains, list them.
            # For a truly dynamic wildcard, Flask-CORS might need a regex,
            # but listing the exact ones is more reliable for now.
        ],
        "methods": ["GET", "POST", "PUT", "DELETE", "OPTIONS"],
        "allow_headers": ["Content-Type", "Authorization", "Accept", "Origin", "X-Requested-With"],
        "supports_credentials": False # Set to True only if you are sending cookies/auth tokens
    }
})

# --- End of CORS Configuration ---


def load_biometric_data(file_path=None):
    """Load biometric data from Excel file with enhanced debugging"""
    print("Attempting to load biometric data from:", file_path)
    try:
        # Use relative path for deployment
        if file_path is None:
            file_path = os.path.join(os.path.dirname(__file__), 'data_biometrics.xlsx')
        
        # Check if file exists
        if not os.path.exists(file_path):
            print(f"Error: File not found at {file_path}")
            return None
            
        df = pd.read_excel(file_path, header=1)
        print("Loaded Data (first 5 rows):")
        print(df.head().to_string())
        print("Columns before renaming:", df.columns.tolist())
        
        # Skip first row and reset index
        df = df.iloc[1:].reset_index(drop=True)
        
        # Define expected columns
        expected_columns = [
            'Employee_ID', 'Employee_Name', 'Date', 'Check_In', 'Check_Out',
            'Working_Hours', 'Late_Minutes', 'Status', 'Late_Flag', 'Is_Late'
        ]
        
        # Always apply expected columns, adjusting for actual column count
        df.columns = expected_columns[:len(df.columns)]
        if len(df.columns) < len(expected_columns):
            # Pad with 'Extra_X' if fewer columns than expected
            df.columns = list(df.columns) + [f'Extra_{i}' for i in range(len(df.columns), len(expected_columns))]
        elif len(df.columns) > len(expected_columns):
            # Truncate if more columns than expected
            df = df.iloc[:, :len(expected_columns)]
            df.columns = expected_columns

        print("Columns after renaming:", df.columns.tolist())
        print("Sample Date values before conversion:", df['Date'].head().tolist())
        
        # Enhanced date handling
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
        print("Date values after conversion:", df['Date'].head().tolist())
        df['Date'] = df['Date'].dt.strftime('%Y-%m-%d').fillna('N/A') # Convert to string after datetime, then fill N/A
        df['Employee_ID'] = df['Employee_ID'].astype(str).str.strip()
        df = df.fillna('N/A').infer_objects(copy=False)
        
        print("Processed Data (first 5 rows):")
        print(df.head().to_string())
        return df
        
    except Exception as e:
        print(f"Error loading Excel file: {e}")
        return None

# Initialize dataframe with debugging
print("Initializing dataframe...")
df = load_biometric_data()
print(f"Dataframe initialized. Data available: {df is not None and not df.empty}")

@app.route('/')
def home():
    """Home route with API information"""
    return jsonify({
        'message': 'Welcome to the Biometric Search API',
        'endpoints': {
            '/api/employees': 'GET - Get all employees',
            '/api/search': 'GET - Search records with query parameters'
        },
        'status': 'active'
    })

@app.route('/api/health', methods=['GET'])
def health_check():
    """Health check endpoint"""
    return jsonify({
        'status': 'healthy',
        'data_loaded': df is not None and not df.empty,
        'total_records': len(df) if df is not None else 0
    })

@app.route('/api/employees', methods=['GET'])
def get_employees():
    """Get all unique employees"""
    print("GET /api/employees endpoint hit")
    if df is None or df.empty:
        print("No data available in DataFrame")
        return jsonify({'error': 'No data available'}), 500
    
    try:
        employees = df[['Employee_ID', 'Employee_Name']].drop_duplicates().to_dict(orient='records')
        print(f"Returning {len(employees)} unique employees")
        return jsonify({'employees': employees})
    except Exception as e:
        print(f"Error getting employees: {e}")
        return jsonify({'error': 'Failed to retrieve employees'}), 500

@app.route('/api/search', methods=['GET'])
def search_records():
    """Search records based on employee ID and date range"""
    print("GET /api/search endpoint hit")
    if df is None or df.empty:
        print("No data available in DataFrame")
        return jsonify({'error': 'No data available'}), 500
    
    try:
        # Get query parameters
        employee_id = request.args.get('employee_id', '').strip()
        from_date = request.args.get('from_date', '').strip()
        to_date = request.args.get('to_date', '').strip()
        
        print(f"Search query: Employee_ID='{employee_id}', From_Date='{from_date}', To_Date='{to_date}'")
        print(f"Available Employee_IDs in DF: {df['Employee_ID'].unique().tolist()}")
        print(f"Available Dates in DF: {df['Date'].unique().tolist()}")
        
        # Start with all records
        result = df.copy()
        
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
        
        return jsonify({
            'records': result[display_columns].to_dict(orient='records'),
            'total_records': len(result)
        })
        
    except Exception as e:
        print(f"Error searching records: {e}")
        return jsonify({'error': 'Failed to search records'}), 500

@app.errorhandler(404)
def not_found(error):
    return jsonify({'error': 'Endpoint not found'}), 404

@app.errorhandler(500)
def internal_error(error):
    return jsonify({'error': 'Internal server error'}), 500


if __name__ == "__main__":
    port = int(os.getenv("PORT", 10000))

    debug_mode = os.getenv("FLASK_ENV", "production") == "development" # This will be False on Render
    app.run(host="0.0.0.0", port=port, debug=debug_mode)
