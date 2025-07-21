import os
from flask import Flask, request, jsonify
from flask_cors import CORS
import pandas as pd

app = Flask(__name__)

# Refined CORS configuration for production and local development
CORS(app, resources={
    r"/api/*": {
        "origins": [
            "http://localhost:3000",  # For local development
            "https://biometrics-dashboard-iq8y63fym-amaan-rathores-projects.vercel.app/"  # For production
        ],
        "methods": ["GET", "POST", "PUT", "DELETE"],
        "allow_headers": ["Content-Type", "Authorization"]
    }
})

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
            df.columns = list(df.columns) + [f'Extra_{i}' for i in range(len(expected_columns) - len(df.columns))]
        
        print("Columns after renaming:", df.columns.tolist())
        print("Sample Date values before conversion:", df['Date'].head().tolist())
        
        # Enhanced date handling
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
        print("Date values after conversion:", df['Date'].head().tolist())
        df['Date'] = df['Date'].fillna('N/A')
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
    if df is None or df.empty:
        print("No data available in DataFrame")
        return jsonify({'error': 'No data available'}), 500
    
    try:
        # Get query parameters
        employee_id = request.args.get('employee_id', '').strip()
        from_date = request.args.get('from_date', '').strip()
        to_date = request.args.get('to_date', '').strip()
        
        print(f"Search query: Employee_ID='{employee_id}', From_Date='{from_date}', To_Date='{to_date}'")
        print(f"Available Employee_IDs: {df['Employee_ID'].unique().tolist()}")
        print(f"Available Dates: {df['Date'].unique().tolist()}")
        
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

# Conditional Gunicorn import and execution for deployment
if 'GUNICORN' in os.environ:  # Check if running under Gunicorn (e.g., on Render)
    from gunicorn.app.base import BaseApplication

    class StandaloneApplication(BaseApplication):
        def __init__(self, app, options=None):
            self.application = app
            super().__init__()

        def load_config(self):
            config = {}
            config['bind'] = f'0.0.0.0:{os.getenv("PORT", "10000")}'
            config['workers'] = 1
            return config

        def load(self):
            return self.application

    if __name__ == "__main__":
        options = {
            'bind': f'0.0.0.0:{os.getenv("PORT", "10000")}',
            'workers': 1,
        }
        StandaloneApplication(app, options).run()
else:
    # Use Flask development server locally (Windows-compatible)
    if __name__ == "__main__":
        port = int(os.getenv("PORT", 10000))
        app.run(host="0.0.0.0", port=port, debug=True)
