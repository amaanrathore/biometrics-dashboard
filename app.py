import os
from flask import Flask, request, jsonify
from flask_cors import CORS
import pandas as pd
from gunicorn.app.base import BaseApplication

app = Flask(__name__)

# Refined CORS configuration for production
CORS(app, resources={
    r"/api/*": {
        "origins": [
            "https://biometrics-dashboard-k2lix7izv-amaan-rathores-projects.vercel.app"
        ],
        "methods": ["GET", "POST", "PUT", "DELETE"],
        "allow_headers": ["Content-Type", "Authorization"]
    }
})

def load_biometric_data(file_path=None):
    """Load biometric data from Excel file"""
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
        print("Columns:", df.columns.tolist())
        
        # Skip first row and reset index
        df = df.iloc[1:].reset_index(drop=True)
        
        # Define expected columns (adjusted to match actual data)
        expected_columns = [
            'Employee_ID', 'Employee_Name', 'Date', 'Check_In', 'Check_Out',
            'Working_Hours', 'Late_Minutes', 'Status', 'Late_Flag', 'Is_Late'
        ]
        
        # Handle column naming
        if len(df.columns) != len(expected_columns):
            print(f"Warning: Expected {len(expected_columns)} columns, found {len(df.columns)}")
            df.columns = expected_columns[:min(len(df.columns), len(expected_columns))] + \
                         [f'Extra_{i}' for i in range(len(df.columns) - len(expected_columns))]
        
        # Process data
        df['Date'] = pd.to_datetime(df['Date'], format='%Y-%m-%d', errors='coerce').dt.strftime('%Y-%m-%d')
        df['Employee_ID'] = df['Employee_ID'].astype(str).str.strip()
        df = df.fillna('N/A').infer_objects(copy=False)
        
        print("Processed Data (first 5 rows):")
        print(df.head().to_string())
        return df
        
    except Exception as e:
        print(f"Error loading Excel file: {e}")
        return None

# Initialize dataframe
df = load_biometric_data()

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

class StandaloneApplication(BaseApplication):
    def __init__(self, app, options=None):
        self.application = app
        super().__init__()

    def load_config(self):
        config = {key: value for key, value in self.cfg.items() if key in self.cfg}
        config.setdefault('bind', '0.0.0.0:' + str(os.getenv("PORT", "10000")))
        return config

    def load(self):
        return self.application

if __name__ == "__main__":
    options = {
        'bind': '0.0.0.0:' + str(os.getenv("PORT", "10000")),
        'workers': 1,
    }
    StandaloneApplication(app, options).run()
