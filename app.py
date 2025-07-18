from flask import Flask, request, jsonify
from flask_cors import CORS
import pandas as pd

app = Flask(__name__)
CORS(app, resources={r"/api/*": {"origins": "http://localhost:3000"}})

def load_biometric_data(file_path='C:/Users/bhara/Desktop/biometrics app/backend/data_biometrics.xlsx'):
    try:
        df = pd.read_excel(file_path, header=1)
        print("Loaded Data (first 5 rows):")
        print(df.head().to_string())
        print("Columns:", df.columns.tolist())
        df = df.iloc[1:].reset_index(drop=True)
        expected_columns = [
            'Employee_ID', 'Employee_Name', 'Date', 'Check_In', 'Check_Out',
            'Working_Hours', 'Late_Minutes', 'Status', 'Late_Flag', 'Is_Late',
            'Unused'
        ]
        if len(df.columns) == 13:
            df.columns = expected_columns
        else:
            print(f"Warning: Expected 13 columns, found {len(df.columns)}")
            df.columns = expected_columns[:min(len(df.columns), len(expected_columns))] + \
                         [f'Extra_{i}' for i in range(len(df.columns) - len(expected_columns))]
        df['Date'] = pd.to_datetime(df['Date'], format='%Y-%m-%d', errors='coerce').dt.strftime('%Y-%m-%d')
        df['Employee_ID'] = df['Employee_ID'].astype(str).str.strip()
        df = df.fillna('N/A').infer_objects(copy=False)
        print("Processed Data (first 5 rows):")
        print(df.head().to_string())
        return df
    except Exception as e:
        print(f"Error loading Excel file: {e}")
        return None

df = load_biometric_data()

@app.route('/')
def home():
    return jsonify({'message': 'Welcome to the Biometric Search API. Use /api/search or /api/employees.'})

@app.route('/api/employees', methods=['GET'])
def get_employees():
    if df is None or df.empty:
        print("No data available in DataFrame")
        return jsonify({'error': 'No data available'}), 500
    employees = df[['Employee_ID', 'Employee_Name']].drop_duplicates().to_dict(orient='records')
    print(f"Returning {len(employees)} unique employees")
    return jsonify({'employees': employees})

@app.route('/api/search', methods=['GET'])
def search_records():
    if df is None or df.empty:
        print("No data available in DataFrame")
        return jsonify({'error': 'No data available'}), 500
    employee_id = request.args.get('employee_id', '').strip()
    from_date = request.args.get('from_date', '').strip()
    to_date = request.args.get('to_date', '').strip()
    print(f"Search query: Employee_ID='{employee_id}', From_Date='{from_date}', To_Date='{to_date}'")
    print(f"Available Employee_IDs: {df['Employee_ID'].unique().tolist()}")
    print(f"Available Dates: {df['Date'].unique().tolist()}")
    
    result = df
    if employee_id:
        result = result[result['Employee_ID'].str.lower() == employee_id.lower()]
    if from_date and to_date:
        result = result[(result['Date'] >= from_date) & (result['Date'] <= to_date)]
    elif from_date:
        result = result[result['Date'] >= from_date]
    elif to_date:
        result = result[result['Date'] <= to_date]
    
    if result.empty:
        print(f"No records found for Employee_ID: '{employee_id}', From_Date: '{from_date}', To_Date: '{to_date}'")
        return jsonify({'records': [], 'message': f"No records found for Employee_ID: {employee_id}, From_Date: {from_date}, To_Date: {to_date}"})
    display_columns = ['Employee_ID', 'Employee_Name', 'Date', 'Check_In', 'Check_Out',
                      'Working_Hours', 'Late_Minutes', 'Status', 'Late_Flag', 'Is_Late']
    return jsonify({'records': result[display_columns].to_dict(orient='records')})

if __name__ == '__main__':
    app.run(debug=True, port=5000)