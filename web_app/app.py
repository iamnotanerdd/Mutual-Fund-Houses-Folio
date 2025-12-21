from flask import Flask, render_template, jsonify
import pandas as pd
import os
import math

app = Flask(__name__)

# Determine the base directory (root of the project)
BASE_DIR = os.path.dirname(os.path.abspath(__file__)) # This is web_app/
ROOT_DIR = os.path.dirname(BASE_DIR) # Parent is root
EXCEL_PATH = os.path.join(ROOT_DIR, "PPFAS_Equity_Analysis.xlsx")

def read_portfolio_data():
    if not os.path.exists(EXCEL_PATH):
        return []

    # Read data skipping top rows (Title=1, Month=2, SubHeader=3) -> Header at Row 3 (index 2)?
    # Actually, let's read with header=2 (Row 3, 0-indexed) which contains Qty, Value, %
    # But that loses the Month info from Row 2.
    
    # Better approach: Read strictly data rows (starting Row 4) and map manually using column knowledge.
    # We know the structure:
    # Cols 0,1,2: Name, ISIN, Rating
    # Cols 3,4,5: Month 1 (Qty, Val, Pct)
    # ...
    
    df = pd.read_excel(EXCEL_PATH, header=None, skiprows=3) # Skip Title(1), Months(2), SubHeaders(3). Data starts Row 4.
    
    # We need the Month names to label the keys.
    # Read Row 2 specifically.
    df_months = pd.read_excel(EXCEL_PATH, header=None, nrows=1, skiprows=1)
    month_row = df_months.iloc[0]
    
    # Extract unique ordered months from the merged cells row
    months = []
    for i in range(3, len(month_row), 3):
        val = month_row[i]
        if pd.notna(val):
            months.append(val)
            
    data = []
    
    for _, row in df.iterrows():
        # Stop if we hit empty name, but allow "Total"
        name = str(row[0])
        if pd.isna(row[0]) or name.lower() == 'nan':
            continue
            
        # Sanitize metadata fields to avoid NaN JSON errors
        def clean_meta(val):
            if pd.isna(val) or str(val).lower() == 'nan': return ""
            return str(val)

        record = {
            "Name": row[0], # Name is already checked above
            "ISIN": clean_meta(row[1]),
            "Rating": clean_meta(row[2]),
            "Months": {}
        }
        
        col_idx = 3
        for m in months:
            # Handle NaNs and floats
            def clean(val):
                if pd.isna(val) or val == "": 
                    return 0.0
                try:
                    return float(val)
                except:
                    return 0.0
                
            qty = clean(row[col_idx])
            val = clean(row[col_idx+1])
            pct = clean(row[col_idx+2])
            
            record["Months"][m] = {
                "Quantity": qty,
                "Value": val,
                "Pct": pct
            }
            col_idx += 3
            
        data.append(record)
        
    return {"months": months, "records": data}

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/data')
def get_data():
    try:
        data = read_portfolio_data()
        return jsonify(data)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5000)
