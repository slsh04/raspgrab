from flask import Flask, render_template, request, redirect, url_for, flash
import pandas as pd
import os
from datetime import datetime
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = 'your-secret-key-here'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# Ensure upload directory exists
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Allowed extensions
ALLOWED_EXTENSIONS = {'xlsx'}

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def parse_schedule(file_path, surname):
    try:
        # Read the Excel file
        df = pd.read_excel(file_path, header=None)
        
        results = []
        
        # Find all cells containing the surname
        for row_idx, row in df.iterrows():
            for col_idx, cell in enumerate(row):
                if pd.notna(cell) and surname.lower() in str(cell).lower():
                    # Get date from column 1 (index 1) in the same row
                    date_val = df.iloc[row_idx, 1] if len(df.columns) > 1 else None
                    
                    # Get time from column 2 (index 2) in the same row (третий столбец)
                    time_val = df.iloc[row_idx, 2] if len(df.columns) > 2 else None
                    
                    # Format date and time if they exist
                    date_str = ""
                    if pd.notna(date_val):
                        if isinstance(date_val, pd.Timestamp):
                            date_str = date_val.strftime('%Y-%m-%d')
                        else:
                            date_str = str(date_val)
                    
                    time_str = ""
                    if pd.notna(time_val):
                        if isinstance(time_val, pd.Timestamp):
                            time_str = time_val.strftime('%H:%M')
                        else:
                            time_str = str(time_val)
                    
                    # Get the subject/text from the cell
                    subject = str(cell) if pd.notna(cell) else ""
                    
                    results.append({
                        'date': date_str,
                        'time': time_str,
                        'subject': subject
                    })
        
        return results
    except Exception as e:
        raise Exception(f"Error parsing file: {str(e)}")

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Check if the post request has the file part
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        
        file = request.files['file']
        surname = request.form.get('surname', '').strip()
        
        # If user does not select file, browser also submit an empty part without filename
        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)
        
        if not surname:
            flash('Please enter a surname')
            return redirect(request.url)
        
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(file_path)
            
            try:
                results = parse_schedule(file_path, surname)
                os.remove(file_path)  # Remove file after processing
                
                if not results:
                    flash(f'No schedule entries found for {surname}')
                    return redirect(request.url)
                
                return render_template('index.html', results=results, surname=surname)
            except Exception as e:
                os.remove(file_path)  # Remove file if error occurs
                flash(str(e))
                return redirect(request.url)
        else:
            flash('Allowed file type is .xlsx only')
            return redirect(request.url)
    
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)