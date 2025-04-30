from flask import Flask, render_template, request
import pandas as pd
import os
from pathlib import Path

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Update paths to use absolute paths
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'static', 'images')
STUDENT_PHOTOS_FOLDER = os.path.join(BASE_DIR, 'static', 'images', 'Student_Photos')
EXCEL_FILE_PATH = os.path.join(BASE_DIR, 'merged_output.xlsx')

# Create directories if they don't exist
os.makedirs(STUDENT_PHOTOS_FOLDER, exist_ok=True)
def load_excel_data(excel_path):
    try:
        df = pd.read_excel(excel_path)
        
        if 'Register No' not in df.columns:
            for i in range(5):
                df = pd.read_excel(excel_path, header=i)
                if 'Register No' in df.columns:
                    break

        if 'Register No' not in df.columns:
            df = pd.read_excel(excel_path, header=None)
            header_row_idx = None
            
            for idx, row in df.iterrows():
                if 'Register No' in row.values or 'Register No'.lower() in [str(x).lower() for x in row.values if isinstance(x, str)]:
                    header_row_idx = idx
                    break
                
            if header_row_idx is not None:
                df.columns = df.iloc[header_row_idx]
                df = df[header_row_idx+1:]

        df.columns = [str(col).strip() for col in df.columns]


        if 'Register No' in df.columns:
            df['Register No'] = df['Register No'].astype(str)
            

            df = df[df['Register No'].str.strip() != '']
            
            print(f"Successfully loaded data with {len(df)} student records")
            return df
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return pd.DataFrame()

def get_student_details(df, register_no):
    try:

        student_row = df[df['Register No'].astype(str) == str(register_no)]
        
        if student_row.empty:
            print(f"No student found with register number: {register_no}")
            return None

        student_data = {
            'name': student_row['Student Name'].values[0] if 'Student Name' in student_row.columns else 'Unknown',
            'register_no': register_no,
            'semesters': [],
            'events': [],
            'courses': [],
            'event_counts': {
                'PAPER PRESENTATION': 0,
                'HACKATHON': 0,
                'PROJECT EXPO': 0,
                'WORKSHOP': 0
            },
            'prize_count': 0,
            'value_added_courses': [], 
            'industrial_visits': [],
            'GitHub Link':None,
            'LinkedIn Link': None,
            'Portfolio Link': None

        }
        
        

        def safe_get(row, col, default='-'):
            try:
                val = row[col].values[0]
                return default if pd.isna(val) else val
            except:
                return default
            
               # Fix for columns with trailing spaces
        # Find GitHub column
        github_col = None
        for col in df.columns:
            if col.strip() == 'GitHub Link':
                github_col = col
                break
        
        # Find LinkedIn column
        linkedin_col = None
        for col in df.columns:
            if col.strip() == 'LinkedIn Link':
                linkedin_col = col
                break
        
        # Find Portfolio column
        portfolio_col = None
        for col in df.columns:
            if col.strip() == 'Portfolio Link':
                portfolio_col = col
                break
        
        # Get values if columns are found
        if github_col:
            student_data['github'] = safe_get(student_row, github_col)
            print(f"Found GitHub in column '{github_col}': {student_data['github']}")
            
        if linkedin_col:
            student_data['linkedin'] = safe_get(student_row, linkedin_col)
            print(f"Found LinkedIn in column '{linkedin_col}': {student_data['linkedin']}")
            
        if portfolio_col:
            student_data['portfolio'] = safe_get(student_row, portfolio_col)
            print(f"Found Portfolio in column '{portfolio_col}': {student_data['portfolio']}")
        
       
        sem1_subjects = [col for col in df.columns if col.startswith('20') and 
                         any(x in col for x in ['BS101', 'CY101', 'EEC101', 'EN101', 'GE101', 'GE102', 'GE103', 'MA101', 'PH101'])]
        sem2_subjects = [col for col in df.columns if col.startswith('20') and 
                         any(x in col for x in ['BE203', 'BS201', 'CS201', 'CS202', 'EC304', 'EEC201', 'EN201', 'GE201', 'MA201', 'TA201'])]
        sem3_subjects = [col for col in df.columns if col.startswith('20') and 
                         any(x in col for x in ['AD301', 'AD302', 'AD303', 'AD304', 'CS402', 'CS404', 'EEC301', 'IT401', 'IT402', 'MA301','TA101'])]
        sem1_marks = {subj: safe_get(student_row, subj) for subj in sem1_subjects}
        student_data['semesters'].append({
            'name': 'SEMESTER 1',
            'subjects': sem1_subjects,
            'marks': sem1_marks,
            'sgpa': safe_get(student_row, 'SGPA 1', 'N/A'),
        })

        sem2_marks = {subj: safe_get(student_row, subj) for subj in sem2_subjects}
        student_data['semesters'].append({
            'name': 'SEMESTER 2',
            'subjects': sem2_subjects,
            'marks': sem2_marks,
            'sgpa': safe_get(student_row, 'SGPA 2', 'N/A'),
          
        })
        
        sem3_marks = {subj: safe_get(student_row, subj) for subj in sem3_subjects}
        student_data['semesters'].append({
            'name': 'SEMESTER 3',
            'subjects': sem3_subjects,
            'marks': sem3_marks,
            'sgpa': safe_get(student_row, 'SGPA 3', 'N/A'),
           
        })
    
        student_data['cgpa'] = safe_get(student_row, 'TOTAL', 'N/A')
        event_columns = [
            'TECHNICAL EVENT', 'PRIZE/PARTICIPATION',
            'TECHNICAL EVENT.1', 'PRIZE/PARTICIPATION.1',
            'TECHNICAL EVENT.2', 'PRIZE/PARTICIPATION.2',
            'TECHNICAL EVENT.3', 'PRIZE/PARTICIPATION.3',
            'TECHNICAL EVENT.4', 'PRIZE/PARTICIPATION.4',
            'TECHNICAL EVENT.5', 'PRIZE/PARTICIPATION.5',
            'TECHNICAL EVENT.6', 'PRIZE/PARTICIPATION.6',
            'TECHNICAL EVENT.7', 'PRIZE/PARTICIPATION.7'
        ]
        
        for i in range(0, len(event_columns), 2):
            event_type_col = event_columns[i]
            prize_col = event_columns[i + 1]
            
            event_type = safe_get(student_row, event_type_col).strip().upper()
            prize = safe_get(student_row, prize_col)
            
            if event_type != '-':
                standardized_event = " ".join(event_type.split())
                if standardized_event in student_data['event_counts']:
                    student_data['event_counts'][event_type] += 1
                else:
                    student_data['event_counts'][event_type] = 1
            
            if prize== 'FIRST':
                student_data['prize_count'] += 1
            elif prize == 'SECOND':
                student_data['prize_count'] += 1
            elif prize == 'THIRD':
                student_data['prize_count'] += 1
            student_data['events'].append({
                'type': event_type,
                'prize': prize
            })

        course_columns = [
            'NAME OF THE COURSE', 'PASS STATUS', 'SCORE',
            
        ]
        for col in course_columns:
            student_data['courses'].append({
                'column': col,
                'value': safe_get(student_row, col)
            })
        
        vac_columns = ['VAC1', 'VAC2']
        iv_columns = ['IV-1', 'IV-2']


        for vac_col in vac_columns:
            vac_value = safe_get(student_row, vac_col)
            if vac_value != '-':
                student_data['value_added_courses'].append({
                    'name': vac_col,
                    'details': vac_value
                })

        for iv_col in iv_columns:
            iv_value = safe_get(student_row, iv_col)
            if iv_value != '-':
                student_data['industrial_visits'].append({
                    'name': iv_col,
                    'details': iv_value
                })

        course_columns = [
            'NAME OF THE COURSE', 'PASS STATUS', 'SCORE',
            '20MA404', '20AD401', '20CS401', '20CS601', 
            '20IT301', '20MC003', '20TA201', 'No of Fail'
        ]
    
        iat_marks = {}
        iat_subjects = ['20MA404', '20AD401', '20CS401', '20CS601', 
                       '20IT301', '20MC003', '20TA(Tamil)']
        
        for subject in iat_subjects:
            iat_marks[subject] = safe_get(student_row, subject)
        

        student_data['iat_marks'] = iat_marks
        student_data['fail_count'] = safe_get(student_row, 'No of Fail', '0')
        
        return student_data
    except Exception as e:
        print(f"Error processing student data: {e}")
        return None

@app.route('/', methods=['GET', 'POST'])
def index():
    result = None
    image_path = None
    error_message = None
    available_students = []
    
    try:
        df = load_excel_data(EXCEL_FILE_PATH)
        if not df.empty and 'Register No' in df.columns:
            available_students = df['Register No'].astype(str).tolist()
    except Exception as e:
        error_message = f"Error loading Excel file: {str(e)}"
        app.logger.error(f"Excel load error: {str(e)}")
    
    if request.method == 'POST':
        register_no = request.form.get('register_no')
        
        try:
            if df is not None and not df.empty:
                result = get_student_details(df, register_no)
                
                if result:
                    # Check for student photo
                    possible_extensions = ['.jpg', '.jpeg', '.png']
                    for ext in possible_extensions:
                        photo_filename = f"{register_no}{ext}"
                        photo_path = os.path.join(STUDENT_PHOTOS_FOLDER, photo_filename)
                        if os.path.exists(photo_path):
                            image_path = f"images/Student_Photos/{photo_filename}"
                            break
                    
                    if not image_path:
                        image_path = 'images/default.jpg'
                        app.logger.warning(f"No image found for student {register_no}")
        except Exception as e:
            error_message = f"Error processing request: {str(e)}"
            app.logger.error(f"Request processing error: {str(e)}")
    
    return render_template('index.html', 
                         result=result, 
                         image_path=image_path, 
                         error_message=error_message,
                         available_students=available_students)
if __name__ == '__main__':
    app.run(debug=True)