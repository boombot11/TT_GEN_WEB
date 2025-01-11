import os
import time
import zipfile
import win32com.client  # Required to interact with Excel
from werkzeug.utils import secure_filename
from flask import Flask, request, jsonify, send_file
import pythoncom
from flask_cors import CORS  # Import CORS

app = Flask(__name__)
CORS(app)
# Configuring the upload folder and allowed extensions
app.config['UPLOAD_FOLDER'] = './uploads'
app.config['ALLOWED_EXTENSIONS'] = {'xlsm'}

# Function to check allowed extensions
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

# Function to run dynamic Excel macro using win32com.client
def execute_dynamic_excel_macro(file_path, macro_name, user_input_lab, user_input_lecture):
    try:
        # Debugging log
        print(f"Running macro '{macro_name}' on {file_path}")
        print(f"Initializing Excel application...")
        pythoncom.CoInitialize()
        # Create an Excel application instance
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False  # Excel will run in the background
        excel.DisplayAlerts = False  # Disable any Excel alerts
        file_path = os.path.abspath(file_path)
        # Open the workbook
        print(f"Opening workbook: {file_path}")
        workbook = excel.Workbooks.Open(file_path)

        # Run the macro with dynamic parameters
        print(f"Running macro: {macro_name}")
        excel.Application.Run(macro_name, user_input_lab, user_input_lecture)

        # Save and close the workbook
        print("Saving and closing the workbook...")
        workbook.Save()
        workbook.Close(False)  # Close without saving changes again
        excel.Quit()  # Quit Excel

        # Adding a 2-second delay to let changes sink in
        time.sleep(2)

        print("Macro executed successfully.")
        return None  # No error
    except Exception as e:
        return str(e)  # Return the error message if something fails

def extract_sheets(excel_file_path, rooms, labs, new_room_file_path, new_lab_file_path, new_teacher_file_path):
    try:
        # Initialize COM for thread safety
        pythoncom.CoInitialize()
       
        # Ensure the input Excel file exists
        if not os.path.exists(excel_file_path):
            print(f"Error: The specified Excel file does not exist: {excel_file_path}")
            return

        # Create an Excel application instance
        print(f"Opening Excel file: {excel_file_path}")
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False  # Excel will run in the background
        excel.DisplayAlerts = False  # Disable any Excel alerts

        # Open the original workbook (.xlsm)
        workbook = excel.Workbooks.Open(excel_file_path)

        # Check if the output workbooks already exist and open them, otherwise create new workbooks
        room_workbook = open_or_create_workbook(excel, new_room_file_path, 'Room')
        lab_workbook = open_or_create_workbook(excel, new_lab_file_path, 'Lab')
        teacher_workbook = open_or_create_workbook(excel, new_teacher_file_path, 'Teacher')

        # Function to check if a sheet name consists of only uppercase letters (indicating teacher sheets)
        def is_teacher_sheet(sheet_name):
            return sheet_name.isupper()  # Only uppercase letters are considered teacher sheets

        # Loop through each sheet in the workbook
        for sheet in workbook.Sheets:
            sheet_name = sheet.Name
            print(f"Processing sheet: {sheet_name}")

            # Check if the sheet is a room sheet
            if sheet_name in rooms:
                print(f"Copying room sheet: {sheet_name}")
                sheet.Copy()  # Copy the sheet to a new workbook
                copied_sheet = excel.ActiveSheet
                print(f"Room sheet copied: {copied_sheet.Name}")
                # Clear existing sheets and move the copied sheet to the room workbook
                copied_sheet.Move(room_workbook.Sheets.Item(room_workbook.Sheets.Count))

            # Check if the sheet is a lab sheet
            elif sheet_name in labs:
                print(f"Copying lab sheet: {sheet_name}")
                sheet.Copy()  # Copy the sheet to a new workbook
                copied_sheet = excel.ActiveSheet
                print(f"Lab sheet copied: {copied_sheet.Name}")
                # Clear existing sheets and move the copied sheet to the lab workbook
                copied_sheet.Move(lab_workbook.Sheets.Item(lab_workbook.Sheets.Count))

            # Check if the sheet name is for a teacher
            elif is_teacher_sheet(sheet_name):
                print(f"Copying teacher sheet: {sheet_name}")
                sheet.Copy()  # Copy the sheet to a new workbook
                copied_sheet = excel.ActiveSheet
                print(f"Teacher sheet copied: {copied_sheet.Name}")
                # Clear existing sheets and move the copied sheet to the teacher workbook
                copied_sheet.Move(teacher_workbook.Sheets.Item(teacher_workbook.Sheets.Count))

        # Save the updated workbooks
        print("Saving updated workbooks...")
        save_workbook(room_workbook, new_room_file_path)
        save_workbook(lab_workbook, new_lab_file_path)
        save_workbook(teacher_workbook, new_teacher_file_path)

        # Close the workbooks
        print("Closing workbooks...")
        close_workbook(room_workbook)
        close_workbook(lab_workbook)
        close_workbook(teacher_workbook)
        close_workbook(workbook)

        # Close the Excel application
        excel.Quit()

        # Release the COM objects to avoid memory leaks
        release_com_objects([workbook, room_workbook, lab_workbook, teacher_workbook, excel])

        # Add a 2-second delay to ensure changes are applied
        time.sleep(2)

        print("Sheets extracted successfully in .xlsx format!")

    except Exception as e:
        print(f"Error during sheet extraction: {str(e)}")


def open_or_create_workbook(excel, file_path, workbook_type):
    if os.path.exists(file_path):
        print(f"Opened existing {workbook_type} workbook: {file_path}")
        return excel.Workbooks.Open(file_path)
    else:
        print(f"Created new {workbook_type} workbook: {file_path}")
        return excel.Workbooks.Add()


def save_workbook(workbook, file_path):
    workbook.SaveAs(file_path, 51)  # 51 corresponds to .xlsx format
    print(f"Saved workbook at: {file_path}")


def close_workbook(workbook):
    workbook.Close()
    print(f"Workbook closed.")


def release_com_objects(objects):
    for obj in objects:
        try:
            del obj
        except Exception as e:
            print(f"Error releasing COM object: {e}")

# Function to create a zip file from a list of files
def create_zip(files, zip_path):
    try:
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for file in files:
                zipf.write(file, os.path.basename(file))  # Add file to zip with basename only
        print(f"Created zip file: {zip_path}")
        return zip_path
    except Exception as e:
        return str(e)

# Main route to upload and process the Excel file
@app.route('/upload-excel', methods=['POST'])
def upload_excel_complex():
    try:
        file = request.files.get('file')
        if not file:
            return jsonify({"error": "No file uploaded"}), 400
        if not allowed_file(file.filename):
            return jsonify({"error": "Only .xlsm files are allowed"}), 400

        classrooms = request.form.get('classrooms', '')
        labs = request.form.get('labs', '')

        # Validate inputs
        if not classrooms or not labs:
            return jsonify({"error": "Classrooms and labs are required"}), 400

        # Save the uploaded file
        filename = f"{int(time.time())}-{secure_filename(file.filename)}"
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)

        # Normalize and get the absolute path of the file
        file_path = os.path.normpath(file_path)
        absolute_file_path = os.path.abspath(file_path)
        print(f"File saved at: {absolute_file_path}")

        macro_name = 'RunAllModules'  # Replace with the actual macro name you want to run
        user_input_lab = labs
        user_input_lecture = classrooms

        # Execute the dynamic macro
        error = execute_dynamic_excel_macro(file_path, macro_name, user_input_lab, user_input_lecture)
        if error:
            return jsonify({"error": "Failed to execute macro", "details": error}), 500

        # Prepare file paths for extracted sheets
        new_room_file_path = os.path.join(app.config['UPLOAD_FOLDER'], f"room.xlsx")
        new_lab_file_path = os.path.join(app.config['UPLOAD_FOLDER'], f"lab.xlsx")
        new_teacher_file_path = os.path.join(app.config['UPLOAD_FOLDER'], f"teachers.xlsx")

        # Normalize paths for new files
        new_room_file_path = os.path.normpath(new_room_file_path)
        new_lab_file_path = os.path.normpath(new_lab_file_path)
        new_teacher_file_path = os.path.normpath(new_teacher_file_path)

        # Debugging log
        print(f"New Room File Path: {new_room_file_path}")
        print(f"New Lab File Path: {new_lab_file_path}")
        print(f"New Teacher File Path: {new_teacher_file_path}")
        new_room_file_path = os.path.abspath(new_room_file_path)
        new_lab_file_path = os.path.abspath(new_lab_file_path)
        new_teacher_file_path = os.path.abspath(new_teacher_file_path)
        
        # Extract the sheets into new files
        error = extract_sheets(absolute_file_path, classrooms.split(), labs.split(), new_room_file_path, new_lab_file_path, new_teacher_file_path)
        if error:
            return jsonify({"error": "Failed to extract sheets", "details": error}), 500

        # Create a zip file containing the extracted sheets
        zip_filename = f"extracted_{filename}.zip"
        zip_path = os.path.join(app.config['UPLOAD_FOLDER'], zip_filename)
        zip_path = os.path.abspath(zip_path)
        create_zip([new_room_file_path, new_lab_file_path, new_teacher_file_path], zip_path)
        response = send_file(zip_path, as_attachment=True, download_name=zip_filename, mimetype='application/zip')

        # Cleanup files after sending the response
        try:
            # Delete the temporary files (Excel sheets and zip file) after response is sent
            os.remove(zip_path)  # Remove the zip file
            os.remove(file_path)  # Remove the original uploaded file
            print(f"Temporary files deleted successfully.")
        except Exception as e:
            print(f"Error during file cleanup: {e}")


        # Return the zip file to the client
        return response;
        
    except Exception as e:
        print(f"Error processing file: {e}")
        return jsonify({"error": "An error occurred while processing the file", "details": str(e)}), 500

def execute_static_excel_macro(file_path, macro_name):
    try:
        # Debugging log - Starting macro execution
        print(f"Attempting to run macro '{macro_name}' on file: {file_path}")
        print("Initializing Excel application...")

        # Normalize and get absolute path
        file_path = os.path.abspath(file_path)
        print(f"Normalized absolute path: {file_path}")  # Debugging the absolute path

        # Initialize COM library (important for multithreaded environments)
        pythoncom.CoInitialize()

        # Create an Excel application instance
        print("Creating Excel application instance...")
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False  # Excel will run in the background
        excel.DisplayAlerts = False  # Disable any Excel alerts

        # Debugging: Excel is created and ready
        print("Excel application created successfully.")
        
        # Open the workbook using absolute path
        print(f"Attempting to open workbook: {file_path}")
        workbook = excel.Workbooks.Open(file_path)
        
        # Debugging: Workbook opened successfully
        print(f"Workbook opened: {workbook.Name}")

        # Run the macro
        print(f"Running macro: {macro_name}...")
        excel.Application.Run(macro_name)

        # Debugging: After macro execution
        print(f"Macro '{macro_name}' executed successfully.")
        
        # Save and close the workbook
        print(f"Saving workbook: {workbook.Name}")
        workbook.Save()
        print(f"Closing workbook: {workbook.Name}")
        workbook.Close(False)  # Close without saving changes again
        excel.Quit()  # Quit Excel

        # Adding a 2-second delay to let changes sink in
        print("Waiting for 2 seconds to ensure changes are applied...")
        time.sleep(2)

        print("Macro executed successfully and workbook closed.")
        return None  # No error
    except Exception as e:
        # Log detailed error message
        print(f"Error during macro execution: {str(e)}")
        return str(e)  # Return the error message if something fails

# Static Excel upload handler
@app.route('/upload/Static', methods=['POST'])
def upload_excel_static():
    try:
        file = request.files.get('file')
        if not file:
            return jsonify({"error": "No file uploaded"}), 400
        if not allowed_file(file.filename):
            return jsonify({"error": "Only .xlsm files are allowed"}), 400

        # Save the uploaded file
        filename = f"{int(time.time())}-{secure_filename(file.filename)}"
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)

        # Normalize and get the absolute path of the file
        file_path = os.path.normpath(file_path)
        absolute_file_path = os.path.abspath(file_path)
        print(f"File saved at: {absolute_file_path}")

        macro_name = 'RunAllModules'  # Replace with the actual macro name you want to run

        # Execute the static macro
        error = execute_static_excel_macro(file_path, macro_name)
        if error:
            return jsonify({"error": "Failed to execute macro", "details": error}), 500

        # Return the processed file to the client
        return send_file(file_path, as_attachment=True, download_name=filename, mimetype='application/vnd.ms-excel.sheet.macroenabled.12')

    except Exception as e:
        print(f"Error processing file: {e}")
        return jsonify({"error": "An error occurred while processing the file", "details": str(e)}), 500


if __name__ == '__main__':
    # Make sure the upload folder exists
    if not os.path.exists(app.config['UPLOAD_FOLDER']):
        os.makedirs(app.config['UPLOAD_FOLDER'])

    # Run the Flask app
    app.run(debug=True)
