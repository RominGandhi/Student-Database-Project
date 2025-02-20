# Student Marking Application Manual

## 1. Accessing the Application

### 1.1 Locate and Open the Worksheet
- Navigate to the `gand6363_Database` folder.
- Open the Microsoft Excel Macro-Enabled Worksheet (`.xlsm`) named `gand6363_FinalDatabase`.

### 1.2 Select the Database
- Go to the "Student Database" Ribbon.
- Click the "Browse" button to open a file dialog.
- Locate and select your desired student database file, then click "Open." Alternatively, you can use the provided file called `Registrat.mdb` to test application functionality.
- Click "Run." If successful, a success message will appear. If not, verify the correct file is selected.

## 2. Using the Options Userform
Once you've successfully connected to the database, an "Options" Userform will appear with the following features:

### 2.1 Course List
- **View Courses**: Displays all courses offered at Wilfrid Laurier University’s Waterloo Campus.
- **Calculate Statistics**: Click on a course, then use the provided buttons to compute the average and standard deviation for that course.

### 2.2 Student List
- **View Student Information**: Displays a list of all students, including their first name, last name, and student ID, in a new worksheet.

### 2.3 Student Search
- **Search by Student ID**: Enter the student ID in the "Student Search" Userform to find specific student details.
- **Error Handling**: Errors will occur if the student ID is invalid or the field is left empty.

### 2.4 Generate Report

#### Course-Specific Report
- Select a course and click "Generate Report."
- The report will be generated in Word with statistical data available in Excel, named in the format `CourseInfo_<CourseName>`.

#### Student-Specific Report
- Enter the student ID and click "Generate Word Report."
- The report will be generated in Word with statistical data available in Excel, named in the format `StudentInfo_<StudentID>`.

## 3. Troubleshooting Common Errors
- **No File Selected**: Ensure a file is selected before clicking "Run."
- **Wrong File Type**: Confirm that the selected file is compatible with the application.
