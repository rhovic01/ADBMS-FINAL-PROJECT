# Professor Dashboard Application

## Overview

The Professor Dashboard is a Windows Forms application designed to help professors manage student grades and information. It provides functionality for importing grade sheets from Excel, managing individual student records, and maintaining a database of student information.

## Features

### 1. Excel Import Functionality

- Import grade sheets directly from Excel files
- Automatically reads data starting from row 8 in Excel
- Supports importing:
  - Student ID
  - Student Name
  - GPA
  - Remarks
- Skips header rows automatically
- Preview data before importing

### 2. Student Record Management

- **Add New Records**

  - Input individual student information
  - Save records to database
  - Automatic validation of data

- **Update Records**

  - Modify existing student information
  - Update GPA and remarks
  - Real-time database updates

- **Delete Records**
  - Remove student records from database
  - Confirmation prompt before deletion

### 3. Batch Operations

- Import multiple student records at once
- Preview before saving
- Transaction support for data integrity
- Error handling and reporting

### 4. Data Fields

The application manages the following information:

- School Year
- Course
- Section
- Subject Code
- Subject Title
- Semester
- Instructor
- Grade ID
- Student ID
- Student Name
- GPA
- Remarks

## Database Structure

The application uses an Access database (register.accdb) with the following table:

**Table: tbl_professor**

- SchoolYear
- Course
- Section
- SubjectCode
- SubjectTitle
- Semester
- Instructor
- GradeID
- StudentID
- StudentName
- GPA
- Remarks

## Key Functions

### ImportFromExcel

- Reads Excel files
- Validates data format
- Previews data in ListView
- Supports batch importing

### SaveProfessorData

- Saves individual student records
- Validates required fields
- Handles database transactions
- Error reporting

### UpdateProfessorData

- Updates existing records
- Maintains data integrity
- Real-time database updates

### DeleteProfessorData

- Removes records safely
- Confirmation before deletion
- Database integrity checks

## Requirements

- Microsoft Access Database Engine (ACE)
- Microsoft Office Interop (for Excel operations)
- .NET Framework
- Windows Operating System

## Error Handling

- Duplicate record checking
- Data validation
- Transaction rollback on errors
- User-friendly error messages

## Best Practices

1. Always preview data before importing
2. Verify student information before saving
3. Use batch import for multiple records
4. Regular database backups recommended

## Notes

- Excel files should follow the specified format
- Student IDs must be unique
- GPA should be in valid numerical format
- Remarks are optional but supported
