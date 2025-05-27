Imports Microsoft.Office.Interop
Imports System.Data.OleDb
Imports System.IO

Public Class PROFESSOR_DASHBOARD
    ' Access database connection string - update path to your register.accdb file
    Private connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=register.accdb;Persist Security Info=False;"
    Private studentsToImport As New List(Of Student)()

    Private Class Student
        Public Property StudentID As String
        Public Property Name As String
        Public Property GPA As String
        Public Property Remarks As String
    End Class

    Private Sub PROFESSOR_DASHBOARD_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Initialize form components
        LoadProfessorData()

        ' Set up OpenFileDialog for Excel import
        OpenFileDialog1.Filter = "Excel Files|*.xlsx;*.xls|All Files|*.*"
        OpenFileDialog1.DefaultExt = "xlsx"
        OpenFileDialog1.Title = "Import Grade Sheet"
    End Sub

    Private Sub LoadProfessorData()
        ' Load existing professor data into form or grid
        Try
            Using connection As New OleDbConnection(connectionString)
                connection.Open()
                Dim query As String = "SELECT * FROM tbl_professor ORDER BY StudentID DESC" ' Changed from DateCreated
                Dim adapter As New OleDbDataAdapter(query, connection)
                Dim dataTable As New DataTable()
                adapter.Fill(dataTable)

                ' If you have a DataGridView, uncomment the next line
                ' dgvProfessor.DataSource = dataTable
            End Using
        Catch ex As Exception
            MessageBox.Show("Error loading professor data: " & ex.Message)
        End Try
    End Sub

    Private Sub OpenFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk
        ImportFromExcel(OpenFileDialog1.FileName)
    End Sub

    Private Sub ImportFromExcel(filePath As String)
        Dim xlApp As Excel.Application = Nothing
        Dim xlWorkbook As Excel.Workbook = Nothing
        Dim xlWorksheet As Excel.Worksheet = Nothing

        Try
            ' Clear previous data
            lvprofessor.Items.Clear()
            studentsToImport.Clear()

            ' Create Excel application
            xlApp = New Excel.Application()
            xlApp.Visible = False
            xlApp.DisplayAlerts = False

            ' Open workbook and worksheet
            xlWorkbook = xlApp.Workbooks.Open(filePath)
            xlWorksheet = CType(xlWorkbook.Sheets(1), Excel.Worksheet)

            ' Read header information
            ReadHeaderInformation(xlWorksheet)

            ' Load student data into ListView
            LoadStudentsToPreview(xlWorksheet)

            MessageBox.Show($"Found {studentsToImport.Count} students in the file. Review and click 'Import Selected' to proceed.")

        Catch ex As Exception
            MessageBox.Show("Error importing from Excel: " & ex.Message)
        Finally
            ' Clean up Excel objects
            If xlWorksheet IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorksheet)
            If xlWorkbook IsNot Nothing Then
                xlWorkbook.Close(False)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook)
            End If
            If xlApp IsNot Nothing Then
                xlApp.Quit()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp)
            End If
        End Try
    End Sub

    Private Sub LoadStudentsToPreview(worksheet As Excel.Worksheet)
        ' Start from row 8 as specified
        Dim row As Integer = 8

        ' Check if row 8 contains headers by looking for typical header text
        If worksheet.Cells(row, 1).Value IsNot Nothing AndAlso
           (worksheet.Cells(row, 1).Value.ToString().ToLower().Contains("id") OrElse
            worksheet.Cells(row, 1).Value.ToString().ToLower().Contains("student") OrElse
            worksheet.Cells(row, 1).Value.ToString().ToLower().Contains("#")) Then
            ' Skip the header row
            row += 1
        End If

        While True
            ' Exit if we reach an empty StudentID
            If worksheet.Cells(row, 1).Value Is Nothing Then Exit While

            ' Create new student object
            Dim student As New Student() With {
                .StudentID = If(worksheet.Cells(row, 1).Value IsNot Nothing, worksheet.Cells(row, 1).Value.ToString(), ""),
                .Name = If(worksheet.Cells(row, 2).Value IsNot Nothing, worksheet.Cells(row, 2).Value.ToString(), ""),
                .GPA = If(worksheet.Cells(row, 3).Value IsNot Nothing, worksheet.Cells(row, 3).Value.ToString(), ""),
                .Remarks = If(worksheet.Cells(row, 4).Value IsNot Nothing, worksheet.Cells(row, 4).Value.ToString(), "")
            }

            ' Add to collection
            studentsToImport.Add(student)

            ' Add to ListView
            Dim lvi As New ListViewItem(student.StudentID)
            lvi.SubItems.Add(student.Name)
            lvi.SubItems.Add(student.GPA)
            lvi.SubItems.Add(student.Remarks)
            lvi.Tag = student ' Store reference to student object
            lvprofessor.Items.Add(lvi)

            row += 1
        End While
    End Sub

    Private Sub ReadHeaderInformation(worksheet As Excel.Worksheet)
        ' Read header information from specified positions
        txtstudsy.Text = If(worksheet.Cells(1, 2).Value IsNot Nothing, worksheet.Cells(1, 2).Value.ToString(), "")
        cbostudcourse.Text = If(worksheet.Cells(2, 2).Value IsNot Nothing, worksheet.Cells(2, 2).Value.ToString(), "")
        txtstudsection.Text = If(worksheet.Cells(3, 2).Value IsNot Nothing, worksheet.Cells(3, 2).Value.ToString(), "")
        txtsubjcode.Text = If(worksheet.Cells(4, 2).Value IsNot Nothing, worksheet.Cells(4, 2).Value.ToString(), "")
        txtsubjtitle.Text = If(worksheet.Cells(5, 2).Value IsNot Nothing, worksheet.Cells(5, 2).Value.ToString(), "")
        txtstudsem.Text = If(worksheet.Cells(6, 2).Value IsNot Nothing, worksheet.Cells(6, 2).Value.ToString(), "")
        txtstudprof.Text = If(worksheet.Cells(1, 4).Value IsNot Nothing, worksheet.Cells(1, 4).Value.ToString(), "")
        txtgradeid.Text = If(worksheet.Cells(2, 4).Value IsNot Nothing, worksheet.Cells(2, 4).Value.ToString(), "")
    End Sub

    Private Sub ImportStudentData(worksheet As Excel.Worksheet)
        Try
            Using connection As New OleDbConnection(connectionString)
                connection.Open()

                ' Start transaction for batch operations
                Dim transaction As OleDbTransaction = connection.BeginTransaction()

                ' Prepare insert command
                Dim insertQuery As String = "INSERT INTO tbl_professor (SchoolYear, Course, Section, SubjectCode, " &
                                    "SubjectTitle, Semester, Instructor, GradeID, StudentID, StudentName, " &
                                    "GPA, Remarks) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"

                Dim command As New OleDbCommand(insertQuery, connection, transaction)

                ' Add parameters (in order)
                command.Parameters.Add("@sy", OleDbType.VarChar)
                command.Parameters.Add("@course", OleDbType.VarChar)
                '... add all other parameters ...

                Dim row As Integer = 9 ' Start from row 8 as specified
                Dim recordsImported As Integer = 0
                Dim recordsSkipped As Integer = 0
                Dim errors As New List(Of String)

                While True
                    ' Exit if we reach an empty StudentID
                    If worksheet.Cells(row, 1).Value Is Nothing OrElse String.IsNullOrWhiteSpace(worksheet.Cells(row, 1).Value.ToString()) Then
                        Exit While
                    End If

                    Try
                        ' Get values from Excel
                        Dim studentID As String = worksheet.Cells(row, 1).Value.ToString().Trim()
                        Dim studentName As String = If(worksheet.Cells(row, 2).Value IsNot Nothing, worksheet.Cells(row, 2).Value.ToString().Trim(), "")
                        Dim gpa As String = If(worksheet.Cells(row, 3).Value IsNot Nothing, worksheet.Cells(row, 3).Value.ToString().Trim(), "")

                        ' Validate required fields
                        If String.IsNullOrEmpty(studentID) Then
                            errors.Add($"Row {row}: Missing StudentID")
                            row += 1
                            Continue While
                        End If

                        ' Check for duplicates
                        Dim checkQuery As String = "SELECT COUNT(*) FROM tbl_professor WHERE StudentID = ? AND SchoolYear = ? AND Course = ? AND Section = ? AND SubjectCode = ?"
                        Using checkCommand As New OleDbCommand(checkQuery, connection, transaction)
                            checkCommand.Parameters.AddWithValue("@studentid", studentID)
                            checkCommand.Parameters.AddWithValue("@sy", txtstudsy.Text)
                            checkCommand.Parameters.AddWithValue("@course", cbostudcourse.Text)
                            checkCommand.Parameters.AddWithValue("@section", txtstudsection.Text)
                            checkCommand.Parameters.AddWithValue("@subjcode", txtsubjcode.Text)

                            Dim exists As Integer = CInt(checkCommand.ExecuteScalar())

                            If exists > 0 Then
                                recordsSkipped += 1
                                row += 1
                                Continue While
                            End If
                        End Using

                        ' Set parameter values
                        command.Parameters("@sy").Value = txtstudsy.Text
                        command.Parameters("@course").Value = cbostudcourse.Text
                        '... set all other parameters ...
                        command.Parameters("@studentid").Value = studentID
                        command.Parameters("@studentname").Value = studentName
                        command.Parameters("@gpa").Value = If(String.IsNullOrEmpty(gpa), DBNull.Value, Convert.ToDouble(gpa))
                        command.Parameters("@remarks").Value = worksheet.Cells(row, 4).Value.ToString()

                        ' Execute insert
                        command.ExecuteNonQuery()
                        recordsImported += 1

                    Catch ex As Exception
                        errors.Add($"Row {row}: Error - {ex.Message}")
                    End Try

                    row += 1
                End While

                ' Commit transaction if no errors
                If errors.Count = 0 Then
                    transaction.Commit()
                    MessageBox.Show($"Import completed!{vbCrLf}Records imported: {recordsImported}{vbCrLf}Records skipped (duplicates): {recordsSkipped}")
                Else
                    transaction.Rollback()
                    MessageBox.Show($"Import completed with errors:{vbCrLf}{String.Join(vbCrLf, errors.Take(10))}" &
                              If(errors.Count > 10, vbCrLf & $"(and {errors.Count - 10} more errors...)", ""))
                End If

                LoadProfessorData() ' Refresh the data view
            End Using
        Catch ex As Exception
            MessageBox.Show("Error importing student data: " & ex.Message)
        End Try
    End Sub

    ' Button event handlers
    Private Sub btnprofsave_Click(sender As Object, e As EventArgs) Handles btnprofsave.Click
        ' Check if we're saving from form fields or ListView selections
        If lvprofessor.CheckedItems.Count > 0 Then
            ' Save multiple selected students
            SaveMultipleStudents()
        Else
            ' Save single student from form fields
            SaveProfessorData()
        End If
    End Sub

    Private Sub SaveMultipleStudents()
        Try
            Using connection As New OleDbConnection(connectionString)
                connection.Open()

                ' Start transaction for batch operations
                Dim transaction As OleDbTransaction = connection.BeginTransaction()

                ' Prepare insert command (same as in SaveProfessorData)
                Dim query As String = "INSERT INTO tbl_professor ([SchoolYear], [Course], [Section], [SubjectCode], " &
                            "[SubjectTitle], [Semester], [Instructor], [GradeID], [StudentID], [StudentName], " &
                            "[GPA], [Remarks]) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"

                Dim command As New OleDbCommand(query, connection, transaction)

                Dim recordsSaved As Integer = 0
                Dim errors As New List(Of String)

                ' Process all checked items in the ListView
                For Each item As ListViewItem In lvprofessor.CheckedItems
                    Dim student As Student = CType(item.Tag, Student)

                    Try
                        ' Clear previous parameters
                        command.Parameters.Clear()

                        ' Add parameters with current values
                        command.Parameters.AddWithValue("", txtstudsy.Text)
                        command.Parameters.AddWithValue("", cbostudcourse.Text)
                        command.Parameters.AddWithValue("", txtstudsection.Text)
                        command.Parameters.AddWithValue("", txtsubjcode.Text)
                        command.Parameters.AddWithValue("", txtsubjtitle.Text)
                        command.Parameters.AddWithValue("", txtstudsem.Text)
                        command.Parameters.AddWithValue("", txtstudprof.Text)
                        command.Parameters.AddWithValue("", txtgradeid.Text)
                        command.Parameters.AddWithValue("", student.StudentID)
                        command.Parameters.AddWithValue("", student.Name)
                        command.Parameters.AddWithValue("", If(String.IsNullOrEmpty(student.GPA), DBNull.Value, CDbl(student.GPA)))
                        command.Parameters.AddWithValue("", student.Remarks)

                        ' Execute insert
                        command.ExecuteNonQuery()
                        recordsSaved += 1

                    Catch ex As Exception
                        errors.Add($"Student ID {student.StudentID}: {ex.Message}")
                    End Try
                Next

                ' Commit transaction if no errors
                If errors.Count = 0 Then
                    transaction.Commit()
                    MessageBox.Show($"Successfully saved {recordsSaved} student records!")
                Else
                    transaction.Rollback()
                    MessageBox.Show($"Completed with errors:{vbCrLf}{String.Join(vbCrLf, errors.Take(10))}" &
                          If(errors.Count > 10, vbCrLf & $"(and {errors.Count - 10} more errors...)", ""))
                End If

                ' Refresh the data view
                LoadProfessorData()
            End Using
        Catch ex As Exception
            MessageBox.Show("Error saving student data: " & ex.Message)
        End Try
    End Sub

    Private Sub SaveProfessorData()
        ' Declare the query variable at the method level
        Dim query As String = "INSERT INTO tbl_professor ([SchoolYear], [Course], [Section], [SubjectCode], " &
                        "[SubjectTitle], [Semester], [Instructor], [GradeID], [StudentID], [StudentName], " &
                        "[GPA], [Remarks]) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"

        Try
            Using connection As New OleDbConnection(connectionString)
                connection.Open()
                Dim command As New OleDbCommand(query, connection)

                ' Add parameters in EXACT order of columns
                command.Parameters.AddWithValue("", txtstudsy.Text)
                command.Parameters.AddWithValue("", cbostudcourse.Text)
                command.Parameters.AddWithValue("", txtstudsection.Text)
                command.Parameters.AddWithValue("", txtsubjcode.Text)
                command.Parameters.AddWithValue("", txtsubjtitle.Text)
                command.Parameters.AddWithValue("", txtstudsem.Text)
                command.Parameters.AddWithValue("", txtstudprof.Text)
                command.Parameters.AddWithValue("", txtgradeid.Text)
                command.Parameters.AddWithValue("", txtstudid.Text)
                command.Parameters.AddWithValue("", txtstudname.Text)
                command.Parameters.AddWithValue("", If(txtstudgpa.Text.Trim() = "", DBNull.Value, CDbl(txtstudgpa.Text)))
                command.Parameters.AddWithValue("", txtstudremarks.Text)

                command.ExecuteNonQuery()
                MessageBox.Show("Data saved successfully!")
                ClearFormFields()
                LoadProfessorData()
            End Using
        Catch ex As Exception
            MessageBox.Show("Error saving data: " & ex.Message & vbCrLf & "Full query: " & query,
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ClearFormFields()
        txtstudid.Clear()
        txtstudname.Clear()
        txtstudgpa.Clear()
        txtstudremarks.Clear()
        ' Add other fields as needed
    End Sub

    Private Sub btnprofupdate_Click(sender As Object, e As EventArgs) Handles btnprofupdate.Click
        UpdateProfessorData()
    End Sub

    Private Sub UpdateProfessorData()
        Try
            Using connection As New OleDbConnection(connectionString)
                connection.Open()
                Dim query As String = "UPDATE tbl_professor SET SchoolYear=?, Course=?, Section=?, " &
                                    "SubjectCode=?, SubjectTitle=?, Semester=?, Instructor=?, GradeID=?, " &
                                    "StudentName=?, GPA=?, Remarks=? WHERE StudentID=?"

                Dim command As New OleDbCommand(query, connection)
                command.Parameters.AddWithValue("@sy", txtstudsy.Text)
                command.Parameters.AddWithValue("@course", cbostudcourse.Text)
                command.Parameters.AddWithValue("@section", txtstudsection.Text)
                command.Parameters.AddWithValue("@subjcode", txtsubjcode.Text)
                command.Parameters.AddWithValue("@subjtitle", txtsubjtitle.Text)
                command.Parameters.AddWithValue("@semester", txtstudsem.Text)
                command.Parameters.AddWithValue("@instructor", txtstudprof.Text)
                command.Parameters.AddWithValue("@gradeid", txtgradeid.Text)
                command.Parameters.AddWithValue("@studentname", txtstudname.Text)
                command.Parameters.AddWithValue("@gpa", If(txtstudgpa.Text.Trim() = "", DBNull.Value, Convert.ToDouble(txtstudgpa.Text)))
                command.Parameters.AddWithValue("@remarks", txtstudremarks.Text)
                command.Parameters.AddWithValue("@studentid", txtstudid.Text)

                Dim rowsAffected As Integer = command.ExecuteNonQuery()
                If rowsAffected > 0 Then
                    MessageBox.Show("Data updated successfully!")
                    LoadProfessorData()
                Else
                    MessageBox.Show("No record found to update.")
                End If
            End Using
        Catch ex As Exception
            MessageBox.Show("Error updating data: " & ex.Message)
        End Try
    End Sub

    Private Sub btnprofdel_Click(sender As Object, e As EventArgs) Handles btnprofdel.Click
        DeleteProfessorData()
    End Sub

    Private Sub DeleteProfessorData()
        If MessageBox.Show("Are you sure you want to delete this record?", "Confirm Delete",
                          MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Try
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    Dim query As String = "DELETE FROM tbl_professor WHERE StudentID = ?"
                    Dim command As New OleDbCommand(query, connection)
                    command.Parameters.AddWithValue("@studentid", txtstudid.Text)

                    Dim rowsAffected As Integer = command.ExecuteNonQuery()
                    If rowsAffected > 0 Then
                        MessageBox.Show("Record deleted successfully!")
                        ClearFormFields()
                        LoadProfessorData()
                    Else
                        MessageBox.Show("No record found to delete.")
                    End If
                End Using
            Catch ex As Exception
                MessageBox.Show("Error deleting record: " & ex.Message)
            End Try
        End If
    End Sub

    Private Sub btnprofbatch_Click(sender As Object, e As EventArgs) Handles btnprofbatch.Click
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            ImportFromExcel(OpenFileDialog1.FileName)
        End If
    End Sub

    Private Sub btnprofback_Click(sender As Object, e As EventArgs) Handles btnprofback.Click
        Me.Hide()
    End Sub

    Private Sub btnprofexit_Click(sender As Object, e As EventArgs) Handles btnprofexit.Click
        Application.Exit()
    End Sub

    Private Sub lvprofessor_DoubleClick(sender As Object, e As EventArgs) Handles lvprofessor.DoubleClick
        ' Handle double-click on ListView items
        If lvprofessor.SelectedItems.Count > 0 Then
            Dim selectedStudent As Student = CType(lvprofessor.SelectedItems(0).Tag, Student)
            ' Display the selected student's details in form controls
            txtstudid.Text = selectedStudent.StudentID
            txtstudname.Text = selectedStudent.Name
            txtstudgpa.Text = selectedStudent.GPA
            txtstudremarks.Text = selectedStudent.Remarks
        End If
    End Sub

    ' Other event handlers that don't need implementation can be removed
    Private Sub txtgradeid_TextChanged(sender As Object, e As EventArgs) Handles txtgradeid.TextChanged
    End Sub
    Private Sub txtstudid_TextChanged(sender As Object, e As EventArgs) Handles txtstudid.TextChanged
    End Sub
    Private Sub txtstudname_TextChanged(sender As Object, e As EventArgs) Handles txtstudname.TextChanged
    End Sub
    Private Sub txtstudgpa_TextChanged(sender As Object, e As EventArgs) Handles txtstudgpa.TextChanged
    End Sub
    Private Sub txtstudremarks_TextChanged(sender As Object, e As EventArgs) Handles txtstudremarks.TextChanged
    End Sub
    Private Sub txtstudsy_TextChanged(sender As Object, e As EventArgs) Handles txtstudsy.TextChanged
    End Sub
    Private Sub cbostudcourse_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbostudcourse.SelectedIndexChanged
    End Sub
    Private Sub txtstudsection_TextChanged(sender As Object, e As EventArgs) Handles txtstudsection.TextChanged
    End Sub
    Private Sub txtsubjcode_TextChanged(sender As Object, e As EventArgs) Handles txtsubjcode.TextChanged
    End Sub
    Private Sub txtsubjtitle_TextChanged(sender As Object, e As EventArgs) Handles txtsubjtitle.TextChanged
    End Sub
    Private Sub txtstudsem_TextChanged(sender As Object, e As EventArgs) Handles txtstudsem.TextChanged
    End Sub
    Private Sub txtstudprof_TextChanged(sender As Object, e As EventArgs) Handles txtstudprof.TextChanged
    End Sub

    Private Sub txtremarks_TextChanged(sender As Object, e As EventArgs) Handles txtstudremarks.TextChanged

    End Sub
End Class