Public Class Form2

    ' Declare constants for admin username and password
    Private Const AdminUsername As String = "ADMIN"
    Private Const AdminPassword As String = "ADMIN123"

    ' Load event
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Optionally, set up default behavior for the checkbox (Show Password)
        cbremember.Checked = False
        txtadminpass.PasswordChar = "*" ' Default behavior: hide password
    End Sub

    ' Event handler for Login button click
    Private Sub btnAdminLogin_Click(sender As Object, e As EventArgs) Handles btnadminlogin.Click
        ' Check if the entered username and password match the admin credentials
        If txtadminuser.Text = AdminUsername AndAlso txtadminpass.Text = AdminPassword Then
            ' If correct, open Form2 (Admin Dashboard or next form)
            MessageBox.Show("Welcome, Admin!", "Login Successful", MessageBoxButtons.OK, MessageBoxIcon.Information)
            ADMINDASHBOARD.Show()
            Me.Hide() ' Hide the current form (Form1) after successful login
        Else
            ' If incorrect, show error message
            MessageBox.Show("Invalid Username or Password. Please try again.", "Login Failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    ' Event handler for "Show Password" checkbox
    Private Sub cbremember_CheckedChanged(sender As Object, e As EventArgs) Handles cbremember.CheckedChanged
        ' Toggle password visibility based on checkbox state
        If cbremember.Checked Then
            txtadminpass.PasswordChar = "" ' Show the password
        Else
            txtadminpass.PasswordChar = "*" ' Hide the password
        End If
    End Sub

End Class
