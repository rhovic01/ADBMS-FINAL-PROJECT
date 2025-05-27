Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Optional: Any code when Form1 loads
    End Sub

    ' Admin button click event
    Private Sub btnadmin_Click(sender As Object, e As EventArgs) Handles btnadmin.Click
        OpenForm("admin")
    End Sub

    ' Professor button click event
    Private Sub btnprofessor_Click(sender As Object, e As EventArgs) Handles btnprofessor.Click
        OpenForm("professor")
    End Sub

    ' Function to open different forms based on role
    Private Sub OpenForm(role As String)
        If role = "admin" Then
            ' Open Form2 for admin
            Dim form2 As New Form2()
            form2.Show()
            Me.Hide()
        ElseIf role = "professor" Then
            ' Open PROF_LOGIN form for professor
            Dim form3 As New PROF_LOGIN() ' Use PROF_LOGIN instead of PROF-LOGIN
            form3.Show()
            Me.Hide()
        End If
    End Sub
End Class
