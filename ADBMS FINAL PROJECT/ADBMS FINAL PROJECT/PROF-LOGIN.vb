Public Class PROF_LOGIN

    Private Sub PROF_LOGIN_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        con() ' Open DB connection
        ' Initially mask the password textbox with asterisk
        txtprofpass.PasswordChar = "*"c
    End Sub

    Private Sub btnproflogin_Click(sender As Object, e As EventArgs) Handles btnproflogin.Click
        ' Validate inputs
        If txtprofuser.Text.Trim() = "" Then
            MsgBox("Please enter your username.", vbInformation, "Missing Username")
            txtprofuser.Focus()
            Exit Sub
        End If

        If txtprofpass.Text.Trim() = "" Then
            MsgBox("Please enter your password.", vbInformation, "Missing Password")
            txtprofpass.Focus()
            Exit Sub
        End If

        Dim rs As New ADODB.Recordset
        Try
            If rs.State = 1 Then rs.Close()

            rs.Open("SELECT * FROM [tbl-adminregister]", cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

            Dim isValidAccount As Boolean = False

            While Not rs.EOF
                Dim dbUser As String = rs.Fields("ADM_USER").Value.ToString()
                Dim dbPass As String = rs.Fields("ADM_PASS").Value.ToString()

                If dbUser = txtprofuser.Text.Trim() Then
                    If dbPass = txtprofpass.Text.Trim() Then
                        isValidAccount = True
                        rs.Close()

                        MsgBox("Login successful! Welcome " & dbUser, vbInformation, "Welcome")

                        ' Open Professor Dashboard
                        Dim dashboard As New PROFESSOR_DASHBOARD()
                        dashboard.Show()

                        ' Close the login form
                        Me.Close()

                        Exit Sub
                    Else
                        MsgBox("Incorrect password.", vbExclamation, "Login Failed")
                        txtprofpass.Focus()
                        rs.Close()
                        Exit Sub
                    End If
                End If
                rs.MoveNext()
            End While

            rs.Close()

            If Not isValidAccount Then
                MsgBox("Account not found or invalid credentials.", vbExclamation, "Login Failed")
                txtprofuser.Focus()
            End If

        Catch ex As Exception
            MsgBox("Error during login: " & ex.Message, vbCritical, "Error")
            If rs.State = 1 Then rs.Close()
        End Try
    End Sub

    Private Sub cbprofremember_CheckedChanged(sender As Object, e As EventArgs) Handles cbprofremember.CheckedChanged
        ' Toggle password visibility
        If cbprofremember.Checked Then
            txtprofpass.PasswordChar = ControlChars.NullChar ' Show password
        Else
            txtprofpass.PasswordChar = "*"c ' Mask password
        End If
    End Sub

End Class
