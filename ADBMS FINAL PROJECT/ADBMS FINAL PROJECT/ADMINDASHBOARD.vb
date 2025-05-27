Public Class ADMINDASHBOARD

    Private Sub ADMINDASHBOARD_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        con() ' Initialize connection
        loadAdminList()
        ' Setup password textbox to mask characters initially
        txtadpass.UseSystemPasswordChar = True
    End Sub

    ' Validate all required fields are filled before saving
    Private Function ValidateFields() As Boolean
        If txtaduser.Text.Trim() = "" Then
            MsgBox("Enter username", vbInformation, "Missing")
            txtaduser.Focus()
            Return False
        ElseIf txtadpass.Text.Trim() = "" Then
            MsgBox("Enter password", vbInformation, "Missing")
            txtadpass.Focus()
            Return False
        ElseIf txtademail.Text.Trim() = "" Then
            MsgBox("Enter email", vbInformation, "Missing")
            txtademail.Focus()
            Return False
        ElseIf txtadfirst.Text.Trim() = "" Then
            MsgBox("Enter first name", vbInformation, "Missing")
            txtadfirst.Focus()
            Return False
        ElseIf txtadlast.Text.Trim() = "" Then
            MsgBox("Enter last name", vbInformation, "Missing")
            txtadlast.Focus()
            Return False
        End If
        Return True
    End Function

    Private Sub btnsave_Click(sender As Object, e As EventArgs) Handles btnsave.Click
        If Not ValidateFields() Then Exit Sub

        Dim rs As New ADODB.Recordset
        Try
            If rs.State = 1 Then rs.Close()

            rs.Open("SELECT * FROM [tbl-adminregister] WHERE ADM_USER='" & txtaduser.Text.Trim() & "'", cn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)

            If Not rs.EOF Then
                MsgBox("Username already exists, choose a different one", vbInformation, "Exist")
                txtaduser.Clear()
                txtaduser.Focus()
                rs.Close()
                Exit Sub
            End If

            rs.AddNew()
            rs.Fields("ADM_USER").Value = txtaduser.Text.Trim()
            rs.Fields("ADM_PASS").Value = txtadpass.Text.Trim()
            rs.Fields("ADM_EMAIL").Value = txtademail.Text.Trim()
            rs.Fields("ADM_FIRST").Value = txtadfirst.Text.Trim()
            rs.Fields("ADM_MIDDLE").Value = txtadmid.Text.Trim()
            rs.Fields("ADM_LAST").Value = txtadlast.Text.Trim()
            rs.Fields("ADM_CONTACT").Value = txtadcon.Text.Trim()
            rs.Fields("ADM_BIRTHDATE").Value = dtpbirth.Value
            rs.Fields("ADM_HOME").Value = txthome.Text.Trim()
            rs.Fields("ADM_STREET").Value = txtstreet.Text.Trim()
            rs.Fields("ADM_BRGY").Value = txtbrgy.Text.Trim()
            rs.Fields("ADM_MUNICITY").Value = txtmunicity.Text.Trim()
            rs.Fields("ADM_PROV").Value = txtprov.Text.Trim()
            rs.Fields("ADM_ZIP").Value = txtzip.Text.Trim()
            rs.Update()

            MsgBox("Admin information was saved successfully", vbInformation, "Save")
            rs.Close()

            clearFields()
            loadAdminList()
        Catch ex As Exception
            MsgBox("Error saving data: " & ex.Message, vbCritical, "Error")
            If rs.State = 1 Then rs.Close()
        End Try
    End Sub

    Private Sub btndele_Click(sender As Object, e As EventArgs) Handles btndele.Click
        If txtaduser.Text.Trim() = "" Then
            MsgBox("Select an admin to delete", vbInformation, "Missing")
            Exit Sub
        End If

        Dim rs As New ADODB.Recordset
        Try
            If rs.State = 1 Then rs.Close()

            rs.Open("SELECT * FROM [tbl-adminregister] WHERE ADM_USER='" & txtaduser.Text.Trim() & "'", cn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)

            If rs.EOF Then
                MsgBox("Admin user not found", vbExclamation, "Not Found")
                rs.Close()
                Exit Sub
            End If

            If MsgBox("Are you sure you want to delete this admin?", vbQuestion + vbYesNo, "Delete") = vbYes Then
                rs.Delete()
                rs.Update()
                MsgBox("Admin was deleted", vbInformation, "Delete")
            End If

            rs.Close()
            clearFields()
            loadAdminList()
        Catch ex As Exception
            MsgBox("Error deleting data: " & ex.Message, vbCritical, "Error")
            If rs.State = 1 Then rs.Close()
        End Try
    End Sub

    Private Sub btnexit_Click(sender As Object, e As EventArgs) Handles btnexit.Click
        Application.Exit()
    End Sub

    Private Sub clearFields()
        txtaduser.Clear()
        txtadpass.Clear()
        txtademail.Clear()
        txtadfirst.Clear()
        txtadmid.Clear()
        txtadlast.Clear()
        txtadcon.Clear()
        dtpbirth.Value = DateTime.Now
        txthome.Clear()
        txtstreet.Clear()
        txtbrgy.Clear()
        txtmunicity.Clear()
        txtprov.Clear()
        txtzip.Clear()
        txtaduser.Focus()
    End Sub

    Private Sub loadAdminList()
        Dim rs As New ADODB.Recordset
        Try
            If rs.State = 1 Then rs.Close()
            rs.Open("SELECT * FROM [tbl-adminregister] ORDER BY ADM_LAST ASC", cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

            lvadmin.Items.Clear()

            While Not rs.EOF
                Dim item As New ListViewItem(rs.Fields("ADM_USER").Value.ToString())
                item.SubItems.Add(rs.Fields("ADM_FIRST").Value.ToString())
                item.SubItems.Add(rs.Fields("ADM_MIDDLE").Value.ToString())
                item.SubItems.Add(rs.Fields("ADM_LAST").Value.ToString())
                item.SubItems.Add(rs.Fields("ADM_EMAIL").Value.ToString())

                Dim addressParts As New List(Of String)
                If Not IsDBNull(rs.Fields("ADM_HOME").Value) AndAlso rs.Fields("ADM_HOME").Value.ToString().Trim() <> "" Then addressParts.Add(rs.Fields("ADM_HOME").Value.ToString().Trim())
                If Not IsDBNull(rs.Fields("ADM_STREET").Value) AndAlso rs.Fields("ADM_STREET").Value.ToString().Trim() <> "" Then addressParts.Add(rs.Fields("ADM_STREET").Value.ToString().Trim())
                If Not IsDBNull(rs.Fields("ADM_BRGY").Value) AndAlso rs.Fields("ADM_BRGY").Value.ToString().Trim() <> "" Then addressParts.Add(rs.Fields("ADM_BRGY").Value.ToString().Trim())
                If Not IsDBNull(rs.Fields("ADM_MUNICITY").Value) AndAlso rs.Fields("ADM_MUNICITY").Value.ToString().Trim() <> "" Then addressParts.Add(rs.Fields("ADM_MUNICITY").Value.ToString().Trim())
                If Not IsDBNull(rs.Fields("ADM_PROV").Value) AndAlso rs.Fields("ADM_PROV").Value.ToString().Trim() <> "" Then addressParts.Add(rs.Fields("ADM_PROV").Value.ToString().Trim())
                If Not IsDBNull(rs.Fields("ADM_ZIP").Value) AndAlso rs.Fields("ADM_ZIP").Value.ToString().Trim() <> "" Then addressParts.Add(rs.Fields("ADM_ZIP").Value.ToString().Trim())

                item.SubItems.Add(String.Join(", ", addressParts))

                lvadmin.Items.Add(item)
                rs.MoveNext()
            End While

            rs.Close()
        Catch ex As Exception
            MsgBox("Error loading admin list: " & ex.Message, vbCritical, "Error")
            If rs.State = 1 Then rs.Close()
        End Try
    End Sub

    ' This sub is called whenever the search textbox text changes
    Private Sub txtsearch_TextChanged(sender As Object, e As EventArgs) Handles txtsearch.TextChanged
        If String.IsNullOrWhiteSpace(txtsearch.Text) Then
            ' If search box empty, load all admins
            loadAdminList()
        Else
            ' If no filter selected, ask user
            If cbofilter.SelectedIndex = -1 Then
                MsgBox("Please select a filter option.", vbInformation, "Filter Missing")
                Exit Sub
            End If

            Dim fieldName As String = ""
            Select Case cbofilter.SelectedIndex
                Case 0 ' USER NAME
                    fieldName = "ADM_USER"
                Case 1 ' FIRST NAME
                    fieldName = "ADM_FIRST"
                Case 2 ' MIDDLE NAME
                    fieldName = "ADM_MIDDLE"
                Case 3 ' LAST NAME
                    fieldName = "ADM_LAST"
                Case 4 ' EMAIL
                    fieldName = "ADM_EMAIL"
                Case 5 ' ADDRESS
                    Dim searchText As String = txtsearch.Text.Trim()
                    Dim sql As String = "SELECT * FROM [tbl-adminregister] WHERE " &
                                        "(ADM_HOME & ADM_STREET & ADM_BRGY & ADM_MUNICITY & ADM_PROV & ADM_ZIP) LIKE '%" & searchText & "%' ORDER BY ADM_LAST ASC"
                    SearchWithSQL(sql)
                    Exit Sub
                Case Else
                    MsgBox("Invalid filter selected.", vbExclamation, "Error")
                    Exit Sub
            End Select

            Dim query As String = $"SELECT * FROM [tbl-adminregister] WHERE {fieldName} LIKE '%{txtsearch.Text.Trim()}%' ORDER BY ADM_LAST ASC"
            SearchWithSQL(query)
        End If
    End Sub

    Private Sub SearchWithSQL(ByVal query As String)
        Dim rs As New ADODB.Recordset
        Try
            If rs.State = 1 Then rs.Close()
            rs.Open(query, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

            lvadmin.Items.Clear()

            If rs.EOF Then
                ' No records found - clear listview but no message to avoid popup on every key press
                rs.Close()
                Exit Sub
            End If

            While Not rs.EOF
                Dim item As New ListViewItem(rs.Fields("ADM_USER").Value.ToString())
                item.SubItems.Add(rs.Fields("ADM_FIRST").Value.ToString())
                item.SubItems.Add(rs.Fields("ADM_MIDDLE").Value.ToString())
                item.SubItems.Add(rs.Fields("ADM_LAST").Value.ToString())
                item.SubItems.Add(rs.Fields("ADM_EMAIL").Value.ToString())

                Dim addressParts As New List(Of String)
                If Not IsDBNull(rs.Fields("ADM_HOME").Value) AndAlso rs.Fields("ADM_HOME").Value.ToString().Trim() <> "" Then addressParts.Add(rs.Fields("ADM_HOME").Value.ToString().Trim())
                If Not IsDBNull(rs.Fields("ADM_STREET").Value) AndAlso rs.Fields("ADM_STREET").Value.ToString().Trim() <> "" Then addressParts.Add(rs.Fields("ADM_STREET").Value.ToString().Trim())
                If Not IsDBNull(rs.Fields("ADM_BRGY").Value) AndAlso rs.Fields("ADM_BRGY").Value.ToString().Trim() <> "" Then addressParts.Add(rs.Fields("ADM_BRGY").Value.ToString().Trim())
                If Not IsDBNull(rs.Fields("ADM_MUNICITY").Value) AndAlso rs.Fields("ADM_MUNICITY").Value.ToString().Trim() <> "" Then addressParts.Add(rs.Fields("ADM_MUNICITY").Value.ToString().Trim())
                If Not IsDBNull(rs.Fields("ADM_PROV").Value) AndAlso rs.Fields("ADM_PROV").Value.ToString().Trim() <> "" Then addressParts.Add(rs.Fields("ADM_PROV").Value.ToString().Trim())
                If Not IsDBNull(rs.Fields("ADM_ZIP").Value) AndAlso rs.Fields("ADM_ZIP").Value.ToString().Trim() <> "" Then addressParts.Add(rs.Fields("ADM_ZIP").Value.ToString().Trim())

                item.SubItems.Add(String.Join(", ", addressParts))

                lvadmin.Items.Add(item)
                rs.MoveNext()
            End While

            rs.Close()
        Catch ex As Exception
            MsgBox("Error performing search: " & ex.Message, vbCritical, "Error")
            If rs.State = 1 Then rs.Close()
        End Try
    End Sub

    ' Checkbox to toggle password visibility
    Private Sub cbshowpass_CheckedChanged(sender As Object, e As EventArgs) Handles cbShowPass.CheckedChanged
        If cbShowPass.Checked Then
            txtadpass.UseSystemPasswordChar = False
        Else
            txtadpass.UseSystemPasswordChar = True
        End If
    End Sub

    ' When an item in ListView is clicked, load details to fields
    Private Sub lvadmin_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lvadmin.SelectedIndexChanged
        If lvadmin.SelectedItems.Count = 0 Then Exit Sub

        Dim selectedUser As String = lvadmin.SelectedItems(0).Text
        Dim rs As New ADODB.Recordset
        Try
            If rs.State = 1 Then rs.Close()

            rs.Open("SELECT * FROM [tbl-adminregister] WHERE ADM_USER='" & selectedUser & "'", cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

            If Not rs.EOF Then
                txtaduser.Text = rs.Fields("ADM_USER").Value.ToString()
                txtadpass.Text = rs.Fields("ADM_PASS").Value.ToString()
                txtademail.Text = rs.Fields("ADM_EMAIL").Value.ToString()
                txtadfirst.Text = rs.Fields("ADM_FIRST").Value.ToString()
                txtadmid.Text = rs.Fields("ADM_MIDDLE").Value.ToString()
                txtadlast.Text = rs.Fields("ADM_LAST").Value.ToString()
                txtadcon.Text = rs.Fields("ADM_CONTACT").Value.ToString()
                dtpbirth.Value = If(IsDBNull(rs.Fields("ADM_BIRTHDATE").Value), DateTime.Now, CDate(rs.Fields("ADM_BIRTHDATE").Value))
                txthome.Text = rs.Fields("ADM_HOME").Value.ToString()
                txtstreet.Text = rs.Fields("ADM_STREET").Value.ToString()
                txtbrgy.Text = rs.Fields("ADM_BRGY").Value.ToString()
                txtmunicity.Text = rs.Fields("ADM_MUNICITY").Value.ToString()
                txtprov.Text = rs.Fields("ADM_PROV").Value.ToString()
                txtzip.Text = rs.Fields("ADM_ZIP").Value.ToString()
            End If
            rs.Close()
        Catch ex As Exception
            MsgBox("Error loading admin data: " & ex.Message, vbCritical, "Error")
            If rs.State = 1 Then rs.Close()
        End Try
    End Sub

    Private Sub btnadclear_Click(sender As Object, e As EventArgs) Handles btnadclear.Click
        clearFields()
    End Sub

    Private Sub btnadupdate_Click(sender As Object, e As EventArgs) Handles btnadupdate.Click
        If Not ValidateFields() Then Exit Sub

        Dim rs As New ADODB.Recordset
        Try
            If rs.State = 1 Then rs.Close()

            rs.Open("SELECT * FROM [tbl-adminregister] WHERE ADM_USER='" & txtaduser.Text.Trim() & "'", cn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)

            If rs.EOF Then
                MsgBox("Admin user not found to update.", vbExclamation, "Not Found")
                rs.Close()
                Exit Sub
            End If

            Dim isChanged As Boolean = False

            ' *** PASSWORD FIELD IS NEVER TOUCHED OR UPDATED ***

            If rs.Fields("ADM_EMAIL").Value.ToString() <> txtademail.Text.Trim() Then
                rs.Fields("ADM_EMAIL").Value = txtademail.Text.Trim()
                isChanged = True
            End If

            If rs.Fields("ADM_FIRST").Value.ToString() <> txtadfirst.Text.Trim() Then
                rs.Fields("ADM_FIRST").Value = txtadfirst.Text.Trim()
                isChanged = True
            End If

            If rs.Fields("ADM_MIDDLE").Value.ToString() <> txtadmid.Text.Trim() Then
                rs.Fields("ADM_MIDDLE").Value = txtadmid.Text.Trim()
                isChanged = True
            End If

            If rs.Fields("ADM_LAST").Value.ToString() <> txtadlast.Text.Trim() Then
                rs.Fields("ADM_LAST").Value = txtadlast.Text.Trim()
                isChanged = True
            End If

            If rs.Fields("ADM_CONTACT").Value.ToString() <> txtadcon.Text.Trim() Then
                rs.Fields("ADM_CONTACT").Value = txtadcon.Text.Trim()
                isChanged = True
            End If

            If CDate(rs.Fields("ADM_BIRTHDATE").Value) <> dtpbirth.Value.Date Then
                rs.Fields("ADM_BIRTHDATE").Value = dtpbirth.Value.Date
                isChanged = True
            End If

            If rs.Fields("ADM_HOME").Value.ToString() <> txthome.Text.Trim() Then
                rs.Fields("ADM_HOME").Value = txthome.Text.Trim()
                isChanged = True
            End If

            If rs.Fields("ADM_STREET").Value.ToString() <> txtstreet.Text.Trim() Then
                rs.Fields("ADM_STREET").Value = txtstreet.Text.Trim()
                isChanged = True
            End If

            If rs.Fields("ADM_BRGY").Value.ToString() <> txtbrgy.Text.Trim() Then
                rs.Fields("ADM_BRGY").Value = txtbrgy.Text.Trim()
                isChanged = True
            End If

            If rs.Fields("ADM_MUNICITY").Value.ToString() <> txtmunicity.Text.Trim() Then
                rs.Fields("ADM_MUNICITY").Value = txtmunicity.Text.Trim()
                isChanged = True
            End If

            If rs.Fields("ADM_PROV").Value.ToString() <> txtprov.Text.Trim() Then
                rs.Fields("ADM_PROV").Value = txtprov.Text.Trim()
                isChanged = True
            End If

            If rs.Fields("ADM_ZIP").Value.ToString() <> txtzip.Text.Trim() Then
                rs.Fields("ADM_ZIP").Value = txtzip.Text.Trim()
                isChanged = True
            End If

            If isChanged Then
                rs.Update()
                MsgBox("Admin information updated successfully.", vbInformation, "Update")
                clearFields()
                loadAdminList()
            Else
                MsgBox("No changes detected to update.", vbInformation, "No Update")
            End If

            rs.Close()

        Catch ex As Exception
            MsgBox("Error updating data: " & ex.Message, vbCritical, "Error")
            If rs.State = 1 Then rs.Close()
        End Try
    End Sub

    Private Sub btnadback_Click(sender As Object, e As EventArgs) Handles btnadback.Click
        Me.Hide()
        Form2.Show()  ' Palitan mo ang ADMINLOGIN sa pangalan ng login form mo
    End Sub

    Private Sub txtaduser_TextChanged(sender As Object, e As EventArgs) Handles txtaduser.TextChanged

    End Sub
End Class
