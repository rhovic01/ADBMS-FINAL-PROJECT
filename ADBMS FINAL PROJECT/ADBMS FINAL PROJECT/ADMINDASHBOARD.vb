Public Class ADMINDASHBOARD
    ' Form Load Event: Initializes the form and populates any necessary data
    Private Sub ADMINDASHBOARD_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Initialize database connection
        Call con()
        ' Load admin register list
        Call loadAdminList()
    End Sub

    ' Validate fields before performing any database operations
    Private Function ValidateFields() As Boolean
        If txtaduser.Text = "" Then
            MsgBox("Enter username", vbInformation, "Missing")
            txtaduser.Focus()
            Return False
        ElseIf txtadpass.Text = "" Then
            MsgBox("Enter password", vbInformation, "Missing")
            txtadpass.Focus()
            Return False
        ElseIf txtademail.Text = "" Then
            MsgBox("Enter email", vbInformation, "Missing")
            txtademail.Focus()
            Return False
        ElseIf txtadfirst.Text = "" Then
            MsgBox("Enter first name", vbInformation, "Missing")
            txtadfirst.Focus()
            Return False
        ElseIf txtadlast.Text = "" Then
            MsgBox("Enter last name", vbInformation, "Missing")
            txtadlast.Focus()
            Return False
        End If
        Return True
    End Function

    ' Save or update admin details in the database
    Private Sub btnsave_Click(sender As Object, e As EventArgs) Handles btnsave.Click
        If Not ValidateFields() Then Exit Sub

        rs = New ADODB.Recordset
        With rs
            If .State = 1 Then .Close()
            ' Check if username already exists in tbl-adminregister
            .Open("SELECT * FROM [tbl-adminregister] WHERE ADM_USER='" & txtaduser.Text & "'", cn, 1, 2)

            If .EOF = False Then
                MsgBox("Username already exists, choose a different one", vbInformation, "Exist")
                txtaduser.Text = ""
                txtaduser.Focus()
                Exit Sub
            End If

            .AddNew()
            ' Insert data into tbl-adminregister
            .Fields("ADM_USER").Value = txtaduser.Text
            .Fields("ADM_PASS").Value = txtadpass.Text
            .Fields("ADM_EMAIL").Value = txtademail.Text
            .Fields("ADM_FIRST").Value = txtadfirst.Text
            .Fields("ADM_MIDDLE").Value = txtadmid.Text
            .Fields("ADM_LAST").Value = txtadlast.Text
            .Fields("ADM_CONTACT").Value = txtadcon.Text
            .Fields("ADM_BIRTHDATE").Value = dtpbirth.Value
            .Fields("ADM_HOME").Value = txthome.Text
            .Fields("ADM_STREET").Value = txtstreet.Text
            .Fields("ADM_BRGY").Value = txtbrgy.Text
            .Fields("ADM_MUNICITY").Value = txtmunicity.Text
            .Fields("ADM_PROV").Value = txtprov.Text
            .Fields("ADM_ZIP").Value = txtzip.Text
            .Update()

            MsgBox("Admin information was saved successfully", vbInformation, "Save")
            .Close()
        End With
        ' Clear fields and refresh list
        Call clearFields()
        Call loadAdminList()
    End Sub

    ' Delete admin data from the database
    Private Sub btndele_Click(sender As Object, e As EventArgs) Handles btndele.Click
        If txtaduser.Text = "" Then
            MsgBox("Select an admin to delete", vbInformation, "Missing")
            Exit Sub
        End If

        rs = New ADODB.Recordset
        With rs
            If .State = 1 Then .Close()
            .Open("SELECT * FROM tbl-adminregister WHERE ADM_USER='" & txtaduser.Text & "'", cn, 1, 2)
            If MsgBox("Are you sure you want to delete this admin?", vbQuestion + vbYesNo, "Delete") = vbNo Then
                Exit Sub
            End If
            .Delete()
            .Update()
            MsgBox("Admin was deleted", vbInformation, "Delete")
            .Close()
        End With
        ' Clear fields and refresh list
        Call clearFields()
        Call loadAdminList()
    End Sub

    ' Exit application
    Private Sub btnexit_Click(sender As Object, e As EventArgs) Handles btnexit.Click
        End
    End Sub

    ' Clears input fields
    Sub clearFields()
        txtaduser.Text = ""
        txtadpass.Text = ""
        txtademail.Text = ""
        txtadfirst.Text = ""
        txtadmid.Text = ""
        txtadlast.Text = ""
        txtadcon.Text = ""
        dtpbirth.Value = DateTime.Now
        txthome.Text = ""
        txtstreet.Text = ""
        txtbrgy.Text = ""
        txtmunicity.Text = ""
        txtprov.Text = ""
        txtzip.Text = ""
        txtaduser.Focus()
    End Sub

    ' Loads the list of admin users into the ListView
    Sub loadAdminList()
        rs = New ADODB.Recordset
        With rs
            If .State = 1 Then .Close()

            Try
                .Open("SELECT * FROM [tbl-adminregister] ORDER BY ADM_LAST ASC", cn, 3, 1)
            Catch ex As Exception
                MsgBox("Error opening recordset: " & ex.Message)
                Exit Sub
            End Try

            If .EOF = True Then Exit Sub

            lvadmin.Items.Clear()

            While Not .EOF
                Try
                    Dim rec As New ListViewItem(.Fields("ADM_USER").Value.ToString())
                    rec.SubItems.Add(.Fields("ADM_FIRST").Value.ToString())
                    rec.SubItems.Add(.Fields("ADM_MIDDLE").Value.ToString())
                    rec.SubItems.Add(.Fields("ADM_LAST").Value.ToString())
                    rec.SubItems.Add(.Fields("ADM_EMAIL").Value.ToString())



                    ' Concatenate address parts into one string
                    Dim addressParts As New List(Of String)

                    If Not IsDBNull(.Fields("ADM_HOME").Value) AndAlso .Fields("ADM_HOME").Value.ToString().Trim() <> "" Then
                        addressParts.Add(.Fields("ADM_HOME").Value.ToString().Trim())
                    End If
                    If Not IsDBNull(.Fields("ADM_STREET").Value) AndAlso .Fields("ADM_STREET").Value.ToString().Trim() <> "" Then
                        addressParts.Add(.Fields("ADM_STREET").Value.ToString().Trim())
                    End If
                    If Not IsDBNull(.Fields("ADM_BRGY").Value) AndAlso .Fields("ADM_BRGY").Value.ToString().Trim() <> "" Then
                        addressParts.Add(.Fields("ADM_BRGY").Value.ToString().Trim())
                    End If
                    If Not IsDBNull(.Fields("ADM_MUNICITY").Value) AndAlso .Fields("ADM_MUNICITY").Value.ToString().Trim() <> "" Then
                        addressParts.Add(.Fields("ADM_MUNICITY").Value.ToString().Trim())
                    End If
                    If Not IsDBNull(.Fields("ADM_PROV").Value) AndAlso .Fields("ADM_PROV").Value.ToString().Trim() <> "" Then
                        addressParts.Add(.Fields("ADM_PROV").Value.ToString().Trim())
                    End If
                    If Not IsDBNull(.Fields("ADM_ZIP").Value) AndAlso .Fields("ADM_ZIP").Value.ToString().Trim() <> "" Then
                        addressParts.Add(.Fields("ADM_ZIP").Value.ToString().Trim())
                    End If

                    Dim fullAddress As String = String.Join(", ", addressParts)

                    rec.SubItems.Add(fullAddress)

                    lvadmin.Items.Add(rec)
                Catch ex As Exception
                    MsgBox("Error processing record: " & ex.Message)
                End Try

                .MoveNext()
            End While

            .Close()
        End With
    End Sub


    ' Search functionality to filter admin records based on a field
    Private Sub txtsearch_TextChanged(sender As Object, e As EventArgs) Handles txtsearch.TextChanged
        If String.IsNullOrEmpty(txtsearch.Text) Then
            ' If search text is empty, exit early (to avoid unnecessary database queries)
            Exit Sub
        End If

        rs = New ADODB.Recordset
        With rs
            If .State = 1 Then .Close()

            ' Check if the filter combobox is selected (index >= 0)
            If cbofilter.SelectedIndex = -1 Then
                MsgBox("Please select a filter option.", vbInformation, "Filter Missing")
                Exit Sub
            End If

            ' Build the query based on the selected filter option
            Dim query As String = ""

            Select Case cbofilter.SelectedIndex
                Case 0 ' Filter by ADM_USER
                    query = "SELECT * FROM tbl-adminregister WHERE ADM_USER LIKE '%" & txtsearch.Text & "%'"
                Case 1 ' Filter by ADM_LAST
                    query = "SELECT * FROM tbl-adminregister WHERE ADM_LAST LIKE '%" & txtsearch.Text & "%'"
                Case 2 ' Filter by ADM_FIRST
                    query = "SELECT * FROM tbl-adminregister WHERE ADM_FIRST LIKE '%" & txtsearch.Text & "%'"
                Case Else
                    Exit Sub ' No valid filter selected
            End Select

            ' Open the recordset with the appropriate query
            .Open(query, cn, 1, 2)

            ' Check if the recordset is empty
            If .EOF Then Exit Sub

            ' Clear the ListView before adding new items
            lvadmin.Items.Clear()

            ' Loop through the recordset and add items to the ListView
            While Not .EOF
                ' Create a new ListViewItem for the current record
                Dim rec As New ListViewItem(.Fields("ADM_USER").Value.ToString())
                rec.SubItems.Add(.Fields("ADM_LAST").Value.ToString())
                rec.SubItems.Add(.Fields("ADM_FIRST").Value.ToString())
                rec.SubItems.Add(.Fields("ADM_EMAIL").Value.ToString())

                ' Add the new item (rec) to the ListView
                lvadmin.Items.Add(rec)

                .MoveNext()
            End While

            ' Close the recordset
            .Close()
        End With
    End Sub


    ' Handle selected admin data from the ListView
    Private Sub lvadmin_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lvadmin.SelectedIndexChanged
        If lvadmin.SelectedItems.Count > 0 Then
            Dim selectedUser As String = lvadmin.SelectedItems(0).Text

            rs = New ADODB.Recordset
            With rs
                If .State = 1 Then .Close()
                .Open("SELECT * FROM tbl-adminregister WHERE ADM_USER='" & selectedUser & "'", cn, 1, 2)
                If .EOF = False Then

                    txtaduser.Text = .Fields("ADM_USER").Value
                    txtadfirst.Text = .Fields("ADM_FIRST").Value
                    txtadmid.Text = .Fields("ADM_MIDDLE").Value
                    txtadpass.Text = .Fields("ADM_PASS").Value
                    txtademail.Text = .Fields("ADM_EMAIL").Value


                    txtadcon.Text = .Fields("ADM_CONTACT").Value

                    dtpbirth.Value = .Fields("ADM_BIRTHDATE").Value
                    txthome.Text = .Fields("ADM_HOME").Value
                    txtstreet.Text = .Fields("ADM_STREET").Value
                    txtbrgy.Text = .Fields("ADM_BRGY").Value
                    txtmunicity.Text = .Fields("ADM_MUNICITY").Value
                    txtprov.Text = .Fields("ADM_PROV").Value
                    txtzip.Text = .Fields("ADM_ZIP").Value
                End If
                .Close()
            End With
        End If
    End Sub

    ' ComboBox Filter Selection Changed
    Private Sub cbofilter_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbofilter.SelectedIndexChanged
        ' You can implement any additional logic here based on the filter selection
    End Sub

    Private Sub txtbrgy_TextChanged(sender As Object, e As EventArgs) Handles txtbrgy.TextChanged

    End Sub
End Class