<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class ADMINDASHBOARD
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.txtsearch = New System.Windows.Forms.TextBox()
        Me.txtzip = New System.Windows.Forms.TextBox()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.txtprov = New System.Windows.Forms.TextBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.txtmunicity = New System.Windows.Forms.TextBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.txtbrgy = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.txtstreet = New System.Windows.Forms.TextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.txthome = New System.Windows.Forms.TextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.cln_username = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.cln_firstname = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.cln_middlename = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ADM_LAST = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.cln_email = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.cln_address = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.Label19 = New System.Windows.Forms.Label()
        Me.cbofilter = New System.Windows.Forms.ComboBox()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.lvadmin = New System.Windows.Forms.ListView()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.dtpbirth = New System.Windows.Forms.DateTimePicker()
        Me.btnexit = New System.Windows.Forms.Button()
        Me.btndele = New System.Windows.Forms.Button()
        Me.btnsave = New System.Windows.Forms.Button()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.txtadfirst = New System.Windows.Forms.TextBox()
        Me.txtadmid = New System.Windows.Forms.TextBox()
        Me.txtadlast = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtadcon = New System.Windows.Forms.TextBox()
        Me.txtademail = New System.Windows.Forms.TextBox()
        Me.txtadpass = New System.Windows.Forms.TextBox()
        Me.txtaduser = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cbShowPass = New System.Windows.Forms.CheckBox()
        Me.btnadclear = New System.Windows.Forms.Button()
        Me.btnadupdate = New System.Windows.Forms.Button()
        Me.btnadback = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtsearch
        '
        Me.txtsearch.Location = New System.Drawing.Point(424, 406)
        Me.txtsearch.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtsearch.Name = "txtsearch"
        Me.txtsearch.Size = New System.Drawing.Size(234, 22)
        Me.txtsearch.TabIndex = 58
        '
        'txtzip
        '
        Me.txtzip.Location = New System.Drawing.Point(102, 163)
        Me.txtzip.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtzip.Name = "txtzip"
        Me.txtzip.Size = New System.Drawing.Size(234, 22)
        Me.txtzip.TabIndex = 27
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Font = New System.Drawing.Font("Arial Narrow", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label17.Location = New System.Drawing.Point(7, 162)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(81, 22)
        Me.Label17.TabIndex = 26
        Me.Label17.Text = "ZIP CODE:"
        '
        'txtprov
        '
        Me.txtprov.Location = New System.Drawing.Point(102, 130)
        Me.txtprov.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtprov.Name = "txtprov"
        Me.txtprov.Size = New System.Drawing.Size(234, 22)
        Me.txtprov.TabIndex = 16
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Font = New System.Drawing.Font("Arial Narrow", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.Location = New System.Drawing.Point(9, 130)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(87, 22)
        Me.Label16.TabIndex = 15
        Me.Label16.Text = "PROVINCE:"
        '
        'txtmunicity
        '
        Me.txtmunicity.Location = New System.Drawing.Point(179, 99)
        Me.txtmunicity.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtmunicity.Name = "txtmunicity"
        Me.txtmunicity.Size = New System.Drawing.Size(234, 22)
        Me.txtmunicity.TabIndex = 14
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Font = New System.Drawing.Font("Arial Narrow", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.Location = New System.Drawing.Point(8, 98)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(150, 22)
        Me.Label15.TabIndex = 13
        Me.Label15.Text = "MUNICIPALITY/CITY:"
        '
        'txtbrgy
        '
        Me.txtbrgy.Location = New System.Drawing.Point(110, 72)
        Me.txtbrgy.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtbrgy.Name = "txtbrgy"
        Me.txtbrgy.Size = New System.Drawing.Size(234, 22)
        Me.txtbrgy.TabIndex = 12
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Font = New System.Drawing.Font("Arial Narrow", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.Location = New System.Drawing.Point(8, 72)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(94, 22)
        Me.Label14.TabIndex = 11
        Me.Label14.Text = "BARANGAY:"
        '
        'txtstreet
        '
        Me.txtstreet.Location = New System.Drawing.Point(110, 46)
        Me.txtstreet.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtstreet.Name = "txtstreet"
        Me.txtstreet.Size = New System.Drawing.Size(234, 22)
        Me.txtstreet.TabIndex = 10
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Font = New System.Drawing.Font("Arial Narrow", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.Location = New System.Drawing.Point(8, 46)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(70, 22)
        Me.Label13.TabIndex = 9
        Me.Label13.Text = "STREET:"
        '
        'txthome
        '
        Me.txthome.Location = New System.Drawing.Point(110, 20)
        Me.txthome.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txthome.Name = "txthome"
        Me.txthome.Size = New System.Drawing.Size(234, 22)
        Me.txthome.TabIndex = 8
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Font = New System.Drawing.Font("Arial Narrow", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.Location = New System.Drawing.Point(8, 20)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(58, 22)
        Me.Label12.TabIndex = 7
        Me.Label12.Text = "HOME:"
        '
        'cln_username
        '
        Me.cln_username.Text = "USER NAME"
        Me.cln_username.Width = 181
        '
        'cln_firstname
        '
        Me.cln_firstname.Text = "FIRST NAME"
        Me.cln_firstname.Width = 175
        '
        'cln_middlename
        '
        Me.cln_middlename.Text = "MIDDLE NAME"
        Me.cln_middlename.Width = 144
        '
        'ADM_LAST
        '
        Me.ADM_LAST.Text = "LAST NAME"
        Me.ADM_LAST.Width = 158
        '
        'cln_email
        '
        Me.cln_email.Text = "EMAIL"
        Me.cln_email.Width = 201
        '
        'cln_address
        '
        Me.cln_address.Text = "ADDRESS"
        Me.cln_address.Width = 459
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Font = New System.Drawing.Font("Arial Narrow", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label19.Location = New System.Drawing.Point(365, 406)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(63, 22)
        Me.Label19.TabIndex = 57
        Me.Label19.Text = "Search:"
        '
        'cbofilter
        '
        Me.cbofilter.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.FileSystem
        Me.cbofilter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbofilter.FormattingEnabled = True
        Me.cbofilter.Items.AddRange(New Object() {"USER NAME", "FIRST NAME", "MIDDLE NAME", "LAST NAME", "EMAIL", "ADDRESS"})
        Me.cbofilter.Location = New System.Drawing.Point(136, 406)
        Me.cbofilter.Name = "cbofilter"
        Me.cbofilter.Size = New System.Drawing.Size(205, 24)
        Me.cbofilter.TabIndex = 56
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Font = New System.Drawing.Font("Arial Narrow", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label18.Location = New System.Drawing.Point(57, 406)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(73, 22)
        Me.Label18.TabIndex = 55
        Me.Label18.Text = "Filter By:"
        '
        'lvadmin
        '
        Me.lvadmin.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.cln_username, Me.cln_firstname, Me.cln_middlename, Me.ADM_LAST, Me.cln_email, Me.cln_address})
        Me.lvadmin.FullRowSelect = True
        Me.lvadmin.Location = New System.Drawing.Point(58, 433)
        Me.lvadmin.Name = "lvadmin"
        Me.lvadmin.Size = New System.Drawing.Size(1164, 229)
        Me.lvadmin.TabIndex = 54
        Me.lvadmin.UseCompatibleStateImageBehavior = False
        Me.lvadmin.View = System.Windows.Forms.View.Details
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txtzip)
        Me.GroupBox1.Controls.Add(Me.Label17)
        Me.GroupBox1.Controls.Add(Me.txtprov)
        Me.GroupBox1.Controls.Add(Me.Label16)
        Me.GroupBox1.Controls.Add(Me.txtmunicity)
        Me.GroupBox1.Controls.Add(Me.Label15)
        Me.GroupBox1.Controls.Add(Me.txtbrgy)
        Me.GroupBox1.Controls.Add(Me.Label14)
        Me.GroupBox1.Controls.Add(Me.txtstreet)
        Me.GroupBox1.Controls.Add(Me.Label13)
        Me.GroupBox1.Controls.Add(Me.txthome)
        Me.GroupBox1.Controls.Add(Me.Label12)
        Me.GroupBox1.Location = New System.Drawing.Point(462, 101)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(588, 205)
        Me.GroupBox1.TabIndex = 53
        Me.GroupBox1.TabStop = False
        '
        'dtpbirth
        '
        Me.dtpbirth.Location = New System.Drawing.Point(180, 338)
        Me.dtpbirth.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.dtpbirth.Name = "dtpbirth"
        Me.dtpbirth.Size = New System.Drawing.Size(234, 22)
        Me.dtpbirth.TabIndex = 52
        '
        'btnexit
        '
        Me.btnexit.Location = New System.Drawing.Point(1055, 284)
        Me.btnexit.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnexit.Name = "btnexit"
        Me.btnexit.Size = New System.Drawing.Size(156, 49)
        Me.btnexit.TabIndex = 51
        Me.btnexit.Text = "EXIT"
        Me.btnexit.UseVisualStyleBackColor = True
        '
        'btndele
        '
        Me.btndele.Location = New System.Drawing.Point(1055, 338)
        Me.btndele.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btndele.Name = "btndele"
        Me.btndele.Size = New System.Drawing.Size(156, 44)
        Me.btndele.TabIndex = 50
        Me.btndele.Text = "DELETE"
        Me.btndele.UseVisualStyleBackColor = True
        '
        'btnsave
        '
        Me.btnsave.Location = New System.Drawing.Point(1056, 110)
        Me.btnsave.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnsave.Name = "btnsave"
        Me.btnsave.Size = New System.Drawing.Size(156, 46)
        Me.btnsave.TabIndex = 49
        Me.btnsave.Text = "SAVE"
        Me.btnsave.UseVisualStyleBackColor = True
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("Arial Narrow", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(456, 61)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(135, 35)
        Me.Label11.TabIndex = 48
        Me.Label11.Text = "ADDRESS"
        '
        'txtadfirst
        '
        Me.txtadfirst.Location = New System.Drawing.Point(180, 232)
        Me.txtadfirst.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtadfirst.Name = "txtadfirst"
        Me.txtadfirst.Size = New System.Drawing.Size(234, 22)
        Me.txtadfirst.TabIndex = 47
        '
        'txtadmid
        '
        Me.txtadmid.Location = New System.Drawing.Point(180, 264)
        Me.txtadmid.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtadmid.Name = "txtadmid"
        Me.txtadmid.Size = New System.Drawing.Size(234, 22)
        Me.txtadmid.TabIndex = 46
        '
        'txtadlast
        '
        Me.txtadlast.Location = New System.Drawing.Point(180, 298)
        Me.txtadlast.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtadlast.Name = "txtadlast"
        Me.txtadlast.Size = New System.Drawing.Size(234, 22)
        Me.txtadlast.TabIndex = 45
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Arial Narrow", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(57, 266)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(116, 22)
        Me.Label10.TabIndex = 44
        Me.Label10.Text = "MIDDLE NAME:"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Arial Narrow", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(57, 298)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(94, 22)
        Me.Label9.TabIndex = 43
        Me.Label9.Text = "LASTNAME:"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Arial Narrow", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(58, 329)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(86, 22)
        Me.Label8.TabIndex = 42
        Me.Label8.Text = "BIRTHDAY:"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Arial Narrow", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(57, 232)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(98, 22)
        Me.Label7.TabIndex = 41
        Me.Label7.Text = "FIRSTNAME:"
        '
        'txtadcon
        '
        Me.txtadcon.Location = New System.Drawing.Point(161, 157)
        Me.txtadcon.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtadcon.Name = "txtadcon"
        Me.txtadcon.Size = New System.Drawing.Size(234, 22)
        Me.txtadcon.TabIndex = 40
        '
        'txtademail
        '
        Me.txtademail.Location = New System.Drawing.Point(161, 130)
        Me.txtademail.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtademail.Name = "txtademail"
        Me.txtademail.Size = New System.Drawing.Size(234, 22)
        Me.txtademail.TabIndex = 39
        '
        'txtadpass
        '
        Me.txtadpass.Location = New System.Drawing.Point(159, 75)
        Me.txtadpass.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtadpass.Name = "txtadpass"
        Me.txtadpass.Size = New System.Drawing.Size(234, 22)
        Me.txtadpass.TabIndex = 38
        '
        'txtaduser
        '
        Me.txtaduser.Location = New System.Drawing.Point(159, 48)
        Me.txtaduser.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtaduser.Name = "txtaduser"
        Me.txtaduser.Size = New System.Drawing.Size(234, 22)
        Me.txtaduser.TabIndex = 37
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Arial Narrow", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(57, 132)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(60, 22)
        Me.Label6.TabIndex = 36
        Me.Label6.Text = "EMAIL:"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Arial Narrow", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(58, 160)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(79, 22)
        Me.Label5.TabIndex = 35
        Me.Label5.Text = "CONTACT"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Arial Narrow", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(57, 191)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(0, 22)
        Me.Label4.TabIndex = 34
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Arial Narrow", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(57, 223)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(0, 22)
        Me.Label3.TabIndex = 33
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Arial Narrow", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(57, 75)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(96, 22)
        Me.Label2.TabIndex = 32
        Me.Label2.Text = "PASSWORD:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Arial Narrow", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(57, 48)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(95, 22)
        Me.Label1.TabIndex = 31
        Me.Label1.Text = "USERNAME:"
        '
        'cbShowPass
        '
        Me.cbShowPass.AutoSize = True
        Me.cbShowPass.Location = New System.Drawing.Point(159, 105)
        Me.cbShowPass.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.cbShowPass.Name = "cbShowPass"
        Me.cbShowPass.Size = New System.Drawing.Size(157, 21)
        Me.cbShowPass.TabIndex = 62
        Me.cbShowPass.Text = "SHOW PASSWORD"
        Me.cbShowPass.UseVisualStyleBackColor = True
        '
        'btnadclear
        '
        Me.btnadclear.Location = New System.Drawing.Point(1054, 230)
        Me.btnadclear.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnadclear.Name = "btnadclear"
        Me.btnadclear.Size = New System.Drawing.Size(157, 47)
        Me.btnadclear.TabIndex = 63
        Me.btnadclear.Text = "CLEAR"
        Me.btnadclear.UseVisualStyleBackColor = True
        '
        'btnadupdate
        '
        Me.btnadupdate.Location = New System.Drawing.Point(1055, 170)
        Me.btnadupdate.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnadupdate.Name = "btnadupdate"
        Me.btnadupdate.Size = New System.Drawing.Size(157, 47)
        Me.btnadupdate.TabIndex = 64
        Me.btnadupdate.Text = "UPDATE"
        Me.btnadupdate.UseVisualStyleBackColor = True
        '
        'btnadback
        '
        Me.btnadback.Location = New System.Drawing.Point(865, 338)
        Me.btnadback.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnadback.Name = "btnadback"
        Me.btnadback.Size = New System.Drawing.Size(156, 44)
        Me.btnadback.TabIndex = 65
        Me.btnadback.Text = "BACK"
        Me.btnadback.UseVisualStyleBackColor = True
        '
        'ADMINDASHBOARD
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1286, 674)
        Me.Controls.Add(Me.btnadback)
        Me.Controls.Add(Me.btnadupdate)
        Me.Controls.Add(Me.btnadclear)
        Me.Controls.Add(Me.cbShowPass)
        Me.Controls.Add(Me.txtsearch)
        Me.Controls.Add(Me.Label19)
        Me.Controls.Add(Me.cbofilter)
        Me.Controls.Add(Me.Label18)
        Me.Controls.Add(Me.lvadmin)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.dtpbirth)
        Me.Controls.Add(Me.btnexit)
        Me.Controls.Add(Me.btndele)
        Me.Controls.Add(Me.btnsave)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.txtadfirst)
        Me.Controls.Add(Me.txtadmid)
        Me.Controls.Add(Me.txtadlast)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.txtadcon)
        Me.Controls.Add(Me.txtademail)
        Me.Controls.Add(Me.txtadpass)
        Me.Controls.Add(Me.txtaduser)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Name = "ADMINDASHBOARD"
        Me.Text = "ADMINDASHBOARD"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txtsearch As TextBox
    Friend WithEvents txtzip As TextBox
    Friend WithEvents Label17 As Label
    Friend WithEvents txtprov As TextBox
    Friend WithEvents Label16 As Label
    Friend WithEvents txtmunicity As TextBox
    Friend WithEvents Label15 As Label
    Friend WithEvents txtbrgy As TextBox
    Friend WithEvents Label14 As Label
    Friend WithEvents txtstreet As TextBox
    Friend WithEvents Label13 As Label
    Friend WithEvents txthome As TextBox
    Friend WithEvents Label12 As Label
    Friend WithEvents cln_username As ColumnHeader
    Friend WithEvents cln_firstname As ColumnHeader
    Friend WithEvents cln_middlename As ColumnHeader
    Friend WithEvents ADM_LAST As ColumnHeader
    Friend WithEvents cln_email As ColumnHeader
    Friend WithEvents cln_address As ColumnHeader
    Friend WithEvents Label19 As Label
    Friend WithEvents cbofilter As ComboBox
    Friend WithEvents Label18 As Label
    Friend WithEvents lvadmin As ListView
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents dtpbirth As DateTimePicker
    Friend WithEvents btnexit As Button
    Friend WithEvents btndele As Button
    Friend WithEvents btnsave As Button
    Friend WithEvents Label11 As Label
    Friend WithEvents txtadfirst As TextBox
    Friend WithEvents txtadmid As TextBox
    Friend WithEvents txtadlast As TextBox
    Friend WithEvents Label10 As Label
    Friend WithEvents Label9 As Label
    Friend WithEvents Label8 As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents txtadcon As TextBox
    Friend WithEvents txtademail As TextBox
    Friend WithEvents txtadpass As TextBox
    Friend WithEvents txtaduser As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents Label1 As Label
    Friend WithEvents cbShowPass As CheckBox
    Friend WithEvents btnadclear As Button
    Friend WithEvents btnadupdate As Button
    Friend WithEvents btnadback As Button
End Class
