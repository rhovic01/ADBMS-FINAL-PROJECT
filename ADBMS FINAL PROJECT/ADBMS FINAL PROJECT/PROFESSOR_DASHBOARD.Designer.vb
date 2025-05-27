<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class PROFESSOR_DASHBOARD
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
        Me.Label30 = New System.Windows.Forms.Label()
        Me.btnprofupdate = New System.Windows.Forms.Button()
        Me.btnprofexit = New System.Windows.Forms.Button()
        Me.btnprofdel = New System.Windows.Forms.Button()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.btnprofsave = New System.Windows.Forms.Button()
        Me.btnprofback = New System.Windows.Forms.Button()
        Me.btnprofbatch = New System.Windows.Forms.Button()
        Me.txtgradeid = New System.Windows.Forms.TextBox()
        Me.txtstudid = New System.Windows.Forms.TextBox()
        Me.txtstudname = New System.Windows.Forms.TextBox()
        Me.txtstudgpa = New System.Windows.Forms.TextBox()
        Me.txtstudsy = New System.Windows.Forms.TextBox()
        Me.cbostudcourse = New System.Windows.Forms.ComboBox()
        Me.txtsubjcode = New System.Windows.Forms.TextBox()
        Me.txtsubjtitle = New System.Windows.Forms.TextBox()
        Me.txtstudsem = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.txtstudsection = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.txtstudprof = New System.Windows.Forms.TextBox()
        Me.txtstudremarks = New System.Windows.Forms.TextBox()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.lvprofessor = New System.Windows.Forms.ListView()
        Me.lvstudid = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.lvstudname = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.lvstudgpa = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.lvstudremarks = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.SuspendLayout()
        '
        'Label30
        '
        Me.Label30.AutoSize = True
        Me.Label30.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!)
        Me.Label30.Location = New System.Drawing.Point(51, 175)
        Me.Label30.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label30.Name = "Label30"
        Me.Label30.Size = New System.Drawing.Size(79, 17)
        Me.Label30.TabIndex = 59
        Me.Label30.Text = "REMARKS:"
        '
        'btnprofupdate
        '
        Me.btnprofupdate.Location = New System.Drawing.Point(786, 299)
        Me.btnprofupdate.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnprofupdate.Name = "btnprofupdate"
        Me.btnprofupdate.Size = New System.Drawing.Size(161, 46)
        Me.btnprofupdate.TabIndex = 53
        Me.btnprofupdate.Text = "UPDATE"
        Me.btnprofupdate.UseVisualStyleBackColor = True
        '
        'btnprofexit
        '
        Me.btnprofexit.Location = New System.Drawing.Point(786, 397)
        Me.btnprofexit.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnprofexit.Name = "btnprofexit"
        Me.btnprofexit.Size = New System.Drawing.Size(161, 46)
        Me.btnprofexit.TabIndex = 52
        Me.btnprofexit.Text = "EXIT"
        Me.btnprofexit.UseVisualStyleBackColor = True
        '
        'btnprofdel
        '
        Me.btnprofdel.Location = New System.Drawing.Point(786, 349)
        Me.btnprofdel.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnprofdel.Name = "btnprofdel"
        Me.btnprofdel.Size = New System.Drawing.Size(161, 46)
        Me.btnprofdel.TabIndex = 51
        Me.btnprofdel.Text = "DELETE"
        Me.btnprofdel.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(51, 39)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(78, 17)
        Me.Label1.TabIndex = 26
        Me.Label1.Text = "GRADE ID:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(52, 243)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(70, 17)
        Me.Label2.TabIndex = 27
        Me.Label2.Text = "COURSE:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(51, 208)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(111, 17)
        Me.Label3.TabIndex = 28
        Me.Label3.Text = "SCHOOL YEAR:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(52, 144)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(41, 17)
        Me.Label4.TabIndex = 29
        Me.Label4.Text = "GPA:"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(52, 109)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(51, 17)
        Me.Label5.TabIndex = 30
        Me.Label5.Text = "NAME:"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(51, 73)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(95, 17)
        Me.Label6.TabIndex = 31
        Me.Label6.Text = "STUDENT ID:"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(544, 39)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(117, 17)
        Me.Label7.TabIndex = 32
        Me.Label7.Text = "SUBJECT CODE:"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(544, 73)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(116, 17)
        Me.Label8.TabIndex = 33
        Me.Label8.Text = "SUBJECT TITLE:"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(544, 109)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(87, 17)
        Me.Label9.TabIndex = 34
        Me.Label9.Text = "SEMESTER:"
        '
        'btnprofsave
        '
        Me.btnprofsave.Location = New System.Drawing.Point(786, 252)
        Me.btnprofsave.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnprofsave.Name = "btnprofsave"
        Me.btnprofsave.Size = New System.Drawing.Size(161, 46)
        Me.btnprofsave.TabIndex = 35
        Me.btnprofsave.Text = "SAVE"
        Me.btnprofsave.UseVisualStyleBackColor = True
        '
        'btnprofback
        '
        Me.btnprofback.Location = New System.Drawing.Point(786, 209)
        Me.btnprofback.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnprofback.Name = "btnprofback"
        Me.btnprofback.Size = New System.Drawing.Size(161, 42)
        Me.btnprofback.TabIndex = 36
        Me.btnprofback.Text = "BACK"
        Me.btnprofback.UseVisualStyleBackColor = True
        '
        'btnprofbatch
        '
        Me.btnprofbatch.Location = New System.Drawing.Point(433, 230)
        Me.btnprofbatch.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnprofbatch.Name = "btnprofbatch"
        Me.btnprofbatch.Size = New System.Drawing.Size(161, 46)
        Me.btnprofbatch.TabIndex = 37
        Me.btnprofbatch.Text = "BATCH UPLOAD"
        Me.btnprofbatch.UseVisualStyleBackColor = True
        '
        'txtgradeid
        '
        Me.txtgradeid.Location = New System.Drawing.Point(170, 39)
        Me.txtgradeid.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtgradeid.Name = "txtgradeid"
        Me.txtgradeid.ReadOnly = True
        Me.txtgradeid.Size = New System.Drawing.Size(208, 22)
        Me.txtgradeid.TabIndex = 38
        '
        'txtstudid
        '
        Me.txtstudid.Location = New System.Drawing.Point(170, 70)
        Me.txtstudid.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtstudid.Name = "txtstudid"
        Me.txtstudid.ReadOnly = True
        Me.txtstudid.Size = New System.Drawing.Size(208, 22)
        Me.txtstudid.TabIndex = 39
        '
        'txtstudname
        '
        Me.txtstudname.Location = New System.Drawing.Point(170, 104)
        Me.txtstudname.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtstudname.Name = "txtstudname"
        Me.txtstudname.Size = New System.Drawing.Size(208, 22)
        Me.txtstudname.TabIndex = 40
        '
        'txtstudgpa
        '
        Me.txtstudgpa.Location = New System.Drawing.Point(170, 141)
        Me.txtstudgpa.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtstudgpa.Name = "txtstudgpa"
        Me.txtstudgpa.Size = New System.Drawing.Size(208, 22)
        Me.txtstudgpa.TabIndex = 41
        '
        'txtstudsy
        '
        Me.txtstudsy.Location = New System.Drawing.Point(170, 205)
        Me.txtstudsy.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtstudsy.Name = "txtstudsy"
        Me.txtstudsy.Size = New System.Drawing.Size(208, 22)
        Me.txtstudsy.TabIndex = 42
        '
        'cbostudcourse
        '
        Me.cbostudcourse.FormattingEnabled = True
        Me.cbostudcourse.Items.AddRange(New Object() {"1st", "2nd", "3rd", "4th"})
        Me.cbostudcourse.Location = New System.Drawing.Point(170, 240)
        Me.cbostudcourse.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.cbostudcourse.Name = "cbostudcourse"
        Me.cbostudcourse.Size = New System.Drawing.Size(200, 24)
        Me.cbostudcourse.TabIndex = 43
        '
        'txtsubjcode
        '
        Me.txtsubjcode.Location = New System.Drawing.Point(667, 36)
        Me.txtsubjcode.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtsubjcode.Name = "txtsubjcode"
        Me.txtsubjcode.Size = New System.Drawing.Size(208, 22)
        Me.txtsubjcode.TabIndex = 44
        '
        'txtsubjtitle
        '
        Me.txtsubjtitle.Location = New System.Drawing.Point(667, 75)
        Me.txtsubjtitle.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtsubjtitle.Name = "txtsubjtitle"
        Me.txtsubjtitle.Size = New System.Drawing.Size(208, 22)
        Me.txtsubjtitle.TabIndex = 45
        '
        'txtstudsem
        '
        Me.txtstudsem.Location = New System.Drawing.Point(667, 111)
        Me.txtstudsem.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtstudsem.Name = "txtstudsem"
        Me.txtstudsem.Size = New System.Drawing.Size(208, 22)
        Me.txtstudsem.TabIndex = 46
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(52, 276)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(72, 17)
        Me.Label10.TabIndex = 47
        Me.Label10.Text = "SECTION:"
        '
        'txtstudsection
        '
        Me.txtstudsection.Location = New System.Drawing.Point(170, 276)
        Me.txtstudsection.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtstudsection.Name = "txtstudsection"
        Me.txtstudsection.Size = New System.Drawing.Size(208, 22)
        Me.txtstudsection.TabIndex = 48
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(544, 146)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(102, 17)
        Me.Label11.TabIndex = 49
        Me.Label11.Text = "INSTRUCTOR:"
        '
        'txtstudprof
        '
        Me.txtstudprof.Location = New System.Drawing.Point(667, 148)
        Me.txtstudprof.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtstudprof.Name = "txtstudprof"
        Me.txtstudprof.Size = New System.Drawing.Size(208, 22)
        Me.txtstudprof.TabIndex = 50
        '
        'txtstudremarks
        '
        Me.txtstudremarks.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!)
        Me.txtstudremarks.Location = New System.Drawing.Point(172, 174)
        Me.txtstudremarks.Margin = New System.Windows.Forms.Padding(4)
        Me.txtstudremarks.Name = "txtstudremarks"
        Me.txtstudremarks.Size = New System.Drawing.Size(206, 22)
        Me.txtstudremarks.TabIndex = 60
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'lvprofessor
        '
        Me.lvprofessor.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.lvstudid, Me.lvstudname, Me.lvstudgpa, Me.lvstudremarks})
        Me.lvprofessor.FullRowSelect = True
        Me.lvprofessor.HideSelection = False
        Me.lvprofessor.Location = New System.Drawing.Point(12, 317)
        Me.lvprofessor.Name = "lvprofessor"
        Me.lvprofessor.Size = New System.Drawing.Size(768, 154)
        Me.lvprofessor.TabIndex = 61
        Me.lvprofessor.UseCompatibleStateImageBehavior = False
        Me.lvprofessor.View = System.Windows.Forms.View.Details
        '
        'lvstudid
        '
        Me.lvstudid.Text = "STUDENT ID"
        Me.lvstudid.Width = 181
        '
        'lvstudname
        '
        Me.lvstudname.Text = "NAME"
        Me.lvstudname.Width = 175
        '
        'lvstudgpa
        '
        Me.lvstudgpa.Text = "GPA"
        Me.lvstudgpa.Width = 144
        '
        'lvstudremarks
        '
        Me.lvstudremarks.Text = "Remarks"
        '
        'PROFESSOR_DASHBOARD
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(979, 483)
        Me.Controls.Add(Me.lvprofessor)
        Me.Controls.Add(Me.txtstudremarks)
        Me.Controls.Add(Me.Label30)
        Me.Controls.Add(Me.btnprofupdate)
        Me.Controls.Add(Me.btnprofexit)
        Me.Controls.Add(Me.btnprofdel)
        Me.Controls.Add(Me.txtstudprof)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.txtstudsection)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.txtstudsem)
        Me.Controls.Add(Me.txtsubjtitle)
        Me.Controls.Add(Me.txtsubjcode)
        Me.Controls.Add(Me.cbostudcourse)
        Me.Controls.Add(Me.txtstudsy)
        Me.Controls.Add(Me.txtstudgpa)
        Me.Controls.Add(Me.txtstudname)
        Me.Controls.Add(Me.txtstudid)
        Me.Controls.Add(Me.txtgradeid)
        Me.Controls.Add(Me.btnprofbatch)
        Me.Controls.Add(Me.btnprofback)
        Me.Controls.Add(Me.btnprofsave)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Name = "PROFESSOR_DASHBOARD"
        Me.Text = "WSS"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label30 As Label
    Friend WithEvents btnprofupdate As Button
    Friend WithEvents btnprofexit As Button
    Friend WithEvents btnprofdel As Button
    Friend WithEvents SaveFileDialog1 As SaveFileDialog
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents Label8 As Label
    Friend WithEvents Label9 As Label
    Friend WithEvents btnprofsave As Button
    Friend WithEvents btnprofback As Button
    Friend WithEvents btnprofbatch As Button
    Friend WithEvents txtgradeid As TextBox
    Friend WithEvents txtstudid As TextBox
    Friend WithEvents txtstudname As TextBox
    Friend WithEvents txtstudgpa As TextBox
    Friend WithEvents txtstudsy As TextBox
    Friend WithEvents cbostudcourse As ComboBox
    Friend WithEvents txtsubjcode As TextBox
    Friend WithEvents txtsubjtitle As TextBox
    Friend WithEvents txtstudsem As TextBox
    Friend WithEvents Label10 As Label
    Friend WithEvents txtstudsection As TextBox
    Friend WithEvents Label11 As Label
    Friend WithEvents txtstudprof As TextBox
    Friend WithEvents txtstudremarks As TextBox
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents lvprofessor As ListView
    Friend WithEvents lvstudid As ColumnHeader
    Friend WithEvents lvstudname As ColumnHeader
    Friend WithEvents lvstudgpa As ColumnHeader
    Friend WithEvents lvstudremarks As ColumnHeader
End Class
