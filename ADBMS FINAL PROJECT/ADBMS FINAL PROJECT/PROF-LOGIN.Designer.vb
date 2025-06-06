<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PROF_LOGIN
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.btnproflogin = New System.Windows.Forms.Button()
        Me.txtprofpass = New System.Windows.Forms.TextBox()
        Me.txtprofuser = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cbprofremember = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Arial Narrow", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(163, 41)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(171, 24)
        Me.Label3.TabIndex = 17
        Me.Label3.Text = "PROFESSOR LOGIN"
        '
        'btnproflogin
        '
        Me.btnproflogin.Location = New System.Drawing.Point(58, 218)
        Me.btnproflogin.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnproflogin.Name = "btnproflogin"
        Me.btnproflogin.Size = New System.Drawing.Size(363, 48)
        Me.btnproflogin.TabIndex = 16
        Me.btnproflogin.Text = "LOGIN"
        Me.btnproflogin.UseVisualStyleBackColor = True
        '
        'txtprofpass
        '
        Me.txtprofpass.Location = New System.Drawing.Point(168, 140)
        Me.txtprofpass.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtprofpass.Name = "txtprofpass"
        Me.txtprofpass.Size = New System.Drawing.Size(252, 26)
        Me.txtprofpass.TabIndex = 15
        '
        'txtprofuser
        '
        Me.txtprofuser.Location = New System.Drawing.Point(169, 90)
        Me.txtprofuser.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtprofuser.Name = "txtprofuser"
        Me.txtprofuser.Size = New System.Drawing.Size(252, 26)
        Me.txtprofuser.TabIndex = 14
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Arial Narrow", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(52, 141)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(111, 24)
        Me.Label2.TabIndex = 13
        Me.Label2.Text = "PASSWORD:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Arial Narrow", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(54, 91)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(109, 24)
        Me.Label1.TabIndex = 12
        Me.Label1.Text = "USERNAME:"
        '
        'cbprofremember
        '
        Me.cbprofremember.AutoSize = True
        Me.cbprofremember.Location = New System.Drawing.Point(168, 174)
        Me.cbprofremember.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cbprofremember.Name = "cbprofremember"
        Me.cbprofremember.Size = New System.Drawing.Size(183, 24)
        Me.cbprofremember.TabIndex = 18
        Me.cbprofremember.Text = "SHOW PASSWORD"
        Me.cbprofremember.UseVisualStyleBackColor = True
        '
        'PROF_LOGIN
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(476, 300)
        Me.Controls.Add(Me.cbprofremember)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.btnproflogin)
        Me.Controls.Add(Me.txtprofpass)
        Me.Controls.Add(Me.txtprofuser)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Name = "PROF_LOGIN"
        Me.Text = "PROF_LOGIN"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label3 As Label
    Friend WithEvents btnproflogin As Button
    Friend WithEvents txtprofpass As TextBox
    Friend WithEvents txtprofuser As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents cbprofremember As CheckBox
End Class
