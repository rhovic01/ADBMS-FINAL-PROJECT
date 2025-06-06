<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form2
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
        Me.btnadminlogin = New System.Windows.Forms.Button()
        Me.txtadminpass = New System.Windows.Forms.TextBox()
        Me.txtadminuser = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        Me.cbremember = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Arial Narrow", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(156, 46)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(106, 22)
        Me.Label3.TabIndex = 18
        Me.Label3.Text = "ADMIN LOGIN"
        '
        'btnadminlogin
        '
        Me.btnadminlogin.Location = New System.Drawing.Point(41, 176)
        Me.btnadminlogin.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnadminlogin.Name = "btnadminlogin"
        Me.btnadminlogin.Size = New System.Drawing.Size(323, 38)
        Me.btnadminlogin.TabIndex = 17
        Me.btnadminlogin.Text = "LOGIN"
        Me.btnadminlogin.UseVisualStyleBackColor = True
        '
        'txtadminpass
        '
        Me.txtadminpass.Location = New System.Drawing.Point(138, 123)
        Me.txtadminpass.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtadminpass.Name = "txtadminpass"
        Me.txtadminpass.Size = New System.Drawing.Size(224, 22)
        Me.txtadminpass.TabIndex = 16
        '
        'txtadminuser
        '
        Me.txtadminuser.Location = New System.Drawing.Point(138, 83)
        Me.txtadminuser.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtadminuser.Name = "txtadminuser"
        Me.txtadminuser.Size = New System.Drawing.Size(224, 22)
        Me.txtadminuser.TabIndex = 15
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Arial Narrow", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(35, 125)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(96, 22)
        Me.Label2.TabIndex = 14
        Me.Label2.Text = "PASSWORD:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Arial Narrow", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(37, 85)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(95, 22)
        Me.Label1.TabIndex = 13
        Me.Label1.Text = "USERNAME:"
        '
        'cbremember
        '
        Me.cbremember.AutoSize = True
        Me.cbremember.Location = New System.Drawing.Point(138, 150)
        Me.cbremember.Name = "cbremember"
        Me.cbremember.Size = New System.Drawing.Size(157, 21)
        Me.cbremember.TabIndex = 19
        Me.cbremember.Text = "SHOW PASSOWRD"
        Me.cbremember.UseVisualStyleBackColor = True
        '
        'Form2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(422, 266)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.btnadminlogin)
        Me.Controls.Add(Me.txtadminpass)
        Me.Controls.Add(Me.txtadminuser)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cbremember)
        Me.Name = "Form2"
        Me.Text = "Form2"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label3 As Label
    Friend WithEvents btnadminlogin As Button
    Friend WithEvents txtadminpass As TextBox
    Friend WithEvents txtadminuser As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents cbremember As CheckBox
End Class
