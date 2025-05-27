<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnprofessor = New System.Windows.Forms.Button()
        Me.btnadmin = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Arial Narrow", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(198, 36)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 22)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "OPTION"
        '
        'btnprofessor
        '
        Me.btnprofessor.Location = New System.Drawing.Point(158, 149)
        Me.btnprofessor.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnprofessor.Name = "btnprofessor"
        Me.btnprofessor.Size = New System.Drawing.Size(169, 58)
        Me.btnprofessor.TabIndex = 4
        Me.btnprofessor.Text = "PROFESSOR"
        Me.btnprofessor.UseVisualStyleBackColor = True
        '
        'btnadmin
        '
        Me.btnadmin.Location = New System.Drawing.Point(158, 80)
        Me.btnadmin.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnadmin.Name = "btnadmin"
        Me.btnadmin.Size = New System.Drawing.Size(169, 58)
        Me.btnadmin.TabIndex = 3
        Me.btnadmin.Text = "ADMIN"
        Me.btnadmin.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(512, 287)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnprofessor)
        Me.Controls.Add(Me.btnadmin)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents btnprofessor As Button
    Friend WithEvents btnadmin As Button
End Class
