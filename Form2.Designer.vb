<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form2
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
        Me.PatternComboBox = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Openbtn = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txbPattern = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'PatternComboBox
        '
        Me.PatternComboBox.AllowDrop = True
        Me.PatternComboBox.FormattingEnabled = True
        Me.PatternComboBox.Items.AddRange(New Object() {"NHK", "Sumitomo", "Miharu"})
        Me.PatternComboBox.Location = New System.Drawing.Point(12, 27)
        Me.PatternComboBox.Name = "PatternComboBox"
        Me.PatternComboBox.Size = New System.Drawing.Size(210, 23)
        Me.PatternComboBox.TabIndex = 12
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(49, 15)
        Me.Label1.TabIndex = 13
        Me.Label1.Text = "Settings"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(147, 139)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 14
        Me.Button1.Text = "Done"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Openbtn
        '
        Me.Openbtn.Location = New System.Drawing.Point(12, 139)
        Me.Openbtn.Name = "Openbtn"
        Me.Openbtn.Size = New System.Drawing.Size(94, 23)
        Me.Openbtn.TabIndex = 15
        Me.Openbtn.Text = "Select Folder"
        Me.Openbtn.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 79)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(66, 15)
        Me.Label2.TabIndex = 17
        Me.Label2.Text = "File Pattern"
        '
        'txbPattern
        '
        Me.txbPattern.Location = New System.Drawing.Point(12, 97)
        Me.txbPattern.Name = "txbPattern"
        Me.txbPattern.Size = New System.Drawing.Size(210, 23)
        Me.txbPattern.TabIndex = 18
        '
        'Form2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(234, 174)
        Me.Controls.Add(Me.txbPattern)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Openbtn)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.PatternComboBox)
        Me.Name = "Form2"
        Me.Text = "Form2"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents PatternComboBox As ComboBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Button1 As Button
    Friend WithEvents Openbtn As Button
    Friend WithEvents Label2 As Label
    Friend WithEvents txbPattern As TextBox
End Class
