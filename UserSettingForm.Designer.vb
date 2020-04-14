<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class UserSettingForm
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
        Me.TBESDuplicateBasePath = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.BTSave = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'TBESDuplicateBasePath
        '
        Me.TBESDuplicateBasePath.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.TBESDuplicateBasePath.Location = New System.Drawing.Point(134, 28)
        Me.TBESDuplicateBasePath.Name = "TBESDuplicateBasePath"
        Me.TBESDuplicateBasePath.Size = New System.Drawing.Size(488, 26)
        Me.TBESDuplicateBasePath.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Label1.Font = New System.Drawing.Font("B Traffic", 9.75!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Label1.Location = New System.Drawing.Point(628, 31)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(160, 21)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "مسیر ذخیره بک آپ امکان سنجی"
        '
        'BTSave
        '
        Me.BTSave.Font = New System.Drawing.Font("B Traffic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.BTSave.Location = New System.Drawing.Point(26, 25)
        Me.BTSave.Name = "BTSave"
        Me.BTSave.Size = New System.Drawing.Size(102, 33)
        Me.BTSave.TabIndex = 2
        Me.BTSave.Text = "ذخیره تنظیمات"
        Me.BTSave.UseVisualStyleBackColor = True
        '
        'UserSettingForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 83)
        Me.Controls.Add(Me.BTSave)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TBESDuplicateBasePath)
        Me.Name = "UserSettingForm"
        Me.Text = "UserSettingForm"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents TBESDuplicateBasePath As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents BTSave As Button
End Class
