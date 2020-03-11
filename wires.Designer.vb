<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class wires
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(wires))
        Me.TabPage3 = New System.Windows.Forms.TabPage()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.TBProductName = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.TBWireDiameter = New System.Windows.Forms.TextBox()
        Me.TBOD = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TBL0 = New System.Windows.Forms.TextBox()
        Me.BTSearch = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TBNt = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.TBProductID = New System.Windows.Forms.TextBox()
        Me.BTModify = New System.Windows.Forms.Button()
        Me.BTNewProduct = New System.Windows.Forms.Button()
        Me.BTUpdateInventory = New System.Windows.Forms.Button()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabPage3
        '
        Me.TabPage3.BackColor = System.Drawing.SystemColors.Control
        Me.TabPage3.Location = New System.Drawing.Point(4, 30)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Size = New System.Drawing.Size(1064, 493)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "لیست سفارشات"
        '
        'TabPage1
        '
        Me.TabPage1.BackColor = System.Drawing.SystemColors.Control
        Me.TabPage1.Controls.Add(Me.GroupBox1)
        Me.TabPage1.Controls.Add(Me.DataGridView1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 30)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(1064, 493)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "موجودی مفتول"
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.BackgroundColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.MenuHighlight
        DataGridViewCellStyle1.Font = New System.Drawing.Font("B Traffic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView1.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle2.Font = New System.Drawing.Font("B Traffic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DataGridView1.DefaultCellStyle = DataGridViewCellStyle2
        Me.DataGridView1.GridColor = System.Drawing.SystemColors.ControlDarkDark
        Me.DataGridView1.Location = New System.Drawing.Point(2, 5)
        Me.DataGridView1.MultiSelect = False
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.MenuHighlight
        DataGridViewCellStyle3.Font = New System.Drawing.Font("B Traffic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView1.RowHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.DataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView1.Size = New System.Drawing.Size(1049, 372)
        Me.DataGridView1.TabIndex = 19
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.BTUpdateInventory)
        Me.GroupBox1.Controls.Add(Me.BTNewProduct)
        Me.GroupBox1.Controls.Add(Me.BTModify)
        Me.GroupBox1.Controls.Add(Me.TBProductID)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.TBNt)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.BTSearch)
        Me.GroupBox1.Controls.Add(Me.TBL0)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.TBOD)
        Me.GroupBox1.Controls.Add(Me.TBWireDiameter)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.TBProductName)
        Me.GroupBox1.Font = New System.Drawing.Font("B Traffic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(6, 384)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.GroupBox1.Size = New System.Drawing.Size(1045, 106)
        Me.GroupBox1.TabIndex = 20
        Me.GroupBox1.TabStop = False
        '
        'TBProductName
        '
        Me.TBProductName.Location = New System.Drawing.Point(732, 24)
        Me.TBProductName.Margin = New System.Windows.Forms.Padding(3, 1, 3, 1)
        Me.TBProductName.Name = "TBProductName"
        Me.TBProductName.Size = New System.Drawing.Size(236, 29)
        Me.TBProductName.TabIndex = 1
        Me.TBProductName.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("B Traffic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Label3.Location = New System.Drawing.Point(363, 25)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(76, 26)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "قطر خارجی"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("B Traffic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Label4.Location = New System.Drawing.Point(224, 25)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(68, 26)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "طول آزاد"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("B Traffic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Label5.Location = New System.Drawing.Point(80, 25)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(71, 26)
        Me.Label5.TabIndex = 8
        Me.Label5.Text = "تعداد حلقه"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TBWireDiameter
        '
        Me.TBWireDiameter.Location = New System.Drawing.Point(442, 24)
        Me.TBWireDiameter.Margin = New System.Windows.Forms.Padding(3, 1, 3, 1)
        Me.TBWireDiameter.Name = "TBWireDiameter"
        Me.TBWireDiameter.Size = New System.Drawing.Size(70, 29)
        Me.TBWireDiameter.TabIndex = 3
        Me.TBWireDiameter.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TBOD
        '
        Me.TBOD.Location = New System.Drawing.Point(295, 24)
        Me.TBOD.Margin = New System.Windows.Forms.Padding(3, 1, 3, 1)
        Me.TBOD.Name = "TBOD"
        Me.TBOD.Size = New System.Drawing.Size(70, 29)
        Me.TBOD.TabIndex = 4
        Me.TBOD.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("B Traffic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Label1.Location = New System.Drawing.Point(968, 25)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(76, 26)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "نام محصول"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TBL0
        '
        Me.TBL0.Location = New System.Drawing.Point(155, 24)
        Me.TBL0.Margin = New System.Windows.Forms.Padding(3, 1, 3, 1)
        Me.TBL0.Name = "TBL0"
        Me.TBL0.Size = New System.Drawing.Size(70, 29)
        Me.TBL0.TabIndex = 5
        Me.TBL0.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'BTSearch
        '
        Me.BTSearch.Font = New System.Drawing.Font("B Traffic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.BTSearch.Location = New System.Drawing.Point(620, 66)
        Me.BTSearch.Margin = New System.Windows.Forms.Padding(1)
        Me.BTSearch.Name = "BTSearch"
        Me.BTSearch.Size = New System.Drawing.Size(63, 33)
        Me.BTSearch.TabIndex = 7
        Me.BTSearch.Text = "جستجو"
        Me.BTSearch.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("B Traffic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Label2.Location = New System.Drawing.Point(512, 25)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(74, 26)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "قطر مفتول"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TBNt
        '
        Me.TBNt.Location = New System.Drawing.Point(8, 24)
        Me.TBNt.Margin = New System.Windows.Forms.Padding(3, 1, 3, 1)
        Me.TBNt.Name = "TBNt"
        Me.TBNt.Size = New System.Drawing.Size(70, 29)
        Me.TBNt.TabIndex = 6
        Me.TBNt.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("B Traffic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Label6.Location = New System.Drawing.Point(681, 25)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(49, 26)
        Me.Label6.TabIndex = 13
        Me.Label6.Text = "کد کالا"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TBProductID
        '
        Me.TBProductID.Location = New System.Drawing.Point(585, 24)
        Me.TBProductID.Margin = New System.Windows.Forms.Padding(3, 1, 3, 1)
        Me.TBProductID.Name = "TBProductID"
        Me.TBProductID.Size = New System.Drawing.Size(98, 29)
        Me.TBProductID.TabIndex = 2
        Me.TBProductID.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'BTModify
        '
        Me.BTModify.Font = New System.Drawing.Font("B Traffic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.BTModify.Location = New System.Drawing.Point(554, 66)
        Me.BTModify.Margin = New System.Windows.Forms.Padding(1)
        Me.BTModify.Name = "BTModify"
        Me.BTModify.Size = New System.Drawing.Size(64, 33)
        Me.BTModify.TabIndex = 8
        Me.BTModify.Text = "ویرایش"
        Me.BTModify.UseVisualStyleBackColor = True
        '
        'BTNewProduct
        '
        Me.BTNewProduct.Font = New System.Drawing.Font("B Traffic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.BTNewProduct.Location = New System.Drawing.Point(447, 66)
        Me.BTNewProduct.Margin = New System.Windows.Forms.Padding(1)
        Me.BTNewProduct.Name = "BTNewProduct"
        Me.BTNewProduct.Size = New System.Drawing.Size(105, 33)
        Me.BTNewProduct.TabIndex = 9
        Me.BTNewProduct.Text = "محصول جدید"
        Me.BTNewProduct.UseVisualStyleBackColor = True
        '
        'BTUpdateInventory
        '
        Me.BTUpdateInventory.Font = New System.Drawing.Font("B Traffic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.BTUpdateInventory.Location = New System.Drawing.Point(8, 66)
        Me.BTUpdateInventory.Margin = New System.Windows.Forms.Padding(1)
        Me.BTUpdateInventory.Name = "BTUpdateInventory"
        Me.BTUpdateInventory.Size = New System.Drawing.Size(149, 33)
        Me.BTUpdateInventory.TabIndex = 10
        Me.BTUpdateInventory.Text = "به‌روزرسانی موجودی"
        Me.BTUpdateInventory.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Font = New System.Drawing.Font("B Traffic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.TabControl1.Location = New System.Drawing.Point(0, 2)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.TabControl1.RightToLeftLayout = True
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(1072, 527)
        Me.TabControl1.TabIndex = 1
        '
        'wires
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1073, 541)
        Me.Controls.Add(Me.TabControl1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "wires"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "مواد اولیه"
        Me.TabPage1.ResumeLayout(False)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TabPage3 As TabPage
    Friend WithEvents TabPage1 As TabPage
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents BTUpdateInventory As Button
    Friend WithEvents BTNewProduct As Button
    Friend WithEvents BTModify As Button
    Friend WithEvents TBProductID As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents TBNt As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents BTSearch As Button
    Friend WithEvents TBL0 As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents TBOD As TextBox
    Friend WithEvents TBWireDiameter As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents TBProductName As TextBox
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents TabControl1 As TabControl
End Class
