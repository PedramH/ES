<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class wires
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(wires))
        Me.TabPage_Wires = New System.Windows.Forms.TabPage()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.TableLayoutPanel2 = New System.Windows.Forms.TableLayoutPanel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TBWireCode = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TBWireDiameter = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.TBWireLengthMin = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.BTUpdateInventory = New System.Windows.Forms.Button()
        Me.BTCheckOrders = New System.Windows.Forms.Button()
        Me.BTSelectWire = New System.Windows.Forms.Button()
        Me.BTShowAllWires = New System.Windows.Forms.Button()
        Me.CBWireType = New System.Windows.Forms.ComboBox()
        Me.TabPage_Orders = New System.Windows.Forms.TabPage()
        Me.TableLayoutPanel3 = New System.Windows.Forms.TableLayoutPanel()
        Me.DataGridView2 = New System.Windows.Forms.DataGridView()
        Me.TableLayoutPanel4 = New System.Windows.Forms.TableLayoutPanel()
        Me.TBWireCodeOrderSearch = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.BTCheckOrder = New System.Windows.Forms.Button()
        Me.BTShowAllOrders = New System.Windows.Forms.Button()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.TBCustomerNameOrderSearch = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.TBProductNameOrderSearch = New System.Windows.Forms.TextBox()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage_Wires.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TableLayoutPanel2.SuspendLayout()
        Me.TabPage_Orders.SuspendLayout()
        Me.TableLayoutPanel3.SuspendLayout()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TableLayoutPanel4.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabPage_Wires
        '
        Me.TabPage_Wires.BackColor = System.Drawing.SystemColors.Control
        Me.TabPage_Wires.Controls.Add(Me.TableLayoutPanel1)
        Me.TabPage_Wires.Location = New System.Drawing.Point(4, 30)
        Me.TabPage_Wires.Name = "TabPage_Wires"
        Me.TabPage_Wires.Size = New System.Drawing.Size(1065, 507)
        Me.TabPage_Wires.TabIndex = 2
        Me.TabPage_Wires.Text = "موجودی مفتول"
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 1
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.DataGridView1, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.TableLayoutPanel2, 0, 1)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 2
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 80.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(1065, 507)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'DataGridView1
        '
        Me.DataGridView1.AccessibleRole = System.Windows.Forms.AccessibleRole.None
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView1.BackgroundColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.MenuHighlight
        DataGridViewCellStyle1.Font = New System.Drawing.Font("B Traffic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView1.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle2.Font = New System.Drawing.Font("B Traffic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DataGridView1.DefaultCellStyle = DataGridViewCellStyle2
        Me.DataGridView1.GridColor = System.Drawing.SystemColors.ControlDarkDark
        Me.DataGridView1.Location = New System.Drawing.Point(3, 3)
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
        Me.DataGridView1.Size = New System.Drawing.Size(1059, 399)
        Me.DataGridView1.TabIndex = 20
        '
        'TableLayoutPanel2
        '
        Me.TableLayoutPanel2.ColumnCount = 10
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 12.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 13.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 9.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 6.0!))
        Me.TableLayoutPanel2.Controls.Add(Me.Label1, 0, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.TBWireCode, 1, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.Label2, 2, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.TBWireDiameter, 3, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.Label6, 4, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.TBWireLengthMin, 5, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.Label3, 6, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.BTUpdateInventory, 8, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.BTCheckOrders, 8, 1)
        Me.TableLayoutPanel2.Controls.Add(Me.BTSelectWire, 7, 1)
        Me.TableLayoutPanel2.Controls.Add(Me.BTShowAllWires, 0, 1)
        Me.TableLayoutPanel2.Controls.Add(Me.CBWireType, 7, 0)
        Me.TableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel2.Location = New System.Drawing.Point(3, 408)
        Me.TableLayoutPanel2.Name = "TableLayoutPanel2"
        Me.TableLayoutPanel2.RowCount = 2
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel2.Size = New System.Drawing.Size(1059, 96)
        Me.TableLayoutPanel2.TabIndex = 21
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label1.Font = New System.Drawing.Font("B Traffic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Label1.Location = New System.Drawing.Point(935, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(121, 48)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "کد کالا"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TBWireCode
        '
        Me.TBWireCode.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TBWireCode.Location = New System.Drawing.Point(830, 9)
        Me.TBWireCode.Margin = New System.Windows.Forms.Padding(3, 1, 3, 1)
        Me.TBWireCode.Name = "TBWireCode"
        Me.TBWireCode.Size = New System.Drawing.Size(99, 29)
        Me.TBWireCode.TabIndex = 6
        Me.TBWireCode.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label2.Font = New System.Drawing.Font("B Traffic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Label2.Location = New System.Drawing.Point(725, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(99, 48)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "قطر مفتول"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TBWireDiameter
        '
        Me.TBWireDiameter.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TBWireDiameter.Location = New System.Drawing.Point(620, 9)
        Me.TBWireDiameter.Margin = New System.Windows.Forms.Padding(3, 1, 3, 1)
        Me.TBWireDiameter.Name = "TBWireDiameter"
        Me.TBWireDiameter.Size = New System.Drawing.Size(99, 29)
        Me.TBWireDiameter.TabIndex = 8
        Me.TBWireDiameter.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label6.Font = New System.Drawing.Font("B Traffic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Label6.Location = New System.Drawing.Point(483, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(131, 48)
        Me.Label6.TabIndex = 17
        Me.Label6.Text = "حداقل طول مفتول"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TBWireLengthMin
        '
        Me.TBWireLengthMin.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TBWireLengthMin.Location = New System.Drawing.Point(378, 9)
        Me.TBWireLengthMin.Margin = New System.Windows.Forms.Padding(3, 1, 3, 1)
        Me.TBWireLengthMin.Name = "TBWireLengthMin"
        Me.TBWireLengthMin.Size = New System.Drawing.Size(99, 29)
        Me.TBWireLengthMin.TabIndex = 18
        Me.TBWireLengthMin.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label3.Font = New System.Drawing.Font("B Traffic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Label3.Location = New System.Drawing.Point(273, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(99, 48)
        Me.Label3.TabIndex = 19
        Me.Label3.Text = "نوع مفتول"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'BTUpdateInventory
        '
        Me.BTUpdateInventory.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel2.SetColumnSpan(Me.BTUpdateInventory, 2)
        Me.BTUpdateInventory.Font = New System.Drawing.Font("B Traffic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.BTUpdateInventory.Location = New System.Drawing.Point(1, 7)
        Me.BTUpdateInventory.Margin = New System.Windows.Forms.Padding(1)
        Me.BTUpdateInventory.Name = "BTUpdateInventory"
        Me.BTUpdateInventory.Size = New System.Drawing.Size(163, 33)
        Me.BTUpdateInventory.TabIndex = 21
        Me.BTUpdateInventory.Text = "به‌روزرسانی موجودی"
        Me.BTUpdateInventory.UseVisualStyleBackColor = True
        '
        'BTCheckOrders
        '
        Me.BTCheckOrders.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel2.SetColumnSpan(Me.BTCheckOrders, 2)
        Me.BTCheckOrders.Font = New System.Drawing.Font("B Traffic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.BTCheckOrders.Location = New System.Drawing.Point(1, 55)
        Me.BTCheckOrders.Margin = New System.Windows.Forms.Padding(1)
        Me.BTCheckOrders.Name = "BTCheckOrders"
        Me.BTCheckOrders.Size = New System.Drawing.Size(163, 33)
        Me.BTCheckOrders.TabIndex = 23
        Me.BTCheckOrders.Text = "جزییات رزرو"
        Me.BTCheckOrders.UseVisualStyleBackColor = True
        '
        'BTSelectWire
        '
        Me.BTSelectWire.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BTSelectWire.Font = New System.Drawing.Font("B Traffic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.BTSelectWire.Location = New System.Drawing.Point(166, 55)
        Me.BTSelectWire.Margin = New System.Windows.Forms.Padding(1)
        Me.BTSelectWire.Name = "BTSelectWire"
        Me.BTSelectWire.Size = New System.Drawing.Size(103, 33)
        Me.BTSelectWire.TabIndex = 22
        Me.BTSelectWire.Text = "انتخاب مفتول"
        Me.BTSelectWire.UseVisualStyleBackColor = True
        Me.BTSelectWire.Visible = False
        '
        'BTShowAllWires
        '
        Me.BTShowAllWires.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BTShowAllWires.Font = New System.Drawing.Font("B Traffic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.BTShowAllWires.Location = New System.Drawing.Point(933, 55)
        Me.BTShowAllWires.Margin = New System.Windows.Forms.Padding(1)
        Me.BTShowAllWires.Name = "BTShowAllWires"
        Me.BTShowAllWires.Size = New System.Drawing.Size(125, 33)
        Me.BTShowAllWires.TabIndex = 24
        Me.BTShowAllWires.Text = "نمایش همه"
        Me.BTShowAllWires.UseVisualStyleBackColor = True
        '
        'CBWireType
        '
        Me.CBWireType.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.CBWireType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CBWireType.FormattingEnabled = True
        Me.CBWireType.Items.AddRange(New Object() {"همه", "مفتول شاخه‌ای", "مفتول کویل"})
        Me.CBWireType.Location = New System.Drawing.Point(168, 13)
        Me.CBWireType.Name = "CBWireType"
        Me.CBWireType.Size = New System.Drawing.Size(99, 29)
        Me.CBWireType.TabIndex = 20
        '
        'TabPage_Orders
        '
        Me.TabPage_Orders.BackColor = System.Drawing.SystemColors.Control
        Me.TabPage_Orders.Controls.Add(Me.TableLayoutPanel3)
        Me.TabPage_Orders.Location = New System.Drawing.Point(4, 30)
        Me.TabPage_Orders.Name = "TabPage_Orders"
        Me.TabPage_Orders.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage_Orders.Size = New System.Drawing.Size(1065, 507)
        Me.TabPage_Orders.TabIndex = 0
        Me.TabPage_Orders.Text = "لیست سفارشات"
        '
        'TableLayoutPanel3
        '
        Me.TableLayoutPanel3.ColumnCount = 1
        Me.TableLayoutPanel3.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel3.Controls.Add(Me.DataGridView2, 0, 0)
        Me.TableLayoutPanel3.Controls.Add(Me.TableLayoutPanel4, 0, 1)
        Me.TableLayoutPanel3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel3.Location = New System.Drawing.Point(3, 3)
        Me.TableLayoutPanel3.Name = "TableLayoutPanel3"
        Me.TableLayoutPanel3.RowCount = 2
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 80.0!))
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20.0!))
        Me.TableLayoutPanel3.Size = New System.Drawing.Size(1059, 501)
        Me.TableLayoutPanel3.TabIndex = 0
        '
        'DataGridView2
        '
        Me.DataGridView2.AllowUserToAddRows = False
        Me.DataGridView2.AllowUserToDeleteRows = False
        Me.DataGridView2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView2.BackgroundColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.MenuHighlight
        DataGridViewCellStyle4.Font = New System.Drawing.Font("B Traffic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView2.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle4
        Me.DataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle5.Font = New System.Drawing.Font("B Traffic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        DataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DataGridView2.DefaultCellStyle = DataGridViewCellStyle5
        Me.DataGridView2.GridColor = System.Drawing.SystemColors.ControlDarkDark
        Me.DataGridView2.Location = New System.Drawing.Point(3, 3)
        Me.DataGridView2.MultiSelect = False
        Me.DataGridView2.Name = "DataGridView2"
        Me.DataGridView2.ReadOnly = True
        Me.DataGridView2.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle6.BackColor = System.Drawing.SystemColors.MenuHighlight
        DataGridViewCellStyle6.Font = New System.Drawing.Font("B Traffic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        DataGridViewCellStyle6.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView2.RowHeadersDefaultCellStyle = DataGridViewCellStyle6
        Me.DataGridView2.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView2.Size = New System.Drawing.Size(1053, 394)
        Me.DataGridView2.TabIndex = 21
        '
        'TableLayoutPanel4
        '
        Me.TableLayoutPanel4.ColumnCount = 8
        Me.TableLayoutPanel4.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 12.5!))
        Me.TableLayoutPanel4.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 12.5!))
        Me.TableLayoutPanel4.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 12.5!))
        Me.TableLayoutPanel4.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 12.5!))
        Me.TableLayoutPanel4.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 12.5!))
        Me.TableLayoutPanel4.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 12.5!))
        Me.TableLayoutPanel4.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 12.5!))
        Me.TableLayoutPanel4.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 12.5!))
        Me.TableLayoutPanel4.Controls.Add(Me.TBWireCodeOrderSearch, 0, 0)
        Me.TableLayoutPanel4.Controls.Add(Me.Label4, 0, 0)
        Me.TableLayoutPanel4.Controls.Add(Me.BTCheckOrder, 7, 0)
        Me.TableLayoutPanel4.Controls.Add(Me.BTShowAllOrders, 6, 0)
        Me.TableLayoutPanel4.Controls.Add(Me.Label5, 4, 0)
        Me.TableLayoutPanel4.Controls.Add(Me.TBCustomerNameOrderSearch, 5, 0)
        Me.TableLayoutPanel4.Controls.Add(Me.Label7, 2, 0)
        Me.TableLayoutPanel4.Controls.Add(Me.TBProductNameOrderSearch, 3, 0)
        Me.TableLayoutPanel4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel4.Location = New System.Drawing.Point(3, 403)
        Me.TableLayoutPanel4.Name = "TableLayoutPanel4"
        Me.TableLayoutPanel4.RowCount = 1
        Me.TableLayoutPanel4.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel4.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 95.0!))
        Me.TableLayoutPanel4.Size = New System.Drawing.Size(1053, 95)
        Me.TableLayoutPanel4.TabIndex = 22
        '
        'TBWireCodeOrderSearch
        '
        Me.TBWireCodeOrderSearch.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TBWireCodeOrderSearch.Location = New System.Drawing.Point(794, 33)
        Me.TBWireCodeOrderSearch.Margin = New System.Windows.Forms.Padding(3, 1, 3, 1)
        Me.TBWireCodeOrderSearch.Name = "TBWireCodeOrderSearch"
        Me.TBWireCodeOrderSearch.Size = New System.Drawing.Size(125, 29)
        Me.TBWireCodeOrderSearch.TabIndex = 7
        Me.TBWireCodeOrderSearch.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label4.Font = New System.Drawing.Font("B Traffic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Label4.Location = New System.Drawing.Point(925, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(125, 95)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "کد مفتول رزرو"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'BTCheckOrder
        '
        Me.BTCheckOrder.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BTCheckOrder.Font = New System.Drawing.Font("B Traffic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.BTCheckOrder.Location = New System.Drawing.Point(1, 31)
        Me.BTCheckOrder.Margin = New System.Windows.Forms.Padding(1)
        Me.BTCheckOrder.Name = "BTCheckOrder"
        Me.BTCheckOrder.Size = New System.Drawing.Size(134, 33)
        Me.BTCheckOrder.TabIndex = 22
        Me.BTCheckOrder.Text = "بررسی سفارش"
        Me.BTCheckOrder.UseVisualStyleBackColor = True
        '
        'BTShowAllOrders
        '
        Me.BTShowAllOrders.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BTShowAllOrders.Font = New System.Drawing.Font("B Traffic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.BTShowAllOrders.Location = New System.Drawing.Point(137, 31)
        Me.BTShowAllOrders.Margin = New System.Windows.Forms.Padding(1)
        Me.BTShowAllOrders.Name = "BTShowAllOrders"
        Me.BTShowAllOrders.Size = New System.Drawing.Size(129, 33)
        Me.BTShowAllOrders.TabIndex = 23
        Me.BTShowAllOrders.Text = "نمایش همه"
        Me.BTShowAllOrders.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label5.Font = New System.Drawing.Font("B Traffic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Label5.Location = New System.Drawing.Point(401, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(125, 95)
        Me.Label5.TabIndex = 24
        Me.Label5.Text = "نام مشتری"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TBCustomerNameOrderSearch
        '
        Me.TBCustomerNameOrderSearch.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TBCustomerNameOrderSearch.Location = New System.Drawing.Point(270, 33)
        Me.TBCustomerNameOrderSearch.Margin = New System.Windows.Forms.Padding(3, 1, 3, 1)
        Me.TBCustomerNameOrderSearch.Name = "TBCustomerNameOrderSearch"
        Me.TBCustomerNameOrderSearch.Size = New System.Drawing.Size(125, 29)
        Me.TBCustomerNameOrderSearch.TabIndex = 27
        Me.TBCustomerNameOrderSearch.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label7.Font = New System.Drawing.Font("B Traffic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Label7.Location = New System.Drawing.Point(663, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(125, 95)
        Me.Label7.TabIndex = 26
        Me.Label7.Text = "نام محصول"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TBProductNameOrderSearch
        '
        Me.TBProductNameOrderSearch.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TBProductNameOrderSearch.Location = New System.Drawing.Point(532, 33)
        Me.TBProductNameOrderSearch.Margin = New System.Windows.Forms.Padding(3, 1, 3, 1)
        Me.TBProductNameOrderSearch.Name = "TBProductNameOrderSearch"
        Me.TBProductNameOrderSearch.Size = New System.Drawing.Size(125, 29)
        Me.TBProductNameOrderSearch.TabIndex = 25
        Me.TBProductNameOrderSearch.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage_Wires)
        Me.TabControl1.Controls.Add(Me.TabPage_Orders)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Font = New System.Drawing.Font("B Traffic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.TabControl1.RightToLeftLayout = True
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(1073, 541)
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
        Me.TabPage_Wires.ResumeLayout(False)
        Me.TableLayoutPanel1.ResumeLayout(False)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TableLayoutPanel2.ResumeLayout(False)
        Me.TableLayoutPanel2.PerformLayout()
        Me.TabPage_Orders.ResumeLayout(False)
        Me.TableLayoutPanel3.ResumeLayout(False)
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TableLayoutPanel4.ResumeLayout(False)
        Me.TableLayoutPanel4.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TabPage_Wires As TabPage
    Friend WithEvents TabPage_Orders As TabPage
    Friend WithEvents TabControl1 As TabControl
    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents TableLayoutPanel2 As TableLayoutPanel
    Friend WithEvents Label1 As Label
    Friend WithEvents TBWireCode As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents TBWireDiameter As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents TBWireLengthMin As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents CBWireType As ComboBox
    Friend WithEvents BTUpdateInventory As Button
    Friend WithEvents BTSelectWire As Button
    Friend WithEvents TableLayoutPanel3 As TableLayoutPanel
    Friend WithEvents DataGridView2 As DataGridView
    Friend WithEvents TableLayoutPanel4 As TableLayoutPanel
    Friend WithEvents Label4 As Label
    Friend WithEvents TBWireCodeOrderSearch As TextBox
    Friend WithEvents BTCheckOrder As Button
    Friend WithEvents BTCheckOrders As Button
    Friend WithEvents BTShowAllWires As Button
    Friend WithEvents BTShowAllOrders As Button
    Friend WithEvents Label5 As Label
    Friend WithEvents TBCustomerNameOrderSearch As TextBox
    Friend WithEvents Label7 As Label
    Friend WithEvents TBProductNameOrderSearch As TextBox
End Class
