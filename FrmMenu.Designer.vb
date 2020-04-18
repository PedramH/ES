<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FrmMenu
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmMenu))
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.Menu_userProfile = New System.Windows.Forms.ToolStripMenuItem()
        Me.M_user_info = New System.Windows.Forms.ToolStripMenuItem()
        Me.M_user_setting = New System.Windows.Forms.ToolStripMenuItem()
        Me.M_user_logout = New System.Windows.Forms.ToolStripMenuItem()
        Me.Menu_emkansanji = New System.Windows.Forms.ToolStripMenuItem()
        Me.M_emkansanji_orders = New System.Windows.Forms.ToolStripMenuItem()
        Me.M_emkansanji_new = New System.Windows.Forms.ToolStripMenuItem()
        Me.Menu_production = New System.Windows.Forms.ToolStripMenuItem()
        Me.M_production_orders = New System.Windows.Forms.ToolStripMenuItem()
        Me.M_prodcution_wireInventory = New System.Windows.Forms.ToolStripMenuItem()
        Me.M_production_data = New System.Windows.Forms.ToolStripMenuItem()
        Me.M_production_shippment = New System.Windows.Forms.ToolStripMenuItem()
        Me.M_production_plan = New System.Windows.Forms.ToolStripMenuItem()
        Me.Menu_customers = New System.Windows.Forms.ToolStripMenuItem()
        Me.M_customers_list = New System.Windows.Forms.ToolStripMenuItem()
        Me.M_customer_new = New System.Windows.Forms.ToolStripMenuItem()
        Me.Menu_products = New System.Windows.Forms.ToolStripMenuItem()
        Me.M_products_list = New System.Windows.Forms.ToolStripMenuItem()
        Me.M_products_new = New System.Windows.Forms.ToolStripMenuItem()
        Me.Menu_warehouse = New System.Windows.Forms.ToolStripMenuItem()
        Me.M_warehouse_mandrels = New System.Windows.Forms.ToolStripMenuItem()
        Me.M_warehouse_wires = New System.Windows.Forms.ToolStripMenuItem()
        Me.M_warehouse_productCode = New System.Windows.Forms.ToolStripMenuItem()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.BTTest = New System.Windows.Forms.Button()
        Me.MenuStrip1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.BackColor = System.Drawing.SystemColors.Control
        Me.MenuStrip1.Font = New System.Drawing.Font("B Traffic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.Menu_userProfile, Me.Menu_emkansanji, Me.Menu_production, Me.Menu_customers, Me.Menu_products, Me.Menu_warehouse})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Padding = New System.Windows.Forms.Padding(8, 3, 0, 3)
        Me.MenuStrip1.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.MenuStrip1.Size = New System.Drawing.Size(784, 31)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'Menu_userProfile
        '
        Me.Menu_userProfile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.M_user_info, Me.M_user_setting, Me.M_user_logout})
        Me.Menu_userProfile.Font = New System.Drawing.Font("B Traffic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Menu_userProfile.Name = "Menu_userProfile"
        Me.Menu_userProfile.Size = New System.Drawing.Size(95, 25)
        Me.Menu_userProfile.Text = "اطلاعات کاربری"
        '
        'M_user_info
        '
        Me.M_user_info.Enabled = False
        Me.M_user_info.Name = "M_user_info"
        Me.M_user_info.Size = New System.Drawing.Size(144, 26)
        Me.M_user_info.Text = "اطلاعات کاربر"
        '
        'M_user_setting
        '
        Me.M_user_setting.Name = "M_user_setting"
        Me.M_user_setting.Size = New System.Drawing.Size(144, 26)
        Me.M_user_setting.Text = "تنظیمات"
        '
        'M_user_logout
        '
        Me.M_user_logout.Name = "M_user_logout"
        Me.M_user_logout.Size = New System.Drawing.Size(144, 26)
        Me.M_user_logout.Text = "خروج"
        '
        'Menu_emkansanji
        '
        Me.Menu_emkansanji.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.M_emkansanji_orders, Me.M_emkansanji_new})
        Me.Menu_emkansanji.Name = "Menu_emkansanji"
        Me.Menu_emkansanji.Size = New System.Drawing.Size(76, 25)
        Me.Menu_emkansanji.Text = "امکان سنجی"
        '
        'M_emkansanji_orders
        '
        Me.M_emkansanji_orders.Name = "M_emkansanji_orders"
        Me.M_emkansanji_orders.Size = New System.Drawing.Size(180, 26)
        Me.M_emkansanji_orders.Text = "لیست سفارشات"
        '
        'M_emkansanji_new
        '
        Me.M_emkansanji_new.Name = "M_emkansanji_new"
        Me.M_emkansanji_new.Size = New System.Drawing.Size(180, 26)
        Me.M_emkansanji_new.Text = "امکان سنجی جدید"
        '
        'Menu_production
        '
        Me.Menu_production.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.M_production_orders, Me.M_prodcution_wireInventory, Me.M_production_data, Me.M_production_shippment, Me.M_production_plan})
        Me.Menu_production.Name = "Menu_production"
        Me.Menu_production.Size = New System.Drawing.Size(45, 25)
        Me.Menu_production.Text = "تولید"
        '
        'M_production_orders
        '
        Me.M_production_orders.Name = "M_production_orders"
        Me.M_production_orders.Size = New System.Drawing.Size(180, 26)
        Me.M_production_orders.Text = "لیست سفارشات"
        '
        'M_prodcution_wireInventory
        '
        Me.M_prodcution_wireInventory.Name = "M_prodcution_wireInventory"
        Me.M_prodcution_wireInventory.Size = New System.Drawing.Size(180, 26)
        Me.M_prodcution_wireInventory.Text = "موجودی مواد اولیه"
        '
        'M_production_data
        '
        Me.M_production_data.Name = "M_production_data"
        Me.M_production_data.Size = New System.Drawing.Size(180, 26)
        Me.M_production_data.Text = "آمار و موجودی تولید"
        '
        'M_production_shippment
        '
        Me.M_production_shippment.Name = "M_production_shippment"
        Me.M_production_shippment.Size = New System.Drawing.Size(180, 26)
        Me.M_production_shippment.Text = "ارسال محصول"
        '
        'M_production_plan
        '
        Me.M_production_plan.Name = "M_production_plan"
        Me.M_production_plan.Size = New System.Drawing.Size(180, 26)
        Me.M_production_plan.Text = "برنامه تولید"
        '
        'Menu_customers
        '
        Me.Menu_customers.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.M_customers_list, Me.M_customer_new})
        Me.Menu_customers.Name = "Menu_customers"
        Me.Menu_customers.Size = New System.Drawing.Size(103, 25)
        Me.Menu_customers.Text = "مدیریت مشتریان"
        '
        'M_customers_list
        '
        Me.M_customers_list.Name = "M_customers_list"
        Me.M_customers_list.Size = New System.Drawing.Size(180, 26)
        Me.M_customers_list.Text = "اطلاعت مشتریان"
        '
        'M_customer_new
        '
        Me.M_customer_new.Name = "M_customer_new"
        Me.M_customer_new.Size = New System.Drawing.Size(180, 26)
        Me.M_customer_new.Text = "ثبت مشتری جدید"
        '
        'Menu_products
        '
        Me.Menu_products.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.M_products_list, Me.M_products_new})
        Me.Menu_products.Name = "Menu_products"
        Me.Menu_products.Size = New System.Drawing.Size(109, 25)
        Me.Menu_products.Text = "مدیریت محصولات"
        '
        'M_products_list
        '
        Me.M_products_list.Name = "M_products_list"
        Me.M_products_list.Size = New System.Drawing.Size(165, 26)
        Me.M_products_list.Text = "اطلاعات محصولات"
        '
        'M_products_new
        '
        Me.M_products_new.Name = "M_products_new"
        Me.M_products_new.Size = New System.Drawing.Size(165, 26)
        Me.M_products_new.Text = "ثبت محصول جدید"
        '
        'Menu_warehouse
        '
        Me.Menu_warehouse.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.M_warehouse_mandrels, Me.M_warehouse_wires, Me.M_warehouse_productCode})
        Me.Menu_warehouse.Name = "Menu_warehouse"
        Me.Menu_warehouse.Size = New System.Drawing.Size(40, 25)
        Me.Menu_warehouse.Text = "انبار"
        '
        'M_warehouse_mandrels
        '
        Me.M_warehouse_mandrels.Name = "M_warehouse_mandrels"
        Me.M_warehouse_mandrels.Size = New System.Drawing.Size(234, 26)
        Me.M_warehouse_mandrels.Text = "لیست مندرل ها"
        '
        'M_warehouse_wires
        '
        Me.M_warehouse_wires.Name = "M_warehouse_wires"
        Me.M_warehouse_wires.Size = New System.Drawing.Size(234, 26)
        Me.M_warehouse_wires.Text = "موجودی مواد اولیه"
        '
        'M_warehouse_productCode
        '
        Me.M_warehouse_productCode.Name = "M_warehouse_productCode"
        Me.M_warehouse_productCode.Size = New System.Drawing.Size(234, 26)
        Me.M_warehouse_productCode.Text = "صدور کد کالا برای محصول جدید"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.BTTest)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 31)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(784, 530)
        Me.Panel1.TabIndex = 1
        '
        'BTTest
        '
        Me.BTTest.Location = New System.Drawing.Point(33, 64)
        Me.BTTest.Name = "BTTest"
        Me.BTTest.Size = New System.Drawing.Size(82, 50)
        Me.BTTest.TabIndex = 0
        Me.BTTest.Text = "test"
        Me.BTTest.UseVisualStyleBackColor = True
        Me.BTTest.Visible = False
        '
        'FrmMenu
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 21.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(784, 561)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Font = New System.Drawing.Font("B Traffic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Name = "FrmMenu"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "EnergySaz ERP"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents Menu_userProfile As ToolStripMenuItem
    Friend WithEvents M_user_setting As ToolStripMenuItem
    Friend WithEvents M_user_logout As ToolStripMenuItem
    Friend WithEvents Menu_emkansanji As ToolStripMenuItem
    Friend WithEvents M_emkansanji_orders As ToolStripMenuItem
    Friend WithEvents M_emkansanji_new As ToolStripMenuItem
    Friend WithEvents Menu_production As ToolStripMenuItem
    Friend WithEvents M_production_plan As ToolStripMenuItem
    Friend WithEvents M_production_orders As ToolStripMenuItem
    Friend WithEvents M_production_data As ToolStripMenuItem
    Friend WithEvents M_prodcution_wireInventory As ToolStripMenuItem
    Friend WithEvents M_production_shippment As ToolStripMenuItem
    Friend WithEvents Menu_customers As ToolStripMenuItem
    Friend WithEvents M_customers_list As ToolStripMenuItem
    Friend WithEvents Menu_products As ToolStripMenuItem
    Friend WithEvents M_products_list As ToolStripMenuItem
    Friend WithEvents M_customer_new As ToolStripMenuItem
    Friend WithEvents M_products_new As ToolStripMenuItem
    Friend WithEvents M_user_info As ToolStripMenuItem
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Menu_warehouse As ToolStripMenuItem
    Friend WithEvents M_warehouse_wires As ToolStripMenuItem
    Friend WithEvents M_warehouse_productCode As ToolStripMenuItem
    Friend WithEvents M_warehouse_mandrels As ToolStripMenuItem
    Friend WithEvents BTTest As Button
End Class
