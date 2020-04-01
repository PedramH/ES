Imports System.Data.OleDb
Imports System.Configuration
Imports Excel = Microsoft.Office.Interop.Excel
Imports Npgsql
Imports System.Text.RegularExpressions


Public Class FrmNewEmkansanji
    Public spring_bs As New BindingSource
    Public customer_bs As New BindingSource

    Public CBA As New Collection '' An array of checkBoxes manufacturing process
    Public CBA_inspection As New Collection '' An array of checkBoxes for inspection

    Protected Overrides Function ProcessCmdKey(ByRef msg As System.Windows.Forms.Message,
                                           ByVal keyData As System.Windows.Forms.Keys) _
                                           As Boolean
        ' Sends Tab key everytime Enter is pressed INSIDE OF A TEXTBOX

        If msg.WParam.ToInt32() = CInt(Keys.Enter) AndAlso TypeOf Me.ActiveControl Is TextBox Then
            If Me.ActiveControl.Name <> "TBComment" Then
                SendKeys.Send("{Tab}")
                Return True
            End If
        End If

        Return MyBase.ProcessCmdKey(msg, keyData)
    End Function



    Private Async Sub mainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '' Add production process checkboxes to a collection for easier handling
        CBA.Add(CHSard)
        CBA.Add(CHGarm)
        CBA.Add(CHStress)
        CBA.Add(CHTemper)
        CBA.Add(CHShot)
        CBA.Add(CHTarak)
        CBA.Add(CHSet)
        CBA.Add(CHSang)
        CBA.Add(CHRang)
        CBA.Add(CHPelak)
        '' Add inspection process checkboxes to a collection for easier handling
        CBA_inspection.Add(CHForceTest)
        CBA_inspection.Add(CHAllInspection)
        CBA_inspection.Add(CHcustomerTolerance)
        CBA_inspection.Add(CHVerifyBeforeShipping)
        CBA_inspection.Add(CHCreepTest)
        CBA_inspection.Add(TBOtherInspection)


        'Loading Springs table into datagridview1 ----- this is the old version
        'Using cn As New OleDbConnection(connectionString)
        '    Using cmd As New OleDbCommand With {.Connection = cn}
        '        cmd.CommandText = "SELECT " & springDataBaseColumnNames & " FROM springDataBase;"

        '        Dim dt As New DataTable With {.TableName = "springDataBase"}
        '        Try
        '            cn.Open()
        '            Dim ds As New DataSet
        '            Dim springDBTable As New DataTable With {.TableName = "springDataBase"}
        '            ds.Tables.Add(springDBTable)
        '            ds.Load(cmd.ExecuteReader(), LoadOption.OverwriteChanges, springDBTable)
        '            spring_bs.DataSource = ds.Tables("springDataBase")
        '            DataGridView1.DataSource = ds.Tables("springDataBase")
        '            'DataGridView1.Columns(0).Visible = False
        '        Catch ex As Exception
        '            MsgBox("خطا در ارتباط با دیتابیس", vbCritical + vbMsgBoxRight, "خطا")
        '            Logger.LogFatal(ex.Message, ex)
        '        Finally
        '            cn.Close()
        '        End Try

        '    End Using
        'End Using

        'Loading Springs table into datagridview1
        Dim sql_command = "SELECT " & springDataBaseColumnNames & " FROM springDataBase;"
        Dim springdt As New DataTable
        Try
            springdt = Await Task(Of DataTable).Run(Function() LoadDataTable(sql_command))
            spring_bs.DataSource = springdt
            DataGridView1.DataSource = spring_bs
            'DataGridView1.Columns(0).Visible = False
        Catch ex As Exception
            MsgBox("خطا در ارتباط با دیتابیس" + Environment.NewLine + ex.Message, vbCritical + vbMsgBoxRight, "خطا")
            Logger.LogFatal("spring data base couldn't be loaded.", ex)
        End Try


        '' ------------------------------------------------------------------------------------------------------------------
        '' ------------------------------------------------------------------------------------------------------------------
        '' ------------------------------------------------------------------------------------------------------------------
        '' Loading Customers info into datagridView2 ------------ This is the old version
        'Using cn As New OleDbConnection(connectionString)
        '    Using cmd As New OleDbCommand With {.Connection = cn}

        '        cmd.CommandText = "SELECT " & customerDataBaseColumnNames & " FROM customers;"

        '        Dim dt As New DataTable With {.TableName = "customers"}
        '        Try
        '            cn.Open()
        '            Dim ds As New DataSet
        '            Dim customersDBTable As New DataTable With {.TableName = "customers"}
        '            ds.Tables.Add(customersDBTable)
        '            ds.Load(cmd.ExecuteReader(), LoadOption.OverwriteChanges, customersDBTable)
        '            DataGridView2.DataSource = ds.Tables("customers")
        '            DataGridView2.Columns(0).Visible = False
        '        Catch ex As Exception
        '            MsgBox("خطا در ارتباط با دیتابیس", vbCritical + vbMsgBoxRight, "خطا")
        '            Logger.LogFatal(ex.Message, ex)
        '        Finally
        '            cn.Close()
        '        End Try

        '    End Using
        'End Using

        '' Loading Customers info into datagridView2
        sql_command = "SELECT " & customerDataBaseColumnNames & " FROM customers;"
        Dim customerdt As New DataTable
        Try
            customerdt = Await Task(Of DataTable).Run(Function() LoadDataTable(sql_command))
            customer_bs.DataSource = customerdt
            DataGridView2.DataSource = customer_bs
            DataGridView2.Columns(0).Visible = False
        Catch ex As Exception
            MsgBox("خطا در ارتباط با دیتابیس" + Environment.NewLine + ex.Message, vbCritical + vbMsgBoxRight, "خطا")
            Logger.LogFatal(ex.Message, ex)
        End Try

    End Sub

    Private Async Sub BTSearch_Click(sender As Object, e As EventArgs) Handles BTSearch.Click
        'This subroutine searches the database based on textbox value

        '' This way it would find [فنر لول داخلی بوژی LSD1] with query [داخلی LSD] its just enough for the word to be inside the name
        ''      and what ever it is between them doesn't matter


        '' ------------------- THIS IS THE OLD WAY-------------------------------------------------------------------------
        'Dim searchTerm = TBProductName.Text.Replace(" ", "%")
        'Using cn As New OleDbConnection(connectionString)
        '    Using cmd As New OleDbCommand With {.Connection = cn}
        '        cmd.CommandText = "SELECT " & springDataBaseColumnNames & " FROM springDataBase WHERE " &
        '            "springDataBase.productName LIKE '%" & searchTerm & "%' AND" & '"springDataBase.productName LIKE '%" & TBProductName.Text & "%' AND" &
        '            " springDataBase.wireDiameter LIKE '%" & TBWireDiameter.Text & "%'" &
        '            " AND springDataBase.OD LIKE '%" & TBOD.Text & "%'" &
        '            " AND springDataBase.L0 LIKE '%" & TBL0.Text & "%'" &
        '            " AND springDataBase.Nt LIKE '%" & TBNt.Text & "%'" & " AND springDataBase.productID LIKE '%" & TBProductID.Text & "%'" & " ;"

        '        Dim dt As New DataTable With {.TableName = "springDataBase"}
        '        Try
        '            cn.Open()
        '            Dim ds As New DataSet
        '            Dim springDBTable As New DataTable With {.TableName = "springDataBase"}
        '            ds.Tables.Add(springDBTable)
        '            ds.Load(cmd.ExecuteReader(), LoadOption.OverwriteChanges, springDBTable)
        '            DataGridView1.DataSource = ds.Tables("springDataBase")
        '            spring_bs.DataSource = ds.Tables("springDataBase")
        '            cn.Close()
        '        Catch ex As Exception
        '            MsgBox("خطا در ارتباط با دیتابیس", vbCritical + vbMsgBoxRight, "خطا")
        '            Logger.LogFatal(ex.Message, ex)
        '        Finally
        '            cn.Close()
        '        End Try
        '    End Using
        'End Using

        Dim searchTerm = TBProductName.Text.Replace(" ", "%")
        Dim sql_command = "SELECT " & springDataBaseColumnNames & " FROM springDataBase WHERE " &
                    "springDataBase.productName LIKE '%" & searchTerm & "%' AND" & '"springDataBase.productName LIKE '%" & TBProductName.Text & "%' AND" &
                    " springDataBase.wireDiameter LIKE '%" & TBWireDiameter.Text & "%'" &
                    " AND springDataBase.OD LIKE '%" & TBOD.Text & "%'" &
                    " AND springDataBase.L0 LIKE '%" & TBL0.Text & "%'" &
                    " AND springDataBase.Nt LIKE '%" & TBNt.Text & "%'" & " AND springDataBase.productID LIKE '%" & TBProductID.Text & "%'" & " ;"

        Try
            Dim dt = Await Task(Of DataTable).Run(Function() LoadDataTable(sql_command))
            spring_bs.DataSource = dt
            DataGridView1.DataSource = spring_bs
        Catch ex As Exception
            MsgBox("خطا در ارتباط با دیتابیس", vbCritical + vbMsgBoxRight, "خطا")
            Logger.LogFatal(ex.Message, ex)
        End Try

    End Sub



    Private Sub BTModify_Click(sender As Object, e As EventArgs) Handles BTModify.Click
        'Dim productDatabaseID As String = DataGridView1.SelectedRows(0).Cells(0).Value.ToString 'ID of the selected product
        'productForm.TBdbID.Text = DataGridView1.SelectedRows(0).Cells(0).Value.ToString 'ID of the selected product
        productForm.TBdbID.Text = DataGridView1.SelectedRows(0).Cells("شماره شناسایی").Value.ToString 'ID of the selected product
        productFormState = "modify"
        productForm.Show()
    End Sub

    Private Sub BTClear_Click(sender As Object, e As EventArgs) Handles BTClear.Click
        '' Clear the search form and show all of the data 
        '' TODO: it's more efficient if you search the data using bindinsource filter instead of a seperate database query
        TBProductID.Text = ""
        TBProductName.Text = ""
        TBL0.Text = ""
        TBNt.Text = ""
        TBOD.Text = ""
        TBWireDiameter.Text = ""
        'Using cn As New OleDbConnection(connectionString)
        '    Using cmd As New OleDbCommand With {.Connection = cn}
        '        cmd.CommandText = "SELECT " & springDataBaseColumnNames & " FROM springDataBase;"
        '        Dim dt As New DataTable With {.TableName = "springDataBase"}
        '        Try
        '            cn.Open()
        '            Dim ds As New DataSet
        '            Dim springDBTable As New DataTable With {.TableName = "springDataBase"}
        '            ds.Tables.Add(springDBTable)
        '            ds.Load(cmd.ExecuteReader(), LoadOption.OverwriteChanges, springDBTable)
        '            DataGridView1.DataSource = ds.Tables("springDataBase")
        '        Catch ex As Exception
        '            MsgBox("خطا در ارتباط با دیتابیس", vbCritical + vbMsgBoxRight, "خطا")
        '            Logger.LogFatal(ex.Message, ex)
        '        Finally
        '            cn.Close()
        '        End Try

        '    End Using
        'End Using

        'Dim sql_command = "SELECT " & springDataBaseColumnNames & " FROM springDataBase;"
        'Dim springdt As New DataTable
        'Try
        '    springdt = Await Task(Of DataTable).Run(Function() LoadDataTable(sql_command))
        '    spring_bs.DataSource = springdt
        '    DataGridView1.DataSource = spring_bs
        '    'DataGridView1.Columns(0).Visible = False
        'Catch ex As Exception
        '    MsgBox("خطا در ارتباط با دیتابیس" + Environment.NewLine + ex.Message, vbCritical + vbMsgBoxRight, "خطا")
        '    Logger.LogFatal("spring data base couldn't be loaded.", ex)
        'End Try
        spring_bs.Filter = ""
    End Sub

    Private Sub BTNewProduct_Click(sender As Object, e As EventArgs) Handles BTNewProduct.Click
        productFormState = "new"
        productForm.Show()
    End Sub

    Private Async Sub BTCustomerSearch_Click(sender As Object, e As EventArgs) Handles BTCustomerSearch.Click
        'This subroutine searches the customer database based on textbox value
        'Using cn As New OleDbConnection(connectionString)
        '    Using cmd As New OleDbCommand With {.Connection = cn}
        '        cmd.CommandText = "SELECT " & customerDataBaseColumnNames & " FROM customers WHERE " &
        '            "customers.customerName LIKE '%" & TBCustomerNameSearch.Text & "%' ;"

        '        Dim dt As New DataTable With {.TableName = "customers"}
        '        Try
        '            cn.Open()
        '            Dim ds As New DataSet
        '            Dim customers As New DataTable With {.TableName = "customers"}
        '            ds.Tables.Add(customers)
        '            ds.Load(cmd.ExecuteReader(), LoadOption.OverwriteChanges, customers)
        '            DataGridView2.DataSource = ds.Tables("customers")
        '        Catch ex As Exception
        '            MsgBox("خطا در ارتباط با دیتابیس", vbCritical + vbMsgBoxRight, "خطا")
        '            Logger.LogFatal(ex.Message, ex)
        '        Finally
        '            cn.Close()
        '        End Try
        '    End Using
        'End Using
        '' TODO: fix the errors
        Dim sql_command = "SELECT " & customerDataBaseColumnNames & " FROM customers WHERE " &
                    "customers.customerName LIKE '%" & TBCustomerNameSearch.Text & "%' ;"
        Dim dt = Await Task(Of DataTable).Run(Function() LoadDataTable(sql_command))
        customer_bs.DataSource = dt
        DataGridView2.DataSource = customer_bs

        'customer_bs.Filter = String.Format("[نام مشتری] LIKE '%{0}%'", TBCustomerNameSearch.Text)
        'customer_bs.Filter = ConstructSearchQuery("نام مشتری", TBCustomerNameSearch.Text)

    End Sub

    Private Sub BTModifyCustomer_Click(sender As Object, e As EventArgs) Handles BTModifyCustomer.Click

        'Send Id of the selected customer to customerForm
        customerForm.TBdbID.Text = DataGridView2.SelectedRows(0).Cells("شماره شناسایی مشتری").Value.ToString
        customerFormState = "modify"
        customerForm.Show()
    End Sub

    Private Sub BTNewCustomer_Click(sender As Object, e As EventArgs) Handles BTNewCustomer.Click
        customerFormState = "new"
        customerForm.Show()
    End Sub

    Private Sub BTClearCustomer_Click(sender As Object, e As EventArgs) Handles BTClearCustomer.Click
        'Clear Customer DataBase Search Form 
        'TBCustomerNameSearch.Text = ""
        'Using cn As New OleDbConnection(connectionString)
        '    Using cmd As New OleDbCommand With {.Connection = cn}
        '        cmd.CommandText = "SELECT " & customerDataBaseColumnNames & " FROM customers;"
        '        Dim dt As New DataTable With {.TableName = "customers"}
        '        Try
        '            cn.Open()
        '            Dim ds As New DataSet
        '            Dim customersDBTable As New DataTable With {.TableName = "customers"}
        '            ds.Tables.Add(customersDBTable)
        '            ds.Load(cmd.ExecuteReader(), LoadOption.OverwriteChanges, customersDBTable)
        '            DataGridView2.DataSource = ds.Tables("customers")
        '            DataGridView2.Columns(0).Visible = False
        '            cn.Close()
        '        Catch ex As Exception
        '            MsgBox("خطا در ارتباط با دیتابیس", vbCritical + vbMsgBoxRight, "خطا")
        '            Logger.LogFatal(ex.Message, ex)
        '        Finally
        '            cn.Close()
        '        End Try
        '    End Using
        'End Using
        customer_bs.Filter = ""
    End Sub
    Private Sub MyTabControl_SelectedIndexChanged(ByVal sender As Object,
                                              ByVal e As System.EventArgs) _
            Handles TabControl1.SelectedIndexChanged

        '' When the third tab is selected, it fills customer and product name for emkansanji
        Dim indexOfSelectedTab As Integer = TabControl1.SelectedIndex
        If indexOfSelectedTab = 2 Then
            On Error Resume Next
            TBEnergySazProductName.Text = DataGridView1.SelectedRows(0).Cells("نام محصول").Value.ToString
            TBCustomerProductName.Text = DataGridView1.SelectedRows(0).Cells("نام محصول").Value.ToString
            TBProductIDES.Text = DataGridView1.SelectedRows(0).Cells("شماره شناسایی").Value.ToString
            '' Parse production process if it exist for the product in the springDataBase
            If Len(DataGridView1.SelectedRows(0).Cells("productionProcess").Value.ToString) > 0 Then
                ParseProductionProcess(CBA, DataGridView1.SelectedRows(0).Cells("productionProcess").Value.ToString)
            End If
            TBCustomerName.Text = DataGridView2.SelectedRows(0).Cells("نام مشتری").Value.ToString
            TBCustomerID.Text = DataGridView2.SelectedRows(0).Cells("شماره شناسایی مشتری").Value.ToString
        End If
    End Sub





    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '' clears the new emkansanji form
        TBProductIDES.Text = ""
        TBCustomerID.Text = ""
        TBCustomerProductName.Text = ""
        TBCustomerDwgNo.Text = ""
        TBQuantity.Text = ""
        TBLetterNo.Text = ""
        TBLetterDate.Text = ""
        TBOrderNo.Text = ""
        TBProccessingDate.Text = ""
        CBStandard.Text = ""
        TBGrade.Text = ""
        TBCustomerProductCode.Text = ""
        TBComment.Text = ""
        TBEnergySazProductName.Text = ""
        TBCustomerName.Text = ""
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        emkanSanjiForm.Show()
    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        wires.Show()
        'Logger.LogInfo("Hello World!")
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ' KaveNegar
        'Dim sender_no = "1000596446"
        'Dim receptor = "09188183115"
        'Dim Message = "تست سرویس ارسال پیام سلام سلام"
        'Dim api As New Kavenegar.KavenegarApi("4578412F307931426F456F6D725A544173357933344D7336364C6139386B4838346C657653793872414F343D")
        'api.Send(sender_no, receptor, Message)

        ''https://developers.ghasedak.io/panel/line
        Try
            Dim Message = "سلام احوالت"
            Dim lineNumber = "50001212124042"
            Dim receptor() As String = {"09188183115"}
            Dim sms As New Ghasedak.Api("9c00cf12398ffbd28551a8d1645e71d07ed8c7acbb46963a2bb285774eb571c4")
            'Dim result = sms.SendSMS(Message, receptor, lineNumber)
            Dim result = sms.Verify(1, "order",
                                      receptor,
                                      "25").Result
            MsgBox(result)
        Catch ex As Ghasedak.Exceptions.ApiException
            Console.WriteLine(ex.Message)
        Catch ex As Ghasedak.Exceptions.ConnectionException
            Console.WriteLine(ex.Message)
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
    End Sub

    Private Sub TBProductName_TextChanged(sender As Object, e As EventArgs) Handles TBProductName.TextChanged

        'Dim filter = String.Format("[نام محصول] LIKE '%{0}%'", searchTerm)
        'Console.WriteLine(searchTerm)
        '' This is singleon way
        'Dim searchTerm = ConstructSearchQuery("نام محصول", TBProductName.Text)
        ''searchTerm = searchTerm & " AND " & ConstructSearchQuery("کد کالا", TBProductID.Text)
        'spring_bs.Filter = searchTerm
        SearchForProduct()
    End Sub
    Private Function ConstructSearchQuery(columnName As String, input As String) As String
        '' When the user's search term is [inner LSD1] the program should show a product with name [inner bogie LSD1 Spring]
        '' Unlike SQL syntax bindingsource's filter property doesn't support wildcard characters in the middle of a string
        '' so we use this function to construct a query for when the search term has spaces
        '' so [inner LSD1] -> [columnName] LIKE '%inner%' AND [columnName] LIKE '%LSD1%'
        '' TODO: think of a faster way to implement this [DONE]

        input = Regex.Replace(input, "[\\/]", "")
        input = input.Replace("*", "[*]")
        input = input.Replace(" ", String.Format("%' AND [{0}] LIKE '%", columnName))
        Dim searchTerm = String.Format("[{0}] LIKE '%", columnName) + input + String.Format("%'")

        'Dim wordArray = input.Split(" ")
        'Dim searchTerm = String.Format(" [{0}] LIKE '%{1}%' ", columnName, wordArray(0))
        'For count = 1 To wordArray.Length - 1
        '    searchTerm += String.Format(" AND [{0}] LIKE '%{1}%' ", columnName, wordArray(count))
        'Next
        Return searchTerm
    End Function

    Private Sub SearchForProduct()
        Dim searchTerm = ConstructSearchQuery("نام محصول", TBProductName.Text)
        'Dim filter = String.Format("[نام محصول] LIKE '%{0}%'", searchTerm)
        'Console.WriteLine(searchTerm)
        searchTerm = searchTerm & " AND " & ConstructSearchQuery("کد کالا", TBProductID.Text) _
                                & " AND " & ConstructSearchQuery("قطر مفتول", TBWireDiameter.Text) _
                                & " AND " & ConstructSearchQuery("قطر خارجی", TBOD.Text) _
                                & " AND " & ConstructSearchQuery("طول آزاد", TBL0.Text) _
                                & " AND " & ConstructSearchQuery("حلقه کل", TBNt.Text)
        spring_bs.Filter = searchTerm
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs)
        '' Send a HTTP POST request to a server including the parametersS
        Using client As New Net.WebClient
            Try
                Dim reqparm As New Specialized.NameValueCollection
                reqparm.Add("param1", TBWireDiameter.Text) ' -> request.form['param1'] = TBWireDiameter.text
                reqparm.Add("param2", "othervalue")
                Dim responsebytes = client.UploadValues("https://pedramh.pythonanywhere.com/", "POST", reqparm)
                Dim responsebody = (New Text.UTF8Encoding).GetString(responsebytes)
                MsgBox(responsebody)
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

        End Using
    End Sub

    Private Sub Button5_Click_1(sender As Object, e As EventArgs) Handles Button5.Click
        'Dim sql = "UPDATE springDataBase SET productionProcess = '1010111110' WHERE productionMethod = 'سرد پیچ' "
        'Dim sql = "SELECT * FROM springDataBase WHERE productionMethod = 'سرد پیچ' "
        Dim ColumnName = " emkansanji.ID, emkansanji.pProcess , springDataBase.productionProcess, springDataBase.productionMethod "
        Dim sql_command = "SELECT " & ColumnName & " FROM ((emkansanji LEFT JOIN springDataBase ON emkansanji.productID = springDataBase.ID) LEFT JOIN customers ON emkansanji.customerID = customers.ID) "
        Dim sql = "UPDATE emkansanji SET pProcess = '0101111110' WHERE productID <> 553 "

        DataGridView1.DataSource = LoadDataTable(sql_command)
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        '' PostgreSQL migration
        Dim dt = New DataTable
        Using con = GetPostgresCon()
            Dim cmd = con.CreateCommand()
            Dim query As String = "SELECT * FROM cities;"
            con.Open()
            cmd.CommandText = query
            'cmd.ExecuteNonQuery()
            dt.Load(cmd.ExecuteReader())
            con.Close()
        End Using
        DataGridView1.DataSource = dt


        'Using cn As New NpgsqlConnection(postgresConString)
        '    Using cmd As New NpgsqlCommand With {.Connection = cn}
        '        cmd.CommandText = "SELECT * FROM cities;" ''Because of the case sensitivity thing the database name should be in double qoutes
        '        Dim dt As New DataTable With {.TableName = "springDataBase"}
        '        'Try
        '        cn.Close()
        '        cn.Close()

        '        cn.Open()
        '        Dim ds As New DataSet
        '        Dim springDBTable As New DataTable With {.TableName = "springDataBase"}
        '        ds.Tables.Add(springDBTable)
        '        ds.Load(cmd.ExecuteReader(), LoadOption.OverwriteChanges, springDBTable)
        '        DataGridView1.DataSource = ds.Tables("springDataBase")
        '        'DataGridView1.Columns("EmployeeID").Visible = False
        '        'Catch ex As Exception
        '        ' very common for a developer to simply ignore errors, unwise.
        '        ' MsgBox("error")
        '        ' End Try
        '        cn.Close()
        '    End Using
        'End Using
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        FrmMenu.Show()
    End Sub

    Private Sub TBCustomerNameSearch_TextChanged(sender As Object, e As EventArgs) Handles TBCustomerNameSearch.TextChanged
        '' filters the data in customer data (datagridview2) in realtime based on the text in search textbox
        customer_bs.Filter = ConstructSearchQuery("نام مشتری", TBCustomerNameSearch.Text)
    End Sub

    Private Sub TBProductID_TextChanged(sender As Object, e As EventArgs) Handles TBProductID.TextChanged
        SearchForProduct()
    End Sub

    Private Sub TBWireDiameter_TextChanged(sender As Object, e As EventArgs) Handles TBWireDiameter.TextChanged
        SearchForProduct()
    End Sub

    Private Sub TBOD_TextChanged(sender As Object, e As EventArgs) Handles TBOD.TextChanged
        SearchForProduct()
    End Sub

    Private Sub TBL0_TextChanged(sender As Object, e As EventArgs) Handles TBL0.TextChanged
        SearchForProduct()
    End Sub

    Private Sub TBNt_TextChanged(sender As Object, e As EventArgs) Handles TBNt.TextChanged
        SearchForProduct()
    End Sub

    '' -----------------------------------------------------------------------------------------------------------------------
    '' -----------------------------------------------------------------------------------------------------------------------
    '' -----------------------------------------------------------------------------------------------------------------------
    '' -----------------------------------------------------------------------------------------------------------------------
    '' ---------------------------- Submit new emkansanji to the table and create the excel file -----------------------------
    '' -----------------------------------------------------------------------------------------------------------------------
    '' -----------------------------------------------------------------------------------------------------------------------
    '' -----------------------------------------------------------------------------------------------------------------------
    '' -----------------------------------------------------------------------------------------------------------------------

    '' generate production process code
    Dim productionProcess As String
    '' gererate inspection process code
    Dim inspectionProcess As String
    '' generate order type code   - a two digit number - first digit New product 1 - changing old product 2 - old product 3
    Dim orderType As String

    Private Function CreateOrderExcelFile() As String



        '' Generate the file path for creating the excel file
        '' path would be : base\year\month\customername\productname.xlsx
        Dim pc As New Globalization.PersianCalendar
        Dim excelFiledir As String = pc.GetYear(Now).ToString & "\" & getMonthName(pc.GetMonth(Now)) & "\" & DataGridView2.SelectedRows(0).Cells("نام مشتری").Value.ToString & "\"
        Dim excelFileName As String = DataGridView1.SelectedRows(0).Cells("نام محصول").Value.ToString


        excelFileName = stripFileName(excelFileName)

        Dim path As String = excelFilesBasePath & excelFiledir
        Dim saveFilePath As String = path & excelFileName 'Complete path of the excel file

        'Check to see if file with this name exist to rename it to prevent overwriting
        saveFilePath = preverntOverwriting(saveFilePath, ".xlsx")

        Try

            '' Open Emkansanji Excel template file to fill in the data
            If excelTemplateFilePath.Substring(0, 1) = "\" Then
                '' Acount for when file address is relative
                excelTemplateFilePath = IO.Directory.GetParent(Application.ExecutablePath).FullName + excelTemplateFilePath
            End If
            Console.WriteLine(excelTemplateFilePath)
            Dim excel As Excel.Application = New Excel.Application
            Dim w As Excel.Workbook = excel.Workbooks.Open(excelTemplateFilePath)

            '' ----------------------------- Populate fields in the emkansanji excel Template -------------------------------
            excel.Range("customerName").Value = NormalizeString(DataGridView2.SelectedRows(0).Cells("نام مشتری").Value.ToString)
            excel.Range("letterNo").Value = NormalizeString(TBLetterNo.Text)
            excel.Range("pName").Value = NormalizeString(TBCustomerProductName.Text)
            excel.Range("letterDate").Value = NormalizeString(TBLetterDate.Text)
            excel.Range("dwgNo").Value = NormalizeString(TBCustomerDwgNo.Text)
            excel.Range("quantity").Value = NormalizeString(TBQuantity.Text)
            excel.Range("sampleQuantity").Value = NormalizeString(TBSampleQuantity.Text)
            excel.Range("pDate").Value = NormalizeString(TBProccessingDate.Text)
            excel.Range("standard").Value = NormalizeString(CBStandard.Text)
            excel.Range("grade").Value = NormalizeString(TBGrade.Text)
            excel.Range("customerProductCode").Value = NormalizeString(TBCustomerProductCode.Text)
            excel.Range("comment").Value = NormalizeString(TBComment.Text)

            '' -------------------------------------------------------------------------------------------------------------
            excel.Range("ESpName").Value = NormalizeString(DataGridView1.SelectedRows(0).Cells("نام محصول").Value.ToString)
            excel.Range("ESProductCode").Value = NormalizeString(DataGridView1.SelectedRows(0).Cells("کد کالا").Value.ToString)
            '' -------------------------------------------------------------------------------------------------------------
            excel.Range("springType").Value = NormalizeString(DataGridView1.SelectedRows(0).Cells("نوع فنر").Value.ToString)
            excel.Range("material").Value = NormalizeString(DataGridView1.SelectedRows(0).Cells("جنس مواد").Value.ToString)
            excel.Range("wireD").Value = NormalizeString(DataGridView1.SelectedRows(0).Cells("قطر مفتول").Value.ToString)
            excel.Range("OD").Value = NormalizeString(DataGridView1.SelectedRows(0).Cells("قطر خارجی").Value.ToString)
            excel.Range("mandrel").Value = NormalizeString(DataGridView1.SelectedRows(0).Cells("قطر شفت").Value.ToString)
            excel.Range("Nt").Value = NormalizeString(DataGridView1.SelectedRows(0).Cells("حلقه کل").Value.ToString)
            excel.Range("Na").Value = NormalizeString(DataGridView1.SelectedRows(0).Cells("حلقه فعال").Value.ToString)
            excel.Range("L0").Value = NormalizeString(DataGridView1.SelectedRows(0).Cells("طول آزاد").Value.ToString)
            excel.Range("coilingDirection").Value = NormalizeString(DataGridView1.SelectedRows(0).Cells("جهت پیچش").Value.ToString)
            excel.Range("springRate").Value = NormalizeString(DataGridView1.SelectedRows(0).Cells("ریت فنر").Value.ToString)
            excel.Range("firstCoil").Value = NormalizeString(DataGridView1.SelectedRows(0).Cells("شکل حلقه ابتدا").Value.ToString)
            excel.Range("lastCoil").Value = NormalizeString(DataGridView1.SelectedRows(0).Cells("شکل حلقه انتها").Value.ToString)
            excel.Range("Force1").Value = NormalizeString(DataGridView1.SelectedRows(0).Cells("F1").Value.ToString)
            excel.Range("Length1").Value = NormalizeString(DataGridView1.SelectedRows(0).Cells("L1").Value.ToString)
            excel.Range("Force2").Value = NormalizeString(DataGridView1.SelectedRows(0).Cells("F2").Value.ToString)
            excel.Range("Length2").Value = NormalizeString(DataGridView1.SelectedRows(0).Cells("L2").Value.ToString)
            excel.Range("Force3").Value = NormalizeString(DataGridView1.SelectedRows(0).Cells("F3").Value.ToString)
            excel.Range("Length3").Value = NormalizeString(DataGridView1.SelectedRows(0).Cells("L3").Value.ToString)
            excel.Range("forceUnit").Value = NormalizeString(DataGridView1.SelectedRows(0).Cells("واحد نیرو").Value.ToString)
            '' -----------------------------------------------------------------------------------------------------------
            excel.Range("wireLength").Value = NormalizeString(DataGridView1.SelectedRows(0).Cells("طول مفتول").Value.ToString)

            '' ------------------------------------------------------------------------------------------------------------
            excel.Range("productionProcess").Value = productionProcess
            excel.Range("inspectionProcess").Value = inspectionProcess
            excel.Range("orderTypeCode").Value = orderType

            'Check if there is an address provided for the second excel file, if not it uses working directory
            Dim saveDuplicatePath As String
            If ConfigurationManager.ConnectionStrings("saveAsFilePath").ConnectionString = "" Then
                saveDuplicatePath = My.Application.Info.DirectoryPath
            Else
                saveDuplicatePath = ConfigurationManager.ConnectionStrings("saveAsFilePath").ConnectionString
            End If
            MkDir(path)  'Make the directory if it doesn't exist
            w.SaveAs(saveFilePath)
            w.SaveAs(saveDuplicatePath & "\" & excelFileName) 'Save another file in the application directory
            w.Close()
            Logger.LogInfo("Excel File Created with path (" + saveFilePath + ")")
            Return saveFilePath
        Catch ex As Exception
            MsgBox("خطا در تکمیل قالب اکسل امکان سنجی. فایل اکسل را چک کرده و مجددا امتحان کنید", vbCritical + vbMsgBoxRight, "خطا")
            Logger.LogFatal(ex.Message, ex)
            Return "abort"
        End Try

    End Function

    Private Async Sub BTSubmit_Click(sender As Object, e As EventArgs) Handles BTSubmit.Click

        LStatus.Text = "در حال ثبت اطلاعات امکان سنجی جدید در دیتابیس ..."
        LStatus.Visible = True

        '' generate production process code
        productionProcess = GenerateProductionProcess(CBA)

        '' gererate inspection process code
        inspectionProcess = GenerateInspectionProcess(CBA_inspection)

        '' generate order type code   - a two digit number - first digit New product 1 - changing old product 2 - old product 3
        orderType = ""
        If RBOldProduct.Checked Then
            orderType = "3"
        ElseIf RBChangeProduct.Checked = True Then
            orderType = "2"
        Else
            orderType = "1"
        End If
        '' second digit -> main order 1 , order amendment :2 
        If RBAmendOrder.Checked = True Then
            orderType = orderType & "2"
        Else
            orderType = orderType & "1"
        End If



        '' ------------------------------------ Writing to data base ---------------------------------------------------------
        'Using cn As New OleDbConnection(connectionString)
        '    Using cmd As New OleDbCommand With {.Connection = cn}


        '        '' generate the query to add value into the table
        '        Dim columnNames As String = " ( productID , customerID , customerProductName , customerDwgNo, quantity, sampleQuantity, " &
        '             " letterNo , letterDate , orderNo , orderType ,dateOfProccessing , standard , " &
        '             " grade , productCode , comment ,pProcess ,inspectionProcess , orderState , excelFilePath ) "

        '        Dim valueString As String = "('" & TBProductIDES.Text & "','" & TBCustomerID.Text & "','" & TBCustomerProductName.Text & "','" & TBCustomerDwgNo.Text & "','" & TBQuantity.Text & "','" & TBSampleQuantity.Text & "','" &
        '                TBLetterNo.Text & "','" & TBLetterDate.Text & "','" & TBOrderNo.Text & "','" & orderType & "','" & TBProccessingDate.Text & "','" & CBStandard.Text & "','" &
        '                TBGrade.Text & "','" & TBCustomerProductCode.Text & "','" & TBComment.Text & "','" & productionProcess & "','" & inspectionProcess & "', 'امکان سنجی اولیه تولید' , '" & saveFilePath & "' )"

        '        cmd.CommandText = "INSERT INTO emkansanji" & columnNames & " VALUES " & valueString & ";"
        '        Try
        '            cn.Open()
        '            cmd.ExecuteReader()
        '            Logger.LogInfo("New EmkanSanji with Product ID = " + TBProductIDES.Text + " And Customer ID = " + TBCustomerID.Text)
        '        Catch ex As Exception
        '            MsgBox("خطا در ارتباط با دیتابیس", vbCritical + vbMsgBoxRight, "خطا")
        '            Logger.LogFatal(ex.Message, ex)
        '        Finally
        '            cn.Close()
        '        End Try
        '    End Using
        'End Using

        '' generate the query to add value into the table

        LStatus.Text = "در حال آماده سازی فایل اکسل امکان سنجی ..."
        Me.Cursor = Cursors.WaitCursor

        Dim saveFilePath = CreateOrderExcelFile()

        'Dim saveFilePath = Await Task(Of String).Run(Function() CreateOrderExcelFile())

        If saveFilePath = "abort" Then
            '' If there is a problem in creation of excel file, data would not be added to the database
            LStatus.Visible = False
            Me.Cursor = Cursors.Default
            MsgBox("امکان سنجی جدید در دیتابیس ثبت نشد!", vbCritical + vbMsgBoxRight, "خطا")
            Exit Sub
        End If

        Dim columnNames As String = " ( productID , customerID , customerProductName , customerDwgNo, quantity, sampleQuantity, " &
                     " letterNo , letterDate , orderNo , orderType ,dateOfProccessing , standard , " &
                     " grade , productCode , comment ,pProcess ,inspectionProcess , orderState , excelFilePath ) "

        Dim valueString As String = "('" & TBProductIDES.Text & "','" & TBCustomerID.Text & "','" & TBCustomerProductName.Text & "','" & TBCustomerDwgNo.Text & "','" & TBQuantity.Text & "','" & TBSampleQuantity.Text & "','" &
                        TBLetterNo.Text & "','" & TBLetterDate.Text & "','" & TBOrderNo.Text & "','" & orderType & "','" & TBProccessingDate.Text & "','" & CBStandard.Text & "','" &
                        TBGrade.Text & "','" & TBCustomerProductCode.Text & "','" & TBComment.Text & "','" & productionProcess & "','" & inspectionProcess & "', 'امکان سنجی اولیه تولید' , '" & saveFilePath & "' )"

        Using con = GetDatabaseCon()
            Dim cmd = con.CreateCommand()
            cmd.CommandText = "INSERT INTO emkansanji" & columnNames & " VALUES " & valueString & ";"
            Try
                Await con.OpenAsync()
                Await cmd.ExecuteNonQueryAsync()
                Logger.LogInfo("New EmkanSanji with Product ID = " + TBProductIDES.Text + " And Customer ID = " + TBCustomerID.Text)
                MsgBox("ثبت امکان سنجی با موفقیت انجام شد", vbInformation + vbMsgBoxRight + RightToLeft, "امکان سنجی")
            Catch ex As Exception
                MsgBox("خطا در ارتباط با دیتابیس. اطلاعات سفارش در دیتابیس ثبت نشد" + Environment.NewLine + "فایل اکسل ساخته شده از پوشه اصلی حذف خواهد شد" + Environment.NewLine + ex.Message, vbCritical + vbMsgBoxRight, "خطا")
                Logger.LogFatal(ex.Message, ex)
                If System.IO.File.Exists(saveFilePath) Then
                    Try
                        System.IO.File.Delete(saveFilePath)
                    Catch
                        '' Do Nothig? 
                    End Try
                End If
            Finally
                con.Close()
            End Try
        End Using
        '' ------------------------------------- END OF WRITING TO DATABASE -----------------------------------------
        LStatus.Visible = False
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        'Dim s1 = GetSaltedHash(TBCustomerProductName.Text, "salt")
        'Dim s2 = GetSaltedHash(TBOrderNo.Text, "salt")
        'If StrComp(s1, s2, False) = 0 Then
        '    Console.WriteLine("Login Successful!")
        'End If
        Console.WriteLine(My.Settings.loggedin)
        Console.WriteLine(My.Settings.loginDate.ToString)
        Console.WriteLine(My.Settings.validation)
        Dim dummydate As Date = New System.DateTime(2020, 4, 6, 12, 0, 0)

        Dim DaysSinceLastLogin As Int32 = dummydate.Subtract(My.Settings.loginDate).Days
        Console.WriteLine(DaysSinceLastLogin)

    End Sub

End Class
