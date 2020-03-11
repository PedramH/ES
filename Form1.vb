Imports System.Data.OleDb

Imports System.Configuration
Imports Excel = Microsoft.Office.Interop.Excel



Public Class mainForm



    Protected Overrides Function ProcessCmdKey(ByRef msg As System.Windows.Forms.Message,
                                           ByVal keyData As System.Windows.Forms.Keys) _
                                           As Boolean
        ' This code send Tab key everytime Enterkey is pressed INSIDE OF A TEXTBOX

        If msg.WParam.ToInt32() = CInt(Keys.Enter) AndAlso TypeOf Me.ActiveControl Is TextBox Then
            If Me.ActiveControl.Name <> "TBComment" Then
                SendKeys.Send("{Tab}")
                Return True
            End If
        End If

        Return MyBase.ProcessCmdKey(msg, keyData)
    End Function



    Private Sub mainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Loading Springs table into datagridview1
        Using cn As New OleDbConnection(connectionString)
            Using cmd As New OleDbCommand With {.Connection = cn}
                cmd.CommandText = "SELECT " & springDataBaseColumnNames & " FROM springDataBase;"
                Dim dt As New DataTable With {.TableName = "springDataBase"}
                Try
                    cn.Open()
                    Dim ds As New DataSet
                    Dim springDBTable As New DataTable With {.TableName = "springDataBase"}
                    ds.Tables.Add(springDBTable)
                    ds.Load(cmd.ExecuteReader(), LoadOption.OverwriteChanges, springDBTable)
                    DataGridView1.DataSource = ds.Tables("springDataBase")
                    DataGridView1.Columns(0).Visible = False
                Catch ex As Exception
                    MsgBox("خطا در ارتباط با دیتابیس", vbCritical + vbMsgBoxRight, "خطا")
                    Logger.LogFatal(ex.Message, ex)
                Finally
                    cn.Close()
                End Try

            End Using
        End Using
        'Loading Customers info into datagridView2
        Using cn As New OleDbConnection(connectionString)
            Using cmd As New OleDbCommand With {.Connection = cn}

                cmd.CommandText = "SELECT " & customerDataBaseColumnNames & " FROM customers;"

                Dim dt As New DataTable With {.TableName = "customers"}
                Try
                    cn.Open()
                    Dim ds As New DataSet
                    Dim customersDBTable As New DataTable With {.TableName = "customers"}
                    ds.Tables.Add(customersDBTable)
                    ds.Load(cmd.ExecuteReader(), LoadOption.OverwriteChanges, customersDBTable)
                    DataGridView2.DataSource = ds.Tables("customers")
                    DataGridView2.Columns(0).Visible = False
                Catch ex As Exception
                    MsgBox("خطا در ارتباط با دیتابیس", vbCritical + vbMsgBoxRight, "خطا")
                    Logger.LogFatal(ex.Message, ex)
                Finally
                    cn.Close()
                End Try

            End Using
        End Using


    End Sub

    Private Sub BTSearch_Click(sender As Object, e As EventArgs) Handles BTSearch.Click
        'This subroutine searches the database based on textbox value
        'Dim connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\ESDB.accdb"

        Using cn As New OleDbConnection(connectionString)
            Using cmd As New OleDbCommand With {.Connection = cn}
                cmd.CommandText = "SELECT " & springDataBaseColumnNames & " FROM springDataBase WHERE " &
                    "springDataBase.productName LIKE '%" & TBProductName.Text & "%' AND" &
                    " springDataBase.wireDiameter LIKE '%" & TBWireDiameter.Text & "%'" &
                    " AND springDataBase.OD LIKE '%" & TBOD.Text & "%'" &
                    " AND springDataBase.L0 LIKE '%" & TBL0.Text & "%'" &
                    " AND springDataBase.Nt LIKE '%" & TBNt.Text & "%'" & " AND springDataBase.productID LIKE '%" & TBProductID.Text & "%'" & " ;"

                Dim dt As New DataTable With {.TableName = "springDataBase"}
                Try
                    cn.Open()
                    Dim ds As New DataSet
                    Dim springDBTable As New DataTable With {.TableName = "springDataBase"}
                    ds.Tables.Add(springDBTable)
                    ds.Load(cmd.ExecuteReader(), LoadOption.OverwriteChanges, springDBTable)
                    DataGridView1.DataSource = ds.Tables("springDataBase")
                    cn.Close()
                Catch ex As Exception
                    MsgBox("خطا در ارتباط با دیتابیس", vbCritical + vbMsgBoxRight, "خطا")
                    Logger.LogFatal(ex.Message, ex)
                Finally
                    cn.Close()
                End Try
            End Using
        End Using
    End Sub



    Private Sub BTModify_Click(sender As Object, e As EventArgs) Handles BTModify.Click
        'Dim productDatabaseID As String = DataGridView1.SelectedRows(0).Cells(0).Value.ToString 'ID of the selected product
        productForm.TBdbID.Text = DataGridView1.SelectedRows(0).Cells(0).Value.ToString 'ID of the selected product
        productFormState = "modify"

        productForm.Show()
    End Sub

    Private Sub BTClear_Click(sender As Object, e As EventArgs) Handles BTClear.Click
        'Dim connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\ESDB.accdb"
        TBProductID.Text = ""
        TBProductName.Text = ""
        TBL0.Text = ""
        TBNt.Text = ""
        TBOD.Text = ""
        TBWireDiameter.Text = ""

        Using cn As New OleDbConnection(connectionString)
            Using cmd As New OleDbCommand With {.Connection = cn}
                cmd.CommandText = "SELECT " & springDataBaseColumnNames & " FROM springDataBase;"
                Dim dt As New DataTable With {.TableName = "springDataBase"}
                Try
                    cn.Open()
                    Dim ds As New DataSet
                    Dim springDBTable As New DataTable With {.TableName = "springDataBase"}
                    ds.Tables.Add(springDBTable)
                    ds.Load(cmd.ExecuteReader(), LoadOption.OverwriteChanges, springDBTable)
                    DataGridView1.DataSource = ds.Tables("springDataBase")
                Catch ex As Exception
                    MsgBox("خطا در ارتباط با دیتابیس", vbCritical + vbMsgBoxRight, "خطا")
                    Logger.LogFatal(ex.Message, ex)
                Finally
                    cn.Close()
                End Try

            End Using
        End Using
    End Sub

    Private Sub BTNewProduct_Click(sender As Object, e As EventArgs) Handles BTNewProduct.Click
        productFormState = "new"
        productForm.Show()
    End Sub

    Private Sub BTCustomerSearch_Click(sender As Object, e As EventArgs) Handles BTCustomerSearch.Click
        'This subroutine searches the customer database based on textbox value


        Using cn As New OleDbConnection(connectionString)
            Using cmd As New OleDbCommand With {.Connection = cn}
                cmd.CommandText = "SELECT " & customerDataBaseColumnNames & " FROM customers WHERE " &
                    "customers.customerName LIKE '%" & TBCustomerNameSearch.Text & "%' ;"

                Dim dt As New DataTable With {.TableName = "customers"}
                Try
                    cn.Open()
                    Dim ds As New DataSet
                    Dim customers As New DataTable With {.TableName = "customers"}
                    ds.Tables.Add(customers)
                    ds.Load(cmd.ExecuteReader(), LoadOption.OverwriteChanges, customers)
                    DataGridView2.DataSource = ds.Tables("customers")
                Catch ex As Exception
                    MsgBox("خطا در ارتباط با دیتابیس", vbCritical + vbMsgBoxRight, "خطا")
                    Logger.LogFatal(ex.Message, ex)
                Finally
                    cn.Close()
                End Try
            End Using
        End Using
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
        TBCustomerNameSearch.Text = ""
        Using cn As New OleDbConnection(connectionString)
            Using cmd As New OleDbCommand With {.Connection = cn}
                cmd.CommandText = "SELECT " & customerDataBaseColumnNames & " FROM customers;"
                Dim dt As New DataTable With {.TableName = "customers"}
                Try
                    cn.Open()
                    Dim ds As New DataSet
                    Dim customersDBTable As New DataTable With {.TableName = "customers"}
                    ds.Tables.Add(customersDBTable)
                    ds.Load(cmd.ExecuteReader(), LoadOption.OverwriteChanges, customersDBTable)
                    DataGridView2.DataSource = ds.Tables("customers")
                    DataGridView2.Columns(0).Visible = False
                    cn.Close()
                Catch ex As Exception
                    MsgBox("خطا در ارتباط با دیتابیس", vbCritical + vbMsgBoxRight, "خطا")
                    Logger.LogFatal(ex.Message, ex)
                Finally
                    cn.Close()
                End Try
            End Using
        End Using
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
            TBCustomerName.Text = DataGridView2.SelectedRows(0).Cells("نام مشتری").Value.ToString
            TBCustomerID.Text = DataGridView2.SelectedRows(0).Cells("شماره شناسایی مشتری").Value.ToString
        End If
    End Sub

    Private Sub BTSubmit_Click(sender As Object, e As EventArgs) Handles BTSubmit.Click

        LStatus.Text = "در حال ثبت اطلاعات امکان سنجی جدید در دیتابیس ..."
        LStatus.Visible = True

        ' Generate the file path for excel file
        Dim pc As New Globalization.PersianCalendar
        Dim excelFiledir As String = pc.GetYear(Now).ToString & "\" & getMonthName(pc.GetMonth(Now)) & "\" & DataGridView2.SelectedRows(0).Cells("نام مشتری").Value.ToString & "\"
        Dim excelFileName As String = DataGridView1.SelectedRows(0).Cells("نام محصول").Value.ToString


        excelFileName = stripFileName(excelFileName)

        Dim path As String = excelFilesBasePath & excelFiledir
        Dim saveFilePath As String = path & excelFileName 'Complete path of the excel file

        'Check to see if file with this name exist to rename it to prevent overwriting
        saveFilePath = preverntOverwriting(saveFilePath, ".xlsx")


        Using cn As New OleDbConnection(connectionString)
            Using cmd As New OleDbCommand With {.Connection = cn}

                Dim columnNames As String = " ( productID , customerID , customerProductName , customerDwgNo, quantity, " &
                     " letterNo , letterDate , orderNo , dateOfProccessing , standard , " &
                     " grade , productCode , comment , orderState , excelFilePath ) "

                Dim valueString As String = "('" & TBProductIDES.Text & "','" & TBCustomerID.Text & "','" & TBCustomerProductName.Text & "','" & TBCustomerDwgNo.Text & "','" & TBQuantity.Text & "','" &
                        TBLetterNo.Text & "','" & TBLetterDate.Text & "','" & TBOrderNo.Text & "','" & TBProccessingDate.Text & "','" & CBStandard.Text & "','" &
                        TBGrade.Text & "','" & TBCustomerProductCode.Text & "','" & TBComment.Text & "', 'امکان سنجی کیفیت' , '" & saveFilePath & "' )"

                cmd.CommandText = "INSERT INTO emkansanji" & columnNames & " VALUES " & valueString & ";"
                Try
                    cn.Open()
                    cmd.ExecuteReader()
                    Logger.LogInfo("New EmkanSanji with Product ID = " + TBProductIDES.Text + " And Customer ID = " + TBCustomerID.Text)
                Catch ex As Exception
                    MsgBox("خطا در ارتباط با دیتابیس", vbCritical + vbMsgBoxRight, "خطا")
                    Logger.LogFatal(ex.Message, ex)
                Finally
                    cn.Close()
                End Try


            End Using
        End Using


        '------------------------------------------------------
        LStatus.Text = "در حال آماده سازی فایل اکسل امکان سنجی ..."
        Me.Cursor = Cursors.WaitCursor


        Try
            Dim excel As Excel.Application = New Excel.Application
            Dim w As Excel.Workbook = excel.Workbooks.Open(excelTemplateFilePath)
            'TODO: Complete this list
            excel.Range("wireD").Value = NormalizeString(DataGridView1.SelectedRows(0).Cells("قطر مفتول").Value.ToString)
            excel.Range("OD").Value = NormalizeString(DataGridView1.SelectedRows(0).Cells("قطر خارجی").Value.ToString)
            excel.Range("ESpName").Value = NormalizeString(DataGridView1.SelectedRows(0).Cells("نام محصول").Value.ToString)
            excel.Range("pName").Value = NormalizeString(TBCustomerProductName.Text)
            excel.Range("Nt").Value = NormalizeString(DataGridView1.SelectedRows(0).Cells("حلقه کل").Value.ToString)
            excel.Range("L0").Value = NormalizeString(DataGridView1.SelectedRows(0).Cells("طول آزاد").Value.ToString)


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
        Catch ex As Exception
            MsgBox("خطا در تکمیل قالب اکسل امکان سنجی. فایل اکسل را چک کرده و مجددا امتحان کنید", vbCritical + vbMsgBoxRight, "خطا")
            Logger.LogFatal(ex.Message, ex)
        End Try

        LStatus.Visible = False
        Me.Cursor = Cursors.Default
        MsgBox("ثبت امکان سنجی با موفقیت انجام شد", vbInformation, "امکان سنجی")

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)
        emkanSanjiForm.Show()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
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
End Class
