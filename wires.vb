Imports System.Data.OleDb
Imports System.Text.RegularExpressions
Imports System.Threading
Public Class wires
    Dim firstTimeEnteringOrdersTab = True

    '' Define a global binding source for the data that is comming from the wires database and going into datagridview1 (Wires)   
    Dim bs As New BindingSource
    '' Define a global binding source for the data that is comming from the wires database and going into datagridview2 (Orders)   
    Dim bs2 As New BindingSource
    Public Async Function LoadWiresData() As Task
        '' This function load data from wire inventory table in the database into a datatable
        ''      it then bind that data table to bs ( a global binding source in this form) then add bs as the data source 
        ''      for datagridview1. 
        ''      This way there wont be a call to data base for each search. we just call the database once then use the 
        ''      binding source for filtering -> less data transfer, probably faster and we can do it on textbox.changed
        ''      because of the increased speed.

        '' Use FLOOR instead of int in postgreSQL
        Dim sql_command = "SELECT" + wiresColumnName + "FROM wireInventory A INNER JOIN wireReserve B ON A.wireCode = B.wireCode;"
        Dim dt = Await Task(Of DataTable).Run(Function() LoadDataTable(sql_command))
        bs.DataSource = dt
        DataGridView1.DataSource = dt
        '' Hide wireType and wireWeight
        DataGridView1.Columns(0).Visible = False
        DataGridView1.Columns(1).Visible = False
    End Function
    Private Function SearchWiresData()
        '' This function bs, a global binding source in this form which is the data source for datagridview1
        Dim wireType = ""
        If CBWireType.Text = "همه" Then
            wireType = ""
        Else
            wireType = CBWireType.Text
        End If
        bs.Filter = String.Format("[کد کالا] LIKE '%{0}%' AND [قطر مفتول] LIKE '%{1}%'  AND ([طول مفتول] >= '{2}' OR [wireType] = 'مفتول کویل' ) AND [wireType] LIKE '%{3}%'",
                                  TBWireCode.Text, TBWireDiameter.Text, TBWireLengthMin.Text, wireType)
        Return True
    End Function
    Public Async Function LoadOrdersData() As Task
        '' This function load data from emkansanji table in the database into a datatable
        ''      it then bind that data table to bs2 ( a global binding source in this form) then add bs2 as the data source 
        ''      for datagridview2. 
        ''      This way there wont be a call to data base for each search. we just call the database once then use the 
        ''      binding source for filtering -> less data transfer, probably faster and we can do it on textbox.changed
        ''      because of the increased speed.

        'Dim sql_command = "SELECT " & ESColumnNames & " FROM ((emkansanji INNER JOIN springDataBase ON emkansanji.productID = springDataBase.ID)
        '                    INNER JOIN customers ON emkansanji.customerID = customers.ID) WHERE
        '                    emkansanji.orderState LIKE '%امکان سنجی%' OR emkansanji.orderState LIKE '%تایید%'  ;"

        Dim sql_command = "SELECT " & ESColumnNames & " FROM ((emkansanji LEFT JOIN springDataBase ON emkansanji.productID = springDataBase.ID)
                            LEFT JOIN customers ON emkansanji.customerID = customers.ID) WHERE
                            emkansanji.orderState LIKE '%امکان سنجی%' OR emkansanji.orderState LIKE '%تایید%'  ;"

        Dim dt = Await Task(Of DataTable).Run(Function() LoadDataTable(sql_command))
        bs2.DataSource = dt
        DataGridView2.DataSource = bs2
        'bs2.Filter = ""
        '' Hide values which are not for the user to see
        DataGridView2.Columns(0).Visible = False
        DataGridView2.Columns(1).Visible = False
        DataGridView2.Columns(2).Visible = False
        DataGridView2.Columns(3).Visible = False
        DataGridView2.Columns(4).Visible = False
        DataGridView2.Columns(5).Visible = False
        DataGridView2.Columns(6).Visible = False
    End Function
    Private Function SearchOrdersData()
        bs2.Filter = String.Format("[کد مفتول رزرو 1] LIKE '%{0}%' OR [کد مفتول رزرو 2] LIKE '%{0}%' OR [کد مفتول رزرو 3] LIKE '%{0}%'", TBWireCodeOrderSearch.Text)
    End Function

    Private Async Sub wires_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        firstTimeEnteringOrdersTab = True
        '' What is visible and what is not
        If wiresFormState = "selection" Then
            BTSelectWire.Visible = True
            wiresFormState = "normal"
        End If

        '' Form Load
        Await LoadWiresData()

    End Sub

    Private Async Sub BTUpdateInventory_Click(sender As Object, e As EventArgs) Handles BTUpdateInventory.Click

        '' This subroutine will generate the inventory Database from excel files provided by rahkaran
        '' Requires ImportExceltoDatatable function

        '' import garm inventory data to a data table
        Dim garmDataTable = Await Task(Of DataTable).Run(Function() ImportExceltoDatatable(excelInventoryGarmPath, "موجودی مواد خط گرم"))
        '' import sard inventory data to a data table
        Dim sardDataTable = Await Task(Of DataTable).Run(Function() ImportExceltoDatatable(excelInventorySardPath, "موجودی مواد خط سرد"))
        '' import purchased inventory data to a data table
        Dim purchasedDataTable = Await Task(Of DataTable).Run(Function() ImportExceltoDatatable(excelInventoryPurchasedPath, "موجودی مواد خریداری شده قطعی"))


        Dim wireDiameter As String
        Dim wireLength As String
        Dim wireWeight As String

        Using cn As New OleDbConnection(connectionString)
            Await cn.OpenAsync()
            Using tran = cn.BeginTransaction()
                Using cmd As New OleDbCommand With {.Connection = cn, .Transaction = tran}

                    Try
                        '' Delete everything in wire inventory
                        cmd.CommandText = "DELETE FROM wireInventory"
                        Await cmd.ExecuteNonQueryAsync()

                        '' Populate the inventory table with data of garm file
                        For Each row As DataRow In garmDataTable.Rows
                            wireDiameter = (Val(row("کد").ToString().Substring(2, 4)) / 100).ToString


                            '' Extract wire length from specification column of rahkaran usind regular expressions 
                            Dim RegexObj As New Regex("[L|l]\s*[:|;]?\s*(\d{4})", RegexOptions.IgnoreCase)
                            Dim specification = row("مشخصه فنی").ToString()
                            wireLength = RegexObj.Match(specification).Groups(1).Value

                            '' Check to see if wire is black steel if it is add black to its name
                            Dim inventoryName = row("عنوان").ToString()
                            Dim RegexBbj2 As New Regex(".*(black).*", RegexOptions.IgnoreCase)
                            If RegexBbj2.Match(specification).Groups(1).Value.ToLower = "black" Then
                                inventoryName += " - (Black)"
                            End If

                            If IsNumeric(wireLength) Then
                                wireWeight = CalculateWireWeight(Val(wireDiameter), Val(wireLength)).ToString
                            Else
                                wireLength = "-"
                                wireWeight = "-"
                            End If
                            cmd.CommandText = String.Format("INSERT INTO wireInventory 
                        (wireCode, wireType, wireDiameter, wireLength, wireSpecification, inventory , inventoryName, wireWeight) 
                        VALUES ( '{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}' );", row("کد").ToString(), "مفتول شاخه‌ای",
                                       wireDiameter, wireLength, row("مشخصه فنی").ToString(), row("مانده (اصلی)").ToString(), inventoryName, wireWeight)
                            Await cmd.ExecuteNonQueryAsync()
                        Next row


                        '' Populate the inventory table with data of sard file
                        For Each row As DataRow In sardDataTable.Rows
                            wireDiameter = (Val(row("کد").ToString().Substring(2, 4)) / 100).ToString
                            cmd.CommandText = String.Format("INSERT INTO wireInventory 
                        (wireCode, wireType, wireDiameter, wireLength, wireSpecification, inventory , inventoryName, wireWeight) 
                        VALUES ( '{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}' );", row("کد").ToString(), "مفتول کویل",
                                       wireDiameter, "-", row("مشخصه فنی").ToString(), row("مانده (اصلی)").ToString(), row("عنوان").ToString(), "-")
                            Await cmd.ExecuteNonQueryAsync()
                        Next row


                        '' Populate the inventory table with data of purchased wires file
                        '' TODO: Add wire weight 
                        For Each row As DataRow In purchasedDataTable.Rows

                            If row("کد").ToString() = "" Then
                                '' prevent empty rows in the files to be inserted in the database
                                Continue For
                            End If

                            cmd.CommandText = String.Format("INSERT INTO wireInventory 
                        (wireCode, wireType, wireDiameter, wireLength, wireSpecification, inventory , inventoryName) 
                        VALUES ( '{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}' );", row("کد").ToString(), row("نوع مفتول").ToString(),
                                       row("قطر مفتول").ToString(), row("طول مفتول").ToString(), row("مشخصه فنی").ToString(), row("موجودی").ToString(), row("عنوان").ToString())

                            'Console.WriteLine(cmd.CommandText)
                            Await cmd.ExecuteNonQueryAsync()
                        Next row

                    Catch ex As Exception
                        MsgBox("انتقال اطلاعات موجودی مواد با خطا مواجه شد", vbCritical + MsgBoxStyle.MsgBoxRight, "خطا")
                        Logger.LogFatal(ex.Message, ex)
                        tran.Rollback()
                        cn.Close()
                        Exit Sub
                    End Try

                    tran.Commit()
                    cn.Close()
                    MsgBox("بروزرسانی اطلاعات موجودی مواد با موفقیت انجام شد.", MsgBoxStyle.MsgBoxRtlReading + MsgBoxStyle.MsgBoxRight + vbInformation, "بروزرسانی اطلاعات مواد")
                End Using
            End Using
        End Using
        Await LoadWiresData()
    End Sub



    Private Sub BTSelectWire_Click(sender As Object, e As EventArgs) Handles BTSelectWire.Click
        '' This button is used when this form is called from emkansanji form as is used to fill the form with information
        '' from the selected wire

        Dim selectedWire As String = DataGridView1.SelectedRows(0).Cells("کد کالا").Value.ToString  'TODO: fix this -> add column name
        Dim selectedWireWeight As String = DataGridView1.SelectedRows(0).Cells("wireWeight").Value.ToString
        Dim selectedWireUnit As String
        If IsNumeric(selectedWireWeight) Then
            selectedWireUnit = "شاخه"
        Else
            selectedWireUnit = "کیلوگرم"
            selectedWireWeight = "-"
        End If
        Select Case wireFormCaller
            Case "wire1"
                emkanSanjiForm.TBMR1.Text = selectedWire
                emkanSanjiForm.Lw1Weight.Text = selectedWireWeight
                emkanSanjiForm.Lw1Unit.Text = selectedWireUnit
            Case "wire2"
                emkanSanjiForm.TBMR2.Text = selectedWire
                emkanSanjiForm.Lw2Weight.Text = selectedWireWeight
                emkanSanjiForm.Lw2Unit.Text = selectedWireUnit
            Case "wire3"
                emkanSanjiForm.TBMR3.Text = selectedWire
                emkanSanjiForm.Lw3Weight.Text = selectedWireWeight
                emkanSanjiForm.Lw3Unit.Text = selectedWireUnit
        End Select
        Me.Dispose()
    End Sub

    Private Async Sub TabPageOrders_Enter(sender As Object, e As EventArgs) Handles TabPage_Orders.Enter
        '' Loading Orders data into datagridview2 only if user actually wants orders data (by going to the tab)
        ''      but ensuring that data are loaded only the first time the tab is selected. 
        If firstTimeEnteringOrdersTab = True Then
            Await LoadOrdersData()
            firstTimeEnteringOrdersTab = False
        End If
    End Sub

    Private Sub TBWireCodeOrderSearch_TextChanged(sender As Object, e As EventArgs) Handles TBWireCodeOrderSearch.TextChanged
        SearchOrdersData()
    End Sub

    Private Sub TBWireCode_TextChanged_1(sender As Object, e As EventArgs) Handles TBWireCode.TextChanged
        SearchWiresData()
    End Sub

    Private Sub TBWireDiameter_TextChanged(sender As Object, e As EventArgs) Handles TBWireDiameter.TextChanged
        SearchWiresData()
    End Sub

    Private Sub TBWireLengthMin_TextChanged(sender As Object, e As EventArgs) Handles TBWireLengthMin.TextChanged
        SearchWiresData()
    End Sub

    Private Sub CBWireType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBWireType.SelectedIndexChanged
        SearchWiresData()
    End Sub

    Private Sub BTCheckOrders_Click(sender As Object, e As EventArgs) Handles BTCheckOrders.Click
        TBWireCodeOrderSearch.Text = DataGridView1.SelectedRows(0).Cells("کد کالا").Value.ToString
        TabControl1.SelectedTab = TabPage_Orders
    End Sub

    Private Sub BTCheckOrder_Click(sender As Object, e As EventArgs) Handles BTCheckOrder.Click
        '' Open an instance of emkansanjiForm, then filter the data to only show the order with exact order number selected
        ''      this way we make sure that the selected row in the gridview is exactly the row that we want. consequently we
        ''      can use BTModify button on that form instead of copy pasting the same code here. then we hide the first tab.
        '' The emkansanji form is capable of quering data with condition but the condition on emkansanjiID is LIKE there 
        ''      when we change that to = , the form can't be loaded without a condition(empty textbox). so we use a seperate filter. 

        Dim EmkansanjiPopUp As New emkanSanjiForm
        '' filter the data in emkansanji form using a bindingsource binded to the datagridview1
        EmkansanjiPopUp.emkansanji_bs.Filter = String.Format("[شماره ردیابی سفارش] = '{0}'", DataGridView2.SelectedRows(0).Cells("شماره ردیابی سفارش").Value.ToString)
        EmkansanjiPopUp.Show()
        EmkansanjiPopUp.thisFormsOwner = "wiresForm"
    End Sub

    Private Sub BTShowAllWires_Click(sender As Object, e As EventArgs) Handles BTShowAllWires.Click
        TBWireCode.Text = ""
        TBWireDiameter.Text = ""
        TBWireLengthMin.Text = ""
        CBWireType.Text = "همه"
    End Sub

    Private Sub BTShowAllOrders_Click(sender As Object, e As EventArgs) Handles BTShowAllOrders.Click
        TBWireCodeOrderSearch.Text = ""
        bs2.Filter = ""
    End Sub

    Private Sub TableLayoutPanel2_Paint(sender As Object, e As PaintEventArgs) Handles TableLayoutPanel2.Paint

    End Sub
End Class