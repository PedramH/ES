
Imports System.Text.RegularExpressions

Public Class wires
    Dim firstTimeEnteringOrdersTab = True

    '' Define a global binding source for the data that is comming from the wires database and going into datagridview1 (Wires)   
    Dim bs As New BindingSource
    '' Define a global binding source for the data that is comming from the wires database and going into datagridview2 (Orders)   
    Dim bs2 As New BindingSource

    Public thisFormsEmkansanjiCaller As emkanSanjiForm
    Public Async Function LoadWiresData() As Task
        '' This function load data from wire inventory table in the database into a datatable
        ''      it then bind that data table to bs ( a global binding source in this form) then add bs as the data source 
        ''      for datagridview1. 
        ''      This way there wont be a call to data base for each search. we just call the database once then use the 
        ''      binding source for filtering -> less data transfer, probably faster and we can do it on textbox.changed
        ''      because of the increased speed.
        Dim sql_command = "SELECT" + wiresColumnName + "FROM wireInventory A LEFT JOIN wireReserve B ON A.wireCode = B.wireCode;"
        If db = "postgres" Then
            sql_command = "SELECT A.wireType, A.wireWeight, A.wireCode AS ""کد کالا"", A.inventoryName AS ""عنوان"" , A.wireDiameter AS ""قطر مفتول"", A.wireLength AS ""طول مفتول"" ,
                                        FLOOR((CAST(A.inventory As real) - CAST(B.preReserve As real) - CAST(B.reserve As real))) As ""مانده موجودي (کيلوگرم)"" , (Case When (A.wireWeight ~ '^\d+(\.\d+)?$') THEN CAST(FLOOR((CAST(A.inventory AS real) - CAST (B.preReserve AS real) - CAST(B.reserve AS real)) / CAST(A.wireWeight AS real)) AS varchar) ELSE '-' END ) AS ""تعداد شاخه"" , 
                                         A.inventory AS ""موجودي فيزيکي(کيلوگرم)"", ( CASE WHEN A.wireWeight ~ '^\d+(\.\d+)?$' THEN CAST(FLOOR( CAST (A.inventory AS real) / CAST(A.wireWeight AS real) ) as varchar) ELSE '-' END) AS ""موجودي فيزيکي (تعداد شاخه)"", 
                                         B.preReserve AS ""رزرو امکان سنجي (کيلوگرم)"" , (CASE WHEN A.wireWeight ~ '^\d+(\.\d+)?$' THEN CAST(FLOOR( CAST(B.preReserve AS real) / CAST (A.wireWeight AS real)) AS varchar) ELSE '-' END) AS ""امکان سنجي (تعداد شاخه)"", 
                                         B.reserve AS ""رزرو توليد (کيلوگرم)"" , (CASE WHEN A.wireWeight ~ '^\d+(\.\d+)?$' THEN CAST (FLOOR( CAST(B.reserve AS real) / CAST(A.wireWeight AS real)) AS varchar) ELSE '-' END) AS ""توليد(تعداد شاخه)""   FROM wireInventory A LEFT JOIN wireReserve B ON A.wireCode = B.wireCode;"
        End If
        Try
            Me.Cursor = Cursors.WaitCursor
            Dim dt = Await Task(Of DataTable).Run(Function() LoadDataTable(sql_command))
            bs.DataSource = dt
            DataGridView1.DataSource = bs.DataSource
            '' Hide wireType and wireWeight
            DataGridView1.Columns(0).Visible = False
            DataGridView1.Columns(1).Visible = False
            '' Formating of columns
            DataGridView1.DefaultCellStyle.Font = New System.Drawing.Font("Arial", 9.85)
            DataGridView1.Columns("عنوان").DefaultCellStyle.Font = New System.Drawing.Font("B Traffic", 9.75)
            Me.Cursor = Cursors.Default
        Catch ex As Exception
            Me.Cursor = Cursors.Default
            MsgBox("خطا در ارتباط با دیتابیس", vbCritical + MsgBoxStyle.MsgBoxRight, RightToLeft)
            Logger.LogFatal(sql_command, ex)
        End Try
        Me.Cursor = Cursors.Default
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

        Dim columnNames = " springDataBase.ID AS [productID] , customers.ID AS [customerID] , springDataBase.wireDiameter AS [wireDiameter], springDataBase.OD AS [OD] , springDataBase.L0 AS [L0] , springDataBase.wireLength AS [wireLength], springDataBase.mandrelDiameter AS [mandrelDiameter], emkansanji.ID AS [شماره ردیابی سفارش], springDataBase.productName AS [نام محصول], emkansanji.customerProductName AS [نام محصول مشتری], customers.customerName AS [نام مشتری],  emkansanji.orderState AS [وضعیت سفارش], 
        emkansanji.quantity AS [تعداد سفارش],emkansanji.sampleQuantity AS [تعداد نمونه] , 
        emkansanji.dateOfProccessing As [تاریخ بررسی], emkansanji.mandrelState As [موجودی مندرل], 
        emkansanji.wireState As [وضعیت موجودی مفتول],
        emkansanji.r1_code As [کد مفتول رزرو 1], (Case When ( (SELECT wireWeight FROM wireInventory WHERE wirecode = emkansanji.r1_code) ~ '^\d+(\.\d+)?$') THEN CAST(  ROUND(CAST(emkansanji.r1_q AS real) / CAST((SELECT wireWeight FROM wireInventory WHERE wirecode = emkansanji.r1_code) AS real)) AS varchar) ELSE CONCAT(CAST(emkansanji.r1_q AS varchar), ' Kg') END ) As [مقدار1],
        emkansanji.r2_code As [کد مفتول رزرو 2], (Case When ( (SELECT wireWeight FROM wireInventory WHERE wirecode = emkansanji.r2_code) ~ '^\d+(\.\d+)?$') THEN CAST(  ROUND(CAST(emkansanji.r2_q AS real) / CAST((SELECT wireWeight FROM wireInventory WHERE wirecode = emkansanji.r2_code) AS real)) AS varchar) ELSE CONCAT(CAST(emkansanji.r2_q AS varchar), ' Kg' )END ) As [مقدار 2],
        emkansanji.r3_code As [کد مفتول رزرو 3], (Case When ( (SELECT wireWeight FROM wireInventory WHERE wirecode = emkansanji.r3_code) ~ '^\d+(\.\d+)?$') THEN CAST(  ROUND(CAST(emkansanji.r3_q AS real) / CAST((SELECT wireWeight FROM wireInventory WHERE wirecode = emkansanji.r3_code) AS real)) AS varchar)ELSE CONCAT(CAST(emkansanji.r3_q AS varchar), ' Kg' )END ) As [مقدار 3],
        emkansanji.productReserve,
        emkansanji.verificationNo As [شماره تاییدیه], emkansanji.verificationDate As [تاریخ تاییدیه], emkansanji.comment As [توضیحات]"
        

        columnNames = MigrateAccessToPostgres(columnNames)

        'Dim sql_command = "Select " & columnNames & " FROM ((emkansanji LEFT JOIN springDataBase On emkansanji.productID = springDataBase.ID)
        '                    LEFT JOIN customers ON emkansanji.customerID = customers.ID) WHERE
        '                    emkansanji.orderState LIKE '%امکان سنجی%' OR emkansanji.orderState LIKE '%تایید%'  ;"
        Dim sql_command = "Select " & columnNames & " FROM ((emkansanji LEFT JOIN springDataBase On emkansanji.productID = springDataBase.ID)
                            LEFT JOIN customers ON emkansanji.customerID = customers.ID) ORDER BY emkansanji.id;"
        'Console.WriteLine(sql_command)
        Try
            Dim dt = Await Task(Of DataTable).Run(Function() LoadDataTable(sql_command))
            bs2.DataSource = dt
            DataGridView2.DataSource = bs2
            'bs2.Filter = ""
            '' Hide values which are not for the user to see

            DataGridView2.Columns("productID").Visible = False
            DataGridView2.Columns("customerID").Visible = False
            DataGridView2.Columns("wireDiameter").Visible = False
            DataGridView2.Columns("OD").Visible = False
            DataGridView2.Columns("L0").Visible = False
            DataGridView2.Columns("wireLength").Visible = False
            DataGridView2.Columns("mandrelDiameter").Visible = False
            DataGridView2.Columns("شماره ردیابی سفارش").Visible = False
            DataGridView2.Columns("نام محصول").Visible = False
            'DataGridView2.Columns("pProcess").Visible = False
            'DataGridView2.Columns("productReserve").Visible = False
            'DataGridView2.Columns("productionProcess").Visible = False
            'DataGridView2.Columns("springInEachPackage").Visible = False
            'DataGridView2.Columns("packagingCost").Visible = False
            'DataGridView2.Columns("doable").Visible = False
            'DataGridView2.Columns("whyNot").Visible = False
            'DataGridView2.Columns("productionReserve").Visible = False
        Catch ex As Exception
            MsgBox("خطا در ارتباط با دیتابیس", vbCritical + MsgBoxStyle.MsgBoxRight, RightToLeft)
            Logger.LogFatal(sql_command, ex)
        End Try
        FormatDatagridview2()
    End Function
    Private Sub SearchOrdersData()
        On Error Resume Next
        If TBWireCodeOrderSearch.Text = "" Then
            bs2.Filter = String.Format("( [نام محصول مشتری] LIKE '%{0}%' OR [نام محصول] LIKE '%{0}%' ) AND [نام مشتری] LIKE '%{1}%' ", TBProductNameOrderSearch.Text, TBCustomerNameOrderSearch.Text)
        Else
            bs2.Filter = String.Format("([کد مفتول رزرو 1] LIKE '%{0}%' OR [کد مفتول رزرو 2] LIKE '%{0}%' OR [کد مفتول رزرو 3] LIKE '%{0}%') AND ( [نام محصول مشتری] LIKE '%{1}%' OR [نام محصول] LIKE '%{1}%' ) AND [نام مشتری] LIKE '%{2}%' ", TBWireCodeOrderSearch.Text, TBProductNameOrderSearch.Text, TBCustomerNameOrderSearch.Text)
        End If
        'bs2.Filter = String.Format("[کد مفتول رزرو 1] LIKE '%{0}%' OR [کد مفتول رزرو 2] LIKE '%{0}%' OR [کد مفتول رزرو 3] LIKE '%{0}%'", TBWireCodeOrderSearch.Text)
    End Sub

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
        Me.Cursor = Cursors.WaitCursor
        '' This subroutine will generate the inventory Database from excel files provided by rahkaran
        '' Requires ImportExceltoDatatable function
        Dim garmDataTable, sardDataTable, purchasedDataTable As DataTable
        Try
            '' import garm inventory data to a data table
            garmDataTable = Await Task(Of DataTable).Run(Function() ImportExceltoDatatable(excelInventoryGarmPath, "موجودی مواد خط گرم"))
            '' import sard inventory data to a data table
            sardDataTable = Await Task(Of DataTable).Run(Function() ImportExceltoDatatable(excelInventorySardPath, "موجودی مواد خط سرد"))
            '' import purchased inventory data to a data table
            purchasedDataTable = Await Task(Of DataTable).Run(Function() ImportExceltoDatatable(excelInventoryPurchasedPath, "موجودی مواد خریداری شده قطعی"))
        Catch ex As Exception
            Me.Cursor = Cursors.Default
            MsgBox("خطا در خواندن اطلاعات موجودی مفتول از فایل اکسل", vbCritical + MsgBoxStyle.MsgBoxRight, RightToLeft)
            Logger.LogFatal("خطا در خواندن اطلاعات موجودی مفتول از فایل اکسل", ex)
            Exit Sub
        End Try
        Me.Cursor = Cursors.WaitCursor
        Dim wireDiameter As String
        Dim wireLength As String
        Dim wireWeight As String

        'Using cn As New OleDbConnection(connectionString)
        '    Await cn.OpenAsync()
        '    Using tran = cn.BeginTransaction()
        '        Using cmd As New OleDbCommand With {.Connection = cn, .Transaction = tran}

        '            Try
        '                '' Delete everything in wire inventory
        '                cmd.CommandText = "DELETE FROM wireInventory"
        '                Await cmd.ExecuteNonQueryAsync()

        '                '' Populate the inventory table with data of garm file
        '                For Each row As DataRow In garmDataTable.Rows
        '                    wireDiameter = (Val(row("کد").ToString().Substring(2, 4)) / 100).ToString


        '                    '' Extract wire length from specification column of rahkaran usind regular expressions 
        '                    Dim RegexObj As New Regex("[L|l]\s*[:|;]?\s*(\d{4})", RegexOptions.IgnoreCase)
        '                    Dim specification = row("مشخصه فنی").ToString()
        '                    wireLength = RegexObj.Match(specification).Groups(1).Value

        '                    '' Check to see if wire is black steel if it is add black to its name
        '                    Dim inventoryName = row("عنوان").ToString()
        '                    Dim RegexBbj2 As New Regex(".*(black).*", RegexOptions.IgnoreCase)
        '                    If RegexBbj2.Match(specification).Groups(1).Value.ToLower = "black" Then
        '                        inventoryName += " - (Black)"
        '                    End If

        '                    If IsNumeric(wireLength) Then
        '                        wireWeight = CalculateWireWeight(Val(wireDiameter), Val(wireLength)).ToString
        '                    Else
        '                        wireLength = "-"
        '                        wireWeight = "-"
        '                    End If
        '                    cmd.CommandText = String.Format("INSERT INTO wireInventory 
        '                (wireCode, wireType, wireDiameter, wireLength, wireSpecification, inventory , inventoryName, wireWeight) 
        '                VALUES ( '{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}' );", row("کد").ToString(), "مفتول شاخه‌ای",
        '                               wireDiameter, wireLength, row("مشخصه فنی").ToString(), row("مانده (اصلی)").ToString(), inventoryName, wireWeight)
        '                    Await cmd.ExecuteNonQueryAsync()
        '                Next row


        '                '' Populate the inventory table with data of sard file
        '                For Each row As DataRow In sardDataTable.Rows
        '                    wireDiameter = (Val(row("کد").ToString().Substring(2, 4)) / 100).ToString
        '                    cmd.CommandText = String.Format("INSERT INTO wireInventory 
        '                (wireCode, wireType, wireDiameter, wireLength, wireSpecification, inventory , inventoryName, wireWeight) 
        '                VALUES ( '{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}' );", row("کد").ToString(), "مفتول کویل",
        '                               wireDiameter, "-", row("مشخصه فنی").ToString(), row("مانده (اصلی)").ToString(), row("عنوان").ToString(), "-")
        '                    Await cmd.ExecuteNonQueryAsync()
        '                Next row


        '                '' Populate the inventory table with data of purchased wires file
        '                '' TODO: Add wire weight 
        '                For Each row As DataRow In purchasedDataTable.Rows

        '                    If row("کد").ToString() = "" Then
        '                        '' prevent empty rows in the files to be inserted in the database
        '                        Continue For
        '                    End If
        '                    If IsNumeric(row("طول مفتول").ToString()) Then
        '                        wireWeight = CalculateWireWeight(Val(row("قطر مفتول").ToString()), Val(row("طول مفتول").ToString())).ToString
        '                    Else
        '                        wireLength = "-"
        '                        wireWeight = "-"
        '                    End If


        '                    cmd.CommandText = String.Format("INSERT INTO wireInventory 
        '                (wireCode, wireType, wireDiameter, wireLength, wireSpecification, inventory , inventoryName,wireWeight) 
        '                VALUES ( '{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}','{7}' );", row("کد").ToString(), row("نوع مفتول").ToString(),
        '                               row("قطر مفتول").ToString(), row("طول مفتول").ToString(), row("مشخصه فنی").ToString(), row("موجودی").ToString(), row("عنوان").ToString(), wireWeight)

        '                    'Console.WriteLine(cmd.CommandText)
        '                    Await cmd.ExecuteNonQueryAsync()
        '                Next row

        '            Catch ex As Exception
        '                MsgBox("انتقال اطلاعات موجودی مواد با خطا مواجه شد", vbCritical + MsgBoxStyle.MsgBoxRight, "خطا")
        '                Logger.LogFatal(ex.Message, ex)
        '                tran.Rollback()
        '                cn.Close()
        '                Exit Sub
        '            End Try

        '            tran.Commit()
        '            cn.Close()
        '            MsgBox("بروزرسانی اطلاعات موجودی مواد با موفقیت انجام شد.", MsgBoxStyle.MsgBoxRtlReading + MsgBoxStyle.MsgBoxRight + vbInformation, "بروزرسانی اطلاعات مواد")
        '        End Using
        '    End Using
        'End Using
        Me.Cursor = Cursors.WaitCursor
        Using cn = GetDatabaseCon()
            Await cn.OpenAsync()
            Using tran = cn.BeginTransaction()
                Using cmd = cn.CreateCommand
                    cmd.Transaction = tran
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
                            If IsNumeric(row("طول مفتول").ToString()) Then
                                wireWeight = CalculateWireWeight(Val(row("قطر مفتول").ToString()), Val(row("طول مفتول").ToString())).ToString
                            Else
                                wireLength = "-"
                                wireWeight = "-"
                            End If


                            cmd.CommandText = String.Format("INSERT INTO wireInventory 
                        (wireCode, wireType, wireDiameter, wireLength, wireSpecification, inventory , inventoryName,wireWeight) 
                        VALUES ( '{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}','{7}' );", row("کد").ToString(), row("نوع مفتول").ToString(),
                                       row("قطر مفتول").ToString(), row("طول مفتول").ToString(), row("مشخصه فنی").ToString(), row("موجودی").ToString(), row("عنوان").ToString(), wireWeight)

                            'Console.WriteLine(cmd.CommandText)
                            Await cmd.ExecuteNonQueryAsync()
                        Next row

                    Catch ex As Exception
                        Me.Cursor = Cursors.Default
                        MsgBox("انتقال اطلاعات موجودی مواد با خطا مواجه شد" + ex.Message, vbCritical + MsgBoxStyle.MsgBoxRight, "خطا")
                        Logger.LogFatal(ex.Message, ex)
                        tran.Rollback()
                        cn.Close()
                        Exit Sub
                    End Try
                    tran.Commit()
                    cn.Close()
                    Me.Cursor = Cursors.Default
                    MsgBox("بروزرسانی اطلاعات موجودی مواد با موفقیت انجام شد.", MsgBoxStyle.MsgBoxRtlReading + MsgBoxStyle.MsgBoxRight + vbInformation, "بروزرسانی اطلاعات مواد")
                End Using
            End Using
        End Using
        Me.Cursor = Cursors.WaitCursor
        Await LoadWiresData()
        Me.Cursor = Cursors.WaitCursor
        Await UpdateReservesTable()
        Me.Cursor = Cursors.Default
    End Sub



    Private Sub BTSelectWire_Click(sender As Object, e As EventArgs) Handles BTSelectWire.Click
        '' This button is used when this form is called from emkansanji form as is used to fill the form with information
        '' from the selected wire

        Dim selectedWire As String = DataGridView1.SelectedRows(0).Cells("کد کالا").Value.ToString  'TODO: fix this -> add column name
        Dim selectedWireWeight As String = DataGridView1.SelectedRows(0).Cells("wireWeight").Value.ToString
        Dim selectedWireUnit As String
        Dim selectedWireD = DataGridView1.SelectedRows(0).Cells("قطر مفتول").Value.ToString
        Dim selectedWireL = DataGridView1.SelectedRows(0).Cells("طول مفتول").Value.ToString

        If IsNumeric(selectedWireWeight) Then
            selectedWireUnit = "شاخه"
        Else
            selectedWireUnit = "کیلوگرم"
            selectedWireWeight = "-"
        End If

        'Select Case wireFormCaller
        '    Case "wire1"
        '        emkanSanjiForm.TBMR1.Text = selectedWire
        '        emkanSanjiForm.Lw1Weight.Text = selectedWireWeight
        '        emkanSanjiForm.Lw1Unit.Text = selectedWireUnit
        '        emkanSanjiForm.LSelectedWireD.Text = selectedWireD
        '        emkanSanjiForm.LSelectedWireL.Text = selectedWireL
        '    Case "wire2"
        '        emkanSanjiForm.TBMR2.Text = selectedWire
        '        emkanSanjiForm.Lw2Weight.Text = selectedWireWeight
        '        emkanSanjiForm.Lw2Unit.Text = selectedWireUnit
        '    Case "wire3"
        '        emkanSanjiForm.TBMR3.Text = selectedWire
        '        emkanSanjiForm.Lw3Weight.Text = selectedWireWeight
        '        emkanSanjiForm.Lw3Unit.Text = selectedWireUnit
        'End Select
        Select Case wireFormCaller
            Case "wire1"
                thisFormsEmkansanjiCaller.TBMR1.Text = selectedWire
                thisFormsEmkansanjiCaller.Lw1Weight.Text = selectedWireWeight
                thisFormsEmkansanjiCaller.Lw1Unit.Text = selectedWireUnit
                thisFormsEmkansanjiCaller.LSelectedWireD.Text = selectedWireD
                thisFormsEmkansanjiCaller.LSelectedWireL.Text = selectedWireL
            Case "wire2"
                thisFormsEmkansanjiCaller.TBMR2.Text = selectedWire
                thisFormsEmkansanjiCaller.Lw2Weight.Text = selectedWireWeight
                thisFormsEmkansanjiCaller.Lw2Unit.Text = selectedWireUnit
            Case "wire3"
                thisFormsEmkansanjiCaller.TBMR3.Text = selectedWire
                thisFormsEmkansanjiCaller.Lw3Weight.Text = selectedWireWeight
                thisFormsEmkansanjiCaller.Lw3Unit.Text = selectedWireUnit
        End Select
        Me.Dispose()
    End Sub

    Private Async Sub TabPageOrders_Enter(sender As Object, e As EventArgs) Handles TabPage_Orders.Enter
        '' Loading Orders data into datagridview2 only if user actually wants orders data (by going to the tab)
        ''      but ensuring that data are loaded only the first time the tab is selected. 
        If firstTimeEnteringOrdersTab = True Then
            Await LoadOrdersData()
            'LoadOrdersData()
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


    Private Sub DataGridView1_CellPainting(sender As Object, e As DataGridViewCellPaintingEventArgs) Handles DataGridView1.CellPainting
        If (e.ColumnIndex = 6 Or e.ColumnIndex = 7) And e.RowIndex >= 0 And IsNumeric(e.Value) Then
            If e.Value < 0 Then
                e.PaintBackground(e.CellBounds, True)
                TextRenderer.DrawText(e.Graphics, e.FormattedValue.ToString(),
              e.CellStyle.Font, e.CellBounds, e.CellStyle.ForeColor,
               TextFormatFlags.HorizontalCenter)
                e.Handled = True
            End If

        End If
        For Each row As DataGridViewRow In DataGridView1.Rows()
            On Error Resume Next '' TODO What the actual fuck
            Dim wire = row.Cells("عنوان").Value().ToString()
            If wire.Contains("Black") Then
                row.Cells("عنوان").Style.BackColor = Color.Black
                row.Cells("عنوان").Style.ForeColor = Color.White
            ElseIf wire.Contains("خرید") Then
                row.Cells("عنوان").Style.BackColor = Color.LawnGreen
            End If
            Dim inv = row.Cells("مانده موجودي (کيلوگرم)").Value().ToString()
            If inv < 0 Then
                row.Cells("مانده موجودي (کيلوگرم)").Style.ForeColor = Color.Red
                row.Cells("تعداد شاخه").Style.ForeColor = Color.Red
            End If
        Next
    End Sub

    Private Sub DataGridView2_CellPainting(sender As Object, e As DataGridViewCellPaintingEventArgs) Handles DataGridView2.CellPainting
        For Each row As DataGridViewRow In DataGridView2.Rows()
            Dim orderState = row.Cells("وضعیت سفارش").Value().ToString()
            If orderState = "تایید شده" Then
                row.DefaultCellStyle.BackColor = System.Drawing.Color.LawnGreen
            ElseIf orderState = "تولید شده" Then
                row.DefaultCellStyle.BackColor = Color.Black
                row.DefaultCellStyle.ForeColor = Color.White
            ElseIf orderState = "منقضی شده" Then
                row.DefaultCellStyle.BackColor = Color.Tomato
            Else
                row.DefaultCellStyle.BackColor = Color.Khaki
            End If
        Next

    End Sub

    Private Sub FormatDatagridview2()
        For Each row As DataGridViewRow In DataGridView2.Rows()
            Dim orderState = row.Cells("وضعیت سفارش").Value().ToString()
            If orderState = "تایید شده" Then
                row.DefaultCellStyle.BackColor = System.Drawing.Color.LawnGreen
            ElseIf orderState = "تولید شده" Then
                row.DefaultCellStyle.BackColor = Color.Black
                row.DefaultCellStyle.ForeColor = Color.White
            ElseIf orderState = "منقضی شده" Then
                row.DefaultCellStyle.BackColor = Color.Tomato
            Else
                row.DefaultCellStyle.BackColor = Color.Khaki
            End If
        Next
    End Sub

    Private Sub TBProductNameOrderSearch_TextChanged(sender As Object, e As EventArgs) Handles TBProductNameOrderSearch.TextChanged
        SearchOrdersData()
    End Sub

    Private Sub TBCustomerNameOrderSearch_TextChanged(sender As Object, e As EventArgs) Handles TBCustomerNameOrderSearch.TextChanged
        SearchOrdersData()
    End Sub
End Class