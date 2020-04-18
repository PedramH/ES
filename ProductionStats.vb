Imports System.Text.RegularExpressions
Public Class ProductionStats
    Public statsCol As New Collection '' An array of checkBoxes 

    Public emkansanji_bs As New BindingSource
    Public spring_bs As New BindingSource

    Public currentStatIsSubmitted = False
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

    Private Sub AnyTextBox_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Cast the 'sender' object into a TextBox (we are sure it is a textbox!)
        'Dim txt As TextBox = DirectCast(sender, TextBox)
        'MessageBox.Show(txt.Name)
        CheckStatsValidity()
    End Sub
    Private Sub ProductionStats_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        statsCol.Add(fanarpichi)
        statsCol.Add(zfanarpichi)
        statsCol.Add(temper)
        statsCol.Add(ztemper)
        statsCol.Add(dtemper)
        statsCol.Add(shot)
        statsCol.Add(zshot)
        statsCol.Add(dshot)
        statsCol.Add(sett)
        statsCol.Add(zset)
        statsCol.Add(dset)
        statsCol.Add(sang)
        statsCol.Add(zsang)
        statsCol.Add(dsang)
        statsCol.Add(tarak)
        statsCol.Add(ztarak)
        statsCol.Add(dtarak)
        statsCol.Add(rang)
        statsCol.Add(drang)
        statsCol.Add(bastebandi)
        statsCol.Add(ersal)
        '' ----------------------------
        'Loop through your list of textboxes and add eventhandlers
        For Each txt As TextBox In statsCol
            AddHandler txt.Leave, AddressOf AnyTextBox_Leave
        Next


        LoadEmkansanjiTable()
        LoadProductData()

        '' -------------------------------
        Dim pc As New Globalization.PersianCalendar
        current_year.Text = pc.GetYear(Now)
        CBMonth.Text = getMonthName(pc.GetMonth(Now))
        current_day.Text = pc.GetDayOfMonth(Now)
        CBShift.Text = "A"
        'Dim excelFiledir As String = pc.GetYear(Now).ToString & "\" & getMonthName(pc.GetMonth(Now)) & "\" & DataGridView2.SelectedRows(0).Cells("نام مشتری").Value.ToString & "\"
    End Sub
    Private Sub CheckStatsValidity()
        For Each tb As TextBox In statsCol
            If IsNumeric(tb.Text) = False Then
                tb.Text = "0"
            End If
            tb.BackColor = Color.White
        Next
        Dim prevNotComplete = Val(Lnimsakht.Text)
        Dim prevComplete = Val(Ltakmil.Text)
        Dim totalProductionLoss = Val(zfanarpichi.Text) + Val(ztemper.Text) + Val(zshot.Text) + Val(zset.Text) + Val(zsang.Text) + Val(ztarak.Text)
        Dim currentNotComplete = prevNotComplete + Val(fanarpichi.Text) - totalProductionLoss - Val(bastebandi.Text)
        Dim currentComplete = prevComplete + Val(bastebandi.Text) - Val(ersal.Text)
        Dim currentRemainingOrderQuantity = Val(LRemainingOrderQuantity.Text) - Val(ersal.Text)
        nimsakht.Text = currentNotComplete
        takmil.Text = currentComplete
        mandesefaresh.Text = currentRemainingOrderQuantity

        If currentNotComplete < 0 Then
            If prevNotComplete + Val(fanarpichi.Text) - totalProductionLoss < 0 Then
                Dim tempval = prevNotComplete + Val(fanarpichi.Text)
                If tempval - Val(zfanarpichi.Text) < 0 Then
                    zfanarpichi.BackColor = System.Drawing.Color.PaleVioletRed
                ElseIf tempval - Val(zfanarpichi.Text) - Val(ztemper.Text) < 0 Then
                    ztemper.BackColor = System.Drawing.Color.PaleVioletRed
                ElseIf tempval - Val(zfanarpichi.Text) - Val(ztemper.Text) - Val(zshot.Text) < 0 Then
                    'shot
                    zshot.BackColor = System.Drawing.Color.PaleVioletRed
                ElseIf tempval - Val(zfanarpichi.Text) - Val(ztemper.Text) - Val(zshot.Text) - Val(zset.Text) < 0 Then
                    'set
                    zset.BackColor = System.Drawing.Color.PaleVioletRed
                ElseIf tempval - Val(zfanarpichi.Text) - Val(ztemper.Text) - Val(zshot.Text) - Val(zset.Text) - Val(zsang.Text) < 0 Then
                    'sang
                    zsang.BackColor = System.Drawing.Color.PaleVioletRed
                ElseIf tempval - Val(zfanarpichi.Text) - Val(ztemper.Text) - Val(zshot.Text) - Val(zset.Text) - Val(zsang.Text) - Val(tarak.Text) < 0 Then
                    'tarak
                    ztarak.BackColor = System.Drawing.Color.PaleVioletRed
                End If
            Else
                bastebandi.BackColor = System.Drawing.Color.PaleVioletRed
            End If
        End If
        If currentComplete < 0 Then
            ersal.BackColor = System.Drawing.Color.PaleVioletRed
        End If

    End Sub

    Private Async Sub LoadEmkansanjiTable()
        '' -------------------------------------------------------------------------------------------------------------------------------------------
        '' --------------------------------------------------------  Load Emkansanji Table -----------------------------------------------------------
        '' -------------------------------------------------------------------------------------------------------------------------------------------
        Dim columnNames = "springDataBase.ID AS [productID] , emkansanji.ID AS [شماره ردیابی سفارش], springDataBase.productName AS [نام محصول], emkansanji.customerProductName AS [نام محصول مشتری],
                            customers.customerName AS [نام مشتری], emkansanji.orderState AS [وضعیت سفارش], emkansanji.quantity AS [تعداد سفارش],emkansanji.sampleQuantity AS [تعداد نمونه], emkansanji.shipped AS [ارسال شده], (CAST(emkansanji.quantity AS integer) - emkansanji.shipped) AS [مانده سفارش],
                            not_complete_inv AS [موجودی نیم‌ساخت] , complete_inv AS [موجودی تکمیل شده] "
        Dim sql_command = "SELECT " & columnNames & " FROM ((emkansanji LEFT JOIN springDataBase ON emkansanji.productID = springDataBase.ID) LEFT JOIN customers ON emkansanji.customerID = customers.ID) WHERE " &
                                 " springDataBase.productName LIKE '%" & TBOrderProductSearch.Text & "%' AND" &
                                 " customers.customerName LIKE '%" & TBOrderCustomerSearch.Text & "%'" &
                                 " ORDER BY emkansanji.ID ;"
        sql_command = MigrateAccessToPostgres(sql_command)
        Try
            Me.Cursor = Cursors.WaitCursor

            Dim dt = Await Task(Of DataTable).Run(Function() LoadDataTable(sql_command))
            Me.Cursor = Cursors.Default

            emkansanji_bs.DataSource = dt
            DataGridView1.DataSource = emkansanji_bs
            'bs2.Filter = ""
            ' Hide values which are not for the user to see
            'DataGridView1.Columns("productionReserve").Visible = False
            Me.Cursor = Cursors.Default
        Catch ex As Exception
            Me.Cursor = Cursors.Default
            MsgBox("خطا در برقرای ارتباط با دیتابیس", vbCritical + RightToLeft + vbMsgBoxRight, "خطا")
            Logger.LogFatal(sql_command, ex)
        End Try

    End Sub
    Private Sub BTInputStatsForOrder_Click(sender As Object, e As EventArgs) Handles BTInputStatsForOrder.Click
        LProductID.Text = DataGridView1.SelectedRows(0).Cells("productID").Value.ToString()
        LproductName.Text = DataGridView1.SelectedRows(0).Cells("نام محصول").Value.ToString()
        LCustomerName.Text = DataGridView1.SelectedRows(0).Cells("نام مشتری").Value.ToString()
        LOrderID.Text = DataGridView1.SelectedRows(0).Cells("شماره ردیابی سفارش").Value.ToString()
        'TBOrderID.Text = DataGridView1.SelectedRows(0).Cells("شماره ردیابی سفارش").Value.ToString()
        'TBOrderID.ReadOnly = True
        LRemainingOrderQuantity.Text = DataGridView1.SelectedRows(0).Cells("مانده سفارش").Value.ToString()
        LTotalOrderQuantity.Text = DataGridView1.SelectedRows(0).Cells("تعداد سفارش").Value.ToString()
        Lnimsakht.Text = DataGridView1.SelectedRows(0).Cells("موجودی نیم‌ساخت").Value.ToString()
        Ltakmil.Text = DataGridView1.SelectedRows(0).Cells("موجودی تکمیل شده").Value.ToString()


        TabControl1.SelectedTab = TabPage3

    End Sub

    Private Async Sub LoadProductData()
        'Loading Springs table into datagridview1
        Dim columnNames As String = " ID AS [productID], productName AS [نام محصول], productID AS [کد کالا],  not_complete_inv AS [موجودی نیم‌ساخت] , complete_inv AS [موجودی تکمیل شده] ,pType AS [نوع فنر] ,productionMethod AS [روش تولید] ,wireDiameter AS [قطر مفتول], " &
        " OD AS [قطر خارجی], L0 AS [طول آزاد], Nt AS [حلقه کل], Nactive AS [حلقه فعال], coilingDirection AS [جهت پیچش], " &
        " mandrelDiameter AS [قطر شفت], wireLength AS [طول مفتول] , productionProcess AS [productionProcess] "

        Dim sql_command = "SELECT " & columnNames & " FROM springDataBase;"

        sql_command = MigrateAccessToPostgres(sql_command)
        Dim springdt As New DataTable
        Try
            springdt = Await Task(Of DataTable).Run(Function() LoadDataTable(sql_command))
            spring_bs.DataSource = springdt
            DataGridView2.DataSource = spring_bs
        Catch ex As Exception
            MsgBox("خطا در ارتباط با دیتابیس" + Environment.NewLine + ex.Message, vbCritical + vbMsgBoxRight, "خطا")
            Logger.LogFatal("spring data base couldn't be loaded.", ex)
        End Try
    End Sub
    Private Sub BTInputStatsForProduct_Click(sender As Object, e As EventArgs) Handles BTInputStatsForProduct.Click
        LProductID.Text = DataGridView2.SelectedRows(0).Cells("productID").Value.ToString()
        LproductName.Text = DataGridView2.SelectedRows(0).Cells("نام محصول").Value.ToString()
        'LCustomerName.Text = DataGridView1.SelectedRows(0).Cells("نام مشتری").Value.ToString()
        'LOrderID.Text = DataGridView1.SelectedRows(0).Cells("شماره ردیابی سفارش").Value.ToString()
        'TBOrderID.Text = DataGridView1.SelectedRows(0).Cells("شماره ردیابی سفارش").Value.ToString()
        'LRemainingOrderQuantity.Text = DataGridView1.SelectedRows(0).Cells("مانده سفارش").Value.ToString()
        'LTotalOrderQuantity.Text = DataGridView1.SelectedRows(0).Cells("تعداد سفارش").Value.ToString()
        Lnimsakht.Text = DataGridView1.SelectedRows(0).Cells("موجودی نیم‌ساخت").Value.ToString()
        Ltakmil.Text = DataGridView1.SelectedRows(0).Cells("موجودی تکمیل شده").Value.ToString()
        TabControl1.SelectedTab = TabPage3

        LOrderID.Text = "-"
        'TBOrderID.Text = "-"
        LCustomerName.Text = "-"
    End Sub

    Private Sub SearchForProduct()
        Dim searchTerm = ConstructSearchQuery("نام محصول", TBProductName.Text)
        'Dim filter = String.Format("[نام محصول] LIKE '%{0}%'", searchTerm)
        'Console.WriteLine(searchTerm)
        searchTerm = searchTerm & " AND " & ConstructSearchQuery("کد کالا", TBProductCode.Text) _
                                & " AND " & ConstructSearchQuery("قطر مفتول", TBProductWireD.Text)
        spring_bs.Filter = searchTerm
    End Sub

    Private Sub SearchForOrder()
        Dim searchTerm = ConstructSearchQuery("نام محصول", TBOrderProductSearch.Text)
        'Dim filter = String.Format("[نام محصول] LIKE '%{0}%'", searchTerm)
        'Console.WriteLine(searchTerm)
        searchTerm = "( " & searchTerm
        searchTerm = searchTerm & " OR " & ConstructSearchQuery("نام محصول مشتری", TBOrderProductSearch.Text) & " )" _
                                & " AND " & ConstructSearchQuery("نام مشتری", TBOrderCustomerSearch.Text)
        If IsNumeric(TBSearchOrderNo.Text) Then
            searchTerm = searchTerm & " AND [شماره ردیابی سفارش] = " & TBSearchOrderNo.Text
        End If
        emkansanji_bs.Filter = searchTerm

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

    Private Sub TBProductName_TextChanged(sender As Object, e As EventArgs) Handles TBProductName.TextChanged
        SearchForProduct()
    End Sub

    Private Sub TBProductCode_TextChanged(sender As Object, e As EventArgs) Handles TBProductCode.TextChanged
        SearchForProduct()
    End Sub

    Private Sub TBProductWireD_TextChanged(sender As Object, e As EventArgs) Handles TBProductWireD.TextChanged
        SearchForProduct()
    End Sub

    Private Sub TBOrderProductSearch_TextChanged(sender As Object, e As EventArgs) Handles TBOrderProductSearch.TextChanged
        SearchForOrder()
    End Sub

    Private Sub TBOrderCustomerSearch_TextChanged(sender As Object, e As EventArgs) Handles TBOrderCustomerSearch.TextChanged
        SearchForOrder()

    End Sub

    Private Sub TBSearchOrderNo_TextChanged(sender As Object, e As EventArgs) Handles TBSearchOrderNo.TextChanged
        SearchForOrder()
    End Sub

    Private Sub BTShowAll_Click(sender As Object, e As EventArgs) Handles BTShowAll.Click
        TBOrderProductSearch.Text = ""
        TBOrderCustomerSearch.Text = ""
        TBSearchOrderNo.Text = ""
        emkansanji_bs.Filter = ""
    End Sub

    Private Sub BTShowAllProducts_Click(sender As Object, e As EventArgs) Handles BTShowAllProducts.Click
        TBProductName.Text = ""
        TBProductCode.Text = ""
        TBProductWireD.Text = ""
        spring_bs.Filter = ""
    End Sub

    Private Async Sub submitstats_Click(sender As Object, e As EventArgs) Handles submitstats.Click
        CheckStatsValidity()

        If currentStatIsSubmitted = True Then
            MsgBox("آمار فعلی قبلا ثبت شده است", vbCritical + MsgBoxStyle.MsgBoxRight, "خطا")
            Exit Sub
        End If
        Me.Cursor = Cursors.WaitCursor()
        Dim orderid = ""
        If IsNumeric(LOrderID.Text) = False Then
            orderid = "0"
        Else
            orderid = LOrderID.Text
        End If

        If ersal.Text <> "" And IsNumeric(LOrderID.Text) = False Then
            MsgBox("ثبت آمار ارسال محصول بدون شماره ردیابی سفارش امکان پذیر نیست.", vbCritical + MsgBoxStyle.MsgBoxRight, "خطا")
            Me.Cursor = Cursors.Default()
            Exit Sub
        End If

        If IsNumeric(current_day.Text) = False Or CBMonth.Text = "" Or IsNumeric(current_year.Text) = False Then
            MsgBox("تاریخ امار ورودی به درستی وارد نشده است.", vbCritical + MsgBoxStyle.MsgBoxRight, "خطا")
            Me.Cursor = Cursors.Default()
            Exit Sub
        End If

        If Val(nimsakht.Text) < 0 Or Val(takmil.Text) < 0 Or Val(mandesefaresh.Text) < 0 Then
            Dim returnVal = MsgBox("در صورت ثبت آمار وارد شده، در موجودی محصولات یا وضعیت سفارشات مغایرت به وجود خواهد آمد. آمار ثبت شود؟", vbCritical + MsgBoxStyle.MsgBoxRight + vbYesNo, "مغایرت آمار")
            If returnVal = vbNo Then
                MsgBox("فرایند ثبت آمار لغو شد", vbInformation + MsgBoxStyle.MsgBoxRight, "لغو ثبت آمار")
                Me.Cursor = Cursors.Default()
                Exit Sub
            End If
        End If

        Using cn = GetDatabaseCon()
            Await cn.OpenAsync()
            Using tran = cn.BeginTransaction()
                Using cmd = cn.CreateCommand
                    cmd.Transaction = tran
                    Try
                        '' Add an entry to the productionstats table
                        cmd.CommandText = String.Format("INSERT INTO productionstats (product_id, order_id, current_day, current_month, current_year, current_shift, fanarpichi, zfanarpichi, temper, ztemper,dtemper,
                        shot, zshot, dshot, set, zset, dset , sang, zsang, dsang, tarak,ztarak,dtarak, rang,drang, bastebandi,ersal)
                                            VALUES ({0},{1},{2},'{3}','{4}','{5}',{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19},{20},{21},{22},{23},{24},{25}
                        , {26} );", LProductID.Text, orderid, current_day.Text, CBMonth.Text, current_year.Text, CBShift.Text, fanarpichi.Text, zfanarpichi.Text, temper.Text, ztemper.Text, dtemper.Text,
                        shot.Text, zshot.Text, dshot.Text, sett.Text, zset.Text, dset.Text, sang.Text, zsang.Text, dsang.Text, tarak.Text, ztarak.Text, dtarak.Text, rang.Text, drang.Text, bastebandi.Text, ersal.Text)
                        Console.WriteLine(cmd.CommandText)
                        Await cmd.ExecuteNonQueryAsync()

                        '' Update inventory of product in springdatabase table
                        cmd.CommandText = String.Format("UPDATE springdatabase SET not_complete_inv = {0}, complete_inv = {1} WHERE id = {2};", nimsakht.Text, takmil.Text, LProductID.Text)
                        Console.WriteLine(cmd.CommandText)
                        Await cmd.ExecuteNonQueryAsync()


                        '' Update order data
                        If ersal.Text <> "" And IsNumeric(LOrderID.Text) Then
                            cmd.CommandText = String.Format("UPDATE emkansanji SET shipped = shipped + {0} WHERE id = {1};", ersal.Text, LOrderID.Text)
                            Console.WriteLine(cmd.CommandText)
                            Await cmd.ExecuteNonQueryAsync()

                        End If


                    Catch ex As Exception
                        MsgBox("ثبت اطلاعات آمار تولید با خطا مواجه شد" + ex.Message, vbCritical + MsgBoxStyle.MsgBoxRight, "خطا")
                        Logger.LogFatal(ex.Message, ex)
                        tran.Rollback()
                        cn.Close()
                        Me.Cursor = Cursors.Default()
                        Exit Sub
                    End Try
                    tran.Commit()
                    cn.Close()
                    Me.Cursor = Cursors.Default()
                    MsgBox("آمار تولید با موفقیت ثبت شد.", MsgBoxStyle.MsgBoxRtlReading + MsgBoxStyle.MsgBoxRight + vbInformation, "آمار تولید")
                    currentStatIsSubmitted = True
                End Using
            End Using
        End Using
    End Sub

    Private Sub BTClear_Click(sender As Object, e As EventArgs) Handles BTClear.Click
        For Each tb As TextBox In statsCol
            tb.Text = ""
        Next
        currentStatIsSubmitted = False
    End Sub
End Class