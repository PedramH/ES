Imports System.ComponentModel
Imports System.Configuration

Imports Excel = Microsoft.Office.Interop.Excel
Public Class emkanSanjiForm
    'binding source for the data gridview data
    Public emkansanji_bs As New BindingSource
    Public thisFormsOwner As String
    Dim tabsHidden As Boolean

    Public CBA As New Collection '' An array of checkBoxes 
    Public CBA_inspection As New Collection '' An array of checkBoxes for inspection
    Protected Overrides Function ProcessCmdKey(ByRef msg As System.Windows.Forms.Message,
                                           ByVal keyData As System.Windows.Forms.Keys) _
                                           As Boolean
        ' This code send Tab key everytime Enterkey is pressed INSIDE OF A TEXTBOX

        If msg.WParam.ToInt32() = CInt(Keys.Enter) AndAlso (TypeOf Me.ActiveControl Is TextBox Or TypeOf Me.ActiveControl Is ComboBox) Then
            If Me.ActiveControl.Name <> "TBComment" Then
                SendKeys.Send("{Tab}")
                Return True
            End If
        End If

        Return MyBase.ProcessCmdKey(msg, keyData)
    End Function

    Private Async Sub emkanSanjiForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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

        ' Hide the tabs which are for editing before user actualy wants to edit something
        TabControl1.TabPages.Remove(TabPage2)
        TabControl1.TabPages.Remove(TabPage3)
        TabControl1.TabPages.Remove(TabPage4)
        tabsHidden = True

        'Loading Springs table into datagridview1
        Await LoadEmkansanjiTable()
        If thisFormsOwner = "wiresForm" Then
            Me.BTModify.PerformClick()
            Me.TabControl1.TabPages.Remove(Me.TabPage1)
        End If

    End Sub



    Private Async Sub BTModify_Click(sender As Object, e As EventArgs) Handles BTModify.Click
        Try
            If tabsHidden = True Then
                '' Unhide the tabs for editing emkansanji data
                TabControl1.TabPages.Add(TabPage2)
                TabControl1.TabPages.Add(TabPage3)
                TabControl1.TabPages.Add(TabPage4)
                tabsHidden = False
            End If

            'TBCustomerName.Text = DataGridView1.SelectedRows(0).Cells("نام مشتری").Value.ToString
            LemkansanjiID.Text = DataGridView1.SelectedRows(0).Cells("شماره ردیابی سفارش").Value.ToString
            TBMEnergySazProductName.Text = DataGridView1.SelectedRows(0).Cells("نام محصول").Value.ToString
            TBMCustomerProductName.Text = DataGridView1.SelectedRows(0).Cells("نام محصول مشتری").Value.ToString
            TBProductIDES.Text = DataGridView1.SelectedRows(0).Cells(0).Value.ToString
            TBCustomerID.Text = DataGridView1.SelectedRows(0).Cells(1).Value.ToString
            TBMCustomerName.Text = DataGridView1.SelectedRows(0).Cells("نام مشتری").Value.ToString
            TBMCustomerDwgNo.Text = DataGridView1.SelectedRows(0).Cells("شماره نقشه").Value.ToString
            TBMCustomerProductCode.Text = DataGridView1.SelectedRows(0).Cells("کد قطعه مشتری").Value.ToString
            TBMLetterNo.Text = DataGridView1.SelectedRows(0).Cells("شماره نامه").Value.ToString
            TBMLetterDate.Text = DataGridView1.SelectedRows(0).Cells("تاریخ نامه").Value.ToString
            TBMProccessingDate.Text = DataGridView1.SelectedRows(0).Cells("تاریخ بررسی").Value.ToString
            CBStandard.Text = DataGridView1.SelectedRows(0).Cells("استاندارد").Value.ToString

            TBGrade.Text = DataGridView1.SelectedRows(0).Cells("گرید").Value.ToString
            TBQuantity.Text = DataGridView1.SelectedRows(0).Cells("تعداد سفارش").Value.ToString
            TBSampleQuantity.Text = DataGridView1.SelectedRows(0).Cells("تعداد نمونه").Value.ToString
            TBMOrderNo.Text = DataGridView1.SelectedRows(0).Cells("شماره سفارش").Value.ToString
            TBComment.Text = DataGridView1.SelectedRows(0).Cells("توضیحات").Value.ToString


            TBMR1.Text = DataGridView1.SelectedRows(0).Cells("کد مفتول رزرو 1").Value.ToString
            TBMR2.Text = DataGridView1.SelectedRows(0).Cells("کد مفتول رزرو 2").Value.ToString
            TBMR3.Text = DataGridView1.SelectedRows(0).Cells("کد مفتول رزرو 3").Value.ToString

            TBMRQ1.Text = DataGridView1.SelectedRows(0).Cells("مقدار1").Value.ToString
            TBMRQ2.Text = DataGridView1.SelectedRows(0).Cells("مقدار 2").Value.ToString
            TBMRQ3.Text = DataGridView1.SelectedRows(0).Cells("مقدار 3").Value.ToString

            TBMVerificationDate.Text = DataGridView1.SelectedRows(0).Cells("تاریخ تاییدیه").Value.ToString
            TBMVerificationNo.Text = DataGridView1.SelectedRows(0).Cells("شماره تاییدیه").Value.ToString
            CBMOrderState.Text = DataGridView1.SelectedRows(0).Cells("وضعیت سفارش").Value.ToString

            LOrderQuantity.Text = DataGridView1.SelectedRows(0).Cells("تعداد سفارش").Value.ToString
            LWireOrderQuantity.Text = DataGridView1.SelectedRows(0).Cells("تعداد سفارش").Value.ToString
            LQuantityDetail.Text = String.Format("{1} عدد نمونه و {0} عدد انبوه", DataGridView1.SelectedRows(0).Cells("تعداد سفارش").Value.ToString, DataGridView1.SelectedRows(0).Cells("تعداد نمونه").Value.ToString)

            LOutsideDiameter.Text = DataGridView1.SelectedRows(0).Cells("OD").Value.ToString
            LFreeLength.Text = DataGridView1.SelectedRows(0).Cells("L0").Value.ToString
            LWireDiameter.Text = DataGridView1.SelectedRows(0).Cells("wireDiameter").Value.ToString
            LWireLength.Text = DataGridView1.SelectedRows(0).Cells("wireLength").Value.ToString
            LMandrelDiameter.Text = DataGridView1.SelectedRows(0).Cells("mandrelDiameter").Value.ToString

            Dim wireState As String = DataGridView1.SelectedRows(0).Cells("وضعیت موجودی مفتول").Value.ToString
            TBProductReserve.Text = DataGridView1.SelectedRows(0).Cells("productReserve").Value.ToString

            '' product Weight 
            LProductWeight.Text = CalculateWireWeight(Val(LWireDiameter.Text), Val(LWireLength.Text))

            '' packaging
            TBPackageCount.Text = DataGridView1.SelectedRows(0).Cells("springInEachPackage").Value.ToString
            TBPCostForEach.Text = DataGridView1.SelectedRows(0).Cells("packagingCost").Value.ToString
            Dim packageType = DataGridView1.SelectedRows(0).Cells("packageType").Value.ToString
            If packageType = "پالت" Then
                Rpack1.Checked = True
            ElseIf packageType = "کارتن" Then
                Rpack2.Checked = True
            Else
                Rpack1.Checked = False
                Rpack2.Checked = False
            End If


            '' Populate the production process check boxes
            Dim processCode = ""
            If Len(DataGridView1.SelectedRows(0).Cells("pProcess").Value.ToString) > 1 Then
                processCode = DataGridView1.SelectedRows(0).Cells("pProcess").Value.ToString 'from emkansanji
            Else
                processCode = DataGridView1.SelectedRows(0).Cells("productionProcess").Value.ToString ' from springDataBase
            End If
            ParseProductionProcess(CBA, processCode)

            '' Populate the inspectionProcess check boxes
            ParseInspectionProcess(CBA_inspection, DataGridView1.SelectedRows(0).Cells("inspectionProcess").Value.ToString)

            '' Populate orderData 
            Dim orderType = DataGridView1.SelectedRows(0).Cells("orderType").Value.ToString
            If Len(orderType) = 2 Then
                If orderType(0) = "1" Then
                    RBNewProduct.Checked = True
                ElseIf orderType(0) = "2" Then
                    RBChangeProduct.Checked = True
                Else
                    RBOldProduct.Checked = True
                End If

                If orderType(1) = "1" Then
                    RBMainOrder.Checked = True
                Else
                    RBAmendOrder.Checked = True
                End If
            End If

            '' Buy Wire 
            Dim buyWireStr As String = DataGridView1.SelectedRows(0).Cells("buyWire").Value.ToString
            If Len(buyWireStr) > 0 Then
                Dim buyWireArray As String() = buyWireStr.Split(New Char() {"-"c})
                TBBuyWireD.Text = buyWireArray(0)
                TBBuyWireLength.Text = buyWireArray(1)
                TBPillCost.Text = buyWireArray(2)
            End If

            '' Buy Mandrel 
            Dim buyMandrelStr As String = DataGridView1.SelectedRows(0).Cells("buyMandrel").Value.ToString
            ' Console.WriteLine(buyMandrelStr)
            If Len(buyMandrelStr) > 0 Then
                Dim buyMandrelArray As String() = buyMandrelStr.Split(New Char() {"-"c})
                TBBuyMandrelD.Text = buyMandrelArray(0)
                TBBuyMandrelL.Text = buyMandrelArray(1)
                TBBuyMandrelPrice.Text = buyMandrelArray(2)
                TBBuyMandrelCost.Text = buyMandrelArray(3)
            End If

            '' Zarfiat Sanji
            '' tedadZayeat - zaman Tahvil ghete - zarfiat khali - zarfiat mojod - mahsol rang shode - mahsol nimsakht
            '' 1-2-3-4-5-6
            Dim zarfiatStr As String = DataGridView1.SelectedRows(0).Cells("zarfiatSanji").Value.ToString
            'Console.WriteLine(zarfiatStr)
            If Len(zarfiatStr) > 0 Then
                Dim zarfiatStrArray As String() = zarfiatStr.Split(New Char() {"-"c})
                TBProductionLoss.Text = zarfiatStrArray(0)
                TBDue.Text = zarfiatStrArray(1)
                TBEmpty.Text = zarfiatStrArray(2)
                TBAvailable.Text = zarfiatStrArray(3)
                TBPRang.Text = zarfiatStrArray(4)
                TBPNimsakht.Text = zarfiatStrArray(5)
            End If

            '' able to produce 
            Dim doable = DataGridView1.SelectedRows(0).Cells("doable").Value.ToString
            If doable = "yes" Then
                RBProducable.Checked = True
            ElseIf doable = "no" Then
                RBNotProducable.Checked = True
                TBReasonOfNotProducing.Text = DataGridView1.SelectedRows(0).Cells("whyNot").Value.ToString
            Else
                RBProducable.Checked = False
                RBNotProducable.Checked = False
            End If


            '' ---------------------------------------------------------------------------------------------------------------------------

            If wireState = "موجود است" Then
                RMaftol1.Checked = True
            ElseIf wireState = "پیل و پولیش شود" Then
                RMaftol2.Checked = True
            ElseIf wireState = "درخواست خرید" Then
                RMaftol3.Checked = True
            ElseIf wireState = "ارسال شده به پیل و پولیش" Then
                RMaftol4.Checked = True
            End If
            '' ---------------------------------------------------------------------------------------------------------------------------
            Dim mandrelState As String = DataGridView1.SelectedRows(0).Cells("موجودی مندرل").Value.ToString
            If mandrelState = "موجود است" Then
                RadioButton5.Checked = True
            Else
                RadioButton6.Checked = True
            End If
            '' ---------------------------------------------------------------------------------------------------------------------------
            '' TODO: This is 3 seprate calls to the database fix this
            If TBMR1.Text <> "" Then
                Dim sql_command = String.Format("SELECT wireWeight, wireDiameter, wireLength FROM wireInventory WHERE wireCode = '{0}'", TBMR1.Text)
                Console.WriteLine(sql_command)
                Dim dt = Await Task(Of DataTable).Run(Function() LoadDataTable(sql_command))
                Dim selectedWireWeight As String = dt.Rows(0)(0).ToString
                LSelectedWireD.Text = dt.Rows(0)(1).ToString
                LSelectedWireL.Text = dt.Rows(0)(2).ToString

                If IsNumeric(selectedWireWeight) Then
                    Lw1Weight.Text = selectedWireWeight
                    Lw1Unit.Text = "شاخه"
                    TBMRQ1.Text = (Math.Round(Val(TBMRQ1.Text) / Val(selectedWireWeight), 0)).ToString
                Else
                    Lw1Weight.Text = "-"
                    Lw1Unit.Text = "کیلوگرم"
                End If

            End If
            If TBMR2.Text <> "" Then
                Dim sql_command = String.Format("SELECT wireWeight FROM wireInventory WHERE wireCode = '{0}'", TBMR2.Text)
                Dim dt = Await Task(Of DataTable).Run(Function() LoadDataTable(sql_command))
                Dim selectedWireWeight As String = dt.Rows(0)(0).ToString

                If IsNumeric(selectedWireWeight) Then
                    Lw2Weight.Text = selectedWireWeight
                    Lw2Unit.Text = "شاخه"
                    TBMRQ2.Text = (Math.Round(Val(TBMRQ2.Text) / Val(selectedWireWeight), 0)).ToString
                Else
                    Lw2Weight.Text = "-"
                    Lw2Unit.Text = "کیلوگرم"
                End If

            End If
            If TBMR3.Text <> "" Then
                Dim sql_command = String.Format("SELECT wireWeight FROM wireInventory WHERE wireCode = '{0}'", TBMR3.Text)
                Dim dt = Await Task(Of DataTable).Run(Function() LoadDataTable(sql_command))
                Dim selectedWireWeight As String = dt.Rows(0)(0).ToString
                If IsNumeric(selectedWireWeight) Then
                    Lw3Weight.Text = selectedWireWeight
                    Lw3Unit.Text = "شاخه"
                    TBMRQ3.Text = (Math.Round(Val(TBMRQ3.Text) / Val(selectedWireWeight), 0)).ToString
                Else
                    Lw3Weight.Text = "-"
                    Lw3Unit.Text = "کیلوگرم"
                End If
            End If
            '' ---------------------------------------------------------------------------------------------------------------------------




            TabControl1.SelectedTab = TabPage2
        Catch ex As Exception
            MsgBox(ex.Message, vbCritical, "خطا")
            Logger.LogFatal("خطا در انتخاب امکان سنجی برای ویراش", ex)
        End Try

    End Sub




    Private Sub HandleWireStateChange()
        If RMaftol1.Checked Then
            LMaftolStatus.Text = "موجود است" ' Maftol Mojod Ast
            GroupBoxReserve.Enabled = True
            GroupBoxPill.Enabled = False
            GroupBoxBuy.Enabled = False
        ElseIf RMaftol2.Checked Then
            LMaftolStatus.Text = "پیل و پولیش شود" 'Maftol bayad Pill va Polish Shavad
            GroupBoxReserve.Enabled = True
            GroupBoxPill.Enabled = True
            GroupBoxBuy.Enabled = False
        ElseIf RMaftol3.Checked Then
            LMaftolStatus.Text = "درخواست خرید" 'Maftol Bayad kharidari shavad
            GroupBoxReserve.Enabled = False
            GroupBoxPill.Enabled = True
            GroupBoxBuy.Enabled = True
        ElseIf RMaftol4.Checked Then
            LMaftolStatus.Text = "ارسال شده به پیل و پولیش" 'Maftol baraye pill va polish ersal shode ast
            GroupBoxReserve.Enabled = True
            GroupBoxPill.Enabled = False
            GroupBoxBuy.Enabled = False
        End If
    End Sub
    Private Sub RMaftol1_CheckedChanged(sender As Object, e As EventArgs) Handles RMaftol1.CheckedChanged
        HandleWireStateChange()
    End Sub

    Private Sub RMaftol2_CheckedChanged(sender As Object, e As EventArgs) Handles RMaftol2.CheckedChanged
        HandleWireStateChange()
    End Sub

    Private Sub RMaftol3_CheckedChanged(sender As Object, e As EventArgs) Handles RMaftol3.CheckedChanged
        HandleWireStateChange()
    End Sub

    Private Sub RMaftol4_CheckedChanged(sender As Object, e As EventArgs) Handles RMaftol4.CheckedChanged
        HandleWireStateChange()
    End Sub

    Private Async Sub LMandrelInventory_Click(sender As Object, e As EventArgs) Handles LMandrelInventory.Click
        'TODO: Check to see if mandrel is in the inventory
        ' Dim mandrelState As Boolean
        If IsNumeric(LMandrelDiameter.Text) Then
            'Using cn As New OleDbConnection(connectionString)
            '    Using cmd As New OleDbCommand With {.Connection = cn}
            '        cmd.CommandText = "SELECT COUNT(*) FROM mandrels WHERE mandrelDiameter = '" + LMandrelDiameter.Text + "' ;"
            '        Try
            '            cn.Open()
            '            If cmd.ExecuteScalar() > 0 Then
            '                'Mandrel is in the inventory
            '                RadioButton5.Checked = True
            '            Else
            '                'Mandrel is not Present
            '                RadioButton6.Checked = True
            '            End If
            '        Catch ex As Exception
            '            MsgBox("خطا در ارتباط با دیتابیس", vbCritical + vbMsgBoxRight, "خطا")
            '            Logger.LogFatal(ex.Message, ex)
            '        Finally
            '            cn.Close()
            '        End Try
            '    End Using
            'End Using
            Using cn = GetDatabaseCon()
                Using cmd = cn.CreateCommand()
                    cmd.CommandText = "SELECT COUNT(*) FROM mandrels WHERE mandrelDiameter = '" + LMandrelDiameter.Text + "' ;"
                    Try
                        Await cn.OpenAsync()
                        If cmd.ExecuteScalar() > 0 Then
                            'Mandrel is in the inventory
                            RadioButton5.Checked = True
                        Else
                            'Mandrel is not Present
                            RadioButton6.Checked = True
                        End If
                    Catch ex As Exception
                        MsgBox("خطا در ارتباط با دیتابیس", vbCritical + vbMsgBoxRight, "خطا")
                        Logger.LogFatal(ex.Message, ex)
                    Finally
                        cn.Close()
                    End Try
                End Using
            End Using
        End If
        mandrels.Show()
        mandrels.SearchMandrelDataBase(LMandrelDiameter.Text)

    End Sub

    Private Sub Rpack1_CheckedChanged(sender As Object, e As EventArgs) Handles Rpack1.CheckedChanged
        If Rpack1.Checked Then
            LPack.Text = "هر پالت شامل"
        ElseIf Rpack2.Checked Then
            LPack.Text = "هر کارتن شامل"
        End If
    End Sub
    Private Sub Rpack2_CheckedChanged(sender As Object, e As EventArgs) Handles Rpack2.CheckedChanged
        If Rpack1.Checked Then
            LPack.Text = "هر پالت شامل"
        ElseIf Rpack2.Checked Then
            LPack.Text = "هر کارتن شامل"
        End If
    End Sub

    Private Sub Label6_Click(sender As Object, e As EventArgs) Handles Label6.Click
        ' TODO: Open Product DataBase
    End Sub

    Private Sub Label24_Click(sender As Object, e As EventArgs) Handles Label24.Click
        ' TODO: Open Customer DataBase
    End Sub
    Private Async Function LoadEmkansanjiTable() As Task
        'Using cn As New OleDbConnection(connectionString)
        '    Using cmd As New OleDbCommand With {.Connection = cn}

        '        Dim emkanSanjiColumnNames As String = " springDataBase.productName, emkansanji.quantity, emkansanji.letterNo, customers.customerName "

        '        'the paranthesis in the query are mandatory
        '        cmd.CommandText = "SELECT " & ESColumnNames & " FROM (emkansanji INNER JOIN springDataBase ON emkansanji.productID = springDataBase.ID) INNER JOIN customers ON emkansanji.customerID = customers.ID" &
        '            " ;"

        '        Dim dt As New DataTable With {.TableName = "emkansanji"}
        '        'Try
        '        cn.Open()
        '        Dim ds As New DataSet
        '        Dim emkansanji As New DataTable With {.TableName = "emkansanji"}
        '        ds.Tables.Add(emkansanji)
        '        ds.Load(cmd.ExecuteReader(), LoadOption.OverwriteChanges, emkansanji)
        '        DataGridView1.DataSource = ds.Tables("emkansanji")
        '        'DataGridView1.Columns(0).Visible = False
        '        cn.Close()
        '        DataGridView1.Columns(0).Visible = False
        '        DataGridView1.Columns(1).Visible = False
        '        DataGridView1.Columns(2).Visible = False
        '        DataGridView1.Columns(3).Visible = False
        '        DataGridView1.Columns(4).Visible = False
        '        DataGridView1.Columns(5).Visible = False
        '        DataGridView1.Columns(6).Visible = False

        '    End Using
        'End Using

        'Using cn As New OleDbConnection(connectionString)
        '    Using cmd As New OleDbCommand With {.Connection = cn}

        '        Dim emkanSanjiColumnNames As String = " springDataBase.productName, emkansanji.quantity, emkansanji.letterNo, customers.customerName "

        '        'the paranthesis in the query are mandatory
        '        cmd.CommandText = "SELECT " & ESColumnNames & " FROM ((emkansanji LEFT JOIN springDataBase ON emkansanji.productID = springDataBase.ID) LEFT JOIN customers ON emkansanji.customerID = customers.ID) WHERE " &
        '            " springDataBase.productName LIKE '%" & TBEnergySazProductName.Text & "%' AND" &
        '         " customers.customerName LIKE '%" & TBCustomerName.Text & "%' AND" &
        '         " emkansanji.customerProductName LIKE '%" & TBCustomerProductName.Text & "%' AND" &
        '         " emkansanji.orderState LIKE '%" & CBOrderState.Text & "%' AND" &
        '         " emkansanji.ID LIKE  '%" & TBEmkansanjiID.Text & "' AND " &
        '        " emkansanji.productCode LIKE '%" & TBCustomerProductCode.Text & "%' AND " &
        '        " emkansanji.orderNo LIKE '%" & TBOrderNo.Text & "%' AND " &
        '        " emkansanji.letterNo LIKE '%" & TBLetterNo.Text & "%' " &
        '        " ;" 'TODO : Search the database based on Reserved wire and coil 

        '        'cmd.CommandText = "SELECT " & ESColumnNames & " FROM ((emkansanji INNER JOIN springDataBase ON emkansanji.productID = springDataBase.ID) INNER JOIN customers ON emkansanji.customerID = customers.ID) ;"
        '        Console.WriteLine(cmd.CommandText)

        '        Dim dt As New DataTable With {.TableName = "emkansanji"}
        '        'Try
        '        cn.Open()
        '        Dim ds As New DataSet
        '        Dim emkansanji As New DataTable With {.TableName = "emkansanji"}
        '        ds.Tables.Add(emkansanji)
        '        ds.Load(cmd.ExecuteReader(), LoadOption.OverwriteChanges, emkansanji)
        '        '' This enables the use of bs.filter method to filter the database
        '        'emkansanji_bs.DataSource = ds.Tables("emkansanji")
        '        emkansanji_bs.DataSource = ds.Tables("emkansanji")
        '        DataGridView1.DataSource = emkansanji_bs

        '        cn.Close()
        '        DataGridView1.Columns(0).Visible = False
        '        DataGridView1.Columns(1).Visible = False
        '        DataGridView1.Columns(2).Visible = False
        '        DataGridView1.Columns(3).Visible = False
        '        DataGridView1.Columns(4).Visible = False
        '        DataGridView1.Columns(5).Visible = False
        '        DataGridView1.Columns(6).Visible = False

        '    End Using
        'End Using


        Dim sql_command = "SELECT " & ESColumnNames & " FROM ((emkansanji LEFT JOIN springDataBase ON emkansanji.productID = springDataBase.ID) LEFT JOIN customers ON emkansanji.customerID = customers.ID) WHERE " &
                                 " springDataBase.productName LIKE '%" & TBEnergySazProductName.Text & "%' AND" &
                                 " customers.customerName LIKE '%" & TBCustomerName.Text & "%' AND" &
                                 " emkansanji.customerProductName LIKE '%" & TBCustomerProductName.Text & "%' AND" &
                                 " emkansanji.orderState LIKE '%" & CBOrderState.Text & "%' AND" &
                                 " emkansanji.ID LIKE  '%" & TBEmkansanjiID.Text & "%' AND " &
                                 " emkansanji.productCode LIKE '%" & TBCustomerProductCode.Text & "%' AND " &
                                 " emkansanji.orderNo LIKE '%" & TBOrderNo.Text & "%' AND " &
                                 " emkansanji.letterNo LIKE '%" & TBLetterNo.Text & "%' " &
                                 " ORDER BY emkansanji.ID ;" 'TODO : Search the database based on Reserved wire and coil 
        Try
            Dim dt = Await Task(Of DataTable).Run(Function() LoadDataTable(sql_command))
            emkansanji_bs.DataSource = dt
            DataGridView1.DataSource = emkansanji_bs
            'bs2.Filter = ""
            ' Hide values which are not for the user to see
            DataGridView1.Columns("productID").Visible = False
            DataGridView1.Columns("customerID").Visible = False
            DataGridView1.Columns("wireDiameter").Visible = False
            DataGridView1.Columns("OD").Visible = False
            DataGridView1.Columns("L0").Visible = False
            DataGridView1.Columns("wireLength").Visible = False
            DataGridView1.Columns("mandrelDiameter").Visible = False
            DataGridView1.Columns("pProcess").Visible = False
            DataGridView1.Columns("productReserve").Visible = False
            DataGridView1.Columns("productionProcess").Visible = False
            DataGridView1.Columns("springInEachPackage").Visible = False
            DataGridView1.Columns("packagingCost").Visible = False
            DataGridView1.Columns("doable").Visible = False
            DataGridView1.Columns("whyNot").Visible = False
            DataGridView1.Columns("productionReserve").Visible = False
        Catch ex As Exception
            MsgBox("خطا در برقرای ارتباط با دیتابیس", vbCritical + RightToLeft + vbMsgBoxRight, "خطا")
            Logger.LogFatal(sql_command, ex)
        End Try




    End Function

    Private Sub Lwire1_Click(sender As Object, e As EventArgs) Handles Lwire1.Click
        Dim wireSelectionForm = New wires()
        wiresFormState = "selection"
        wireFormCaller = "wire1"
        wireSelectionForm.Show()
    End Sub

    Private Sub Lwire2_Click(sender As Object, e As EventArgs) Handles Lwire2.Click
        Dim wireSelectionForm = New wires()
        wiresFormState = "selection"
        wireFormCaller = "wire2"
        wireSelectionForm.Show()
    End Sub

    Private Sub Lwire3_Click(sender As Object, e As EventArgs) Handles Lwire3.Click
        Dim wireSelectionForm = New wires()
        wiresFormState = "selection"
        wireFormCaller = "wire3"
        wireSelectionForm.Show()
    End Sub

    Private Async Sub BTEmkansanjiSearch_Click(sender As Object, e As EventArgs) Handles BTEmkansanjiSearch.Click
        ''LoadEmkansanjiTable()
        'Using cn As New OleDbConnection(connectionString)
        '    Using cmd As New OleDbCommand With {.Connection = cn}

        '        Dim emkanSanjiColumnNames As String = " springDataBase.productName, emkansanji.quantity, emkansanji.letterNo, customers.customerName "

        '        'the paranthesis in the query are mandatory
        '        cmd.CommandText = "SELECT " & ESColumnNames & " FROM ((emkansanji INNER JOIN springDataBase ON emkansanji.productID = springDataBase.ID) INNER JOIN customers ON emkansanji.customerID = customers.ID) WHERE " &
        '            " springDataBase.productName LIKE '%" & TBEnergySazProductName.Text & "%' AND" &
        '         " customers.customerName LIKE '%" & TBCustomerName.Text & "%' AND" &
        '         " emkansanji.customerProductName LIKE '%" & TBCustomerProductName.Text & "%' AND" &
        '         " emkansanji.orderState LIKE '%" & CBOrderState.Text & "%' AND" &
        '         " emkansanji.ID LIKE  '%" & TBEmkansanjiID.Text & "%' AND " &
        '        " emkansanji.productCode LIKE '%" & TBCustomerProductCode.Text & "%' AND " &
        '        " emkansanji.orderNo LIKE '%" & TBOrderNo.Text & "%' AND " &
        '        " emkansanji.letterNo LIKE '%" & TBLetterNo.Text & "%' " &
        '        " ;" 'TODO : Search the database based on Reserved wire and coil 
        '        DataGridView1.Columns(0).Visible = False
        '        DataGridView1.Columns(1).Visible = False
        '        DataGridView1.Columns(2).Visible = False
        '        DataGridView1.Columns(3).Visible = False
        '        DataGridView1.Columns(4).Visible = False
        '        DataGridView1.Columns(5).Visible = False
        '        DataGridView1.Columns(6).Visible = False

        '        Dim dt As New DataTable With {.TableName = "emkansanji"}
        '        'Try
        '        cn.Open()
        '        Dim ds As New DataSet
        '        Dim emkansanji As New DataTable With {.TableName = "emkansanji"}
        '        ds.Tables.Add(emkansanji)
        '        ds.Load(cmd.ExecuteReader(), LoadOption.OverwriteChanges, emkansanji)
        '        DataGridView1.DataSource = ds.Tables("emkansanji")
        '        cn.Close()

        '    End Using
        'End Using



        'the paranthesis in the query are mandatory
        Dim sql_command = "SELECT " & ESColumnNames & " FROM ((emkansanji INNER JOIN springDataBase ON emkansanji.productID = springDataBase.ID) INNER JOIN customers ON emkansanji.customerID = customers.ID) WHERE " &
                    " springDataBase.productName LIKE '%" & TBEnergySazProductName.Text & "%' AND" &
                 " customers.customerName LIKE '%" & TBCustomerName.Text & "%' AND" &
                 " emkansanji.customerProductName LIKE '%" & TBCustomerProductName.Text & "%' AND" &
                 " emkansanji.orderState LIKE '%" & CBOrderState.Text & "%' AND" &
                 " emkansanji.ID LIKE  '%" & TBEmkansanjiID.Text & "%' AND " &
                " emkansanji.productCode LIKE '%" & TBCustomerProductCode.Text & "%' AND " &
                " emkansanji.orderNo LIKE '%" & TBOrderNo.Text & "%' AND " &
                " emkansanji.letterNo LIKE '%" & TBLetterNo.Text & "%' " &
                " ;" 'TODO : Search the database based on Reserved wire and coil 
        Try
            DataGridView1.Columns("productID").Visible = False
            DataGridView1.Columns("customerID").Visible = False
            DataGridView1.Columns("wireDiameter").Visible = False
            DataGridView1.Columns("OD").Visible = False
            DataGridView1.Columns("L0").Visible = False
            DataGridView1.Columns("wireLength").Visible = False
            DataGridView1.Columns("mandrelDiameter").Visible = False
            DataGridView1.Columns("pProcess").Visible = False
            DataGridView1.Columns("productReserve").Visible = False
            DataGridView1.Columns("productionProcess").Visible = False
            DataGridView1.Columns("springInEachPackage").Visible = False
            DataGridView1.Columns("packagingCost").Visible = False
            DataGridView1.Columns("doable").Visible = False
            DataGridView1.Columns("whyNot").Visible = False
            DataGridView1.Columns("productionReserve").Visible = False
            'Try
            Dim dt = Await Task(Of DataTable).Run(Function() LoadDataTable(sql_command))
            emkansanji_bs.DataSource = dt
            DataGridView1.DataSource = emkansanji_bs.DataSource
        Catch ex As Exception
            MsgBox("خطا در برقرای ارتباط با دیتابیس", vbCritical + RightToLeft + vbMsgBoxRight, "خطا")
            Logger.LogFatal(sql_command, ex)
        End Try

    End Sub

    Private Async Sub emkanSanjiForm_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        '' If this form is called as a pop-up form from wires form then update wire and order data when closing
        If thisFormsOwner = "wiresForm" Then
            Await wires.LoadWiresData()
            Await wires.LoadOrdersData()
            Me.Dispose()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim sql = " UPDATE emkansanji SET productID = 553 WHERE ISNULL(productID)"
        Dim sql2 = "SELECT * FROM emkansanji WHERE ISNULL(productID) "
        Dim dt = LoadDataTable(sql)
        DataGridView1.DataSource = dt
        MsgBox("done")
    End Sub

    Private Sub Label47_Click(sender As Object, e As EventArgs) Handles Label47.Click
        '' TODO
    End Sub

    Private Sub HandleMandrelStateChange()
        If RadioButton5.Checked = True Then 'Mandrel is available
            GroupBoxBuyMandrel.Enabled = False
        ElseIf RadioButton6.Checked = True Then
            GroupBoxBuyMandrel.Enabled = True
        End If
    End Sub
    Private Sub RadioButton5_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton5.CheckedChanged
        HandleMandrelStateChange()
    End Sub

    Private Sub RadioButton6_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton6.CheckedChanged
        HandleMandrelStateChange()
    End Sub

    Private Sub LPack_Click(sender As Object, e As EventArgs) Handles LPack.Click
        Try
            Dim maxInLength = Int(TBPackageL.Text / LOutsideDiameter.Text)
            Dim maxInWidth = Int(TBPackageW.Text / LOutsideDiameter.Text)
            Dim NoOfRows = Int(TBPackageH.Text / LFreeLength.Text)

            TBPackageCount.Text = maxInLength * maxInWidth * NoOfRows
            TBPCostForEach.Text = Math.Round((Val(TBPCost.Text) / (Int(LOrderQuantity.Text / TBPackageCount.Text) + 1)) / 10000, 0) * 10000
        Catch ex As Exception
            MsgBox(" پارامتر هار ورودی را کنترل کنید" + ex.Message, vbCritical + vbMsgBoxRight + RightToLeft, "خطا")
        End Try
    End Sub

    Private Sub RBProducable_CheckedChanged(sender As Object, e As EventArgs) Handles RBProducable.CheckedChanged
        If RBProducable.Checked = True Then
            TBReasonOfNotProducing.Enabled = False
            LReasonOfNotProducing.Enabled = False
        Else
            TBReasonOfNotProducing.Enabled = True
            LReasonOfNotProducing.Enabled = True
        End If
    End Sub

    Private Sub BTModifyES2_Click(sender As Object, e As EventArgs) Handles BTModifyES2.Click
        BTModifyES.PerformClick()
    End Sub

    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CHShot.CheckedChanged

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        '' Add checkboxes to collection (array)

        'MsgBox(CBA.Count.ToString)
        'Dim processCode = GenerateProductionProcess(CBA)
        'MsgBox(processCode)
        Dim code = InputBox("EnterSomething", "title", "1111111111")
        ParseProductionProcess(CBA, code)

    End Sub

    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs) Handles GroupBox2.Enter

    End Sub


    '' ----------------------------------------------------------------------------------------------------------------------
    '' ----------------------------------------------------------------------------------------------------------------------
    '' ---------------------- Declration of variables that are used to update the database and excel file -------------------
    '' ----------------------------------------------------------------------------------------------------------------------
    '' ----------------------------------------------------------------------------------------------------------------------

    Dim mandrelState As String
    Dim r1_q, r2_q, r3_q As String
    Dim productionProcess As String
    Dim inspectionProcess As String
    Dim orderType As String
    Dim buyWireData As String
    Dim buyMandrelData As String
    Dim zarfiatData As String
    Dim producable As String
    Dim whyNotProducable As String
    Dim packageType As String
    Private Sub ModifyEmkansanjiExcelFile()
        ' Generate the file path for excel file -- it's really not needed now but just for the future
        Dim pc As New Globalization.PersianCalendar
        Dim excelFiledir As String = pc.GetYear(Now).ToString & "\" & getMonthName(pc.GetMonth(Now)) & "\" & TBMCustomerName.Text & "\"
        Dim excelFileName As String = TBMEnergySazProductName.Text

        excelFileName = stripFileName(excelFileName)

        Dim path As String = excelFilesBasePath & excelFiledir
        Dim saveFilePath As String = path & excelFileName 'Complete path of the excel file

        'Check to see if file with this name exist to rename it to prevent overwriting
        saveFilePath = preverntOverwriting(saveFilePath, ".xlsx")



        '------------------------------------------------------
        LStatus.Visible = True
        LStatus.Text = "در حال آماده سازی فایل اکسل امکان سنجی ..."
        Me.Cursor = Cursors.WaitCursor

        Try
            '' We use the old address in the db to just modify the file
            Dim excelFilePath = DataGridView1.SelectedRows(0).Cells("excelFilePath").Value.ToString()
            If System.IO.File.Exists(excelFilePath) = False Then
                MsgBox("فایل اکسل این امکان سنجی در محل تعیین شده وجود ندارد", vbCritical + RightToLeft + vbMsgBoxRtlReading, "")
                Logger.LogInfo("Tried to modify excel file but path doesn't exist -> " & excelFilePath)
                '' TODO:Create it again
                Exit Try
            End If

            Dim excel As Excel.Application = New Excel.Application
            Dim w As Excel.Workbook = excel.Workbooks.Open(excelFilePath)

            '' ----------------------------- Populate fields in the emkansanji excel Template -------------------------------
            excel.Range("customerName").Value = NormalizeString(TBMCustomerName.Text)
            excel.Range("letterNo").Value = NormalizeString(TBMLetterNo.Text)
            excel.Range("pName").Value = NormalizeString(TBMCustomerProductName.Text)
            excel.Range("letterDate").Value = NormalizeString(TBMLetterDate.Text)
            excel.Range("dwgNo").Value = NormalizeString(TBMCustomerDwgNo.Text)
            excel.Range("quantity").Value = NormalizeString(TBQuantity.Text)
            excel.Range("sampleQuantity").Value = NormalizeString(TBSampleQuantity.Text)
            excel.Range("pDate").Value = NormalizeString(TBMProccessingDate.Text)
            excel.Range("standard").Value = NormalizeString(CBStandard.Text)
            excel.Range("grade").Value = NormalizeString(TBGrade.Text)
            excel.Range("customerProductCode").Value = NormalizeString(TBMCustomerProductCode.Text)
            excel.Range("comment").Value = NormalizeString(TBComment.Text)

            '' -------------------------------------------------------------------------------------------------------------
            excel.Range("ESpName").Value = NormalizeString(TBMEnergySazProductName.Text)
            'excel.Range("ESProductCode").Value = NormalizeString(DataGridView1.SelectedRows(0).Cells("کد کالا").Value.ToString)
            '' -------------------------------------------------------------------------------------------------------------
            'excel.Range("springType").Value = NormalizeString(DataGridView1.SelectedRows(0).Cells("نوع فنر").Value.ToString)
            'excel.Range("material").Value = NormalizeString(DataGridView1.SelectedRows(0).Cells("جنس مواد").Value.ToString)
            'excel.Range("wireD").Value = NormalizeString(DataGridView1.SelectedRows(0).Cells("قطر مفتول").Value.ToString)
            'excel.Range("OD").Value = NormalizeString(DataGridView1.SelectedRows(0).Cells("قطر خارجی").Value.ToString)
            'excel.Range("mandrel").Value = NormalizeString(DataGridView1.SelectedRows(0).Cells("قطر شفت").Value.ToString)
            'excel.Range("Nt").Value = NormalizeString(DataGridView1.SelectedRows(0).Cells("حلقه کل").Value.ToString)
            'excel.Range("Na").Value = NormalizeString(DataGridView1.SelectedRows(0).Cells("حلقه فعال").Value.ToString)
            'excel.Range("L0").Value = NormalizeString(DataGridView1.SelectedRows(0).Cells("طول آزاد").Value.ToString)
            'excel.Range("coilingDirection").Value = NormalizeString(DataGridView1.SelectedRows(0).Cells("جهت پیچش").Value.ToString)
            'excel.Range("springRate").Value = NormalizeString(DataGridView1.SelectedRows(0).Cells("ریت فنر").Value.ToString)
            'excel.Range("firstCoil").Value = NormalizeString(DataGridView1.SelectedRows(0).Cells("شکل حلقه ابتدا").Value.ToString)
            'excel.Range("lastCoil").Value = NormalizeString(DataGridView1.SelectedRows(0).Cells("شکل حلقه انتها").Value.ToString)
            'excel.Range("Force1").Value = NormalizeString(DataGridView1.SelectedRows(0).Cells("F1").Value.ToString)
            'excel.Range("Length1").Value = NormalizeString(DataGridView1.SelectedRows(0).Cells("L1").Value.ToString)
            'excel.Range("Force2").Value = NormalizeString(DataGridView1.SelectedRows(0).Cells("F2").Value.ToString)
            'excel.Range("Length2").Value = NormalizeString(DataGridView1.SelectedRows(0).Cells("L2").Value.ToString)
            'excel.Range("Force3").Value = NormalizeString(DataGridView1.SelectedRows(0).Cells("F3").Value.ToString)
            'excel.Range("Length3").Value = NormalizeString(DataGridView1.SelectedRows(0).Cells("L3").Value.ToString)
            'excel.Range("forceUnit").Value = NormalizeString(DataGridView1.SelectedRows(0).Cells("واحد نیرو").Value.ToString)
            '' -----------------------------------------------------------------------------------------------------------
            'excel.Range("wireLength").Value = NormalizeString(DataGridView1.SelectedRows(0).Cells("طول مفتول").Value.ToString)

            '' ------------------------------------------------------------------------------------------------------------
            excel.Range("productionProcess").Value = productionProcess
            excel.Range("inspectionProcess").Value = inspectionProcess
            excel.Range("orderTypeCode").Value = orderType



            '' TODO: Add the rest of them

            '' Wire state : 
            If RMaftol1.Checked Then
                excel.Range("wireAvailable").Value = "TRUE"
                excel.Range("pillWire").Value = "FALSE"
                excel.Range("buyWire").Value = "FALSE"
                excel.Range("inventoryWireD").Value = LSelectedWireD.Text
                excel.Range("inventoryWireLength").Value = LSelectedWireL.Text
            ElseIf RMaftol2.Checked Then
                excel.Range("pillWire").Value = "TRUE"
                excel.Range("buyWire").Value = "FALSE"
                excel.Range("wireAvailable").Value = "FALSE"
                excel.Range("inventoryWireD").Value = LSelectedWireD.Text
                excel.Range("inventoryWireLength").Value = LSelectedWireL.Text
                excel.Range("pillingCost").Value = NormalizeString(TBPillCost.Text)
            ElseIf RMaftol3.Checked Then
                excel.Range("buyWire").Value = "TRUE"
                excel.Range("pillWire").Value = "FALSE"
                excel.Range("wireAvailable").Value = "FALSE"
                excel.Range("inventoryWireD").Value = NormalizeString(TBBuyWireD.Text)
                excel.Range("inventoryWireLength").Value = NormalizeString(TBBuyWireLength.Text)
                excel.Range("pillingCost").Value = NormalizeString(TBPillCost.Text)
            End If



            '' Mandrel State: 
            If RadioButton6.Checked Then '' Mandrel is not in the inventory
                excel.Range("buyMandrel").Value = "TRUE"
                excel.Range("buyMandrelLength").Value = NormalizeString(TBBuyMandrelL.Text)
                excel.Range("buyMandrelPrice").Value = NormalizeString(TBBuyMandrelPrice.Text)
                excel.Range("buyMandrelCost").Value = NormalizeString(TBBuyMandrelCost.Text)
                excel.Range("buyMandrelDiameter").Value = NormalizeString(TBBuyMandrelD.Text)
            ElseIf RadioButton5.Checked Then
                excel.Range("buyMandrel").Value = "FALSE"
            End If

            '' Packaging
            excel.Range("springIneachPackage").Value = NormalizeString(TBPackageCount.Text)
            excel.Range("packageCostForEach").Value = NormalizeString(TBPCostForEach.Text)
            excel.Range("packageType").Value = NormalizeString(LPack.Text)

            '' Due Date
            excel.Range("dueDate").Value = NormalizeString(TBDue.Text)

            '' Production Capacity
            'excel.Range("sampleQuantity").Value = NormalizeString(LPack.Text) this is done in upper code
            excel.Range("empty").Value = NormalizeString(TBEmpty.Text)
            excel.Range("capacity").Value = NormalizeString(TBAvailable.Text)
            excel.Range("rangShode").Value = NormalizeString(TBPRang.Text)
            excel.Range("nimsakht").Value = NormalizeString(TBPNimsakht.Text)

            '' comment 
            excel.Range("productionLoss").Value = NormalizeString(TBProductionLoss.Text)
            excel.Range("comment").Value = NormalizeString(TBComment.Text)

            '' Final section 
            If RBProducable.Checked Then
                excel.Range("producable").Value = "TRUE"
                excel.Range("notProducable").Value = "FALSE"
                excel.Range("whyNot").Value = ""
            ElseIf RBNotProducable.Checked Then
                excel.Range("producable").Value = "FALSE"
                excel.Range("notProducable").Value = "TRUE"
                excel.Range("whyNot").Value = TBReasonOfNotProducing.Text
            End If
            '' ------------------------------------------------------------------------------------------------------------

            'Check if there is an address provided for the second excel file, if not it uses working directory
            Dim saveDuplicatePath As String
            If ConfigurationManager.ConnectionStrings("saveAsFilePath").ConnectionString = "" Then
                saveDuplicatePath = My.Application.Info.DirectoryPath
            Else
                saveDuplicatePath = ConfigurationManager.ConnectionStrings("saveAsFilePath").ConnectionString
            End If
            'MkDir(path)  'Make the directory if it doesn't exist
            'w.SaveAs(excelFilePath)
            w.Save()
            w.SaveAs(saveDuplicatePath & "\" & excelFileName) 'Save another file in the application directory
            w.Close()
            Logger.LogInfo("Excel File Modified with path (" + saveFilePath + ")")
        Catch ex As Exception
            MsgBox("خطا در تکمیل قالب اکسل امکان سنجی. فایل اکسل را چک کرده و مجددا امتحان کنید", vbCritical + vbMsgBoxRight, "خطا")
            Logger.LogFatal(ex.Message, ex)

        End Try
        LStatus.Visible = False
        'Me.Cursor = Cursors.Default
    End Sub
    Private Async Sub BTModifyES_Click(sender As Object, e As EventArgs) Handles BTModifyES.Click
        '' Establish mandrel state
        If RadioButton5.Checked = True Then
            mandrelState = "موجود است"
        Else
            mandrelState = "ساخته شود"
        End If

        '' Establish wire reservation
        '' TODO: This produce heavy truncation errors

        If Lw1Unit.Text = "شاخه" Then
            r1_q = Math.Round(Val(TBMRQ1.Text) * Val(Lw1Weight.Text), 2).ToString
        Else
            r1_q = TBMRQ1.Text
        End If
        If Lw2Unit.Text = "شاخه" Then
            r2_q = Math.Round(Val(TBMRQ2.Text) * Val(Lw2Weight.Text), 2).ToString
        Else
            r2_q = TBMRQ2.Text
        End If
        If Lw2Unit.Text = "شاخه" Then
            r3_q = Math.Round(Val(TBMRQ3.Text) * Val(Lw3Weight.Text), 2).ToString
        Else
            r3_q = TBMRQ3.Text
        End If

        '' Production process 
        productionProcess = GenerateProductionProcess(CBA)

        '' Inspection process 
        inspectionProcess = GenerateInspectionProcess(CBA)

        '' Order Type
        '' generate order type code   - a two digit number - first digit New product 1 - changing old product 2 - old product 3
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

        '' buy wire data 
        buyWireData = TBBuyWireD.Text & "-" & TBBuyWireLength.Text & "-" & TBPillCost.Text

        '' buy Mandrel Data 
        buyMandrelData = TBBuyMandrelD.Text & "-" & TBBuyMandrelL.Text & "-" & TBBuyMandrelPrice.Text & "-" & TBBuyMandrelCost.Text

        ''production loss and zarfiatsanji
        zarfiatData = TBProductionLoss.Text & "-" & TBDue.Text & "-" & TBEmpty.Text & "-" & TBAvailable.Text & "-" & TBPRang.Text & "-" & TBPNimsakht.Text

        '' production doable? 

        If RBProducable.Checked = True Then
            producable = "yes"
            whyNotProducable = ""
        ElseIf RBNotProducable.Checked = True Then
            producable = "no"
            whyNotProducable = TBReasonOfNotProducing.Text
        Else
            producable = ""
            whyNotProducable = ""
        End If

        '' Packaging 
        If Rpack2.Checked = True Then
            packageType = "کارتن"
        Else
            packageType = "پالت"
        End If

        '' Construct the sql command to update the emkansanji table with new data
        Dim sql_command As String = "UPDATE emkansanji SET" &
                            " productID = '" & TBProductIDES.Text & "'," &
                            " pProcess = '" & productionProcess & "'," &
                             " customerID = '" & TBCustomerID.Text & "'," &
                              " customerProductName = '" & TBMCustomerProductName.Text & "'," &
                              " customerDwgNo = '" & TBMCustomerDwgNo.Text & "'," &
                              " quantity = '" & TBQuantity.Text & "'," &
                              " sampleQuantity = '" & TBSampleQuantity.Text & "'," &
                              " letterNo = '" & TBMLetterNo.Text & "'," &
                              " letterDate = '" & TBMLetterDate.Text & "'," &
                              " orderNo = '" & TBMOrderNo.Text & "'," &
                              " standard = '" & CBStandard.Text & "'," &
                              " dateOfProccessing = '" & TBMProccessingDate.Text & "'," &
                              " grade = '" & TBGrade.Text & "'," &
                              " productCode = '" & TBMCustomerProductCode.Text & "'," &
                              " comment = '" & TBComment.Text & "'," &
                              " orderState = '" & CBMOrderState.Text & "'," &
                              " r1_code = '" & TBMR1.Text & "'," &
                              " r1_q = '" & r1_q & "'," &
                              " r2_code = '" & TBMR2.Text & "'," &
                              " r2_q = '" & r2_q & "'," &
                              " r3_code = '" & TBMR3.Text & "'," &
                             " r3_q = '" & r3_q & "'," &
                              " wireState = '" & LMaftolStatus.Text & "'," &
                             " verificationNo = '" & TBMVerificationNo.Text & "'," &
                             " mandrelState = '" & mandrelState & "'," &
                             " buyMandrel = '" & buyMandrelData & "'," &
                             " springInEachPackage = '" & TBPackageCount.Text & "'," &
                             " packagingCost = '" & TBPCostForEach.Text & "'," &
                             " packageType = '" & packageType & "'," &
                             " doable = '" & producable & "'," &
                              " whyNot = '" & whyNotProducable & "'," &
                              " buyWire = '" & buyWireData & "'," &
                              " zarfiatSanji = '" & zarfiatData & "'," &
                              " productReserve = '" & TBProductReserve.Text & "'," &
                              " inspectionProcess = '" & inspectionProcess & "'," &
                              " orderType = '" & orderType & "'," &
                             " verificationDate = '" & TBMVerificationDate.Text & "'" &
                             " WHERE ID = " & LemkansanjiID.Text & ";"


        Try
            If CheckChangeExcel.Checked Then
                ModifyEmkansanjiExcelFile()
            End If
            Dim dt = Await Task(Of DataTable).Run(Function() LoadDataTable(sql_command))
            Logger.LogInfo(String.Format("Modified order {0}", LemkansanjiID.Text))
            MsgBox("ویرایش مشخصات سفارش با موفقیت انجام شد", vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, "ویرایش امکان سنجی")
        Catch ex As Exception
            Logger.LogFatal(String.Format("Error Updating order info:  {0}", sql_command), ex)
            MsgBox("خطا در ویرایش مشخصات سفارش", vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, "ویرایش امکان سنجی")
        End Try
        '' Load the table in the datagrid view so the new changes are visiable in it then update wire reserve data
        Await LoadEmkansanjiTable()
        Await UpdateReservesTable()
        Me.Cursor = Cursors.Default
        'MsgBox("ویرایش مشخصات سفارش با موفقیت انجام شد", vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, "ویرایش امکان سنجی")
    End Sub


End Class