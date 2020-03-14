Imports System.Data.OleDb


Imports Excel = Microsoft.Office.Interop.Excel
Public Class emkanSanjiForm

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

    Private Sub emkanSanjiForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Loading Springs table into datagridview1
        LoadEmkansanjiTable()
    End Sub

    Private Sub Label8_Click(sender As Object, e As EventArgs) Handles Label8.Click

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Using cn As New OleDbConnection(connectionString)
            Using cmd As New OleDbCommand With {.Connection = cn}

                Dim emkanSanjiColumnNames As String = " springDataBase.productName, emkansanji.quantity, emkansanji.letterNo, customers.customerName "

                'the paranthesis in the query are mandatory
                cmd.CommandText = "SELECT " & ESColumnNames & " FROM ((emkansanji INNER JOIN springDataBase ON emkansanji.productID = springDataBase.ID) INNER JOIN customers ON emkansanji.customerID = customers.ID) WHERE " &
                    " springDataBase.productName LIKE '%" & TBEnergySazProductName.Text & "%' AND" &
                 " customers.customerName LIKE '%" & TBCustomerName.Text & "%' AND" &
                 " emkansanji.customerProductName LIKE '%" & TBCustomerProductName.Text & "%' AND" &
                 " emkansanji.orderState LIKE '%" & CBOrderState.Text & "%' AND" &
                 " emkansanji.ID LIKE  '%" & TBEmkansanjiID.Text & "%' AND " &
                " emkansanji.productCode LIKE '%" & TBCustomerProductCode.Text & "%' AND " &
                " emkansanji.orderNo LIKE '%" & TBOrderNo.Text & "%' AND " &
                " emkansanji.letterNo LIKE '%" & TBLetterNo.Text & "%' " &
                " ;" 'TODO : Search the database based on Reserved wire and coil 
                DataGridView1.Columns(0).Visible = False
                DataGridView1.Columns(1).Visible = False
                DataGridView1.Columns(2).Visible = False
                DataGridView1.Columns(3).Visible = False
                DataGridView1.Columns(4).Visible = False
                DataGridView1.Columns(5).Visible = False
                DataGridView1.Columns(6).Visible = False

                Dim dt As New DataTable With {.TableName = "emkansanji"}
                'Try
                cn.Open()
                Dim ds As New DataSet
                Dim emkansanji As New DataTable With {.TableName = "emkansanji"}
                ds.Tables.Add(emkansanji)
                ds.Load(cmd.ExecuteReader(), LoadOption.OverwriteChanges, emkansanji)
                DataGridView1.DataSource = ds.Tables("emkansanji")
                cn.Close()

            End Using
        End Using
    End Sub

    Private Async Sub BTModify_Click(sender As Object, e As EventArgs) Handles BTModify.Click
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
        LOutsideDiameter.Text = DataGridView1.SelectedRows(0).Cells("OD").Value.ToString
        LFreeLength.Text = DataGridView1.SelectedRows(0).Cells("L0").Value.ToString
        LWireDiameter.Text = DataGridView1.SelectedRows(0).Cells("wireDiameter").Value.ToString
        LWireLength.Text = DataGridView1.SelectedRows(0).Cells("wireLength").Value.ToString
        LMandrelDiameter.Text = DataGridView1.SelectedRows(0).Cells("mandrelDiameter").Value.ToString

        Dim wireState As String = DataGridView1.SelectedRows(0).Cells("وضعیت موجودی مفتول").Value.ToString
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
        If TBMR1.Text <> "" Then
            Dim sql_command = String.Format("SELECT wireWeight FROM wireInventory WHERE wireCode = '{0}'", TBMR1.Text)
            Console.WriteLine(sql_command)
            Dim dt = Await Task(Of DataTable).Run(Function() LoadDataTable(sql_command))
            Dim selectedWireWeight As String = dt.Rows(0)(0).ToString
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

    End Sub

    Private Async Sub BTModifyES_Click(sender As Object, e As EventArgs) Handles BTModifyES.Click
        Dim mandrelState As String
        If RadioButton5.Checked = True Then
            mandrelState = "موجود است"
        Else
            mandrelState = "ساخته شود"
        End If
        Dim r1_q, r2_q, r3_q As String
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
        Dim sql_command As String = "UPDATE emkansanji SET" &
                            " productID = '" & TBProductIDES.Text & "'," &
                             " customerID = '" & TBCustomerID.Text & "'," &
                              " customerProductName = '" & TBMCustomerProductName.Text & "'," &
                              " customerDwgNo = '" & TBMCustomerDwgNo.Text & "'," &
                              " quantity = '" & TBQuantity.Text & "'," &
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
                             " verificationDate = '" & TBMVerificationDate.Text & "'" &
                             " WHERE ID = " & LemkansanjiID.Text & ";"
        Dim dt = Await Task(Of DataTable).Run(Function() LoadDataTable(sql_command))
        MsgBox("done")
        LoadEmkansanjiTable()
    End Sub

    Private Sub RMaftol1_CheckedChanged(sender As Object, e As EventArgs) Handles RMaftol1.CheckedChanged
        If RMaftol1.Checked Then
            LMaftolStatus.Text = "موجود است" ' Maftol Mojod Ast
        ElseIf RMaftol2.Checked Then
            LMaftolStatus.Text = "پیل و پولیش شود" 'Maftol bayad Pill va Polish Shavad
        ElseIf RMaftol3.Checked Then
            LMaftolStatus.Text = "درخواست خرید" 'Maftol Bayad kharidari shavad
        ElseIf RMaftol4.Checked Then
            LMaftolStatus.Text = "ارسال شده به پیل و پولیش" 'Maftol baraye pill va polish ersal shode ast
        End If
    End Sub

    Private Sub RMaftol2_CheckedChanged(sender As Object, e As EventArgs) Handles RMaftol2.CheckedChanged
        If RMaftol1.Checked Then
            LMaftolStatus.Text = "موجود است" ' Maftol Mojod Ast
        ElseIf RMaftol2.Checked Then
            LMaftolStatus.Text = "پیل و پولیش شود" 'Maftol bayad Pill va Polish Shavad
        ElseIf RMaftol3.Checked Then
            LMaftolStatus.Text = "درخواست خرید" 'Maftol Bayad kharidari shavad
        ElseIf RMaftol4.Checked Then
            LMaftolStatus.Text = "ارسال شده به پیل و پولیش" 'Maftol baraye pill va polish ersal shode ast
        End If
    End Sub

    Private Sub RMaftol3_CheckedChanged(sender As Object, e As EventArgs) Handles RMaftol3.CheckedChanged
        If RMaftol1.Checked Then
            LMaftolStatus.Text = "موجود است" ' Maftol Mojod Ast
        ElseIf RMaftol2.Checked Then
            LMaftolStatus.Text = "پیل و پولیش شود" 'Maftol bayad Pill va Polish Shavad
        ElseIf RMaftol3.Checked Then
            LMaftolStatus.Text = "درخواست خرید" 'Maftol Bayad kharidari shavad
        ElseIf RMaftol4.Checked Then
            LMaftolStatus.Text = "ارسال شده به پیل و پولیش" 'Maftol baraye pill va polish ersal shode ast
        End If
    End Sub

    Private Sub RMaftol4_CheckedChanged(sender As Object, e As EventArgs) Handles RMaftol4.CheckedChanged
        If RMaftol1.Checked Then
            LMaftolStatus.Text = "موجود است" ' Maftol Mojod Ast
        ElseIf RMaftol2.Checked Then
            LMaftolStatus.Text = "پیل و پولیش شود" 'Maftol bayad Pill va Polish Shavad
        ElseIf RMaftol3.Checked Then
            LMaftolStatus.Text = "درخواست خرید" 'Maftol Bayad kharidari shavad
        ElseIf RMaftol4.Checked Then
            LMaftolStatus.Text = "ارسال شده به پیل و پولیش" 'Maftol baraye pill va polish ersal shode ast
        End If
    End Sub

    Private Sub LMandrelInventory_Click(sender As Object, e As EventArgs) Handles LMandrelInventory.Click
        'TODO: Check to see if mandrel is in the inventory
        ' Dim mandrelState As Boolean
        If IsNumeric(LMandrelDiameter.Text) Then
            Using cn As New OleDbConnection(connectionString)
                Using cmd As New OleDbCommand With {.Connection = cn}
                    cmd.CommandText = "SELECT COUNT(*) FROM mandrels WHERE mandrelDiameter = '" + LMandrelDiameter.Text + "' ;"
                    Try
                        cn.Open()
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
    Private Function LoadEmkansanjiTable()
        Using cn As New OleDbConnection(connectionString)
            Using cmd As New OleDbCommand With {.Connection = cn}

                Dim emkanSanjiColumnNames As String = " springDataBase.productName, emkansanji.quantity, emkansanji.letterNo, customers.customerName "

                'the paranthesis in the query are mandatory
                cmd.CommandText = "SELECT " & ESColumnNames & " FROM (emkansanji INNER JOIN springDataBase ON emkansanji.productID = springDataBase.ID) INNER JOIN customers ON emkansanji.customerID = customers.ID" &
                    " ;"

                Dim dt As New DataTable With {.TableName = "emkansanji"}
                'Try
                cn.Open()
                Dim ds As New DataSet
                Dim emkansanji As New DataTable With {.TableName = "emkansanji"}
                ds.Tables.Add(emkansanji)
                ds.Load(cmd.ExecuteReader(), LoadOption.OverwriteChanges, emkansanji)
                DataGridView1.DataSource = ds.Tables("emkansanji")
                'DataGridView1.Columns(0).Visible = False
                cn.Close()
                DataGridView1.Columns(0).Visible = False
                DataGridView1.Columns(1).Visible = False
                DataGridView1.Columns(2).Visible = False
                DataGridView1.Columns(3).Visible = False
                DataGridView1.Columns(4).Visible = False
                DataGridView1.Columns(5).Visible = False
                DataGridView1.Columns(6).Visible = False

            End Using
        End Using
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
End Class