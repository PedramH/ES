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
                'Catch ex As Exception
                ' very common for a developer to simply ignore errors, unwise.
                ' MsgBox("error")
                ' End Try
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

    Private Sub BTModify_Click(sender As Object, e As EventArgs) Handles BTModify.Click
        'TBCustomerName.Text = DataGridView1.SelectedRows(0).Cells("نام مشتری").Value.ToString
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

        TBMValidationDate.Text = DataGridView1.SelectedRows(0).Cells("تاریخ تاییدیه").Value.ToString
        TBMValidationNo.Text = DataGridView1.SelectedRows(0).Cells("شماره تاییدیه").Value.ToString
        CBMOrderState.Text = DataGridView1.SelectedRows(0).Cells("وضعیت سفارش").Value.ToString

        LOrderQuantity.Text = DataGridView1.SelectedRows(0).Cells("تعداد سفارش").Value.ToString
        LOutsideDiameter.Text = DataGridView1.SelectedRows(0).Cells("OD").Value.ToString
        LFreeLength.Text = DataGridView1.SelectedRows(0).Cells("L0").Value.ToString
        LWireDiameter.Text = DataGridView1.SelectedRows(0).Cells("wireDiameter").Value.ToString
        LWireLength.Text = DataGridView1.SelectedRows(0).Cells("wireLength").Value.ToString
        LMandrelDiameter.Text = DataGridView1.SelectedRows(0).Cells("mandrelDiameter").Value.ToString

        Dim wireState As String = DataGridView1.SelectedRows(0).Cells("وضعیت موجودی مفتول").Value.ToString
        If wireState = "موجود است" Then
            RMaftol1.Checked = True
        ElseIf wireState = "پیل و پولیش شود" Then
            RMaftol2.Checked = True
        ElseIf wireState = "درخواست خرید" Then
            RMaftol3.Checked = True
        ElseIf wireState = "ارسال شده به پیل و پولیش" Then
            RMaftol4.Checked = True
        End If

        TabControl1.SelectedTab = TabPage2

    End Sub

    Private Sub BTModifyES_Click(sender As Object, e As EventArgs) Handles BTModifyES.Click
        '    Dim answer As String = MsgBox("در صورت تایید مشخصات این امکان سنجی به صورتی دائمی تغییر خواهد کرد", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical + MsgBoxStyle.DefaultButton2 + MsgBoxStyle.MsgBoxRight, Title:="ویرایش مشخصات امکان سنجی")
        '    If answer = vbOK Then
        '        Using cn As New OleDbConnection(connectionString)
        '            Using cmd As New OleDbCommand With {.Connection = cn}
        '                cmd.CommandText = "UPDATE emkansanji SET" &
        '                " productName = '" & TBProductName.Text & "'," &
        '                 " productID = '" & TBProductID.Text & "'," &
        '                  " wireDiameter = '" & TBWireDiameter.Text & "'," &
        '                  " OD = '" & TBOD.Text & "'," &
        '                  " L0 = '" & TBL0.Text & "'," &
        '                  " Nt = '" & TBNt.Text & "'," &
        '                  " Nactive = '" & TBNActive.Text & "'," &
        '                  " coilingDirection = '" & CBCoilingDirection.Text & "'," &
        '                  " startCoilType = '" & CBScoilType.Text & "'," &
        '                  " endCoilType = '" & CBEcoilType.Text & "'," &
        '                  " mandrelDiameter = '" & TBMandrelDiameter.Text & "'," &
        '                  " wireLength = '" & TBWireLength.Text & "'," &
        '                  " springRate = '" & TBSpringRate.Text & "'," &
        '                  " material = '" & CBMaterial.Text & "'," &
        '                  " pType = '" & CBspringType.Text & "'," &
        '                  " dwgNo = '" & TBDwgNo.Text & "'," &
        '                  " solidStress = '" & TBSolidStress.Text & "'," &
        '                  " comment = '" & TBComment.Text & "'," &
        '                 " F1 = '" & TBF1.Text & "'," &
        '                  " F2 = '" & TBF2.Text & "'," &
        '                 " F3 = '" & TBF3.Text & "'," &
        '                 " L1 = '" & TBL1.Text & "'," &
        '                 " L2 = '" & TBL2.Text & "'," &
        '                 " L3 = '" & TBL2.Text & "'" &
        '                 "WHERE ID = " & TBdbID.Text & ";"

        '                'Try
        '                cn.Open()
        '                cmd.ExecuteReader()
        '                cn.Close()
        '                MsgBox("ویرایش اطلاعات با موفقیت انجام شد", vbInformation, "ویرایش مشخصات محصول")


        '            End Using
        '        End Using
        '    End If
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

    Private Sub Label36_Click(sender As Object, e As EventArgs) Handles Label36.Click
        'TODO: Check to see if mandrel is in the inventory
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
End Class