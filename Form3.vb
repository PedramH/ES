Imports System.ComponentModel
Public Class customerForm

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

    Private Async Sub BTModifyCustomer_Click(sender As Object, e As EventArgs) Handles BTModifyCustomer.Click
        Dim answer As String = MsgBox("در صورت تایید مشخصات مشتری به صورتی دائمی تغییر خواهد کرد", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical + MsgBoxStyle.DefaultButton2 + MsgBoxStyle.MsgBoxRight, Title:="ویرایش مشخصات مشتری")
        If answer = vbOK Then
            'Using cn As New OleDbConnection(connectionString)
            '    Using cmd As New OleDbCommand With {.Connection = cn}
            '        cmd.CommandText = "UPDATE customers SET" &
            '        " customerName = '" & TBCustomerName.Text & "'," &
            '        " fieldOfWork = '" & TBFieldOfWork.Text & "'," &
            '        " shenaseMelli = '" & TBShenaseMelli.Text & "'," &
            '        " codeEghtesadi = '" & TBCodeEghtesadi.Text & "'," &
            '        " postCode = '" & TBPostCode.Text & "'," &
            '        " ads1 = '" & TBAds1.Text & "'," &
            '        " ads2 = '" & TBAds2.Text & "'," &
            '        " p1 = '" & TBP_1_Name.Text & "'," &
            '        " p1_job = '" & TBP_1_job.Text & "'," &
            '        " p1_phone = '" & TBP_1_phone.Text & "'," &
            '        " p1_mobile = '" & TBP_1_mobile.Text & "'," &
            '        " p1_email = '" & TBP_1_email.Text & "'," &
            '        " p2 = '" & TBP_2_Name.Text & "'," &
            '        " p2_job = '" & TBP_2_job.Text & "'," &
            '        " p2_phone = '" & TBP_2_phone.Text & "'," &
            '        " p2_mobile = '" & TBP_2_mobile.Text & "'," &
            '        " p2_email = '" & TBP_2_email.Text & "'," &
            '        " p3 = '" & TBP_3_Name.Text & "'," &
            '        " p3_job = '" & TBP_3_job.Text & "'," &
            '        " p3_phone = '" & TBP_3_phone.Text & "'," &
            '        " p3_mobile = '" & TBP_3_mobile.Text & "'," &
            '        " p3_email = '" & TBP_3_email.Text & "'," &
            '        " requirements = '" & TBRequierments.Text & "'," &
            '        " comment = '" & TBComment.Text & "'" &
            '        "WHERE ID = " & TBdbID.Text & ";"

            '        Try
            '            cn.Open()
            '            cmd.ExecuteReader()
            '            cn.Close()
            '            MsgBox("ویرایش اطلاعات با موفقیت انجام شد", vbInformation, "ویرایش مشخصات مشتری")
            '            Logger.LogInfo("Modified information of Customer Name : " + TBCustomerName.Text + " - Customer ID: " + TBdbID.Text)
            '        Catch ex As Exception
            '            MsgBox("خطا در ثبت اطلاعات مشتری. پارامتر های ورودی را کنترل کرده و مجددا سعی کنید", vbCritical + vbMsgBoxRight, "خطا در ثبت اطلاعات مشتری")
            '            Logger.LogFatal(ex.Message, ex)
            '        End Try

            '    End Using
            'End Using
            Using cn = GetDatabaseCon()
                Using cmd = cn.CreateCommand()
                    cmd.CommandText = "UPDATE customers SET" &
                    " customerName = '" & TBCustomerName.Text & "'," &
                    " fieldOfWork = '" & TBFieldOfWork.Text & "'," &
                    " shenaseMelli = '" & TBShenaseMelli.Text & "'," &
                    " codeEghtesadi = '" & TBCodeEghtesadi.Text & "'," &
                    " postCode = '" & TBPostCode.Text & "'," &
                    " ads1 = '" & TBAds1.Text & "'," &
                    " ads2 = '" & TBAds2.Text & "'," &
                    " p1 = '" & TBP_1_Name.Text & "'," &
                    " p1_job = '" & TBP_1_job.Text & "'," &
                    " p1_phone = '" & TBP_1_phone.Text & "'," &
                    " p1_mobile = '" & TBP_1_mobile.Text & "'," &
                    " p1_email = '" & TBP_1_email.Text & "'," &
                    " p2 = '" & TBP_2_Name.Text & "'," &
                    " p2_job = '" & TBP_2_job.Text & "'," &
                    " p2_phone = '" & TBP_2_phone.Text & "'," &
                    " p2_mobile = '" & TBP_2_mobile.Text & "'," &
                    " p2_email = '" & TBP_2_email.Text & "'," &
                    " p3 = '" & TBP_3_Name.Text & "'," &
                    " p3_job = '" & TBP_3_job.Text & "'," &
                    " p3_phone = '" & TBP_3_phone.Text & "'," &
                    " p3_mobile = '" & TBP_3_mobile.Text & "'," &
                    " p3_email = '" & TBP_3_email.Text & "'," &
                    " requirements = '" & TBRequierments.Text & "'," &
                    " comment = '" & TBComment.Text & "'" &
                    "WHERE ID = " & TBdbID.Text & ";"

                    Try
                        Await cn.OpenAsync()
                        Await cmd.ExecuteNonQueryAsync()
                        cn.Close()
                        MsgBox("ویرایش اطلاعات با موفقیت انجام شد", vbInformation, "ویرایش مشخصات مشتری")
                        Logger.LogInfo("Modified information of Customer Name : " + TBCustomerName.Text + " - Customer ID: " + TBdbID.Text)
                    Catch ex As Exception
                        MsgBox("خطا در ثبت اطلاعات مشتری. پارامتر های ورودی را کنترل کرده و مجددا سعی کنید", vbCritical + vbMsgBoxRight, "خطا در ثبت اطلاعات مشتری")
                        Logger.LogFatal(cmd.CommandText, ex)
                    End Try

                End Using
            End Using



        End If
    End Sub

    Private Async Sub BTDeleteCustomer_Click(sender As Object, e As EventArgs) Handles BTDeleteCustomer.Click
        Dim answer As String = MsgBox("در صورت تایید مشخصات این مشتری به صورت دائمی حذف خواهد شد", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical + MsgBoxStyle.DefaultButton2 + MsgBoxStyle.MsgBoxRight, Title:="حذف اطلاعات مشتری")
        If answer = vbOK Then
            'Dim i As String ' Number of emkansanji with this customer
            ''Check to see if an emkansanji with this customer is present
            'Try
            '    Using cn As New OleDbConnection(connectionString)
            '        Using cmd As New OleDbCommand With {.Connection = cn}
            '            cmd.CommandText = "SELECT COUNT(*) FROM emkansanji Where customerID = " & TBdbID.Text & " ;"
            '            'Try
            '            cn.Open()
            '            i = cmd.ExecuteScalar
            '            cn.Close()
            '        End Using
            '    End Using

            '    If i = 0 Then
            '        Using cn As New OleDbConnection(connectionString)
            '            Using cmd As New OleDbCommand With {.Connection = cn}
            '                cmd.CommandText = "DELETE FROM customers Where ID = " & TBdbID.Text & " ;"
            '                'Try
            '                cn.Open()
            '                cmd.ExecuteReader()
            '                cn.Close()
            '                MsgBox("مشتری از دیتابیس حذف شد", vbInformation, "حذف مشتری")
            '                Logger.LogInfo("Deleting customer  Name: " + TBCustomerName.Text + " - ID: " + TBdbID.Text)
            '            End Using
            '        End Using
            '    Else
            '        MsgBox("امکان سنجی این مشتری در دیتابیس ثبت شده است، بنابراین امکان حذف آن وجود ندارد", vbCritical + MsgBoxStyle.MsgBoxRight, "حذف مشتری")
            '    End If
            'Catch ex As Exception
            '    MsgBox("خطا در حذف اطلاعات مشتری. پارامتر های ورودی را کنترل کرده و مجددا سعی کنید", vbCritical + vbMsgBoxRight, "خطا در حذف اطلاعات مشتری")
            '    Logger.LogFatal(ex.Message, ex)
            'End Try


            Dim i As String ' Number of emkansanji with this customer
            'Check to see if an emkansanji with this customer is present
            Try
                Using cn = GetDatabaseCon()
                    Using cmd = cn.CreateCommand
                        cmd.CommandText = "SELECT COUNT(*) FROM emkansanji Where customerID = " & TBdbID.Text & " ;"
                        'Try
                        Await cn.OpenAsync()
                        i = cmd.ExecuteScalar()
                        cn.Close()
                    End Using
                End Using

                If i = 0 Then
                    Using cn = GetDatabaseCon()
                        Using cmd = cn.CreateCommand()
                            cmd.CommandText = "DELETE FROM customers Where ID = " & TBdbID.Text & " ;"
                            'Try
                            Await cn.OpenAsync()
                            Await cmd.ExecuteNonQueryAsync()
                            cn.Close()
                            MsgBox("مشتری از دیتابیس حذف شد", vbInformation, "حذف مشتری")
                            Logger.LogInfo("Deleting customer  Name: " + TBCustomerName.Text + " - ID: " + TBdbID.Text)
                        End Using
                    End Using
                Else
                    MsgBox("امکان سنجی این مشتری در دیتابیس ثبت شده است، بنابراین امکان حذف آن وجود ندارد", vbCritical + MsgBoxStyle.MsgBoxRight, "حذف مشتری")
                End If
            Catch ex As Exception
                MsgBox("خطا در حذف اطلاعات مشتری. پارامتر های ورودی را کنترل کرده و مجددا سعی کنید", vbCritical + vbMsgBoxRight, "خطا در حذف اطلاعات مشتری")
                Logger.LogFatal(ex.Message, ex)
            End Try


        End If
    End Sub

    Private Async Sub BTNewCustomer_Click(sender As Object, e As EventArgs) Handles BTNewCustomer.Click
        Dim answer As String = MsgBox("در صورت تایید مشتری جدید با مشخصات ذکر شده ثبت خواهد شد", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical + MsgBoxStyle.DefaultButton2 + MsgBoxStyle.MsgBoxRight, Title:="ثبت محصول جدید")
        If answer = vbOK Then
            'Using cn As New OleDbConnection(connectionString)
            '    Using cmd As New OleDbCommand With {.Connection = cn}

            '        Dim columnNames As String = " ( customerName , fieldOfWork , shenaseMelli , codeEghtesadi, postCode, " &
            '         " ads1 , ads2 , p1, p1_job , p1_phone , " &
            '         " p1_mobile , p1_email , p2 , p2_job , p2_phone , " &
            '        " p2_mobile, p2_email ,p3,p3_job,p3_phone,p3_mobile,p3_email,requirements,  comment ) "

            '        Dim valueString As String = "('" & TBCustomerName.Text & "','" & TBFieldOfWork.Text & "','" & TBShenaseMelli.Text & "','" & TBCodeEghtesadi.Text & "','" & TBPostCode.Text & "','" &
            '            TBAds1.Text & "','" & TBAds2.Text & "','" & TBP_1_Name.Text & "','" & TBP_1_job.Text & "','" & TBP_1_phone.Text & "','" &
            '            TBP_1_mobile.Text & "','" & TBP_1_email.Text & "','" & TBP_2_Name.Text & "','" & TBP_2_job.Text & "','" & TBP_2_phone.Text & "','" &
            '            TBP_2_mobile.Text & "','" & TBP_2_email.Text & "','" & TBP_3_Name.Text & "','" & TBP_3_job.Text & "','" & TBP_3_phone.Text & "','" & TBP_3_mobile.Text & "','" & TBP_3_email.Text & "','" & TBRequierments.Text & "','" & TBComment.Text & "' )"

            '        cmd.CommandText = "INSERT INTO customers" & columnNames & " VALUES " & valueString & ";"
            '        Try
            '            cn.Open()
            '            cmd.ExecuteReader()
            '            MsgBox("ثبت اطلاعات مشتری با موفقیت انجام شد", vbInformation, "مشخصات مشتری")
            '            cn.Close()
            '            Logger.LogInfo("New Customer Added With Name = " + TBCustomerName.Text)
            '        Catch ex As Exception
            '            MsgBox("خطا در ثبت اطلاعات مشتری. پارامتر های ورودی را کنترل کرده و مجددا سعی کنید", vbCritical + vbMsgBoxRight, "خطا در ثبت اطلاعات مشتری")
            '            Logger.LogFatal(ex.Message, ex)
            '        End Try
            '    End Using
            'End Using

            Using cn = GetDatabaseCon()
                Using cmd = cn.CreateCommand()
                    Dim columnNames As String = " ( customerName , fieldOfWork , shenaseMelli , codeEghtesadi, postCode, " &
                     " ads1 , ads2 , p1, p1_job , p1_phone , " &
                     " p1_mobile , p1_email , p2 , p2_job , p2_phone , " &
                    " p2_mobile, p2_email ,p3,p3_job,p3_phone,p3_mobile,p3_email,requirements,  comment ) "

                    Dim valueString As String = "('" & TBCustomerName.Text & "','" & TBFieldOfWork.Text & "','" & TBShenaseMelli.Text & "','" & TBCodeEghtesadi.Text & "','" & TBPostCode.Text & "','" &
                        TBAds1.Text & "','" & TBAds2.Text & "','" & TBP_1_Name.Text & "','" & TBP_1_job.Text & "','" & TBP_1_phone.Text & "','" &
                        TBP_1_mobile.Text & "','" & TBP_1_email.Text & "','" & TBP_2_Name.Text & "','" & TBP_2_job.Text & "','" & TBP_2_phone.Text & "','" &
                        TBP_2_mobile.Text & "','" & TBP_2_email.Text & "','" & TBP_3_Name.Text & "','" & TBP_3_job.Text & "','" & TBP_3_phone.Text & "','" & TBP_3_mobile.Text & "','" & TBP_3_email.Text & "','" & TBRequierments.Text & "','" & TBComment.Text & "' )"

                    cmd.CommandText = "INSERT INTO customers" & columnNames & " VALUES " & valueString & ";"
                    Try
                        Await cn.OpenAsync()
                        Await cmd.ExecuteNonQueryAsync()
                        MsgBox("ثبت اطلاعات مشتری با موفقیت انجام شد", vbInformation, "مشخصات مشتری")
                        cn.Close()
                        Logger.LogInfo("New Customer Added With Name = " + TBCustomerName.Text)
                    Catch ex As Exception
                        MsgBox("خطا در ثبت اطلاعات مشتری. پارامتر های ورودی را کنترل کرده و مجددا سعی کنید", vbCritical + vbMsgBoxRight, "خطا در ثبت اطلاعات مشتری")
                        Logger.LogFatal(cmd.CommandText, ex)
                    End Try
                End Using
            End Using
        End If
    End Sub

    Private Sub customerForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Select Case customerFormState
            Case "modify"
                Me.BTNewCustomer.Enabled = False
                Me.BTModifyCustomer.Enabled = True
                Me.BTDeleteCustomer.Enabled = True
                LoadCustomerInfo()
            Case "new"
                Me.BTNewCustomer.Enabled = True
                Me.BTModifyCustomer.Enabled = False
                Me.BTDeleteCustomer.Enabled = False
            Case "view"
                Me.BTNewCustomer.Enabled = False
                Me.BTModifyCustomer.Enabled = False
                Me.BTDeleteCustomer.Enabled = False
                LoadCustomerInfo()
        End Select
    End Sub

    Private Async Sub LoadCustomerInfo()
        '' Load customer info into the form
        'Using cn As New OleDbConnection(connectionString)
        '    Using cmd As New OleDbCommand With {.Connection = cn}
        '        cmd.CommandText = "SELECT * FROM customers WHERE customers.ID = " & TBdbID.Text & ";"

        '        Dim dt As New DataTable With {.TableName = "customers"}

        '        Dim ds As New DataSet
        '        Dim customers As New DataTable With {.TableName = "customers"}
        '        ds.Tables.Add(customers)
        '        Try
        '            cn.Open()
        '            ds.Load(cmd.ExecuteReader(), LoadOption.OverwriteChanges, customers)
        '            cn.Close()
        '        Catch ex As Exception
        '            MsgBox("خطا در دریافت اطلاعات مشتری", vbCritical + vbMsgBoxRight, "خطا")
        '            Logger.LogFatal(ex.Message, ex)
        '        End Try
        '        Me.TBCustomerName.Text = ds.Tables("customers").Rows(0)("customerName").ToString()
        '        Me.TBFieldOfWork.Text = ds.Tables("customers").Rows(0)("fieldOfWork").ToString()
        '        Me.TBShenaseMelli.Text = ds.Tables("customers").Rows(0)("shenaseMelli").ToString()
        '        Me.TBCodeEghtesadi.Text = ds.Tables("customers").Rows(0)("codeEghtesadi").ToString()
        '        Me.TBPostCode.Text = ds.Tables("customers").Rows(0)("postCode").ToString()

        '        Me.TBAds1.Text = ds.Tables("customers").Rows(0)("ads1").ToString()
        '        Me.TBAds2.Text = ds.Tables("customers").Rows(0)("ads2").ToString()

        '        Me.TBP_1_Name.Text = ds.Tables("customers").Rows(0)("p1").ToString()
        '        Me.TBP_1_job.Text = ds.Tables("customers").Rows(0)("p1_job").ToString()
        '        Me.TBP_1_phone.Text = ds.Tables("customers").Rows(0)("p1_phone").ToString()
        '        Me.TBP_1_mobile.Text = ds.Tables("customers").Rows(0)("p1_mobile").ToString()
        '        Me.TBP_1_email.Text = ds.Tables("customers").Rows(0)("p1_email").ToString()

        '        Me.TBP_2_Name.Text = ds.Tables("customers").Rows(0)("p2").ToString()
        '        Me.TBP_2_job.Text = ds.Tables("customers").Rows(0)("p2_job").ToString()
        '        Me.TBP_2_phone.Text = ds.Tables("customers").Rows(0)("p2_phone").ToString()
        '        Me.TBP_2_mobile.Text = ds.Tables("customers").Rows(0)("p2_mobile").ToString()
        '        Me.TBP_2_email.Text = ds.Tables("customers").Rows(0)("p2_email").ToString()

        '        Me.TBP_3_Name.Text = ds.Tables("customers").Rows(0)("p3").ToString()
        '        Me.TBP_3_job.Text = ds.Tables("customers").Rows(0)("p3_job").ToString()
        '        Me.TBP_3_phone.Text = ds.Tables("customers").Rows(0)("p3_phone").ToString()
        '        Me.TBP_3_mobile.Text = ds.Tables("customers").Rows(0)("p3_mobile").ToString()
        '        Me.TBP_3_email.Text = ds.Tables("customers").Rows(0)("p3_email").ToString()

        '    End Using
        'End Using

        Dim sql_command = "SELECT * FROM customers WHERE customers.ID = " & TBdbID.Text & ";"

        Try
            Dim dt = Await Task(Of DataTable).Run(Function() LoadDataTable(sql_command))
            Me.TBCustomerName.Text = dt.Rows(0)("customerName").ToString()
            Me.TBFieldOfWork.Text = dt.Rows(0)("fieldOfWork").ToString()
            Me.TBShenaseMelli.Text = dt.Rows(0)("shenaseMelli").ToString()
            Me.TBCodeEghtesadi.Text = dt.Rows(0)("codeEghtesadi").ToString()
            Me.TBPostCode.Text = dt.Rows(0)("postCode").ToString()

            Me.TBAds1.Text = dt.Rows(0)("ads1").ToString()
            Me.TBAds2.Text = dt.Rows(0)("ads2").ToString()

            Me.TBP_1_Name.Text = dt.Rows(0)("p1").ToString()
            Me.TBP_1_job.Text = dt.Rows(0)("p1_job").ToString()
            Me.TBP_1_phone.Text = dt.Rows(0)("p1_phone").ToString()
            Me.TBP_1_mobile.Text = dt.Rows(0)("p1_mobile").ToString()
            Me.TBP_1_email.Text = dt.Rows(0)("p1_email").ToString()

            Me.TBP_2_Name.Text = dt.Rows(0)("p2").ToString()
            Me.TBP_2_job.Text = dt.Rows(0)("p2_job").ToString()
            Me.TBP_2_phone.Text = dt.Rows(0)("p2_phone").ToString()
            Me.TBP_2_mobile.Text = dt.Rows(0)("p2_mobile").ToString()
            Me.TBP_2_email.Text = dt.Rows(0)("p2_email").ToString()

            Me.TBP_3_Name.Text = dt.Rows(0)("p3").ToString()
            Me.TBP_3_job.Text = dt.Rows(0)("p3_job").ToString()
            Me.TBP_3_phone.Text = dt.Rows(0)("p3_phone").ToString()
            Me.TBP_3_mobile.Text = dt.Rows(0)("p3_mobile").ToString()
            Me.TBP_3_email.Text = dt.Rows(0)("p3_email").ToString()

        Catch ex As Exception
            MsgBox("خطا در دریافت اطلاعات مشتری", vbCritical + vbMsgBoxRight, "خطا")
            Logger.LogFatal(ex.Message, ex)
        End Try


    End Sub

    Private Sub customerForm_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        On Error Resume Next
        FrmNewEmkansanji.BTCustomerSearch.PerformClick()
    End Sub
End Class