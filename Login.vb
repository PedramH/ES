
Public Class LoginForm
    Protected Overrides Function ProcessCmdKey(ByRef msg As System.Windows.Forms.Message,
                                           ByVal keyData As System.Windows.Forms.Keys) _
                                           As Boolean

        If msg.WParam.ToInt32() = CInt(Keys.Enter) AndAlso TypeOf Me.ActiveControl Is TextBox Then
            Authenticate()
            Return True
        End If

        Return MyBase.ProcessCmdKey(msg, keyData)
    End Function
    Private Async Sub Authenticate()

        'Using cn = GetDatabaseCon()
        '    Using cmd = cn.CreateCommand()
        '        cmd.CommandText = "SELECT COUNT(*) FROM users WHERE userID = '" + TBuserName.Text + "';"
        '        Try
        '            Await cn.OpenAsync()
        '            Dim i As String = cmd.ExecuteScalar()
        '            cn.Close()
        '            If i > 0 Then
        '                Dim dt As New DataTable With {.TableName = "userDataTable"}
        '                cmd.CommandText = "SELECT * FROM users WHERE userID = '" + TBuserName.Text + "';"
        '                cn.Open()
        '                Dim ds As New DataSet
        '                Dim userDataTable As New DataTable With {.TableName = "userDataTable"}
        '                ds.Tables.Add(userDataTable)
        '                ds.Load(cmd.ExecuteReader(), LoadOption.OverwriteChanges, userDataTable)
        '                cn.Close()
        '                'MsgBox(userDataTable.Rows(0)("password"))
        '                If TBPassword.Text = userDataTable.Rows(0)("password") Then
        '                    loggedInUser = userDataTable.Rows(0)("userID")
        '                    loggedInUserName = userDataTable.Rows(0)("userName")
        '                    loggedInUserGroup = userDataTable.Rows(0)("userGroup")
        '                    FrmMenu.Show()
        '                    Me.Close()
        '                Else
        '                    MsgBox("کلمه عبور وارد شده اشتباه است", vbCritical, "ورود ناموفق")
        '                End If

        '            Else
        '                MsgBox("نام کاربری وارد شده اشتباه است", vbCritical, "ورود ناموفق")
        '            End If
        '        Catch ex As Exception
        '            MsgBox("خطا در اتصال به دیتابیس. پارامتر های ورودی را چک کرده و مجددا سعی کنید", vbCritical + vbMsgBoxRight, "خطا در اتصال")
        '            Logger.LogFatal(ex.Message, ex)
        '        End Try

        '    End Using
        'End Using

        Dim i As String = 0
        Me.Cursor = Cursors.WaitCursor
        Using cn = GetDatabaseCon()
            Using cmd = cn.CreateCommand()
                cmd.CommandText = "SELECT COUNT(*) FROM users WHERE userID = '" + TBuserName.Text + "';"
                Try
                    Await cn.OpenAsync()
                    i = cmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    MsgBox("خطا در اتصال به دیتابیس. پارامتر های ورودی را چک کرده و مجددا سعی کنید", vbCritical + vbMsgBoxRight, "خطا در اتصال")
                    Logger.LogFatal(ex.Message, ex)
                End Try
            End Using
        End Using
        Try
            If i > 0 Then

                Dim sql_command = "SELECT * FROM users WHERE userID = '" + TBuserName.Text + "';"
                Dim userDataTable = Await Task(Of DataTable).Run(Function() LoadDataTable(sql_command))
                'MsgBox(userDataTable.Rows(0)("password"))
                If StrComp(GetSaltedHash(TBPassword.Text, "salt"), userDataTable.Rows(0)("password"), False) = 0 Then
                    loggedInUser = userDataTable.Rows(0)("userID")
                    loggedInUserName = userDataTable.Rows(0)("userName")
                    loggedInUserGroup = userDataTable.Rows(0)("userGroup")
                    If CHRemember.Checked Then
                        My.Settings.loggedin = loggedInUser
                        My.Settings.loginDate = Now
                        My.Settings.usersName = loggedInUserName
                        My.Settings.userGroup = loggedInUserGroup
                        My.Settings.validation = GetSaltedHash(loggedInUser & loggedInUserName & loggedInUserGroup, My.Settings.loginDate.ToString())
                    End If
                    FrmMenu.Show()
                    Me.Close()
                Else
                    MsgBox("کلمه عبور وارد شده اشتباه است", vbCritical, "ورود ناموفق")
                End If

            Else
                MsgBox("نام کاربری وارد شده اشتباه است", vbCritical, "ورود ناموفق")
            End If
            Me.Cursor = Cursors.Default
        Catch ex As Exception
            Me.Cursor = Cursors.Default
            MsgBox("خطا در اتصال به دیتابیس. پارامتر های ورودی را چک کرده و مجددا سعی کنید", vbCritical + vbMsgBoxRight, "خطا در اتصال")
            Logger.LogFatal(ex.Message, ex)
        End Try


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Authenticate()
    End Sub

    Private Async Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        '' change password form
        Me.Cursor = Cursors.WaitCursor
        Dim i As String = 0
        Using cn = GetDatabaseCon()
            Using cmd = cn.CreateCommand()
                cmd.CommandText = "SELECT COUNT(*) FROM users WHERE userID = '" + TBuserName.Text + "';"
                Try
                    Await cn.OpenAsync()
                    i = cmd.ExecuteScalar()
                    cn.Close()
                Catch ex As Exception
                    MsgBox("خطا در اتصال به دیتابیس. پارامتر های ورودی را چک کرده و مجددا سعی کنید", vbCritical + vbMsgBoxRight, "خطا در اتصال")
                    Logger.LogFatal(ex.Message, ex)
                End Try
            End Using
        End Using
        Try
            If i > 0 Then

                Dim sql_command = "SELECT * FROM users WHERE userID = '" + TBuserName.Text + "';"
                Dim userDataTable = Await Task(Of DataTable).Run(Function() LoadDataTable(sql_command))
                'MsgBox(userDataTable.Rows(0)("password"))
                'If TBPassword.Text = userDataTable.Rows(0)("password") Then
                '    loggedInUser = userDataTable.Rows(0)("userID")
                '    loggedInUserName = userDataTable.Rows(0)("userName")
                '    loggedInUserGroup = userDataTable.Rows(0)("userGroup")
                '    ChangePasswordForm.Show()
                '    Me.Close()
                'Else
                '    MsgBox("کلمه عبور وارد شده اشتباه است", vbCritical, "ورود ناموفق")
                'End If
                If StrComp(GetSaltedHash(TBPassword.Text, "salt"), userDataTable.Rows(0)("password"), False) = 0 Then
                    loggedInUser = userDataTable.Rows(0)("userID")
                    loggedInUserName = userDataTable.Rows(0)("userName")
                    loggedInUserGroup = userDataTable.Rows(0)("userGroup")
                    ChangePasswordForm.Show()
                    Me.Close()
                Else
                    MsgBox("کلمه عبور وارد شده اشتباه است", vbCritical, "ورود ناموفق")
                End If

            Else
                MsgBox("نام کاربری وارد شده اشتباه است", vbCritical, "ورود ناموفق")
            End If
            Me.Cursor = Cursors.Default
        Catch ex As Exception
            Me.Cursor = Cursors.Default
            MsgBox("خطا در اتصال به دیتابیس. پارامتر های ورودی را چک کرده و مجددا سعی کنید", vbCritical + vbMsgBoxRight, "خطا در اتصال")
            Logger.LogFatal(ex.Message, ex)
        End Try
    End Sub

    Private Sub LoginForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        '' PostgreSQL uses " for alias but access use [] so we change it here
        If db = "postgres" Then
            springDataBaseColumnNames = MigrateAccessToPostgres(springDataBaseColumnNames)
            customerDataBaseColumnNames = MigrateAccessToPostgres(customerDataBaseColumnNames)
            ESColumnNames = MigrateAccessToPostgres(ESColumnNames)
            mandrelsColumnName = MigrateAccessToPostgres(mandrelsColumnName)
            wiresColumnName = MigrateAccessToPostgres(wiresColumnName)
        End If



        If debugMode = True Then
            loggedInUser = "Pedram"
            loggedInUserName = "پدرام یوسفی"
            loggedInUserGroup = "Admin"
            FrmNewEmkansanji.Show()
            Me.Close()
        End If
        If My.Settings.loggedin <> "" And Now.Subtract(My.Settings.loginDate).Days < 7 Then
            Dim validationHash = GetSaltedHash(My.Settings.loggedin & My.Settings.usersName & My.Settings.userGroup, My.Settings.loginDate.ToString())
            If StrComp(validationHash, My.Settings.validation, False) = 0 Then
                loggedInUser = My.Settings.loggedin
                loggedInUserName = My.Settings.usersName
                loggedInUserGroup = My.Settings.userGroup
                FrmMenu.Show()
                Me.Close()
            End If
        End If
    End Sub




End Class