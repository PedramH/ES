Imports System.Data.OleDb
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
    Private Sub Authenticate()
        Using cn As New OleDbConnection(connectionString)
            Using cmd As New OleDbCommand With {.Connection = cn}
                cmd.CommandText = "SELECT COUNT(*) FROM users WHERE userID = '" + TBuserName.Text + "';"
                Try
                    cn.Open()
                    Dim i As String = cmd.ExecuteScalar()
                    cn.Close()
                    If i > 0 Then
                        Dim dt As New DataTable With {.TableName = "userDataTable"}
                        cmd.CommandText = "SELECT * FROM users WHERE userID = '" + TBuserName.Text + "';"
                        cn.Open()
                        Dim ds As New DataSet
                        Dim userDataTable As New DataTable With {.TableName = "userDataTable"}
                        ds.Tables.Add(userDataTable)
                        ds.Load(cmd.ExecuteReader(), LoadOption.OverwriteChanges, userDataTable)
                        cn.Close()
                        'MsgBox(userDataTable.Rows(0)("password"))
                        If TBPassword.Text = userDataTable.Rows(0)("password") Then
                            loggedInUser = userDataTable.Rows(0)("userID")
                            loggedInUserName = userDataTable.Rows(0)("userName")
                            loggedInUserGroup = userDataTable.Rows(0)("userGroup")
                            mainForm.Show()
                            Me.Close()
                        Else
                            MsgBox("کلمه عبور وارد شده اشتباه است", vbCritical, "ورود ناموفق")
                        End If

                    Else
                        MsgBox("نام کاربری وارد شده اشتباه است", vbCritical, "ورود ناموفق")
                    End If
                Catch ex As Exception
                    MsgBox("خطا در اتصال به دیتابیس. پارامتر های ورودی را چک کرده و مجددا سعی کنید", vbCritical + vbMsgBoxRight, "خطا در اتصال")
                    Logger.LogFatal(ex.Message, ex)
                End Try

            End Using
        End Using
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Authenticate()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Using cn As New OleDbConnection(connectionString)
            Using cmd As New OleDbCommand With {.Connection = cn}
                cmd.CommandText = "SELECT COUNT(*) FROM users WHERE userID = '" + TBuserName.Text + "';"
                Try
                    cn.Open()
                    Dim i As String = cmd.ExecuteScalar()
                    cn.Close()
                    If i > 0 Then
                        Dim dt As New DataTable With {.TableName = "userDataTable"}
                        cmd.CommandText = "SELECT * FROM users WHERE userID = '" + TBuserName.Text + "';"
                        cn.Open()
                        Dim ds As New DataSet
                        Dim userDataTable As New DataTable With {.TableName = "userDataTable"}
                        ds.Tables.Add(userDataTable)
                        ds.Load(cmd.ExecuteReader(), LoadOption.OverwriteChanges, userDataTable)
                        cn.Close()
                        'MsgBox(userDataTable.Rows(0)("password"))
                        If TBPassword.Text = userDataTable.Rows(0)("password") Then
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
                Catch ex As Exception
                    MsgBox("خطا در اتصال به دیتابیس. پارامتر های ورودی را چک کرده و مجددا سعی کنید", vbCritical + vbMsgBoxRight, "خطا در اتصال")
                    Logger.LogFatal(ex.Message, ex)
                End Try

            End Using
        End Using
    End Sub

    Private Sub LoginForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If debugMode = True Then
            mainForm.Show()
            Me.Close()
        End If
    End Sub
End Class