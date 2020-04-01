Imports System.ComponentModel
Imports System.Data.OleDb
Public Class ChangePasswordForm
    Private Async Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TBpass.Text = TBpass2.Text Then
            Using cn = GetDatabaseCon()
                Using cmd = cn.CreateCommand()
                    cmd.CommandText = "UPDATE users SET" &
                         " users.password = '" & GetSaltedHash(TBpass.Text, "salt") & "'" &
                         " WHERE userID = '" & loggedInUser & "';"
                    Try
                        Await cn.OpenAsync()
                        Await cmd.ExecuteNonQueryAsync()
                        cn.Close()
                        MsgBox("کلمه عبور با موفقیت تغییر کرد", vbInformation + MsgBoxStyle.MsgBoxRight, "تغییر کلمه عبور")
                        Logger.LogInfo("Password Changed")
                    Catch ex As Exception
                        MsgBox("خطا در اتصال به دیتابیس. پارامتر های ورودی را چک کرده و مجددا سعی کنید", vbCritical + vbMsgBoxRight, "خطا در اتصال")
                        Logger.LogFatal(ex.Message, ex)
                    End Try
                End Using
            End Using
        Else
            MsgBox("عبارات وارد شده با هم یکسان نیستند", vbCritical + MsgBoxStyle.MsgBoxRight, "تغییر کلمه عبور")
        End If
        Me.Close()
    End Sub

    Private Sub ChangePasswordForm_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        LoginForm.Show()
    End Sub

    Private Sub ChangePasswordForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If loggedInUserGroup = "Admin" Then BTPasswordReset.Visible = True
    End Sub

    Private Async Sub BTPasswordReset_Click(sender As Object, e As EventArgs) Handles BTPasswordReset.Click
        Using cn = GetDatabaseCon()
            Using cmd = cn.CreateCommand()
                cmd.CommandText = "UPDATE users SET" &
                     " users.password = '" & GetSaltedHash(TBpass2.Text, "salt") & "'" &
                     " WHERE userID = '" & TBpass.Text & "';"
                'Console.WriteLine(cmd.CommandText)
                Try
                    Await cn.OpenAsync()
                    Await cmd.ExecuteNonQueryAsync()
                    cn.Close()
                    MsgBox("کلمه عبور با موفقیت تغییر کرد", vbInformation + MsgBoxStyle.MsgBoxRight, "تغییر کلمه عبور")
                    Logger.LogInfo("Password Changed")
                Catch ex As Exception
                    MsgBox("خطا در اتصال به دیتابیس. پارامتر های ورودی را چک کرده و مجددا سعی کنید", vbCritical + vbMsgBoxRight, "خطا در اتصال")
                    Logger.LogFatal(ex.Message, ex)
                End Try
            End Using
        End Using
    End Sub
End Class