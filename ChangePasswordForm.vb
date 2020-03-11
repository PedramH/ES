Imports System.ComponentModel
Imports System.Data.OleDb
Public Class ChangePasswordForm
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TBpass.Text = TBpass2.Text Then
            Using cn As New OleDbConnection(connectionString)
                Using cmd As New OleDbCommand With {.Connection = cn}
                    cmd.CommandText = "UPDATE users SET" &
                         " users.password = '" & TBpass.Text & "'" &
                         " WHERE userID = '" & loggedInUser & "';"
                    Try
                        cn.Open()
                        cmd.ExecuteReader()
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
    End Sub

    Private Sub ChangePasswordForm_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        LoginForm.Show()
    End Sub
End Class