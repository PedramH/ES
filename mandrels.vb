
Public Class mandrels
    Private Sub mandrels_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Cursor = Cursors.WaitCursor
        ''Loading Mandrels table into datagridview
        SearchMandrelDataBase("")
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub BTCustomerSearch_Click(sender As Object, e As EventArgs) Handles BTCustomerSearch.Click
        SearchMandrelDataBase(TBMandrelSearch.Text)
    End Sub

    Function SearchMandrelDataBase(query As String)
        '' Gets a query value and search the mandrel database by that value
        Using cn = GetDatabaseCon()
            Using cmd = cn.CreateCommand()
                cmd.CommandText = "SELECT " & mandrelsColumnName & " FROM mandrels WHERE mandrelDiameter LIKE '%" + query + "%' ;"
                Using dt As New DataTable With {.TableName = "mandrels"}
                    Try
                        cn.Open()
                        Using ds As New DataSet
                            Dim mandrels As New DataTable With {.TableName = "mandrels"}
                            ds.Tables.Add(mandrels)
                            ds.Load(cmd.ExecuteReader(), LoadOption.OverwriteChanges, mandrels)
                            DataGridView1.DataSource = ds.Tables("mandrels")
                            DataGridView1.Columns(0).Visible = False
                        End Using
                    Catch ex As Exception
                        MsgBox("خطا در ارتباط با دیتابیس", vbCritical + vbMsgBoxRight, "خطا")
                        Logger.LogFatal(ex.Message, ex)
                    Finally
                        cn.Close()
                    End Try
                End Using
            End Using
        End Using
        Return True
    End Function

    Private Sub BTClearCustomer_Click(sender As Object, e As EventArgs) Handles BTClearCustomer.Click

        'Clear Seach Form Then Load the grid view again
        TBMandrelSearch.Text = ""
        SearchMandrelDataBase("")
    End Sub
End Class