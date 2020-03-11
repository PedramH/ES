Imports System.Data.OleDb

Public Class wires
    Private Async Sub wires_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '' Form Load
        Dim sql_command = "SELECT * FROM wireInventory;"
        Dim dt = Await Task(Of DataTable).Run(Function() LoadDataTable(sql_command))
        DataGridView1.DataSource = dt
    End Sub

    Private Sub BTUpdateInventory_Click(sender As Object, e As EventArgs) Handles BTUpdateInventory.Click
        '' TODO: Make this async

        '' This subroutine will generate the inventory Database from excel files provided by rahkaran
        '' Requires ImportExceltoDatatable function
        Dim testDataTable As New DataTable With {.TableName = "garmInventory"}
        testDataTable = ImportExceltoDatatable(excelInventoryGarm)
        DataGridView1.DataSource = testDataTable

        Using cn As New OleDbConnection(connectionString)
            Using cmd As New OleDbCommand With {.Connection = cn}
                cn.Open()
                '' Delete The data in the reserve table
                cmd.CommandText = "DELETE FROM wireInventory"
                cmd.ExecuteReader()
                cn.Close()

                cn.Open()
                Dim wireDiameter As String
                Dim wireLength As String
                For Each row As DataRow In testDataTable.Rows
                    Try
                        '    cmd.CommandText = "
                        'INSERT INTO wireInventory (wireCode, wireType, wireDiameter, wireWeight, preReserve, reserve)
                        'VALUES (" + row("wireCode").ToString() + " , " + "1, " + row("رزرو امکان سنجی").ToString() + " , " + row("رزرو تولید").ToString() + " ) ;"
                        wireDiameter = (Val(row("کد").ToString().Substring(2, 4)) / 100).ToString

                        cmd.CommandText = String.Format("INSERT INTO wireInventory 
                        (wireCode, wireType, wireDiameter, wireLength, wireSpecification, inventory) 
                        VALUES ( '{0}', '{1}', '{2}', '{3}', '{4}', '{5}' );", row("کد").ToString(), "مفتول شاخه‌ای", wireDiameter, "1000", row("مشخصه فنی").ToString(), row("مانده (اصلی)").ToString())
                        Dim dr As OleDbDataReader
                        dr = cmd.ExecuteReader
                        dr.Close()
                    Catch ex As Exception
                        Console.WriteLine(cmd.CommandText)
                        Console.WriteLine(ex.Message)
                        'Console.WriteLine(row("wireCode").ToString())
                        'Console.WriteLine(row("رزرو امکان سنجی").ToString())
                        'Console.WriteLine(row("رزرو تولید").ToString())
                    End Try

                Next row
                cn.Close()
            End Using
        End Using
    End Sub
End Class