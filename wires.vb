Imports System.Data.OleDb

Public Class wires
    Private Async Sub wires_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '' Form Load
        Dim sql_command = "SELECT * FROM wireInventory;"
        Dim dt = Await Task(Of DataTable).Run(Function() LoadDataTable(sql_command))
        DataGridView1.DataSource = dt
    End Sub

    Private Async Sub BTUpdateInventory_Click(sender As Object, e As EventArgs) Handles BTUpdateInventory.Click

        '' This subroutine will generate the inventory Database from excel files provided by rahkaran
        '' Requires ImportExceltoDatatable function

        '' Delete everything in wire inventory

        'Dim dt = Await Task(Of DataTable).Run(Function() LoadDataTable("DELETE FROM wireInventory"))

        '' import garm inventory data to a data table
        Dim garmDataTable = Await Task(Of DataTable).Run(Function() ImportExceltoDatatable(excelInventoryGarmPath, "موجودی مواد خط گرم"))
        '' import sard inventory data to a data table
        Dim sardDataTable = Await Task(Of DataTable).Run(Function() ImportExceltoDatatable(excelInventorySardPath, "موجودی مواد خط سرد"))

        Dim purchasedDataTable = Await Task(Of DataTable).Run(Function() ImportExceltoDatatable(excelInventoryPurchasedPath, "موجودی مواد خریداری شده قطعی"))


        Dim wireDiameter As String
        Dim wireLength As String
        Dim wireWeight As String

        Using cn As New OleDbConnection(connectionString)
            Await cn.OpenAsync()
            Using tran = cn.BeginTransaction()
                Using cmd As New OleDbCommand With {.Connection = cn, .Transaction = tran}

                    Try
                        '' Delete everything in wire inventory
                        cmd.CommandText = "DELETE FROM wireInventory"
                        Await cmd.ExecuteNonQueryAsync()

                        '' Populate the inventory table with data of garm file
                        For Each row As DataRow In garmDataTable.Rows
                            wireDiameter = (Val(row("کد").ToString().Substring(2, 4)) / 100).ToString
                            wireLength = 1000 '' TODO: change this 
                            If IsNumeric(wireLength) Then
                                wireWeight = CalculateWireWeight(Val(wireDiameter), Val(wireLength)).ToString
                            Else
                                wireWeight = "-"
                            End If
                            cmd.CommandText = String.Format("INSERT INTO wireInventory 
                        (wireCode, wireType, wireDiameter, wireLength, wireSpecification, inventory , inventoryName, wireWeight) 
                        VALUES ( '{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}' );", row("کد").ToString(), "مفتول شاخه‌ای",
                                       wireDiameter, "1000", row("مشخصه فنی").ToString(), row("مانده (اصلی)").ToString(), row("عنوان").ToString(), wireWeight)
                            Await cmd.ExecuteNonQueryAsync()
                        Next row

                        '' Populate the inventory table with data of sard file
                        For Each row As DataRow In sardDataTable.Rows
                            wireDiameter = (Val(row("کد").ToString().Substring(2, 4)) / 100).ToString
                            cmd.CommandText = String.Format("INSERT INTO wireInventory 
                        (wireCode, wireType, wireDiameter, wireLength, wireSpecification, inventory , inventoryName, wireWeight) 
                        VALUES ( '{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}' );", row("کد").ToString(), "مفتول کویل",
                                       wireDiameter, "-", row("مشخصه فنی").ToString(), row("مانده (اصلی)").ToString(), row("عنوان").ToString(), "-")
                            Await cmd.ExecuteNonQueryAsync()
                        Next row

                        '' Populate the inventory table with data of purchased wires file
                        '' TODO: Add wire weight 
                        For Each row As DataRow In purchasedDataTable.Rows

                            If row("کد").ToString() = "" Then
                                '' prevent empty rows in the files to be inserted in the database
                                Continue For
                            End If

                            cmd.CommandText = String.Format("INSERT INTO wireInventory 
                        (wireCode, wireType, wireDiameter, wireLength, wireSpecification, inventory , inventoryName) 
                        VALUES ( '{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}' );", row("کد").ToString(), row("نوع مفتول").ToString(),
                                       row("قطر مفتول").ToString(), row("طول مفتول").ToString(), row("مشخصه فنی").ToString(), row("موجودی").ToString(), row("عنوان").ToString())
                            Console.WriteLine(cmd.CommandText)
                            Await cmd.ExecuteNonQueryAsync()
                        Next row

                    Catch ex As Exception
                        MsgBox("انتقال اطلاعات موجودی مواد با خطا مواجه شد", vbCritical + MsgBoxStyle.MsgBoxRight, "خطا")
                        Logger.LogFatal(ex.Message, ex)
                        tran.Rollback()
                        cn.Close()
                        Exit Sub
                    End Try

                    tran.Commit()
                    cn.Close()
                    MsgBox("بروزرسانی اطلاعات موجودی مواد با موفقیت انجام شد.", MsgBoxStyle.MsgBoxRtlReading + MsgBoxStyle.MsgBoxRight + vbInformation, "بروزرسانی اطلاعات مواد")
                End Using
            End Using
        End Using
    End Sub

    Private Async Sub BTTest_Click(sender As Object, e As EventArgs) Handles BTTest.Click
        Dim sql_command = "SELECT IIF( ISNUMERIC(wireLength) , wireLength*5 , '-') as [test] FROM wireInventory;"
        'Dim sql_command = "SELECT IIF( wireType = 'مفتول شاخه‌ای' , wireLength*5 , '-') as [test] FROM wireInventory;"
        Dim dt = Await Task(Of DataTable).Run(Function() LoadDataTable(sql_command))
        DataGridView1.DataSource = dt
    End Sub

    Private Sub BTModify_Click(sender As Object, e As EventArgs) Handles BTModify.Click
        '' Opens another instance of a form
        ShowForm(New mainForm)
    End Sub

    Private Sub ShowForm(WhichForm As Form)
        With WhichForm
            .Show()
            .BringToFront()
        End With
    End Sub
End Class