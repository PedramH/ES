Imports System.Data.OleDb

Public Class wires
    Private Async Sub wires_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        '' What is visible and what is not
        If wiresFormState = "selection" Then
            BTSelectWire.Visible = True
            wiresFormState = "normal"
        End If

        '' Form Load
        Dim sql_command = "SELECT * FROM wireInventory;"
        Dim dt = Await Task(Of DataTable).Run(Function() LoadDataTable(sql_command))
        DataGridView1.DataSource = dt
    End Sub

    Private Async Sub BTUpdateInventory_Click(sender As Object, e As EventArgs) Handles BTUpdateInventory.Click

        '' This subroutine will generate the inventory Database from excel files provided by rahkaran
        '' Requires ImportExceltoDatatable function

        '' import garm inventory data to a data table
        Dim garmDataTable = Await Task(Of DataTable).Run(Function() ImportExceltoDatatable(excelInventoryGarmPath, "موجودی مواد خط گرم"))
        '' import sard inventory data to a data table
        Dim sardDataTable = Await Task(Of DataTable).Run(Function() ImportExceltoDatatable(excelInventorySardPath, "موجودی مواد خط سرد"))
        '' import purchased inventory data to a data table
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



    Private Async Sub BTModify_Click(sender As Object, e As EventArgs) Handles BTModify.Click
        Await UpdateReservesTable()
    End Sub

    Private Sub ShowForm(WhichForm As Form)
        With WhichForm
            .Show()
            .BringToFront()
        End With
    End Sub

    Private Sub BTSelectWire_Click(sender As Object, e As EventArgs) Handles BTSelectWire.Click
        Dim selectedWire As String = DataGridView1.SelectedRows(0).Cells(0).Value.ToString  'TODO: fix this -> add column name
        Dim selectedWireWeight As String = DataGridView1.SelectedRows(0).Cells("wireWeight").Value.ToString
        Dim selectedWireUnit As String
        If IsNumeric(selectedWireWeight) Then
            selectedWireUnit = "شاخه"
        Else
            selectedWireUnit = "کیلوگرم"
            selectedWireWeight = "-"
        End If
        Select Case wireFormCaller
            Case "wire1"
                emkanSanjiForm.TBMR1.Text = selectedWire
                emkanSanjiForm.Lw1Weight.Text = selectedWireWeight
                emkanSanjiForm.Lw1Unit.Text = selectedWireUnit
            Case "wire2"
                emkanSanjiForm.TBMR2.Text = selectedWire
                emkanSanjiForm.Lw2Weight.Text = selectedWireWeight
                emkanSanjiForm.Lw2Unit.Text = selectedWireUnit
            Case "wire3"
                emkanSanjiForm.TBMR3.Text = selectedWire
                emkanSanjiForm.Lw3Weight.Text = selectedWireWeight
                emkanSanjiForm.Lw3Unit.Text = selectedWireUnit
        End Select
        Me.Dispose()
    End Sub
    Private Async Function UpdateReservesTable() As Task
        Try
            '' ------------------------------------------------  Generating the reserves table  -----------------------------------------------------------
            Dim sql_command = "
                    SELECT wireInventory.wireCode AS [wireCode],
                    SUM(IIF(emkansanji.orderState LIKE '%امکان سنجی%' AND wireInventory.wireCode = emkansanji.r1_code,emkansanji.r1_q,0)) + SUM(IIF(emkansanji.orderState LIKE '%امکان سنجی%' AND wireInventory.wireCode = emkansanji.r2_code,emkansanji.r2_q,0)) + SUM(IIF(emkansanji.orderState LIKE '%امکان سنجی%' AND wireInventory.wireCode = emkansanji.r3_code,emkansanji.r3_q,0)) AS [رزرو امکان سنجی] ,
                    SUM(IIF(emkansanji.orderState LIKE '%تایید%' AND wireInventory.wireCode = emkansanji.r1_code,emkansanji.r1_q,0)) + SUM(IIF(emkansanji.orderState LIKE '%تایید%' AND wireInventory.wireCode = emkansanji.r2_code,emkansanji.r2_q,0))  + SUM(IIF(emkansanji.orderState LIKE '%تایید%' AND wireInventory.wireCode = emkansanji.r3_code,emkansanji.r3_q,0)) AS [رزرو تولید]
                    FROM wireInventory  
                    LEFT JOIN emkansanji ON (wireInventory.wireCode = emkansanji.r1_code OR wireInventory.wireCode = emkansanji.r2_code OR wireInventory.wireCode = emkansanji.r3_code)
                    GROUP BY wireInventory.wireCode
                    ;"
            Dim dt = Await Task(Of DataTable).Run(Function() LoadDataTable(sql_command))
            '' --------------------------------------  Updating the reserves table in the database with new data  ------------------------------------------
            Using cn As New OleDbConnection(connectionString)
                Await cn.OpenAsync()
                Using tran = cn.BeginTransaction()
                    Using cmd As New OleDbCommand With {.Connection = cn, .Transaction = tran}
                        Try
                            '' Delete everything in wire wire reserve table
                            cmd.CommandText = "DELETE FROM wireReserve"
                            Await cmd.ExecuteNonQueryAsync()

                            '' Populate the inventory table with data of reserves query
                            For Each row As DataRow In dt.Rows

                                cmd.CommandText = String.Format("INSERT INTO wireReserve (wireCode, preReserve, reserve) 
                                                                VALUES ('{0}', '{1}', '{2}') ; ", row("wireCode").ToString, row("رزرو امکان سنجی").ToString, row("رزرو تولید").ToString)

                                Await cmd.ExecuteNonQueryAsync()
                            Next row
                        Catch ex As Exception
                            MsgBox("بروزرسانی اطلاعات رزرو مواد با خطا مواجه شد", vbCritical + MsgBoxStyle.MsgBoxRight, "خطا")
                            Logger.LogFatal(ex.Message, ex)
                            tran.Rollback()
                            cn.Close()
                            Exit Function
                        End Try
                        tran.Commit()
                        cn.Close()
                        MsgBox("بروزرسانی اطلاعات رزرو مواد با موفقیت انجام شد.", MsgBoxStyle.MsgBoxRtlReading + MsgBoxStyle.MsgBoxRight + vbInformation, "بروزرسانی اطلاعات مواد")
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MsgBox("بروزرسانی اطلاعات رزرو مواد با خطا مواجه شد.", MsgBoxStyle.MsgBoxRtlReading + MsgBoxStyle.MsgBoxRight + vbInformation, "بروزرسانی اطلاعات مواد")
            Logger.LogFatal(ex.Message, ex)
        End Try
    End Function
End Class