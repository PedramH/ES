Imports System.ComponentModel
Public Class productForm
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

    Private Async Sub BTModify_Click(sender As Object, e As EventArgs) Handles BTModify.Click

        Dim answer As String = MsgBox("در صورت تایید مشخصات محصول به صورتی دائمی تغییر خواهد کرد", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical + MsgBoxStyle.DefaultButton2 + MsgBoxStyle.MsgBoxRight, Title:="ویرایش مشخصات محصول")
        If answer = vbOK Then

            'Using cn As New OleDbConnection(connectionString)
            '    Using cmd As New OleDbCommand With {.Connection = cn}
            '        cmd.CommandText = "UPDATE springDataBase SET" &
            '        " productName = '" & TBProductName.Text & "'," &
            '         " productID = '" & TBProductID.Text & "'," &
            '          " wireDiameter = '" & TBWireDiameter.Text & "'," &
            '          " OD = '" & TBOD.Text & "'," &
            '          " L0 = '" & TBL0.Text & "'," &
            '          " Nt = '" & TBNt.Text & "'," &
            '          " Nactive = '" & TBNActive.Text & "'," &
            '          " coilingDirection = '" & CBCoilingDirection.Text & "'," &
            '          " startCoilType = '" & CBScoilType.Text & "'," &
            '          " endCoilType = '" & CBEcoilType.Text & "'," &
            '          " tipThickness = '" & TBtipThickness.Text & "'," &
            '          " mandrelDiameter = '" & TBMandrelDiameter.Text & "'," &
            '          " wireLength = '" & TBWireLength.Text & "'," &
            '          " springRate = '" & TBSpringRate.Text & "'," &
            '          " material = '" & CBMaterial.Text & "'," &
            '          " pType = '" & CBspringType.Text & "'," &
            '          " dwgNo = '" & TBDwgNo.Text & "'," &
            '          " solidStress = '" & TBSolidStress.Text & "'," &
            '          " solidLoad = '" & TBMaxLoad.Text & "'," &
            '          " comment = '" & TBComment.Text & "'," &
            '         " F1 = '" & TBF1.Text & "'," &
            '          " F2 = '" & TBF2.Text & "'," &
            '         " F3 = '" & TBF3.Text & "'," &
            '         " L1 = '" & TBL1.Text & "'," &
            '         " L2 = '" & TBL2.Text & "'," &
            '         " L3 = '" & TBL2.Text & "'," &
            '         " productionMethod = '" & CBProductionMethod.Text & "'" &
            '         " WHERE ID = " & TBdbID.Text & ";"

            '        Try
            '            cn.Open()
            '            cmd.ExecuteReader()
            '            cn.Close()
            '            Logger.LogInfo("Modified Product with ID = " + TBdbID.Text)
            '            MsgBox("ویرایش اطلاعات با موفقیت انجام شد", vbInformation, "ویرایش مشخصات محصول")
            '        Catch ex As Exception
            '            MsgBox("پارامتر های وارد شده را بررسی کنید و در صورت تکرار اطلاع دهید", vbCritical + MsgBoxStyle.MsgBoxRight, "خطا در ویرایش اطلاعات")
            '            Logger.LogFatal("Modifying Product Data Base", ex)
            '        End Try



            '    End Using
            'End Using
            Me.Cursor = Cursors.WaitCursor
            Using cn = GetDatabaseCon()
                Dim cmd = cn.CreateCommand()
                cmd.CommandText = "UPDATE springDataBase SET" &
                    " productName = '" & TBProductName.Text & "'," &
                     " productID = '" & TBProductID.Text & "'," &
                      " wireDiameter = '" & TBWireDiameter.Text & "'," &
                      " OD = '" & TBOD.Text & "'," &
                      " L0 = '" & TBL0.Text & "'," &
                      " Nt = '" & TBNt.Text & "'," &
                      " Nactive = '" & TBNActive.Text & "'," &
                      " coilingDirection = '" & CBCoilingDirection.Text & "'," &
                      " startCoilType = '" & CBScoilType.Text & "'," &
                      " endCoilType = '" & CBEcoilType.Text & "'," &
                      " tipThickness = '" & TBtipThickness.Text & "'," &
                      " mandrelDiameter = '" & TBMandrelDiameter.Text & "'," &
                      " wireLength = '" & TBWireLength.Text & "'," &
                      " springRate = '" & TBSpringRate.Text & "'," &
                      " material = '" & CBMaterial.Text & "'," &
                      " pType = '" & CBspringType.Text & "'," &
                      " dwgNo = '" & TBDwgNo.Text & "'," &
                      " solidStress = '" & TBSolidStress.Text & "'," &
                      " solidLoad = '" & TBMaxLoad.Text & "'," &
                      " comment = '" & TBComment.Text & "'," &
                      " forceUnit = '" & CBForceUnit.Text & "'," &
                     " F1 = '" & TBF1.Text & "'," &
                      " F2 = '" & TBF2.Text & "'," &
                     " F3 = '" & TBF3.Text & "'," &
                     " L1 = '" & TBL1.Text & "'," &
                     " L2 = '" & TBL2.Text & "'," &
                     " L3 = '" & TBL2.Text & "'," &
                     " productionMethod = '" & CBProductionMethod.Text & "'" &
                     " WHERE ID = " & TBdbID.Text & ";"
                Try
                    Await cn.OpenAsync()
                    Await cmd.ExecuteNonQueryAsync()
                    cn.Close()
                    Me.Cursor = Cursors.Default
                    Logger.LogInfo("Modified Product with ID = " + TBdbID.Text)
                    MsgBox("ویرایش اطلاعات با موفقیت انجام شد", vbInformation, "ویرایش مشخصات محصول")
                Catch ex As Exception
                    Me.Cursor = Cursors.Default
                    MsgBox("پارامتر های وارد شده را بررسی کنید و در صورت تکرار اطلاع دهید", vbCritical + MsgBoxStyle.MsgBoxRight, "خطا در ویرایش اطلاعات")
                    Logger.LogFatal("Modifying Product Data Base", ex)
                End Try
            End Using



        End If

    End Sub

    Private Async Sub BTDelete_Click(sender As Object, e As EventArgs) Handles BTDelete.Click
        Dim answer As String = MsgBox("در صورت تایید مشخصات این محصول به صورت دائمی حذف خواهد شد", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical + MsgBoxStyle.DefaultButton2 + MsgBoxStyle.MsgBoxRight, Title:="حذف محصول")
        Dim i As String = 1
        If answer = vbOK Then

            ''Check to see if an emkansanji with this product is present
            'Using cn As New OleDbConnection(connectionString)
            '    Using cmd As New OleDbCommand With {.Connection = cn}
            '        cmd.CommandText = "SELECT COUNT(*) FROM emkansanji Where productID = " & TBdbID.Text & " ;"
            '        Try
            '            cn.Open()
            '            i = cmd.ExecuteScalar
            '            cn.Close()
            '        Catch ex As Exception
            '            MsgBox("پارامتر های وارد شده را بررسی کنید و در صورت تکرار اطلاع دهید", vbCritical + MsgBoxStyle.MsgBoxRight, "خطا در حذف اطلاعات")
            '            Logger.LogFatal(ex.Message, ex)
            '        End Try

            '    End Using
            'End Using
            'If i = 0 Then
            '    Using cn As New OleDbConnection(connectionString)
            '        Using cmd As New OleDbCommand With {.Connection = cn}
            '            cmd.CommandText = "DELETE FROM springDataBase Where ID = " & TBdbID.Text & " ;"
            '            Try
            '                cn.Open()
            '                cmd.ExecuteReader()
            '                cn.Close()
            '                MsgBox("محصول از دیتابیس حذف شد", vbInformation, "حذف محصول")
            '                Logger.LogInfo("Deleted Product Name = " + TBProductName.Text)
            '            Catch ex As Exception
            '                MsgBox("پارامتر های وارد شده را بررسی کنید و در صورت تکرار اطلاع دهید", vbCritical + MsgBoxStyle.MsgBoxRight, "خطا در حذف اطلاعات")
            '                Logger.LogFatal(ex.Message, ex)
            '            End Try

            '        End Using
            '    End Using
            'Else
            '    MsgBox("این محصول قبلا امکان سنجی شده است و امکان حذف آن وجود ندارد", MsgBoxStyle.Critical + MsgBoxStyle.MsgBoxRight, "حذف محصول")
            'End If

            'Check to see if an emkansanji with this product is present
            Me.Cursor = Cursors.Default
            Using cn = GetDatabaseCon()

                Using cmd = cn.CreateCommand()
                    cmd.CommandText = "SELECT COUNT(*) FROM emkansanji Where productID = " & TBdbID.Text & " ;"
                    Try
                        Await cn.OpenAsync()
                        i = cmd.ExecuteScalar()
                        cn.Close()
                        Me.Cursor = Cursors.Default
                    Catch ex As Exception
                        Me.Cursor = Cursors.Default

                        MsgBox("پارامتر های وارد شده را بررسی کنید و در صورت تکرار اطلاع دهید", vbCritical + MsgBoxStyle.MsgBoxRight, "خطا در حذف اطلاعات")
                        Logger.LogFatal(ex.Message, ex)
                    End Try

                End Using
            End Using
            If i = 0 Then
                Me.Cursor = Cursors.WaitCursor

                Using cn = GetDatabaseCon()
                    Using cmd = cn.CreateCommand()
                        cmd.CommandText = "DELETE FROM springDataBase Where ID = " & TBdbID.Text & " ;"
                        Try
                            Await cn.OpenAsync
                            Await cmd.ExecuteNonQueryAsync()
                            cn.Close()
                            Me.Cursor = Cursors.Default
                            MsgBox("محصول از دیتابیس حذف شد", vbInformation, "حذف محصول")
                            Logger.LogInfo("Deleted Product Name = " + TBProductName.Text)
                        Catch ex As Exception
                            Me.Cursor = Cursors.Default
                            MsgBox("پارامتر های وارد شده را بررسی کنید و در صورت تکرار اطلاع دهید", vbCritical + MsgBoxStyle.MsgBoxRight, "خطا در حذف اطلاعات")
                            Logger.LogFatal(ex.Message, ex)
                        End Try
                    End Using
                End Using
            Else
                MsgBox("این محصول قبلا امکان سنجی شده است و امکان حذف آن وجود ندارد", MsgBoxStyle.Critical + MsgBoxStyle.MsgBoxRight, "حذف محصول")
            End If
        End If
    End Sub

    Private Async Sub BTNew_Click(sender As Object, e As EventArgs) Handles BTNew.Click

        Dim answer As String = MsgBox("در صورت تایید محصولی جدید با مشخصات ذکر شده ثبت خواهد شد", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical + MsgBoxStyle.DefaultButton2 + MsgBoxStyle.MsgBoxRight, Title:="ثبت محصول جدید")
        If answer = vbOK Then
            '' Generating production process
            Dim productionProcess As String
            If CBProductionMethod.Text = "سرد پیچ" Then
                productionProcess = "1010"

            ElseIf CBProductionMethod.Text = "گرم پیچ" Then
                productionProcess = "0101"
            Else
                If Val(TBWireDiameter.Text) < 13 Then
                    productionProcess = "1010"
                Else
                    productionProcess = "0101"
                End If
            End If
            If CBEcoilType.Text = "بسته و سنگ خورده" Or CBEcoilType.Text = "باز و سنگ خورده" Or CBScoilType.Text = "بسته و سنگ خورده" Or CBScoilType.Text = "باز و سنگ خورده" Then
                productionProcess = productionProcess & "111110"
            Else
                productionProcess = productionProcess & "111010"
            End If


            'Using cn As New OleDbConnection(connectionString)
            '    Using cmd As New OleDbCommand With {.Connection = cn}

            '        Dim columnNames As String = " ( productName , productID , wireDiameter , pType, productionMethod, productionProcess ,dwgNo, " &
            '         " OD , L0 , Nt, Nactive , coilingDirection , " &
            '         " mandrelDiameter , wireLength , startCoilType , endCoilType , tipThickness , material , " &
            '        " solidStress, solidLoad , springRate ,F1,L1,F2,L2,F3,L3,  comment ) "

            '        Dim valueString As String = "('" & TBProductName.Text & "','" & TBProductID.Text & "','" & TBWireDiameter.Text & "','" & CBspringType.Text & "','" & CBProductionMethod.Text & "','" & productionProcess & "','" & TBDwgNo.Text & "','" &
            '            TBOD.Text & "','" & TBL0.Text & "','" & TBNt.Text & "','" & TBNActive.Text & "','" & CBCoilingDirection.Text & "','" &
            '            TBMandrelDiameter.Text & "','" & TBWireLength.Text & "','" & CBScoilType.Text & "','" & CBEcoilType.Text & "','" & TBtipThickness.Text & "','" & CBMaterial.Text & "','" &
            '            TBSolidStress.Text & "','" & TBMaxLoad.Text & "','" & TBSpringRate.Text & "','" & TBF1.Text & "','" & TBL1.Text & "','" & TBF2.Text & "','" & TBL2.Text & "','" & TBF3.Text & "','" & TBL3.Text & "','" & TBComment.Text & "' )"

            '        cmd.CommandText = "SELECT * FROM springDataBase WHERE productID = '" & TBProductID.Text & "';"

            '        'Check to see if product ID is a duplicate value
            '        Try
            '            cn.Open()
            '            If cmd.ExecuteReader().HasRows() And TBProductID.Text <> "" Then
            '                MsgBox("کد کالای وارد شده تکرای است", MsgBoxStyle.Critical, "ثبت محصول جدید")
            '                cn.Close()
            '            Else
            '                cn.Close()
            '                cmd.CommandText = "INSERT INTO springDataBase" & columnNames & " VALUES " & valueString & ";"
            '                cn.Open()
            '                cmd.ExecuteReader()
            '                cn.Close()
            '                MsgBox("ثبت محصول با موفقیت انجام شد", vbInformation, "ویرایش مشخصات محصول")
            '                Logger.LogInfo("New product added to the database with Name = " + TBProductName.Text)
            '            End If
            '            'cn.Close()
            '        Catch ex As Exception
            '            MsgBox("پارامتر های وارد شده را بررسی کنید ", vbCritical + MsgBoxStyle.MsgBoxRight, "خطا در ثبت اطلاعات")
            '            Logger.LogFatal(ex.Message, ex)
            '        End Try
            '    End Using
            'End Using
            Me.Cursor = Cursors.WaitCursor

            Using cn = GetDatabaseCon()
                Using cmd = cn.CreateCommand()

                    Dim columnNames As String = " ( productName , productID , wireDiameter , pType, productionMethod, productionProcess ,dwgNo, " &
                     " OD , L0 , Nt, Nactive , coilingDirection , " &
                     " mandrelDiameter , wireLength , startCoilType , endCoilType , tipThickness , material , " &
                    " solidStress, solidLoad , springRate, forceUnit ,F1,L1,F2,L2,F3,L3,  comment ) "

                    Dim valueString As String = "('" & TBProductName.Text & "','" & TBProductID.Text & "','" & TBWireDiameter.Text & "','" & CBspringType.Text & "','" & CBProductionMethod.Text & "','" & productionProcess & "','" & TBDwgNo.Text & "','" &
                        TBOD.Text & "','" & TBL0.Text & "','" & TBNt.Text & "','" & TBNActive.Text & "','" & CBCoilingDirection.Text & "','" &
                        TBMandrelDiameter.Text & "','" & TBWireLength.Text & "','" & CBScoilType.Text & "','" & CBEcoilType.Text & "','" & TBtipThickness.Text & "','" & CBMaterial.Text & "','" &
                        TBSolidStress.Text & "','" & TBMaxLoad.Text & "','" & TBSpringRate.Text & "','" & CBForceUnit.Text & "','" & TBF1.Text & "','" & TBL1.Text & "','" & TBF2.Text & "','" & TBL2.Text & "','" & TBF3.Text & "','" & TBL3.Text & "','" & TBComment.Text & "' )"

                    cmd.CommandText = "SELECT * FROM springDataBase WHERE productID = '" & TBProductID.Text & "';"

                    'Check to see if product ID is a duplicate value
                    Try
                        Await cn.OpenAsync()
                        If cmd.ExecuteReader().HasRows() And TBProductID.Text <> "" Then
                            Me.Cursor = Cursors.Default
                            MsgBox("کد کالای وارد شده تکرای است", MsgBoxStyle.Critical, "ثبت محصول جدید")
                            cn.Close()
                        Else
                            cn.Close()
                            cmd.CommandText = "INSERT INTO springDataBase" & columnNames & " VALUES " & valueString & ";"
                            Await cn.OpenAsync()
                            Await cmd.ExecuteNonQueryAsync()
                            cn.Close()
                            Me.Cursor = Cursors.Default
                            MsgBox("ثبت محصول با موفقیت انجام شد", vbInformation, "ویرایش مشخصات محصول")
                            Logger.LogInfo("New product added to the database with Name = " + TBProductName.Text)
                        End If
                        'cn.Close()
                    Catch ex As Exception
                        Me.Cursor = Cursors.Default

                        MsgBox("پارامتر های وارد شده را بررسی کنید ", vbCritical + MsgBoxStyle.MsgBoxRight, "خطا در ثبت اطلاعات")
                        Logger.LogFatal(ex.Message, ex)
                    End Try
                End Using
            End Using
        End If
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub Label24_Click(sender As Object, e As EventArgs) Handles Label24.Click
        '''' Calculating Spring solid Stress
        Try
            Dim springRate, G, solidLength, solidLoad, meanD, activeCoil As Double
            meanD = TBOD.Text - TBWireDiameter.Text
            activeCoil = TBNActive.Text
            If CBMaterial.Text = "54SiCr6" Then
                G = 79500.5
            Else
                G = 79299.5
            End If
            springRate = (G * TBWireDiameter.Text ^ 4) / (8 * activeCoil * meanD ^ 3)
            Dim groundThickness As Double = (1 - TBtipThickness.Text / 100) * 2
            solidLength = (TBNt.Text + 1 - groundThickness) * TBWireDiameter.Text   'TODO: Compensate for spring end types
            solidLoad = (TBL0.Text - solidLength) * springRate
            TBSolidStress.Text = Math.Round((8 * solidLoad * meanD) / (Math.PI * TBWireDiameter.Text ^ 3), 2)
        Catch ex As Exception
            MsgBox("پارامتر های فنی محصول به درستی وارد نشده است", vbCritical + MsgBoxStyle.MsgBoxRight, "خطا در انجام محاسبه")
            Logger.LogFatal("Error Calculating Solid Stress", ex)
        End Try

    End Sub

    Private Sub Label14_Click(sender As Object, e As EventArgs) Handles Label14.Click

        '''' Calculating Spring Rate
        Try
            Dim springRate, G, meanD, activeCoil As Double
            meanD = TBOD.Text - TBWireDiameter.Text
            activeCoil = TBNActive.Text

            If CBMaterial.Text = "54SiCr6" Then
                G = 79500.5
            Else
                G = 79299.5
            End If
            springRate = (G * TBWireDiameter.Text ^ 4) / (8 * activeCoil * meanD ^ 3)
            TBSpringRate.Text = Math.Round(springRate, 2)
        Catch ex As Exception
            MsgBox("پارامتر های فنی محصول به درستی وارد نشده است", vbCritical + MsgBoxStyle.MsgBoxRight, "خطا در انجام محاسبه")
            Logger.LogFatal("Error Calculating Spring Rate", ex)
        End Try

    End Sub

    Private Sub Label7_Click(sender As Object, e As EventArgs) Handles Label7.Click
        ''Calculating Mandrel diameter

        Try
            TBMandrelDiameter.Text = TBOD.Text - TBWireDiameter.Text * 2
        Catch ex As Exception
            MsgBox("پارامتر های فنی محصول به درستی وارد نشده است", vbCritical + MsgBoxStyle.MsgBoxRight, "خطا در انجام محاسبه")
            Logger.LogFatal("Error Calculating mandrel diameter", ex)
        End Try
    End Sub

    Private Sub Label13_Click(sender As Object, e As EventArgs) Handles Label13.Click
        'Calculating wire length
        Try
            Dim noOfClosedEnds As String = 0
            If CBScoilType.Text = "" Or CBScoilType.Text = "بسته" Or CBScoilType.Text = "بسته و سنگ خورده" Or CBScoilType.Text = "بسته و فورج شده" Then
                noOfClosedEnds += 1
            End If

            If CBEcoilType.Text = "" Or CBEcoilType.Text = "بسته" Or CBEcoilType.Text = "بسته و سنگ خورده" Or CBEcoilType.Text = "بسته و فورج شده" Then
                noOfClosedEnds += 1
            End If

            Dim meanD, closedCoilLength, openCoilLength As Double
            meanD = Val(TBMandrelDiameter.Text) + Val(TBWireDiameter.Text)

            Dim groundThickness As Double = (1 - TBtipThickness.Text / 100) * 2

            closedCoilLength = ((Math.PI * meanD * ((TBNt.Text - TBNActive.Text) / 2)) ^ 2 + (((TBNt.Text - TBNActive.Text) / 2) * TBWireDiameter.Text) ^ 2) ^ 0.5
            openCoilLength = ((TBL0.Text + ((groundThickness - (TBNt.Text - TBNActive.Text) - 1)) * TBWireDiameter.Text) ^ 2 + (Math.PI * meanD * (TBNActive.Text)) ^ 2) ^ 0.5
            TBWireLength.Text = Math.Round(noOfClosedEnds * closedCoilLength + openCoilLength, 0)
        Catch ex As Exception
            MsgBox("پارامتر های فنی محصول به درستی وارد نشده است", vbCritical + MsgBoxStyle.MsgBoxRight, "خطا در انجام محاسبه")
            Logger.LogFatal("Error Calculating wire Length", ex)
        End Try


    End Sub

    Private Sub productForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Cursor = Cursors.WaitCursor
        Select Case productFormState
            Case "modify"
                BTNew.Enabled = False
                BTModify.Enabled = True
                BTDelete.Enabled = True
                PopulateForm()
            Case "new"
                BTNew.Enabled = True
                BTModify.Enabled = False
                BTDelete.Enabled = False
            Case "view"
                BTNew.Enabled = False
                BTModify.Enabled = False
                BTDelete.Enabled = False
                PopulateForm()
        End Select
        HandleUserPermissions()
        Me.Cursor = Cursors.Default
    End Sub
    Private Sub productForm_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        On Error Resume Next
        FrmNewEmkansanji.BTSearch.PerformClick()
    End Sub

    Private Sub Label27_Click(sender As Object, e As EventArgs) Handles Label27.Click
        '''' Calculating Spring solid Load
        Try
            Dim springRate, G, solidLength, meanD, activeCoil As Double
            meanD = TBOD.Text - TBWireDiameter.Text
            activeCoil = TBNActive.Text
            If CBMaterial.Text = "54SiCr6" Then
                G = 79500.5
            Else
                G = 79299.5
            End If

            springRate = (G * TBWireDiameter.Text ^ 4) / (8 * activeCoil * meanD ^ 3)
            Dim groundThickness As Double = (1 - TBtipThickness.Text / 100) * 2
            solidLength = (TBNt.Text + 1 - groundThickness) * TBWireDiameter.Text   'TODO: Compensate for spring end types
            TBMaxLoad.Text = Math.Round((TBL0.Text - solidLength) * springRate, 2)
        Catch ex As Exception
            MsgBox("پارامتر های فنی محصول به درستی وارد نشده است", vbCritical + MsgBoxStyle.MsgBoxRight, "خطا در انجام محاسبه")
            Logger.LogFatal("Error Calculating Solid Load", ex)
        End Try
    End Sub





    ''----------------------- Functions ------------------------
    Private Async Sub PopulateForm()
        'Using cn As New OleDbConnection(connectionString)
        '    Using cmd As New OleDbCommand With {.Connection = cn}
        '        cmd.CommandText = "SELECT * FROM springDataBase WHERE springDataBase.ID = " & TBdbID.Text & ";"
        '        Dim dt As New DataTable With {.TableName = "springDataBase"}
        '        'Try
        '        cn.Open()
        '        Dim ds As New DataSet
        '        Dim springDBTable As New DataTable With {.TableName = "springDataBase"}
        '        ds.Tables.Add(springDBTable)
        '        ds.Load(cmd.ExecuteReader(), LoadOption.OverwriteChanges, springDBTable)
        '        cn.Close()

        '        Me.TBProductName.Text = ds.Tables("springDataBase").Rows(0)("productName").ToString()
        '        Me.TBProductID.Text = ds.Tables("springDataBase").Rows(0)("productID").ToString()
        '        Me.CBProductionMethod.Text = ds.Tables("springDataBase").Rows(0)("productionMethod").ToString()
        '        Me.TBWireDiameter.Text = ds.Tables("springDataBase").Rows(0)("wireDiameter").ToString()
        '        Me.TBOD.Text = ds.Tables("springDataBase").Rows(0)("OD").ToString()
        '        Me.TBL0.Text = ds.Tables("springDataBase").Rows(0)("L0").ToString()
        '        Me.TBNt.Text = ds.Tables("springDataBase").Rows(0)("Nt").ToString()
        '        Me.TBNActive.Text = ds.Tables("springDataBase").Rows(0)("Nactive").ToString()

        '        ' TODO TODO TODO Shomare naghseh

        '        Me.CBCoilingDirection.Text = ds.Tables("springDataBase").Rows(0)("coilingDirection").ToString()
        '        Me.CBScoilType.Text = ds.Tables("springDataBase").Rows(0)("startCoilType").ToString()
        '        Me.CBEcoilType.Text = ds.Tables("springDataBase").Rows(0)("endCoilType").ToString()
        '        Me.TBtipThickness.Text = ds.Tables("springDataBase").Rows(0)("tipThickness").ToString()

        '        Me.TBMandrelDiameter.Text = ds.Tables("springDataBase").Rows(0)("mandrelDiameter").ToString()
        '        Me.TBWireLength.Text = ds.Tables("springDataBase").Rows(0)("wireLength").ToString()
        '        Me.TBSpringRate.Text = ds.Tables("springDataBase").Rows(0)("springRate").ToString()
        '        Me.CBMaterial.Text = ds.Tables("springDataBase").Rows(0)("material").ToString()

        '        Me.TBF1.Text = ds.Tables("springDataBase").Rows(0)("F1").ToString()
        '        Me.TBF2.Text = ds.Tables("springDataBase").Rows(0)("F2").ToString()
        '        Me.TBF3.Text = ds.Tables("springDataBase").Rows(0)("F3").ToString()

        '        Me.TBL1.Text = ds.Tables("springDataBase").Rows(0)("L1").ToString()
        '        Me.TBL2.Text = ds.Tables("springDataBase").Rows(0)("L2").ToString()
        '        Me.TBL3.Text = ds.Tables("springDataBase").Rows(0)("L3").ToString()

        '        Me.TBSolidStress.Text = ds.Tables("springDataBase").Rows(0)("solidStress").ToString()
        '        Me.TBMaxLoad.Text = ds.Tables("springDataBase").Rows(0)("solidLoad").ToString()
        '        Me.TBDwgNo.Text = ds.Tables("springDataBase").Rows(0)("dwgNo").ToString()
        '        Me.CBspringType.Text = ds.Tables("springDataBase").Rows(0)("pType").ToString()


        '        Me.TBComment.Text = ds.Tables("springDataBase").Rows(0)("comment").ToString()



        '    End Using
        'End Using

        Dim sql_command = "SELECT * FROM springDataBase WHERE springDataBase.ID = " & TBdbID.Text & ";"
        Try
            Dim dt = Await Task(Of DataTable).Run(Function() LoadDataTable(sql_command))
            Me.TBProductName.Text = dt.Rows(0)("productName").ToString()
            Me.TBProductID.Text = dt.Rows(0)("productID").ToString()
            Me.CBProductionMethod.Text = dt.Rows(0)("productionMethod").ToString()
            Me.TBWireDiameter.Text = dt.Rows(0)("wireDiameter").ToString()
            Me.TBOD.Text = dt.Rows(0)("OD").ToString()
            Me.TBL0.Text = dt.Rows(0)("L0").ToString()
            Me.TBNt.Text = dt.Rows(0)("Nt").ToString()
            Me.TBNActive.Text = dt.Rows(0)("Nactive").ToString()

            Me.CBCoilingDirection.Text = dt.Rows(0)("coilingDirection").ToString()
            Me.CBScoilType.Text = dt.Rows(0)("startCoilType").ToString()
            Me.CBEcoilType.Text = dt.Rows(0)("endCoilType").ToString()
            Me.TBtipThickness.Text = dt.Rows(0)("tipThickness").ToString()

            Me.TBMandrelDiameter.Text = dt.Rows(0)("mandrelDiameter").ToString()
            Me.TBWireLength.Text = dt.Rows(0)("wireLength").ToString()
            Me.TBSpringRate.Text = dt.Rows(0)("springRate").ToString()
            Me.CBMaterial.Text = dt.Rows(0)("material").ToString()

            Me.TBF1.Text = dt.Rows(0)("F1").ToString()
            Me.TBF2.Text = dt.Rows(0)("F2").ToString()
            Me.TBF3.Text = dt.Rows(0)("F3").ToString()

            Me.TBL1.Text = dt.Rows(0)("L1").ToString()
            Me.TBL2.Text = dt.Rows(0)("L2").ToString()
            Me.TBL3.Text = dt.Rows(0)("L3").ToString()

            Me.CBForceUnit.Text = dt.Rows(0)("forceUnit").ToString()
            Me.TBSolidStress.Text = dt.Rows(0)("solidStress").ToString()
            Me.TBMaxLoad.Text = dt.Rows(0)("solidLoad").ToString()
            Me.TBDwgNo.Text = dt.Rows(0)("dwgNo").ToString()
            Me.CBspringType.Text = dt.Rows(0)("pType").ToString()

            Me.TBComment.Text = dt.Rows(0)("comment").ToString()
        Catch ex As Exception
            MsgBox("خطا در بارگزاری اطلاعات محصول", vbCritical + RightToLeft + vbMsgBoxRight, "خطا")
            Logger.LogFatal(sql_command, ex)
        End Try


    End Sub

    Function CalculateLength(springRate As Double, L0 As Double, load As Double)
        Return Math.Round(L0 - (load / springRate), 2)
    End Function
    Function CalculateLoad(springRate As Double, L0 As Double, length As Double)
        Return Math.Round((L0 - length) * springRate, 2)
    End Function

    Private Sub Label18_Click(sender As Object, e As EventArgs) Handles Label18.Click
        Try
            TBL1.Text = CalculateLength(TBSpringRate.Text, TBL0.Text, TBF1.Text).ToString
        Catch ex As Exception
            MsgBox("پارامتر های فنی محصول به درستی وارد نشده است", vbCritical + MsgBoxStyle.MsgBoxRight, "خطا در انجام محاسبه")
            Logger.LogFatal("Error Calculating load length data", ex)
        End Try
    End Sub
    Private Sub Label19_Click(sender As Object, e As EventArgs) Handles Label19.Click
        Try
            TBL2.Text = CalculateLength(TBSpringRate.Text, TBL0.Text, TBF2.Text).ToString
        Catch ex As Exception
            MsgBox("پارامتر های فنی محصول به درستی وارد نشده است", vbCritical + MsgBoxStyle.MsgBoxRight, "خطا در انجام محاسبه")
            Logger.LogFatal("Error Calculating load length data", ex)
        End Try
    End Sub
    Private Sub Label20_Click(sender As Object, e As EventArgs) Handles Label20.Click
        Try
            TBL3.Text = CalculateLength(TBSpringRate.Text, TBL0.Text, TBF3.Text).ToString
        Catch ex As Exception
            MsgBox("پارامتر های فنی محصول به درستی وارد نشده است", vbCritical + MsgBoxStyle.MsgBoxRight, "خطا در انجام محاسبه")
            Logger.LogFatal("Error Calculating load length data", ex)
        End Try
    End Sub

    Private Sub Label15_Click(sender As Object, e As EventArgs) Handles Label15.Click
        Try
            TBF1.Text = CalculateLength(TBSpringRate.Text, TBL0.Text, TBL1.Text).ToString
        Catch ex As Exception
            MsgBox("پارامتر های فنی محصول به درستی وارد نشده است", vbCritical + MsgBoxStyle.MsgBoxRight, "خطا در انجام محاسبه")
            Logger.LogFatal("Error Calculating load length data", ex)
        End Try
    End Sub

    Private Sub Label16_Click(sender As Object, e As EventArgs) Handles Label16.Click
        Try
            TBF2.Text = CalculateLength(TBSpringRate.Text, TBL0.Text, TBL2.Text).ToString
        Catch ex As Exception
            MsgBox("پارامتر های فنی محصول به درستی وارد نشده است", vbCritical + MsgBoxStyle.MsgBoxRight, "خطا در انجام محاسبه")
            Logger.LogFatal("Error Calculating load length data", ex)
        End Try
    End Sub

    Private Sub Label17_Click(sender As Object, e As EventArgs) Handles Label17.Click
        Try
            TBF3.Text = CalculateLength(TBSpringRate.Text, TBL0.Text, TBL3.Text).ToString
        Catch ex As Exception
            MsgBox("پارامتر های فنی محصول به درستی وارد نشده است", vbCritical + MsgBoxStyle.MsgBoxRight, "خطا در انجام محاسبه")
            Logger.LogFatal("Error Calculating load length data", ex)
        End Try
    End Sub

    Private Sub BTCheckMandrelInventory_Click(sender As Object, e As EventArgs) Handles BTCheckMandrelInventory.Click
        mandrels.Show()
        mandrels.SearchMandrelDataBase(TBMandrelDiameter.Text)
    End Sub
    Private Sub HandleUserPermissions()
        If loggedInUserGroup <> "Admin" And loggedInUserGroup <> "QC" Then
            BTModify.Enabled = False
            BTNew.Enabled = False
            BTDelete.Enabled = False
        End If
        If loggedInUserGroup = "Anbar" Then
            '' disable all textboxes
            For Each tb As TextBox In GroupBox1.Controls.OfType(Of TextBox)()
                tb.ReadOnly = True
            Next
            For Each cb As ComboBox In GroupBox1.Controls.OfType(Of ComboBox)()
                cb.Enabled = False
            Next
            BTDelete.Enabled = False
            BTNew.Enabled = False
            BTCheckMandrelInventory.Enabled = False
            BTWireInventory.Enabled = False
            TBProductID.ReadOnly = False '' need to be able to modify product ID 
            BTModify.Enabled = True
        End If
    End Sub

    Private Sub Label8_Click(sender As Object, e As EventArgs) Handles Label8.Click
        Try
            TBNActive.Text = Val(TBNt.Text) - 1.5
        Catch ex As Exception
            MsgBox("پارامتر های فنی محصول به درستی وارد نشده است", vbCritical + MsgBoxStyle.MsgBoxRight, "خطا در انجام محاسبه")
            Logger.LogFatal("Error Calculating Active Coil", ex)
        End Try
    End Sub

    Private Sub TBWireLength_TextChanged(sender As Object, e As EventArgs) Handles TBWireLength.TextChanged
        If IsNumeric(TBWireLength.Text) And CBProductionMethod.Text <> "سرد پیچ" Then
            If Val(TBWireLength.Text) > My.Settings.wireLengthThreshold Then
                TBWireLength.BackColor = Color.Red
                TBWireLength.ForeColor = SystemColors.Window
            Else
                TBWireLength.BackColor = SystemColors.Window
                TBWireLength.ForeColor = SystemColors.WindowText
            End If
        Else
            TBWireLength.BackColor = SystemColors.Window
            TBWireLength.ForeColor = SystemColors.WindowText
        End If

    End Sub

    Private Sub TBSolidStress_TextChanged(sender As Object, e As EventArgs) Handles TBSolidStress.TextChanged
        If IsNumeric(TBSolidStress.Text) Then
            If Val(TBSolidStress.Text) > My.Settings.stressThreashold Then
                TBSolidStress.BackColor = Color.Red
                TBSolidStress.ForeColor = SystemColors.Window
            Else
                TBSolidStress.BackColor = SystemColors.Window
                TBSolidStress.ForeColor = SystemColors.WindowText
            End If
        Else
            TBSolidStress.BackColor = SystemColors.Window
            TBSolidStress.ForeColor = SystemColors.WindowText
        End If
    End Sub
End Class