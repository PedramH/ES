Imports System.Threading
Public Class UserSettingForm
    Private Sub UserSettingForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TBESDuplicateBasePath.Text = My.Settings.duplicateESExcel
    End Sub

    Private Sub BTSave_Click(sender As Object, e As EventArgs) Handles BTSave.Click
        My.Settings.duplicateESExcel = TBESDuplicateBasePath.Text
        MsgBox("تنظیمات با موفقیت ذخیره شد", vbInformation + vbMsgBoxRtlReading, "ذخیره تنظیمات")
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click
        '' Check to see if the filepath provided in the config file exist, if not ask for the path
        '' This portion of the code uses a seprate thread with STA, because winforms can't open a openfilediaglog()
        ''    in the same thread as the form! For whatever fucked up reason.

        '' TODO: there is some bug here! :-?
        'Dim filePath = ""
        'Dim t As New Thread(
        '    Sub()
        '        'Dim fd As OpenFileDialog = New OpenFileDialog()
        '        Dim fd As OpenFileDialog = New OpenFileDialog()
        '        fd.Title = "Open File Dialog"
        '        fd.InitialDirectory = ""
        '        'fd.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm"
        '        'fd.FilterIndex = 2
        '        fd.RestoreDirectory = True
        '        If fd.ShowDialog() = DialogResult.OK Then
        '            filePath = fd.FileName
        '        ElseIf fd.ShowDialog() = DialogResult.Cancel Then
        '            Exit Sub
        '        End If

        '    End Sub
        ')

        ''' Run the code from a thread that joins the STA Thread
        't.SetApartmentState(ApartmentState.STA)
        't.Start()
        't.Join()
        'If filePath <> "" Then TBESDuplicateBasePath.Text = filePath
    End Sub
End Class