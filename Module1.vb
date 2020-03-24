Imports System.Configuration
Imports System.Threading
Imports Scripting
Imports System.Text.RegularExpressions
Imports System.Data.OleDb



Public Module GlobalVariables

    Public debugMode As Boolean = True

    '' -------------------------------------------------------  Configurations  -----------------------------------------------------

    'Public connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\ESDB.accdb"
    Public connectionString As String = ConfigurationManager.ConnectionStrings("ConString").ConnectionString & "Jet OLEDB:Database Password=esdbpassword;"

    'Public excelTemplateFilePath As String = "D:\ES.xlsx"
    Public excelTemplateFilePath As String = ConfigurationManager.ConnectionStrings("excelPath").ConnectionString
    Public excelFilesBasePath As String = ConfigurationManager.ConnectionStrings("excelBasePath").ConnectionString
    '' TODO: Make these configurable
    Public excelInventoryGarmPath As String = "D:\Academic\EnergySaz\Emkansanji\takeHome\موجودي مواد\garm.xlsx"
    Public excelInventorySardPath As String = "D:\Academic\EnergySaz\Emkansanji\takeHome\موجودي مواد\sard.xlsx"
    Public excelInventoryPurchasedPath As String = "D:\Academic\EnergySaz\Emkansanji\takeHome\موجودي مواد\purchased.xlsx"


    '' ----------------------------------------------------------  User Info  --------------------------------------------------------

    Public loggedInUser As String = "Pedram"
    Public loggedInUserName As String = "پدرام یوسفی"
    Public loggedInUserGroup As String = "Admin"

    '' -------------------------------------------------------  Query Column Names  --------------------------------------------------
    Public springDataBaseColumnNames As String = " ID AS [شماره شناسایی], productName AS [نام محصول], productID AS [کد کالا], productionMethod AS [روش تولید] ,wireDiameter AS [قطر مفتول], " &
     " OD AS [قطر خارجی], L0 AS [طول آزاد], Nt AS [حلقه کل], Nactive AS [حلقه فعال], coilingDirection AS [جهت پیچش], " &
     " mandrelDiameter AS [قطر شفت], wireLength AS [طول مفتول], startCoilType AS [شکل حلقه ابتدا], endCoilType AS [شکل حلقه انتها], tipThickness AS [ضخامت لبه] ,material AS [جنس مواد], " &
     " solidStress AS [تنش حداکثر], solidLoad AS [حداکثر نیرو] , springRate AS [ریت فنر] ,F1,L1,F2,L2,F3,L3,  comment AS [توضیحات] "


    Public customerDataBaseColumnNames As String = " ID AS [شماره شناسایی مشتری], customerName AS [نام مشتری], fieldOfWork AS [زمینه کاری], shenaseMelli AS [شناسه ملی], " &
     " codeEghtesadi AS [کد اقتصادی], postCode AS [کد پستی], ads1 AS [آدرس 1], ads2 AS [آدرس 2], p1 AS [رابط اول], " &
     " p1_job AS [سمت], p1_phone AS [تلفن], p1_mobile AS [تلفن همراه], p1_email AS [ایمیل], p2 AS [رابط دوم], " &
     " p2_job AS [سمت 2], p2_phone AS [تلفن 2], p2_mobile AS [تلفن همراه 2], p2_email AS [ایمیل 2], p3 AS [رابط سوم], " &
     " p3_job AS [سمت 3], p3_phone AS [تلفن 3], p3_mobile AS [تلفن همراه 3], p3_email AS [ایمیل 3], " &
     " requirements AS [الزامات مشتری],  comment AS [توضیحات] "

    Public ESColumnNames As String = " springDataBase.ID, customers.ID, springDataBase.wireDiameter, springDataBase.OD, springDataBase.L0, springDataBase.wireLength, springDataBase.mandrelDiameter , emkansanji.ID AS [شماره ردیابی سفارش], springDataBase.productName AS [نام محصول], customers.customerName AS [نام مشتری], emkansanji.customerProductName AS [نام محصول مشتری], emkansanji.orderState AS [وضعیت سفارش], " &
     " emkansanji.customerDwgNo AS [شماره نقشه], emkansanji.quantity AS [تعداد سفارش], emkansanji.letterNo AS [شماره نامه], emkansanji.letterDate AS [تاریخ نامه], emkansanji.orderNo AS [شماره سفارش], " &
     " emkansanji.dateOfProccessing AS [تاریخ بررسی], emkansanji.standard AS [استاندارد], emkansanji.grade AS [گرید], emkansanji.productCode AS [کد قطعه مشتری], emkansanji.mandrelState AS [موجودی مندرل], " &
     " emkansanji.r1_code AS [کد مفتول رزرو 1], emkansanji.r1_q AS [مقدار1], emkansanji.r2_code AS [کد مفتول رزرو 2], emkansanji.r2_q AS [مقدار 2], emkansanji.r3_code AS [کد مفتول رزرو 3], emkansanji.r3_q AS [مقدار 3], emkansanji.wireState AS [وضعیت موجودی مفتول], " &
     " emkansanji.verificationNo AS [شماره تاییدیه], emkansanji.verificationDate AS [تاریخ تاییدیه], emkansanji.comment AS [توضیحات] "

    Public mandrelsColumnName As String = " ID AS [شماره شناسایی], mandrelCode AS [کد کالا] , mandrelDiameter AS [قطر شفت] "

    '' A is wireInventory table as B is wire reserves table
    Public wiresColumnName As String = " A.wireType, A.wireWeight, A.wireCode AS [کد کالا], A.inventoryName AS [عنوان] , A.wireDiameter AS [قطر مفتول], A.wireLength AS [طول مفتول] ,
                                        (A.inventory - B.preReserve - B.reserve) AS [مانده موجودی] , IIF(ISNUMERIC (A.wireWeight), INT((A.inventory - B.preReserve - B.reserve) / A.wireWeight) , '-' ) AS [تعداد شاخه] , 
                                         A.inventory AS [موجودی فیزیکی], IIF( ISNUMERIC(A.wireWeight), INT(A.inventory / A.wireWeight) , '-' ) AS [موجودی فیزیکی (تعداد شاخه)], 
                                         B.preReserve AS [رزرو امکان سنجی] , IIF( ISNUMERIC(A.wireWeight), INT(B.preReserve / A.wireWeight) , '-' ) AS [امکان سنجی (تعداد شاخه)], 
                                         B.reserve AS [رزرو تولید] , IIF( ISNUMERIC(A.wireWeight), INT(B.reserve / A.wireWeight) , '-' ) AS [تولید(تعداد شاخه)]   "



    '' ----------------------------------------------------  Form State Variables  ----------------------------------------------------
    Public productFormState As String = "modify"
    Public customerFormState As String = "modify"

    Public wiresFormState As String = "normal"  ' normal - selection 
    Public wireFormCaller As String = ""

End Module

Public Module globalFunctions

    Public Function LoadDataTable(sql As String) As DataTable
        '' This function gets an SQL Query then returns data in a datatable
        Dim dt = New DataTable
        Using dbcon As New OleDbConnection(connectionString)
            Using cmd As New OleDbCommand(sql, dbcon)
                dbcon.Open()
                dt.Load(cmd.ExecuteReader())
                dbcon.Close()
            End Using
        End Using
        Return dt
    End Function


    Public Function ImportExceltoDatatable(filePath As String, fileDesc As String) As DataTable
        '' This function import data from an excel file and return the data in a datatable


        '' Check to see if the filepath provided in the config file exist, if not ask for the path
        '' This portion of the code uses a seprate thread with STA, because winforms can't open a openfilediaglog()
        ''    in the same thread as the form! For whatever fucked up reason.

        '' TODO: there is some bug here! :-?
        Dim t As New Thread(
            Sub()
                While (System.IO.File.Exists(filePath) = False)
                    MsgBox(String.Format("فایل {0} یافت نشد. لطفا این فایل را انتخاب کنید.", fileDesc), MsgBoxStyle.Critical + MsgBoxStyle.MsgBoxRtlReading + vbMsgBoxRight, "خطا")
                    Dim fd As OpenFileDialog = New OpenFileDialog()
                    fd.Title = "Open File Dialog"
                    fd.InitialDirectory = "C:\"
                    fd.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm"
                    fd.FilterIndex = 2
                    fd.RestoreDirectory = True
                    If fd.ShowDialog() = DialogResult.OK Then
                        filePath = fd.FileName
                    ElseIf fd.ShowDialog() = DialogResult.Cancel Then
                        MsgBox(String.Format("عملیات خواندن اطلاعات از فایل {0} به انتخاب کاربر لغو شد.", fileDesc), MsgBoxStyle.Critical + MsgBoxStyle.MsgBoxRtlReading + vbMsgBoxRight, "خطا")
                        Exit Sub
                    End If
                End While
            End Sub
        )

        '' Run the code from a thread that joins the STA Thread
        t.SetApartmentState(ApartmentState.STA)
        t.Start()
        t.Join()

        Dim dt As New DataTable
                Try
                    Dim ds As New DataSet()
                    Dim constring As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filePath & ";Extended Properties=""Excel 12.0;HDR=YES;"""
                    Dim con As New OleDbConnection(constring & "")
                    con.Open()
                    Dim myTableName = con.GetSchema("Tables").Rows(0)("TABLE_NAME")
                    Dim sqlquery As String = String.Format("SELECT * FROM [{0}]", myTableName) ' "Select * From " & myTableName  
                    Dim da As New OleDbDataAdapter(sqlquery, con)
                    da.Fill(ds)
                    dt = ds.Tables(0)
                    con.Close()
                    Return dt
                Catch ex As Exception
                    MsgBox(String.Format("خطا در خواندن اطلاعات از فایل {0}. فایل را بررسی کنید و مجددا سعی کنید.", fileDesc), MsgBoxStyle.Critical + MsgBoxStyle.MsgBoxRtlReading + vbMsgBoxRight, "خطا")
                    Logger.LogFatal(ex.Message, ex)
                    Return dt
                End Try
    End Function
    Public Async Function UpdateReservesTable() As Task
        '' This function updates the wire reserves table based on emkansanji table
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
    Public Function CalculateWireWeight(d As Double, L As Double) As Double
        '' Calculates the weight of each wire rod
        Dim rho As Double = 0.00000783
        Return Math.Round(((d * d * Math.PI) / 4) * L * rho, 2)
    End Function
    Public Function getDate()
        Dim pc As New Globalization.PersianCalendar
        Return pc.GetYear(Now).ToString & "-" & pc.GetMonth(Now).ToString & "-" & pc.GetDayOfMonth(Now).ToString
    End Function


    'requires reference to Microsoft Scripting Runtime
    Public Function MkDir(path As String)

        'Dim fso As New FileSystemObject
        'Dim path As String

        'If Not fso.FolderExists(path) Then

        ' doesn't exist, so create the folder
        'fso.CreateFolder(path)
        'End If
        Try
            System.IO.Directory.CreateDirectory(path)
        Catch ex As Exception
            MsgBox("خطا در ساخت مسیر ذخیره فایل. اجازه های دسترسی بررسی شود", vbCritical + vbMsgBoxRight, "خطا")
        End Try
        Return True
    End Function

    Public Function getMonthName(month As Integer)
        Select Case month
            Case 1
                Return "فروردین"
            Case 2
                Return "اردیبهشت"
            Case 3
                Return "خرداد"
            Case 4
                Return "تیر"
            Case 5
                Return "مرداد"
            Case 6
                Return "شهریور"
            Case 7
                Return "مهر"
            Case 8
                Return "آبان"
            Case 9
                Return "آذر"
            Case 10
                Return "دی"
            Case 11
                Return "بهمن"
            Case 12
                Return "اسفند"
            Case Else
                Return True
        End Select
    End Function

    Public Function stripFileName(fileName As String)
        fileName = Regex.Replace(fileName, "[*\\/]", ".")
        fileName = Regex.Replace(fileName, "[:|<>""]", "")
        Return fileName
    End Function

    Public Function NormalizeName(inputStr As String)
        'This function normalizes the search query and prevent SQL injection
        inputStr = Regex.Replace(inputStr, "[\\/'"";=]", "")
        inputStr = Regex.Replace(inputStr, "[ك]", "ک")
        inputStr = Regex.Replace(inputStr, "[ي]", "ی")
        Return inputStr
    End Function
    Public Function NormalizeString(inputStr As String)
        ' This funtion prevent certain characters (* / \ " ' ; = ) from being in the name inputs
        inputStr = Regex.Replace(inputStr, "[*\\/'"";=]", "")
        Return inputStr
    End Function


    Public Function preverntOverwriting(filePath As String, fileExtension As String)
        'This funtion checks to see if a file exist, if it does it adds a number in paranthesis line name(i) to the filename
        Dim i As Integer = 0
        If System.IO.File.Exists(filePath & fileExtension) Then
            'The file exists
            Do
                i += 1
            Loop While System.IO.File.Exists(filePath & "(" & i.ToString() & ")" & fileExtension)
            Return filePath & "(" & i.ToString() & ")" & fileExtension
        End If
        Return filePath & fileExtension
    End Function
End Module


'   TODO: 
'       [✔] Add IST like calculations to productForm 
'       [✔] Fix the tab order of form
'       [✔] Do something for tolid 
'       [  ] Deploy using a real database with correct information
'               [  ] Migrate Tolid's data to the new format
'       [✔] Add a config file
'       [  ] Fix the Functionality of Modify Emkansanji Form 
'       [  ] Create a product and customer search form 
'       [  ] Add usergroup and different user permissions
'               [✔] Implement a login system
'                       [  ] Implement a hashing system to store passwords - Just for the fun of it :) 
'               [  ] Enable/Disable Form controls based on usergroup
'               [  ] Restrict user's permission to modify different parts of the database 
'       [✔] Add a logging system
'       [  ] Error Handling and Logging
'               [✔] ProductForm
'               [✔] CustomerForm
'               [✔] Login and Change password Form
'               [✔] Main form
'               [✔] Module1
'               [  ] emkansanjiForm
'               [  ] wires
'       [✔] Add functionality of making emkansanji excel file
'       [  ] State of wire and mandrel and packaging should be available in the emkansanji database
'       [  ] Add product reservation 
'       [  ] The Excel file should be opened from inside the program
'       [✔] Add Production Method to springDataBase
'       [✔] Mandrel DataBase
'               [✔] Check if mandrel is present by clicking on a button
'       [✔] Add an aproppriate Icon
'       [✔] Prevent deleting of a product or customer for which emkansanji exists
'       [✔] Prevent using of unwanted characters in the file name
'       [✔] Prevent overwriting previous excel files with the same name
'       [✔] Password Protect The DataBase
'       [✔] Try a better database server (Preferably PostgreSQL) -> it works fine with minimal change
'       [  ] Migrate to postgreSQL
'       [  ] Make all calls to database async
'       [  ] Make a script to change the reserve for every emkansanji where a future bought wire is used when the wire arrives to the factory
'       [✔] Generate wire reservation table
'       [  ] Consider wire state in generation of reservation table
'       [✔] Use regex to extract wire Length
'       [  ] Add print functionality to wires data
'       [  ] Wires
'               [✔] Add searching functionality 
'               [✔] Add list of all orders
'               [✔] Test the new search system in wires form, if it's good enough maybe change everything else to this method? 
'               [  ] Add formating based on value to grid views
'               [  ] Properly Sort and format wires data 
'       [  ] Make a main form from which every form is accessible
'       [  ] Update reserves table using the function after each change to the emkansanji table
'       [  ] Update mandrel data from rahkaran excel file. it shouldn't be hard but it will save some headache in the future
'       [  ] Measure time difference between local and over the network queries  
'       [✔] Update data in the wireForm after changing an emkansanji datum originated from that form
'       [  ] Customer can get all their open orders states from telegram, providing that they know their specific customer ID(long string) 
'       [  ] Change logging mode from file to db
'       [  ] Make a copy of rahkaran excel files then read the copy to prevent file open in another computer errors. 
'       [  ] Add something to compensate wire reservation for orders that only a part of them is produced
'       [  ] Build purchased file with the data from mojodimaftol.excel
'       [  ] Use NormalizeName function to prevent sql injection and compensate for difference in farsi and arabic ی  ک characters
'       [  ]
'       [  ] 
'       [  ] 
'       [  ] 
'
'
