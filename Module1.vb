Imports System.Configuration
Imports Scripting
Imports System.Text.RegularExpressions



Public Module GlobalVariables

    Public debugMode As Boolean = True

    '' -------------------------------------------------------  Configurations  -----------------------------------------------------

    'Public connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\ESDB.accdb"
    Public connectionString As String = ConfigurationManager.ConnectionStrings("ConString").ConnectionString & "Jet OLEDB:Database Password=esdbpassword;"

    'Public excelTemplateFilePath As String = "D:\ES.xlsx"
    Public excelTemplateFilePath As String = ConfigurationManager.ConnectionStrings("excelPath").ConnectionString
    Public excelFilesBasePath As String = ConfigurationManager.ConnectionStrings("excelBasePath").ConnectionString

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
     " emkansanji.dateOfProccessing AS [تاریخ بررسی], emkansanji.standard AS [استاندارد], emkansanji.grade AS [گرید], emkansanji.productCode AS [کد قطعه مشتری], " &
     " emkansanji.r1_code AS [کد مفتول رزرو 1], emkansanji.r1_q AS [مقدار1], emkansanji.r2_code AS [کد مفتول رزرو 2], emkansanji.r2_q AS [مقدار 2], emkansanji.r3_code AS [کد مفتول رزرو 3], emkansanji.r3_q AS [مقدار 3], emkansanji.wireState AS [وضعیت موجودی مفتول], " &
     " emkansanji.verificationNo AS [شماره تاییدیه], emkansanji.verificationDate AS [تاریخ تاییدیه], emkansanji.comment AS [توضیحات] "

    '' ----------------------------------------------------  Form State Variables  ----------------------------------------------------
    Public productFormState As String = "modify"
    Public customerFormState As String = "modify"


End Module

Public Module globalFunctions
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
'       [  ] Do something for tolid 
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
'               [  ] 
'       [✔] Add functionality of making emkansanji excel file
'       [  ] State of wire and mandrel and packaging should be available in the emkansanji database
'       [  ] Add product reservation 
'       [  ] The Excel file should be opened from inside the program
'       [✔] Add Production Method to springDataBase
'       [  ] Mandrel DataBase
'               [  ] Check if mandrel is present by clicking on mandrel Label
'       [✔] Add an aproppriate Icon
'       [✔] Prevent deleting of a product or customer for which emkansanji exists
'       [✔] Prevent using of unwanted characters in the file name
'       [✔] Prevent overwriting previous excel files with the same name
'       [✔] Password Protect The DataBase
'       [  ] Try a better database server (Preferably PostgreSQL)
'       [  ]
'       [  ]
'       [  ]
'
'
