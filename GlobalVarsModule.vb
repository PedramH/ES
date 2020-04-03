Imports System.Configuration

Module GlobalVarsModule

    '' ------------------------------------------ Debug Mode: circumvents the login system ------------------------------------------
    Public debugMode As Boolean = False
    Public db As String = "postgres" 'access or postgres

    '' -------------------------------------------------------  Configurations  -----------------------------------------------------


    Public connectionString As String = ConfigurationManager.ConnectionStrings("ConString").ConnectionString & "Jet OLEDB:Database Password=esdbpassword;"
    'Public postgresConString As String = "Host=198.143.181.131;Port=5432;Username=postgres;Password=picher;Database=mydb"
    Public postgresConString As String = "Host=185.97.117.81;Port=5432;Username=postgres;Password=esdbpassword;Database=esdb" 'arvan-cloud


    Public excelTemplateFilePath As String = ConfigurationManager.ConnectionStrings("excelPath").ConnectionString
    Public excelFilesBasePath As String = ConfigurationManager.ConnectionStrings("excelBasePath").ConnectionString

    '' TODO: Make these configurable
    Public excelInventoryGarmPath As String = "D:\Academic\EnergySaz\Emkansanji\takeHome\موجودي مواد\garm.xlsx"
    Public excelInventorySardPath As String = "D:\Academic\EnergySaz\Emkansanji\takeHome\موجودي مواد\sard.xlsx"
    Public excelInventoryPurchasedPath As String = "D:\Academic\EnergySaz\Emkansanji\takeHome\موجودي مواد\purchased.xlsx"


    '' ----------------------------------------------------------  User Info  --------------------------------------------------------

    Public loggedInUser As String = ""
    Public loggedInUserName As String = ""
    Public loggedInUserGroup As String = ""

    '' -------------------------------------------------------  Query Column Names  --------------------------------------------------

    Public springDataBaseColumnNames As String = " ID AS [شماره شناسایی], productName AS [نام محصول], productID AS [کد کالا], pType AS [نوع فنر] ,productionMethod AS [روش تولید] ,wireDiameter AS [قطر مفتول], " &
     " OD AS [قطر خارجی], L0 AS [طول آزاد], Nt AS [حلقه کل], Nactive AS [حلقه فعال], coilingDirection AS [جهت پیچش], " &
     " mandrelDiameter AS [قطر شفت], wireLength AS [طول مفتول], startCoilType AS [شکل حلقه ابتدا], endCoilType AS [شکل حلقه انتها], tipThickness AS [ضخامت لبه] ,material AS [جنس مواد], " &
     " solidStress AS [تنش حداکثر], solidLoad AS [حداکثر نیرو] , springRate AS [ریت فنر] ,F1,L1,F2,L2,F3,L3, forceUnit AS [واحد نیرو] ,comment AS [توضیحات], productionProcess AS [productionProcess] "

    Public customerDataBaseColumnNames As String = " ID AS [شماره شناسایی مشتری], customerName AS [نام مشتری], fieldOfWork AS [زمینه کاری], shenaseMelli AS [شناسه ملی], " &
     " codeEghtesadi AS [کد اقتصادی], postCode AS [کد پستی], ads1 AS [آدرس 1], ads2 AS [آدرس 2], p1 AS [رابط اول], " &
     " p1_job AS [سمت], p1_phone AS [تلفن], p1_mobile AS [تلفن همراه], p1_email AS [ایمیل], p2 AS [رابط دوم], " &
     " p2_job AS [سمت 2], p2_phone AS [تلفن 2], p2_mobile AS [تلفن همراه 2], p2_email AS [ایمیل 2], p3 AS [رابط سوم], " &
     " p3_job AS [سمت 3], p3_phone AS [تلفن 3], p3_mobile AS [تلفن همراه 3], p3_email AS [ایمیل 3], " &
     " requirements AS [الزامات مشتری],  comment AS [توضیحات] "

    Public ESColumnNames As String = " springDataBase.ID AS [productID] , customers.ID AS [customerID] , springDataBase.wireDiameter AS [wireDiameter], springDataBase.OD AS [OD] , springDataBase.L0 AS [L0] , springDataBase.wireLength AS [wireLength], springDataBase.mandrelDiameter AS [mandrelDiameter], emkansanji.ID AS [شماره ردیابی سفارش], springDataBase.productName AS [نام محصول], customers.customerName AS [نام مشتری], emkansanji.customerProductName AS [نام محصول مشتری], emkansanji.orderState AS [وضعیت سفارش], " &
     " emkansanji.customerDwgNo AS [شماره نقشه], emkansanji.quantity AS [تعداد سفارش],emkansanji.sampleQuantity AS [تعداد نمونه] ,emkansanji.letterNo AS [شماره نامه], emkansanji.letterDate AS [تاریخ نامه], emkansanji.orderNo AS [شماره سفارش], " &
     " emkansanji.dateOfProccessing AS [تاریخ بررسی], emkansanji.standard AS [استاندارد], emkansanji.grade AS [گرید], emkansanji.productCode AS [کد قطعه مشتری], emkansanji.mandrelState AS [موجودی مندرل], " &
     " emkansanji.r1_code AS [کد مفتول رزرو 1], emkansanji.r1_q AS [مقدار1], emkansanji.r2_code AS [کد مفتول رزرو 2], emkansanji.r2_q AS [مقدار 2], emkansanji.r3_code AS [کد مفتول رزرو 3], emkansanji.r3_q AS [مقدار 3], emkansanji.wireState AS [وضعیت موجودی مفتول], emkansanji.productReserve, " &
     " emkansanji.verificationNo AS [شماره تاییدیه], emkansanji.verificationDate AS [تاریخ تاییدیه], emkansanji.comment AS [توضیحات], emkansanji.pProcess AS [pProcess] , emkansanji.productReserve AS [productionReserve] ," &
     "springDataBase.productionProcess AS [productionProcess] , emkansanji.springInEachPackage AS [springInEachPackage] , emkansanji.packagingCost AS [packagingCost] , emkansanji.doable AS [doable] , emkansanji.whyNot AS [whyNot], " &
     " emkansanji.buyWire AS [buyWire], emkansanji.buyMandrel AS [buyMandrel] , emkansanji.zarfiatSanji AS [zarfiatSanji], emkansanji.packageType AS [packageType], emkansanji.inspectionProcess AS [inspectionProcess] , emkansanji.orderType AS [orderType] , emkansanji.excelFilePath AS [excelFilePath] "

    Public mandrelsColumnName As String = " ID AS [شماره شناسایی], mandrelCode AS [کد کالا] , mandrelDiameter AS [قطر شفت] "

    '' A is wireInventory table as B is wire reserves table
    'Public wiresColumnName As String = " A.wireType, A.wireWeight, A.wireCode AS [کد کالا], A.inventoryName AS [عنوان] , A.wireDiameter AS [قطر مفتول], A.wireLength AS [طول مفتول] ,
    '                                    (A.inventory - B.preReserve - B.reserve) AS [مانده موجودی] , IIF(ISNUMERIC (A.wireWeight), INT((A.inventory - B.preReserve - B.reserve) / A.wireWeight) , '-' ) AS [تعداد شاخه] , 
    '                                     A.inventory AS [موجودی فیزیکی], IIF( ISNUMERIC(A.wireWeight), INT(A.inventory / A.wireWeight) , '-' ) AS [موجودی فیزیکی (تعداد شاخه)], 
    '                                     B.preReserve AS [رزرو امکان سنجی] , IIF( ISNUMERIC(A.wireWeight), INT(B.preReserve / A.wireWeight) , '-' ) AS [امکان سنجی (تعداد شاخه)], 
    '                                     B.reserve AS [رزرو تولید] , IIF( ISNUMERIC(A.wireWeight), INT(B.reserve / A.wireWeight) , '-' ) AS [تولید(تعداد شاخه)]   "
    Public wiresColumnName As String = " A.wireType, A.wireWeight, A.wireCode AS [کد کالا], A.inventoryName AS [عنوان] , A.wireDiameter AS [قطر مفتول], A.wireLength AS [طول مفتول] ,
                                        INT((A.inventory - B.preReserve - B.reserve)) AS [مانده موجودی (کیلوگرم)] , IIF(ISNUMERIC (A.wireWeight), INT((A.inventory - B.preReserve - B.reserve) / A.wireWeight) , '-' ) AS [تعداد شاخه] , 
                                         A.inventory AS [موجودی فیزیکی(کیلوگرم)], IIF( ISNUMERIC(A.wireWeight), INT(A.inventory / A.wireWeight) , '-' ) AS [موجودی فیزیکی (تعداد شاخه)], 
                                         B.preReserve AS [رزرو امکان سنجی (کیلوگرم)] , IIF( ISNUMERIC(A.wireWeight), INT(B.preReserve / A.wireWeight) , '-' ) AS [امکان سنجی (تعداد شاخه)], 
                                         B.reserve AS [رزرو تولید (کیلوگرم)] , IIF( ISNUMERIC(A.wireWeight), INT(B.reserve / A.wireWeight) , '-' ) AS [تولید(تعداد شاخه)]   "



    '' ----------------------------------------------------  Form State Variables  ----------------------------------------------------
    Public productFormState As String = "modify"
    Public customerFormState As String = "modify"


    '' this variable determines the the state of select wire button (enable or disable)
    Public wiresFormState As String = "normal"  ' normal - selection 
    '' wireFormCaller is used to determine which of the wire reservastion spots called the wire form. 
    Public wireFormCaller As String = ""
    '' ----------------------------------------------------  Order state numbering system  ----------------------------------------------------
    Public userGroupOrder As New Dictionary(Of String, String)
End Module
