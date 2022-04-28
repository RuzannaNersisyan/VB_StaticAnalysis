'USEUNIT SWIFT_International_Payorder_Library
'USEUNIT Library_Common
'USEUNIT Library_Colour
'USEUNIT Library_CheckDB
'USEUNIT Constants
Option Explicit
Dim sDATE, fDATE, settingsPath, max, min, rand, fileFrom, fileTo, what, fWith
Sub SWIFT_300_Inernational_RightClick_Test()
    'Համակարգ մուտք գործել ARMSOFT օգտագործողով
    Call Initialize_AsBank("bank", sDATE, fDATE)
    Login("ARMSOFT")
'----------------------------------------------------
'--------------- Կարգավորումների ներմուծում --------------
'----------------------------------------------------
    Log.Message "Կարգավորումների ներմուծում ",,,DivideColor
    settingsPath = Project.Path & "Stores\SWIFT\HT300\Settings\Setting_2.txt"
    Call Settings_Import(settingsPath)
    Login("ARMSOFT")
'-----------------------------------------------------------------------------
'------ "S.W.I.F.T. ԱՇՏ/Պարամետրեր"-ում կատարել համապատասխան փոփոխությունները-------
'-----------------------------------------------------------------------------
    Log.Message "-- S.W.I.F.T. ԱՇՏ/Պարամետրեր-ում կատարել համապատասխան փոփոխությունները --",,,DivideColor  
    'Մուտք գործել "S.W.I.F.T. ԱՇՏ"
    Call ChangeWorkspace(c_SWIFT)
    Call wTreeView.DblClickItem("|S.W.I.F.T. ²Þî                  |ä³ñ³Ù»ïñ»ñ")
    BuiltIn.Delay(3000)
    'Նոր փաստաթղթի համարի գեներացում
    max=100
    min=999
    Randomize
    rand = Int((max-min+1)*Rnd+min)
    fileFrom = Project.Path &"Stores\SWIFT\HT300\ImportFile\IA000390.RJE"
    fileTo = Project.Path &"Stores\SWIFT\HT300\ImportFile\Import\IA000391.RJE"
    what = "CITI2111089856"
    fWith = "CITI2111089" & rand
    
    Log.Message(fWith)
    
    'Վերարտագրում է նախօրոք տրված ֆայլը մեկ այլ ֆայլի մեջ` կատարելով փոփոխություն
    Call Read_Write_File_Wiht_replace(fileFrom,fileTo,what,fWith)
    
    If SearchInPttel("frmPttel",1, "SWOUT") Then
        Call wMainForm.MainMenu.Click(c_AllActions)
        Call wMainForm.PopupMenu.Click(c_ToEdit)
        BuiltIn.Delay(2000)
        'Լրացնում է "Արժեք" դաշտը
        Call Rekvizit_Fill("Dialog",1,"General","VALUE","^A[Del]" & Project.Path & "Stores\SWIFT\HT300\ImportFile\Import\")
        'Սեղմել "Կատարել"
        Call ClickCmdButton(2, "Î³ï³ñ»É")
        BuiltIn.Delay(2000)
        Call Close_Window(wMDIClient, "frmPttel")
        Login("ARMSOFT")
    Else
        Log.Error "Can Not find (SWOUT)Parameter row!",,,ErrorColor
    End If
'-----------------------------------------------------------------------------
'----------------- Կատարել Ընդունել SWIFT համակարգից գործողությունը ------------------
'-----------------------------------------------------------------------------
    Log.Message "Կատարել Ընդունել SWIFT համակարգից գործողությունը",,,DivideColor
    
    'Մուտք գործել "S.W.I.F.T. ԱՇՏ"
    Call ChangeWorkspace(c_SWIFT)
    Call wTreeView.DblClickItem("|S.W.I.F.T. ²Þî                  |Ð³Õáñ¹³·ñáõÃÛáõÝÝ»ñÇ ÁÝ¹áõÝáõÙ|ÀÝ¹áõÝ»É S.W.I.F.T. Ñ³Ù³Ï³ñ·Çó")
    Call ClickCmdButton(5, "OK") 
    
    
    'Մուտք գործել Փոխանցումներ/Ստացված փաստաթղթեր թղթապանակ
    Call wTreeView.DblClickItem("|S.W.I.F.T. ²Þî                  |öáË³ÝóáõÙÝ»ñ|êï³óí³Í ÷áË³ÝóáõÙÝ»ñ")
    Call Rekvizit_Fill("Dialog",1,"General", "PERN",aqDateTime.Today)
    Call Rekvizit_Fill("Dialog",1,"General", "PERK",aqDateTime.Today)
    Call ClickCmdButton(2, "Î³ï³ñ»É") 
    BuiltIn.Delay(4000)    


End Sub

Sub Test_Initialize_SWIFT_300_RC()
    sDATE = "20020101"
    fDATE = "20260101"   
End Sub