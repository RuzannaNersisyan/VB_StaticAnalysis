'USEUNIT  Library_Common
'USEUNIT Library_Colour
'USEUNIT OLAP_Library
'USEUNIT Card_Library
'USEUNIT Mortgage_Library
'USEUNIT SWIFT_International_Payorder_Library
'USEUNIT Constants
Option Explicit

'Test Case Id - 160004

Sub Check_Reports_1()
  
    Dim sDATE,fDATE
    Dim Client,PlasticCart,CardAccTrans,ReceivedTrans,RecClearingTrans
    Dim Path1,Path2,resultWorksheet,exists
    Dim SortArr(4)
    SortArr(0) = "fKEY"
     
    'Համակարգ մուտք գործել ARMSOFT օգտագործողով
    sDATE = "20030101"
    fDATE = "20260101"
    Call Initialize_AsBank("bank_Report", sDATE, fDATE)
    Login("ARMSOFT")

    Call SaveRAM_RowsLimit("100")
    
    'Մուտք գործել "Պլաստիկ քարտերի ԱՇՏ (SV)"
    Call ChangeWorkspace(c_CardsSV)    
     
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''--- Պլաստիկ Քարտեր ԱՇՏ/Հաճախորդներ թղթապանակ --''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "--- Հաճախորդներ թղթապանակ (1) ---" ,,, DivideColor  
    
    Set Client = New_Clients()  
        Client.ClosedClients = 1

    Call GoToClients_PlasticCarts(Client)  
    Call CheckPttel_RowCount("frmPttel", 26111)
    Call ColumnSorting(SortArr, 1, "frmPttel")

    Path1 = Project.Path & "Stores\Reports\Plastic Cards\Actual\Actual_1.xlsx"
    Path2 = Project.Path & "Stores\Reports\Plastic Cards\Expected\Expected_1.xlsx"
    resultWorksheet = Project.Path & "Stores\Reports\Plastic Cards\Result\Result_1.xlsx"
    
    'Արտահանել և Ð³Ù»Ù³ï»É »ñÏáõ EXCEL ý³ÛÉ»ñ
    Call ExportToExcel("frmPttel",Path1)
    Call CompareTwoExcelFiles(Path1, Path2, resultWorksheet)  
    Call CloseAllExcelFiles()
    
    Call Close_Pttel("frmPttel")    
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''--- Պլաստիկ Քարտեր ԱՇՏ/Հաճախորդներ թղթապանակ ---'''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "--- Հաճախորդներ թղթապանակ (2) ---" ,,, DivideColor  
    
    Set Client = New_Clients()   
        Client.AccountMask = "33181921800"
        Client.DeepSearchByClientName = 1
        Client.Note = "1"
        Client.Note2 = "03"
        Client.Note3 = "01"
        Client.ClosedClients = 1
        Client.Reminders = 1
        Client.SocialInfo = 1
        Client.OtherInfo = 1
        Client.IncludeClosed = 1
        Client.BankIdDate = 1
        Client.FillInto = "1" 
        
    Call GoToClients_PlasticCarts(Client)  
    
    Path1 = Project.Path & "Stores\Reports\Plastic Cards\Actual\Actual_2.xlsx"
    Path2 = Project.Path & "Stores\Reports\Plastic Cards\Expected\Expected_2.xlsx"
    resultWorksheet = Project.Path & "Stores\Reports\Plastic Cards\Result\Result_2.xlsx"
        
    'Î³ï³ñáõÙ ¿ ëïáõ·áõÙ,»Ã» ÝÙ³Ý ³ÝáõÝáí ý³ÛÉ Ï³ ïñí³Í ÃÕÃ³å³Ý³ÏáõÙ ,çÝçáõÙ ¿   
    exists = aqFile.Exists(Path1)
    If exists Then
        aqFileSystem.DeleteFile(Path1)
    End If
    'Արտահանել և Ð³Ù»Ù³ï»É »ñÏáõ EXCEL ý³ÛÉ»ñ
    BuiltIn.Delay(3000)
    If Sys.Process("EXCEL").Exists Then
        Sys.Process("EXCEL").Window("XLMAIN", "* - Excel", 1).Window("XLDESK", "", 1).Window("EXCEL7", "*", 1).Keys("[F12]")
        Sys.Process("EXCEL").Window("#32770", "Save As", 1).Keys(Path1 & "[Enter]")
    Else 
        Log.Error "Excel does not Open!" ,,,ErrorColor
    End If 
        
    Call CompareTwoExcelFiles(Path1, Path2, resultWorksheet)  
    Call CloseAllExcelFiles() 
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''--- Պլաստիկ Քարտեր ԱՇՏ/Պլաստիկ քարտեր ---''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "--- Պլաստիկ քարտեր (1)---" ,,, DivideColor   
    
    SortArr(0) = "CardNum"
    Set PlasticCart = New_PlasticCarts()
    With PlasticCart
        .CardName = "Ð³×³Ëáñ¹ 00014031"
        .Client = "00014031"
        .Division = "P00"
        .Department = "061"
        .DatePeriod_1 = "240813"
        .DatePeriod_2 = "310815"
        .ValidFrom_1 = "240813"
        .CardType = "345"
        .CardStandard = "102"
        .Curr = "000"
        .Note2 = "POO"
        .Company = "139"
        .MobileServices = 0
        .Closed = 1
        .Limits = 1
        .OverLimits = 1
        .ClientInfo = 1
        .ExistsChanges = 1
        .OtherInfo = 1
        .View = "VCards"
        .FillInto = "0"
    End With
        
    Call GoToPlasticCarts_PlasticCarts(PlasticCart)  
    Call CheckPttel_RowCount("frmPttel", 3)
    Call ColumnSorting(SortArr, 1, "frmPttel")
    
    Path1 = Project.Path & "Stores\Reports\Plastic Cards\Actual\Actual_3.txt"
    Path2 = Project.Path & "Stores\Reports\Plastic Cards\Expected\Expected_3.txt"
    
    'Արտահանել և Ð³Ù»Ù³ï»É »ñÏáõ xml ý³ÛÉ»ñ
    Call ExportToTXTFromPttel("frmPttel",Path1)
    Call Compare_Files(Path2, Path1, "")
    Call Close_Pttel("frmPttel")
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''--- Պլաստիկ Քարտեր ԱՇՏ/Պլաստիկ քարտեր ---''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "--- Պլաստիկ քարտեր (2)---" ,,, DivideColor     
    
    Set PlasticCart = New_PlasticCarts()
        
    Call GoToPlasticCarts_PlasticCarts(PlasticCart)  
    Call CheckPttel_RowCount("frmPttel", 13604)
    Call ColumnSorting(SortArr, 1, "frmPttel")
    
    Path1 = Project.Path & "Stores\Reports\Plastic Cards\Actual\Actual_4.xlsx"
    Path2 = Project.Path & "Stores\Reports\Plastic Cards\Expected\Expected_4.xlsx"
    resultWorksheet = Project.Path & "Stores\Reports\Plastic Cards\Result\Result_4.xlsx"
    
    'Արտահանել և Ð³Ù»Ù³ï»É »ñÏáõ EXCEL ý³ÛÉ»ñ
    Call ExportToExcel("frmPttel",Path1)
    Call CompareTwoExcelFiles(Path1, Path2, resultWorksheet)  
    Call CloseAllExcelFiles()
    
    Call Close_Pttel("frmPttel")
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''--- Պլաստիկ Քարտեր ԱՇՏ/Քարտային հաշիվների գործողություններ ---''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "--- Քարտային հաշիվների գործողություններ ---" ,,, DivideColor
    
    SortArr(0) = "CardNum"
    SortArr(1) = "CURSUM"
    SortArr(2) = "COMMENT"
    SortArr(3) = "DVLOPED"
    SortArr(4) = "OPDATE"
    
    Set CardAccTrans = New_CardAccountsTrans()
    With CardAccTrans
        .DatePeriod_1 = "^A[Del]"&"010114"
        .DatePeriod_2 = "^A[Del]"&"011120"
        .CardNumber = "5160880000???880"
        .AccountMask = "???70120100"
        .ShowAllTransactions = 1
        .View = "CdAcTrns"
        .FillInto = "0"
    End With
    Call GoToCardAccTrans_PlasticCarts(CardAccTrans) 
    Call CheckPttel_RowCount("frmPttel", 18)
    Call ColumnSorting(SortArr, 5, "frmPttel")
    
    Path1 = Project.Path & "Stores\Reports\Plastic Cards\Actual\Actual_5.xlsx"
    Path2 = Project.Path & "Stores\Reports\Plastic Cards\Expected\Expected_5.xlsx"
    resultWorksheet = Project.Path & "Stores\Reports\Plastic Cards\Result\Result_5.xlsx"
    
    'Արտահանել և Ð³Ù»Ù³ï»É »ñÏáõ EXCEL ý³ÛÉ»ñ
    Call ExportToExcel("frmPttel",Path1)
    Call CompareTwoExcelFiles(Path1, Path2, resultWorksheet)  
    Call CloseAllExcelFiles()
        
    Call Close_Pttel("frmPttel")  
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''--- Պլաստիկ Քարտեր ԱՇՏ/Ստացված գործողություններ ---''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "--- Ստացված գործողություններ ---" ,,, DivideColor  
    
    SortArr(0) = "fDATE"
    SortArr(1) = "fBILLSUM"
    SortArr(2) = "fCARD"
    
    Set ReceivedTrans = New_ReceivedTransactions()
    With ReceivedTrans
        .FileDate_1 = "^A[Del]"&"260710"
        .FileDate_2 = "^A[Del]"&"010111"
        .CardsTransactions = 1
        .MerchantPointTransactions = 1
        .ShowMadeTransactions = 1
        .ShowAllRows = 1
        .View = "VRecTrns"
        .FillInto = "0"
    End With
    Call GoToReceivedTrans_PlasticCarts(ReceivedTrans) 
    Call CheckPttel_RowCount("frmPttel", 81339)
    Call ColumnSorting(SortArr, 3, "frmPttel")
    
    Path1 = Project.Path & "Stores\Reports\Plastic Cards\Actual\Actual_6.xlsx"
    Path2 = Project.Path & "Stores\Reports\Plastic Cards\Expected\Expected_6.xlsx"
    resultWorksheet = Project.Path & "Stores\Reports\Plastic Cards\Result\Result_6.xlsx"
    
    'Արտահանել և Ð³Ù»Ù³ï»É »ñÏáõ EXCEL ý³ÛÉ»ñ
    Call ExportToExcel("frmPttel",Path1)
    Call CompareTwoExcelFiles(Path1, Path2, resultWorksheet)  
    Call CloseAllExcelFiles()
    Call Close_Pttel("frmPttel") 
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''--- Պլաստիկ Քարտեր ԱՇՏ/Ստացված հանրագումարներ ---''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "--- Ստացված հանրագումարներ ---" ,,, DivideColor  
    
    Set RecClearingTrans = New_ReceivedClearingTransactions()
    With RecClearingTrans
        .FileDate_1 = "140510"
        .FileDate_2 = "010120"
        .CardBank = "21"
        .ShowMadeTransactions = 1
        .View = "VRcClear"
        .FillInto = "0"
    End With

    Call GoToRecClearingTrans_PlasticCarts(RecClearingTrans) 
    Call CheckPttel_RowCount("frmPttel", 55943)
    
    BuiltIn.Delay(1000)
    Call wMainForm.MainMenu.Click(c_Views & "|" & "Հերթական համար -> առաջին սյուն")
    BuiltIn.Delay(8000)

    
    Path1 = Project.Path & "Stores\Reports\Plastic Cards\Actual\Actual_7.xlsx"
    Path2 = Project.Path & "Stores\Reports\Plastic Cards\Expected\Expected_7.xlsx"
    resultWorksheet = Project.Path & "Stores\Reports\Plastic Cards\Result\Result_7.xlsx"
    
    'Արտահանել և Ð³Ù»Ù³ï»É »ñÏáõ EXCEL ý³ÛÉ»ñ
    Call ExportToExcel("frmPttel",Path1)
    Call CompareTwoExcelFiles(Path1, Path2, resultWorksheet)
    Call CloseAllExcelFiles()
    Call Close_Pttel("frmPttel") 
    Call Close_AsBank() 

End Sub