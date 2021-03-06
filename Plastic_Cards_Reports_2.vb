'USEUNIT  Library_Common
'USEUNIT Library_Colour
'USEUNIT OLAP_Library
'USEUNIT Card_Library
'USEUNIT Mortgage_Library
'USEUNIT SWIFT_International_Payorder_Library
'USEUNIT Constants
Option Explicit

'Test Case Id - 160923

Sub Check_Reports_2()
  
    Dim sDATE,fDATE
    Dim MCReceivedTrans,SharedTerminalOperation,SMS_Messages
    Dim SentRecPayment_FilesHistory,SentReceived_FilesHistory
    Dim Path1,Path2,resultWorksheet,exists
    Dim SortArr(3)
     
    'Համակարգ մուտք գործել ARMSOFT օգտագործողով
    sDATE = "20030101"
    fDATE = "20260101"
    Call Initialize_AsBank("bank_Report", sDATE, fDATE)
    Login("ARMSOFT")
    
    Call SaveRAM_RowsLimit("100")
    
    'Մուտք գործել "Պլաստիկ քարտերի ԱՇՏ (SV)"
    Call ChangeWorkspace(c_CardsSV)
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''--- Պլաստիկ Քարտեր ԱՇՏ/MC Ստացված գործողություններ ---'''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "--- MC Ստացված գործողություններ ---" ,,, DivideColor  
    
    SortArr(0) = "fCARD"
    SortArr(1) = "FILEDATE"
    SortArr(2) = "fDATE"
    SortArr(3) = "fBILLSUM"
    
    Set MCReceivedTrans = New_MCReceivedTrans()
    With MCReceivedTrans
        .FileDate_1 = "031007"
        .FileDate_2 = "140808"
        .CardNumber = "????260006001938"
        .ShowAllTransactions = 1
        .View = "MCRcTrns"
        .FillInto = "0"
    End With
    
    Call GoToMCReceivedTrans_PlasticCarts(MCReceivedTrans) 
    Call CheckPttel_RowCount("frmPttel", 62)
    Call ColumnSorting(SortArr, 4, "frmPttel")
    
    Path1 = Project.Path & "Stores\Reports\Plastic Cards\Actual\Actual_8.xlsx"
    Path2 = Project.Path & "Stores\Reports\Plastic Cards\Expected\Expected_8.xlsx"
    resultWorksheet = Project.Path & "Stores\Reports\Plastic Cards\Result\Result_8.xlsx"
    
    'Արտահանել և Ð³Ù»Ù³ï»É »ñÏáõ EXCEL ý³ÛÉ»ñ
    Call ExportToExcel("frmPttel",Path1)
    Call CompareTwoExcelFiles(Path1, Path2, resultWorksheet)  
    Call CloseAllExcelFiles()
    
    Call Close_Pttel("frmPttel")
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''---  Պլաստիկ Քարտեր ԱՇՏ/Համատեղ տերմինալ գործողություններ ---''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "--- Համատեղ տերմինալ գործողություններ ---" ,,, DivideColor  
    
    Set SharedTerminalOperation = New_SharedTerminalOperations()
    With SharedTerminalOperation
        .FileDate_1 = "^A[Del]"
        .FileDate_2 = "^A[Del]"
        .CardNumber = "4335432108239446"
        .ShowMadeTransactions = 1
        .ShowArchivedOpers = 0
        .ShowInsideBankingCommission = 1
        .View = "VrtTrns\2"
        .FillInto = "1"
    End With
    
    Call GoToSharedTermOperation_PlasticCarts(SharedTerminalOperation) 
    
    Path1 = Project.Path & "Stores\Reports\Plastic Cards\Actual\Actual_9.xlsx"
    Path2 = Project.Path & "Stores\Reports\Plastic Cards\Expected\Expected_9.xlsx"
    resultWorksheet = Project.Path & "Stores\Reports\Plastic Cards\Result\Result_9.xlsx"
    
    'Արտահանել և Ð³Ù»Ù³ï»É »ñÏáõ EXCEL ý³ÛÉ»ñ
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

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''---  Պլաստիկ Քարտեր ԱՇՏ/SMS Հաղորդագրություն ---'''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "--- SMS Հաղորդագրություն ---" ,,, DivideColor  
    
    SortArr(0) = "FILEDATE"
    SortArr(1) = "CARDNUM"
    Set SMS_Messages = New_SMS_Messages()
    With SMS_Messages
        .FileDate_1 = "^A[Del]"
        .FileDate_2 = "^A[Del]"
        .ShowProcessed = 1
        .Archive = 0
        .View = "vACSMS"
        .FillInto = "0"
    End With

    Call GoToSMS_Messages_PlasticCarts(SMS_Messages)
    Call CheckPttel_RowCount("frmPttel", 35457)
    Call ColumnSorting(SortArr, 2, "frmPttel")

    Path1 = Project.Path & "Stores\Reports\Plastic Cards\Actual\Actual_10.xlsx"
    Path2 = Project.Path & "Stores\Reports\Plastic Cards\Expected\Expected_10.xlsx"
    resultWorksheet = Project.Path & "Stores\Reports\Plastic Cards\Result\Result_10.xlsx"
    
    'Արտահանել և Ð³Ù»Ù³ï»É »ñÏáõ EXCEL ý³ÛÉ»ñ
    Call ExportToExcel("frmPttel",Path1)
    Call CompareTwoExcelFiles(Path1, Path2, resultWorksheet)  
    Call CloseAllExcelFiles() 
    Call Close_Pttel("frmPttel")
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''--- Պլաստիկ Քարտեր ԱՇՏ/Ուղարկված/Ստացված Ֆայլերի պատմություն ---''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "--- Ուղարկված/Ստացված Ֆայլերի պատմություն ---" ,,, DivideColor  
    
    SortArr(0) = "fAPPLID"
    
    Set SentReceived_FilesHistory = New_SentReceived_FilesHistory()
    With SentReceived_FilesHistory
        .FileDate_1 = "010114"
        .FileDate_2 = "010120"
        .ShowOnlyErrors = 1
        .View = "SVLOGV"
        .FillInto = "0"
    End With
    
    Call GoToSentReceived_FilesHistory_PlasticCarts(SentReceived_FilesHistory)
    Call CheckPttel_RowCount("frmPttel", 4)
    Call ColumnSorting(SortArr, 1, "frmPttel")
    
    Path1 = Project.Path & "Stores\Reports\Plastic Cards\Actual\Actual_11.txt"
    Path2 = Project.Path & "Stores\Reports\Plastic Cards\Expected\Expected_11.txt"
    
    'Արտահանել և Ð³Ù»Ù³ï»É »ñÏáõ txt ý³ÛÉ»ñ
    Call ExportToTXTFromPttel("frmPttel",Path1)
    Call Compare_Files(Path2, Path1, "")  
    Call Close_Pttel("frmPttel") 
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''--- Պլաստիկ Քարտեր ԱՇՏ/Ուղարկված/Ստացված վճարային Ֆայլերի պատմություն---''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "--- Ուղարկված/Ստացված վճարային Ֆայլերի պատմություն ---" ,,, DivideColor  
    
    SortArr(0) = "fDATE"
    SortArr(1) = "fCARDNUMBER"
    SortArr(2) = "fSUMMA"
    
    Set SentRecPayment_FilesHistory = New_SentRecPayment_FilesHistory()
    With SentRecPayment_FilesHistory
        .FileDate_1 = "010113"
        .FileDate_2 = "010120"
        .PaymentFileState = ""
        .View = "PAYMVIEW"
        .FillInto = "0"
    End With
    
    Call GoToSentRecPayment_FilesHistory_PlasticCarts(SentRecPayment_FilesHistory)
    Call CheckPttel_RowCount("frmPttel", 85725)
    Call ColumnSorting(SortArr, 3, "frmPttel")
    
    Path1 = Project.Path & "Stores\Reports\Plastic Cards\Actual\Actual_12.xlsx"
    Path2 = Project.Path & "Stores\Reports\Plastic Cards\Expected\Expected_12.xlsx"
    resultWorksheet = Project.Path & "Stores\Reports\Plastic Cards\Result\Result_12.xlsx"
    
    'Արտահանել և Ð³Ù»Ù³ï»É »ñÏáõ EXCEL ý³ÛÉ»ñ
    Call ExportToExcel("frmPttel",Path1)
    Call CompareTwoExcelFiles(Path1, Path2, resultWorksheet)  
    Call CloseAllExcelFiles()
    Call Close_Pttel("frmPttel")
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''--- Պլաստիկ Քարտեր ԱՇՏ/Ուղարկված/Քարտային վճարումներ ---''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "--- Քարտային վճարումներ թղթապանակի ստուգում ---" ,,, DivideColor  
    
    SortArr(0) = "fCARDCODE"
    Call wTreeView.DblClickItem("|äÉ³ëïÇÏ ù³ñï»ñÇ ²Þî (SV)|ÂÕÃ³å³Ý³ÏÝ»ñ|ø³ñï³ÛÇÝ í×³ñáõÙÝ»ñ")
    BuiltIn.Delay(2000) 
    Call CheckPttel_RowCount("frmPttel", 62)
    Call ColumnSorting(SortArr, 1, "frmPttel")
    
    Path1 = Project.Path & "Stores\Reports\Plastic Cards\Actual\Actual_13.xlsx"
    Path2 = Project.Path & "Stores\Reports\Plastic Cards\Expected\Expected_13.xlsx"
    resultWorksheet = Project.Path & "Stores\Reports\Plastic Cards\Result\Result_13.xlsx"
    
    'Արտահանել և Ð³Ù»Ù³ï»É »ñÏáõ EXCEL ý³ÛÉ»ñ
    Call ExportToExcel("frmPttel",Path1)
    Call CompareTwoExcelFiles(Path1, Path2, resultWorksheet)  
    Call CloseAllExcelFiles()
    Call Close_Pttel("frmPttel")
    Call Close_AsBank()

End Sub    