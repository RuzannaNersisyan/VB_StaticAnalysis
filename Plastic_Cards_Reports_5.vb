'USEUNIT  Library_Common
'USEUNIT Library_Colour
'USEUNIT OLAP_Library
'USEUNIT Card_Library
'USEUNIT Mortgage_Library
'USEUNIT SWIFT_International_Payorder_Library
'USEUNIT Constants
Option Explicit

'Test Case Id - 161030

Sub Check_Reports_5()
  
    Dim sDATE,fDATE
    Dim Path1,Path2,resultWorksheet,exists
    Dim CardSystemsExchangeRates,DealingExchangeRates,CBExchangeRates
    Dim SortArr(3)
     
    'Համակարգ մուտք գործել ARMSOFT օգտագործողով
    sDATE = "20030101"
    fDATE = "20260101"
    Call Initialize_AsBank("bank_Report", sDATE, fDATE)
    Login("ARMSOFT")
    
    Call SaveRAM_RowsLimit("100")
    
    'Մուտք գործել "Պլաստիկ քարտերի ԱՇՏ (SV)"
    Call ChangeWorkspace(c_CardsSV)
    
    SortArr(0) = "fDATE"
    SortArr(1) = "CUR1"
    SortArr(2) = "VALUE1"
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''--- Պլաստիկ Քարտեր ԱՇՏ/Հաշվետվություններ, մատյաններ/Արտարժույթների փոխարժեքներ/Քարտային համակարգերի փոխարժ. ---'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "--- Քարտային համակարգերի փոխարժ ---" ,,, DivideColor  

    Set CardSystemsExchangeRates = New_CardSystemsExchangeRates()
    
    With CardSystemsExchangeRates
        .DatePeriod_Start = "010103"
        .DatePeriod_End = "010120"
        .Curr = "045"
    End With
    
    Call GoTo_CardSystemsExchangeRates_PlasticCards(CardSystemsExchangeRates)
    Call CheckPttel_RowCount("frmPttel", 822)
    Call ColumnSorting(SortArr, 3, "frmPttel")
    
    Path1 = Project.Path & "Stores\Reports\Plastic Cards\Actual\Actual_25.xlsx"
    Path2 = Project.Path & "Stores\Reports\Plastic Cards\Expected\Expected_25.xlsx"
    resultWorksheet = Project.Path & "Stores\Reports\Plastic Cards\Result\Result_25.xlsx"
    
    'Արտահանել և Ð³Ù»Ù³ï»É »ñÏáõ EXCEL ý³ÛÉ»ñ
    Call ExportToExcel("frmPttel",Path1)
    Call CompareTwoExcelFiles(Path1, Path2, resultWorksheet)  
    Call CloseAllExcelFiles()
    Call Close_Pttel("frmPttel")
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''--- Պլաստիկ Քարտեր ԱՇՏ/Հաշվետվություններ, մատյաններ/Արտարժույթների փոխարժեքներ/Դիլինգային փոխարժեքներ ---''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "--- Դիլինգային փոխարժեքներ (1)---" ,,, DivideColor  

    Set DealingExchangeRates = New_DealingExchangeRates()
    With DealingExchangeRates
        .DatePeriod_Start = "010108"
        .DatePeriod_End = "010120"
        .Curr1 = ""
        .Curr2 = ""
        .RateType = "0"
        .ShowTheLastRates = 0
        .Division = ""
    End With
    
    Call GoTo_DealingExchangeRates_PlasticCards(DealingExchangeRates)
    Call CheckPttel_RowCount("frmPttel", 29277)
    Call wMainForm.MainMenu.Click(c_Views & "|" & "Սորտավորած Ստեղծման ամսաթիվ")
    BuiltIn.Delay(3000)
        
    Path1 = Project.Path & "Stores\Reports\Plastic Cards\Actual\Actual_26.xlsx"
    Path2 = Project.Path & "Stores\Reports\Plastic Cards\Expected\Expected_26.xlsx"
    resultWorksheet = Project.Path & "Stores\Reports\Plastic Cards\Result\Result_26.xlsx"
    
    'Արտահանել և Ð³Ù»Ù³ï»É »ñÏáõ EXCEL ý³ÛÉ»ñ
    Call ExportToExcel("frmPttel",Path1)
    Call CompareTwoExcelFiles(Path1, Path2, resultWorksheet)  
    Call CloseAllExcelFiles()
    Call Close_Pttel("frmPttel")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''--- Պլաստիկ Քարտեր ԱՇՏ/Հաշվետվություններ, մատյաններ/Արտարժույթների փոխարժեքներ/Դիլինգային փոխարժեքներ ---''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "--- Դիլինգային փոխարժեքներ (2)---" ,,, DivideColor  

    Set DealingExchangeRates = New_DealingExchangeRates()
    With DealingExchangeRates
        .DatePeriod_Start = "121206"
        .DatePeriod_End = "010120"
        .Curr1 = "001"
        .Curr2 = "000"
        .RateType = "1"
        .ShowTheLastRates = 1
        .Division = "P00"
    End With
    
    Call GoTo_DealingExchangeRates_PlasticCards(DealingExchangeRates)
    Call CheckPttel_RowCount("frmPttel", 0)
    Call Close_Pttel("frmPttel")
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''--- Պլաստիկ Քարտեր ԱՇՏ/Հաշվետվություններ, մատյաններ/Արտարժույթների փոխարժեքներ/ՀՀ ԿԲ փոխարժեքներ ---''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    SortArr(0) = "fPARID"
    SortArr(1) = "fDATE"
    SortArr(2) = "VALUE"
    Log.Message "--- ՀՀ ԿԲ փոխարժեքներ ---" ,,, DivideColor  
    
    Set CBExchangeRates = New_CBExchangeRates()
    With CBExchangeRates
        .DatePeriod_Start = "010103"
        .DatePeriod_End = "010120"
        .Curr = ""
    End With
    
    Call GoTo_CBExchangeRates_PlasticCards(CBExchangeRates)
    Call CheckPttel_RowCount("frmPttel", 29585)
    Call ColumnSorting(SortArr, 3, "frmPttel")
    
    Path1 = Project.Path & "Stores\Reports\Plastic Cards\Actual\Actual_27.xlsx"
    Path2 = Project.Path & "Stores\Reports\Plastic Cards\Expected\Expected_27.xlsx"
    resultWorksheet = Project.Path & "Stores\Reports\Plastic Cards\Result\Result_27.xlsx"
    
    'Արտահանել և Ð³Ù»Ù³ï»É »ñÏáõ EXCEL ý³ÛÉ»ñ
    Call ExportToExcel("frmPttel",Path1)
    Call CompareTwoExcelFiles(Path1, Path2, resultWorksheet)  
    Call CloseAllExcelFiles()
    Call Close_Pttel("frmPttel")
    
    Call Close_AsBank()
End Sub    