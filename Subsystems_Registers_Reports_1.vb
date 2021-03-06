'USEUNIT  Library_Common
'USEUNIT Library_Colour
'USEUNIT OLAP_Library
'USEUNIT Card_Library
'USEUNIT Mortgage_Library
'USEUNIT SWIFT_International_Payorder_Library
'USEUNIT Library_Contracts
'USEUNIT Constants
Option Explicit

'Test Case Id - 161550

Sub Subsystems_Registers_Reports_1()
  
    Dim sDATE,fDATE
    Dim Path1,Path2,resultWorksheet,folderName,Exists
    Dim SummaryOfContracts,BankAllSecurities,BankOwnSecurities
    Dim SortArr(3)
     
    'Համակարգ մուտք գործել ARMSOFT օգտագործողով
    sDATE = "20030101"
    fDATE = "20260101"
    Call Initialize_AsBank("bank_Report", sDATE, fDATE)
    Login("ARMSOFT")

    Call SaveRAM_RowsLimit("100")
    
    'Մուտք գործել "Ենթահամակարգեր (ՀԾ)"
    Call ChangeWorkspace(c_Subsystems)    
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''-- Ենթահամակարգեր (ՀԾ)|Հաշվ, մատյաններ|Ամփոփ|Ստացված գրավ, երաշխավորություն|Պայմանագրերի ամփոփում" --'''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "--- Ստացված գրավ, երաշխավորություն/Պայմանագրերի ամփոփում ---" ,,, DivideColor  
    
    SortArr(0) = "fCODE"
    SortArr(1) = "FCURMORT"
    SortArr(2) = "FSUMMA"
    
    Set SummaryOfContracts = New_SummaryOfContracts()
		With SummaryOfContracts
				.common.groupExist = true
				.common.date = "29/05/16"
				.common.agreeKind = "2"
				.common.agreeType = "12"
				.common.agreePaperN = "00009168"
				.common.agreeN = "SG1531"
				.common.curr = "000"
				.common.preferredCurr = "000"
				.common.client = "00009168"
				.common.clientName = "Ð³×³Ëáñ¹ 00009168"
				.common.clientInsurance = "00009168"
				.common.clientNameInsurance = "Ð³×³Ëáñ¹ 00009168"
				.common.groupInsurance = ""
				.common.isSignedStart = "19/07/02"
				.common.isSignedEnd = "17/10/16"
				.common.show = "01,02,03,04,06,09"
				.additional.insuranceType = ""
				.additional.insuranceN = "TV12824"
				.additional.guaranteed_InsuranceType = ""
				.additional.guaranteed_InsuranceN = ""
				.additional.office = "P04"
				.additional.group  = "05"
				.additional.accessType = "N11"
		End with

    folderName = "|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|Ð³ßí»ïíáõÃÛáõÝÝ»ñ,  Ù³ïÛ³ÝÝ»ñ|²Ù÷á÷|êï³óí³Í ·ñ³í, »ñ³ßË³íáñáõÃÛáõÝ|"
    Call GoTo_SummaryOfContracts(folderName, "ä³ÛÙ³Ý³·ñ»ñÇ ³Ù÷á÷áõÙ", SummaryOfContracts)
    Call WaitForExecutionProgress()
    Call CheckPttel_RowCount("frmPttel", 1)
    Call columnSorting(SortArr, 1, "frmPttel")
    
    Path1 = Project.Path & "Stores\Reports\Subsystems\Reports Registers\Actual\Actual_1.txt"
    Path2 = Project.Path & "Stores\Reports\Subsystems\Reports Registers\Expected\Expected_1.txt"
    
    'Արտահանել և Ð³Ù»Ù³ï»É »ñÏáõ xml ý³ÛÉ»ñ
    Call ExportToTXTFromPttel("frmPttel",Path1)
    Call Compare_Files(Path2, Path1, "")
    Call Close_Pttel("frmPttel")
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''-- Ենթահամակարգեր (ՀԾ)|Հաշվ, մատյաններ|Ամփոփ|Ստացված գրավ, երաշխավորություն|Պայմանագրերի ամփոփում (Միայն փակված)" --'''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "--- Ստացված գրավ, երաշխավորություն/Պայմանագրերի ամփոփում (Միայն փակված) ---" ,,, DivideColor  
    
		Set SummaryOfContracts = New_SummaryOfContracts()
		With SummaryOfContracts    
		    .common.closeDateExists = true
				.common.groupExist = true
		End with
    
    Call GoTo_SummaryOfContracts(folderName, "ä³ÛÙ³Ý³·ñ»ñÇ ³Ù÷á÷áõÙ (ØÇ³ÛÝ ÷³Ïí³Í)", SummaryOfContracts)
    Call WaitForExecutionProgress()
    
    Call CheckPttel_RowCount("frmPttel", 13998)
    Call columnSorting(SortArr, 3, "frmPttel")
    
    Path1 = Project.Path & "Stores\Reports\Subsystems\Reports Registers\Actual\Actual_2.xlsx"
    Path2 = Project.Path & "Stores\Reports\Subsystems\Reports Registers\Expected\Expected_2.xlsx"
    resultWorksheet = Project.Path & "Stores\Reports\Subsystems\Reports Registers\Result\Result_2.xlsx"
    
    'Արտահանել և Ð³Ù»Ù³ï»É »ñÏáõ EXCEL ý³ÛÉ»ñ
    Call ExportToExcel("frmPttel",Path1)
    Call CompareTwoExcelFiles(Path1, Path2, resultWorksheet)  
    Call CloseAllExcelFiles()
    Call Close_Pttel("frmPttel")  

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''-- Ենթահամակարգեր (ՀԾ)|Հաշվ, մատյաններ|Ամփոփ|Տրամադրված գրավ, երաշխավորություն|Պայմանագրերի ամփոփում" --'''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "--- Տրամադրված գրավ, երաշխավորություն/Պայմանագրերի ամփոփում ---" ,,, DivideColor  
    
    Set SummaryOfContracts = New_SummaryOfContracts()
		With SummaryOfContracts
				.common.checkBoxExist = true
				.common.LRCodeExist = true
				.common.date = ""
				.common.agreeKind = ""
				.common.agreeType = ""
				.common.agreePaperN = "00000242"
				.common.agreeN = ""
				.common.curr = ""
				.common.preferredCurr = ""
				.common.client = "00000242"
				.common.clientName = "Ð³×³Ëáñ¹ 00000242"
				.common.groupInsurance = ""
				.common.isSignedStart = "21/04/03"
				.common.isSignedEnd = "30/12/15"
				.common.showWithoutExpiredPart = 1
				.common.showNoWriteOffs = 0
				.common.show = "01,02,03,04,06,09"
				.additional.insuranceType = ""
				.additional.insuranceN = ""
				.additional.guaranteed_InsuranceType = ""
				.additional.guaranteed_InsuranceN = ""
				.additional.office = ""
				.additional.group  = "02"
				.additional.accessType = "M21"
		End with

		folderName = "|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|Ð³ßí»ïíáõÃÛáõÝÝ»ñ,  Ù³ïÛ³ÝÝ»ñ|²Ù÷á÷|îñ³Ù³¹ñí³Í ·ñ³í, »ñ³ßË³íáñáõÃÛáõÝ|"
    Call GoTo_SummaryOfContracts(folderName, "ä³ÛÙ³Ý³·ñ»ñÇ ³Ù÷á÷áõÙ", SummaryOfContracts)
    Call WaitForExecutionProgress()
    Call CheckPttel_RowCount("frmPttel", 64)
    Call columnSorting(SortArr, 3, "frmPttel")
    
    Path1 = Project.Path & "Stores\Reports\Subsystems\Reports Registers\Actual\Actual_3.txt"
    Path2 = Project.Path & "Stores\Reports\Subsystems\Reports Registers\Expected\Expected_3.txt"
    
    'Արտահանել և Ð³Ù»Ù³ï»É »ñÏáõ xml ý³ÛÉ»ñ
    Call ExportToTXTFromPttel("frmPttel",Path1)
    Call Compare_Files(Path2, Path1, "")
    Call Close_Pttel("frmPttel")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''-- Ենթահամակարգեր (ՀԾ)|Հաշվ, մատյաններ|Ամփոփ|Տրամադրված գրավ, երաշխավորություն|Պայմանագրերի ամփոփում (Միայն փակված)" --'''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "--- Տրամադրված գրավ, երաշխավորություն/Պայմանագրերի ամփոփում (Միայն փակված) ---" ,,, DivideColor  
    
		Set SummaryOfContracts = New_SummaryOfContracts()
		With SummaryOfContracts    
				.common.checkBoxExist = true
				.common.LRCodeExist = true
				.common.date = "12/12/12"
				.common.closeDateExists = true
				.common.closeDateStart = "16/07/04"
				.common.closeDateEnd = "16/07/04"
				.common.agreeKind = "4"
				.common.agreeType = "2"
				.common.agreePaperN = "00000242"
				.common.agreeN = "TV12265"
				.common.curr = "045"
				.common.preferredCurr = "045"
				.common.client = "00000242"
				.common.clientName = "Ð³×³Ëáñ¹ 00000242"
				.common.isSignedStart = "24/09/03"
				.common.isSignedEnd = "07/06/04"
				.common.showWithoutExpiredPart = 1
				.common.showNoWriteOffs = 1
				.common.show = "02,03,04"
        .common.fill = "1"
				.additional.insuranceType = ""
				.additional.insuranceN = ""
				.additional.guaranteed_InsuranceType = ""
				.additional.guaranteed_InsuranceN = ""
				.additional.office = "P00"
				.additional.group  = "02"
				.additional.accessType = "M21"
		End with
    
    Call GoTo_SummaryOfContracts(folderName, "ä³ÛÙ³Ý³·ñ»ñÇ ³Ù÷á÷áõÙ (ØÇ³ÛÝ ÷³Ïí³Í)", SummaryOfContracts)
    
    Path1 = Project.Path & "Stores\Reports\Subsystems\Reports Registers\Actual\Actual_4.xlsx"
    Path2 = Project.Path & "Stores\Reports\Subsystems\Reports Registers\Expected\Expected_4.xlsx"
    resultWorksheet = Project.Path & "Stores\Reports\Subsystems\Reports Registers\Result\Result_4.xlsx"
    
    'Î³ï³ñáõÙ ¿ ëïáõ·áõÙ,»Ã» ÝÙ³Ý ³ÝáõÝáí ý³ÛÉ Ï³ ïñí³Í ÃÕÃ³å³Ý³ÏáõÙ ,çÝçáõÙ ¿   
    Exists = aqFile.Exists(Path1)
    If Exists Then
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

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''-- Ենթահամակարգեր (ՀԾ)|Հաշվ, մատյաններ|Ամփոփ|Բանկի բոլոր արժեթղթերը" --'''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "--- Բանկի բոլոր արժեթղթերը ---" ,,, DivideColor      
    

    folderName = "|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|Ð³ßí»ïíáõÃÛáõÝÝ»ñ,  Ù³ïÛ³ÝÝ»ñ|²Ù÷á÷|"
    
    Set BankAllSecurities = New_BankAllSecurities()
		With BankAllSecurities
		    .startDate = "01/01/03"
		    .endDate = "01/01/03"
		End with
    
    Call GoTo_BankAllSecurities(folderName, BankAllSecurities)
    Call WaitForExecutionProgress()
    
    Call CheckPttel_RowCount("frmPttel", 2)
    Call columnSorting(SortArr, 3, "frmPttel")
    
    Path1 = Project.Path & "Stores\Reports\Subsystems\Reports Registers\Actual\Actual_5.xlsx"
    Path2 = Project.Path & "Stores\Reports\Subsystems\Reports Registers\Expected\Expected_5.xlsx"
    resultWorksheet = Project.Path & "Stores\Reports\Subsystems\Reports Registers\Result\Result_5.xlsx"
    
    'Արտահանել և Ð³Ù»Ù³ï»É »ñÏáõ EXCEL ý³ÛÉ»ñ
    Call ExportToExcel("frmPttel",Path1)
    Call CompareTwoExcelFiles(Path1, Path2, resultWorksheet)  
    Call CloseAllExcelFiles()
    Call Close_Pttel("frmPttel")
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''-- Ենթահամակարգեր (ՀԾ)|Հաշվ, մատյաններ|Ամփոփ|Բանկի սեփական արժեթղթերը" --'''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "--- Բանկի սեփական արժեթղթերը ---" ,,, DivideColor      
    
    SortArr(0) = "fCOM"
    Set BankOwnSecurities = New_BankOwnSecurities()
		With BankOwnSecurities
				.date = "13/11/11"
        .yieldCurveDate = "27/05/21"
				.issue = ""
				.showWithoutRepo = 1
        .summaryByReleases = 1
		End With
    
    Call GoTo_bankOwnSecurities(folderName, BankOwnSecurities)
    Call WaitForExecutionProgress()
    
    Call CheckPttel_RowCount("frmPttel", 7)
    Call columnSorting(SortArr, 1, "frmPttel")
    
    Path1 = Project.Path & "Stores\Reports\Subsystems\Reports Registers\Actual\Actual_6.xlsx"
    Path2 = Project.Path & "Stores\Reports\Subsystems\Reports Registers\Expected\Expected_6.xlsx"
    resultWorksheet = Project.Path & "Stores\Reports\Subsystems\Reports Registers\Result\Result_6.xlsx"
    
    'Արտահանել և Ð³Ù»Ù³ï»É »ñÏáõ EXCEL ý³ÛÉ»ñ
    Call ExportToExcel("frmPttel",Path1)
    Call CompareTwoExcelFiles(Path1, Path2, resultWorksheet)  
    Call CloseAllExcelFiles()
    Call Close_Pttel("frmPttel")     

    Call Close_AsBank() 
End Sub