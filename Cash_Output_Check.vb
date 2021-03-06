'USEUNIT Library_Common
'USEUNIT Library_Contracts
'USEUNIT Constants
'USEUNIT Library_Colour
'USEUNIT DAHK_Library_Filter
'USEUNIT CashOutput_Confirmpases_Library
'USEUNIT Payment_Except_Library
'USEUNIT Payment_Order_ConfirmPhases_Library
'USEUNIT Akreditiv_Library
'USEUNIT Main_Accountant_Filter_Library
'USEUNIT SWIFT_International_Payorder_Library
'USEUNIT Library_CheckDB 
'USEUNIT Online_PaySys_Library
Option Explicit

'Test Case ID 179157

Dim sDate, eDate, folderName, expectedFile, actualFilePath, actualFile, currHour
Dim cashOutputCreate, cashOutputEdit, workingDocs, verifyDoc, currentDate, param
Dim fBODY, dbo_FOLDERS(2), dbo_PAYMENTS

Sub Cash_Output_Check_Test()
				Call Test_Inintialize()

				' Համակարգ մուտք գործել ARMSOFT օգտագործողով
				Log.Message "Համակարգ մուտք գործել ARMSOFT օգտագործողով", "", pmNormal, DivideColor
				Call Test_StartUp()
    
    ' Կարգավորումների ներմուծում
    If Not Input_Config("X:\Testing\CashOutput confirm phases\CashOutput_Allconditions.txt") Then
        Log.Error "The configuration doesn't input", "", pmNormal, ErrorColor
    End If
    
    ' Մուտք Գլխավոր հաշվապահի ԱՇՏ
				Call ChangeWorkspace(c_ChiefAcc)
				
				' Մուտք գործել Հաշիվներ թղթապանակ
				Log.Message "Մուտք գործել Հաշիվներ թղթապանակ", "", pmNormal, DivideColor
				Call OpenAccauntsFolder(folderName & "Ð³ßÇíÝ»ñ","1","","72110253300","","","","","","",0,"","","","","",0,0,0,"","","","","","ACCS","0")		
				Call CheckPttel_RowCount("frmPttel", 1) 
		
				' Ստեղծել Կանխիկ ելք փաստաթղթի սևագիր
				Log.Message "Ստեղծել Կանխիկ ելք փաստաթղթի սևագիր", "", pmNormal, DivideColor
				Call Create_Cash_Output(cashOutputCreate, "Î³ï³ñ»É")
    
    ' Քաղվածքի պահպանում 
				Log.Message "Քաղվածքի պահպանում", "", pmNormal, DivideColor
    Call SaveDoc(actualFilePath, actualFile) 
				
				' Փակել Քաղվածքի պատուհանը 
				Call Close_Window(wMDIClient, "FrmSpr")
    
    ' Փակել Հաշիվներ թղթապանակը
				Call Close_Window(wMDIClient, "frmPttel")
    
    ' Կանխիկ ելք փաստաթղթի ստեղծումից հետո SQL ստուգում
    Log.Message "Կանխիկ ելք փաստաթղթի ստեղծումից հետո SQL ստուգում", "", pmNormal, DivideColor
    Call DB_Initialize()
    Call Check_DB_Create()
				
				' Փաստացի քաղվածքի համեմատում սպասվողի հետ
				Log.Message "Փաստացի քաղվածքի համեմատում սպասվողի հետ", "", pmNormal, DivideColor
				param = "N\s\d{1,6}\s*.\d{1,10}\s{0,}.|Date\s\d{1,2}.\d{1,2}.\d{1,2}\s(\d{1,2}:\d{1,2})*"
    Call Compare_Files(actualFilePath & actualFile, expectedFile, param)
				
				' Մուտք գործել Աշխատանքային փաստաթղթեր թղթապանակ
				folderName = "|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|"
				Call GoTo_MainAccWorkingDocuments(folderName, workingDocs)
				
				' Ստուգել ստեղծված փաստաթղթի արժեքները և խմբագրել այն 
				Log.Message "Ստուգել ստեղծված փաստաթղթի արժեքները և խմբագրել այն", "", pmNormal, DivideColor
    With cashOutputCreate
        .chargeTab.chargeAccForCheck = "000001100  "
        .coinTab.coinForCheck = "43.25"
        .chargeTab.operType = "1"
        .chargeTab.operPlace = "3"
    End With
    If aqDateTime.Compare(aqConvert.DateTimeToFormatStr(aqDateTime.Time, "%H:%M"), "16:00") < 0 Then
        cashOutputCreate.chargeTab.timeForCheck = "1"
    Else
        cashOutputCreate.chargeTab.timeForCheck = "2"
    End If
				If SearchInPttel("frmPttel", 2, cashOutputCreate.commonTab.docNum) Then
    				Call Edit_Cash_Output(cashOutputCreate, cashOutputEdit, "Î³ï³ñ»É")
    Else
        Log.Error "Can't find searched row.", "", pmNormal, ErrorColor
    End If
    
    ' Կանխիկ ելք փաստաթուղթը խմբագրելուց հետո SQL ստուգում
    Log.Message "Կանխիկ ելք փաստաթուղթը խմբագրելուց հետո SQL ստուգում", "", pmNormal, DivideColor
    Call Check_DB_Edit()
				
				' Ուղարկել դրամարկղ
				Log.Message "Ուղարկել դրամարկղ", "", pmNormal, DivideColor
				Call Online_PaySys_Send_To_Verify(1)
    
    ' Ուղարկել դրամարկղից հետո SQL ստուգում
    Log.Message "Ուղարկել դրամարկղից հետո SQL ստուգում", "", pmNormal, DivideColor
    Call Check_DB_SendToVerify()
				
				' Մուտք գործել Աշխատանքային փաստաթղթեր թղթապանակ
				Call GoTo_MainAccWorkingDocuments(folderName, workingDocs)
    
    ' Ստուգել, որ առկա է մեր ավելացրած փաստաթուղթը
				Log.Message "Ստուգել, որ առկա է մեր ավելացրած փաստաթուղթը", "", pmNormal, DivideColor
    With cashOutputEdit
        .coinTab.coinForCheck = "0.00"
        .chargeTab.chargeCurrForCheck = "002"
        .chargeTab.chargeAccForCheck = "72110253300  "
        .chargeTab.legalStatusForCheck = "33"
        .commonTab.idGiveDateForCheck = "13/03/2013"
        .commonTab.idValidUntilForCheck = "13/03/2023"
        .chargeTab.nonResidentForCheck = 0
        .chargeTab.operAreaForCheck = "6"
    End With
				If SearchInPttel("frmPttel", 2, cashOutputEdit.commonTab.docNum) Then
        wMDIClient.vbObject("frmPttel").Keys("^w")
        If wMDIClient.WaitVBObject("frmASDocForm", 3000).Exists Then
            Call Check_Cash_Output(cashOutputEdit)
            Call ClickCmdButton(1, "OK") 
        Else 
            Log.Error "Can't open frmASDocForm window.", "", pmNormal, ErrorColor
        End If
    Else
        Log.Error "Can't find searched row.", "", pmNormal, ErrorColor
    End If
				
				' Վավերացնել փաստաթուղթը
				Log.Message "Վավերացնել փաստաթուղթը", "", pmNormal, DivideColor
				Call Validate_Doc()
				
				' Փակել Աշխատանքային փաստաթղթեր թղթապանակը
				Call Close_Window(wMDIClient, "frmPttel")
    
    ' Վավերացումից հետո SQL ստուգում
    Log.Message "Վավերացումից հետո SQL ստուգում", "", pmNormal, DivideColor
    Call Check_DB_Validate()
				
				' Մուտք գործել Ստեղծված փաստաթղթեր թղթապանակ
				currentDate = aqConvert.DateTimeToFormatStr(aqDateTime.Now(),"%d%m%y")
				Call OpenCreatedDocFolder(folderName & "êï»ÕÍí³Í ÷³ëï³ÃÕÃ»ñ", currentDate, currentDate, null, "KasRsOrd")
				
				' Ստուգել, որ առկա է մեր ավելացրած փաստաթուղթը
				Log.Message "Ստուգել, որ առկա է մեր ավելացրած փաստաթուղթը", "", pmNormal, DivideColor
				If SearchInPttel("frmPttel", 2, cashOutputEdit.fIsn) Then
    				wMDIClient.vbObject("frmPttel").Keys("^w")
        If wMDIClient.WaitVBObject("frmASDocForm", 3000).Exists Then
            Call Check_Cash_Output(cashOutputEdit)
            Call ClickCmdButton(1, "OK") 
        Else 
            Log.Error "Can't open frmASDocForm window.", "", pmNormal, ErrorColor
        End If
    Else
        Log.Error "Can't find searched row.", "", pmNormal, ErrorColor
    End If
				
				' Ջնջել Կանխիկի ելք փաստաթուղթը
				Log.Message "Ջնջել Կանխիկի ելք փաստաթուղթը", "", pmNormal, DivideColor
				Call SearchAndDelete("frmPttel", 2, cashOutputEdit.fIsn, "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ")
				
				' Փակել Ստեղծված փաստաթղթեր թղթապանակը
				Call Close_Window(wMDIClient, "frmPttel")
    
    ' Ջնջելուց հետո SQL ստուգում
    Log.Message "Ջնջելուց հետո SQL ստուգում", "", pmNormal, DivideColor
    Call Check_DB_Delete()
				
				' Փակել ծրագիրը
				Call Close_AsBank()
End	Sub

Sub Test_StartUp()
				Call Initialize_AsBank("bank", sDate, eDate)   
				Login("ARMSOFT")
				' Մուտք Գլխավոր հաշվապահի ԱՇՏ
				Call ChangeWorkspace(c_ChiefAcc)
End Sub

Sub Test_Inintialize()
				sDate = "20030101"
				eDate = "20250101"
		
				folderName = "|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|"
				expectedFile = Project.Path &  "Stores\Cash_Input_Output\Expected\Expected_Cash_Output_Create.txt"
				actualFilePath = Project.Path &  "Stores\Cash_Input_Output\Actual\"
    actualFile = "Actual_Cash_Output_Create.txt"
		
				Set cashOutputCreate = New_CashOutput(0, 0, 0)
				With cashOutputCreate
								.commonTab.office = "00"
        .commonTab.department = "1"
        .commonTab.date = "22/02/22"
        .commonTab.dateForCheck = "22/02/22"
        .commonTab.cashRegister = "001"
        .commonTab.cashRegisterAcc = "000001102  "
        .commonTab.curr = "002"
        .commonTab.accDebet = "72110253300"
        .commonTab.amount = "5000.00"
        .commonTab.amountForCheck = "5,000.00"
        .commonTab.cashierChar = "051"
        .commonTab.base = "Ð³Ù³Ó³ÛÝ é»åá-Ñ³Ù³Ó³ÛÝ³·ñÇ"
        .commonTab.aim = "Ð³Ù³Ó³ÛÝ ·ñ³íÇ å³ÛÙ³Ý³·ñÇ Éñ³óÙ³Ý"
        .commonTab.payer = "00034855"
        .commonTab.payerLegalStatus = "ֆիզԱնձ"
        .commonTab.name = "Ø»ñ"
        .commonTab.surname = "´³ÝÏÇ"
        .commonTab.id = "CC00034855"
        .commonTab.idForCheck = "CC00034855"
        .commonTab.idType = "05"
        .commonTab.idTypeForCheck = "05"
        .commonTab.idGivenBy = "001"
        .commonTab.idGivenByForCheck = "001"
        .commonTab.idGiveDate = "12062013"
        .commonTab.idGiveDateForCheck = "12/06/2013"
        .commonTab.idValidUntil = "12062025"
        .commonTab.idValidUntilForCheck = "12/06/2025"
        .commonTab.birthDate = "26051988"
        .commonTab.birthDateForCheck = "26/05/1988"
        .commonTab.citizenship = ""
        .commonTab.country = "AT"
        .commonTab.residence = "990000002"
        .commonTab.city = "ìÇ»Ý³"
        .commonTab.street = "15"
        .commonTab.apartment = "µÝ. 69    "
        .commonTab.house = "ïáõÝ 6    "
        .commonTab.email = ""
        .commonTab.emailForCheck = ""
        .chargeTab.office = .commonTab.office
        .chargeTab.department = .commonTab.department
        .chargeTab.chargeAcc = "000001100  "
        .chargeTab.chargeAccForCheck = "72110253300  "
        .chargeTab.chargeCurr = "000"
        .chargeTab.chargeCurrForCheck = "000"
        .chargeTab.cbExchangeRate = "1.0000/1"
        .chargeTab.chargeType = "TT"
        .chargeTab.chargeAmount = "0.00"
        .chargeTab.chargeAmoForCheck = "0.00"
        .chargeTab.chargePercent = "0.0000"
        .chargeTab.chargePerForCheck = "0.0000"
        .chargeTab.incomeAcc = "77796901013"
        .chargeTab.incomeAccCurr = "000"
        .chargeTab.buyAndSell = ""
        .chargeTab.buyAndSellForCheck = ""
        .chargeTab.operType = ""
        .chargeTab.operPlace = ""
        .chargeTab.operArea = "7"
        .chargeTab.operAreaForCheck = "7"
        .chargeTab.nonResident = 1
        .chargeTab.nonResidentForCheck = 1
        .chargeTab.legalStatus = "21"
        .chargeTab.legalStatusForCheck = "21"
        .chargeTab.comment = ""
        .chargeTab.commentForCheck = ""
        .chargeTab.clientAgreeData = "üÇ½ÇÏ³Ï³Ý ³ÝÓ"
        .coinTab.coin = "43.25"
        .coinTab.coinForCheck = "0.00"
        .coinTab.coinPayCurr = "000"
        .coinTab.coinBuyAndSell = "1"
        .coinTab.coinPayAcc = "000001100  "
        .coinTab.coinExchangeRate = "782.9295/1"
        .coinTab.coinCBExchangeRate = "782.9300/1"
        .coinTab.coinPayAmount = "33861.70"
        .coinTab.coinPayAmountForCheck = "33,861.70"
        .coinTab.amountWithMainCurr = "4,956.75"
        .coinTab.amountCurrForCheck = "4,956.75"
        .coinTab.incomeOutChange = "000931900  "
        .coinTab.damagesOutChange = "001434300  "
				End With
				
				Set cashOutputEdit = New_CashOutput(1, 1, 0)
				With cashOutputEdit
								.commonTab.office = "00"
        .commonTab.department = "1"
        .commonTab.date = "250122"
        .commonTab.dateForCheck = "25/01/22"
        .commonTab.cashRegister = "001"
        .commonTab.cashRegisterAcc = "000001102  "
        .commonTab.curr = "002"
        .commonTab.accDebet = "72110253300"
        .commonTab.amount = "254.31"
        .commonTab.amountForCheck = "254.31"
        .commonTab.cashierChar = "06 "
        .commonTab.base = "Üáñ Ù»ÏÝ³µ³Ý."
        .commonTab.aim = "ÐÐ Î´-áõÙ ÃÕÃ³Ïó³ÛÇÝ Ñ³ßíÇ ³Ùñ³óáõÙ"
        .commonTab.payer = "00034855"
        .commonTab.payerLegalStatus = "ֆիզԱնձ"
        .commonTab.name = "Ø»ñ"
        .commonTab.surname = "´³ÝÏÇ"
        .commonTab.id = "VF00034855"
        .commonTab.idForCheck = "CC00034855"
        .commonTab.idType = "13"
        .commonTab.idTypeForCheck = "05"
        .commonTab.idGivenBy = "013"
        .commonTab.idGivenByForCheck = "001"
        .commonTab.idGiveDate = "13032013"
        .commonTab.idGiveDateForCheck = "12/06/2013"
        .commonTab.idValidUntil = "13032023"
        .commonTab.idValidUntilForCheck = "12/06/2025"
        .commonTab.birthDate = "26051988"
        .commonTab.birthDateForCheck = "26/05/1988"
        .commonTab.citizenship = ""
        .commonTab.country = "AT"
        .commonTab.residence = "990000002"
        .commonTab.city = "ìÇ»Ý³"
        .commonTab.street = "15"
        .commonTab.apartment = "µÝ. 69    "
        .commonTab.house = "ïáõÝ 6    "
        .commonTab.email = "ViennaEmpares@mail.ru"
        .commonTab.emailForCheck = ""
        .chargeTab.office = .commonTab.office
        .chargeTab.department = .commonTab.department
        .chargeTab.chargeAcc = "72110253300 "
        .chargeTab.chargeAccForCheck = "000001100 "
        .chargeTab.chargeCurr = "002"
        .chargeTab.chargeCurrForCheck = "002"
        .chargeTab.cbExchangeRate = "782.9300/1"
        .chargeTab.chargeType = "01"
        .chargeTab.chargeAmount = "0.26"
        .chargeTab.chargeAmoForCheck = "0.26"
        .chargeTab.chargePercent = "0.1004"
        .chargeTab.chargePerForCheck = "0.1004"
        .chargeTab.incomeAcc = "000439300  "
        .chargeTab.incomeAccCurr = "000"
        .chargeTab.buyAndSell = "1"
        .chargeTab.buyAndSellForCheck = "1"
        .chargeTab.operType = "1"
        .chargeTab.operPlace = "3"
        .chargeTab.operArea = "6"
        .chargeTab.operAreaForCheck = "7"
        .chargeTab.nonResident = 0
        .chargeTab.nonResidentForCheck = 1
        .chargeTab.legalStatus = "33"
        .chargeTab.legalStatusForCheck = "21"
        .chargeTab.comment = "²ñï³ñÅ.ÙÇçí×. ·³ÝÓáõÙ"
        .chargeTab.commentForCheck = "²ñï³ñÅ.ÙÇçí×. ·³ÝÓáõÙ"
        .chargeTab.clientAgreeData = "üÇ½ÇÏ³Ï³Ý ³ÝÓ"
        .coinTab.coin = "0.00"
        .coinTab.coinForCheck = "43.25"
        .coinTab.coinPayCurr = "000"
        .coinTab.coinBuyAndSell = "1"
        .coinTab.coinPayAcc = "000001100  "
        .coinTab.coinExchangeRate = "782.9300/1"
        .coinTab.coinCBExchangeRate = "782.9300/1"
        .coinTab.coinPayAmount = "0.00"
        .coinTab.coinPayAmountForCheck = "0.00"
        .coinTab.amountWithMainCurr = "254.31"
        .coinTab.amountCurrForCheck = "254.31"
        .coinTab.incomeOutChange = "000931900  "
        .coinTab.damagesOutChange = "001434300  "
        .attachedTab.addFiles(0) = Project.Path & "Stores\Attach file\Photo.jpg"
        .attachedTab.fileName(0) = "Photo.jpg"
        .attachedTab.linkName(0) = "attachedLink_1"
        .attachedTab.addLinks(0) = Project.Path & "Stores\Attach file\Photo.jpg"
				End With
				
				Set workingDocs = New_MainAccWorkingDocuments()
				With workingDocs
								.startDate = cashOutputEdit.commonTab.date
								.endDate = cashOutputCreate.commonTab.date
				End With
				
				Set verifyDoc = New_VerificationDocument()
				verifyDoc.DocType = "KasRsOrd"
End Sub

Sub DB_Initialize()
    Dim i 
    For i = 0 To 1
        Set dbo_FOLDERS(i) = New_DB_FOLDERS()
        With dbo_FOLDERS(i)
            .fKEY = cashOutputCreate.fIsn
            .fISN = cashOutputCreate.fIsn
            .fNAME = "KasRsOrd"
            .fSTATUS = "5"
            .fCOM = "Î³ÝËÇÏ »Éù"
            .fECOM = "Cash Withdrawal Advice"
        End With
    Next
    With dbo_FOLDERS(0)
        .fFOLDERID = "C.123283164"
        .fSPEC = "²Ùë³ÃÇí- 22/02/22 N- " & cashOutputCreate.commonTab.docNum & " ¶áõÙ³ñ-             5,000.00 ²ñÅ.- 002 [Üáñ]"
    End With
    With dbo_FOLDERS(1)
        .fFOLDERID = "Oper.20220222"
        .fSPEC = cashOutputCreate.commonTab.docNum & "777007211025330077700000001102           5000.00002Üáñ                                                   77Ø»ñ ´³ÝÏÇ                       CC00034855 001 12/06/2013                              Ð        Ð³Ù³Ó³ÛÝ ·ñ³íÇ å³ÛÙ³Ý³·ñÇ Éñ³óÙ³Ý Ð³Ù³Ó³ÛÝ é»åá-Ñ³Ù³Ó³ÛÝ³·ñÇ                                                                                "
        .fDCBRANCH = "00 "
        .fDCDEPART = "1  "
    End With
    
    Set dbo_PAYMENTS = New_DB_PAYMENTS()
    With dbo_PAYMENTS
        .fDOCTYPE = "KasRsOrd"
        .fDATE = "2022-01-25"
        .fSTATE = "14"
        .fCLIENT = "00034855"
        .fACCDB = "7770072110253300"
        .fPAYER = "Ø»ñ ´³ÝÏÇ"
        .fCUR = "002"
        .fSUMMA = "254.31"
        .fSUMMAAMD = "199106.9283"
        .fSUMMAUSD = "497.7673"
        .fCOM = "ÐÐ Î´-áõÙ ÃÕÃ³Ïó³ÛÇÝ Ñ³ßíÇ ³Ùñ³óáõÙ Üáñ Ù»ÏÝ³µ³Ý.                                                                                           "
        .fPASSPORT = "VF00034855 013 13/03/2013"
        .fCOUNTRY = "AM"
        .fACSBRANCH = "00 "
        .fACSDEPART = "1  "
    End With
End Sub

Sub Check_DB_Create()
    'SQL Ստուգում DOCLOG աղուսյակում համար
    Log.Message "SQL Ստուգում DOCLOG աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCLOG", "fISN", cashOutputCreate.fIsn, 2)
    Call CheckDB_DOCLOG(cashOutputCreate.fIsn, "77", "N", "1", "", 1)
    Call CheckDB_DOCLOG(cashOutputCreate.fIsn, "77", "C", "2", "", 1)
  
    'SQL Ստուգում DOCS աղուսյակում 
    Log.Message "SQL Ստուգում DOCS աղուսյակում", "", pmNormal, SqlDivideColor
    fBODY = " ACSBRANCH:00 ACSDEPART:1 BLREP:0 OPERTYPE:MSC TYPECODE:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28 USERID:  77 DOCNUM:" & cashOutputCreate.commonTab.docNum & " DATE:20220222 ACCDB:72110253300 CUR:002 KASSA:001 ACCCR:000001102 SUMMA:5000 KASSIMV:051 BASE:Ð³Ù³Ó³ÛÝ é»åá-Ñ³Ù³Ó³ÛÝ³·ñÇ AIM:Ð³Ù³Ó³ÛÝ ·ñ³íÇ å³ÛÙ³Ý³·ñÇ Éñ³óÙ³Ý CLICODE:00034855 RECEIVER:Ø»ñ RECEIVERLASTNAME:´³ÝÏÇ PASSNUM:CC00034855 PASTYPE:05 PASBY:001 DATEPASS:20130612 DATEEXPIRE:20250612 ATEBIRTH:19880526 COUNTRY:AT COMMUNITY:990000002 CITY:ìÇ»Ý³ APARTMENT:µÝ. 69 ADDRESS:15 BUILDNUM:ïáõÝ 6 FROMPAYORD:0 ACSBRANCHINC:00 ACSDEPARTINC:1 CHRGACC:000001100 TYPECODE2:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28 CHRGCUR:000 CHRGCBCRS:1/1 PAYSCALE:TT CHRGINC:77796901013 FRSHNOCRG:0 CURTES:1 CURVAIR:3 VOLORT:7 NONREZ:1 JURSTAT:21 AGRDETAILS:üÇ½ÇÏ³Ï³Ý ³ÝÓ PAYSYSIN:Ð XSUM:43.25 XCUR:000 XACC:000001100 XDLCRS:782.9295/1 XDLCRSNAME:000 / 002 XCBCRS:782.9300/1 XCBCRSNAME:000 / 002 XCUPUSA:1 XCURSUM:33861.7 XSUMMAIN:4956.75 XINC:000931900 XEXP:001434300 NOTSENDABLE:0  "  
    fBODY = Replace(fBODY, " ", "%")
    Call CheckQueryRowCount("DOCS", "fISN", cashOutputCreate.fIsn, 1)
    Call CheckDB_DOCS(cashOutputCreate.fIsn, "KasRsOrd", "2", fBODY, 1)
  
    'SQL Ստուգում FOLDERS աղուսյակում 
    Log.Message "SQL Ստուգում FOLDERS աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("FOLDERS", "fISN", cashOutputCreate.fIsn, 2)
    Call CheckDB_FOLDERS(dbo_FOLDERS(0), 1)
    Call CheckDB_FOLDERS(dbo_FOLDERS(1), 1)
    
    'SQL Ստուգում HI աղուսյակում համար
    Log.Message "SQL Ստուգում HI աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HI", "fBASE", cashOutputCreate.fIsn, 4)
    Call Check_HI_CE_accounting ("2022-02-22", cashOutputCreate.fIsn, "11", "1630422", "3880788.30", "002", "4956.75", "MSC", "C")
    Call Check_HI_CE_accounting ("2022-02-22", cashOutputCreate.fIsn, "11", "1606020695", "3880788.30", "002", "4956.75", "MSC", "D")
    Call Check_HI_CE_accounting ("2022-02-22", cashOutputCreate.fIsn, "11", "1630170", "33861.70", "000", "33861.70", "CEX", "C")
    Call Check_HI_CE_accounting ("2022-02-22", cashOutputCreate.fIsn, "11", "1606020695", "33861.70", "002", "43.25", "CEX", "D")
End Sub

Sub Check_DB_Edit()
    'SQL Ստուգում DOCLOG աղուսյակում համար
    Log.Message "SQL Ստուգում DOCLOG աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCLOG", "fISN", cashOutputEdit.fIsn, 3)
    Call CheckDB_DOCLOG(cashOutputEdit.fIsn, "77", "E", "2", "", 1)
  
    'SQL Ստուգում DOCS աղուսյակում 
    Log.Message "SQL Ստուգում DOCS աղուսյակում", "", pmNormal, SqlDivideColor
    fBODY = " ACSBRANCH:00 ACSDEPART:1 BLREP:0 OPERTYPE:MSC TYPECODE:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28 USERID:  77 DOCNUM:" & cashOutputEdit.commonTab.docNum & " DATE:20220125 ACCDB:72110253300 CUR:002 KASSA:001 ACCCR:000001102 SUMMA:254.31 TOTAL:5000 KASSIMV:06 BASE:Üáñ Ù»ÏÝ³µ³Ý. AIM:ÐÐ Î´-áõÙ ÃÕÃ³Ïó³ÛÇÝ Ñ³ßíÇ ³Ùñ³óáõÙ CLICODE:00034855 RECEIVER:Ø»ñ RECEIVERLASTNAME:´³ÝÏÇ PASSNUM:VF00034855 PASTYPE:13 PASBY:013 DATEPASS:20130313 DATEEXPIRE:20230313 DATEBIRTH:19880526 COUNTRY:AT COMMUNITY:990000002 CITY:ìÇ»Ý³ APARTMENT:µÝ. 69 ADDRESS:15 BUILDNUM:ïáõÝ 6 EMAIL:ViennaEmpares@mail.ru FROMPAYORD:0 ACSBRANCHINC:00 ACSDEPARTINC:1 CHRGACC:72110253300 TYPECODE2:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28 CHRGCUR:002 CHRGCBCRS:782.9300/1 PAYSCALE:01 CHRGSUM:0.26 PRSNT:0.1004 CHRGINC:000439300 FRSHNOCRG:0 CUPUSA:1 CURTES:1 CURVAIR:3 VOLORT:6 NONREZ:0 JURSTAT:33 COMM:²ñï³ñÅ.ÙÇçí×. ·³ÝÓáõÙ AGRDETAILS:üÇ½ÇÏ³Ï³Ý ³ÝÓ PAYSYSIN:Ð XCUR:000 XACC:000001100 XDLCRS:782.9300/1 XDLCRSNAME:000 / 002 XCBCRS:782.9300/1 XCBCRSNAME:000 / 002 XCUPUSA:1 XSUMMAIN:254.31 XINC:000931900 XEXP:001434300 NOTSENDABLE:0 "   
    fBODY = Replace(fBODY, " ", "%")
    Call CheckQueryRowCount("DOCS", "fISN", cashOutputEdit.fIsn, 1)
    Call CheckDB_DOCS(cashOutputEdit.fIsn, "KasRsOrd", "2", fBODY, 1)
    
    'SQL Ստուգում DOCSATTACH աղուսյակում 
    Log.Message "SQL Ստուգում DOCSATTACH աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCSATTACH", "fISN", cashOutputEdit.fIsn, 2)
    Call CheckDB_DOCSATTACH(cashOutputEdit.fIsn, Project.Path &  "Stores\Attach file\Photo.jpg", "1", "attachedLink_1                                    ", 1)
    Call CheckDB_DOCSATTACH(cashOutputEdit.fIsn, "Photo.jpg", "0", "", 1)
  
    'SQL Ստուգում FOLDERS աղուսյակում 
    Log.Message "SQL Ստուգում FOLDERS աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("FOLDERS", "fISN", cashOutputEdit.fIsn, 2)
    dbo_FOLDERS(0).fSPEC = "²Ùë³ÃÇí- 25/01/22 N- " & cashOutputEdit.commonTab.docNum & " ¶áõÙ³ñ-               254.31 ²ñÅ.- 002 [ÊÙµ³·ñíáÕ]"
    Call CheckDB_FOLDERS(dbo_FOLDERS(0), 1)
    With dbo_FOLDERS(1)
        .fFOLDERID = "Oper.20220125"
        .fSPEC = cashOutputEdit.commonTab.docNum & "777007211025330077700000001102            254.31002ÊÙµ³·ñíáÕ                                             77Ø»ñ ´³ÝÏÇ                       VF00034855 013 13/03/2013                              Ð        ÐÐ Î´-áõÙ ÃÕÃ³Ïó³ÛÇÝ Ñ³ßíÇ ³Ùñ³óáõÙ Üáñ Ù»ÏÝ³µ³Ý.                                                                                           "
    End With
    Call CheckDB_FOLDERS(dbo_FOLDERS(1), 1)
    
    'SQL Ստուգում HI աղուսյակում համար
    Log.Message "SQL Ստուգում HI աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HI", "fBASE", cashOutputEdit.fIsn, 4)
    Call Check_HI_CE_accounting ("2022-01-25", cashOutputEdit.fIsn, "11", "1630422", "199106.90", "002", "254.31", "MSC", "C")
    Call Check_HI_CE_accounting ("2022-01-25", cashOutputEdit.fIsn, "11", "1606020695", "199106.90", "002", "254.31", "MSC", "D")
    Call Check_HI_CE_accounting ("2022-01-25", cashOutputEdit.fIsn, "11", "1629203", "203.60", "000", "203.60", "FEX", "C")
    Call Check_HI_CE_accounting ("2022-01-25", cashOutputEdit.fIsn, "11", "1606020695", "203.60", "002", "0.26", "FEX", "D")
End Sub

Sub Check_DB_SendToVerify()
    'SQL Ստուգում DOCLOG աղուսյակում համար
    Log.Message "SQL Ստուգում DOCLOG աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCLOG", "fISN", cashOutputEdit.fIsn, 4)
    Call CheckDB_DOCLOG(cashOutputEdit.fIsn, "77", "M", "11", "àõÕ³ñÏí»É ¿ ¹ñ³Ù³ñÏÕ", 1)
  
    'SQL Ստուգում DOCS աղուսյակում 
    Log.Message "SQL Ստուգում DOCS աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCS", "fISN", cashOutputEdit.fIsn, 1)
    Call CheckDB_DOCS(cashOutputEdit.fIsn, "KasRsOrd", "11", fBODY, 1)
  
    'SQL Ստուգում FOLDERS աղուսյակում 
    Log.Message "SQL Ստուգում FOLDERS աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("FOLDERS", "fISN", cashOutputEdit.fIsn, 3)
    With dbo_FOLDERS(0)
        .fSTATUS = "4"
        .fSPEC = "²Ùë³ÃÇí- 25/01/22 N- " & cashOutputEdit.commonTab.docNum & " ¶áõÙ³ñ-               254.31 ²ñÅ.- 002 [àõÕ³ñÏí»É ¿ ¹ñ³Ù³ñÏÕ]"
    End With
    Call CheckDB_FOLDERS(dbo_FOLDERS(0), 1)
    With dbo_FOLDERS(1)
        .fSTATUS = "4"
        .fFOLDERID = "CashOper.20220125"
        .fSPEC = cashOutputEdit.commonTab.docNum & "777007211025330077700000001102            254.31002àõÕ³ñÏí»É ¿ ¹ñ³Ù³ñÏÕ  77ÐÐ Î´-áõÙ ÃÕÃ³Ïó³ÛÇÝ Ñ³ßíÇ ³Ùñ³óØ»ñ ´³ÝÏÇ                       VF00034855 013 13/03/2013       001ºÉù            254.31002                 0.00   "
    End With
    Call CheckDB_FOLDERS(dbo_FOLDERS(1), 1)
    Set dbo_FOLDERS(2) = New_DB_FOLDERS()
    With dbo_FOLDERS(2)
        .fKEY = cashOutputEdit.fIsn
        .fISN = cashOutputEdit.fIsn
        .fNAME = "KasRsOrd"
        .fSTATUS = "4"
        .fFOLDERID = "Oper.20220125"
        .fCOM = "Î³ÝËÇÏ »Éù"
        .fSPEC = cashOutputEdit.commonTab.docNum & "777007211025330077700000001102            254.31002àõÕ³ñÏí»É ¿ ¹ñ³Ù³ñÏÕ                                  77Ø»ñ ´³ÝÏÇ                       VF00034855 013 13/03/2013       001                    Ð        ÐÐ Î´-áõÙ ÃÕÃ³Ïó³ÛÇÝ Ñ³ßíÇ ³Ùñ³óáõÙ Üáñ Ù»ÏÝ³µ³Ý.                                                                                           "
        .fECOM = "Cash Withdrawal Advice"
        .fDCBRANCH = "00 "
        .fDCDEPART = "1  "
    End With
    Call CheckDB_FOLDERS(dbo_FOLDERS(2), 1)
End Sub

Sub Check_DB_Validate()
    'SQL Ստուգում DOCLOG աղուսյակում համար
    Log.Message "SQL Ստուգում DOCLOG աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCLOG", "fISN", cashOutputEdit.fIsn, 6)
    Call CheckDB_DOCLOG(cashOutputEdit.fIsn, "77", "W", "12", "", 1)
    Call CheckDB_DOCLOG(cashOutputEdit.fIsn, "77", "M", "14", "¶ñ³Ýóí»É »Ý Ó¨³Ï»ñåáõÙÝ»ñÁ", 1)
  
    'SQL Ստուգում DOCS աղուսյակում 
    Log.Message "SQL Ստուգում DOCS աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCS", "fISN", cashOutputEdit.fIsn, 1)
    Call CheckDB_DOCS(cashOutputEdit.fIsn, "KasRsOrd", "14", fBODY, 1)
    
    'SQL Ստուգում FOLDERS աղուսյակում 
    Log.Message "SQL Ստուգում FOLDERS աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("FOLDERS", "fISN", cashOutputEdit.fIsn, 0)
  
    'SQL Ստուգում PAYMENTS աղուսյակում 
    Log.Message "SQL Ստուգում PAYMENTS աղուսյակում", "", pmNormal, SqlDivideColor
    With dbo_PAYMENTS
        .fISN = cashOutputEdit.fIsn
        .fDOCNUM = cashOutputEdit.commonTab.docNum
    End With
    Call CheckQueryRowCount("PAYMENTS", "fISN", cashOutputEdit.fIsn, 1)
    Call CheckDB_PAYMENTS(dbo_PAYMENTS, 1)
    
    'SQL Ստուգում HI աղուսյակում համար
    Log.Message "SQL Ստուգում HI աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HI", "fBASE", cashOutputEdit.fIsn, 5)
    Call Check_HI_CE_accounting ("2022-01-25", cashOutputEdit.fIsn, "01", "1630422", "199106.90", "002", "254.31", "MSC", "C")
    Call Check_HI_CE_accounting ("2022-01-25", cashOutputEdit.fIsn, "01", "1606020695", "199106.90", "002", "254.31", "MSC", "D")
    Call Check_HI_CE_accounting ("2022-01-25", cashOutputEdit.fIsn, "01", "1629203", "203.60", "000", "203.60", "FEX", "C")
    Call Check_HI_CE_accounting ("2022-01-25", cashOutputEdit.fIsn, "01", "1606020695", "203.60", "002", "0.26", "FEX", "D")
    Call Check_HI_CE_accounting ("2022-01-25", cashOutputEdit.fIsn, "CE", "1578251", "203.60", "002", "0.26", "PUR", "D")
End Sub

Sub Check_DB_Delete()
    'SQL Ստուգում DOCLOG աղուսյակում համար
    Log.Message "SQL Ստուգում DOCLOG աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCLOG", "fISN", cashOutputEdit.fIsn, 7)
    Call CheckDB_DOCLOG(cashOutputEdit.fIsn, "77", "D", "999", "", 1)
				
    'SQL Ստուգում DOCS աղուսյակում համար
    Log.Message "SQL Ստուգում DOCS աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCS", "fISN", cashOutputEdit.fIsn, 1)
    Call CheckDB_DOCS(cashOutputEdit.fIsn, "KasRsOrd", "999", fBODY, 1)
		
    'SQL Ստուգում FOLDERS աղուսյակում 
    Log.Message "SQL Ստուգում FOLDERS աղուսյակում", "", pmNormal, SqlDivideColor
    With dbo_FOLDERS(0)
        .fNAME = "KasRsOrd"
        .fSTATUS = "0"
        .fFOLDERID = ".R." & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%Y%m%d")
        .fSPEC = Left_Align(Get_Compname_DOCLOG(cashOutputEdit.fIsn), 16) & "GlavBux ARMSOFT                       1114 "
        .fCOM = ""
        .fECOM = ""
        .fDCBRANCH = "00 "
        .fDCDEPART = "1  "
    End With
    Call CheckQueryRowCount("FOLDERS", "fISN", cashOutputEdit.fIsn, 1)
    Call CheckDB_FOLDERS(dbo_FOLDERS(0), 1)
End Sub