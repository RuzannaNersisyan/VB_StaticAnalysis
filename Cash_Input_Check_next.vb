'USEUNIT Library_Common
'USEUNIT Library_Contracts
'USEUNIT Constants
'USEUNIT Library_Colour
'USEUNIT DAHK_Library_Filter
'USEUNIT CashInput_Confirmphases_Library
'USEUNIT Payment_Except_Library
'USEUNIT Akreditiv_Library
'USEUNIT Main_Accountant_Filter_Library
'USEUNIT SWIFT_International_Payorder_Library
'USEUNIT Library_CheckDB 
Option Explicit

'Test Case ID 183406

Dim sDate, eDate, folderName, expectedFile, expectedFileNext, actualFilePath, actualFile, actualFilePathNext, actualFileNext
Dim cashInputCreate, cashInputNextCreate, workingDocs, verifyDoc, currentDate, param
Dim fBODY, fBODYNext, dbo_FOLDERS(3), dbo_FOLDERSNext(3)

Sub Cash_Input_Check_Next_Test()
				Call Test_Inintialize()

				' Համակարգ մուտք գործել ARMSOFT օգտագործողով
				Log.Message "Համակարգ մուտք գործել ARMSOFT օգտագործողով", "", pmNormal, DivideColor
				Call Test_StartUp()
				
				' Ստեղծել Կանխիկ մուտք փաստաթուղթ
				Log.Message "Ստեղծել Կանխիկ մուտք փաստաթուղթ", "", pmNormal, DivideColor
    wTreeView.DblClickItem(folderName & "Î³ÝËÇÏ Ùáõïù")
    If wMDIClient.WaitvbObject("frmASDocForm", 3000).Exists Then 
        'ISN-ի վերագրում փոփոխականին
        cashInputCreate.fIsn = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
        Call Fill_CashInput(cashInputCreate, "Ð³çáñ¹Á")
        cashInputCreate.coinTab.coinForCheck = "0.15"
    Else 
        Log.Error "Can't open frmASDocForm window.", "", pmNormal, ErrorColor 
    End If
    
    ' Քաղվածքի պահպանում 
				Log.Message "Քաղվածքի պահպանում", "", pmNormal, DivideColor
    wMDIClient.vbObject("frmASDocForm").Keys("^[Tab]")
    Call SaveDoc(actualFilePath, actualFile) 
				
				' Փակել Քաղվածքի պատուհանը 
				Call Close_Window(wMDIClient, "FrmSpr")
    
    ' Կանխիկ մուտք փաստաթղթի ստեղծումից հետո SQL ստուգում
    Log.Message "Կանխիկ մուտք փաստաթղթի ստեղծումից հետո SQL ստուգում", "", pmNormal, DivideColor
    Call DB_Initialize()
    Call Check_DB_Create()
    
    ' Փաստացի քաղվածքի համեմատում սպասվողի հետ
				Log.Message "Փաստացի քաղվածքի համեմատում սպասվողի հետ", "", pmNormal, DivideColor
				param = "N\s\d{1,6}\s*.\d{1,10}\s{0,}.|Date\s\d{1,2}.\d{1,2}.\d{1,2}\s(\d{1,2}:\d{1,2})*"
    Call Compare_Files(actualFilePath & actualFile, expectedFile, param)
    
    ' Ստեղծել Կանխիկ մուտք փաստաթուղթ հաջորդից
    Log.Message "Ստեղծել Կանխիկ մուտք փաստաթուղթ հաջորդից", "", pmNormal, DivideColor
    wMDIClient.vbObject("frmExplorer").Keys("^[Tab]")
    If wMDIClient.WaitvbObject("frmASDocForm", 3000).Exists Then 
        'ISN-ի վերագրում փոփոխականին
        cashInputNextCreate.fIsn = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
        Call Fill_CashInput(cashInputNextCreate, "Î³ï³ñ»É")
    Else 
        Log.Error "Can't open frmASDocForm window.", "", pmNormal, ErrorColor 
    End If
    
				' Քաղվածքի պահպանում 
				Log.Message "Քաղվածքի պահպանում", "", pmNormal, DivideColor
    Call SaveDoc(actualFilePathNext, actualFileNext) 
				
				' Փակել Քաղվածքի պատուհանը 
				Call Close_Window(wMDIClient, "FrmSpr")
    
    ' Կանխիկ մուտք փաստաթուղթը հաջորդից ստեղծումից հետո SQL ստուգում
    Log.Message "Կանխիկ մուտք փաստաթուղթը հաջորդից ստեղծումից հետո SQL ստուգում", "", pmNormal, DivideColor
    Call Check_DB_Next_Create()
				
				' Փաստացի քաղվածքի համեմատում սպասվողի հետ
				Log.Message "Փաստացի քաղվածքի համեմատում սպասվողի հետ", "", pmNormal, DivideColor
    Call Compare_Files(actualFilePathNext & actualFileNext, expectedFileNext, param)
				
				' Մուտք գործել Աշխատանքային փաստաթղթեր թղթապանակ
    folderName = "|Ð³×³Ëáñ¹Ç ëå³ë³ñÏáõÙ ¨ ¹ñ³Ù³ñÏÕ |ÂÕÃ³å³Ý³ÏÝ»ñ|"
				Call GoTo_MainAccWorkingDocuments(folderName, workingDocs)
				
				' Ստուգել ստեղծված փաստաթղթի արժեքները
				Log.Message "Ստուգել ստեղծված փաստաթղթի արժեքները", "", pmNormal, DivideColor
    If SearchInPttel("frmPttel", 2, cashInputCreate.commonTab.docNum) Then
        wMDIClient.vbObject("frmPttel").Keys("^w")
        If wMDIClient.WaitVBObject("frmASDocForm", 3000).Exists Then
            Call Check_Cash_Input(cashInputCreate)
            Call ClickCmdButton(1, "OK") 
        Else 
            Log.Error "Can't open frmASDocForm window.", "", pmNormal, ErrorColor
        End If
    Else
        Log.Error "Can't find searched row.", "", pmNormal, ErrorColor
    End If
    
    ' Ուղարկել հաստատման
				Log.Message "Ուղարկել հաստատման", "", pmNormal, DivideColor
				Call SendToVerify_Contrct(3, 2, "Î³ï³ñ»É")
    
    ' Ուղարկել հաստատման-ից հետո SQL ստուգում
    Log.Message "Ուղարկել հաստատման-ից հետո SQL ստուգում", "", pmNormal, DivideColor
    Call Check_DB_SendToVerify()
    
    ' Մուտք գործել Աշխատանքային փաստաթղթեր թղթապանակ
    folderName = "|Ð³×³Ëáñ¹Ç ëå³ë³ñÏáõÙ ¨ ¹ñ³Ù³ñÏÕ |ÂÕÃ³å³Ý³ÏÝ»ñ|"
				Call GoTo_MainAccWorkingDocuments(folderName, workingDocs)
    
    ' Ստուգել հաջորդից ստեղծված փաստաթղթի արժեքները
				Log.Message "Ստուգել հաջորդից ստեղծված փաստաթղթի արժեքները", "", pmNormal, DivideColor
				If SearchInPttel("frmPttel", 2, cashInputNextCreate.commonTab.docNum) Then
        wMDIClient.vbObject("frmPttel").Keys("^w")
        If wMDIClient.WaitVBObject("frmASDocForm", 3000).Exists Then
            Call Check_Cash_Input(cashInputNextCreate)
            Call ClickCmdButton(1, "OK") 
        Else 
            Log.Error "Can't open frmASDocForm window.", "", pmNormal, ErrorColor
        End If
    Else
        Log.Error "Can't find searched row.", "", pmNormal, ErrorColor
    End If
				
				' Ուղարկել հաստատման
				Log.Message "Ուղարկել հաստատման", "", pmNormal, DivideColor
				Call SendToVerify_Contrct(3, 2, "Î³ï³ñ»É")
    
    ' Ուղարկել հաստատման-ից հետո SQL ստուգում
    Log.Message "Ուղարկել հաստատման-ից հետո SQL ստուգում", "", pmNormal, DivideColor
    Call Check_DB_SendToVerify_Next()
				
    ' Մուտք Գլխավոր հաշվապահի ԱՇՏ
				Call ChangeWorkspace(c_ChiefAcc)
    folderName = "|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|"
    
				' Մուտք գործել Հաստատվող փաստաթղթեր (|) թղթապանակ
				Call GoToVerificationDocument(folderName & "Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ (I)", verifyDoc)
				
				' Վավերացնել փաստաթղթերը
				Log.Message "Վավերացնել փաստաթղթերը", "", pmNormal, DivideColor
    If SearchInPttel("frmPttel", 1, cashInputCreate.fIsn) Then
    				Call Validate_Doc()
    Else
        Log.Error "Can't find searched row.", "", pmNormal, ErrorColor
    End If
    If SearchInPttel("frmPttel", 1, cashInputNextCreate.fIsn) Then
    				Call Validate_Doc()
    Else
        Log.Error "Can't find searched row.", "", pmNormal, ErrorColor
    End If
				
				' Փակել Հաստատվող փաստաթղթեր (|) թղթապանակը
				Call Close_Window(wMDIClient, "frmPttel")
    
    ' Վավերացումից հետո SQL ստուգում
    Log.Message "Վավերացումից հետո SQL ստուգում", "", pmNormal, DivideColor
    Call Check_DB_Validate()
				
				' Մուտք գործել Ստեղծված փաստաթղթեր թղթապանակ
				currentDate = aqConvert.DateTimeToFormatStr(aqDateTime.Now(),"%d%m%y")
				Call OpenCreatedDocFolder(folderName & "êï»ÕÍí³Í ÷³ëï³ÃÕÃ»ñ", currentDate, currentDate, null, "KasPrOrd")
				
				' Ստուգել, որ առկա են մեր ավելացրած փաստաթղթերը
				Log.Message "Ստուգել, որ առկա է մեր ավելացրած փաստաթղթերը", "", pmNormal, DivideColor
				If SearchInPttel("frmPttel", 2, cashInputCreate.fIsn) Then
    				wMDIClient.vbObject("frmPttel").Keys("^w")
        If wMDIClient.WaitVBObject("frmASDocForm", 3000).Exists Then
            Call Check_Cash_Input(cashInputCreate)
            Call ClickCmdButton(1, "OK") 
        Else 
            Log.Error "Can't open frmASDocForm window.", "", pmNormal, ErrorColor
        End If
    Else
        Log.Error "Can't find searched row.", "", pmNormal, ErrorColor
    End If
    
    If SearchInPttel("frmPttel", 2, cashInputNextCreate.fIsn) Then
    				wMDIClient.vbObject("frmPttel").Keys("^w")
        If wMDIClient.WaitVBObject("frmASDocForm", 3000).Exists Then
            Call Check_Cash_Input(cashInputNextCreate)
            Call ClickCmdButton(1, "OK") 
        Else 
            Log.Error "Can't open frmASDocForm window.", "", pmNormal, ErrorColor
        End If
    Else
        Log.Error "Can't find searched row.", "", pmNormal, ErrorColor
    End If
				
				' Ջնջել Կանխիկի մուտք փաստաթղթերը
				Log.Message "Ջնջել Կանխիկի մուտք փաստաթղթերը", "", pmNormal, DivideColor
				Call SearchAndDelete("frmPttel", 2, cashInputCreate.fIsn, "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ")
    Call SearchAndDelete("frmPttel", 2, cashInputNextCreate.fIsn, "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ")
				
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
				' Մուտք Հաճախորդի սպասարկում և դրամարկղ ընդլայնված ԱՇՏ
				Call ChangeWorkspace(c_CustomerService)
End Sub

Sub Test_Inintialize()
				sDate = "20030101"
				eDate = "20250101"
		
				folderName = "|Ð³×³Ëáñ¹Ç ëå³ë³ñÏáõÙ ¨ ¹ñ³Ù³ñÏÕ |Üáñ ÷³ëï³ÃÕÃ»ñ|ì×³ñ³ÛÇÝ ÷³ëï³ÃÕÃ»ñ|"
    
				expectedFile = Project.Path &  "Stores\Cash_Input_Output\Expected\Expected_Cash_Input_Next_Create.txt"
				actualFilePath = Project.Path &  "Stores\Cash_Input_Output\Actual\"
    actualFile = "Actual_Cash_Input_Next_Create.txt"
    
    expectedFileNext = Project.Path &  "Stores\Cash_Input_Output\Expected\Expected_Cash_Input_Next_Create2.txt"
				actualFilePathNext = Project.Path &  "Stores\Cash_Input_Output\Actual\"
    actualFileNext = "Actual_Cash_Input_Next_Create2.txt"
		
				Set cashInputCreate = New_CashInput(1, 1, 0)
				With cashInputCreate
								.commonTab.office = "00"
        .commonTab.department = "1"
        .commonTab.date = "080222"
        .commonTab.dateForCheck = "08/02/22"
        .commonTab.cashRegister = "001"
        .commonTab.cashRegisterAcc = "77798311011"
        .commonTab.curr = "003"
        .commonTab.accCredit = "72110175100"
        .commonTab.amount = "26543.12"
        .commonTab.amountForCheck = "26,543.12"
        .commonTab.cashierChar = "04 "
        .commonTab.base = "ØÇçÝáñ¹³í×³ñ"
        .commonTab.aim = "²ßË³ï³í³ñÓÇ í×³ñáõÙ"
        .commonTab.payer = "00034848"
        .commonTab.name = "²½³ï³ÝÇ"
        .commonTab.surname = "·ÛáõÕ³å»ï³ñ³Ý"
        .commonTab.id = "AG00034848"
		      .commonTab.idForCheck = "AG00034848"
        .commonTab.idType = "06"
		      .commonTab.idTypeForCheck = "06"
        .commonTab.idGivenBy = "045"
		      .commonTab.idGivenByForCheck = "045"
        .commonTab.idGiveDate = "12062006"
        .commonTab.idGiveDateForCheck = "12/06/2006"
        .commonTab.idValidUntil = "25042016"
        .commonTab.idValidUntilForCheck = "25/04/2016"
        .commonTab.birthDate = "19101988"
        .commonTab.birthDateForCheck = "19/10/1988"
        .commonTab.citizenship = "4"
        .commonTab.country = "DK"
        .commonTab.residence = "990000002"
        .commonTab.city = "Îáå»ÝÑ³·»Ý"
        .commonTab.street = "Ð³ÛïÝÇÝ»ñÇ ÷."
        .commonTab.apartment = "µÝ. 26    "
        .commonTab.house = "ß»Ýù 8    "
        .commonTab.email = "personFromDenmark@person.dk"
		      .commonTab.emailForCheck = "personFromDenmark@person.dk"
        .chargeTab.office = .commonTab.office
        .chargeTab.department = .commonTab.department
        .chargeTab.chargeAcc = "000001101 "
        .chargeTab.chargeAccForCheck = "000001100 "
        .chargeTab.chargeCurr = "001"
        .chargeTab.chargeCurrForCheck = "001"
        .chargeTab.cbExchangeRate = "400.0000/1"
        .chargeTab.chargeType = "01"
        .chargeTab.chargeAmount = "29.86"
        .chargeTab.chargeAmoForCheck = "29.86"
        .chargeTab.chargePercent = "0.1000"
        .chargeTab.chargePerForCheck = "0.1000"
        .chargeTab.incomeAcc = "00000451000"
        .chargeTab.incomeAccCurr = "000"
        .chargeTab.buyAndSell = "1"
        .chargeTab.buyAndSellForCheck = "1"
        .chargeTab.operType = "1"
        .chargeTab.operPlace = "3"
        .chargeTab.operArea = "9X"
        .chargeTab.operAreaForCheck = "9X"
        .chargeTab.nonResident = 0
        .chargeTab.nonResidentForCheck = 0
        .chargeTab.legalStatus = "32"
        .chargeTab.legalStatusForCheck = "32"
        .chargeTab.comment = "²ñï³ñÅ.ÙÇçí×. ·³ÝÓáõÙ"
        .chargeTab.commentForCheck = "²ñï³ñÅ.ÙÇçí×. ·³ÝÓáõÙ"
        .chargeTab.clientAgreeData = "Ø»ñ ëÇñ»ÉÇ Ñ³×³Ëáñ¹Ç ïíÛ³ÉÝ»ñ"
        .coinTab.coin = "0.15"
        .coinTab.coinForCheck = "0.00"
        .coinTab.coinPayCurr = "000"
        .coinTab.coinBuyAndSell = "2"
        .coinTab.coinPayAcc = "000001100  "
        .coinTab.coinExchangeRate = "535.3333/1"
        .coinTab.coinCBExchangeRate = "450.0000/1"
        .coinTab.coinPayAmount = "80.30"
        .coinTab.coinPayAmountForCheck = "80.30"
        .coinTab.amountWithMainCurr = "26,542.97"
        .coinTab.amountCurrForCheck = "26,542.97"
        .coinTab.incomeOutChange = "000931900  "
        .coinTab.damagesOutChange = "001434300  "
        .attachedTab.addFiles(0) = Project.Path & "Stores\Attach file\Photo.jpg"
        .attachedTab.fileName(0) = "Photo.jpg"
        .attachedTab.linkName(0) = "attachedLink_1"
        .attachedTab.addLinks(0) = Project.Path & "Stores\Attach file\Photo.jpg"
				End With
				
    Set cashInputNextCreate = New_CashInput(0, 0, 0)
				With cashInputNextCreate
        .commonTab.office = "00"
        .commonTab.department = "1"
        .commonTab.date = "080222"
        .commonTab.dateForCheck = "08/02/22"
        .commonTab.cashRegister = "001"
        .commonTab.cashRegisterAcc = "000001100  "
        .commonTab.curr = "000"
        .commonTab.accCredit = "72110177100"
        .commonTab.amount = "48736.49"
        .commonTab.amountForCheck = "48,736.50"
        .commonTab.cashierChar = "031"
        .commonTab.base = "ÇÝÏ³ë³ïáñÇ íÏ³Û³Ï³Ý"
        .commonTab.aim = "Ð³Ù³Ó³ÛÝ é»åá-Ñ³Ù³Ó³ÛÝ³·ñÇ"
        .commonTab.payer = "00034854"
        .commonTab.name = "àã å»ï³Ï³Ý"
        .commonTab.surname = "ÑÇÙÝ³ñÏ"
        .chargeTab.office = .commonTab.office
        .chargeTab.department = .commonTab.department
        .chargeTab.chargeAcc = "72110177100"
        .chargeTab.chargeAccForCheck = "72110177100"
        .chargeTab.chargeCurr = "000"
        .chargeTab.chargeCurrForCheck = "000"
        .chargeTab.cbExchangeRate = "1.0000/1"
        .chargeTab.operArea = "9X"
        .chargeTab.operAreaForCheck = "9X"
        .chargeTab.legalStatus = "41"
        .chargeTab.legalStatusForCheck = "41"
				End With
				
				Set workingDocs = New_MainAccWorkingDocuments()
				With workingDocs
								.startDate = cashInputCreate.commonTab.date
								.endDate = cashInputNextCreate.commonTab.date
				End With
				
				Set verifyDoc = New_VerificationDocument()
				verifyDoc.DocType = "KasPrOrd"
End Sub

Sub DB_Initialize()		
    Dim i 
    For i = 0 To 1
        Set dbo_FOLDERS(i) = New_DB_FOLDERS()
        With dbo_FOLDERS(i)
            .fKEY = cashInputCreate.fIsn
            .fISN = cashInputCreate.fIsn
            .fNAME = "KasPrOrd"
            .fSTATUS = "5"
            .fCOM = "Î³ÝËÇÏ Ùáõïù"
            .fECOM = "Cash Deposit Advice"
        End With
    Next
    With dbo_FOLDERS(0)
        .fFOLDERID = "C.1012283530"
        .fSPEC = "²Ùë³ÃÇí- 08/02/22 N- " & cashInputCreate.commonTab.docNum & " ¶áõÙ³ñ-            26,543.12 ²ñÅ.- 003 [Üáñ]"
    End With
    With dbo_FOLDERS(1)
        .fFOLDERID = "Oper.20220208"
        .fSPEC = cashInputCreate.commonTab.docNum & "77700777983110117770072110175100        26543.12003Üáñ                                                   77²½³ï³ÝÇ ·ÛáõÕ³å»ï³ñ³Ý           AG00034848 045 12/06/2006                                       ²ßË³ï³í³ñÓÇ í×³ñáõÙ ØÇçÝáñ¹³í×³ñ                                                                                                            "
        .fDCBRANCH = "00 "
        .fDCDEPART = "1  "
    End With
    
    For i = 0 To 1
        Set dbo_FOLDERSNext(i) = New_DB_FOLDERS()
        With dbo_FOLDERSNext(i)
            .fNAME = "KasPrOrd"
            .fSTATUS = "5"
            .fCOM = "Î³ÝËÇÏ Ùáõïù"
            .fECOM = "Cash Deposit Advice"
        End With
    Next
    dbo_FOLDERSNext(0).fFOLDERID = "C.566887393"
    With dbo_FOLDERSNext(1)
        .fFOLDERID = "Oper.20220208"
        .fDCBRANCH = "00 "
        .fDCDEPART = "1  "
    End With
End	Sub

Sub Check_DB_Create()
    'SQL Ստուգում DOCLOG աղուսյակում համար
    Log.Message "SQL Ստուգում DOCLOG աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCLOG", "fISN", cashInputCreate.fIsn, 2)
    Call CheckDB_DOCLOG(cashInputCreate.fIsn, "77", "N", "1", "", 1)
    Call CheckDB_DOCLOG(cashInputCreate.fIsn, "77", "C", "2", "", 1)
  
    'SQL Ստուգում DOCS աղուսյակում 
    Log.Message "SQL Ստուգում DOCS աղուսյակում", "", pmNormal, SqlDivideColor
    fBODY = " ACSBRANCH:00 ACSDEPART:1 BLREP:0 OPERTYPE:MSC TYPECODE:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28 USERID:  77 DOCNUM:" & cashInputCreate.commonTab.docNum & " DATE:20220208 KASSA:001 ACCDB:77798311011 CUR:003 ACCCR:72110175100 SUMMA:26543.12 KASSIMV:04 BASE:ØÇçÝáñ¹³í×³ñ AIM:²ßË³ï³í³ñÓÇ í×³ñáõÙ CLICODE:00034848 PAYER:²½³ï³ÝÇ PAYERLASTNAME:·ÛáõÕ³å»ï³ñ³Ý PASSNUM:AG00034848 PASTYPE:06 PASBY:045 DATEPASS:20060612 DATEEXPIRE:20160425 DATEBIRTH:19881019 CITIZENSHIP:4 COUNTRY:DK COMMUNITY:990000002 CITY:Îáå»ÝÑ³·»Ý APARTMENT:µÝ. 26 ADDRESS:Ð³ÛïÝÇÝ»ñÇ ÷. BUILDNUM:ß»Ýù 8 EMAIL:personFromDenmark@person.dk ACSBRANCHINC:00 ACSDEPARTINC:1 CHRGACC:000001101 TYPECODE2:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28 CHRGCUR:001 CHRGCBCRS:400.0000/1 PAYSCALE:01 CHRGSUM:29.86 PRSNT:0.1 CHRGINC:00000451000 CUPUSA:1 CURTES:1 CURVAIR:3 VOLORT:9X NONREZ:0 JURSTAT:32 COMM:²ñï³ñÅ.ÙÇçí×. ·³ÝÓáõÙ AGRDETAILS:Ø»ñ ëÇñ»ÉÇ Ñ³×³Ëáñ¹Ç ïíÛ³ÉÝ»ñ XSUM:0.15 XCUR:000 XACC:000001100 XDLCRS:535.3333/1 XDLCRSNAME:000 / 003 XCBCRS:450.0000/1 XCBCRSNAME:000 / 003 XCUPUSA:2 XCURSUM:80.3 XSUMMAIN:26542.97 XINC:000931900 XEXP:001434300 USEOVERLIMIT:0 NOTSENDABLE:0  " 
    fBODY = Replace(fBODY, " ", "%")
    Call CheckQueryRowCount("DOCS", "fISN", cashInputCreate.fIsn, 1)
    Call CheckDB_DOCS(cashInputCreate.fIsn, "KasPrOrd", "2", fBODY, 1)
    
    'SQL Ստուգում DOCSATTACH աղուսյակում 
    Log.Message "SQL Ստուգում DOCSATTACH աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCSATTACH", "fISN", cashInputCreate.fIsn, 2)
    Call CheckDB_DOCSATTACH(cashInputCreate.fIsn, Project.Path & "Stores\Attach file\Photo.jpg", "1", "attachedLink_1                                    ", 1)
    Call CheckDB_DOCSATTACH(cashInputCreate.fIsn, "Photo.jpg", "0", "", 1)    
  
    'SQL Ստուգում FOLDERS աղուսյակում 
    Log.Message "SQL Ստուգում FOLDERS աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("FOLDERS", "fISN", cashInputCreate.fIsn, 2)
    Call CheckDB_FOLDERS(dbo_FOLDERS(0), 1)
    Call CheckDB_FOLDERS(dbo_FOLDERS(1), 1)
  
    'SQL Ստուգում HI աղուսյակում համար
    Log.Message "SQL Ստուգում HI աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HI", "fBASE", cashInputCreate.fIsn, 8)
    Call Check_HI_CE_accounting ("2022-02-08", cashInputCreate.fIsn, "11", "1757142", "11944336.50", "003", "26542.97", "MSC", "D")
    Call Check_HI_CE_accounting ("2022-02-08", cashInputCreate.fIsn, "11", "1093631245", "11944336.50", "003", "26542.97", "MSC", "C")
    Call Check_HI_CE_accounting ("2022-02-08", cashInputCreate.fIsn, "11", "1629177", "12.80", "000", "12.80", "MSC", "C")
    Call Check_HI_CE_accounting ("2022-02-08", cashInputCreate.fIsn, "11", "1093631245", "12.80", "003", "0.00", "MSC", "D")
    Call Check_HI_CE_accounting ("2022-02-08", cashInputCreate.fIsn, "11", "1630170", "80.30", "000", "80.30", "CEX", "D")
    Call Check_HI_CE_accounting ("2022-02-08", cashInputCreate.fIsn, "11", "1093631245", "80.30", "003", "0.15", "CEX", "C")
    Call Check_HI_CE_accounting ("2022-02-08", cashInputCreate.fIsn, "11", "1630171", "11944.00", "001", "29.86", "FEX", "D")
    Call Check_HI_CE_accounting ("2022-02-08", cashInputCreate.fIsn, "11", "354210522", "11944.00", "000", "11944.00", "FEX", "C")
End Sub

Sub Check_DB_Next_Create()
    'SQL Ստուգում DOCLOG աղուսյակում համար
    Log.Message "SQL Ստուգում DOCLOG աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCLOG", "fISN", cashInputNextCreate.fIsn, 2)
    Call CheckDB_DOCLOG(cashInputNextCreate.fIsn, "77", "N", "1", "", 1)
    Call CheckDB_DOCLOG(cashInputNextCreate.fIsn, "77", "C", "2", "", 1)
  
    'SQL Ստուգում DOCS աղուսյակում 
    Log.Message "SQL Ստուգում DOCS աղուսյակում", "", pmNormal, SqlDivideColor
    fBODYNext = " ACSBRANCH:00 ACSDEPART:1 BLREP:0 OPERTYPE:MSC TYPECODE:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28 USERID:  77 DOCNUM:"& cashInputNextCreate.commonTab.docNum & " DATE:20220208 KASSA:001 ACCDB:000001100 CUR:000 ACCCR:72110177100 SUMMA:48736.5 KASSIMV:031 BASE:ÇÝÏ³ë³ïáñÇ íÏ³Û³Ï³Ý AIM:Ð³Ù³Ó³ÛÝ é»åá-Ñ³Ù³Ó³ÛÝ³·ñÇ CLICODE:00034854 PAYER:àã å»ï³Ï³Ý PAYERLASTNAME:ÑÇÙÝ³ñÏ ACSBRANCHINC:00 ACSDEPARTINC:1 CHRGACC:72110177100 TYPECODE2:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28 CHRGCUR:000 CHRGCBCRS:1/1 VOLORT:9X NONREZ:0 JURSTAT:41  XDLCRSNAME:000 / 003 XCBCRSNAME:000 / 003 USEOVERLIMIT:0 NOTSENDABLE:0  "
    fBODYNext = Replace(fBODYNext, " ", "%")
    Call CheckQueryRowCount("DOCS", "fISN", cashInputNextCreate.fIsn, 1)
    Call CheckDB_DOCS(cashInputNextCreate.fIsn, "KasPrOrd", "2", fBODYNext, 1) 
  
    'SQL Ստուգում FOLDERS աղուսյակում 
    Log.Message "SQL Ստուգում FOLDERS աղուսյակում", "", pmNormal, SqlDivideColor
    With dbo_FOLDERSNext(0)
        .fKEY = cashInputNextCreate.fIsn
        .fISN = cashInputNextCreate.fIsn
        .fSPEC = "²Ùë³ÃÇí- 08/02/22 N- " & cashInputNextCreate.commonTab.docNum & " ¶áõÙ³ñ-            48,736.50 ²ñÅ.- 000 [Üáñ]"
    End With
    With dbo_FOLDERSNext(1)
        .fKEY = cashInputNextCreate.fIsn
        .fISN = cashInputNextCreate.fIsn
        .fSPEC = cashInputNextCreate.commonTab.docNum & "77700000001100  7770072110177100        48736.50000Üáñ                                                   77àã å»ï³Ï³Ý ÑÇÙÝ³ñÏ                                                                              Ð³Ù³Ó³ÛÝ é»åá-Ñ³Ù³Ó³ÛÝ³·ñÇ ÇÝÏ³ë³ïáñÇ íÏ³Û³Ï³Ý                                                                                              "
    End With
    Call CheckQueryRowCount("FOLDERS", "fISN", cashInputNextCreate.fIsn, 2)
    Call CheckDB_FOLDERS(dbo_FOLDERSNext(0), 1)
    Call CheckDB_FOLDERS(dbo_FOLDERSNext(1), 1)
  
    'SQL Ստուգում HI աղուսյակում համար
    Log.Message "SQL Ստուգում HI աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HI", "fBASE", cashInputNextCreate.fIsn, 2)
    Call Check_HI_CE_accounting ("2022-02-08", cashInputNextCreate.fIsn, "11", "1630170", "48736.50", "000", "48736.50", "MSC", "D")
    Call Check_HI_CE_accounting ("2022-02-08", cashInputNextCreate.fIsn, "11", "441594693", "48736.50", "000", "48736.50", "MSC", "C")
End Sub

Sub Check_DB_SendToVerify()
    'SQL Ստուգում DOCLOG աղուսյակում համար
    Log.Message "SQL Ստուգում DOCLOG աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCLOG", "fISN", cashInputCreate.fIsn, 3)
    Call CheckDB_DOCLOG(cashInputCreate.fIsn, "77", "M", "101", "àõÕ³ñÏí»É ¿ Ñ³ëï³ïÙ³Ý", 1)
  
    'SQL Ստուգում DOCS աղուսյակում 
    Log.Message "SQL Ստուգում DOCS աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCS", "fISN", cashInputCreate.fIsn, 1)
    Call CheckDB_DOCS(cashInputCreate.fIsn, "KasPrOrd", "101", fBODY, 1)
  
    'SQL Ստուգում FOLDERS աղուսյակում 
    Log.Message "SQL Ստուգում FOLDERS աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("FOLDERS", "fISN", cashInputCreate.fIsn, 3)
    With dbo_FOLDERS(0)
        .fSTATUS = "0"
        .fSPEC = "²Ùë³ÃÇí- 08/02/22 N- " & cashInputCreate.commonTab.docNum & " ¶áõÙ³ñ-            26,543.12 ²ñÅ.- 003 [àõÕ³ñÏí³Í I Ñ³ëï³ïÙ³Ý]"
    End With
    Call CheckDB_FOLDERS(dbo_FOLDERS(0), 1)
    With dbo_FOLDERS(1)
        .fSTATUS = "0"
        .fSPEC = cashInputCreate.commonTab.docNum & "77700777983110117770072110175100        26543.12003àõÕ³ñÏí³Í I Ñ³ëï³ïÙ³Ý                                 77²½³ï³ÝÇ ·ÛáõÕ³å»ï³ñ³Ý                                           001                             ²ßË³ï³í³ñÓÇ í×³ñáõÙ ØÇçÝáñ¹³í×³ñ                                                                                                            "
    End With
    Call CheckDB_FOLDERS(dbo_FOLDERS(1), 1)
    Set dbo_FOLDERS(2) = New_DB_FOLDERS()
    With dbo_FOLDERS(2)
        .fKEY = cashInputCreate.fIsn
        .fISN = cashInputCreate.fIsn
        .fNAME = "KasPrOrd"
        .fSTATUS = "4"
        .fFOLDERID = "Ver.20220208001"
        .fCOM = "Î³ÝËÇÏ Ùáõïù"
        .fSPEC = cashInputCreate.commonTab.docNum & "77700777983110117770072110175100        26543.12003  77²ßË³ï³í³ñÓÇ í×³ñáõÙ             ØÇçÝáñ¹³í×³ñ                    ²½³ï³ÝÇ ·ÛáõÕ³å»ï³ñ³Ý           "
        .fECOM = "Cash Deposit Advice"
        .fDCBRANCH = "00 "
        .fDCDEPART = "1  "
    End With
    Call CheckDB_FOLDERS(dbo_FOLDERS(2), 1)
End Sub

Sub Check_DB_SendToVerify_Next()
    'SQL Ստուգում DOCLOG աղուսյակում համար
    Log.Message "SQL Ստուգում DOCLOG աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCLOG", "fISN", cashInputNextCreate.fIsn, 3)
    Call CheckDB_DOCLOG(cashInputNextCreate.fIsn, "77", "M", "101", "àõÕ³ñÏí»É ¿ Ñ³ëï³ïÙ³Ý", 1)
  
    'SQL Ստուգում DOCS աղուսյակում 
    Log.Message "SQL Ստուգում DOCS աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCS", "fISN", cashInputNextCreate.fIsn, 1)
    Call CheckDB_DOCS(cashInputNextCreate.fIsn, "KasPrOrd", "101", fBODYNext, 1)
  
    'SQL Ստուգում FOLDERS աղուսյակում 
    Log.Message "SQL Ստուգում FOLDERS աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("FOLDERS", "fISN", cashInputNextCreate.fIsn, 3)
    With dbo_FOLDERSNext(0)
        .fSTATUS = "0"
        .fSPEC = "²Ùë³ÃÇí- 08/02/22 N- " & cashInputNextCreate.commonTab.docNum & " ¶áõÙ³ñ-            48,736.50 ²ñÅ.- 000 [àõÕ³ñÏí³Í I Ñ³ëï³ïÙ³Ý]"
    End With
    Call CheckDB_FOLDERS(dbo_FOLDERSNext(0), 1)
    With dbo_FOLDERSNext(1)
        .fSTATUS = "0"
        .fSPEC = cashInputNextCreate.commonTab.docNum & "77700000001100  7770072110177100        48736.50000àõÕ³ñÏí³Í I Ñ³ëï³ïÙ³Ý                                 77àã å»ï³Ï³Ý ÑÇÙÝ³ñÏ                                              001                             Ð³Ù³Ó³ÛÝ é»åá-Ñ³Ù³Ó³ÛÝ³·ñÇ ÇÝÏ³ë³ïáñÇ íÏ³Û³Ï³Ý                                                                                              "
    End With
    Call CheckDB_FOLDERS(dbo_FOLDERSNext(1), 1)
    Set dbo_FOLDERSNext(2) = New_DB_FOLDERS()
    With dbo_FOLDERSNext(2)
        .fKEY = cashInputNextCreate.fIsn
        .fISN = cashInputNextCreate.fIsn
        .fNAME = "KasPrOrd"
        .fSTATUS = "4"
        .fFOLDERID = "Ver.20220208001"
        .fCOM = "Î³ÝËÇÏ Ùáõïù"
        .fSPEC = cashInputNextCreate.commonTab.docNum & "77700000001100  7770072110177100        48736.50000  77Ð³Ù³Ó³ÛÝ é»åá-Ñ³Ù³Ó³ÛÝ³·ñÇ      ÇÝÏ³ë³ïáñÇ íÏ³Û³Ï³Ý             àã å»ï³Ï³Ý ÑÇÙÝ³ñÏ              "
        .fECOM = "Cash Deposit Advice"
        .fDCBRANCH = "00 "
        .fDCDEPART = "1  "
    End With
    Call CheckDB_FOLDERS(dbo_FOLDERSNext(2), 1)
End Sub

Sub Check_DB_Validate()
    'SQL Ստուգում DOCLOG աղուսյակում համար
    Log.Message "SQL Ստուգում DOCLOG աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCLOG", "fISN", cashInputCreate.fIsn, 5)
    Call CheckDB_DOCLOG(cashInputCreate.fIsn, "77", "W", "102", "", 1)
    Call CheckDB_DOCLOG(cashInputCreate.fIsn, "77", "C", "15", "", 1)
    Call CheckQueryRowCount("DOCLOG", "fISN", cashInputNextCreate.fIsn, 5)
    Call CheckDB_DOCLOG(cashInputNextCreate.fIsn, "77", "W", "102", "", 1)
    Call CheckDB_DOCLOG(cashInputNextCreate.fIsn, "77", "C", "15", "", 1)
  
    'SQL Ստուգում DOCS աղուսյակում 
    Log.Message "SQL Ստուգում DOCS աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCS", "fISN", cashInputCreate.fIsn, 1)
    Call CheckDB_DOCS(cashInputCreate.fIsn, "KasPrOrd", "15", fBODY, 1)
    Call CheckQueryRowCount("DOCS", "fISN", cashInputNextCreate.fIsn, 1)
    Call CheckDB_DOCS(cashInputNextCreate.fIsn, "KasPrOrd", "15", fBODYNext, 1)
  
    'SQL Ստուգում FOLDERS աղուսյակում 
    Log.Message "SQL Ստուգում FOLDERS աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("FOLDERS", "fISN", cashInputCreate.fIsn, 2)
    With dbo_FOLDERS(0)
        .fSTATUS = "4"
        .fSPEC = "²Ùë³ÃÇí- 08/02/22 N- " & cashInputCreate.commonTab.docNum & " ¶áõÙ³ñ-            26,543.12 ²ñÅ.- 003 [Ð³ëï³ïí³Í]"
    End With
    Call CheckDB_FOLDERS(dbo_FOLDERS(0), 1)
    With dbo_FOLDERS(1)
        .fFOLDERID = "Oper.20220208"
        .fSTATUS = "4"
        .fSPEC = cashInputCreate.commonTab.docNum & "77700777983110117770072110175100        26543.12003Ð³ëï³ïí³Í                                             77²½³ï³ÝÇ ·ÛáõÕ³å»ï³ñ³Ý           AG00034848 045 12/06/2006                                       ²ßË³ï³í³ñÓÇ í×³ñáõÙ ØÇçÝáñ¹³í×³ñ                                                                                                            "
    End With
    Call CheckDB_FOLDERS(dbo_FOLDERS(1), 1)
    Call CheckQueryRowCount("FOLDERS", "fISN", cashInputNextCreate.fIsn, 2)
    With dbo_FOLDERSNext(0)
        .fSTATUS = "4"
        .fSPEC = "²Ùë³ÃÇí- 08/02/22 N- " & cashInputNextCreate.commonTab.docNum & " ¶áõÙ³ñ-            48,736.50 ²ñÅ.- 000 [Ð³ëï³ïí³Í]"
    End With
    Call CheckDB_FOLDERS(dbo_FOLDERSNext(0), 1)
    With dbo_FOLDERSNext(1)
        .fFOLDERID = "Oper.20220208"
        .fSTATUS = "4"
        .fSPEC = cashInputNextCreate.commonTab.docNum & "77700000001100  7770072110177100        48736.50000Ð³ëï³ïí³Í                                             77àã å»ï³Ï³Ý ÑÇÙÝ³ñÏ                                                                              Ð³Ù³Ó³ÛÝ é»åá-Ñ³Ù³Ó³ÛÝ³·ñÇ ÇÝÏ³ë³ïáñÇ íÏ³Û³Ï³Ý                                                                                              "
    End With
    Call CheckDB_FOLDERS(dbo_FOLDERSNext(1), 1)
End Sub

Sub Check_DB_Delete()
    'SQL Ստուգում DOCLOG աղուսյակում համար
    Log.Message "SQL Ստուգում DOCLOG աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCLOG", "fISN", cashInputCreate.fIsn, 6)
    Call CheckDB_DOCLOG(cashInputCreate.fIsn, "77", "D", "999", "", 1)
    Call CheckQueryRowCount("DOCLOG", "fISN", cashInputNextCreate.fIsn, 6)
    Call CheckDB_DOCLOG(cashInputNextCreate.fIsn, "77", "D", "999", "", 1)
				
    'SQL Ստուգում DOCS աղուսյակում համար
    Log.Message "SQL Ստուգում DOCS աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCS", "fISN", cashInputCreate.fIsn, 1)
    Call CheckDB_DOCS(cashInputCreate.fIsn, "KasPrOrd", "999", fBODY, 1)
    Call CheckQueryRowCount("DOCS", "fISN", cashInputNextCreate.fIsn, 1)
    Call CheckDB_DOCS(cashInputNextCreate.fIsn, "KasPrOrd", "999", fBODYNext, 1)
		
    'SQL Ստուգում FOLDERS աղուսյակում 
    Log.Message "SQL Ստուգում FOLDERS աղուսյակում", "", pmNormal, SqlDivideColor
    With dbo_FOLDERS(0)
        .fKEY = cashInputCreate.fIsn
        .fISN = cashInputCreate.fIsn
        .fNAME = "KasPrOrd"
        .fSTATUS = "0"
        .fFOLDERID = ".R." & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%Y%m%d")
        .fSPEC = Left_Align(Get_Compname_DOCLOG(cashInputCreate.fIsn), 16) & "GlavBux ARMSOFT                       0115"
        .fCOM = ""
        .fECOM = ""
        .fDCBRANCH = "00 "
        .fDCDEPART = "1  "
    End With
    Call CheckQueryRowCount("FOLDERS", "fISN", cashInputCreate.fIsn, 1)
    Call CheckDB_FOLDERS(dbo_FOLDERS(0), 1)
    Call CheckQueryRowCount("FOLDERS", "fISN", cashInputNextCreate.fIsn, 1)
    With dbo_FOLDERS(0)
        .fKEY = cashInputNextCreate.fIsn
        .fISN = cashInputNextCreate.fIsn
        .fSPEC = Left_Align(Get_Compname_DOCLOG(cashInputNextCreate.fIsn), 16) & "GlavBux ARMSOFT                       0115"
    End With
    Call CheckDB_FOLDERS(dbo_FOLDERS(0), 1)
End Sub