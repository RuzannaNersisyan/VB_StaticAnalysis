'USEUNIT Library_Common  
'USEUNIT Library_Colour
'USEUNIT Library_CheckDB  
'USEUNIT Loan_Agreements_Library 
'USEUNIT Akreditiv_Library
'USEUNIT Credit_Line_Library
'USEUNIT Loan_Agreemnts_With_Schedule_Library
'USEUNIT Deposit_Contract_Library
'USEUNIT Loan_Agreements_With_Schedule_Linear_Library
'USEUNIT Group_Operations_Library
'USEUNIT Overdraft_NewCases_Library
'USEUNIT Constants
'USEUNIT Mortgage_Library
Option Explicit

'Test Case N 165360

Dim loan, folderName, sDATE, fDATE
Dim dbo_CONTRACTS, dbo_FOLDERS(5), fBODY
Dim i, obj1, obj2

Sub Loan_Attracted_With_Limit_Test()
		Call Test_Initialize()

		'Ð³ÙÏ³ñ· Ùáõïù ·áñÍ»É ARMSOFT û·ï³·áñÍáÕáí
		Log.Message "Համակարգ մուտք գործել ARMSOFT օգտագործողով", "", pmNormal, DivideColor
    Call Test_StartUp()
		
		'êï»ÕÍ»É ì³ñÏ³ÛÇÝ ·ÇÍ å³ÛÙ³Ý³·Çñ
		Log.Message "Ստեղծել Վարկային գիծ պայմանագիր", "", pmNormal, DivideColor
		Call loan.CreateAttrLoan(folderName & "Üáñ å³ÛÙ³Ý³·ñÇ ëï»ÕÍáõÙ")
		
		BuiltIn.Delay(2000)
		Call Close_Pttel("frmPttel")
		
		'ì³ñÏ³ÛÇÝ ·ÇÍ å³ÛÙ³Ý³·ñÇ ëï»ÕÍáõÙÇó Ñ»ïá SQL ëïáõ·áõÙ
		Log.Message "Վարկային գիծ պայմանագրի ստեղծումից հետո SQL ստուգում", "", pmNormal, DivideColor
		Call DB_Inirtialize()
		Call Check_DB_CreateAttrLoan()
		
		'ä³ÛÙ³Ý³·ÇñÁ áõÕ³ñÏ»É Ñ³ëï³ïÙ³Ý 
		Log.Message "Պայմանագրին ուղղարկել հաստատման", "", pmNormal, DivideColor
    loan.SendToVerify(folderName & "²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ")
		
		' ä³ÛÙ³Ý³·ÇñÁ Ñ³ëï³ïÙ³Ý áõÕ³ñÏ»Éáõó Ñ»ïá SQL ëïáõ·áõÙ
		Log.Message "Պայմանագիրը հաստատման ուղարկելուց հետո SQL ստուգում", "", pmNormal, SqlDivideColor
		Call Check_DB_SendToVerify()
		
		'ì³í»ñ³óÝ»É å³ÛÙ³Ý³·ÇñÁ
		Log.Message "Վավերացնել պայմանագիրը", "", pmNormal, DivideColor
    loan.Verify(folderName & "Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I")
		
		'ä³ÛÙ³Ý³·ñÇ í³í»ñ³óáõÙÇó Ñ»ïá SQL ëïáõ·áõÙ
		Log.Message "Պայմանագրի վավերացումից հետո SQL ստուգում", "", pmNormal, SqlDivideColor
		Call Check_DB_VerifyContract()
  
		'ä³ÛÙ³Ý³·ñ»ñ ÃÕÃ³å³Ý³ÏáõÙ ÷³ëï³ÃÕÃÇ ³éÏ³ÛáõÃÛ³Ý ëïáõ·áõÙ  
		Log.Message "Պայմանագրեր թղթապանակում փաստատթղթի առկայության ստուգում", "", pmNormal, DivideColor
    Call LetterOfCredit_Filter_Fill(folderName, loan.DocLevel, loan.DocNum)
		
		'¶³ÝÓáõÙ Ý»ñ·ñ³íáõÙÇó
		Log.Message "Գանձում ներգրավումից", "", pmNormal, DivideColor
    Call ChargeForAttraction("", loan.Date, 100, "", "")
		
		'¶³ÝÓáõÙ Ý»ñ·ñ³íáõÙÇó Ñ»ïá SQL ëïáõ·áõÙ
		Log.Message "Գանձում ներգրավումից հետո SQL ստուգում", "", pmNormal, SqlDivideColor
		Call Check_DB_ChargeForAttr()
  
		'ì³ñÏÇ Ý»ñ·ñ³íáõÙ
    Log.Message "Վարկի ներգրավում", "", pmNormal, DivideColor
    Call Attraction(c_LoanAttraction, loan.Date, loan.Limit, "", "")
		
		'ì³ñÏÇ Ý»ñ·ñ³íáõÙÇó Ñ»ïá SQL ëïáõ·áõÙ
		Log.Message "Վարկի ներգրավումից հետո SQL ստուգում", "", pmNormal, SqlDivideColor
		Call Check_DB_Attraction()

    'îáÏáëÝ»ñÇ Ñ³ßí³ñÏ
    Log.Message "Տոկոսների հաշվարկ", "", pmNormal, DivideColor
    Call Calculate_Percents("010821", "010821", false)
		
		'îáÏáëÝ»ñÇ Ñ³ßí³ñÏÇó Ñ»ïá SQL ëïáõ·áõÙ
		Log.Message "Տոկոսների հաշվարկից հետո SQL ստուգում", "", pmNormal, SqlDivideColor
		Call Check_DB_CalcPercents()
		
		'ê³ÑÙ³Ý³ã³÷Ç ÷á÷áËáõÙ
		Log.Message "Սահմանաչափի փոփոխում", "", pmNormal, DivideColor
    Call Change_Limit("010821" , 200000)
		
		'ê³ÑÙ³Ý³ã³÷Ç ÷á÷áËáõÙÇó Ñ»ïá SQL ëïáõ·áõÙ
		Log.Message "Սահմանաչափի փոփոխումից հետո SQL ստուգում", "", pmNormal, SqlDivideColor
		Call Check_DB_ChangeLimit()
    
		'îáÏáëÝ»ñÇ Ï³åÇï³É³óáõÙ
    Log.Message "Տոկոսների կապիտալացում", "", pmNormal, DivideColor
    Call Percent_Capitalization(null , "020821", "")
		
		'îáÏáëÝ»ñÇ Ï³åÇï³É³óáõÙÇó Ñ»ïá SQL ëïáõ·áõÙ
		Log.Message "Տոկոսների կապիտալացումից հետո SQL ստուգում", "", pmNormal, SqlDivideColor
		Call Check_DB_PerCapitalization()
		
		'¶Í³ÛÝáõÃÛ³Ý ¹³¹³ñ»óáõÙ
		Log.Message "Գծայնության դադարեցում", "", pmNormal, DivideColor
    Call Credit_Line_Stop_Recovery_DocFill("020821", 1)
		
		BuiltIn.Delay(3000)
		Call Close_Pttel("frmPttel")
		
		'¶Í³ÛÝáõÃÛ³Ý ¹³¹³ñ»óáõÙÇó Ñ»ïá SQL ëïáõ·áõÙ
		Log.Message "Գծայնության դադարեցումից հետո SQL ստուգում", "", pmNormal, SqlDivideColor
		Call Check_DB_LineStopRec()
		
		'´áÉáñ ÷³ëï³ÃÕÃ»ñÇ çÝçáõÙ
		Log.Message "Բոլոր փաստաթղթերի ջնջում", "", pmNormal, DivideColor
    Call DeleteAllActions("|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|ä³ÛÙ³Ý³·ñ»ñ|Ü»ñ·ñ³íí³Í ÙÇçáóÝ»ñ|Ü»ñ·ñ³íí³Í í³ñÏ»ñ", loan.DocNum, "^a[Del]", "^a[Del]")
		
		'´áÉáñ ÷³ëï³ÃÕÃ»ñÇ çÝçáõÙÇó Ñ»ïá SQL ëïáõ·áõÙ
		Log.Message "Բոլոր փաստաթղթերի ջնջումից հետո SQL ստուգում", "", pmNormal, SqlDivideColor
		Call Check_DB_DeleteAllActions()
		
		Call Close_AsBank() 
End	Sub

Sub Test_StartUp()
		Call Initialize_AsBank("bank", sDATE, fDATE)
    Login("ARMSOFT")
		Call ChangeWorkspace(c_Subsystems)
End	Sub

Sub Test_Initialize()
		folderName = "|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|ä³ÛÙ³Ý³·ñ»ñ|Ü»ñ·ñ³íí³Í ÙÇçáóÝ»ñ|Ü»ñ·ñ³íí³Í í³ñÏ»ñ|"
		
		sDATE = "20030101"
		fDATE = "20260101"  
		
		Set loan = New_LoanDocument()
		With loan
      .DocType = "ì³ñÏ³ÛÇÝ ·ÇÍ"
      .CalcAcc = "00000113032"   
      '.Client = "00034852"                                 
      .Limit = 100000
      .Date = "020721" 
      .GiveDate = "020721"
      .Term = "020722"
      .FirstDate = "020721"
      .PaperCode = 555
		End With
End Sub

Sub DB_Inirtialize()
    Set dbo_CONTRACTS = New_DB_CONTRACTS()
    dbo_CONTRACTS.fDGISN = loan.fBASE
    dbo_CONTRACTS.fDGPARENTISN = loan.fBASE
    dbo_CONTRACTS.fDGISN1 = loan.fBASE
    dbo_CONTRACTS.fDGISN3 = loan.fBASE
    dbo_CONTRACTS.fDGAGRKIND = 1
    dbo_CONTRACTS.fDGSTATE = 1
    dbo_CONTRACTS.fDGTYPENAME = "D2AS21  "
    dbo_CONTRACTS.fDGCODE = loan.DocNum
    dbo_CONTRACTS.fDGCAPTION = "ý²àôêî"
    dbo_CONTRACTS.fDGCLICODE = "00034852"
    dbo_CONTRACTS.fDGCUR = "000"
    dbo_CONTRACTS.fDGSUMMA = "100000.00"
    dbo_CONTRACTS.fDGALLSUMMA = "0.00"
    dbo_CONTRACTS.fDGRISKDEGREE = "0.00"
    dbo_CONTRACTS.fDGRISKDEGNB = "0.00"
    dbo_CONTRACTS.fDGACSBRANCH = "00 "
    dbo_CONTRACTS.fDGACSDEPART = "1  "
    dbo_CONTRACTS.fDGACSTYPE = "D20 "
		
    For i = 0 To 4
      Set dbo_FOLDERS(i) = New_DB_FOLDERS()
      dbo_FOLDERS(i).fKEY = loan.fBASE
      dbo_FOLDERS(i).fISN = loan.fBASE
      dbo_FOLDERS(i).fNAME = "D2AS21  "
      dbo_FOLDERS(i).fSTATUS = "1"
    Next
    dbo_FOLDERS(0).fFOLDERID = "Agr." & loan.fBASE
    dbo_FOLDERS(0).fCOM = "ì³ñÏ³ÛÇÝ ·ÇÍ"
    dbo_FOLDERS(0).fSPEC = "1ì³ñÏ³ÛÇÝ ·ÇÍ- "& loan.DocNum & " {ý²àôêî}"
    dbo_FOLDERS(1).fFOLDERID = "C.903824400"
    dbo_FOLDERS(1).fCOM = " ì³ñÏ³ÛÇÝ ·ÇÍ (Ý³Ë³·ÇÍ)"
    dbo_FOLDERS(1).fSPEC = loan.DocNum & " (ý²àôêî),     100000 - Ð³ÛÏ³Ï³Ý ¹ñ³Ù"
    dbo_FOLDERS(2).fFOLDERID = "SSWork.CRD220210702" 
    dbo_FOLDERS(2).fCOM = "ì³ñÏ³ÛÇÝ ·ÇÍ"
    dbo_FOLDERS(2).fSPEC = loan.DocNum & "    D20 20210702            0.0077  00034852Üáñ å³ÛÙ³Ý³·Çñ      "
    dbo_FOLDERS(2).fECOM = "Credit Line"
    dbo_FOLDERS(2).fDCBRANCH = "00 "
    dbo_FOLDERS(2).fDCDEPART = "1  "
End	Sub

Sub Check_DB_CreateAttrLoan()
    'SQL Ստուգում CONTRACTS աղյուսակում 
    Log.Message "SQL Ստուգում CONTRACTS աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("CONTRACTS", "fDGISN", loan.fBASE, 1)
    Call CheckDB_CONTRACTS(dbo_CONTRACTS, 1)
  
    'SQL Ստուգում DOCLOG աղյուսակում
    Log.Message "SQL Ստուգում DOCLOG աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCLOG", "fISN", loan.fBASE, 1)
    Call CheckDB_DOCLOG(loan.fBASE, "77", "N", "1", "", 1)
  
    'SQL Ստուգում DOCP աղյուսակում  
    Log.Message "SQL Ստուգում DOCP աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCP", "fPARENTISN", loan.fBASE, 1)
    Call CheckDB_DOCP("443871031", "Acc", loan.fBASE, 1)
  
    'SQL Ստուգում DOCS աղյուսակում 
    Log.Message "SQL Ստուգում DOCS աղյուսակում", "", pmNormal, SqlDivideColor
  	fBODY = " CODE:" & loan.DocNum & " CLICOD:00034852 NAME:ý²àôêî CURRENCY:000 ACCACC:00000113032 SUMMA:100000 CHRGFIRSTDAY:0 AUTODEBT:1 DATE:20210702 ACSBRANCH:00 ACSDEPART:1 ACSTYPE:D20 KINDSCALE:1 PCAGR:12.0000/365 PCNOCHOOSE:8.0000/365 TAXVALUE:10 PCNDERAUTO:0 PCPENAGR:0/1 PCPENPER:0/1 DATEGIVE:20210702 DATEAGR:20220702 AUTODATE:0 REPAYADVANCE:100 AUTOCAP:0 ONSTPER:0 "
    fBODY = Replace(fBODY, " ", "%")
    Call CheckQueryRowCount("DOCS", "fISN", loan.fBASE, 1)
    Call CheckDB_DOCS(loan.fBASE, "D2AS21  ", "1", fBODY, 1)
  
  		'SQL Ստուգում DOCSG աղյուսակում 
    Log.Message "SQL Ստուգում DOCSG աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCSG", "fISN", loan.fBASE, 12)
  
    'SQL Ստուգում FOLDERS աղյուսակում 
    Log.Message "SQL Ստուգում FOLDERS աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("FOLDERS", "fISN", loan.fBASE, 3)
    For i = 0 To 2
      Call CheckDB_FOLDERS(dbo_FOLDERS(i), 1)
    Next
  
    'SQL Ստուգում RESNUMBERS աղյուսակում 
    Log.Message "SQL Ստուգում RESNUMBERS աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("RESNUMBERS", "fISN", loan.fBASE, 1)
    Call CheckDB_RESNUMBERS(loan.fBASE, "D ", loan.DocNum, 1)	
End	Sub

Sub Check_DB_SendToVerify()
    'SQL Ստուգում CONTRACTS աղյուսակում 
    Log.Message "SQL Ստուգում CONTRACTS աղյուսակում", "", pmNormal, SqlDivideColor
    dbo_CONTRACTS.fDGSTATE = 101
    Call CheckQueryRowCount("CONTRACTS", "fDGISN", loan.fBASE, 1)
    Call CheckDB_CONTRACTS(dbo_CONTRACTS, 1)
  
    'SQL Ստուգում DOCLOG աղյուսակում
    Log.Message "SQL Ստուգում DOCLOG աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCLOG", "fISN", loan.fBASE, 3)
    Call CheckDB_DOCLOG(loan.fBASE, "77", "M", "99", "àõÕ³ñÏí»É ¿ Ñ³ëï³ïÙ³Ý", 1)
  	Call CheckDB_DOCLOG(loan.fBASE, "77", "C", "101", "", 1)

    'SQL Ստուգում DOCS աղյուսակում
    Log.Message "SQL Ստուգում DOCS աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCS", "fISN", loan.fBASE, 1)
    Call CheckDB_DOCS(loan.fBASE, "D2AS21  ", "101", fBODY, 1)
  
  	'SQL Ստուգում FOLDERS աղյուսակում 
    Log.Message "SQL Ստուգում FOLDERS աղյուսակում", "", pmNormal, SqlDivideColor
    dbo_FOLDERS(3).fKEY = loan.fBASE
    dbo_FOLDERS(3).fISN = loan.fBASE
    dbo_FOLDERS(3).fNAME = "D2AS21  "
    dbo_FOLDERS(3).fSTATUS = "4"
    dbo_FOLDERS(3).fFOLDERID = "SSConf.CRD2001" 
    dbo_FOLDERS(3).fCOM = "ì³ñÏ³ÛÇÝ ·ÇÍ"
    dbo_FOLDERS(3).fSPEC = loan.DocNum & "    D20 20210702            0.0077  00034852"
    dbo_FOLDERS(3).fECOM = "Credit Line"
    dbo_FOLDERS(3).fDCBRANCH = "00 "
    dbo_FOLDERS(3).fDCDEPART = "1  "
    dbo_FOLDERS(0).fSTATUS = "0"
    dbo_FOLDERS(1).fSTATUS = "0"
    dbo_FOLDERS(2).fSTATUS = "0"
    dbo_FOLDERS(1).fCOM = " ì³ñÏ³ÛÇÝ ·ÇÍ" 
    dbo_FOLDERS(2).fSPEC = loan.DocNum & "    D20 20210702            0.0077  00034852àõÕ³ñÏí³Í I Ñ³ëï³ïÙ³"
    Call CheckQueryRowCount("FOLDERS", "fISN", loan.fBASE, 4)
    For i = 0 To 3
      Call CheckDB_FOLDERS(dbo_FOLDERS(i), 1)
    Next
End	Sub

Sub Check_DB_VerifyContract()		
    'SQL Ստուգում CONTRACTS աղյուսակում 
    Log.Message "SQL Ստուգում CONTRACTS աղյուսակում", "", pmNormal, SqlDivideColor
    dbo_CONTRACTS.fDGSTATE = 7
    Call CheckQueryRowCount("CONTRACTS", "fDGISN", loan.fBASE, 1)
    Call CheckDB_CONTRACTS(dbo_CONTRACTS, 1)
  
    'SQL Ստուգում DAGRACCS աղյուսակում 
    Log.Message "SQL Ստուգում DAGRACCS աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DAGRACCS", "fAGRISN", loan.fBASE, 1)
  
    'SQL Ստուգում DOCLOG աղյուսակում
    Log.Message "SQL Ստուգում DOCLOG աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCLOG", "fISN", loan.fBASE, 5)
    Call CheckDB_DOCLOG(loan.fBASE, "77", "W", "102", "", 1)
    Call CheckDB_DOCLOG(loan.fBASE, "77", "T", "7", "", 1)
		
    'SQL Ստուգում DOCP աղյուսակում  
    Log.Message "SQL Ստուգում DOCP աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCP", "fPARENTISN", loan.fBASE, 1)

    'SQL Ստուգում DOCS աղյուսակում
    Log.Message "SQL Ստուգում DOCS աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCS", "fISN", loan.fBASE, 1)
    Call CheckDB_DOCS(loan.fBASE, "D2AS21  ", "7", fBODY, 1)
  
    'SQL Ստուգում FOLDERS աղյուսակում 
    Log.Message "SQL Ստուգում FOLDERS աղյուսակում", "", pmNormal, SqlDivideColor
    dbo_FOLDERS(0).fSTATUS = "1"
    dbo_FOLDERS(1).fSTATUS = "1"
    Call CheckQueryRowCount("FOLDERS", "fISN", loan.fBASE, 2)
    for i = 0 to 1
    Call CheckDB_FOLDERS(dbo_FOLDERS(i), 1)
    next
  
    'SQL Ստուգում HIF  աղյուսակում 
    Log.Message "SQL Ստուգում HIF աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HIF", "fOBJECT", loan.fBASE, 26)
		
    'SQL Ստուգում HI աղյուսակում 
    Log.Message "SQL Ստուգում HI աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HI", "fBASE", loan.fBASE, 2)
    obj1 = Get_ColumnValueSQL("HI", "fOBJECT", "fBASE = " & loan.fBASE & " and fDBCR = 'D' and fSPEC like '%                  0.00²é³í»É³·áõÛÝ ë³ÑÙ³Ý³ã³÷                           %     8000000 %'")
    Call Check_HI_CE_accounting ("2021-07-02", loan.fBASE, "02", obj1, "100000.00", "000", "100000.00", "MSC", "D")
    obj2 = Get_ColumnValueSQL("HI", "fOBJECT", "fBASE = " & loan.fBASE & " and fDBCR = 'D' and fSPEC like '%                  0.00²é³í»É³·áõÛÝ ë³ÑÙ³Ý³ã³÷                           %     999998  %'")
    Call Check_HI_CE_accounting ("2021-07-02", loan.fBASE, "02", obj2, "100000.00", "000", "100000.00", "MSC", "D")
		
    'SQL Ստուգում HIREST  աղյուսակում 
    Log.Message "SQL Ստուգում HIREST աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HIREST", "fOBJECT", obj1, 3)		
    Call CheckDB_HIREST("02", obj1, "100000.00", "000", "100000.00", 1)
    Call CheckQueryRowCount("HIREST", "fOBJECT", obj2, 3)		
    Call CheckDB_HIREST("02", obj2, "100000.00", "000", "100000.00", 1)
End	Sub

Sub Check_DB_ChargeForAttr()  
    'SQL Ստուգում DAGRACCS աղյուսակում 
    Log.Message "SQL Ստուգում DAGRACCS աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DAGRACCS", "fAGRISN", loan.fBASE, 1)
  
    'SQL Ստուգում DOCLOG աղյուսակում
    Log.Message "SQL Ստուգում DOCLOG աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCLOG", "fISN", loan.fBASE, 5)
    Call CheckDB_DOCLOG(loan.fBASE, "77", "W", "102", "", 1)
    Call CheckDB_DOCLOG(loan.fBASE, "77", "T", "7", "", 1)
		
    'SQL Ստուգում DOCP աղյուսակում  
    Log.Message "SQL Ստուգում DOCP աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCP", "fPARENTISN", loan.fBASE, 1)

    'SQL Ստուգում DOCS աղյուսակում
    Log.Message "SQL Ստուգում DOCS աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCS", "fISN", loan.fBASE, 1)
    Call CheckDB_DOCS(loan.fBASE, "D2AS21  ", "7", fBODY, 1)
  
    'SQL Ստուգում FOLDERS աղյուսակում 
    Log.Message "SQL Ստուգում FOLDERS աղյուսակում", "", pmNormal, SqlDivideColor
    dbo_FOLDERS(0).fSTATUS = "1"
    dbo_FOLDERS(1).fSTATUS = "1"
    Call CheckQueryRowCount("FOLDERS", "fISN", loan.fBASE, 2)
    For i = 0 To 1
      Call CheckDB_FOLDERS(dbo_FOLDERS(i), 1)
    Next

    'SQL Ստուգում HI աղյուսակում 
    Log.Message "SQL Ստուգում HI աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HI", "fBASE", loan.fBASE, 2)
    obj1 = Get_ColumnValueSQL("HI", "fOBJECT", "fBASE = " & loan.fBASE & " and fDBCR = 'D' and fSPEC like '%                  0.00²é³í»É³·áõÛÝ ë³ÑÙ³Ý³ã³÷                           %     999998  %'")
    Call Check_HI_CE_accounting ("2021-07-02", loan.fBASE, "02", obj1, "100000.00", "000", "100000.00", "MSC", "D")
    obj2 = Get_ColumnValueSQL("HI", "fOBJECT", "fBASE = " & loan.fBASE & " and fDBCR = 'D' and fSPEC like '%                  0.00²é³í»É³·áõÛÝ ë³ÑÙ³Ý³ã³÷                           %     8000000 %'")
    Call Check_HI_CE_accounting ("2021-07-02", loan.fBASE, "02", obj2, "100000.00", "000", "100000.00", "MSC", "D")
		
    'SQL Ստուգում HIR  աղյուսակում 
    Log.Message "SQL Ստուգում HIR աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HIR", "fOBJECT", loan.fBASE, 2)
    Call Check_HIR("2021-07-02", "R^", loan.fBASE, "000", "90.00", "PAY", "D")
    Call Check_HIR("2021-07-02", "R^", loan.fBASE, "000", "10.00", "TAX", "D")
		
    'SQL Ստուգում HIREST  աղյուսակում 
    Log.Message "SQL Ստուգում HIREST աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HIREST", "fOBJECT", obj1, 3)		
    Call CheckDB_HIREST("02", obj1, "100000.00", "000", "100000.00", 1)
    Call CheckQueryRowCount("HIREST", "fOBJECT", obj2, 3)		
    Call CheckDB_HIREST("02", obj2, "100000.00", "000", "100000.00", 1)
		
    'SQL Ստուգում HIRREST  աղյուսակում 
    Log.Message "SQL Ստուգում HIRREST աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HIRREST", "fOBJECT", loan.fBASE, 1)
    Call CheckDB_HIRREST("R^", loan.fBASE, "100.00", "2021-07-02", 1)
End	Sub

Sub Check_DB_Attraction()
    'SQL Ստուգում DOCLOG աղյուսակում
    Log.Message "SQL Ստուգում DOCLOG աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCLOG", "fISN", loan.fBASE, 5)
    Call CheckDB_DOCLOG(loan.fBASE, "77", "W", "102", "", 1)
    Call CheckDB_DOCLOG(loan.fBASE, "77", "T", "7", "", 1)
		
    'SQL Ստուգում DOCS աղյուսակում
    Log.Message "SQL Ստուգում DOCS աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCS", "fISN", loan.fBASE, 1)
    Call CheckDB_DOCS(loan.fBASE, "D2AS21  ", "7", fBODY, 1)
		
    'SQL Ստուգում HI աղյուսակում 
    Log.Message "SQL Ստուգում HI աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HI", "fBASE", loan.fBASE, 2)
    Call Check_HI_CE_accounting ("2021-07-02", loan.fBASE, "02", obj1, "100000.00", "000", "100000.00", "MSC", "D")
    Call Check_HI_CE_accounting ("2021-07-02", loan.fBASE, "02", obj2, "100000.00", "000", "100000.00", "MSC", "D")
		
    'SQL Ստուգում HIR  աղյուսակում 
    Log.Message "SQL Ստուգում HIR աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HIR", "fOBJECT", loan.fBASE, 3)
    Call Check_HIR("2021-07-02", "R1", loan.fBASE, "000", "100000.00", "AGR", "D")
		
    'SQL Ստուգում HIREST  աղյուսակում 
    Log.Message "SQL Ստուգում HIREST աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HIREST", "fOBJECT", obj1, 3)		
    Call CheckQueryRowCount("HIREST", "fOBJECT", obj2, 3)		
		
    'SQL Ստուգում HIRREST  աղյուսակում 
    Log.Message "SQL Ստուգում HIRREST աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HIRREST", "fOBJECT", loan.fBASE, 2)
    Call CheckDB_HIRREST("R1", loan.fBASE, "100000.00", "2021-07-02", 1)
End	Sub

Sub Check_DB_CalcPercents()
    'SQL Ստուգում DAGRACCS աղյուսակում 
    Log.Message "SQL Ստուգում DAGRACCS աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DAGRACCS", "fAGRISN", loan.fBASE, 1)
		
    'SQL Ստուգում DOCLOG աղյուսակում
    Log.Message "SQL Ստուգում DOCLOG աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCLOG", "fISN", loan.fBASE, 5)
		
    'SQL Ստուգում DOCP աղյուսակում  
    Log.Message "SQL Ստուգում DOCP աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCP", "fPARENTISN", loan.fBASE, 1)
		
    'SQL Ստուգում DOCS աղյուսակում
    Log.Message "SQL Ստուգում DOCS աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCS", "fISN", loan.fBASE, 1)
    Call CheckDB_DOCS(loan.fBASE, "D2AS21  ", "7", fBODY, 1)
		
    'SQL Ստուգում FOLDERS աղյուսակում 
    Log.Message "SQL Ստուգում FOLDERS աղյուսակում", "", pmNormal, SqlDivideColor
    dbo_FOLDERS(0).fSTATUS = "1"
    dbo_FOLDERS(1).fSTATUS = "1"
    Call CheckQueryRowCount("FOLDERS", "fISN", loan.fBASE, 2)
    for i = 0 to 1
    Call CheckDB_FOLDERS(dbo_FOLDERS(i), 1)
    next
		
    'SQL Ստուգում HI աղյուսակում 
    Log.Message "SQL Ստուգում HI աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HI", "fBASE", loan.fBASE, 2)
    Call Check_HI_CE_accounting ("2021-07-02", loan.fBASE, "02", obj1, "100000.00", "000", "100000.00", "MSC", "D")
    Call Check_HI_CE_accounting ("2021-07-02", loan.fBASE, "02", obj2, "100000.00", "000", "100000.00", "MSC", "D")
		
    'SQL Ստուգում HIF  աղյուսակում 
    Log.Message "SQL Ստուգում HIF աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HIF", "fOBJECT", loan.fBASE, 27)
		
    'SQL Ստուգում HIR  աղյուսակում 
    Log.Message "SQL Ստուգում HIR աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HIR", "fOBJECT", loan.fBASE, 5)
    Call Check_HIR("2021-08-01", "R2", loan.fBASE, "000", "986.30", "PER", "D")
    Call Check_HIR("2021-08-02", "R¸", loan.fBASE, "000", "986.30", "PRJ", "D")
		
    'SQL Ստուգում HIREST  աղյուսակում 
    Log.Message "SQL Ստուգում HIREST աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HIREST", "fOBJECT", obj1, 3)		
    Call CheckQueryRowCount("HIREST", "fOBJECT", obj2, 3)		
		
    'SQL Ստուգում HIRREST  աղյուսակում 
    Log.Message "SQL Ստուգում HIRREST աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HIRREST", "fOBJECT", loan.fBASE, 4)
    Call CheckDB_HIRREST("R2", loan.fBASE, "986.30", "2021-08-01", 1)
    Call CheckDB_HIRREST("R¸", loan.fBASE, "986.30", "2021-08-02", 1)
		
    'SQL Ստուգում HIT  աղյուսակում 
    Log.Message "SQL Ստուգում HIT աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HIT", "fOBJECT", loan.fBASE, 1)
    Call Check_HIT("2021-08-01", "N2", loan.fBASE, "000", "986.30", "PER", "D")
End	Sub

Sub Check_DB_ChangeLimit()
    'SQL Ստուգում DOCLOG աղյուսակում
    Log.Message "SQL Ստուգում DOCLOG աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCLOG", "fISN", loan.fBASE, 5)
		
    'SQL Ստուգում DOCS աղյուսակում
    Log.Message "SQL Ստուգում DOCS աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCS", "fISN", loan.fBASE, 1)
    Call CheckDB_DOCS(loan.fBASE, "D2AS21  ", "7", fBODY, 1)
		
    'SQL Ստուգում FOLDERS աղյուսակում 
    Log.Message "SQL Ստուգում FOLDERS աղյուսակում", "", pmNormal, SqlDivideColor
    dbo_FOLDERS(0).fSTATUS = "1"
    dbo_FOLDERS(1).fSTATUS = "1"
    Call CheckQueryRowCount("FOLDERS", "fISN", loan.fBASE, 2)
    For i = 0 To 1
      Call CheckDB_FOLDERS(dbo_FOLDERS(i), 1)
    Next
		
    'SQL Ստուգում HI աղյուսակում 
    Log.Message "SQL Ստուգում HI աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HI", "fBASE", loan.fBASE, 2)
    Call Check_HI_CE_accounting ("2021-07-02", loan.fBASE, "02", obj1, "100000.00", "000", "100000.00", "MSC", "D")
    Call Check_HI_CE_accounting ("2021-07-02", loan.fBASE, "02", obj2, "100000.00", "000", "100000.00", "MSC", "D")
		
    'SQL Ստուգում HIF  աղյուսակում 
    Log.Message "SQL Ստուգում HIF աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HIF", "fOBJECT", loan.fBASE, 28)
		
    'SQL Ստուգում HIREST  աղյուսակում 
    Log.Message "SQL Ստուգում HIREST աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HIREST", "fOBJECT", obj1, 3)		
    Call CheckQueryRowCount("HIREST", "fOBJECT", obj2, 3)		
End	Sub

Sub Check_DB_PerCapitalization()
    'SQL Ստուգում DOCLOG աղյուսակում
    Log.Message "SQL Ստուգում DOCLOG աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCLOG", "fISN", loan.fBASE, 5)
		
    'SQL Ստուգում DOCS աղյուսակում
    Log.Message "SQL Ստուգում DOCS աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCS", "fISN", loan.fBASE, 1)
    Call CheckDB_DOCS(loan.fBASE, "D2AS21  ", "7", fBODY, 1)
		
    'SQL Ստուգում HI աղյուսակում 
    Log.Message "SQL Ստուգում HI աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HI", "fBASE", loan.fBASE, 2)
    Call Check_HI_CE_accounting ("2021-07-02", loan.fBASE, "02", obj1, "100000.00", "000", "100000.00", "MSC", "D")
    Call Check_HI_CE_accounting ("2021-07-02", loan.fBASE, "02", obj2, "100000.00", "000", "100000.00", "MSC", "D")
		
    'SQL Ստուգում HIR  աղյուսակում 
    Log.Message "SQL Ստուգում HIR աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HIR", "fOBJECT", loan.fBASE, 10)
    Call Check_HIR("2021-08-01", "R2", loan.fBASE, "000", "986.30", "PER", "D")
    Call Check_HIR("2021-08-02", "R¸", loan.fBASE, "000", "986.30", "PRJ", "D")
		
    'SQL Ստուգում HIREST  աղյուսակում 
    Log.Message "SQL Ստուգում HIREST աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HIREST", "fOBJECT", obj1, 3)		
    Call CheckQueryRowCount("HIREST", "fOBJECT", obj2, 3)		
		
    'SQL Ստուգում HIRREST  աղյուսակում 
    Log.Message "SQL Ստուգում HIRREST աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HIRREST", "fOBJECT", loan.fBASE, 4)
    Call CheckDB_HIRREST("R1", loan.fBASE, "100887.70", "2021-08-02", 1)
    Call CheckDB_HIRREST("R2", loan.fBASE, "0.00", "2021-08-02", 1)
    Call CheckDB_HIRREST("R^", loan.fBASE, "100.00", "2021-07-02", 1)
    Call CheckDB_HIRREST("R¸", loan.fBASE, "0.00", "2021-08-02", 1)
End	Sub

Sub Check_DB_LineStopRec()
    'SQL Ստուգում DOCLOG աղյուսակում
    Log.Message "SQL Ստուգում DOCLOG աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCLOG", "fISN", loan.fBASE, 5)
		
    'SQL Ստուգում DOCS աղյուսակում
    Log.Message "SQL Ստուգում DOCS աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCS", "fISN", loan.fBASE, 1)
    Call CheckDB_DOCS(loan.fBASE, "D2AS21  ", "7", fBODY, 1)
		
    'SQL Ստուգում HI աղյուսակում 
    Log.Message "SQL Ստուգում HI աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HI", "fBASE", loan.fBASE, 2)
    Call Check_HI_CE_accounting ("2021-07-02", loan.fBASE, "02", obj1, "100000.00", "000", "100000.00", "MSC", "D")
    Call Check_HI_CE_accounting ("2021-07-02", loan.fBASE, "02", obj2, "100000.00", "000", "100000.00", "MSC", "D")
		
    'SQL Ստուգում HIR  աղյուսակում 
    Log.Message "SQL Ստուգում HIR աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HIR", "fOBJECT", loan.fBASE, 10)
    Call Check_HIR("2021-08-02", "R1", loan.fBASE, "000", "887.70", "CAP", "D")
    Call Check_HIR("2021-08-02", "R2", loan.fBASE, "000", "887.70", "CAP", "C")
    Call Check_HIR("2021-08-02", "R2", loan.fBASE, "000", "98.60", "TXC", "C")
    Call Check_HIR("2021-08-02", "R¸", loan.fBASE, "000", "887.70", "CAP", "C")
    Call Check_HIR("2021-08-02", "R¸", loan.fBASE, "000", "98.60", "TXC", "C")
		
    'SQL Ստուգում HIREST  աղյուսակում 
    Log.Message "SQL Ստուգում HIREST աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HIREST", "fOBJECT", obj1, 3)		
    Call CheckQueryRowCount("HIREST", "fOBJECT", obj2, 3)		
		
    'SQL Ստուգում HIRREST  աղյուսակում 
    Log.Message "SQL Ստուգում HIRREST աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HIRREST", "fOBJECT", loan.fBASE, 4)
    Call CheckDB_HIRREST("R1", loan.fBASE, "100887.70", "2021-08-02", 1)
    Call CheckDB_HIRREST("R2", loan.fBASE, "0.00", "2021-08-02", 1)
    Call CheckDB_HIRREST("R^", loan.fBASE, "100.00", "2021-07-02", 1)
    Call CheckDB_HIRREST("R¸", loan.fBASE, "0.00", "2021-08-02", 1)
End	Sub

Sub Check_DB_DeleteAllActions()
    'SQL Ստուգում DOCLOG աղյուսակում
    Log.Message "SQL Ստուգում DOCLOG աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCLOG", "fISN", loan.fBASE, 6)
    Call CheckDB_DOCLOG(loan.fBASE, "77", "D", "999", "", 1)
		
    'SQL Ստուգում DOCS աղյուսակում
    Log.Message "SQL Ստուգում DOCS աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCS", "fISN", loan.fBASE, 1)
    Call CheckDB_DOCS(loan.fBASE, "D2AS21  ", "999", fBODY, 1)
		
    'SQL Ստուգում FOLDERS աղյուսակում 
    Log.Message "SQL Ստուգում FOLDERS աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("FOLDERS", "fISN", loan.fBASE, 1)
    dbo_FOLDERS(3).fKEY = loan.fBASE
    dbo_FOLDERS(3).fISN = loan.fBASE
    dbo_FOLDERS(3).fNAME = "D2AS21  "
    dbo_FOLDERS(3).fSTATUS = "0"
    dbo_FOLDERS(3).fFOLDERID = ".R." & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%Y%m%d")
    dbo_FOLDERS(3).fCOM = ""
    dbo_FOLDERS(3).fSPEC = Left_Align(Get_Compname_DOCLOG(loan.fBASE), 16) &  "Cred&DepARMSOFT                       007  "
    dbo_FOLDERS(3).fECOM = ""
    dbo_FOLDERS(3).fDCBRANCH = "00 "
    dbo_FOLDERS(3).fDCDEPART = "1  "
    Call CheckDB_FOLDERS(dbo_FOLDERS(3), 1)
End	Sub