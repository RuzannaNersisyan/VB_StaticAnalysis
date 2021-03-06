'USEUNIT Library_Common  
'USEUNIT Library_Colour
'USEUNIT Library_Common 
'USEUNIT Financial_Leasing_Library 
'USEUNIT Akreditiv_Library
'USEUNIT Loan_Agreements_Library 
'USEUNIT Loan_Agreemnts_With_Schedule_Library
'USEUNIT Overdraft_NewCases_Library
'USEUNIT Group_Operations_Library
'USEUNIT Constants
'USEUNIT Deposit_Contract_Library
'USEUNIT Library_CheckDB
'USEUNIT Mortgage_Library
Option Explicit

'Test case ID 146218		

Dim FolderName, fDATE, sDATE
Dim Leasing, agreementAllOperations1, documentType, obj
Dim dbo_CONTRACTS, fBODY, dbo_FOLDERS(4), dbo_FOLDERSVerify(3), fBASE


Sub Financial_Leasing_Test()
    Call Test_Initialize()

    ''1. Համակարգ մուտք գործել ARMSOFT օգտագործողով
    Log.Message "Համակարգ մուտք գործել ARMSOFT օգտագործողով", "", pmNormal, DivideColor
    Call Test_StartUp()
  
    ''2. Մուտք գործել "Ենթահամակրգեր(ՀԾ)"
    Log.Message "Մուտք գործել ""Ենթահամակարգեր(ՀԾ)""", "", pmNormal, DivideColor
    Call ChangeWorkspace(c_Subsystems)  

    ''3. Լիզինգի պայմանագրի ստեղծում
    Log.Message "Լիզինգի պայմանագրի ստեղծում", "", pmNormal, DivideColor
    Call Leasing.CreateLeasing(FolderName & "Üáñ å³ÛÙ³Ý³·ñÇ ëï»ÕÍáõÙ")
    
    BuiltIn.Delay(3000)
    Call Close_Pttel("frmPttel")
		
    Log.Message "Ստեղծել Լիզինգի պայմանագիր-ի SQL ստուգում", "", pmNormal, SqlDivideColor
    Call DB_Initialize()
    Call Check_DB_AfterCreatingLeasingDoc()
		
    ''4. Պայմանագիրը ուղարկել հաստատման
    Log.Message "Պայմանագիրը ուղարկել հաստատման", "", pmNormal, DivideColor
    Leasing.SendToVerify(FolderName & "²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ")
		
    Log.Message "Պայմանագիրը ուղարկել հաստատման SQL ստուգում", "", pmNormal, SqlDivideColor
    Call Check_DB_AfterSendToVerify()
		
    ''5. Վավերացնել պայմանագիրը
    Log.Message "Վավերացնել պայմանագիրը", "", pmNormal, DivideColor
    Leasing.Verify(FolderName & "Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I")
  
    Call LetterOfCredit_Filter_Fill(FolderName, Leasing.DocLevel, Leasing.DocNum)
  
    Log.Message "Վավերացնել պայմանագիրը SQL ստուգում", "", pmNormal, SqlDivideColor
    Call Check_DB_VerifyContract()
		
    ''6. Գանձում տրամադրումից
    Log.Message "Գանձում տրամադրումից", "", pmNormal, DivideColor
    Call Collect_From_Provision("251020", "10000", 2, Leasing.CalcAcc, fBASE)
  
    Log.Message "Գանձում տրամադրումից SQL ստուգում", "", pmNormal, SqlDivideColor
    Call Check_DB_CollectFromPrev()
		
    ''7. Լիզինգի տրամադրում
    Log.Message "Լիզինգի տրամադրում", "", pmNormal, DivideColor
    Call Give_Leasing("251020")
  
    Log.Message "Լիզինգի տրամադրում SQL ստուգում", "", pmNormal, SqlDivideColor
    Call Check_DB_GiveLeasing()
		
    ''8. Տոկոսների հաշվարկ
    Log.Message "Տոկոսների հաշվարկ", "", pmNormal, DivideColor
    Call Calculate_Percents("241120", "241120", False)

    Log.Message "Տոկոսների հաշվարկ SQL ստուգում", "", pmNormal, SqlDivideColor
    Call Check_DB_CalcPercents()		
		
    ''9. Պարտքերի մարում
    Log.Message "Պարտքերի մարում", "", pmNormal, DivideColor
    Call Leasing_Fade_Debt(fBASE, "251120", "101220", "", 2, Leasing.CalcAcc, Leasing.DocNum)  
  
    Log.Message "Պարտքերի մարում SQL ստուգում", "", pmNormal, SqlDivideColor
    Call Check_DB_FadeDebt()
		    
    ''10. Տոկոսադրույքներ
    Log.Message "Տոկոսադրույքներ", "", pmNormal, DivideColor
    Call ChangeRete("251120", 85, 80)
    
    Log.Message "Տոկոսադրույքներ SQL ստուգում", "", pmNormal, SqlDivideColor
    Call Check_DB_ChangeRate()
		
    ''11. Արդյունավետ տոկոսադրույք
    Log.Message "Արդյունավետ տոկոսադրույք", "", pmNormal, DivideColor
    Call ChangeEffRete("251120", "", "")
  
    Log.Message "Արդյունավետ տոկոսադրույք SQL ստուգում", "", pmNormal, SqlDivideColor
    Call Check_DB_ChangeEffRete()
		   
    ''12. Տոկոսների հաշվարկ
    Log.Message "Տոկոսների հաշվարկ", "", pmNormal, DivideColor
    Call Calculate_Percents("251120", "251120", False)
     
    Log.Message "Տոկոսների հաշվարկ SQL ստուգում", "", pmNormal, SqlDivideColor
    Call Check_DB_CalcPercents2()
		 
    ''13. Օբյեկտիվ ռիսկի դասիչ
    Log.Message "Օբյեկտիվ ռիսկի դասիչ", "", pmNormal, DivideColor
    Call ObjectiveRisk("261120", "04")
   
    Log.Message "Օբյեկտիվ ռիսկի դասիչ SQL ստուգում", "", pmNormal, SqlDivideColor
    Call Check_DB_ObjRisk()
		 
    ''14. Ռիսկի դասիչ և պահուստավորման տոկոս
    Log.Message "Ռիսկի դասիչ և պահուստավորման տոկոս", "", pmNormal, DivideColor
    Call FillDoc_Risk_Classifier("261120", "05", 100)
   
    Log.Message "Ռիսկի դասիչ և պահուստավորման տոկոս SQL ստուգում", "", pmNormal, SqlDivideColor
    Call Check_DB_RiskClassifier()
		 
    ''15. Պահուստավորում
    Log.Message "Պահուստավորում", "", pmNormal, DivideColor
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_Opers & "|" & c_Store & "|" & c_Store)
    Call Rekvizit_Fill("Document", 1, "General", "DATE", "^A[Del]" & "261120")
    With wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame").VBObject("DocGrid")
      .Row = 0
      .Col = 1 
      .Keys("10")  
    End With 
    Call ClickCmdButton(1, "Î³ï³ñ»É") 
   
    Log.Message "Պահուստավորում SQL ստուգում", "", pmNormal, SqlDivideColor
    Call Check_DB_Backup()
		 
    ''16. Դուրս գրում
    Log.Message "Դուրս գրում", "", pmNormal, DivideColor
    Call FillDoc_WriteOut("261120", fBASE)
   
    Log.Message "Դուրս գրում SQL ստուգում", "", pmNormal, SqlDivideColor
    Call Check_DB_WriteOut()
		 
    ''17. Դուրս գրածի վերականգնում
    Log.Message "Դուրս գրածի վերականգնում", "", pmNormal, DivideColor
    Call WriteOffReconstruction("261120", "", "")
   
    Log.Message "Դուրս գրածի վերականգնում SQL ստուգում", "", pmNormal, SqlDivideColor
    Call Check_DB_WriteOffRec()
		 
    ''18. Պարտքերի մարում
    Log.Message "Պարտքերի մարում", "", pmNormal, DivideColor
    Call Leasing_Fade_Debt(fBASE, "261120", "251220", "", 2, Leasing.CalcAcc, Leasing.DocNum)  
   
    Log.Message "Պարտքերի մարում SQL ստուգում", "", pmNormal, SqlDivideColor
    Call Check_DB_FadeDebt2()
		 
    '		'Պայմանագրի փակում
    '  Log.Message "Պայմանագրի փակում", "", pmNormal, DivideColor
    '  Leasing.CloseDate = "261120"
    '  Leasing.CloseAgr()
    '
    '  'Պայմանագրի բացում
    '  Log.Message "Պայմանագրի բացում", "", pmNormal, DivideColor
    '  Leasing.OpenAgr()
  
    BuiltIn.Delay(3000) 
    Call Close_Pttel("frmPttel")

    ''19. æÝç»É µáÉáñ ·áñÍáÕáõÃÛáõÝÝ»ñÁ
    Log.Message "Բոլոր փաստաթղթերի ջնջում", "", pmNormal, DivideColor
    agreementAllOperations1.agreementN = Leasing.DocNum
    Call Delete_AgreementAllOperations(FolderName, agreementAllOperations1, "frmPttel", 4, documentType)
		
    Log.Message "Բոլոր փաստաթղթերի ջնջում SQL ստուգում", "", pmNormal, SqlDivideColor
    Call Check_DB_Delete_AgreementAllOperations()
		
    Call Close_AsBank()
End Sub

Sub Test_Initialize()
		FolderName = "|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|ä³ÛÙ³Ý³·ñ»ñ|î»Õ³µ³ßËí³Í ÙÇçáóÝ»ñ|üÇÝ³Ýë³Ï³Ý ÉÇ½ÇÝ· (ï»Õ³µ³ßËí³Í)|"
	
		sDATE = "20200101"
		fDATE = "20260101"  
	
		Set Leasing = New_LeasingDoc()
        Leasing.CalcAcc = "00000113032"
        Leasing.PermAsAcc = "72110332100"
        Leasing.Date = "221020"
        Leasing.GiveDate = "221020"
        Leasing.StartDate = "251020"
        Leasing.Summa = 10000
        Leasing.BuyPrice = 10
        Leasing.PaperCode = 111
        Leasing.Term = "221021"
        Leasing.DatesFillType = 1
        Leasing.DocType = "ÈÇ½ÇÝ·"
        Leasing.LastDate = Leasing.Term
        Leasing.office = "00"
      	Leasing.department = "1"	
        Leasing.accessType = "C40"	
  
		Set agreementAllOperations1 = New_AgreementAllOperations()
        agreementAllOperations1.startDate = "01/01/20"
        agreementAllOperations1.endDate = "01/01/23"
		
	Redim documentType(13)
    		documentType(13) = "üÇÝ³Ýë³Ï³Ý ÉÇ½ÇÝ·"
        documentType(12) = "¶³ÝÓáõÙ ïñ³Ù³¹ñáõÙÇó"
        documentType(11) = "¶áõÙ³ñÇ ïñ³Ù³¹ñáõÙ"
        documentType(10) = "îáÏáëÝ»ñÇ Ñ³ßí³ñÏáõÙ"
        documentType(9) = "ä³ñïù»ñÇ Ù³ñÙ³Ý Ñ³Ûï"
        documentType(8) = "îáÏáë³¹ñáõÛùÝ»ñ"
        documentType(7) = "²ñ¹ÛáõÝ³í»ï ïáÏáë³¹ñáõÛù"
        documentType(6) = "îáÏáëÝ»ñÇ Ñ³ßí³ñÏáõÙ"
        documentType(5) = "úµÛ»ÏïÇí éÇëÏÇ ¹³ëÇã"
        documentType(4) = "èÇëÏÇ ¹³ëÇã ¨ å³Ñáõëï³íáñÙ³Ý ïáÏáë"
    		documentType(3) = "ä³Ñáõëï³íáñáõÙ"
    		documentType(2) = "¸áõñë ·ñáõÙ"
    		documentType(1) = "¸áõñë ·ñí³ÍÇ í»ñ³Ï³Ý·ÝáõÙ"
    		documentType(0) = "ä³ñïù»ñÇ Ù³ñÙ³Ý Ñ³Ûï"
End Sub

Sub Test_StartUp()
		Call Initialize_AsBank("bank", sDATE, fDATE)
    Login("ARMSOFT")
End	Sub

Sub DB_Initialize()
		Dim i
		Set dbo_CONTRACTS = New_DB_CONTRACTS()
    dbo_CONTRACTS.fDGISN = Leasing.ISN
    dbo_CONTRACTS.fDGPARENTISN = Leasing.ISN
    dbo_CONTRACTS.fDGISN1 = Leasing.ISN
    dbo_CONTRACTS.fDGISN3 = Leasing.ISN
    dbo_CONTRACTS.fDGAGRKIND = 3
    dbo_CONTRACTS.fDGSTATE = 1
    dbo_CONTRACTS.fDGTYPENAME = "C4Diagr "
    dbo_CONTRACTS.fDGCODE = Leasing.DocNum
    dbo_CONTRACTS.fDGPPRCODE = "111"
    dbo_CONTRACTS.fDGCAPTION = "ý²àôêî"
    dbo_CONTRACTS.fDGCLICODE = "00034852"
    dbo_CONTRACTS.fDGCUR = "000"
    dbo_CONTRACTS.fDGSUMMA = "10000.00"
    dbo_CONTRACTS.fDGALLSUMMA = "10652.80"
    dbo_CONTRACTS.fDGRISKDEGREE = "0.00"
    dbo_CONTRACTS.fDGRISKDEGNB = "0.00"
    dbo_CONTRACTS.fDGSCHEDULE = "9"
    dbo_CONTRACTS.fDGDISTRICT = "001"
    dbo_CONTRACTS.fDGACSBRANCH = "00"
    dbo_CONTRACTS.fDGACSDEPART = "1"
    dbo_CONTRACTS.fDGACSTYPE = "C40"
    dbo_CONTRACTS.fDGAIM = "00"
    dbo_CONTRACTS.fDGUSAGEFIELD = "01.001"
    dbo_CONTRACTS.fDGCOUNTRY = "AM "
    dbo_CONTRACTS.fDGREGION = "010000008"
    dbo_CONTRACTS.fDGCRDTCODE = Leasing.CreditCode
		
		For i = 0 to 2
		  Set dbo_FOLDERS(i) = New_DB_FOLDERS()
		  dbo_FOLDERS(i).fKEY = Leasing.ISN
		  dbo_FOLDERS(i).fISN = Leasing.ISN
		  dbo_FOLDERS(i).fNAME = "C4Diagr "
		  dbo_FOLDERS(i).fSTATUS = "1"
    Next
    dbo_FOLDERS(0).fFOLDERID = "Agr." & Leasing.ISN
    dbo_FOLDERS(0).fCOM = "üÇÝ³Ýë³Ï³Ý ÉÇ½ÇÝ·"
    dbo_FOLDERS(0).fSPEC = "1üÇÝ³Ýë³Ï³Ý ÉÇ½ÇÝ·- " & Leasing.DocNum & " {ý²àôêî}"
    dbo_FOLDERS(1).fFOLDERID = "C.903824400"
    dbo_FOLDERS(1).fCOM = " üÇÝ³Ýë³Ï³Ý ÉÇ½ÇÝ· (Ý³Ë³·ÇÍ)"
    dbo_FOLDERS(1).fSPEC = Leasing.DocNum & " (ý²àôêî),     10652.8 - Ð³ÛÏ³Ï³Ý ¹ñ³Ù"
    dbo_FOLDERS(2).fFOLDERID = "SSWork.CRC420201022"
    dbo_FOLDERS(2).fCOM = "üÇÝ³Ýë³Ï³Ý ÉÇ½ÇÝ·"
    dbo_FOLDERS(2).fSPEC = Leasing.DocNum & "          C40 20201022            0.0077  00034852Üáñ å³ÛÙ³Ý³·Çñ      "
    dbo_FOLDERS(2).fECOM = "üÇÝ³Ýë³Ï³Ý ÉÇ½ÇÝ·"
    dbo_FOLDERS(2).fDCBRANCH = "00 "
    dbo_FOLDERS(2).fDCDEPART = "1  "
End	Sub

Sub Check_DB_AfterCreatingLeasingDoc()
		Dim i
  'SQL Ստուգում CONTRACTS աղուսյակում 
  Log.Message "SQL Ստուգում CONTRACTS աղուսյակում", "", pmNormal, SqlDivideColor
  Call CheckQueryRowCount("CONTRACTS", "fDGISN", Leasing.ISN, 1)
  Call CheckDB_CONTRACTS(dbo_CONTRACTS, 1)
  
  'SQL Ստուգում DOCLOG աղուսյակում համար
  Log.Message "SQL Ստուգում DOCLOG աղուսյակում", "", pmNormal, SqlDivideColor
  Call CheckQueryRowCount("DOCLOG", "fISN", Leasing.ISN, 1)
  Call CheckDB_DOCLOG(Leasing.ISN, "77", "N", "1", "", 1)
  
  'SQL Ստուգում DOCP աղուսյակում  
  Log.Message "SQL Ստուգում DOCP աղուսյակում", "", pmNormal, SqlDivideColor
  Call CheckQueryRowCount("DOCP", "fPARENTISN", Leasing.ISN, 1)
  Call CheckDB_DOCP("443871031", "Acc", Leasing.ISN, 1)
  
  'SQL Ստուգում DOCS աղուսյակում 
  Log.Message "SQL Ստուգում DOCS աղուսյակում", "", pmNormal, SqlDivideColor
		fBODY = " CODE:" & Leasing.DocNum & " CLICOD:00034852 NAME:ý²àôêî CURRENCY:000 ACCACC:00000113032 ALLSUMMA:10652.8 SUMMA:10000 ISDISCOUNT:1 DATE:20201022 ACSBRANCH:00 ACSDEPART:1 ACSTYPE:C40 AUTODEBT:1 DEBTJPART1:2 DEBTJPART:0 USECLICONNSCH:0 ONLYOVERDUE:0 KINDSCALE:2 PCAGR:12.0000/365 PCNDERAUTO:0 PCPENAGR:0/1 PCPENPER:0/1 PCLOSS:0/1 CALCFINPER:0 CALCJOUTS:0 DATEGIVE:20201022 DATEBEGCHRG:20201025 CONSTPER:1 AUTODATEA:0 REFRPERSUM:0 SECTOR:U2 USAGEFIELD:01.001 AIM:00 SCHEDULE:9 GUARANTEE:9 COUNTRY:AM LRDISTR:001 REGION:010000008 PERRES:1 REDUCEOVRDDAYS:0 WEIGHTAMDRISK:0 PPRCODE:111 SUBJRISK:0 CHRGFIRSTDAY:1 AUTOCAP:0 GIVEN:0 ISNBOUT:0 PUTINLR:1 NOTCLASS:0 OTHERCOLLATERAL:0 "
  fBODY = Replace(fBODY, " ", "%")
  Call CheckQueryRowCount("DOCS", "fISN", Leasing.ISN, 1)
  Call CheckDB_DOCS(Leasing.ISN, "C4Diagr  ", "1", fBODY, 1)
  
  'SQL Ստուգում DOCSG աղուսյակում 
  Log.Message "SQL Ստուգում DOCSG աղուսյակում", "", pmNormal, SqlDivideColor
  Call CheckQueryRowCount("DOCSG", "fISN", Leasing.ISN, 48)
  
  'SQL Ստուգում FOLDERS աղուսյակում 
  Log.Message "SQL Ստուգում FOLDERS աղուսյակում", "", pmNormal, SqlDivideColor
  Call CheckQueryRowCount("FOLDERS", "fISN", Leasing.ISN, 3)
  for i = 0 to 2
    Call CheckDB_FOLDERS(dbo_FOLDERS(i), 1)
  next
  
  'SQL Ստուգում RESNUMBERS աղուսյակում 
  Log.Message "SQL Ստուգում RESNUMBERS աղուսյակում", "", pmNormal, SqlDivideColor
  Call CheckQueryRowCount("RESNUMBERS", "fISN", Leasing.ISN, 1)
  Call CheckDB_RESNUMBERS(Leasing.ISN, "C", Leasing.DocNum, 1)		
End	Sub

Sub Check_DB_AfterSendToVerify()
    Dim i
    'SQL Ստուգում CONTRACTS աղուսյակում 
    Log.Message "SQL Ստուգում CONTRACTS աղուսյակում", "", pmNormal, SqlDivideColor
    dbo_CONTRACTS.fDGSTATE = 101
    dbo_CONTRACTS.fDGCRDTCODE = Leasing.CreditCode
    Call CheckQueryRowCount("CONTRACTS", "fDGISN", Leasing.ISN, 1)
    Call CheckDB_CONTRACTS(dbo_CONTRACTS, 1)
  
    'SQL Ստուգում DOCLOG աղուսյակում համար
    Log.Message "SQL Ստուգում DOCLOG աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCLOG", "fISN", Leasing.ISN, 3)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "N", "1", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "M", "99", "àõÕ³ñÏí»É ¿ Ñ³ëï³ïÙ³Ý", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "C", "101", "", 1)
  
    'SQL Ստուգում DOCS աղուսյակում համար
    Log.Message "SQL Ստուգում DOCS աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCS", "fISN", Leasing.ISN, 1)
    Call CheckDB_DOCS(Leasing.ISN, "C4Diagr  ", "101", fBODY, 1)
  
    'SQL Ստուգում FOLDERS աղուսյակում 
    Log.Message "SQL Ստուգում FOLDERS աղուսյակում", "", pmNormal, SqlDivideColor
    Set dbo_FOLDERS(3) = New_DB_FOLDERS()
    dbo_FOLDERS(3).fKEY = Leasing.ISN
    dbo_FOLDERS(3).fISN = Leasing.ISN
    For i = 0 To 3
      dbo_FOLDERS(i).fNAME = "C4Diagr "
      dbo_FOLDERS(i).fSTATUS = "0"
    Next
    dbo_FOLDERS(0).fFOLDERID = "Agr." & Leasing.ISN
    dbo_FOLDERS(0).fCOM = "üÇÝ³Ýë³Ï³Ý ÉÇ½ÇÝ·"
    dbo_FOLDERS(0).fSPEC = "1üÇÝ³Ýë³Ï³Ý ÉÇ½ÇÝ·- "& Leasing.DocNum & " {ý²àôêî}"
    dbo_FOLDERS(1).fFOLDERID = "C.903824400"
    dbo_FOLDERS(1).fCOM = " üÇÝ³Ýë³Ï³Ý ÉÇ½ÇÝ·"
    dbo_FOLDERS(1).fSPEC = Leasing.DocNum & " (ý²àôêî),     10652.8 - Ð³ÛÏ³Ï³Ý ¹ñ³Ù"
    dbo_FOLDERS(2).fFOLDERID = "SSConf.CRC4001"
    dbo_FOLDERS(2).fCOM = "üÇÝ³Ýë³Ï³Ý ÉÇ½ÇÝ·"
    dbo_FOLDERS(2).fSTATUS = "4"
    dbo_FOLDERS(2).fSPEC = Leasing.DocNum & "          C40 20201022            0.0077  00034852"
    dbo_FOLDERS(2).fDCBRANCH = "00 "
    dbo_FOLDERS(2).fDCDEPART = "1  "
    dbo_FOLDERS(2).fECOM = "üÇÝ³Ýë³Ï³Ý ÉÇ½ÇÝ·"
    dbo_FOLDERS(3).fFOLDERID = "SSWork.CRC420201022" 
    dbo_FOLDERS(3).fCOM = "üÇÝ³Ýë³Ï³Ý ÉÇ½ÇÝ·"
    dbo_FOLDERS(3).fSPEC = Leasing.DocNum & "          C40 20201022            0.0077  00034852àõÕ³ñÏí³Í I Ñ³ëï³ïÙ³"
    dbo_FOLDERS(3).fECOM = "üÇÝ³Ýë³Ï³Ý ÉÇ½ÇÝ·"
    dbo_FOLDERS(3).fDCBRANCH = "00 "
    dbo_FOLDERS(3).fDCDEPART = "1  "
    Call CheckQueryRowCount("FOLDERS", "fISN", Leasing.ISN, 4)
    For i = 0 To 3
      Call CheckDB_FOLDERS(dbo_FOLDERS(i), 1)
    Next
End	Sub

Sub Check_DB_VerifyContract()
    Dim i  
    'SQL Ստուգում CONTRACTS աղուսյակում 
    Log.Message "SQL Ստուգում CONTRACTS աղուսյակում", "", pmNormal, SqlDivideColor
    dbo_CONTRACTS.fDGSTATE = 7
    dbo_CONTRACTS.fDGCRDTCODE = Leasing.CreditCode
    Call CheckQueryRowCount("CONTRACTS", "fDGISN", Leasing.ISN, 1)
    Call CheckDB_CONTRACTS(dbo_CONTRACTS, 1)
  
    'SQL Ստուգում CAGRACCS աղուսյակում 
    Log.Message "SQL Ստուգում CAGRACCS աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("CAGRACCS", "fAGRISN", Leasing.ISN, 1)
  
    'SQL Ստուգում DOCLOG աղուսյակում համար
    Log.Message "SQL Ստուգում DOCLOG աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCLOG", "fISN", Leasing.ISN, 5)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "N", "1", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "M", "99", "àõÕ³ñÏí»É ¿ Ñ³ëï³ïÙ³Ý", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "C", "101", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "W", "102", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "T", "7", "", 1)
  
    'SQL Ստուգում DOCP աղուսյակում  
    Log.Message "SQL Ստուգում DOCP աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCP", "fPARENTISN", Leasing.ISN, 13)

    'SQL Ստուգում DOCS աղուսյակում համար
    Log.Message "SQL Ստուգում DOCS աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCS", "fISN", Leasing.ISN, 1)
    Call CheckDB_DOCS(Leasing.ISN, "C4Diagr  ", "7", fBODY, 1)
  
    'SQL Ստուգում FOLDERS աղուսյակում 
    Log.Message "SQL Ստուգում FOLDERS աղուսյակում", "", pmNormal, SqlDivideColor
    For i = 0 to 3
        Set dbo_FOLDERSVerify(i) = New_DB_FOLDERS()
        dbo_FOLDERSVerify(i).fKEY = Leasing.ISN
        dbo_FOLDERSVerify(i).fISN = Leasing.ISN
        dbo_FOLDERSVerify(i).fNAME = "C4Diagr "
        dbo_FOLDERSVerify(i).fSTATUS = "1"
    Next
    
    dbo_FOLDERSVerify(0).fFOLDERID = "Agr." & Leasing.ISN
    dbo_FOLDERSVerify(0).fSPEC = "1üÇÝ³Ýë³Ï³Ý ÉÇ½ÇÝ·- " & Leasing.DocNum & " {ý²àôêî}"
    dbo_FOLDERSVerify(0).fCOM = "üÇÝ³Ýë³Ï³Ý ÉÇ½ÇÝ·"
    dbo_FOLDERSVerify(1).fFOLDERID = "C.903824400"
    dbo_FOLDERSVerify(1).fSPEC = Leasing.DocNum & " (ý²àôêî),     10652.8 - Ð³ÛÏ³Ï³Ý ¹ñ³Ù"
    dbo_FOLDERSVerify(1).fCOM = " üÇÝ³Ýë³Ï³Ý ÉÇ½ÇÝ·"
	
    dbo_FOLDERSVerify(1).fECOM = "1"
    dbo_FOLDERSVerify(2).fFOLDERID = "LOANREGISTER"
    dbo_FOLDERSVerify(2).fSPEC = "C43" & Leasing.DocNum & "          111                               0                                                                                                                                                             0.00                                                                                                                                                                                                                                                                                               "
    dbo_FOLDERSVerify(2).fCOM = "ý²àôêî"
    dbo_FOLDERSVerify(3).fFOLDERID = "LOANREGISTER2"
    dbo_FOLDERSVerify(3).fSPEC = "0"
    dbo_FOLDERSVerify(3).fCOM = "ý²àôêî"
    Call CheckQueryRowCount("FOLDERS", "fISN", Leasing.ISN, 4)
    For i = 0 To 3
        Call CheckDB_FOLDERS(dbo_FOLDERSVerify(i), 1)
    Next
  
    'SQL Ստուգում HIF  աղուսյակում 
    Log.Message "SQL Ստուգում HIF աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HIF", "fOBJECT", Leasing.ISN, 4)
End	Sub

Sub Check_DB_CollectFromPrev()
    Dim i  
    'SQL Ստուգում CAGRACCS աղուսյակում 
    Log.Message "SQL Ստուգում CAGRACCS աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("CAGRACCS", "fAGRISN", Leasing.ISN, 1)
  
    'SQL Ստուգում DOCLOG աղուսյակում համար
    Log.Message "SQL Ստուգում DOCLOG աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCLOG", "fISN", Leasing.ISN, 5)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "N", "1", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "M", "99", "àõÕ³ñÏí»É ¿ Ñ³ëï³ïÙ³Ý", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "C", "101", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "W", "102", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "T", "7", "", 1)
  
    'SQL Ստուգում DOCP աղուսյակում  
    Log.Message "SQL Ստուգում DOCP աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCP", "fPARENTISN", Leasing.ISN, 13)

    'SQL Ստուգում DOCS աղուսյակում համար
    Log.Message "SQL Ստուգում DOCS աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCS", "fISN", Leasing.ISN, 1)
    Call CheckDB_DOCS(Leasing.ISN, "C4Diagr  ", "7", fBODY, 1)
  
    'SQL Ստուգում FOLDERS աղուսյակում 
    Log.Message "SQL Ստուգում FOLDERS աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("FOLDERS", "fISN", Leasing.ISN, 4)
    for i = 0 to 3
        Call CheckDB_FOLDERS(dbo_FOLDERSVerify(i), 1)
    next

    'SQL Ստուգում HI աղուսյակում համար
    Log.Message "SQL Ստուգում HI աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HI", "fBASE", fBASE, 2)
    obj = Get_ColumnValueSQL("HI", "fOBJECT", "fBASE = " & fBASE & " and fDBCR = 'C'")
    Call Check_HI_CE_accounting ("2020-10-25", fBASE, "01", "443871031", "10000.00", "000", "10000.00", "FEE", "D")
    Call Check_HI_CE_accounting ("2020-10-25", fBASE, "01", obj, "10000.00", "000", "10000.00", "FEE", "C")
		
    'SQL Ստուգում HIR աղուսյակում 
    Log.Message "SQL Ստուգում HIR աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HIR", "fOBJECT", Leasing.ISN, 1)
    Call Check_HIR("2020-10-25", "R^", Leasing.ISN, "000", "10000.00", "PAY", "D")
  
    'SQL Ստուգում HIREST  աղուսյակում 
    Log.Message "SQL Ստուգում HIREST աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HIREST", "fOBJECT", "443871031", 10)

    'SQL Ստուգում HIRREST  աղուսյակում 
    Log.Message "SQL Ստուգում HIRREST աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HIRREST", "fOBJECT", Leasing.ISN, 1)
    Call CheckDB_HIRREST("R^", Leasing.ISN, "10000.00", "2020-10-25", 1)		
End	Sub

Sub Check_DB_GiveLeasing()		
    'SQL Ստուգում CONTRACTS աղուսյակում 
    Log.Message "SQL Ստուգում CONTRACTS աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("CONTRACTS", "fDGISN", Leasing.ISN, 1)
    Call CheckDB_CONTRACTS(dbo_CONTRACTS, 1)
		
    'SQL Ստուգում DOCLOG աղուսյակում համար
    Log.Message "SQL Ստուգում DOCLOG աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCLOG", "fISN", Leasing.ISN, 6)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "N", "1", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "M", "99", "àõÕ³ñÏí»É ¿ Ñ³ëï³ïÙ³Ý", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "C", "101", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "W", "102", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "T", "7", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "E", "7", "", 1)
  
    'SQL Ստուգում DOCP աղուսյակում  
    Log.Message "SQL Ստուգում DOCP աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCP", "fPARENTISN", Leasing.ISN, 14)

    'SQL Ստուգում DOCS աղուսյակում համար
    Log.Message "SQL Ստուգում DOCS աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCS", "fISN", Leasing.ISN, 1)
    fBODY = " CODE:" & Leasing.DocNum & " CRDTCODE:"&Leasing.CreditCode&" CLICOD:00034852 NAME:ý²àôêî CURRENCY:000 ACCACC:00000113032 ACCPERMAS:72110332100 ALLSUMMA:10652.8 SUMMA:10000 ISDISCOUNT:1 DATE:20201022 ACSBRANCH:00 ACSDEPART:1 ACSTYPE:C40 AUTODEBT:1 DEBTJPART1:2 DEBTJPART:0 USECLICONNSCH:0 ONLYOVERDUE:0 KINDSCALE:2 PCAGR:12.0000/365 PCNDERAUTO:0 PCPENAGR:0/1 PCPENPER:0/1 PCLOSS:0/1 CALCFINPER:0 CALCJOUTS:0 DATEGIVE:20201022 DATEBEGCHRG:20201025 CONSTPER:1 AUTODATEA:0 REFRPERSUM:0 SECTOR:U2 USAGEFIELD:01.001 AIM:00 SCHEDULE:9 GUARANTEE:9 COUNTRY:AM LRDISTR:001 REGION:010000008 PERRES:1 REDUCEOVRDDAYS:0 WEIGHTAMDRISK:0 PPRCODE:111 SUBJRISK:0 CHRGFIRSTDAY:1 AUTOCAP:0 GIVEN:1 ISNBOUT:0 PUTINLR:1 NOTCLASS:0 OTHERCOLLATERAL:0 "
    fBODY = Replace(fBODY, " ", "%")
    Call CheckDB_DOCS(Leasing.ISN, "C4Diagr  ", "7", fBODY, 1)
		
    'SQL Ստուգում DOCSG աղուսյակում համար
    Log.Message "SQL Ստուգում DOCSG աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCSG", "fISN", Leasing.ISN, 72)
		
    'SQL Ստուգում HI աղուսյակում համար
    Log.Message "SQL Ստուգում HI աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HI", "fBASE", fBASE, 2)
    obj = Get_ColumnValueSQL("HI", "fOBJECT", "fBASE = " & fBASE & " and fDBCR = 'C'")
    Call Check_HI_CE_accounting ("2020-10-25", fBASE, "01", "443871031", "10000.00", "000", "10000.00", "FEE", "D")
    Call Check_HI_CE_accounting ("2020-10-25", fBASE, "01", obj, "10000.00", "000", "10000.00", "FEE", "C")
		
    'SQL Ստուգում HIR աղուսյակում 
    Log.Message "SQL Ստուգում HIR աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HIR", "fOBJECT", Leasing.ISN, 1)
    Call Check_HIR("2020-10-25", "R^", Leasing.ISN, "000", "10000.00", "PAY", "D")
  
    'SQL Ստուգում HIREST  աղուսյակում 
    Log.Message "SQL Ստուգում HIREST աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HIREST", "fOBJECT", "443871031", 10)

    'SQL Ստուգում HIRREST  աղուսյակում 
    Log.Message "SQL Ստուգում HIRREST աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HIRREST", "fOBJECT", Leasing.ISN, 1)
    Call CheckDB_HIRREST("R^", Leasing.ISN, "10000.00", "2020-10-25", 1)		
End	Sub

Sub Check_DB_CalcPercents()		
    Dim i, fObj(11)
    'SQL Ստուգում DOCLOG աղուսյակում համար
    Log.Message "SQL Ստուգում DOCLOG աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCLOG", "fISN", Leasing.ISN, 6)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "N", "1", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "M", "99", "àõÕ³ñÏí»É ¿ Ñ³ëï³ïÙ³Ý", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "C", "101", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "W", "102", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "T", "7", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "E", "7", "", 1)
 
    'SQL Ստուգում DOCS աղուսյակում համար
    Log.Message "SQL Ստուգում DOCS աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCS", "fISN", Leasing.ISN, 1)
    Call CheckDB_DOCS(Leasing.ISN, "C4Diagr  ", "7", fBODY, 1)
		
    'SQL Ստուգում DOCSG աղուսյակում համար
    Log.Message "SQL Ստուգում DOCSG աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCSG", "fISN", Leasing.ISN, 72)
		
    'SQL Ստուգում HI աղուսյակում համար
    Log.Message "SQL Ստուգում HI աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HI", "fBASE", fBASE, 2)
    obj = Get_ColumnValueSQL("HI", "fOBJECT", "fBASE = " & fBASE & " and fDBCR = 'C'")
    Call Check_HI_CE_accounting ("2020-10-25", fBASE, "01", "443871031", "10000.00", "000", "10000.00", "FEE", "D")
    Call Check_HI_CE_accounting ("2020-10-25", fBASE, "01", obj, "10000.00", "000", "10000.00", "FEE", "C")
		
    'SQL Ստուգում HIF  աղուսյակում 
    Log.Message "SQL Ստուգում HIF աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HIF", "fOBJECT", Leasing.ISN, 5)
		
    'SQL Ստուգում HIR աղուսյակում 
    Log.Message "SQL Ստուգում HIR աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HIR", "fOBJECT", Leasing.ISN, 1)
    Call Check_HIR("2020-10-25", "R^", Leasing.ISN, "000", "10000.00", "PAY", "D")
  
    'SQL Ստուգում HIREST  աղուսյակում 
    Log.Message "SQL Ստուգում HIREST աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HIREST", "fOBJECT", "443871031", 10)

    'SQL Ստուգում HIRREST  աղուսյակում 
    Log.Message "SQL Ստուգում HIRREST աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HIRREST", "fOBJECT", Leasing.ISN, 1)
    Call CheckDB_HIRREST("R^", Leasing.ISN, "10000.00", "2020-10-25", 1)		
		
    'SQL Ստուգում HIT  աղուսյակում 
    Log.Message "SQL Ստուգում HIT աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HIT", "fACR", Leasing.ISN, 12)
End	Sub

Sub Check_DB_FadeDebt()
    'SQL Ստուգում DOCLOG աղուսյակում համար
    Log.Message "SQL Ստուգում DOCLOG աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCLOG", "fISN", Leasing.ISN, 6)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "N", "1", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "M", "99", "àõÕ³ñÏí»É ¿ Ñ³ëï³ïÙ³Ý", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "C", "101", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "W", "102", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "T", "7", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "E", "7", "", 1)
 
    'SQL Ստուգում DOCS աղուսյակում համար
    Log.Message "SQL Ստուգում DOCS աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCS", "fISN", Leasing.ISN, 1)
    Call CheckDB_DOCS(Leasing.ISN, "C4Diagr  ", "7", fBODY, 1)
		
    'SQL Ստուգում DOCSG աղուսյակում համար
    Log.Message "SQL Ստուգում DOCSG աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCSG", "fISN", Leasing.ISN, 72)
		
    'SQL Ստուգում HI աղուսյակում համար
    Log.Message "SQL Ստուգում HI աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HI", "fBASE", fBASE, 2)
    obj = Get_ColumnValueSQL("HI", "fOBJECT", "fBASE = " & fBASE & " and fDBCR = 'C'")
    Call Check_HI_CE_accounting ("2020-10-25", fBASE, "01", "443871031", "10000.00", "000", "10000.00", "FEE", "D")
    Call Check_HI_CE_accounting ("2020-10-25", fBASE, "01", obj, "10000.00", "000", "10000.00", "FEE", "C")
		
    'SQL Ստուգում HIR աղուսյակում 
    Log.Message "SQL Ստուգում HIR աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HIR", "fOBJECT", Leasing.ISN, 1)
    Call Check_HIR("2020-10-25", "R^", Leasing.ISN, "000", "10000.00", "PAY", "D")
  
    'SQL Ստուգում HIREST  աղուսյակում 
    Log.Message "SQL Ստուգում HIREST աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HIREST", "fOBJECT", "443871031", 10)

    'SQL Ստուգում HIRREST  աղուսյակում 
    Log.Message "SQL Ստուգում HIRREST աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HIRREST", "fOBJECT", Leasing.ISN, 1)
    Call CheckDB_HIRREST("R^", Leasing.ISN, "10000.00", "2020-10-25", 1)		
End	Sub

Sub Check_DB_ChangeRate()
    Dim i
    'SQL Ստուգում CONTRACTS աղուսյակում 
    Log.Message "SQL Ստուգում CONTRACTS աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("CONTRACTS", "fDGISN", Leasing.ISN, 1)
    Call CheckDB_CONTRACTS(dbo_CONTRACTS, 1)

    'SQL Ստուգում DOCLOG աղուսյակում համար
    Log.Message "SQL Ստուգում DOCLOG աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCLOG", "fISN", Leasing.ISN, 7)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "N", "1", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "M", "99", "àõÕ³ñÏí»É ¿ Ñ³ëï³ïÙ³Ý", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "C", "101", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "W", "102", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "T", "7", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "E", "7", "", 2)
 
    'SQL Ստուգում DOCS աղուսյակում համար
    Log.Message "SQL Ստուգում DOCS աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCS", "fISN", Leasing.ISN, 1)
    Call CheckDB_DOCS(Leasing.ISN, "C4Diagr  ", "7", fBODY, 1)
		
    'SQL Ստուգում FOLDERS աղուսյակում 
    Log.Message "SQL Ստուգում FOLDERS աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("FOLDERS", "fISN", Leasing.ISN, 4)
    For i = 0 To 3
        Call CheckDB_FOLDERS(dbo_FOLDERSVerify(i), 1)
    Next
End	Sub

Sub Check_DB_ChangeEffRete()
    'SQL Ստուգում DOCLOG աղուսյակում համար
    Log.Message "SQL Ստուգում DOCLOG աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCLOG", "fISN", Leasing.ISN, 7)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "N", "1", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "M", "99", "àõÕ³ñÏí»É ¿ Ñ³ëï³ïÙ³Ý", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "C", "101", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "W", "102", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "T", "7", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "E", "7", "", 2)
 
    'SQL Ստուգում DOCS աղուսյակում համար
    Log.Message "SQL Ստուգում DOCS աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCS", "fISN", Leasing.ISN, 1)
    Call CheckDB_DOCS(Leasing.ISN, "C4Diagr  ", "7", fBODY, 1)
		
    'SQL Ստուգում HIF  աղուսյակում 
    Log.Message "SQL Ստուգում HIF աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HIF", "fOBJECT", Leasing.ISN, 7)
End	Sub

Sub Check_DB_CalcPercents2()
    'SQL Ստուգում DOCLOG աղուսյակում համար
    Log.Message "SQL Ստուգում DOCLOG աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCLOG", "fISN", Leasing.ISN, 7)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "N", "1", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "M", "99", "àõÕ³ñÏí»É ¿ Ñ³ëï³ïÙ³Ý", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "C", "101", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "W", "102", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "T", "7", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "E", "7", "", 2)
 
    'SQL Ստուգում DOCS աղուսյակում համար
    Log.Message "SQL Ստուգում DOCS աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCS", "fISN", Leasing.ISN, 1)
    Call CheckDB_DOCS(Leasing.ISN, "C4Diagr  ", "7", fBODY, 1)
		
    'SQL Ստուգում DOCSG աղուսյակում համար
    Log.Message "SQL Ստուգում DOCSG աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCSG", "fISN", Leasing.ISN, 72)
		
    'SQL Ստուգում HI աղուսյակում համար
    Log.Message "SQL Ստուգում HI աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HI", "fBASE", fBASE, 2)
    obj = Get_ColumnValueSQL("HI", "fOBJECT", "fBASE = " & fBASE & " and fDBCR = 'C'")
    Call Check_HI_CE_accounting ("2020-10-25", fBASE, "01", "443871031", "10000.00", "000", "10000.00", "FEE", "D")
    Call Check_HI_CE_accounting ("2020-10-25", fBASE, "01", obj, "10000.00", "000", "10000.00", "FEE", "C")
End	Sub

Sub Check_DB_ObjRisk()
    'SQL Ստուգում DOCLOG աղուսյակում համար
    Log.Message "SQL Ստուգում DOCLOG աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCLOG", "fISN", Leasing.ISN, 7)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "N", "1", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "M", "99", "àõÕ³ñÏí»É ¿ Ñ³ëï³ïÙ³Ý", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "C", "101", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "W", "102", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "T", "7", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "E", "7", "", 2)
 
    'SQL Ստուգում DOCS աղուսյակում համար
    Log.Message "SQL Ստուգում DOCS աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCS", "fISN", Leasing.ISN, 1)
    Call CheckDB_DOCS(Leasing.ISN, "C4Diagr  ", "7", fBODY, 1)
End	Sub

Sub Check_DB_RiskClassifier()
    'SQL Ստուգում DOCLOG աղուսյակում համար
    Log.Message "SQL Ստուգում DOCLOG աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCLOG", "fISN", Leasing.ISN, 7)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "N", "1", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "M", "99", "àõÕ³ñÏí»É ¿ Ñ³ëï³ïÙ³Ý", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "C", "101", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "W", "102", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "T", "7", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "E", "7", "", 2)
 
    'SQL Ստուգում DOCS աղուսյակում համար
    Log.Message "SQL Ստուգում DOCS աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCS", "fISN", Leasing.ISN, 1)
    Call CheckDB_DOCS(Leasing.ISN, "C4Diagr  ", "7", fBODY, 1)
End	Sub

Sub Check_DB_Backup()
    'SQL Ստուգում DOCLOG աղուսյակում համար
    Log.Message "SQL Ստուգում DOCLOG աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCLOG", "fISN", Leasing.ISN, 7)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "N", "1", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "M", "99", "àõÕ³ñÏí»É ¿ Ñ³ëï³ïÙ³Ý", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "C", "101", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "W", "102", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "T", "7", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "E", "7", "", 2)
 
    'SQL Ստուգում DOCS աղուսյակում համար
    Log.Message "SQL Ստուգում DOCS աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCS", "fISN", Leasing.ISN, 1)
    Call CheckDB_DOCS(Leasing.ISN, "C4Diagr  ", "7", fBODY, 1)
		
    'SQL Ստուգում DOCSG աղուսյակում համար
    Log.Message "SQL Ստուգում DOCSG աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCSG", "fISN", Leasing.ISN, 72)
		
    'SQL Ստուգում HI աղուսյակում համար
    Log.Message "SQL Ստուգում HI աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HI", "fBASE", fBASE, 2)
    obj = Get_ColumnValueSQL("HI", "fOBJECT", "fBASE = " & fBASE & " and fDBCR = 'C'")
    Call Check_HI_CE_accounting ("2020-10-25", fBASE, "01", "443871031", "10000.00", "000", "10000.00", "FEE", "D")
    Call Check_HI_CE_accounting ("2020-10-25", fBASE, "01", obj, "10000.00", "000", "10000.00", "FEE", "C")
		
    'SQL Ստուգում HIREST  աղուսյակում 
    Log.Message "SQL Ստուգում HIREST աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HIREST", "fOBJECT", "443871031", 10)
End	Sub

Sub Check_DB_WriteOut()
    'SQL Ստուգում CAGRACCS աղուսյակում 
    Log.Message "SQL Ստուգում CAGRACCS աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("CAGRACCS", "fAGRISN", Leasing.ISN, 1)

    'SQL Ստուգում DOCLOG աղուսյակում համար
    Log.Message "SQL Ստուգում DOCLOG աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCLOG", "fISN", Leasing.ISN, 7)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "N", "1", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "M", "99", "àõÕ³ñÏí»É ¿ Ñ³ëï³ïÙ³Ý", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "C", "101", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "W", "102", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "T", "7", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "E", "7", "", 2)
 
    'SQL Ստուգում DOCP աղուսյակում  
    Log.Message "SQL Ստուգում DOCP աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCP", "fPARENTISN", Leasing.ISN, 14)
		
    'SQL Ստուգում DOCS աղուսյակում համար
    Log.Message "SQL Ստուգում DOCS աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCS", "fISN", Leasing.ISN, 1)
    Call CheckDB_DOCS(Leasing.ISN, "C4Diagr  ", "7", fBODY, 1)
		
    'SQL Ստուգում DOCSG աղուսյակում համար
    Log.Message "SQL Ստուգում DOCSG աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCSG", "fISN", Leasing.ISN, 72)
		
    'SQL Ստուգում FOLDERS աղուսյակում 
    Log.Message "SQL Ստուգում FOLDERS աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("FOLDERS", "fISN", Leasing.ISN, 4)
		
    'SQL Ստուգում HI աղուսյակում համար
    Log.Message "SQL Ստուգում HI աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HI", "fBASE", fBASE, 8)
    'Call Check_HI_CE_accounting ("2020-11-26", fBASE, "01", "*", "9210.80", "000", "9210.80", "MSC", "D")
    'Call Check_HI_CE_accounting ("2020-11-26", fBASE, "01", "*", "3.30", "000", "3.30", "MSC", "D")
    'Call Check_HI_CE_accounting ("2020-11-26", fBASE, "01", "*", "3.30", "000", "3.30", "MSC", "C")
    'Call Check_HI_CE_accounting ("2020-11-26", fBASE, "01", "*", "554.20", "000", "554.20", "MSC", "D")
    'Call Check_HI_CE_accounting ("2020-11-26", fBASE, "01", "*", "9210.80", "000", "9210.80", "MSC", "C")
    'Call Check_HI_CE_accounting ("2020-11-26", fBASE, "01", "*", "554.20", "000", "554.20", "MSC", "C")
    'Call Check_HI_CE_accounting ("2020-11-26", fBASE, "02", "*", "3.30", "000", "3.30", "MSC", "D")
    'Call Check_HI_CE_accounting ("2020-11-26", fBASE, "02", "*", "9210.80", "000", "9210.80", "MSC", "D")
		
    'SQL Ստուգում HIR աղուսյակում 
    Log.Message "SQL Ստուգում HIR աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HIR", "fOBJECT", Leasing.ISN, 1)
    Call Check_HIR("2020-10-25", "R^", Leasing.ISN, "000", "10000.00", "PAY", "D")
		
    'SQL Ստուգում HIREST  աղուսյակում 
    Log.Message "SQL Ստուգում HIREST աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HIREST", "fOBJECT", "443871031", 10)
		
    'SQL Ստուգում HIRREST  աղուսյակում 
    Log.Message "SQL Ստուգում HIRREST աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HIRREST", "fOBJECT", Leasing.ISN, 1)
    Call CheckDB_HIRREST("R^", Leasing.ISN, "10000.00", "2020-10-25", 1)		
End	Sub

Sub Check_DB_WriteOffRec()
    'SQL Ստուգում DOCLOG աղուսյակում համար
    Log.Message "SQL Ստուգում DOCLOG աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCLOG", "fISN", Leasing.ISN, 7)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "N", "1", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "M", "99", "àõÕ³ñÏí»É ¿ Ñ³ëï³ïÙ³Ý", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "C", "101", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "W", "102", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "T", "7", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "E", "7", "", 2)
 
    'SQL Ստուգում DOCS աղուսյակում համար
    Log.Message "SQL Ստուգում DOCS աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCS", "fISN", Leasing.ISN, 1)
    Call CheckDB_DOCS(Leasing.ISN, "C4Diagr  ", "7", fBODY, 1)
		
    'SQL Ստուգում DOCSG աղուսյակում համար
    Log.Message "SQL Ստուգում DOCSG աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCSG", "fISN", Leasing.ISN, 72)
	
    'SQL Ստուգում HI աղուսյակում համար
    Log.Message "SQL Ստուգում HI աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HI", "fBASE", fBASE, 8)
    'Call Check_HI_CE_accounting ("2020-11-26", fBASE, "01", "*", "9210.80", "000", "9210.80", "MSC", "D")
    'Call Check_HI_CE_accounting ("2020-11-26", fBASE, "01", "*", "3.30", "000", "3.30", "MSC", "D")
    'Call Check_HI_CE_accounting ("2020-11-26", fBASE, "01", "*", "3.30", "000", "3.30", "MSC", "C")
    'Call Check_HI_CE_accounting ("2020-11-26", fBASE, "01", "*", "554.20", "000", "554.20", "MSC", "D")
    'Call Check_HI_CE_accounting ("2020-11-26", fBASE, "01", "*", "9210.80", "000", "9210.80", "MSC", "C")
    'Call Check_HI_CE_accounting ("2020-11-26", fBASE, "01", "*", "554.20", "000", "554.20", "MSC", "C")
    'Call Check_HI_CE_accounting ("2020-11-26", fBASE, "02", "*", "3.30", "000", "3.30", "MSC", "D")
    'Call Check_HI_CE_accounting ("2020-11-26", fBASE, "02", "*", "9210.80", "000", "9210.80", "MSC", "D")
		
    'SQL Ստուգում HIR աղուսյակում 
    Log.Message "SQL Ստուգում HIR աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HIR", "fOBJECT", Leasing.ISN, 1)
    Call Check_HIR("2020-10-25", "R^", Leasing.ISN, "000", "10000.00", "PAY", "D")
		
    'SQL Ստուգում HIREST  աղուսյակում 
    Log.Message "SQL Ստուգում HIREST աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HIREST", "fOBJECT", "443871031", 10)
		
    'SQL Ստուգում HIRREST  աղուսյակում 
    Log.Message "SQL Ստուգում HIRREST աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HIRREST", "fOBJECT", Leasing.ISN, 1)
    Call CheckDB_HIRREST("R^", Leasing.ISN, "10000.00", "2020-10-25", 1)		
End	Sub

Sub Check_DB_FadeDebt2()
    'SQL Ստուգում DOCLOG աղուսյակում համար
    Log.Message "SQL Ստուգում DOCLOG աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCLOG", "fISN", Leasing.ISN, 7)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "N", "1", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "M", "99", "àõÕ³ñÏí»É ¿ Ñ³ëï³ïÙ³Ý", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "C", "101", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "W", "102", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "T", "7", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "E", "7", "", 2)
 
    'SQL Ստուգում DOCS աղուսյակում համար
    Log.Message "SQL Ստուգում DOCS աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCS", "fISN", Leasing.ISN, 1)
    Call CheckDB_DOCS(Leasing.ISN, "C4Diagr  ", "7", fBODY, 1)
		
    'SQL Ստուգում DOCSG աղուսյակում համար
    Log.Message "SQL Ստուգում DOCSG աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCSG", "fISN", Leasing.ISN, 72)
	
    'SQL Ստուգում HI աղուսյակում համար
    Log.Message "SQL Ստուգում HI աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HI", "fBASE", fBASE, 8)
    'Call Check_HI_CE_accounting ("2020-11-26", fBASE, "01", "*", "9210.80", "000", "9210.80", "MSC", "D")
    'Call Check_HI_CE_accounting ("2020-11-26", fBASE, "01", "*", "3.30", "000", "3.30", "MSC", "D")
    'Call Check_HI_CE_accounting ("2020-11-26", fBASE, "01", "*", "3.30", "000", "3.30", "MSC", "C")
    'Call Check_HI_CE_accounting ("2020-11-26", fBASE, "01", "*", "554.20", "000", "554.20", "MSC", "D")
    'Call Check_HI_CE_accounting ("2020-11-26", fBASE, "01", "*", "9210.80", "000", "9210.80", "MSC", "C")
    'Call Check_HI_CE_accounting ("2020-11-26", fBASE, "01", "*", "554.20", "000", "554.20", "MSC", "C")
    'Call Check_HI_CE_accounting ("2020-11-26", fBASE, "02", "*", "3.30", "000", "3.30", "MSC", "D")
    'Call Check_HI_CE_accounting ("2020-11-26", fBASE, "02", "*", "9210.80", "000", "9210.80", "MSC", "D")
		
    'SQL Ստուգում HIR աղուսյակում 
    Log.Message "SQL Ստուգում HIR աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HIR", "fOBJECT", Leasing.ISN, 1)
    Call Check_HIR("2020-10-25", "R^", Leasing.ISN, "000", "10000.00", "PAY", "D")
		
    'SQL Ստուգում HIREST  աղուսյակում 
    Log.Message "SQL Ստուգում HIREST աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HIREST", "fOBJECT", "443871031", 10)
		
    'SQL Ստուգում HIRREST  աղուսյակում 
    Log.Message "SQL Ստուգում HIRREST աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HIRREST", "fOBJECT", Leasing.ISN, 1)
    Call CheckDB_HIRREST("R^", Leasing.ISN, "10000.00", "2020-10-25", 1)		
End	Sub

Sub Check_DB_Delete_AgreementAllOperations()
    'SQL Ստուգում DOCLOG աղուսյակում համար
    Log.Message "SQL Ստուգում DOCLOG աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCLOG", "fISN", Leasing.ISN, 9)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "N", "1", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "M", "99", "àõÕ³ñÏí»É ¿ Ñ³ëï³ïÙ³Ý", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "C", "101", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "W", "102", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "T", "7", "", 1)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "E", "7", "", 3)
    Call CheckDB_DOCLOG(Leasing.ISN, "77", "D", "999", "", 1)
				
    'SQL Ստուգում DOCS աղուսյակում համար
    Log.Message "SQL Ստուգում DOCS աղուսյակում", "", pmNormal, SqlDivideColor
    fBODY = " CODE:" & Leasing.DocNum & " CRDTCODE:"&Leasing.CreditCode&" CLICOD:00034852 NAME:ý²àôêî CURRENCY:000 ACCACC:00000113032 ACCPERMAS:72110332100 ALLSUMMA:10652.8 SUMMA:10000 ISDISCOUNT:1 DATE:20201022 ACSBRANCH:00 ACSDEPART:1 ACSTYPE:C40 AUTODEBT:1 DEBTJPART1:2 DEBTJPART:0 USECLICONNSCH:0 ONLYOVERDUE:0 KINDSCALE:2 PCAGR:12.0000/365 PCNDERAUTO:0 PCPENAGR:0/1 PCPENPER:0/1 PCLOSS:0/1 CALCFINPER:0 CALCJOUTS:0 DATEGIVE:20201022 DATEBEGCHRG:20201025 CONSTPER:1 AUTODATEA:0 REFRPERSUM:0 SECTOR:U2 USAGEFIELD:01.001 AIM:00 SCHEDULE:9 GUARANTEE:9 COUNTRY:AM LRDISTR:001 REGION:010000008 PERRES:1 REDUCEOVRDDAYS:0 WEIGHTAMDRISK:0 PPRCODE:111 SUBJRISK:0 CHRGFIRSTDAY:1 AUTOCAP:0 GIVEN:0 ISNBOUT:0 PUTINLR:1 NOTCLASS:0 OTHERCOLLATERAL:0 "
    fBODY = Replace(fBODY, " ", "%")
    Call CheckQueryRowCount("DOCS", "fISN", Leasing.ISN, 1)
    Call CheckDB_DOCS(Leasing.ISN, "C4Diagr ", "999", fBODY, 1)
		
    'SQL Ստուգում FOLDERS աղուսյակում 
    Log.Message "SQL Ստուգում FOLDERS աղուսյակում", "", pmNormal, SqlDivideColor
    dbo_FOLDERSVerify(0).fKEY = Leasing.ISN
    dbo_FOLDERSVerify(0).fISN = Leasing.ISN
    dbo_FOLDERSVerify(0).fNAME = "C4Diagr "
    dbo_FOLDERSVerify(0).fSTATUS = "0"
    dbo_FOLDERSVerify(0).fFOLDERID = ".R." & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%Y%m%d")' & Leasing.ISN
    dbo_FOLDERSVerify(0).fSPEC = Left_Align(Get_Compname_DOCLOG(Leasing.ISN), 16) & "Cred&DepARMSOFT                       007  "
    dbo_FOLDERSVerify(0).fCOM = ""
    dbo_FOLDERSVerify(0).fDCBRANCH	= "00 "
    dbo_FOLDERSVerify(0).fECOM = ""
    dbo_FOLDERSVerify(0).fDCDEPART = "1  "
    Call CheckQueryRowCount("FOLDERS", "fISN", Leasing.ISN, 1)
    Call CheckDB_FOLDERS(dbo_FOLDERSVerify(0), 1)
End	Sub