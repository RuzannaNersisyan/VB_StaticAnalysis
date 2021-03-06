'USEUNIT Library_Common 
'USEUNIT Library_Colour
'USEUNIT Constants
'USEUNIT Loan_Agreements_Library 
'USEUNIT Overlimit_Library
'USEUNIT Akreditiv_Library
'USEUNIT Subsystems_SQL_Library
'USEUNIT Library_CheckDB
'USEUNIT Subsystems_Special_Library
'USEUNIT Library_Contracts

Option Explicit
'Test Case Id - 145724
Dim AccWithOverlimit,ContractFillter
Dim fADB

Sub Check_OverlimitFromAccount()
    
    Dim sDATE,fDATE
    Call Test_InitializeFor_OverlimitFromAccount()
     
    'Համակարգ մուտք գործել ARMSOFT օգտագործողով
    sDATE = "20140101"
    fDATE = "20201125"
    Call Initialize_AsBank("bank", sDATE, fDATE)
    Login("ARMSOFT")

    'Մուտք գործել "Գերածախս"
    Call ChangeWorkspace(c_Overlimit) 
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''-Բացել "Գերածախս ունեցող հաշիվներ" թղթապանակը-'''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "Open Accounts with Overlimit Doc",,,DivideColor
    
    Call wTreeView.DblClickItem("|¶»ñ³Í³Ëë|¶»ñ³Í³Ëë áõÝ»óáÕ Ñ³ßÇíÝ»ñ|")
    BuiltIn.Delay(delay_middle)
    Call Fill_AccWithOverlimit(AccWithOverlimit)
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''-Հայտնցած տողի վրա կատարել աջ կլիկ - Գերածախսի բացում (խմբ.)-''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "Right Click - Open Overlimit Action",,,DivideColor    
    
    AccountIsn = OpenOverimitFromAccount("251120")
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''-Պայմանագրեր թղթապանակում փաստատթղթի առկայության ստուգում-'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "Check (Existing Contract) Function",,,DivideColor
    
    Call ExistsContract_Filter_Fill("|¶»ñ³Í³Ëë|",ContractFillter,1)
    AccountParentIsn = GetAccountIsnOverlimit()
    fISN = GetIsn()

    Log.Message "SQL Check After Right Click - Open Overlimit Action",,,SqlDivideColor
    Log.Message "fISN = "& fISN,,,SqlDivideColor 
    Log.Message "AccountParentIsn = "& AccountParentIsn,,,SqlDivideColor    
    
    Call SQL_Initialize_OverlimitFromAccount(fISN,AccountParentIsn)
    
    'SQL Ստուգում DOCS աղուսյակում  
    fBODY = "  CODE:01046643311  CRDTCODE:777000000227L001  CLICOD:00000022  NAME:²·³Ã ö³ÛÉ³ï³ÏÛ³Ý  CURRENCY:000  ACCACC:01046643311  AUTODEBT:1  ACCCONNMODE:3  USECLICONNSCH:0  DATE:20201125  DATEGIVE:20201125  ACSBRANCH:00  ACSDEPART:2  ACSTYPE:CO1  KINDSCALE:1  PCPENAGR:0/1  CONSTPER:0  SECTOR:F  SCHEDULE:9  GUARANTEE:9  PERRES:1  PPRCODE:01046643311  SUBJRISK:0  CHRGFIRSTDAY:1  GIVEN:1  PUTINLR:0  NOTCLASS:0  OTHERCOLLATERAL:0  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",fISN,1)
    Call CheckDB_DOCS(fISN,"COSimpl","7",fBODY,1)

    'SQL Ստուգում DOCS աղուսյակում ¶»ñ³Í³ËëÇ Ñ³ßí³å³Ñ³Ï³Ý Ñ³í»Éí³Í-Ի համար
    fBODY = "  CODE:01046643311  CURRENCY:000  CLICOD:00000022  JURSTAT:21  VOLORT:7  PETBUJ:2  REZ:1  RELBANK:0  RABBANK:0  ACCAGR:01080793012  ACCACC:01046643311  FILLACCS:0  OPENACCS:0  TYPEPEN:0  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",AccountParentIsn,1)
    Call CheckDB_DOCS(AccountParentIsn,"COAgrAcc","2",fBODY,1)
    
    'SQL Ստուգում DOCLOG աղուսյակում 
    Call CheckQueryRowCount("DOCLOG","fISN",fISN,4)
    Call CheckDB_DOCLOG(fISN,"77","N","1","",1)
    Call CheckDB_DOCLOG(fISN,"77","T","7","",1)
    Call CheckDB_DOCLOG(fISN,"77","E","7","",2)
    
    'SQL Ստուգում DOCLOG աղուսյակում ¶»ñ³Í³ËëÇ Ñ³ßí³å³Ñ³Ï³Ý Ñ³í»Éí³Í-Ի համար
    Call CheckQueryRowCount("DOCLOG","fISN",AccountParentIsn,3)
    Call CheckDB_DOCLOG(AccountParentIsn,"77","N","4","",1)
    Call CheckDB_DOCLOG(AccountParentIsn,"77","C","2","",1)
    Call CheckDB_DOCLOG(AccountParentIsn,"77","E","2","",1)
    
    'SQL Ստուգում DOCSG աղուսյակում 
    Call CheckQueryRowCount("DOCSG","fISN",AccountParentIsn,8)
    Call CheckDB_DOCSG(AccountParentIsn,"ACCSRES","0","ACCRES","00000453201",1)
    Call CheckDB_DOCSG(AccountParentIsn,"ACCSRES","0","ACCRESEXP","73030381000",1)
    Call CheckDB_DOCSG(AccountParentIsn,"ACCSRES","0","RISK","01",1)
    Call CheckDB_DOCSG(AccountParentIsn,"ACCSRES","1","RISK","02",1)
    Call CheckDB_DOCSG(AccountParentIsn,"ACCSRES","2","RISK","03",1)
    Call CheckDB_DOCSG(AccountParentIsn,"ACCSRES","3","RISK","04",1)
    Call CheckDB_DOCSG(AccountParentIsn,"ACCSRES","4","RISK","05",1)
    
    'SQL Ստուգում DOCP աղուսյակում  
    Call CheckQueryRowCount("DOCP","fPARENTISN",AccountParentIsn,5)
    
    'SQL Ստուգում ACCOUNTS  և HIREST աղուսյակներում
    'Row 1
    Query = "SELECT fISN FROM (SELECT *, ROW_NUMBER() OVER (ORDER BY fDCD ASC) AS rownum FROM DOCP where fPARENTISN = "&AccountParentIsn&" and fISN <> 1630358) AS foo WHERE  rownum > 0 AND rownum <= 1 "
    AccountIsn = my_Row_Count(Query) 
    Call CheckQueryRowCount("ACCOUNTS","fISN",AccountIsn,1)
    Call CheckDB_DOCP(AccountIsn,"Acc     ",AccountParentIsn,1)
    Call CheckDB_HIREST("01", AccountIsn,"999970.70","000","999970.70",1)    
    'Row 2
    Query = "SELECT fISN FROM (SELECT *, ROW_NUMBER() OVER (ORDER BY fDCD ASC) AS rownum FROM DOCP where fPARENTISN = "&AccountParentIsn&" and fISN <> 1630358) AS foo WHERE  rownum > 1 AND rownum <= 2 "
    AccountIsn = my_Row_Count(Query) 
    Call CheckQueryRowCount("DOCP","fPARENTISN",AccountParentIsn,5)
    Call CheckDB_DOCP(AccountIsn,"Acc     ",AccountParentIsn,1)
    Call CheckQueryRowCount("HIREST","fOBJECT",AccountIsn,1)
    'Row 3
    Query = "SELECT fISN FROM (SELECT *, ROW_NUMBER() OVER (ORDER BY fDCD ASC) AS rownum FROM DOCP where fPARENTISN = "&AccountParentIsn&" and fISN <> 1630358) AS foo WHERE  rownum > 2 AND rownum <= 3 "
    AccountIsn = my_Row_Count(Query) 
    Call CheckQueryRowCount("ACCOUNTS","fISN",AccountIsn,1)
    Call CheckDB_DOCP(AccountIsn,"Acc     ",AccountParentIsn,1)
    Call CheckQueryRowCount("HIREST","fOBJECT",AccountIsn,1)
    'Row 4
    Query = "SELECT fISN FROM (SELECT *, ROW_NUMBER() OVER (ORDER BY fDCD ASC) AS rownum FROM DOCP where fPARENTISN = "&AccountParentIsn&" and fISN <> 1630358) AS foo WHERE  rownum > 3 AND rownum <= 4 "
    AccountIsn = my_Row_Count(Query) 
    Call CheckQueryRowCount("ACCOUNTS","fISN",AccountIsn,1)
    Call CheckDB_DOCP(AccountIsn,"Acc     ",AccountParentIsn,1)
    Call CheckQueryRowCount("HIREST","fOBJECT",AccountIsn,1)
    
    'SQL Ստուգում CONTRACTS աղուսյակում 
    Call CheckQueryRowCount("CONTRACTS","fDGISN",fISN,1)
    Call CheckDB_CONTRACTS(dbCONTRACT,1)
    
    'SQL Ստուգում CAGRACCS աղուսյակում 
    Call CheckQueryRowCount("CAGRACCS","fAGRISN",fISN,1)
    
    'SQL Ստուգում FOLDERS աղուսյակում 
    Call CheckQueryRowCount("FOLDERS","fISN",fISN,3)
    Call CheckDB_FOLDERS(dbFOLDERS(1),1)
    Call CheckDB_FOLDERS(dbFOLDERS(2),1)
    Call CheckDB_FOLDERS(dbFOLDERS(3),1)

    Call CheckQueryRowCount("FOLDERS","fISN",AccountParentIsn,2)
    Call CheckDB_FOLDERS(dbFOLDERS(7),1)
    Call CheckDB_FOLDERS(dbFOLDERS(8),1)
    
    'SQL Ստուգում HIF  աղուսյակում 
    Call Check_HIF("2020-11-25", "N0", fISN, "0.00", "1.00", "PPA", Null)
    Call Check_HIF("2020-11-25", "N0", fISN, "0.00", "0.00", "LIM", Null)
   
   'SQL Ստուգում RESNUMBERS  աղուսյակում 
    Call CheckDB_RESNUMBERS(fISN,"C","01046643311   ",1)

    'SQL Ստուգում HI աղուսյակում 
    Query = "Select fBASE from HI WHERE fSUM = '999970.70'"
    fBASE(0) = my_Row_Count(Query)
    Query = "Select fADB from HI WHERE fSUM = '999970.70'"
    fADB = my_Row_Count(Query)
    Call CheckQueryRowCount("HI","fBASE",fBASE(0),2)
    Call Check_HI_CE_accounting ("20201125",fBASE(0), "01", "1630358", "999970.70", "000", "999970.70", "OVD", "C") 
    Call Check_HI_CE_accounting ("20201125",fBASE(0), "01", fADB, "999970.70", "000", "999970.70", "OVD", "D") 

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''-"Տոկոսների հաշվարկ" գործողության կատարում-''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''    
    Log.Message "Check RC (Calculate Percents/Տոկոսների հաշվարկ) Function",,,DivideColor
    
    Call CalculatePercents(CalcPercents,"",False)
    
    Log.Message "SQL Check After RC (Calculate Percents/Տոկոսների հաշվարկ) Function",,,SqlDivideColor
    Log.Message "fISN = " & CalcPercents.Isn,,,SqlDivideColor
    
    'SQL Ստուգում DOCS աղուսյակում 
    fBODY = "  CODE:01046643311  DATECHARGE:20201125  DATE:20201125  SUMAGRPEN:22222222.1/11111111.1  SUMALLPEN:22222222.1/11111111.1  COMMENT:NewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOut1  ACSBRANCH:01  ACSDEPART:4  ACSTYPE:CO1  USERID:  77  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",CalcPercents.Isn,1)
    Call CheckDB_DOCS(CalcPercents.Isn,"CODSChrg","5",fBODY,1)

    'SQL Ստուգում DOCLOG աղուսյակում 
    Call CheckQueryRowCount("DOCLOG","fISN",CalcPercents.Isn,4)
    Call CheckDB_DOCLOG(CalcPercents.Isn,"77","N","1","",1)
    Call CheckDB_DOCLOG(CalcPercents.Isn,"77","M","99","àõÕ³ñÏí»É ¿ Ñ³ëï³ïÙ³Ý",1)
    Call CheckDB_DOCLOG(CalcPercents.Isn,"77","T","2","",1)
    Call CheckDB_DOCLOG(CalcPercents.Isn,"77","C","5","",1)
    
    'SQL Ստուգում HIF աղուսյակում 
    Call Check_HIF ("20201125", "N0", fISN, "0.00", "0.00", "AGJ", "1")
    Call Check_HIF ("20201125", "N0", fISN, "0.00", "0.00", "DTC", "20201125")
    Call Check_HIF ("20201125", "N0", fISN, "0.00", "0.00", "LIM", Null)
    Call Check_HIF ("20201125", "N0", fISN, "0.00", "1.00", "PPA", Null)
    
    'SQL Ստուգում HIR աղուսյակում 
    Call CheckQueryRowCount("HIR","fOBJECT",fISN,4)
    Call Check_HIR("20201125", "R1", fISN, "000", "999970.70", "AGR", "D")
    Call Check_HIR("20201125", "R3", fISN, "000", "22222222.10", "PNA", "D")
    Call Check_HIR("20201125", "R7", fISN, "000", "11111111.10", "PNA", "D")
    Call Check_HIR("20201125", "RÄ", fISN, "000", "999970.70", "AGJ", "D")
    
    'SQL Ստուգում HIT աղուսյակում 
    Call CheckQueryRowCount("HIT","fOBJECT",fISN,2)
    Call Check_HIT("20201125", "N3", fISN, "000", "22222222.10", "PNA", "D")
    Call Check_HIT("20201125", "N7", fISN, "000", "11111111.10", "PNA", "D")
    
    'SQL Ստուգում HIRREST  աղուսյակում 
    Call CheckQueryRowCount("HIRREST","fOBJECT",fISN,4)
    Call CheckDB_HIRREST("R3",fISN,"22222222.10","20201125",1)
    Call CheckDB_HIRREST("R7",fISN,"11111111.10","20201125",1)    
   
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''-"Տոկոսադրույքներ" գործողության կատարում-'''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''      
    Log.Message "Check RC (Overlimit Rete) Function",,,DivideColor    
    
    ActionIsn(1) = ChangeOverlimitRete("01046643311", "251120", "1092.1001", "29","ChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChan1")

    Log.Message "SQL Check After RC (Overlimit Rete/Տոկոսադրույքներ) Function",,,SqlDivideColor
    Log.Message "fISN = " & ActionIsn(1),,,SqlDivideColor
    
    'SQL Ստուգում DOCS աղուսյակում  
    fBODY = "  CODE:01046643311  DATE:20201125  PCPENAGR:1092.1001/29  COMMENT:ChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChan1  USERID:  77  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",ActionIsn(1),1)
    Call CheckDB_DOCS(ActionIsn(1),"COTSPC  ","5",fBODY,1)
    
    'SQL Ստուգում DOCLOG աղուսյակում 
    Call CheckQueryRowCount("DOCLOG","fISN",fISN,4)
    Call CheckQueryRowCount("DOCLOG","fISN",ActionIsn(1),4)
    Call CheckDB_DOCLOG(ActionIsn(1),"77","N","1","",1)
    Call CheckDB_DOCLOG(ActionIsn(1),"77","M","99","àõÕ³ñÏí»É ¿ Ñ³ëï³ïÙ³Ý",1)
    Call CheckDB_DOCLOG(ActionIsn(1),"77","T","2","",1)
    Call CheckDB_DOCLOG(ActionIsn(1),"77","C","5","",1)
    
    'SQL Ստուգում HIF աղուսյակում 
    Call Check_HIF ("20201125", "N0", fISN, "1092.1001", "29.00", "PPA", Null)
    
    'SQL Ստուգում FOLDERS  աղուսյակում 0
    Set dbFOLDERS_ForRate = New_DB_FOLDERS()
        dbFOLDERS_ForRate.fFOLDERID = "Agr." & fISN
        dbFOLDERS_ForRate.fNAME = "COTSPC  "
        dbFOLDERS_ForRate.fKEY = ActionIsn(1)
        dbFOLDERS_ForRate.fISN = ActionIsn(1)
        dbFOLDERS_ForRate.fSTATUS = "1"
        dbFOLDERS_ForRate.fCOM = "îáÏáë³¹ñáõÛùÝ»ñ"
        dbFOLDERS_ForRate.fSPEC = "1îáÏáë³¹ñáõÛùÝ»ñ`  25/11/20,  { ,  1092.1001/29 }"  
        
    Call CheckDB_FOLDERS(dbFOLDERS_ForRate,1)    
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''-"Պայմաններ և վիճակներ/Օբյեկտիվ ռիսկի դասիչ" գործողության կատարում-''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''      
    Log.Message "Check RC (Objective Risk/Օբյեկտիվ ռիսկի դասիչ) Function",,,DivideColor    

    ActionIsn(2) = Objective_Risk("01046643311","251120", "02", "ChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChan1",False)
                                                                                         
    Log.Message "SQL Check After RC (Objective Risk/Օբյեկտիվ ռիսկի դասիչ) Function",,,SqlDivideColor
    Log.Message "fISN = " & ActionIsn(2),,,SqlDivideColor
    
    'SQL Ստուգում DOCS աղուսյակում    
    fBODY = "  CODE:01046643311  DATE:20201125  RISK:02  COMMENT:ChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChan1  USERID:  77  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",ActionIsn(2),1)
    Call CheckDB_DOCS(ActionIsn(2),"COTSORC ","5",fBODY,1)
    
    'SQL Ստուգում DOCLOG աղուսյակում 
    Call CheckQueryRowCount("DOCLOG","fISN",ActionIsn(2),4)
    Call CheckDB_DOCLOG(ActionIsn(2),"77","N","1","",1)
    Call CheckDB_DOCLOG(ActionIsn(2),"77","M","99","àõÕ³ñÏí»É ¿ Ñ³ëï³ïÙ³Ý",1)
    Call CheckDB_DOCLOG(ActionIsn(2),"77","T","2","",1)
    Call CheckDB_DOCLOG(ActionIsn(2),"77","C","5","",1)
    
    'SQL Ստուգում HIF աղուսյակում 
    Call Check_HIF ("20201125", "N0", fISN, "0.00", "0.00", "ORC", "02")   
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''-"Ռիսկի դասիչ և պահուստավորման տոկոս" գործողության կատարում-''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''      
    Log.Message "Check RC (Risk Classifier/Ռիսկի դասիչ և պահուստավորման տոկոս) Function",,,DivideColor   
    
    ActionIsn(3) = Create_Risk_Classifier("01046643311","251120", "01", "2", "ChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChan1")

    Log.Message "SQL Check After RC (Risk Classifier/Ռիսկի դասիչ և պահուստավորման տոկոս) Function",,,SqlDivideColor
    Log.Message "fISN = " & ActionIsn(3),,,SqlDivideColor 
    
    'SQL Ստուգում DOCS աղուսյակում 
    fBODY = "  CODE:01046643311  DATE:20201125  RISK:01  PERRES:2  COMMENT:ChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChan1  USERID:  77  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",ActionIsn(3),1)
    Call CheckDB_DOCS(ActionIsn(3),"COTSRsPr","5",fBODY,1)
    
    'SQL Ստուգում DOCLOG աղուսյակում 
    Call CheckQueryRowCount("DOCLOG","fISN",ActionIsn(3),4)
    Call CheckDB_DOCLOG(ActionIsn(3),"77","N","1","",1)
    Call CheckDB_DOCLOG(ActionIsn(3),"77","M","99","àõÕ³ñÏí»É ¿ Ñ³ëï³ïÙ³Ý",1)
    Call CheckDB_DOCLOG(ActionIsn(3),"77","T","2","",1)
    Call CheckDB_DOCLOG(ActionIsn(3),"77","C","5","",1)
    
    'SQL Ստուգում HIF աղուսյակում
    Call Check_HIF("20201125", "N0", fISN, "2.00", "0.00", "PRS", Null)  
    Call Check_HIF("20201125", "N0", fISN, "0.00", "0.00", "RSK", "01")    

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''-"Պահուստավորում" գործողության կատարում-'''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''      
    Log.Message "Check RC (Store/Պահուստավորում) Function",,,DivideColor   

    Call Doc_Store(NewStore)
    
    Log.Message "SQL Check After RC (Store/Պահուստավորում) Function",,,SqlDivideColor
    Log.Message "fISN = " & NewStore.Isn,,,SqlDivideColor  
    
    'SQL Ստուգում DOCS աղուսյակում 
    fBODY = "  CODE:01046643311  DATE:20201125  SUMRES:100000000  COMMENT:NewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOut1  ACSBRANCH:01  ACSDEPART:3  ACSTYPE:CO1  USERID:  77  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",NewStore.Isn,1)
    Call CheckDB_DOCS(NewStore.Isn,"CODSRes ","5",fBODY,1)
    
    'SQL Ստուգում DOCLOG աղուսյակում 
    Call CheckQueryRowCount("DOCLOG","fISN",NewStore.Isn,4)
    Call CheckDB_DOCLOG(NewStore.Isn,"77","N","1","",1)
    Call CheckDB_DOCLOG(NewStore.Isn,"77","M","99","àõÕ³ñÏí»É ¿ Ñ³ëï³ïÙ³Ý",1)
    Call CheckDB_DOCLOG(NewStore.Isn,"77","T","2","",1)
    Call CheckDB_DOCLOG(NewStore.Isn,"77","C","5","",1)
    
    'SQL Ստուգում HI աղուսյակում 
    Call CheckQueryRowCount("HI","fBASE",NewStore.Isn,2)
    Query = "Select fOBJECT From HI Where fOP = 'RST' and fDBCR = 'C' and fSUM = '100000000.00'"
    fOBJECT(0) = my_Row_Count(Query)  
    Call Check_HI_CE_accounting ("20201125",NewStore.Isn, "01", fOBJECT(0), "100000000.00", "000", "100000000.00", "RST", "C") 
    
    'SQL Ստուգում HIREST  աղուսյակում համապատասխան fOBJECT-ով
    Call CheckDB_HIREST("01",fOBJECT(0),"-100000000.00","000","-100000000.00",1)
    
    'SQL Ստուգում HI աղուսյակում 
    Query = "Select fOBJECT From HI Where fOP = 'RST' and fDBCR = 'D' and fSUM = '100000000.00'"
    fOBJECT(1) = my_Row_Count(Query)  
    Call Check_HI_CE_accounting ("20201125",NewStore.Isn, "01", fOBJECT(1), "100000000.00", "000", "100000000.00", "RST", "D") 
    
    'SQL Ստուգում HIREST  աղուսյակում համապատասխան fOBJECT-ով
    Call CheckDB_HIREST("01",fOBJECT(1),"100000000.00","000","100000000.00",1)
    
    'SQL Ստուգում ACCOUNTS աղուսյակում համապատասխան fOBJECT-ով
    Call CheckQueryRowCount("ACCOUNTS","fISN",fOBJECT(0),1)
    Call CheckQueryRowCount("ACCOUNTS","fISN",fOBJECT(1),1)
    
    'SQL Ստուգում HIR աղուսյակում 
    Call CheckQueryRowCount("HIR","fOBJECT",fISN,5)
    Call Check_HIR("20201125", "R1", fISN, "000", "999970.70", "AGR", "D")
    Call Check_HIR("20201125", "R3", fISN, "000", "22222222.10", "PNA", "D")
    Call Check_HIR("20201125", "R4", fISN, "000", "100000000.00", "RES", "D")
    Call Check_HIR("20201125", "R7", fISN, "000", "11111111.10", "PNA", "D")
    Call Check_HIR("20201125", "RÄ", fISN, "000", "999970.70", "AGJ", "D")
    
    'SQL Ստուգում HIRREST  աղուսյակում 
    Call CheckQueryRowCount("HIRREST","fOBJECT",fISN,5)
    Call CheckDB_HIRREST("R1",fISN,"999970.70","20201125",1)  
    Call CheckDB_HIRREST("R3",fISN,"22222222.10","20201125",1)  
    Call CheckDB_HIRREST("R4",fISN,"100000000.00","20201125",1)  
    Call CheckDB_HIRREST("R7",fISN,"11111111.10","20201125",1)  
    Call CheckDB_HIRREST("RÄ",fISN,"999970.70","20201125",1)  
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''-"Տոկոսների հաշվարկ" գործողության կատարում-''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''    
    Log.Message "Check RC Create (Calculate Percents/Տոկոսների հաշվարկ) Function",,,DivideColor

    ExpectedMessage = "¶»ñ³Í³Ëë³ÛÇÝ å³ÛÙ³Ý³·Çñª  01046643311  /²·³Ã ö³ÛÉ³ï³ÏÛ³Ý/"& vbCrLf &"--------------------------------------------------------------------------------------------------------------"& vbCrLf &""& vbCrLf &"25/11/20-ÇÝ Ï³ï³ñí»É ¿ ïáÏáëÇ Ñ³ßí³ñÏáõÙ"
    Call CalculatePercents(CalcPercents_2,ExpectedMessage,True)
    
    Log.Message "SQL Check After RC (Calculate Percents/Տոկոսների հաշվարկ) Function",,,SqlDivideColor
    Log.Message "fISN = " & CalcPercents_2.Isn,,,SqlDivideColor

    'SQL Ստուգում DOCS աղուսյակում 
    fBODY = "  CODE:01046643311  DATECHARGE:20201125  DATE:20201125  SUMAGRPEN:987.1/123.1  SUMALLPEN:987.1/123.1  ACSBRANCH:00  ACSDEPART:4  ACSTYPE:CO1  USERID:  77  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",CalcPercents_2.Isn,1)
    Call CheckDB_DOCS(CalcPercents_2.Isn,"CODSChrg","5",fBODY,1)

    'SQL Ստուգում DOCLOG աղուսյակում 
    Call CheckQueryRowCount("DOCLOG","fISN",CalcPercents_2.Isn,4)
    Call CheckDB_DOCLOG(CalcPercents_2.Isn,"77","N","1","",1)
    Call CheckDB_DOCLOG(CalcPercents_2.Isn,"77","M","99","àõÕ³ñÏí»É ¿ Ñ³ëï³ïÙ³Ý",1)
    Call CheckDB_DOCLOG(CalcPercents_2.Isn,"77","T","2","",1)
    Call CheckDB_DOCLOG(CalcPercents_2.Isn,"77","C","5","",1)
    
    'SQL Ստուգում HIF աղուսյակում 
    Call CheckQueryRowCount("HIF","fBASE",CalcPercents_2.Isn,1)
    
    Query = "Select * from HIF where fDATE = '20201125' and fTYPE = 'N0' and fOBJECT = "&fISN&" and fOP = 'DTC' and fSUM = 0.00 and fCURSUM = 0.00 and fSPEC = 20201125	and fBASE = "&CalcPercents_2.Isn
    If my_Row_Count(Query) = 1 Then
        Log.Message "HIF record is Correct" 
    End If
    
    'SQL Ստուգում HIR աղուսյակում 
    Call CheckQueryRowCount("HIR","fBASE",CalcPercents_2.Isn,2)
    Call Check_HIR("20201125", "R3", fISN, "000", "987.10", "PNA", "D")
    Call Check_HIR("20201125", "R7", fISN, "000", "123.10", "PNA", "D")
    
    'SQL Ստուգում HIT աղուսյակում 
    Call CheckQueryRowCount("HIT","fBASE",CalcPercents_2.Isn,2)
    Call Check_HIT("20201125", "N3", fISN, "000", "987.10", "PNA", "D")
    Call Check_HIT("20201125", "N7", fISN, "000", "123.10", "PNA", "D")
    
    'SQL Ստուգում HIRREST  աղուսյակում 
    Call CheckQueryRowCount("HIRREST","fOBJECT",fISN,5)
    Call CheckDB_HIRREST("R1",fISN,"999970.70","20201125",1)
    Call CheckDB_HIRREST("R3",fISN,"22223209.20","20201125",1)
    Call CheckDB_HIRREST("R4",fISN,"100000000.00","20201125",1)
    Call CheckDB_HIRREST("R7",fISN,"11111234.20","20201125",1)
    Call CheckDB_HIRREST("RÄ",fISN,"999970.70","20201125",1)
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''-"Դուրս գրում" գործողության կատարում-''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''    
    Log.Message "Check RC (Write Out/Դուրս գրում) Function",,,DivideColor      
    
    Call Create_WriteOut(NewWriteOut)
    
    Log.Message "SQL Check After RC (Write Out/Դուրս գրում) Function",,,SqlDivideColor
    Log.Message "fISN = " & NewWriteOut.Isn,,,SqlDivideColor
    
    'SQL Ստուգում DOCS աղուսյակում 
    fBODY = "  CODE:01046643311  DATE:20201125  SUMAGR:100  SUMFINE:764  SUMMA:864  COMMENT:NewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOut1  ACSBRANCH:01  ACSDEPART:4  ACSTYPE:CO1  USERID:  77  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",NewWriteOut.Isn,1)
    Call CheckDB_DOCS(NewWriteOut.Isn,"CODSOut ","5",fBODY,1)

    'SQL Ստուգում DOCLOG աղուսյակում 
    Call CheckQueryRowCount("DOCLOG","fISN",NewWriteOut.Isn,4)
    Call CheckDB_DOCLOG(NewWriteOut.Isn,"77","N","1","",1)
    Call CheckDB_DOCLOG(NewWriteOut.Isn,"77","M","99","àõÕ³ñÏí»É ¿ Ñ³ëï³ïÙ³Ý",1)
    Call CheckDB_DOCLOG(NewWriteOut.Isn,"77","T","2","",1)
    Call CheckDB_DOCLOG(NewWriteOut.Isn,"77","C","5","",1)
    
    Query = "Select fOBJECT From HI Where fOP = 'MSC' and fDBCR = 'D' and fTYPE = '01' AND fBASE = "&NewWriteOut.Isn
    fOBJECT(2) = my_Row_Count(Query)      
    'SQL Ստուգում HI աղուսյակում  համապատասխան fOBJECT-ով     
    Call Check_HI_CE_accounting ("20201125",NewWriteOut.Isn, "01", fOBJECT(2), "100.00", "000", "100.00", "MSC", "D")
    'SQL Ստուգում HIREST աղուսյակում համապատասխան fOBJECT-ով
    Call CheckDB_HIREST("01",fOBJECT(2),"-99999900.00","000","-99999900.00",1)
    'SQL Ստուգում DOCP աղուսյակում համապատասխան fOBJECT-ով
    Call CheckDB_DOCP(fOBJECT(2),"Acc     ",AccountParentIsn,1)
    
    Query = "Select fOBJECT From HI Where fOP = 'MSC' and fDBCR = 'C' and fTYPE = '01' AND fBASE = "&NewWriteOut.Isn
    fOBJECT(3) = my_Row_Count(Query)      
    'SQL Ստուգում HI աղուսյակում  համապատասխան fOBJECT-ով         
    Call Check_HI_CE_accounting ("20201125",NewWriteOut.Isn, "01", fOBJECT(3), "100.00", "000", "100.00", "MSC", "C")
    'SQL Ստուգում HIREST աղուսյակում համապատասխան fOBJECT-ով
    Call CheckDB_HIREST("01",fOBJECT(3),"999870.70","000","999870.70",1)
    'SQL Ստուգում DOCP աղուսյակում համապատասխան fOBJECT-ով
    Call CheckDB_DOCP(fOBJECT(3),"Acc   ",AccountParentIsn,1)
    
    Query = "Select fOBJECT From HI Where fOP = 'MSC' and fDBCR = 'D' and fTYPE = '02' AND fBASE = "&NewWriteOut.Isn
    fOBJECT(4) = my_Row_Count(Query)      
    'SQL Ստուգում HI աղուսյակում  համապատասխան fOBJECT-ով         
    Call Check_HI_CE_accounting ("20201125",NewWriteOut.Isn, "02", fOBJECT(4), "100.00", "000", "100.00", "MSC", "D")
    'SQL Ստուգում HIREST աղուսյակում համապատասխան fOBJECT-ով
    Call CheckDB_HIREST("02",fOBJECT(4),"0.00","000","0.00",1)
    'SQL Ստուգում DOCP աղուսյակում համապատասխան fOBJECT-ով
    Call CheckDB_DOCP(fOBJECT(4),"NBAcc   ",AccountParentIsn,1)
    
    'SQL Ստուգում NBACCOUNTS  աղուսյակում համապատասխան fOBJECT-ով
    Call CheckQueryRowCount("NBACCOUNTS","fISN",fOBJECT(4),1)
    
    'SQL Ստուգում HIR աղուսյակում 
    Call CheckQueryRowCount("HIR","fBASE",NewWriteOut.Isn,3)
    Call Check_HIR("20201125", "R4", fISN, "000", "100.00", "OUT", "C")
    Call Check_HIR("20201125", "R5", fISN, "000", "100.00", "OUT", "D")
    Call Check_HIR("20201125", "R7", fISN, "000", "764.00", "OUT", "D")
    
    'SQL Ստուգում HIRREST  աղուսյակում 
    Call CheckQueryRowCount("HIRREST","fOBJECT",fISN,6)
    Call CheckDB_HIRREST("R1",fISN,"999970.70","20201125",1)
    Call CheckDB_HIRREST("R3",fISN,"22223209.20","20201125",1)
    Call CheckDB_HIRREST("R4",fISN,"99999900.00","20201125",1)
    Call CheckDB_HIRREST("R5",fISN,"100.00","20201125",1)
    Call CheckDB_HIRREST("R7",fISN,"11111998.20","20201125",1)
    Call CheckDB_HIRREST("RÄ",fISN,"999970.70","20201125",1)
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''-"Դուրս գրածի վերականգնում" գործողության կատարում-''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''    
    Log.Message "Check RC (Write Off Reconstruction/Դուրս գրածի վերականգնում) Function",,,DivideColor       
    
    Call WriteOut_Reconstruction(NewWriteOff,True)
    
    Log.Message "SQL Check After RC (Write Off Reconstruction/Դուրս գրածի վերականգնում) Function",,,SqlDivideColor
    Log.Message "fISN = " & NewWriteOff.Isn,,,SqlDivideColor
    
    'SQL Ստուգում DOCS աղուսյակում 
    fBODY = "  CODE:01046643311  DATE:20201125  SUMAGR:100  SUMFINE:11111998.2  SUMMA:11112098.2  COMMENT:wWriteOffwWriteOffwWriteOffwWriteOffwWriteOffwWriteOffwWriteOffwWriteOffwWriteOffwWriteOffwWriteOff1  ACSBRANCH:01  ACSDEPART:3  ACSTYPE:CO1  USERID:  77  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",NewWriteOff.Isn,1)
    Call CheckDB_DOCS(NewWriteOff.Isn,"CODSInc ","5",fBODY,1)

    'SQL Ստուգում DOCLOG աղուսյակում 
    Call CheckQueryRowCount("DOCLOG","fISN",NewWriteOff.Isn,4)
    Call CheckDB_DOCLOG(NewWriteOff.Isn,"77","N","1","",1)
    Call CheckDB_DOCLOG(NewWriteOff.Isn,"77","M","99","àõÕ³ñÏí»É ¿ Ñ³ëï³ïÙ³Ý",1)
    Call CheckDB_DOCLOG(NewWriteOff.Isn,"77","T","2","",1)
    Call CheckDB_DOCLOG(NewWriteOff.Isn,"77","C","5","",1)
    
    Query = "Select fOBJECT From HI Where fOP = 'MSC' and fDBCR = 'C' and fSUM = '100.00' and fBASEDEPART = '3' AND fBASE ="&NewWriteOff.Isn
    fOBJECT(5) = my_Row_Count(Query)      
    'SQL Ստուգում HI աղուսյակում  համապատասխան fOBJECT-ով         
    Call Check_HI_CE_accounting ("20201125",NewWriteOff.Isn, "01", fOBJECT(5), "100.00", "000", "100.00", "MSC", "C")
    'SQL Ստուգում HIREST աղուսյակում համապատասխան fOBJECT-ով
    Call CheckDB_HIREST("01",fOBJECT(5),"-100000000.00","000","-100000000.00",1)
    
    Query = "Select fOBJECT From HI Where fOP = 'MSC' and fDBCR = 'D' and fSUM = '100.00' and fBASEDEPART = '3' AND fBASE ="&NewWriteOff.Isn
    fOBJECT(6) = my_Row_Count(Query)      
    'SQL Ստուգում HI աղուսյակում  համապատասխան fOBJECT-ով         
    Call Check_HI_CE_accounting ("20201125",NewWriteOff.Isn, "01", fOBJECT(6), "100.00", "000", "100.00", "MSC", "D")
    'SQL Ստուգում HIREST աղուսյակում համապատասխան fOBJECT-ով
    Call CheckDB_HIREST("01",fOBJECT(6),"999970.70","000","999970.70",1)
    
    Query = "Select fOBJECT From HI Where fOP = 'MSC' and fDBCR = 'C' and fSUM = '100.00' and fBASEDEPART = '2' AND fBASE = "&NewWriteOff.Isn
    fOBJECT(7) = my_Row_Count(Query)      
    'SQL Ստուգում HI աղուսյակում  համապատասխան fOBJECT-ով         
    Call Check_HI_CE_accounting ("20201125",NewWriteOff.Isn, "02", fOBJECT(7), "100.00", "000", "100.00", "MSC", "C")
    
    'SQL Ստուգում HIR աղուսյակում 
    Call CheckQueryRowCount("HIR","fBASE",NewWriteOff.Isn,4)
    Call Check_HIR("20201125", "R4", fISN, "000", "100.00", "INC", "D")
    Call Check_HIR("20201125", "R5", fISN, "000", "100.00", "INC", "C")
    Call Check_HIR("20201125", "R7", fISN, "000", "11111998.20", "INC", "C")
    Call Check_HIR("20201125", "RI", fISN, "000", "11111234.20", "IR7", "D")
    
    'SQL Ստուգում HIRREST  աղուսյակում 
    Call CheckQueryRowCount("HIRREST","fOBJECT",fISN,7)
    Call CheckDB_HIRREST("R1",fISN,"999970.70","20201125",1)
    Call CheckDB_HIRREST("R3",fISN,"22223209.20","20201125",1)
    Call CheckDB_HIRREST("R4",fISN,"100000000.00","20201125",1)
    Call CheckDB_HIRREST("R5",fISN,"0.00","20201125",1)
    Call CheckDB_HIRREST("R7",fISN,"0.00","20201125",1)
    Call CheckDB_HIRREST("RI",fISN,"11111234.20","20201125",1)
    Call CheckDB_HIRREST("RÄ",fISN,"999970.70","20201125",1)
    
    wMDIClient.VBObject("frmPttel").Close
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''-Հեռացնել բոլոր գործողությունները-'''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''      
    Log.Message "Deleted All Action",,,DivideColor 
 
    Call wTreeView.DblClickItem("|¶»ñ³Í³Ëë|¶áñÍáÕáõÃÛáõÝÝ»ñ")
    wMDIClient.Refresh
    
    BuiltIn.Delay(delay_small)
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    BuiltIn.Delay(delay_small)
    Set DocForm = wMDIClient.VBObject("frmPttel") 
        
    If WaitForPttel("frmPttel") Then
        Call SearchAndDelete("frmPttel", 4, "11111998.2", "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ") 
        Call SearchAndDelete("frmPttel", 5, "¶»ñ³Í³ËëÇ ¹áõñë·ñáõÙ", "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ")    
        Call SearchAndDelete("frmPttel", 4, "987.1", "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ") 
        Call SearchAndDelete("frmPttel", 4, "100000000", "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ") 
        wMDIClient.VBObject("frmPttel").Close
    Else
        Log.Error "Can Not Open գործողությունները Window",,,ErrorColor         
    End If 
    If DocForm.Exists Then
        Log.Error "Can Not Close գործողությունները Window",,,ErrorColor
    End If 
     
    Call wTreeView.DblClickItem("|¶»ñ³Í³Ëë|Üáñ ÷³ëï³Ã., ÃÕÃ³å³Ý³ÏÝ»ñ, Ñ³ßí»ïíáõÃÛáõÝÝ»ñ|¶áñÍáÕáõÃÛáõÝÝ»ñ, ÷á÷áËáõÃÛáõÝÝ»ñ|ä³ÛÙ³Ý³·ñÇ µáÉáñ ·áñÍáÕáõÃÛáõÝÝ»ñ")
    wMDIClient.Refresh
    
    BuiltIn.Delay(delay_small)
    Call Rekvizit_Fill("Dialog", 1, "General", "START", "251120" & "[Tab]" & "251120")
    Call Rekvizit_Fill("Dialog", 1, "General", "NUM", "01046643311")
    
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    BuiltIn.Delay(delay_small)
    
    If WaitForPttel("frmPttel") Then
        Call SearchAndDelete("frmPttel", 3, "èÇëÏÇ ¹³ëÇã", "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ") 
        Call SearchAndDelete("frmPttel", 3, "úµÛ»ÏïÇí éÇëÏÇ ¹³ëÇã", "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ") 
        Call SearchAndDelete("frmPttel", 4, "îáÏáë³¹ñáõÛùÝ»ñ", "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ") 
        Call SearchAndDelete("frmPttel", 3, "îáÏáëÝ»ñÇ Ñ³ßí³ñÏÙ³Ý ³Ùë³ÃÇí", "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ") 
        Call SearchAndDelete("frmPttel", 3, "Ä³ÙÏ»ï³Ýó ·»ñ³Í³ËëÇ ïáõÛÅ", "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ") 
        wMDIClient.VBObject("frmPttel").Close
    Else
        Log.Error "Can Not Open Պայմանագրի բոլոր գործողությունները Window",,,ErrorColor         
    End If 
    If DocForm.Exists Then
        Log.Error "Can Not Close Պայմանագրի բոլոր գործողությունները Window",,,ErrorColor
    End If 
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''-SQL Check After Deleted All Action-'''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 
    Log.Message "SQL Check After Deleted All Action",,,SqlDivideColor
    
    'SQL Ստուգում DOCS աղուսյակում  
    fBODY = "  CODE:01046643311  CRDTCODE:777000000227L001  CLICOD:00000022  NAME:²·³Ã ö³ÛÉ³ï³ÏÛ³Ý  CURRENCY:000  ACCACC:01046643311  AUTODEBT:1  ACCCONNMODE:3  USECLICONNSCH:0  DATE:20201125  DATEGIVE:20201125  ACSBRANCH:00  ACSDEPART:2  ACSTYPE:CO1  KINDSCALE:1  PCPENAGR:0/1  CONSTPER:0  SECTOR:F  SCHEDULE:9  GUARANTEE:9  PERRES:1  PPRCODE:01046643311  SUBJRISK:0  CHRGFIRSTDAY:1  GIVEN:1  PUTINLR:0  NOTCLASS:0  OTHERCOLLATERAL:0  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckDB_DOCS(fISN,"COSimpl","999",fBODY,1)
    
    fBODY = "  CODE:01046643311  DATECHARGE:20201125  DATE:20201125  SUMAGRPEN:22222222.1/11111111.1  SUMALLPEN:22222222.1/11111111.1  COMMENT:NewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOut1  ACSBRANCH:01  ACSDEPART:4  ACSTYPE:CO1  USERID:  77  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",CalcPercents.Isn,1)
    Call CheckDB_DOCS(CalcPercents.Isn,"CODSChrg","999",fBODY,1)
    
    fBODY = "  CODE:01046643311  DATE:20201125  PCPENAGR:1092.1001/29  COMMENT:ChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChan1  USERID:  77  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",ActionIsn(1),1)
    Call CheckDB_DOCS(ActionIsn(1),"COTSPC  ","999",fBODY,1)
    
    'SQL Ստուգում DOCS աղուսյակում    
    fBODY = "  CODE:01046643311  DATE:20201125  RISK:02  COMMENT:ChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChan1  USERID:  77  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",ActionIsn(2),1)
    Call CheckDB_DOCS(ActionIsn(2),"COTSORC ","999",fBODY,1)
    
    fBODY = "  CODE:01046643311  DATE:20201125  RISK:01  PERRES:2  COMMENT:ChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChan1  USERID:  77  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",ActionIsn(3),1)
    Call CheckDB_DOCS(ActionIsn(3),"COTSRsPr","999",fBODY,1)
    
    fBODY = "  CODE:01046643311  DATE:20201125  RISK:01  PERRES:2  COMMENT:ChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChan1  USERID:  77  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",ActionIsn(3),1)
    Call CheckDB_DOCS(ActionIsn(3),"COTSRsPr","999",fBODY,1)
    
    fBODY = "  CODE:01046643311  DATE:20201125  SUMRES:100000000  COMMENT:NewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOut1  ACSBRANCH:01  ACSDEPART:3  ACSTYPE:CO1  USERID:  77  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",NewStore.Isn,1)
    Call CheckDB_DOCS(NewStore.Isn,"CODSRes ","999",fBODY,1)
    
    fBODY = "  CODE:01046643311  DATECHARGE:20201125  DATE:20201125  SUMAGRPEN:987.1/123.1  SUMALLPEN:987.1/123.1  ACSBRANCH:00  ACSDEPART:4  ACSTYPE:CO1  USERID:  77  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",CalcPercents_2.Isn,1)
    Call CheckDB_DOCS(CalcPercents_2.Isn,"CODSChrg","999",fBODY,1)
    
    fBODY = "  CODE:01046643311  DATECHARGE:20201125  DATE:20201125  SUMAGRPEN:987.1/123.1  SUMALLPEN:987.1/123.1  ACSBRANCH:00  ACSDEPART:4  ACSTYPE:CO1  USERID:  77  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",CalcPercents_2.Isn,1)
    Call CheckDB_DOCS(CalcPercents_2.Isn,"CODSChrg","999",fBODY,1)
    
    fBODY = "  CODE:01046643311  DATE:20201125  SUMAGR:100  SUMFINE:764  SUMMA:864  COMMENT:NewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOut1  ACSBRANCH:01  ACSDEPART:4  ACSTYPE:CO1  USERID:  77  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",NewWriteOut.Isn,1)
    Call CheckDB_DOCS(NewWriteOut.Isn,"CODSOut ","999",fBODY,1)
    
    fBODY = "  CODE:01046643311  DATE:20201125  SUMAGR:100  SUMFINE:11111998.2  SUMMA:11112098.2  COMMENT:wWriteOffwWriteOffwWriteOffwWriteOffwWriteOffwWriteOffwWriteOffwWriteOffwWriteOffwWriteOffwWriteOff1  ACSBRANCH:01  ACSDEPART:3  ACSTYPE:CO1  USERID:  77  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",NewWriteOff.Isn,1)
    Call CheckDB_DOCS(NewWriteOff.Isn,"CODSInc ","999",fBODY,1)
    
    'SQL Ստուգում DOCLOG աղուսյակում 
    Call CheckQueryRowCount("DOCLOG","fISN",fISN,6)
    Call CheckDB_DOCLOG(fISN,"77","M","7","1 Ï³ñ·Ç ³ñ·»Éí³Í ·áñÍáÕáõÃÛáõÝ",1)
    Call CheckDB_DOCLOG(fISN,"77","D","999","",1)
    
    Set dbFOLDERS(9) = New_DB_FOLDERS()
        dbFOLDERS(9).fFOLDERID = ".R." & aqConvert.DateTimeToFormatStr(aqDateTime.Now(),"%Y%m%d")
        dbFOLDERS(9).fNAME = "COSimpl "
        dbFOLDERS(9).fKEY = fISN
        dbFOLDERS(9).fISN = fISN
        dbFOLDERS(9).fSTATUS = "0"
        dbFOLDERS(9).fSPEC = Left_Align(Get_Compname_DOCLOG(fISN), 16)&"CredO   ARMSOFT                       007  "
        dbFOLDERS(9).fDCBRANCH = "00"
        dbFOLDERS(9).fDCDEPART = "2"
    Call CheckDB_FOLDERS(dbFOLDERS(9),1)
    
        dbFOLDERS(9).fNAME = "CODSChrg"
        dbFOLDERS(9).fKEY = CalcPercents.Isn
        dbFOLDERS(9).fISN = CalcPercents.Isn
        dbFOLDERS(9).fSPEC = Left_Align(Get_Compname_DOCLOG(fISN), 16)&"CredO   ARMSOFT                       005  "
        dbFOLDERS(9).fDCBRANCH = "01"
        dbFOLDERS(9).fDCDEPART = "4"
    Call CheckDB_FOLDERS(dbFOLDERS(9),1)
    
        dbFOLDERS(9).fNAME = "COTSPC  "
        dbFOLDERS(9).fKEY = ActionIsn(1)
        dbFOLDERS(9).fISN = ActionIsn(1)
        dbFOLDERS(9).fSPEC = Left_Align(Get_Compname_DOCLOG(fISN), 16)&"CredO   ARMSOFT                       005  "
        dbFOLDERS(9).fDCBRANCH = ""
        dbFOLDERS(9).fDCDEPART = ""
    Call CheckDB_FOLDERS(dbFOLDERS(9),1)
    
        dbFOLDERS(9).fNAME = "COTSORC "
        dbFOLDERS(9).fKEY = ActionIsn(2)
        dbFOLDERS(9).fISN = ActionIsn(2)
        dbFOLDERS(9).fSPEC = Left_Align(Get_Compname_DOCLOG(fISN), 16)&"CredO   ARMSOFT                       005  "
    Call CheckDB_FOLDERS(dbFOLDERS(9),1)
    
        dbFOLDERS(9).fNAME = "COTSRsPr"
        dbFOLDERS(9).fKEY = ActionIsn(3)
        dbFOLDERS(9).fISN = ActionIsn(3)
        dbFOLDERS(9).fSPEC = Left_Align(Get_Compname_DOCLOG(fISN), 16)&"CredO   ARMSOFT                       005  "
    Call CheckDB_FOLDERS(dbFOLDERS(9),1)
    
        dbFOLDERS(9).fNAME = "CODSRes "
        dbFOLDERS(9).fKEY = NewStore.Isn
        dbFOLDERS(9).fISN = NewStore.Isn
        dbFOLDERS(9).fSPEC = Left_Align(Get_Compname_DOCLOG(fISN), 16)&"CredO   ARMSOFT                       115  "
        dbFOLDERS(9).fDCBRANCH = "01"
        dbFOLDERS(9).fDCDEPART = "3"
    Call CheckDB_FOLDERS(dbFOLDERS(9),1)
    
        dbFOLDERS(9).fNAME = "CODSChrg"
        dbFOLDERS(9).fKEY = CalcPercents_2.Isn
        dbFOLDERS(9).fISN = CalcPercents_2.Isn
        dbFOLDERS(9).fSPEC = Left_Align(Get_Compname_DOCLOG(fISN), 16)&"CredO   ARMSOFT                       005  "
        dbFOLDERS(9).fDCBRANCH = "00"
        dbFOLDERS(9).fDCDEPART = "4"
    Call CheckDB_FOLDERS(dbFOLDERS(9),1)
    
        dbFOLDERS(9).fNAME = "CODSOut "
        dbFOLDERS(9).fKEY = NewWriteOut.Isn
        dbFOLDERS(9).fISN = NewWriteOut.Isn
        dbFOLDERS(9).fSPEC = Left_Align(Get_Compname_DOCLOG(fISN), 16)&"CredO   ARMSOFT                       115  "
        dbFOLDERS(9).fDCBRANCH = "01"
        dbFOLDERS(9).fDCDEPART = "4"
    Call CheckDB_FOLDERS(dbFOLDERS(9),1)
        
        dbFOLDERS(9).fNAME = "CODSInc "
        dbFOLDERS(9).fKEY = NewWriteOff.Isn
        dbFOLDERS(9).fISN = NewWriteOff.Isn
        dbFOLDERS(9).fSPEC = Left_Align(Get_Compname_DOCLOG(fISN), 16)&"CredO   ARMSOFT                       115  "
        dbFOLDERS(9).fDCBRANCH = "01"
        dbFOLDERS(9).fDCDEPART = "3"
    Call CheckDB_FOLDERS(dbFOLDERS(9),1)
    
    'SQL Ստուգում HI աղուսյակում 
    Call CheckQueryRowCount("HI","fOBJECT",fISN,0)
    
    'SQL Ստուգում HIR աղուսյակում 
    Call CheckQueryRowCount("HIR","fOBJECT",fISN,0)
    
    'SQL Ստուգում HIF աղուսյակում 
    Call CheckQueryRowCount("HIF","fOBJECT",fISN,0)

    'SQL Ստուգում HIRREST  աղուսյակում 
    Call CheckQueryRowCount("HIRREST","fOBJECT",fISN,0)
    
    'SQL Ստուգում CONTRACTS աղուսյակում 
    Call CheckQueryRowCount("CONTRACTS","fDGISN",fISN,0)
    
    'SQL Ստուգում CAGRACCS աղուսյակում 
    Call CheckQueryRowCount("CAGRACCS","fAGRISN",fISN,0)
    
    'SQL Ստուգում DOCP աղուսյակում  
    Call CheckQueryRowCount("DOCP","fPARENTISN",fISN,0)
    Call CheckQueryRowCount("DOCP","fPARENTISN",AccountParentIsn,0)
    
    'SQL Ստուգում HIREST աղուսյակում  
    Call CheckQueryRowCount("HIREST","fOBJECT",fOBJECT(0),0)
    Call CheckQueryRowCount("HIREST","fOBJECT",fOBJECT(1),0)
    Call CheckQueryRowCount("HIREST","fOBJECT",fOBJECT(2),0)
    Call CheckQueryRowCount("HIREST","fOBJECT",fOBJECT(3),0)
    Call CheckQueryRowCount("HIREST","fOBJECT",fOBJECT(4),0)
    Call CheckQueryRowCount("HIREST","fOBJECT",fOBJECT(5),0)
    Call CheckQueryRowCount("HIREST","fOBJECT",fOBJECT(6),0)
    Call CheckQueryRowCount("HIREST","fOBJECT",fOBJECT(7),0)
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''-Գերածախս ունեցող հաշիվներ թղթապանակում փաստատթղթի առկայության ստուգում-''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 
    Call ExistsAccWithOverlimit_Filter_Fill(AccWithOverlimit, 1)   
    wMDIClient.VBObject("frmPttel").Close
    
    Call Close_AsBank()
End Sub    

Sub Test_InitializeFor_OverlimitFromAccount()
        '"Գերածախս ունեցող հաշիվներ" Ֆիլտրի լրացման արժեքներ
    Set AccWithOverlimit = New_AccountsWithOverlimit()
        AccWithOverlimit.Curr = "000"
        AccWithOverlimit.Client = "00000022"
        AccWithOverlimit.AccountMask = "01046643311"
        AccWithOverlimit.Division = "00"
        AccWithOverlimit.Department = "2"
        AccWithOverlimit.AccessType = ""
            
        '"Պայմանագրեր" Ֆիլտրի լրացման արժեքներ
    Set ContractFillter = New_ContractOverlimit()
        ContractFillter.AgreementLevel = "1"
        ContractFillter.AgreementSpecies = "5"
        ContractFillter.AgreementN = "01046643311"
        ContractFillter.Curr = "000"
        ContractFillter.Client = "00000022"
        ContractFillter.ShowClosed = "1"
        ContractFillter.Division = "00"
        ContractFillter.Department = "2"
        ContractFillter.AccessType = "CO1"
        
        '"Աջ կլիկ/Գերածախս" Պատուհանի լրացման արժեքներ
    Set RcOptionOverlimit = New_RcOverlimit()
        RcOptionOverlimit.ExpectedAgreementN = "01046643311"
        RcOptionOverlimit.Date = "251120"
        RcOptionOverlimit.Sum = "50121500.01"
        RcOptionOverlimit.CashOrNo = "2"
        RcOptionOverlimit.CalcAcc = "01046643311"
        RcOptionOverlimit.Comment = "Test_For_Overlimit1"
        RcOptionOverlimit.Division = "01"
        RcOptionOverlimit.Department = "4"
        
        '"Աջ կլիկ/Տոկոսների հաշվարկ" Պատուհանի լրացման արժեքներ
    Set CalcPercents = New_RcCalculatePercents()
        CalcPercents.ExpectedAgreementN = "01046643311"
        CalcPercents.CalculationDate = "251120"
        CalcPercents.OperationDate = "251120" 
        CalcPercents.FineOnPastDueSum = "22222222.09"
        CalcPercents.FineOnPastDueSum2 = "11111111.09"
        CalcPercents.TotalPenalty = "22222222.10"
        CalcPercents.TotalPenalty2 = "11111111.10"
        CalcPercents.Comment =  "NewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOut1"
        CalcPercents.Division =  "01"
        CalcPercents.Department =  "4"
    
    Set CalcPercents_2 = New_RcCalculatePercents()
        CalcPercents_2.ExpectedAgreementN = "01046643311"
        CalcPercents_2.CalculationDate = "251120"
        CalcPercents_2.OperationDate = "251120" 
        CalcPercents_2.FineOnPastDueSum = "987.09"
        CalcPercents_2.FineOnPastDueSum2 = "123.09"
        CalcPercents_2.TotalPenalty = "987.10"
        CalcPercents_2.TotalPenalty2 = "123.10"
        CalcPercents_2.Division =  "00"
        CalcPercents_2.Department =  "4" 
        
        '"Աջ կլիկ/Պահուստավորում" Պատուհանի լրացման արժեքներ
    Set NewStore = New_RcStore()    
        NewStore.ExpectedAgreementN = "01046643311"
        NewStore.Date = "251120"
        NewStore.Provision = "99999999.99"
        NewStore.UnProvision = ""
        NewStore.Comment =  "NewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOut1"
        NewStore.Division = "01"
        NewStore.Department = "3"
        
        '"Աջ կլիկ/Դուրս գրում" Պատուհանի լրացման արժեքներ
    Set NewWriteOut = New_RcWriteOut()
        NewWriteOut.ExpectedAgreementN = "01046643311"
        NewWriteOut.Date = "251120"
        NewWriteOut.ExpectedBaseSum = "0.00"
        NewWriteOut.BaseSum = "100.00"
        NewWriteOut.ExpectedFineOnPastSum = "0.00"
        NewWriteOut.FineOnPastSum = "764.00"
        NewWriteOut.TotalSum = "864.00"
        NewWriteOut.Comment = "NewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOut1"
        NewWriteOut.Division = "01"
        NewWriteOut.Department = "4"

        '"Աջ կլիկ/Դուրս գրածի վերականգնում" Պատուհանի լրացման արժեքներ
    Set NewWriteOff = New_RcWriteOut()
        NewWriteOff.ExpectedAgreementN = "01046643311"
        NewWriteOff.Date = "251120"
        NewWriteOff.ExpectedBaseSum = "100.00"
        NewWriteOff.BaseSum = "100.00"
        NewWriteOff.ExpectedFineOnPastSum = "11,111,998.20"
        NewWriteOff.FineOnPastSum = "11,111,998.20"
        NewWriteOff.TotalSum = "11,112,098.20"
        NewWriteOff.Comment = "wWriteOffwWriteOffwWriteOffwWriteOffwWriteOffwWriteOffwWriteOffwWriteOffwWriteOffwWriteOffwWriteOff1"
        NewWriteOff.Division = "01"
        NewWriteOff.Department = "3"
End Sub

Sub SQL_Initialize_OverlimitFromAccount(fISN,fISN2)
       
    Set dbCONTRACT = New_DB_CONTRACTS()
        dbCONTRACT.fDGISN = fISN
        dbCONTRACT.fDGPARENTISN = fISN
        dbCONTRACT.fDGISN1 = fISN
        dbCONTRACT.fDGISN3 = fISN
        dbCONTRACT.fDGAGRKIND = "5"
        dbCONTRACT.fDGSTATE = "7"
        dbCONTRACT.fDGTYPENAME = "COSimpl "
        dbCONTRACT.fDGCODE = "01046643311   "
        dbCONTRACT.fDGPPRCODE = "01046643311"
        dbCONTRACT.fDGCAPTION = "²·³Ã ö³ÛÉ³ï³ÏÛ³Ý"
        dbCONTRACT.fDGCLICODE = "00000022"
        dbCONTRACT.fDGCUR = "000"
        dbCONTRACT.fDGSUMMA = "0.00"
        dbCONTRACT.fDGALLSUMMA = "0.00"
        dbCONTRACT.fDGRISKDEGREE = "0.00"
        dbCONTRACT.fDGRISKDEGNB = "0.00"
        dbCONTRACT.fDGSCHEDULE = "9"
        dbCONTRACT.fDGDISTRICT = "   "
        dbCONTRACT.fDGACSBRANCH = "00"
        dbCONTRACT.fDGACSDEPART = "2"
        dbCONTRACT.fDGACSTYPE = "CO1"
        dbCONTRACT.fDGCRDTCODE = "777000000227L001"
        
    Set dbFOLDERS(1) = New_DB_FOLDERS()
        dbFOLDERS(1).fFOLDERID = "AGROVERLIM"
        dbFOLDERS(1).fNAME = "COSimpl "
        dbFOLDERS(1).fKEY = "01046643311"
        dbFOLDERS(1).fISN = fISN
        dbFOLDERS(1).fSTATUS = "1"
        dbFOLDERS(1).fCOM = "²·³Ã ö³ÛÉ³ï³ÏÛ³Ý"
        dbFOLDERS(1).fSPEC = "01046643311   CO1 0000"
        dbFOLDERS(1).fDCBRANCH = "00"
        dbFOLDERS(1).fDCDEPART = "2"

    Set dbFOLDERS(2) = New_DB_FOLDERS()
        dbFOLDERS(2).fFOLDERID = "Agr."&fISN
        dbFOLDERS(2).fNAME = "COSimpl "
        dbFOLDERS(2).fKEY = fISN
        dbFOLDERS(2).fISN = fISN
        dbFOLDERS(2).fSTATUS = "1"
        dbFOLDERS(2).fCOM = "¶»ñ³Í³Ëë ³ÝÅ³ÙÏ»ï"
        dbFOLDERS(2).fSPEC = "1¶»ñ³Í³Ëë ³ÝÅ³ÙÏ»ï- 01046643311 {²·³Ã ö³ÛÉ³ï³ÏÛ³Ý}"
     
    Set dbFOLDERS(3) = New_DB_FOLDERS()
        dbFOLDERS(3).fFOLDERID = "C.1628336"
        dbFOLDERS(3).fNAME = "COSimpl "
        dbFOLDERS(3).fKEY = fISN
        dbFOLDERS(3).fISN = fISN
        dbFOLDERS(3).fSTATUS = "1"
        dbFOLDERS(3).fCOM = " ¶»ñ³Í³Ëë ³ÝÅ³ÙÏ»ï"
        dbFOLDERS(3).fECOM = "1"
        dbFOLDERS(3).fSPEC = "01046643311 (²·³Ã ö³ÛÉ³ï³ÏÛ³Ý),     0 - Ð³ÛÏ³Ï³Ý ¹ñ³Ù"   
                
    Set dbFOLDERS(4) = New_DB_FOLDERS()
        dbFOLDERS(4).fFOLDERID = "ALLACCSACC"
        dbFOLDERS(4).fNAME = "COAgrAcc"
        dbFOLDERS(4).fKEY = fISN
        dbFOLDERS(4).fISN = fISN2
        dbFOLDERS(4).fSTATUS = "1"
        dbFOLDERS(4).fCOM = "¶»ñ³Í³ËëÇ Ñ³ßí³å³Ñ³Ï³Ý Ñ³í»Éí³Í"
        dbFOLDERS(4).fSPEC = "001046643311                                            1"   
    
    Set dbFOLDERS(5) = New_DB_FOLDERS()
        dbFOLDERS(5).fFOLDERID = "ALLACCSGEN"
        dbFOLDERS(5).fNAME = "COAgrAcc"
        dbFOLDERS(5).fKEY = fISN
        dbFOLDERS(5).fISN = fISN2
        dbFOLDERS(5).fSTATUS = "1"
        dbFOLDERS(5).fCOM = "¶»ñ³Í³ËëÇ Ñ³ßí³å³Ñ³Ï³Ý Ñ³í»Éí³Í"
        dbFOLDERS(5).fSPEC = "01080793012"  
    
    Set dbFOLDERS(6) = New_DB_FOLDERS()
        dbFOLDERS(6).fFOLDERID = "ALLACCSRES"
        dbFOLDERS(6).fNAME = "COAgrAcc"
        dbFOLDERS(6).fKEY = fISN
        dbFOLDERS(6).fISN = fISN2
        dbFOLDERS(6).fSTATUS = "1"
        dbFOLDERS(6).fCOM = "¶»ñ³Í³ËëÇ Ñ³ßí³å³Ñ³Ï³Ý Ñ³í»Éí³Í"
        dbFOLDERS(6).fSPEC = "000004532017303038100072112153000"  
        
    Set dbFOLDERS(7) = New_DB_FOLDERS()
        dbFOLDERS(7).fFOLDERID = "Agr."&fISN
        dbFOLDERS(7).fNAME = "COAgrAcc"
        dbFOLDERS(7).fKEY = fISN2
        dbFOLDERS(7).fISN = fISN2
        dbFOLDERS(7).fSTATUS = "1"
        dbFOLDERS(7).fCOM = "¶»ñ³Í³ËëÇ Ñ³ßí³å³Ñ³Ï³Ý Ñ³í»Éí³Í"
        dbFOLDERS(7).fSPEC = "1¶»ñ³Í³ËëÇ Ñ³ßí³å³Ñ³Ï³Ý Ñ³í»Éí³Í- 01046643311   "
        
    Set dbFOLDERS(8) = New_DB_FOLDERS()
        dbFOLDERS(8).fFOLDERID = "CAGRACCS"
        dbFOLDERS(8).fNAME = "COAgrAcc"
        dbFOLDERS(8).fKEY = "01046643311   "
        dbFOLDERS(8).fISN = fISN2
        dbFOLDERS(8).fSTATUS = "1"
        dbFOLDERS(8).fCOM = "¶»ñ³Í³ËëÇ Ñ³ßí³å³Ñ³Ï³Ý Ñ³í»Éí³Í"
End Sub
