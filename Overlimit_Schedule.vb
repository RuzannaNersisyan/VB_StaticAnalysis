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
'Test Case Id - 145220

Sub Check_Overlimit_Schedule()
    
    Dim sDATE,fDATE
    Call Test_Overlimit_Schedule_Init()
     
    'Համակարգ մուտք գործել ARMSOFT օգտագործողով
    sDATE = "20140101"
    fDATE = "20200101"
    Call Initialize_AsBank("bank", sDATE, fDATE)
    Login("ARMSOFT")

    'Մուտք գործել "Գերածախս"
    Call ChangeWorkspace(c_Overlimit) 

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''-Üáñ å³ÛÙ³Ý³·ñÇ ëï»ÕÍáõÙ-'''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "Create Overlimit Contract",,,DivideColor
    
    Call wTreeView.DblClickItem("|¶»ñ³Í³Ëë|Üáñ å³ÛÙ³Ý³·ñÇ ëï»ÕÍáõÙ|")
    
    'Ստեղծել Գերածախսի(Overlimit) պայմանագրիր
    Call CreateOverlimitDoc(NewOverlimitDoc)

    Log.Message "SQL Check After Create Overlimit Doc",,,SqlDivideColor
    Log.Message "fISN = " & NewOverlimitDoc.isn,,,SqlDivideColor

    fISN = NewOverlimitDoc.isn
    fBODY = "  CODE:00000113032  CRDTCODE:"&NewOverlimitDoc.GeneralTab.ExpectedCreditCode&"  CLICOD:00034852  NAME:ý²àôêî  CURRENCY:000  ACCACC:00000113032  COMMENT:MaxSinvolsSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSS1  AUTODEBT:1  ACCCONNMODE:3  USECLICONNSCH:1  ACCCONNSCH:002  DATE:20200101  DATEGIVE:20200101  ACSBRANCH:01  ACSDEPART:4  ACSTYPE:CO1  DATESFILLTYPE:2  AGRMARBEG:20200101  AGRMARFIN:20790101  AGRPERIOD:10/29  PASSOVDIRECTION:2  PASSOVTYPE:13  FIXEDROWSINSCH:0  SUMSDATESFILLTYPE:2  SUMSFILLTYPE:01  MIXEDSUMSINSCH:0  KINDSCALE:1  PCPENAGR:11.1111/29  CONSTPER:0  SECTOR:A  USAGEFIELD:03.006  AIM:04  INTERORG:BSTDB  SCHEDULE:9  GUARANTEE:1  COUNTRY:CH  LRDISTR:004  REGION:050000009  PERRES:1  NOTE:01  NOTE2:022  NOTE3:03  PPRCODE:88888888888888888881  SUBJRISK:1  CHRGFIRSTDAY:1  GIVEN:0  NTFMODE:1  SENDSTMADRS:2  PUTINLR:0  LRMRTCUR:000  ACRANOTE:02  NOTCLASS:1  REVISIONREASON:1  REPSOURCE:4  MORTSUBJECT:16  OTHERCOLLATERAL:1  "
    fBODY = Replace(fBODY, "  ", "%")
    Call SQL_Init_Overlimit_Schedule(fISN,"")
    dbCONTRACT.fDGCRDTCODE = NewOverlimitDoc.GeneralTab.ExpectedCreditCode
    
    'SQL Ստուգում DOCS աղուսյակում  
    Call CheckQueryRowCount("DOCS","fISN",fISN,1)
    Call CheckDB_DOCS(fISN,"COUniver","1",fBODY,1)

    'SQL Ստուգում DOCP աղուսյակում  
    Call CheckQueryRowCount("DOCP","fPARENTISN",fISN,1)
    Call CheckDB_DOCP("443871031","Acc     ",fISN,1)
    
    'SQL Ստուգում DOCLOG աղուսյակում 
    Call CheckQueryRowCount("DOCLOG","fISN",fISN,1)
    Call CheckDB_DOCLOG(fISN,"77","N","1","",1)

    'SQL Ստուգում DOCSG աղուսյակում 
    Call CheckQueryRowCount("DOCSG","fISN",fISN,4)
    Call CheckDB_DOCSG(fISN,"NOTES","0","CODE","00001",1)
    Call CheckDB_DOCSG(fISN,"NOTES","0","FILL","0",1)
    Call CheckDB_DOCSG(fISN,"NOTES","0","NAME","Ø»Ãá¹áÉá·Ç³                                       ",1)
    Call CheckDB_DOCSG(fISN,"NOTES","0","TYPE","C(1)",1)
    
    'SQL Ստուգում RESNUMBERS աղուսյակում 
    Call CheckQueryRowCount("RESNUMBERS","fISN",fISN,1)
    Call CheckDB_RESNUMBERS(fISN,"C","00000113032   ",1)
    
    'SQL Ստուգում AGRNOTES աղուսյակում 
    fVALUES = "Type:C(1) Value:  "
    fVALUES = Replace(fVALUES, "  ", vbCrLf)
    Call CheckQueryRowCount("AGRNOTES","fAGRISN",fISN,1)
    Call CheckDB_AGRNOTES(fISN,fVALUES,1)
    
    'SQL Ստուգում CONTRACTS աղուսյակում
    Call CheckQueryRowCount("CONTRACTS","fDGISN",fISN,1)
    Call CheckDB_CONTRACTS(dbCONTRACT,1)
    
    'SQL Ստուգում FOLDERS աղուսյակում 
    Call CheckQueryRowCount("FOLDERS","fISN",fISN,6)
    Call CheckDB_FOLDERS(dbFOLDERS(1),1)
    Call CheckDB_FOLDERS(dbFOLDERS(2),1)
    Call CheckDB_FOLDERS(dbFOLDERS(3),1)
    Call CheckDB_FOLDERS(dbFOLDERS(4),1)
    Call CheckDB_FOLDERS(dbFOLDERS(5),1)
    Call CheckDB_FOLDERS(dbFOLDERS(6),1)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''-Պայմանագիրը ուղարկում է հաստատման-''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''    
    Log.Message "Send To Verify Contract ",,,DivideColor 
    
    Call SendToVerify_Contrct(2, 5, "²Ûá")
    
      dbFOLDERS(1).fSTATUS = "0"
      dbFOLDERS(2).fSTATUS = "0"
      dbFOLDERS(3).fSTATUS = "0"
      dbFOLDERS(3).fCOM = " ¶»ñ³Í³Ëë (·ñ³ýÇÏáí å³ÛÙ.)"
      dbFOLDERS(5).fSTATUS = "0"
      dbFOLDERS(5).fSPEC = "00000113032   CO1 20200101            0.0077  00034852àõÕ³ñÏí³Í I Ñ³ëï³ïÙ³"
      dbFOLDERS(6).fSTATUS = "0"
      dbCONTRACT.fDGSTATE = "101"
      BuiltIn.Delay(500)
    
    'SQL Ստուգում DOCS աղուսյակում  
    Call CheckQueryRowCount("DOCS","fISN",fISN,1)
    Call CheckDB_DOCS(fISN,"COUniver","101",fBODY,1)
    
    'SQL Ստուգում DOCLOG աղուսյակում 
    Call CheckQueryRowCount("DOCLOG","fISN",fISN,3)
    Call CheckDB_DOCLOG(fISN,"77","M","99","àõÕ³ñÏí»É ¿ Ñ³ëï³ïÙ³Ý",1)
    Call CheckDB_DOCLOG(fISN,"77","C","101","",1)
    
    'SQL Ստուգում CONTRACTS աղուսյակում 
    Call CheckQueryRowCount("CONTRACTS","fDGISN",fISN,1)
    Call CheckDB_CONTRACTS(dbCONTRACT,1)
    
    'SQL Ստուգում FOLDERS աղուսյակում 
    Call CheckQueryRowCount("FOLDERS","fISN",fISN,7)
    Call CheckDB_FOLDERS(dbFOLDERS(1),1)
    Call CheckDB_FOLDERS(dbFOLDERS(2),1)
    Call CheckDB_FOLDERS(dbFOLDERS(3),1)
    Call CheckDB_FOLDERS(dbFOLDERS(4),1)
    Call CheckDB_FOLDERS(dbFOLDERS(5),1)
    Call CheckDB_FOLDERS(dbFOLDERS(6),1)
    Call CheckDB_FOLDERS(dbFOLDERS(7),1)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''-'Հաստատում է պայմանագիրը(Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I)-''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''     
    Log.Message "Verify Contract",,,DivideColor

    Call Verify_Contract("|¶»ñ³Í³Ëë|Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I",VerifyOverlimit1) 
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''-Պայմանագրեր թղթապանակում փաստատթղթի առկայության ստուգում-'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "Check (Existing Contract) Function",,,DivideColor
    
    Call ExistsContract_Filter_Fill("|¶»ñ³Í³Ëë|",ContractNew,1)
    AccountParentIsn = GetAccountIsnOverlimit()

    Log.Message "SQL Check After Verify Contract",,,SqlDivideColor
    Log.Message "AccountParentIsn = "& AccountParentIsn,,,SqlDivideColor
    
    Call SQL_Init_Overlimit_Schedule(fISN,AccountParentIsn) 
    dbFOLDERS(3).fCOM = " ¶»ñ³Í³Ëë (·ñ³ýÇÏáí å³ÛÙ.)"
    dbFOLDERS(3).fECOM = "1"
    dbCONTRACT.fDGSTATE = "7"  
    BuiltIn.Delay(500)
    
    'SQL Ստուգում DOCS աղուսյակում  
    Call CheckQueryRowCount("DOCS","fISN",fISN,1)
    Call CheckDB_DOCS(fISN,"COUniver","7",fBODY,1)

    'SQL Ստուգում DOCS աղուսյակում ¶»ñ³Í³ËëÇ Ñ³ßí³å³Ñ³Ï³Ý Ñ³í»Éí³Í-Ի համար
    fBODY = "  CODE:00000113032  CURRENCY:000  CLICOD:00034852  JURSTAT:21  VOLORT:7  PETBUJ:2  REZ:1  RELBANK:0  RABBANK:0  ACCAGR:01080793012  ACCACC:00000113032  FILLACCS:0  OPENACCS:0  TYPEPEN:0  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",AccountParentIsn,1)
    Call CheckDB_DOCS(AccountParentIsn,"COAgrAcc","2",fBODY,1)
    
    'SQL Ստուգում DOCLOG աղուսյակում 
    Call CheckQueryRowCount("DOCLOG","fISN",fISN,5)
    Call CheckDB_DOCLOG(fISN,"77","W","102","",1)
    Call CheckDB_DOCLOG(fISN,"77","T","7","",1)
    
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
'    
    'SQL Ստուգում ACCOUNTS  և HIREST աղուսյակներում
    'Row 1
    Query = "SELECT fISN FROM (SELECT *, ROW_NUMBER() OVER (ORDER BY fISN ASC) AS rownum FROM DOCP where fPARENTISN = "&AccountParentIsn&" and fISN <> 443871031) AS foo WHERE  rownum > 0 AND rownum <= 1 "
    AccountIsn = my_Row_Count(Query) 
    Call CheckQueryRowCount("ACCOUNTS","fISN",AccountIsn,1)
    Call CheckDB_DOCP(AccountIsn,"Acc     ",AccountParentIsn,1)
    Call CheckQueryRowCount("HIREST","fOBJECT",AccountIsn,1)
    'Row 2
    Query = "SELECT fISN FROM (SELECT *, ROW_NUMBER() OVER (ORDER BY fISN ASC) AS rownum FROM DOCP where fPARENTISN = "&AccountParentIsn&" and fISN <> 443871031) AS foo WHERE  rownum > 1 AND rownum <= 2 "
    AccountIsn = my_Row_Count(Query) 
    Call CheckQueryRowCount("DOCP","fPARENTISN",AccountParentIsn,5)
    Call CheckDB_DOCP(AccountIsn,"Acc     ",AccountParentIsn,1)
    Call CheckQueryRowCount("HIREST","fOBJECT",AccountIsn,1)
    'Row 3
    Query = "SELECT fISN FROM (SELECT *, ROW_NUMBER() OVER (ORDER BY fISN ASC) AS rownum FROM DOCP where fPARENTISN = "&AccountParentIsn&" and fISN <> 443871031) AS foo WHERE  rownum > 2 AND rownum <= 3 "
    AccountIsn = my_Row_Count(Query) 
    Call CheckQueryRowCount("ACCOUNTS","fISN",AccountIsn,1)
    Call CheckDB_DOCP(AccountIsn,"Acc     ",AccountParentIsn,1)
    Call CheckQueryRowCount("HIREST","fOBJECT",AccountIsn,1)
    'Row 4
    Query = "SELECT fISN FROM (SELECT *, ROW_NUMBER() OVER (ORDER BY fISN ASC) AS rownum FROM DOCP where fPARENTISN = "&AccountParentIsn&" and fISN <> 443871031) AS foo WHERE  rownum > 3 AND rownum <= 4 "
    AccountIsn = my_Row_Count(Query) 
    Call CheckQueryRowCount("ACCOUNTS","fISN",AccountIsn,1)
    Call CheckDB_DOCP(AccountIsn,"Acc     ",AccountParentIsn,1)
    Call CheckQueryRowCount("HIREST","fOBJECT",AccountIsn,1)
    
    'SQL Ստուգում CONTRACTS աղուսյակում 
    dbCONTRACT.fDGCRDTCODE = NewOverlimitDoc.GeneralTab.ExpectedCreditCode
    Call CheckQueryRowCount("CONTRACTS","fDGISN",fISN,1)
    Call CheckDB_CONTRACTS(dbCONTRACT,1)
    
    'SQL Ստուգում CAGRACCS աղուսյակում 
    Call CheckQueryRowCount("CAGRACCS","fAGRISN",fISN,1)
    
    'SQL Ստուգում FOLDERS աղուսյակում 
    Call CheckQueryRowCount("FOLDERS","fISN",fISN,5)
    Call CheckDB_FOLDERS(dbFOLDERS(1),1)
    Call CheckDB_FOLDERS(dbFOLDERS(2),1)
    Call CheckDB_FOLDERS(dbFOLDERS(3),1)
    Call CheckDB_FOLDERS(dbFOLDERS(4),1)
    Call CheckDB_FOLDERS(dbFOLDERS(6),1)
    
    Call CheckQueryRowCount("FOLDERS","fISN",AccountParentIsn,2)
    Call CheckDB_FOLDERS(dbFOLDERS(11),1)
    Call CheckDB_FOLDERS(dbFOLDERS(12),1)
    
    'SQL Ստուգում HIF  աղուսյակում 
    Call Check_HIF("2020-01-01", "N0", fISN, "11.1111", "29.00", "PPA", Null)
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''-"Գերածախս" գործողության կատարում-'''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "Check RC Create New (Overlimit/Գերածախս) Function",,,DivideColor
    
    Call Give_Overlimit(RcOptionOverlimit)

    Log.Message "SQL Check After RC Create New (Overlimit/Գերածախս) Function",,,SqlDivideColor
    Log.Message "fISN = " & RcOptionOverlimit.Isn,,,SqlDivideColor
     
    'SQL Ստուգում FOLDERS աղուսյակում 
    Query = "Select fBASE from AGRSCHEDULE WHERE fINC = '1' AND fAGRISN = " & fISN
    fBASE(0) = my_Row_Count(Query)   
    dbFOLDERS(13).fISN =  fBASE(0)
    Call CheckDB_FOLDERS(dbFOLDERS(13),1)
    
    'SQL Ստուգում DOCS աղուսյակում 
    fBODY = "  CODE:00000113032  CRDTCODE:"&NewOverlimitDoc.GeneralTab.ExpectedCreditCode&"  CLICOD:00034852  NAME:ý²àôêî  CURRENCY:000  ACCACC:00000113032  COMMENT:MaxSinvolsSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSS1  AUTODEBT:1  ACCCONNMODE:3  USECLICONNSCH:1  ACCCONNSCH:002  DATE:20200101  DATEGIVE:20200101  ACSBRANCH:01  ACSDEPART:4  ACSTYPE:CO1  DATESFILLTYPE:2  AGRMARBEG:20200101  AGRMARFIN:20790101  AGRPERIOD:10/29  PASSOVDIRECTION:2  PASSOVTYPE:13  FIXEDROWSINSCH:0  SUMSDATESFILLTYPE:2  SUMSFILLTYPE:01  MIXEDSUMSINSCH:0  KINDSCALE:1  PCPENAGR:11.1111/29  CONSTPER:0  SECTOR:A  USAGEFIELD:03.006  AIM:04  INTERORG:BSTDB  SCHEDULE:9  GUARANTEE:1  COUNTRY:CH  LRDISTR:004  REGION:050000009  PERRES:1  NOTE:01  NOTE2:022  NOTE3:03  PPRCODE:88888888888888888881  SUBJRISK:1  CHRGFIRSTDAY:1  GIVEN:1  OPRISNLIST:"&RcOptionOverlimit.Isn&"  NTFMODE:1  SENDSTMADRS:2  PUTINLR:0  LRMRTCUR:000  ACRANOTE:02  NOTCLASS:1  REVISIONREASON:1  REPSOURCE:4  MORTSUBJECT:16  OTHERCOLLATERAL:1  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",fISN,1) 
    Call CheckDB_DOCS(fISN,"COUniver","7",fBODY,1)
    
    fBODY = "  CODE:00000113032  DATE:20200101  SUMMA:100121500  CASHORNO:2  ACCCORR:00000113032  COMMENT:RcOptionOverlimitRcOptionOverlimitRcOptionOverlimitRcOptionOverlimitRcOptionOverlimitRcOptionOverli1  MARDOCISN:"&fBASE(0)&"  ACSBRANCH:01  ACSDEPART:4  ACSTYPE:CO1  USERID:  77  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",RcOptionOverlimit.Isn,1)
    Call CheckDB_DOCS(RcOptionOverlimit.Isn,"CODSAgr ","5",fBODY,1)

    'SQL Ստուգում DOCLOG աղուսյակում 
    Call CheckQueryRowCount("DOCLOG","fISN",RcOptionOverlimit.Isn,4)
    Call CheckDB_DOCLOG(RcOptionOverlimit.Isn,"77","N","1","",1)
    Call CheckDB_DOCLOG(RcOptionOverlimit.Isn,"77","M","99","àõÕ³ñÏí»É ¿ Ñ³ëï³ïÙ³Ý",1)
    Call CheckDB_DOCLOG(RcOptionOverlimit.Isn,"77","T","2","",1)
    Call CheckDB_DOCLOG(RcOptionOverlimit.Isn,"77","C","5","",1)
    
    Call CheckQueryRowCount("DOCLOG","fISN",fISN,6)
    Call CheckDB_DOCLOG(fISN,"77","E","7","",1)

    'SQL Ստուգում HIR աղուսյակում 
    Call CheckQueryRowCount("HIR","fOBJECT",fISN,1)
    Call Check_HIR("2020-01-01", "R1", fISN, "000", "100121500.00", "AGR", "D")

    'SQL Ստուգում HIRREST  աղուսյակում 
    Call CheckQueryRowCount("HIRREST","fOBJECT",fISN,1)
    Call CheckDB_HIRREST("R1",fISN,"100121500.00","2020-01-01",1)
    
    'SQL Ստուգում HI աղուսյակում 
    Call CheckQueryRowCount("HI","fBASE",RcOptionOverlimit.Isn,2)
    Call Check_HI_CE_accounting ("20200101",RcOptionOverlimit.Isn, "01", "443871031", "100121500.00", "000", "100121500.00", "OVD", "C") 
    
    Query = "Select fOBJECT From HI Where fBASE = "& RcOptionOverlimit.Isn &" and fDBCR = 'D'"
    fOBJECT(0) = my_Row_Count(Query)                  
    Call Check_HI_CE_accounting ("20200101",RcOptionOverlimit.Isn, "01",  fOBJECT(0), "100121500.00", "000", "100121500.00", "OVD", "D")
    
    'SQL Ստուգում HIREST  աղուսյակում 
    Call CheckDB_HIREST("01", fOBJECT(0),"100121500.00","000","100121500.00",1)    

    'SQL Ստուգում AGRSCHEDULE աղուսյակում 
    Call Check_AGRSCHEDULE(fISN, "20200101", "20200101", "1", "0", "3")
    
    'SQL Ստուգում AGRSCHEDULEVALUES աղուսյակում 
    Call Check_AGRSCHEDULEVALUES(fISN, "1", "20201202", "100121500.00", "1", "0")
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''-"Տոկոսների հաշվարկ" գործողության կատարում-''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''    
    Log.Message "Check RC (Calculate Percents/Տոկոսների հաշվարկ) Function",,,DivideColor
    
    Call CalculatePercents(CalcPercents,"",False)
    
    Log.Message "SQL Check After RC (Calculate Percents/Տոկոսների հաշվարկ) Function",,,SqlDivideColor
    Log.Message "fISN = " & CalcPercents.Isn,,,SqlDivideColor

    'SQL Ստուգում DOCS աղուսյակում 
    fBODY = "  CODE:00000113032  DATECHARGE:20200101  DATE:20200101  SUMAGRPEN:22222222.1/11111111.1  SUMALLPEN:22222222.1/11111111.1  COMMENT:CalcPercentsCalcPercentsCalcPercentsCalcPercentsCalcPercentsCalcPercentsCalcPercentsCalcPercentsCal1  ACSBRANCH:01  ACSDEPART:4  ACSTYPE:CO1  USERID:  77  "
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
    Call Check_HIF ("20200101", "N0", fISN, "0.00", "0.00", "AGJ", "0")
    Call Check_HIF ("20200101", "N0", fISN, "0.00", "0.00", "DTC", "20200101")
    Call Check_HIF ("20200101", "N0", fISN, "0.00", "0.00", "LIM", Null)
    Call Check_HIF ("20200101", "N0", fISN, "11.1111", "29.00", "PPA", Null)
    
    'SQL Ստուգում HIR աղուսյակում 
    Call CheckQueryRowCount("HIR","fOBJECT",fISN,3)
    Call Check_HIR("20200101", "R3", fISN, "000", "22222222.10", "PNA", "D")
    Call Check_HIR("20200101", "R7", fISN, "000", "11111111.10", "PNA", "D")
    
    'SQL Ստուգում HIT աղուսյակում 
    Call CheckQueryRowCount("HIT","fOBJECT",fISN,2)
    Call Check_HIT("20200101", "N3", fISN, "000", "22222222.10", "PNA", "D")
    Call Check_HIT("20200101", "N7", fISN, "000", "11111111.10", "PNA", "D")
    
    'SQL Ստուգում HIRREST  աղուսյակում 
    Call CheckQueryRowCount("HIRREST","fOBJECT",fISN,3)
    Call CheckDB_HIRREST("R3",fISN,"22222222.10","20200101",1)
    Call CheckDB_HIRREST("R7",fISN,"11111111.10","20200101",1)    
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''-"Պարտքերի մարում" գործողության կատարում-'''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''      
    Log.Message "Check RC (Overlimit Repay/Պարտքերի մարում) Function",,,DivideColor

    Call Overlimit_Repay(NewOverlimitRepay)
    
    Log.Message "SQL Check After RC (Overlimit Repay/Պարտքերի մարում) Function",,,SqlDivideColor
    Log.Message "fISN = " & NewOverlimitRepay.Isn,,,SqlDivideColor
    
    'SQL Ստուգում FOLDERS աղուսյակում 
    Query = "Select fBASE from AGRSCHEDULE WHERE fINC = '2' AND fAGRISN = " & fISN
    fBASE(1) = my_Row_Count(Query)   
    dbFOLDERS(13).fISN =  fBASE(1)
    dbFOLDERS(13).fSPEC = "2Ø³ñáõÙÝ»ñÇ ·ñ³ýÇÏ`  01/01/20"     
    Call CheckDB_FOLDERS(dbFOLDERS(13),1)
    
    'SQL Ստուգում DOCS աղուսյակում  
    fBODY = "  CODE:00000113032  DATE:20200101  SUMAGR:22200022.1  SUMFINE:11111111  SUMMA:33311133.1  CASHORNO:2  ISPUSA:0  ACCCORR:00000113032  COMMENT:NewOverlimitRepayNewOverlimitRepayNewOverlimitRepayNewOverlimitRepayNewOverlimitRepayNewOverlimitRe1  MARDOCISN:"&fBASE(1)&"  ACSBRANCH:01  ACSDEPART:4  ACSTYPE:CO1  USERID:  77  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",NewOverlimitRepay.Isn,1)
    Call CheckDB_DOCS(NewOverlimitRepay.Isn,"CODSDebt","5",fBODY,1)
    
    'SQL Ստուգում DOCP աղուսյակում  
    Call CheckQueryRowCount("DOCP","fPARENTISN",AccountParentIsn,7)

    'SQL Ստուգում DOCLOG աղուսյակում 
    Call CheckQueryRowCount("DOCLOG","fISN",fISN,6)
    Call CheckQueryRowCount("DOCLOG","fISN",NewOverlimitRepay.Isn,4)
    Call CheckDB_DOCLOG(NewOverlimitRepay.Isn,"77","N","1","",1)
    Call CheckDB_DOCLOG(NewOverlimitRepay.Isn,"77","M","99","àõÕ³ñÏí»É ¿ Ñ³ëï³ïÙ³Ý",1)
    Call CheckDB_DOCLOG(NewOverlimitRepay.Isn,"77","T","2","",1)
    Call CheckDB_DOCLOG(NewOverlimitRepay.Isn,"77","C","5","",1)
    
    'SQL Ստուգում CAGRACCS աղուսյակում 
    Call CheckQueryRowCount("CAGRACCS","fAGRISN",fISN,1)
    
    'SQL Ստուգում HI աղուսյակում 
    Call CheckQueryRowCount("HI","fBASE",NewOverlimitRepay.Isn,6)
    Call Check_HI_CE_accounting ("20200101",NewOverlimitRepay.Isn, "01", "443871031", "22200022.10", "000", "22200022.10", "OVD", "D") 
    Call Check_HI_CE_accounting ("20200101",NewOverlimitRepay.Isn, "01", "443871031", "11111111.00", "000", "11111111.00", "PEN", "D") 
    
    Query = "Select fOBJECT From HI Where fOP = 'OVD' and fDBCR = 'C' and fSUM = '22200022.10'"
    fOBJECT(1) = my_Row_Count(Query)                  
    Call Check_HI_CE_accounting ("20200101",NewOverlimitRepay.Isn, "01", fOBJECT(1), "22200022.10", "000", "22200022.10", "OVD", "C")
    'SQL Ստուգում HIREST  աղուսյակում 
    Call CheckDB_HIREST("01",fOBJECT(1),"77921477.90","000","77921477.90",1)
    'SQL Ստուգում DOCP աղուսյակում համապատասխան fOBJECT-ով
    Call CheckDB_DOCP(fOBJECT(1),"Acc     ",AccountParentIsn,1)
    
    Query = "Select fOBJECT From HI Where fOP = 'PRC' and fDBCR = 'C' and fSUM = '11111111.00'"
    fOBJECT(2) = my_Row_Count(Query)                  
    Call Check_HI_CE_accounting ("20200101",NewOverlimitRepay.Isn, "01", fOBJECT(2), "11111111.00", "000", "11111111.00", "PRC", "C")
    'SQL Ստուգում HIREST  աղուսյակում 
    Call CheckDB_HIREST("01",fOBJECT(2),"-11111111.00","000","-11111111.00",1)
    'SQL Ստուգում DOCP աղուսյակում համապատասխան fOBJECT-ով
    Call CheckDB_DOCP(fOBJECT(2),"Acc     ",AccountParentIsn,1)
    
    Query = "Select fOBJECT From HI Where fOP = 'PRC' and fDBCR = 'D' and fSUM = '11111111.00'"
    fOBJECT(3) = my_Row_Count(Query) 
    Call Check_HI_CE_accounting ("20200101",NewOverlimitRepay.Isn, "01", fOBJECT(3), "11111111.00", "000", "11111111.00", "PRC", "D")
    Call Check_HI_CE_accounting ("20200101",NewOverlimitRepay.Isn, "01", fOBJECT(3), "11111111.00", "000", "11111111.00", "PEN", "C")
    'SQL Ստուգում DOCP աղուսյակում համապատասխան fOBJECT-ով
    Call CheckDB_DOCP(fOBJECT(3),"Acc     ",AccountParentIsn,1)
    
    'SQL Ստուգում HIR աղուսյակում 
    Call CheckQueryRowCount("HIR","fBASE",NewOverlimitRepay.Isn,2)
    Call Check_HIR("20200101", "R1", fISN, "000", "22200022.10", "DBT", "C")
    Call Check_HIR("20200101", "R3", fISN, "000", "11111111.00", "DBT", "C")
    
    'SQL Ստուգում HIRREST  աղուսյակում 
    Call CheckQueryRowCount("HIRREST","fOBJECT",fISN,3)
    Call CheckDB_HIRREST("R1",fISN,"77921477.90","20200101",1)
    Call CheckDB_HIRREST("R3",fISN,"11111111.10","20200101",1)
    Call CheckDB_HIRREST("R7",fISN,"11111111.10","20200101",1)
    
    'SQL Ստուգում AGRSCHEDULE աղուսյակում 
    Call Check_AGRSCHEDULE(fISN, "20200101", "20200101", "2", "0", "2")
    
    'SQL Ստուգում AGRSCHEDULEVALUES աղուսյակում 
    Call Check_AGRSCHEDULEVALUES(fISN, "2", "20201202", "77921477.90", "1", "0")
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''-"Տոկոսադրույքներ" գործողության կատարում-'''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''      
    Log.Message "Check RC (Overlimit Rete) Function",,,DivideColor    
    
    ActionIsn(1) = ChangeOverlimitRete("00000113032", "010120", "1092.1001", "29","ChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChan1")

    Log.Message "SQL Check After RC (Overlimit Rete/Տոկոսադրույքներ) Function",,,SqlDivideColor
    Log.Message "fISN = " & ActionIsn(1),,,SqlDivideColor
    
    'SQL Ստուգում DOCS աղուսյակում  
    fBODY = "  CODE:00000113032  DATE:20200101  PCPENAGR:1092.1001/29  COMMENT:ChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChan1  USERID:  77  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",ActionIsn(1),1)
    Call CheckDB_DOCS(ActionIsn(1),"COTSPC  ","5",fBODY,1)
    
    'SQL Ստուգում DOCLOG աղուսյակում 
    Call CheckQueryRowCount("DOCLOG","fISN",fISN,6)
    Call CheckQueryRowCount("DOCLOG","fISN",ActionIsn(1),4)
    Call CheckDB_DOCLOG(ActionIsn(1),"77","N","1","",1)
    Call CheckDB_DOCLOG(ActionIsn(1),"77","M","99","àõÕ³ñÏí»É ¿ Ñ³ëï³ïÙ³Ý",1)
    Call CheckDB_DOCLOG(ActionIsn(1),"77","T","2","",1)
    Call CheckDB_DOCLOG(ActionIsn(1),"77","C","5","",1)
    
    'SQL Ստուգում HIF աղուսյակում 
    Call Check_HIF ("20200101", "N0", fISN, "1092.1001", "29.00", "PPA", Null)
    
    'SQL Ստուգում FOLDERS  աղուսյակում 0
    Set dbFOLDERS_ForRate = New_DB_FOLDERS()
        dbFOLDERS_ForRate.fFOLDERID = "Agr." & fISN
        dbFOLDERS_ForRate.fNAME = "COTSPC  "
        dbFOLDERS_ForRate.fKEY = ActionIsn(1)
        dbFOLDERS_ForRate.fISN = ActionIsn(1)
        dbFOLDERS_ForRate.fSTATUS = "1"
        dbFOLDERS_ForRate.fCOM = "îáÏáë³¹ñáõÛùÝ»ñ"
        dbFOLDERS_ForRate.fSPEC = "1îáÏáë³¹ñáõÛùÝ»ñ`  01/01/20,  { ,  1092.1001/29 }"  
        
    Call CheckDB_FOLDERS(dbFOLDERS_ForRate,1)    
'    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''-"Պայմաններ և վիճակներ/Օբյեկտիվ ռիսկի դասիչ" գործողության կատարում-''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''      
    Log.Message "Check RC (Objective Risk/Օբյեկտիվ ռիսկի դասիչ) Function",,,DivideColor    

    ActionIsn(2) = Objective_Risk("00000113032","010120", "02", "ChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChan1",False)
                                                                                         
    Log.Message "SQL Check After RC (Objective Risk/Օբյեկտիվ ռիսկի դասիչ) Function",,,SqlDivideColor
    Log.Message "fISN = " & ActionIsn(2),,,SqlDivideColor
                                                                                                            
    'SQL Ստուգում DOCS աղուսյակում 
    fBODY = "  CODE:00000113032  DATE:20200101  RISK:02  COMMENT:ChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChan1  USERID:  77  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",ActionIsn(2),1)
    Call CheckDB_DOCS(ActionIsn(2),"COTSORC ","5",fBODY,1)
    
    'SQL Ստուգում DOCLOG աղուսյակում 
    Call CheckQueryRowCount("DOCLOG","fISN",fISN,6)
    
    Call CheckQueryRowCount("DOCLOG","fISN",ActionIsn(2),4)
    Call CheckDB_DOCLOG(ActionIsn(2),"77","N","1","",1)
    Call CheckDB_DOCLOG(ActionIsn(2),"77","M","99","àõÕ³ñÏí»É ¿ Ñ³ëï³ïÙ³Ý",1)
    Call CheckDB_DOCLOG(ActionIsn(2),"77","T","2","",1)
    Call CheckDB_DOCLOG(ActionIsn(2),"77","C","5","",1)
    
    'SQL Ստուգում HIF աղուսյակում 
    Call Check_HIF ("20200101", "N0", fISN, "0.00", "0.00", "ORC", "02")   

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''-"Ռիսկի դասիչ և պահուստավորման տոկոս" գործողության կատարում-''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''      
    Log.Message "Check RC (Risk Classifier/Ռիսկի դասիչ և պահուստավորման տոկոս) Function",,,DivideColor   
    
    ActionIsn(3) = Create_Risk_Classifier("00000113032","010120", "01", "2", "ChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChan1")

    Log.Message "SQL Check After RC (Risk Classifier/Ռիսկի դասիչ և պահուստավորման տոկոս) Function",,,SqlDivideColor
    Log.Message "fISN = " & ActionIsn(3),,,SqlDivideColor
                                                                                                            
    'SQL Ստուգում DOCS աղուսյակում 
    fBODY = "  CODE:00000113032  DATE:20200101  RISK:01  PERRES:2  COMMENT:ChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChan1  USERID:  77  "
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
    Call Check_HIF("20200101", "N0", fISN, "2.00", "0.00", "PRS", Null)  
    Call Check_HIF("20200101", "N0", fISN, "0.00", "0.00", "RSK", "01")     

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''-"Պահուստավորում" գործողության կատարում-'''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''      
    Log.Message "Check RC (Store/Պահուստավորում) Function",,,DivideColor   

    Call Doc_Store(NewStore)
    
    Log.Message "SQL Check After RC (Store/Պահուստավորում) Function",,,SqlDivideColor
    Log.Message "fISN = " & NewStore.Isn,,,SqlDivideColor
    
    'SQL Ստուգում DOCS աղուսյակում 
    fBODY = "  CODE:00000113032  DATE:20200101  SUMRES:100000000  COMMENT:NewStoreNewStoreNewStoreNewStoreNewStoreNewStoreNewStoreNewStoreNewStoreNewStoreNewStoreNewStoreNew1  ACSBRANCH:01  ACSDEPART:3  ACSTYPE:CO1  USERID:  77  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",NewStore.Isn,1)
    Call CheckDB_DOCS(NewStore.Isn,"CODSRes ","5",fBODY,1)
    
    'SQL Ստուգում DOCLOG աղուսյակում 
    Call CheckQueryRowCount("DOCLOG","fISN",fISN,6)
    
    Call CheckQueryRowCount("DOCLOG","fISN",NewStore.Isn,4)
    Call CheckDB_DOCLOG(NewStore.Isn,"77","N","1","",1)
    Call CheckDB_DOCLOG(NewStore.Isn,"77","M","99","àõÕ³ñÏí»É ¿ Ñ³ëï³ïÙ³Ý",1)
    Call CheckDB_DOCLOG(NewStore.Isn,"77","T","2","",1)
    Call CheckDB_DOCLOG(NewStore.Isn,"77","C","5","",1)
    
    'SQL Ստուգում HI աղուսյակում 
    Call CheckQueryRowCount("HI","fBASE",NewStore.Isn,2)
    Query = "Select fOBJECT From HI Where fOP = 'RST' and fDBCR = 'C' and fSUM = '100000000.00'"
    fOBJECT(4) = my_Row_Count(Query)  
    Call Check_HI_CE_accounting ("20200101",NewStore.Isn, "01", fOBJECT(4), "100000000.00", "000", "100000000.00", "RST", "C") 
    
    'SQL Ստուգում HIREST  աղուսյակում համապատասխան fOBJECT-ով
    Call CheckDB_HIREST("01",fOBJECT(4),"-100000000.00","000","-100000000.00",1)
    
    'SQL Ստուգում HI աղուսյակում 
    Query = "Select fOBJECT From HI Where fOP = 'RST' and fDBCR = 'D' and fSUM = '100000000.00'"
    fOBJECT(5) = my_Row_Count(Query)  
    Call Check_HI_CE_accounting ("20200101",NewStore.Isn, "01", fOBJECT(5), "100000000.00", "000", "100000000.00", "RST", "D") 
    
    'SQL Ստուգում HIREST  աղուսյակում համապատասխան fOBJECT-ով
    Call CheckDB_HIREST("01",fOBJECT(5),"100000000.00","000","100000000.00",1)
    
    'SQL Ստուգում ACCOUNTS աղուսյակում համապատասխան fOBJECT-ով
    Call CheckQueryRowCount("ACCOUNTS","fISN",fOBJECT(4),1)
    Call CheckQueryRowCount("ACCOUNTS","fISN",fOBJECT(5),1)
    
    'SQL Ստուգում HIR աղուսյակում 
    Call CheckQueryRowCount("HIR","fOBJECT",fISN,6)
    Call Check_HIR("20200101", "R1", fISN, "000", "100121500.00", "AGR", "D")
    Call Check_HIR("20200101", "R1", fISN, "000", "22200022.10", "DBT", "C")
    Call Check_HIR("20200101", "R3", fISN, "000", "22222222.10", "PNA", "D")
    Call Check_HIR("20200101", "R3", fISN, "000", "11111111.00", "DBT", "C")
    Call Check_HIR("20200101", "R4", fISN, "000", "100000000.00", "RES", "D")
    Call Check_HIR("20200101", "R7", fISN, "000", "11111111.10", "PNA", "D")
    
    'SQL Ստուգում HIRREST  աղուսյակում 
    Call CheckQueryRowCount("HIRREST","fOBJECT",fISN,4)
    Call CheckDB_HIRREST("R4",fISN,"100000000.00","20200101",1)    

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''-"Տոկոսների հաշվարկ" գործողության կատարում-''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''    
    Log.Message "Check RC Create (Calculate Percents/Տոկոսների հաշվարկ) Function",,,DivideColor
    
    ExpectedMessage = "¶»ñ³Í³Ëë³ÛÇÝ å³ÛÙ³Ý³·Çñª  00000113032  /ý²àôêî/"& vbCrLf &"--------------------------------------------------------------------------------------------------------------"& vbCrLf &""& vbCrLf &"01/01/20-ÇÝ Ï³ï³ñí»É ¿ ïáÏáëÇ Ñ³ßí³ñÏáõÙ"
    Call CalculatePercents(CalcPercents_2,ExpectedMessage,True)
    
    Log.Message "SQL Check After RC (Calculate Percents/Տոկոսների հաշվարկ) Function",,,SqlDivideColor
    Log.Message "fISN = " & CalcPercents_2.Isn,,,SqlDivideColor

    'SQL Ստուգում DOCS աղուսյակում 
    fBODY = "  CODE:00000113032  DATECHARGE:20200101  DATE:20200101  SUMAGRPEN:987.1/123.1  SUMALLPEN:987.1/123.1  ACSBRANCH:00  ACSDEPART:4  ACSTYPE:CO1  USERID:  77  "
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
    
    Query = "Select * from HIF where fDATE = '20200101' and fTYPE = 'N0' and fOBJECT = "&fISN&" and fOP = 'DTC' and fSUM = 0.00 and fCURSUM = 0.00 and fSPEC = 20200101	and fBASE = "&CalcPercents_2.Isn
    If my_Row_Count(Query) = 1 Then
        Log.Message "HIF record is Correct" 
    End If
    
    'SQL Ստուգում HIR աղուսյակում 
    Call CheckQueryRowCount("HIR","fBASE",CalcPercents_2.Isn,2)
    Call Check_HIR("20200101", "R3", fISN, "000", "987.10", "PNA", "D")
    Call Check_HIR("20200101", "R7", fISN, "000", "123.10", "PNA", "D")
    
    'SQL Ստուգում HIT աղուսյակում 
    Call CheckQueryRowCount("HIT","fBASE",CalcPercents_2.Isn,2)
    Call Check_HIT("20200101", "N3", fISN, "000", "987.10", "PNA", "D")
    Call Check_HIT("20200101", "N7", fISN, "000", "123.10", "PNA", "D")
    
    'SQL Ստուգում HIRREST  աղուսյակում 
    Call CheckQueryRowCount("HIRREST","fOBJECT",fISN,4)
    Call CheckDB_HIRREST("R1",fISN,"77921477.90","20200101",1)
    Call CheckDB_HIRREST("R3",fISN,"11112098.20","20200101",1)
    Call CheckDB_HIRREST("R4",fISN,"100000000.00","20200101",1)
    Call CheckDB_HIRREST("R7",fISN,"11111234.20","20200101",1)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''-"Դուրս գրում" գործողության կատարում-''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''    
    Log.Message "Check RC (Write Out/Դուրս գրում) Function",,,DivideColor      
    
    Call Create_WriteOut(NewWriteOut)
    
    Log.Message "SQL Check After RC (Write Out/Դուրս գրում) Function",,,SqlDivideColor
    Log.Message "fISN = " & NewWriteOut.Isn,,,SqlDivideColor
    
    'SQL Ստուգում DOCS աղուսյակում 
    fBODY = "  CODE:00000113032  DATE:20200101  SUMAGR:100  SUMFINE:764  SUMMA:864  COMMENT:NewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOut1  ACSBRANCH:01  ACSDEPART:4  ACSTYPE:CO1  USERID:  77  "
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
    fOBJECT(6) = my_Row_Count(Query)      
    'SQL Ստուգում HI աղուսյակում  համապատասխան fOBJECT-ով     
    Call Check_HI_CE_accounting ("20200101",NewWriteOut.Isn, "01", fOBJECT(6), "100.00", "000", "100.00", "MSC", "D")
    'SQL Ստուգում HIREST աղուսյակում համապատասխան fOBJECT-ով
    Call CheckDB_HIREST("01",fOBJECT(6),"-99999900.00","000","-99999900.00",1)
    'SQL Ստուգում DOCP աղուսյակում համապատասխան fOBJECT-ով
    Call CheckDB_DOCP(fOBJECT(6),"Acc     ",AccountParentIsn,1)
    
    Query = "Select fOBJECT From HI Where fOP = 'MSC' and fDBCR = 'C' and fTYPE = '01' AND fBASE = "&NewWriteOut.Isn
    fOBJECT(7) = my_Row_Count(Query)      
    'SQL Ստուգում HI աղուսյակում  համապատասխան fOBJECT-ով         
    Call Check_HI_CE_accounting ("20200101",NewWriteOut.Isn, "01", fOBJECT(7), "100.00", "000", "100.00", "MSC", "C")
    'SQL Ստուգում HIREST աղուսյակում համապատասխան fOBJECT-ով
    Call CheckDB_HIREST("01",fOBJECT(7),"77921377.90","000","77921377.90",1)
    'SQL Ստուգում DOCP աղուսյակում համապատասխան fOBJECT-ով
    Call CheckDB_DOCP(fOBJECT(7),"Acc   ",AccountParentIsn,1)
    
    Query = "Select fOBJECT From HI Where fOP = 'MSC' and fDBCR = 'D' and fTYPE = '02' AND fBASE = "&NewWriteOut.Isn
    fOBJECT(8) = my_Row_Count(Query)      
    'SQL Ստուգում HI աղուսյակում  համապատասխան fOBJECT-ով         
    Call Check_HI_CE_accounting ("20200101",NewWriteOut.Isn, "02", fOBJECT(8), "100.00", "000", "100.00", "MSC", "D")
    'SQL Ստուգում HIREST աղուսյակում համապատասխան fOBJECT-ով
    Call CheckDB_HIREST("02",fOBJECT(8),"0.00","000","0.00",1)
    'SQL Ստուգում DOCP աղուսյակում համապատասխան fOBJECT-ով
    Call CheckDB_DOCP(fOBJECT(8),"NBAcc   ",AccountParentIsn,1)
    'SQL Ստուգում NBACCOUNTS  աղուսյակում համապատասխան fOBJECT-ով
    Call CheckQueryRowCount("NBACCOUNTS","fISN",fOBJECT(8),1)
    
    'SQL Ստուգում HIR աղուսյակում 
    Call CheckQueryRowCount("HIR","fBASE",NewWriteOut.Isn,3)
    Call Check_HIR("20200101", "R4", fISN, "000", "100.00", "OUT", "C")
    Call Check_HIR("20200101", "R5", fISN, "000", "100.00", "OUT", "D")
    Call Check_HIR("20200101", "R7", fISN, "000", "764.00", "OUT", "D")
    
    'SQL Ստուգում HIRREST  աղուսյակում 
    Call CheckQueryRowCount("HIRREST","fOBJECT",fISN,5)
    Call CheckDB_HIRREST("R1",fISN,"77921477.90","20200101",1)
    Call CheckDB_HIRREST("R3",fISN,"11112098.20","20200101",1)
    Call CheckDB_HIRREST("R4",fISN,"99999900.00","20200101",1)
    Call CheckDB_HIRREST("R5",fISN,"100.00","20200101",1)
    Call CheckDB_HIRREST("R7",fISN,"11111998.20","20200101",1)
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''-"Դուրս գրածի վերականգնում" գործողության կատարում-''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''    
    Log.Message "Check RC (Write Off Reconstruction/Դուրս գրածի վերականգնում) Function",,,DivideColor       
    
    Call WriteOut_Reconstruction(NewWriteOff,True)
    
    Log.Message "SQL Check After RC (Write Off Reconstruction/Դուրս գրածի վերականգնում) Function",,,SqlDivideColor
    Log.Message "fISN = " & NewWriteOff.Isn,,,SqlDivideColor
    
    'SQL Ստուգում DOCS աղուսյակում 
    fBODY = "  CODE:00000113032  DATE:20200101  SUMAGR:100  SUMFINE:11111998.2  SUMMA:11112098.2  COMMENT:wWriteOffwWriteOffwWriteOffwWriteOffwWriteOffwWriteOffwWriteOffwWriteOffwWriteOffwWriteOffwWriteOff1  ACSBRANCH:01  ACSDEPART:3  ACSTYPE:CO1  USERID:  77  "
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
    fOBJECT(9) = my_Row_Count(Query)      
    'SQL Ստուգում HI աղուսյակում  համապատասխան fOBJECT-ով         
    Call Check_HI_CE_accounting ("20200101",NewWriteOff.Isn, "01", fOBJECT(9), "100.00", "000", "100.00", "MSC", "C")
    'SQL Ստուգում HIREST աղուսյակում համապատասխան fOBJECT-ով
    Call CheckDB_HIREST("01",fOBJECT(9),"-100000000.00","000","-100000000.00",1)
    
    Query = "Select fOBJECT From HI Where fOP = 'MSC' and fDBCR = 'D' and fSUM = '100.00' and fBASEDEPART = '3' AND fBASE ="&NewWriteOff.Isn
    fOBJECT(10) = my_Row_Count(Query)      
    'SQL Ստուգում HI աղուսյակում  համապատասխան fOBJECT-ով         
    Call Check_HI_CE_accounting ("20200101",NewWriteOff.Isn, "01", fOBJECT(10), "100.00", "000", "100.00", "MSC", "D")
    'SQL Ստուգում HIREST աղուսյակում համապատասխան fOBJECT-ով
    Call CheckDB_HIREST("01",fOBJECT(10),"77921477.90","000","77921477.90",1)
    
    Query = "Select fOBJECT From HI Where fOP = 'MSC' and fDBCR = 'C' and fSUM = '100.00' and fBASEDEPART = '4' AND fBASE = "&NewWriteOff.Isn
    fOBJECT(11) = my_Row_Count(Query)      
    'SQL Ստուգում HI աղուսյակում  համապատասխան fOBJECT-ով         
    Call Check_HI_CE_accounting ("20200101",NewWriteOff.Isn, "02", fOBJECT(11), "100.00", "000", "100.00", "MSC", "C")
    
    'SQL Ստուգում HIR աղուսյակում 
    Call CheckQueryRowCount("HIR","fBASE",NewWriteOff.Isn,4)
    Call Check_HIR("20200101", "R4", fISN, "000", "100.00", "INC", "D")
    Call Check_HIR("20200101", "R5", fISN, "000", "100.00", "INC", "C")
    Call Check_HIR("20200101", "R7", fISN, "000", "11111998.20", "INC", "C")
    Call Check_HIR("20200101", "RI", fISN, "000", "11111234.20", "IR7", "D")
    
    'SQL Ստուգում HIRREST  աղուսյակում 
    Call CheckQueryRowCount("HIRREST","fOBJECT",fISN,6)
    Call CheckDB_HIRREST("R1",fISN,"77921477.90","20200101",1)
    Call CheckDB_HIRREST("R3",fISN,"11112098.20","20200101",1)
    Call CheckDB_HIRREST("R4",fISN,"100000000.00","20200101",1)
    Call CheckDB_HIRREST("R5",fISN,"0.00","20200101",1)
    Call CheckDB_HIRREST("R7",fISN,"0.00","20200101",1)
    Call CheckDB_HIRREST("RI",fISN,"11111234.20","20200101",1)
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''-"Պարտքերի մարում" գործողության կատարում-'''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''      
    Log.Message "Check RC (Overlimit Repay/Պարտքերի մարում) Function",,,DivideColor

    Call Overlimit_Repay(NewOverlimitRepay2)

    Log.Message "SQL Check After RC (Overlimit Repay/Պարտքերի մարում) Function",,,SqlDivideColor
    Log.Message "fISN = " & NewOverlimitRepay2.Isn,,,SqlDivideColor
    
    'SQL Ստուգում FOLDERS աղուսյակում 
    Query = "Select fBASE from AGRSCHEDULE WHERE fINC = '3' AND fAGRISN = " & fISN
    fBASE(2) = my_Row_Count(Query)   
    dbFOLDERS(13).fISN =  fBASE(2)
    dbFOLDERS(13).fSPEC = "2Ø³ñáõÙÝ»ñÇ ·ñ³ýÇÏ`  01/01/20"     
    Call CheckDB_FOLDERS(dbFOLDERS(13),1)
    
    'SQL Ստուգում DOCS աղուսյակում  
    fBODY = "  CODE:00000113032  DATE:20200101  SUMAGR:77921477.9  SUMFINE:11112098.2  SUMMA:89033576.1  CASHORNO:2  ISPUSA:0  ACCCORR:00000113032  COMMENT:NewOverlimitRepayNewOverlimitRepayNewOverlimitRepayNewOverlimitRepayNewOverlimitRepayNewOverlimitRe1  MARDOCISN:"&fBASE(2)&"  ACSBRANCH:01  ACSDEPART:4  ACSTYPE:CO1  USERID:  77  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",NewOverlimitRepay2.Isn,1)
    Call CheckDB_DOCS(NewOverlimitRepay2.Isn,"CODSDebt","5",fBODY,1)
    
    'SQL Ստուգում DOCLOG աղուսյակում 
    Call CheckQueryRowCount("DOCLOG","fISN",NewOverlimitRepay2.Isn,4)
    Call CheckDB_DOCLOG(NewOverlimitRepay2.Isn,"77","N","1","",1)
    Call CheckDB_DOCLOG(NewOverlimitRepay2.Isn,"77","M","99","àõÕ³ñÏí»É ¿ Ñ³ëï³ïÙ³Ý",1)
    Call CheckDB_DOCLOG(NewOverlimitRepay2.Isn,"77","T","2","",1)
    Call CheckDB_DOCLOG(NewOverlimitRepay2.Isn,"77","C","5","",1)
    
    'SQL Ստուգում HI աղուսյակում 
    Call CheckQueryRowCount("HI","fBASE",NewOverlimitRepay2.Isn,6)
    Call Check_HI_CE_accounting ("20200101",NewOverlimitRepay2.Isn, "01", "443871031", "77921477.90", "000", "77921477.90", "OVD", "D") 
    Call Check_HI_CE_accounting ("20200101",NewOverlimitRepay2.Isn, "01", "443871031", "11112098.20", "000", "11112098.20", "PEN", "D") 
    
    'SQL Ստուգում HIREST  աղուսյակում 
    Call CheckDB_HIREST("01","443871031","-977896790.80","000","-977896790.80",1)

    Query = "Select fOBJECT From HI Where fOP = 'PRC' and fDBCR = 'C' and fBASE = "& NewOverlimitRepay2.Isn
    fOBJECT(12) = my_Row_Count(Query)                  
    Call Check_HI_CE_accounting ("20200101",NewOverlimitRepay2.Isn, "01", fOBJECT(12), "11112098.20", "000", "11112098.20", "PRC", "C")
    'SQL Ստուգում HIREST  աղուսյակում 
    Call CheckDB_HIREST("01",fOBJECT(12),"-22223209.20","000","-22223209.20",1)

    Query = "Select fOBJECT From HI Where fOP = 'PRC' and fDBCR = 'D' and fBASE = "& NewOverlimitRepay2.Isn
    fOBJECT(13) = my_Row_Count(Query) 
    Call Check_HI_CE_accounting ("20200101",NewOverlimitRepay2.Isn, "01", fOBJECT(13), "11112098.20", "000", "11112098.20", "PRC", "D")
    Call Check_HI_CE_accounting ("20200101",NewOverlimitRepay2.Isn, "01", fOBJECT(13), "11112098.20", "000", "11112098.20", "PEN", "C")
    
    Query = "Select fOBJECT From HI Where fOP = 'OVD' and fDBCR = 'C' and fBASE = "& NewOverlimitRepay2.Isn
    fOBJECT(14) = my_Row_Count(Query) 
    Call Check_HI_CE_accounting ("20200101",NewOverlimitRepay2.Isn, "01", fOBJECT(14), "77921477.90", "000", "77921477.90", "OVD", "C")

    'SQL Ստուգում HIR աղուսյակում 
    Call CheckQueryRowCount("HIR","fBASE",NewOverlimitRepay2.Isn,2)
    Call Check_HIR("20200101", "R1", fISN, "000", "77921477.90", "DBT", "C")
    Call Check_HIR("20200101", "R3", fISN, "000", "11112098.20", "DBT", "C")
    
    'SQL Ստուգում HIRREST  աղուսյակում 
    Call CheckQueryRowCount("HIRREST","fOBJECT",fISN,6)
    Call CheckDB_HIRREST("R1",fISN,"0.00","20200101",1)
    Call CheckDB_HIRREST("R3",fISN,"0.00","20200101",1)
    Call CheckDB_HIRREST("R4",fISN,"100000000.00","20200101",1)
    Call CheckDB_HIRREST("R5",fISN,"0.00","20200101",1)
    Call CheckDB_HIRREST("R7",fISN,"0.00","20200101",1)
    Call CheckDB_HIRREST("RI",fISN,"11111234.20","20200101",1)
    
    'SQL Ստուգում AGRSCHEDULE աղուսյակում 
    Call Check_AGRSCHEDULE(fISN, "20200101", "20200101", "3", "0", "2")
    
    'SQL Ստուգում AGRSCHEDULEVALUES աղուսյակում 
    Call Check_AGRSCHEDULEVALUES(fISN, "3", "20201202", "0.00", "1", "0")
   
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''-"Պայմանագրի փակում" գործողության կատարում-'''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''      
    Log.Message "Check RC (Close Contract Date/Պայմանագրի փակում) Function",,,DivideColor    
    
    Call CloseContract("010120")

    Log.Message "SQL Check After RC(Close Contract Date/Պայմանագրի փակում) Function",,,SqlDivideColor
    Call SQL_Init_Overlimit_Schedule(fISN,AccountParentIsn) 
    
    'SQL Ստուգում DOCS աղուսյակում  
    fBODY = "  CODE:00000113032  CRDTCODE:"&NewOverlimitDoc.GeneralTab.ExpectedCreditCode&"  CLICOD:00034852  NAME:ý²àôêî  CURRENCY:000  ACCACC:00000113032  COMMENT:MaxSinvolsSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSS1  AUTODEBT:1  ACCCONNMODE:3  USECLICONNSCH:1  ACCCONNSCH:002  DATE:20200101  DATEGIVE:20200101  ACSBRANCH:01  ACSDEPART:4  ACSTYPE:CO1  DATESFILLTYPE:2  AGRMARBEG:20200101  AGRMARFIN:20790101  AGRPERIOD:10/29  PASSOVDIRECTION:2  PASSOVTYPE:13  FIXEDROWSINSCH:0  SUMSDATESFILLTYPE:2  SUMSFILLTYPE:01  MIXEDSUMSINSCH:0  KINDSCALE:1  PCPENAGR:11.1111/29  CONSTPER:0  SECTOR:A  USAGEFIELD:03.006  AIM:04  INTERORG:BSTDB  SCHEDULE:9  GUARANTEE:1  COUNTRY:CH  LRDISTR:004  REGION:050000009  PERRES:1  NOTE:01  NOTE2:022  NOTE3:03  PPRCODE:88888888888888888881  DATECLOSE:20200101  SUBJRISK:1  CHRGFIRSTDAY:1  GIVEN:1  OPRISNLIST:"&RcOptionOverlimit.Isn&"  NTFMODE:1  SENDSTMADRS:2  PUTINLR:0  LRMRTCUR:000  ACRANOTE:02  NOTCLASS:1  REVISIONREASON:1  REPSOURCE:4  MORTSUBJECT:16  OTHERCOLLATERAL:1  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",fISN,1)
    Call CheckDB_DOCS(fISN,"COUniver","7",fBODY,1)
    
    'SQL Ստուգում DOCLOG աղուսյակում 
    Call CheckQueryRowCount("DOCLOG","fISN",fISN,9)
    Call CheckDB_DOCLOG(fISN,"77","N","1","",1)
    Call CheckDB_DOCLOG(fISN,"77","M","99","àõÕ³ñÏí»É ¿ Ñ³ëï³ïÙ³Ý",1)
    Call CheckDB_DOCLOG(fISN,"77","C","101","",1)
    Call CheckDB_DOCLOG(fISN,"77","W","102","",1)
    Call CheckDB_DOCLOG(fISN,"77","T","7","",1)
    Call CheckDB_DOCLOG(fISN,"77","E","7","",1)
    Call CheckDB_DOCLOG(fISN,"77","M","7","1 Ï³ñ·Ç ³ñ·»Éí³Í ·áñÍáÕáõÃÛáõÝ",1)
    Call CheckDB_DOCLOG(fISN,"77","M","77","ä³ÛÙ³Ý³·ñÇ ÷³ÏáõÙ",1)
    Call CheckDB_DOCLOG(fISN,"77","C","7","",1)
        
    'SQL Ստուգում DOCP աղուսյակում
    Call CheckQueryRowCount("DOCP","fISN",fISN,0)
    Call CheckQueryRowCount("DOCP","fPARENTISN",AccountParentIsn,0)
        
    'SQL Ստուգում CONTRACTS աղուսյակում 
    dbCONTRACT.fDGSTATE = "7"
    dbCONTRACT.fDGCRDTCODE = NewOverlimitDoc.GeneralTab.ExpectedCreditCode
    Call CheckQueryRowCount("CONTRACTS","fDGISN",fISN,1)
    Call CheckDB_CONTRACTS(dbCONTRACT,1)
    
    'SQL Ստուգում FOLDERS աղուսյակում 
    dbFOLDERS(1).fSPEC = "00000113032   CO1 1000"
    dbFOLDERS(3).fCOM = " ¶»ñ³Í³Ëë (·ñ³ýÇÏáí å³ÛÙ.)"
    dbFOLDERS(3).fSPEC = "00000113032 (ý²àôêî),     0 - Ð³ÛÏ³Ï³Ý ¹ñ³Ù[ö³Ïí³Í]"   
              
    Call CheckQueryRowCount("FOLDERS","fISN",fISN,5)
    Call CheckDB_FOLDERS(dbFOLDERS(1),1)
    Call CheckDB_FOLDERS(dbFOLDERS(2),1)
    Call CheckDB_FOLDERS(dbFOLDERS(3),1)
    Call CheckDB_FOLDERS(dbFOLDERS(4),1)
    Call CheckDB_FOLDERS(dbFOLDERS(6),1)    
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''-"Պայմանագրի Բացում" գործողության կատարում-'''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''  
    Log.Message "Check RC (Open Contract Date/Պայմանագրի Բացում) Function",,,DivideColor   
    
    Call OpenContract()
    wMDIClient.VBObject("frmPttel").Close
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''-Հեռացնել բոլոր գործողությունները-'''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''      

    Log.Message "Deleted All Action",,,DivideColor 
 
    Call wTreeView.DblClickItem("|¶»ñ³Í³Ëë|¶áñÍáÕáõÃÛáõÝÝ»ñ")
    wMDIClient.Refresh
    
	  BuiltIn.Delay(delay_small)
    Call Rekvizit_Fill("Dialog", 1, "General", "START", "^A[Del]") 
    Call Rekvizit_Fill("Dialog", 1, "General", "END", "^A[Del]") 
    
    BuiltIn.Delay(delay_small)
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    BuiltIn.Delay(delay_small)
    Set DocForm = wMDIClient.VBObject("frmPttel") 
        
    If WaitForPttel("frmPttel") Then
        Call SearchAndDelete("frmPttel", 4, "77921477.9", "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ")
        Call SearchAndDelete("frmPttel", 5, "¶»ñ³Í³ËëÇ í»ñ³Ï³Ý·ÝáõÙ", "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ") 
        Call SearchAndDelete("frmPttel", 5, "¶»ñ³Í³ËëÇ ¹áõñë·ñáõÙ", "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ")    
        Call SearchAndDelete("frmPttel", 4, "987.1", "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ") 
        Call SearchAndDelete("frmPttel", 5, "ä³Ñáõëï³íáñáõÙ", "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ") 
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
    BuiltIn.Delay(delay_small)
    Call Rekvizit_Fill("Dialog", 1, "General", "START", "^A[Del]" & "010120") 
    Call Rekvizit_Fill("Dialog", 1, "General", "END", "^A[Del]") 
    Call Rekvizit_Fill("Dialog", 1, "General", "NUM", "00000113032")
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    BuiltIn.Delay(delay_small)
    
    Set DocForm = wMDIClient.VBObject("frmPttel")
    If WaitForPttel("frmPttel") Then
        Call SearchAndDelete("frmPttel", 3, "èÇëÏÇ ¹³ëÇã", "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ") 
        Call SearchAndDelete("frmPttel", 3, "úµÛ»ÏïÇí éÇëÏÇ ¹³ëÇã", "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ") 
        Call SearchAndDelete("frmPttel", 4, "îáÏáë³¹ñáõÛùÝ»ñ", "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ") 
        Call SearchAndDelete("frmPttel", 3, "Å³ÙÏ»ï³Ýó ·áõÙ³ñÇ ïáõÛÅÇ Ù³ñáõÙ", "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ") 
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
    fBODY = "  CODE:00000113032  CRDTCODE:"&NewOverlimitDoc.GeneralTab.ExpectedCreditCode&"  CLICOD:00034852  NAME:ý²àôêî  CURRENCY:000  ACCACC:00000113032  COMMENT:MaxSinvolsSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSS1  AUTODEBT:1  ACCCONNMODE:3  USECLICONNSCH:1  ACCCONNSCH:002  DATE:20200101  DATEGIVE:20200101  ACSBRANCH:01  ACSDEPART:4  ACSTYPE:CO1  DATESFILLTYPE:2  AGRMARBEG:20200101  AGRMARFIN:20790101  AGRPERIOD:10/29  PASSOVDIRECTION:2  PASSOVTYPE:13  FIXEDROWSINSCH:0  SUMSDATESFILLTYPE:2  SUMSFILLTYPE:01  MIXEDSUMSINSCH:0  KINDSCALE:1  PCPENAGR:11.1111/29  CONSTPER:0  SECTOR:A  USAGEFIELD:03.006  AIM:04  INTERORG:BSTDB  SCHEDULE:9  GUARANTEE:1  COUNTRY:CH  LRDISTR:004  REGION:050000009  PERRES:1  NOTE:01  NOTE2:022  NOTE3:03  PPRCODE:88888888888888888881  SUBJRISK:1  CHRGFIRSTDAY:1  GIVEN:1  OPRISNLIST:"&RcOptionOverlimit.Isn&"  NTFMODE:1  SENDSTMADRS:2  PUTINLR:0  LRMRTCUR:000  ACRANOTE:02  NOTCLASS:1  REVISIONREASON:1  REPSOURCE:4  MORTSUBJECT:16  OTHERCOLLATERAL:1  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckDB_DOCS(fISN,"COUniver","999",fBODY,1)
    
    fBODY = "  CODE:00000113032  DATE:20200101  SUMMA:100121500  CASHORNO:2  ACCCORR:00000113032  COMMENT:RcOptionOverlimitRcOptionOverlimitRcOptionOverlimitRcOptionOverlimitRcOptionOverlimitRcOptionOverli1  MARDOCISN:"&fBASE(0)&"  ACSBRANCH:01  ACSDEPART:4  ACSTYPE:CO1  USERID:  77  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckDB_DOCS(RcOptionOverlimit.Isn,"CODSAgr ","999",fBODY,1)
    
    fBODY = "  CODE:00000113032  DATE:20200101  PCPENAGR:1092.1001/29  COMMENT:ChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChan1  USERID:  77  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckDB_DOCS(ActionIsn(1),"COTSPC  ","999",fBODY,1)
    
    fBODY = "  CODE:00000113032  DATE:20200101  RISK:02  COMMENT:ChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChan1  USERID:  77  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckDB_DOCS(ActionIsn(2),"COTSORC ","999",fBODY,1)
    
    fBODY = "  CODE:00000113032  DATE:20200101  RISK:01  PERRES:2  COMMENT:ChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChangeOverlimitReteChan1  USERID:  77  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckDB_DOCS(ActionIsn(3),"COTSRsPr","999",fBODY,1)
    
    fBODY = "  CODE:00000113032  DATE:20200101  SUMRES:100000000  COMMENT:NewStoreNewStoreNewStoreNewStoreNewStoreNewStoreNewStoreNewStoreNewStoreNewStoreNewStoreNewStoreNew1  ACSBRANCH:01  ACSDEPART:3  ACSTYPE:CO1  USERID:  77  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckDB_DOCS(NewStore.Isn,"CODSRes ","999",fBODY,1)
    
    fBODY = "  CODE:00000113032  DATECHARGE:20200101  DATE:20200101  SUMAGRPEN:987.1/123.1  SUMALLPEN:987.1/123.1  ACSBRANCH:00  ACSDEPART:4  ACSTYPE:CO1  USERID:  77  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckDB_DOCS(CalcPercents_2.Isn,"CODSChrg","999",fBODY,1)
    
    fBODY = "  CODE:00000113032  DATE:20200101  SUMAGR:100  SUMFINE:764  SUMMA:864  COMMENT:NewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOutNewWriteOut1  ACSBRANCH:01  ACSDEPART:4  ACSTYPE:CO1  USERID:  77  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckDB_DOCS(NewWriteOut.Isn,"CODSOut ","999",fBODY,1)
    
    fBODY = "  CODE:00000113032  DATE:20200101  SUMAGR:100  SUMFINE:11111998.2  SUMMA:11112098.2  COMMENT:wWriteOffwWriteOffwWriteOffwWriteOffwWriteOffwWriteOffwWriteOffwWriteOffwWriteOffwWriteOffwWriteOff1  ACSBRANCH:01  ACSDEPART:3  ACSTYPE:CO1  USERID:  77  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckDB_DOCS(NewWriteOff.Isn,"CODSInc ","999",fBODY,1)
    
    fBODY = "  CODE:00000113032  DATE:20200101  SUMAGR:77921477.9  SUMFINE:11112098.2  SUMMA:89033576.1  CASHORNO:2  ISPUSA:0  ACCCORR:00000113032  COMMENT:NewOverlimitRepayNewOverlimitRepayNewOverlimitRepayNewOverlimitRepayNewOverlimitRepayNewOverlimitRe1  MARDOCISN:"&fBASE(2)&"  ACSBRANCH:01  ACSDEPART:4  ACSTYPE:CO1  USERID:  77  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckDB_DOCS(NewOverlimitRepay2.Isn,"CODSDebt","999",fBODY,1)
    
    'SQL Ստուգում DOCLOG աղուսյակում 
    Call CheckQueryRowCount("DOCLOG","fISN",fISN,11)
    Call CheckDB_DOCLOG(fISN,"77","M","7","ä³ÛÙ³Ý³·ñÇ µ³óáõÙ",1)
    Call CheckDB_DOCLOG(fISN,"77","D","999","",1)
    
    Set dbFOLDERS(13) = New_DB_FOLDERS()
        dbFOLDERS(13).fFOLDERID = ".R." & aqConvert.DateTimeToFormatStr(aqDateTime.Now(),"%Y%m%d")
        dbFOLDERS(13).fNAME = "COUniver"
        dbFOLDERS(13).fKEY = fISN
        dbFOLDERS(13).fISN = fISN
        dbFOLDERS(13).fSTATUS = "0"
        dbFOLDERS(13).fSPEC = Left_Align(Get_Compname_DOCLOG(fISN), 16)&"CredO   ARMSOFT                       007  "
        dbFOLDERS(13).fDCBRANCH = "01"
        dbFOLDERS(13).fDCDEPART = "4"
    Call CheckDB_FOLDERS(dbFOLDERS(13),1)
    
        dbFOLDERS(13).fNAME = "CODSAgr "
        dbFOLDERS(13).fKEY = RcOptionOverlimit.Isn
        dbFOLDERS(13).fISN = RcOptionOverlimit.Isn
        dbFOLDERS(13).fSPEC = Left_Align(Get_Compname_DOCLOG(fISN), 16)&"CredO   ARMSOFT                       115  "
        BuiltIn.Delay(delay_small)
    Call CheckDB_FOLDERS(dbFOLDERS(13),1)
    
        dbFOLDERS(13).fNAME = "CODSChrg"
        dbFOLDERS(13).fKEY = CalcPercents.Isn
        dbFOLDERS(13).fISN = CalcPercents.Isn
        dbFOLDERS(13).fSPEC = Left_Align(Get_Compname_DOCLOG(fISN), 16)&"CredO   ARMSOFT                       005  "
    Call CheckDB_FOLDERS(dbFOLDERS(13),1)
    
        dbFOLDERS(13).fNAME = "CODSDebt"
        dbFOLDERS(13).fKEY = NewOverlimitRepay.Isn
        dbFOLDERS(13).fISN = NewOverlimitRepay.Isn
        dbFOLDERS(13).fSPEC = Left_Align(Get_Compname_DOCLOG(fISN), 16)&"CredO   ARMSOFT                       115  "
    Call CheckDB_FOLDERS(dbFOLDERS(13),1)
    
        dbFOLDERS(13).fNAME = "COTSPC  "
        dbFOLDERS(13).fKEY = ActionIsn(1)
        dbFOLDERS(13).fISN = ActionIsn(1)
        dbFOLDERS(13).fSPEC = Left_Align(Get_Compname_DOCLOG(fISN), 16)&"CredO   ARMSOFT                       005  "
        dbFOLDERS(13).fDCBRANCH = ""
        dbFOLDERS(13).fDCDEPART = ""
    Call CheckDB_FOLDERS(dbFOLDERS(13),1)
    
        dbFOLDERS(13).fNAME = "COTSORC "
        dbFOLDERS(13).fKEY = ActionIsn(2)
        dbFOLDERS(13).fISN = ActionIsn(2)
        dbFOLDERS(13).fSPEC = Left_Align(Get_Compname_DOCLOG(fISN), 16)&"CredO   ARMSOFT                       005  "
    Call CheckDB_FOLDERS(dbFOLDERS(13),1)
    
        dbFOLDERS(13).fNAME = "COTSRsPr"
        dbFOLDERS(13).fKEY = ActionIsn(3)
        dbFOLDERS(13).fISN = ActionIsn(3)
        dbFOLDERS(13).fSPEC = Left_Align(Get_Compname_DOCLOG(fISN), 16)&"CredO   ARMSOFT                       005  "
    Call CheckDB_FOLDERS(dbFOLDERS(13),1)
    
        dbFOLDERS(13).fNAME = "CODSRes "
        dbFOLDERS(13).fKEY = NewStore.Isn
        dbFOLDERS(13).fISN = NewStore.Isn
        dbFOLDERS(13).fSPEC = Left_Align(Get_Compname_DOCLOG(fISN), 16)&"CredO   ARMSOFT                       115  "
        dbFOLDERS(13).fDCBRANCH = "01"
        dbFOLDERS(13).fDCDEPART = "3"
    Call CheckDB_FOLDERS(dbFOLDERS(13),1)
    
        dbFOLDERS(13).fNAME = "CODSChrg"
        dbFOLDERS(13).fKEY = CalcPercents_2.Isn
        dbFOLDERS(13).fISN = CalcPercents_2.Isn
        dbFOLDERS(13).fSPEC = Left_Align(Get_Compname_DOCLOG(fISN), 16)&"CredO   ARMSOFT                       005  "
        dbFOLDERS(13).fDCBRANCH = "00"
        dbFOLDERS(13).fDCDEPART = "4"
    Call CheckDB_FOLDERS(dbFOLDERS(13),1)
    
        dbFOLDERS(13).fNAME = "CODSOut "
        dbFOLDERS(13).fKEY = NewWriteOut.Isn
        dbFOLDERS(13).fISN = NewWriteOut.Isn
        dbFOLDERS(13).fSPEC = Left_Align(Get_Compname_DOCLOG(fISN), 16)&"CredO   ARMSOFT                       115  "
        dbFOLDERS(13).fDCBRANCH = "01"
        dbFOLDERS(13).fDCDEPART = "4"
    Call CheckDB_FOLDERS(dbFOLDERS(13),1)
        
        dbFOLDERS(13).fNAME = "CODSInc "
        dbFOLDERS(13).fKEY = NewWriteOff.Isn
        dbFOLDERS(13).fISN = NewWriteOff.Isn
        dbFOLDERS(13).fSPEC = Left_Align(Get_Compname_DOCLOG(fISN), 16)&"CredO   ARMSOFT                       115  "
        dbFOLDERS(13).fDCBRANCH = "01"
        dbFOLDERS(13).fDCDEPART = "3"
    Call CheckDB_FOLDERS(dbFOLDERS(13),1)
    
        dbFOLDERS(13).fNAME = "CODSDebt"
        dbFOLDERS(13).fKEY = NewOverlimitRepay2.Isn
        dbFOLDERS(13).fISN = NewOverlimitRepay2.Isn
        dbFOLDERS(13).fSPEC = Left_Align(Get_Compname_DOCLOG(fISN), 16)&"CredO   ARMSOFT                       115  "
        dbFOLDERS(13).fDCBRANCH = "01"
        dbFOLDERS(13).fDCDEPART = "4"
    Call CheckDB_FOLDERS(dbFOLDERS(13),1)
    
        dbFOLDERS(13).fNAME = "COTSDtUn"
        dbFOLDERS(13).fKEY = fBASE(0)
        dbFOLDERS(13).fISN = fBASE(0)
        dbFOLDERS(13).fSPEC = Left_Align(Get_Compname_DOCLOG(fISN), 16)&"CredO   ARMSOFT                       005  "
        dbFOLDERS(13).fDCBRANCH = ""
        dbFOLDERS(13).fDCDEPART = ""
    Call CheckDB_FOLDERS(dbFOLDERS(13),1)
    
        dbFOLDERS(13).fNAME = "COTSDtUn"
        dbFOLDERS(13).fKEY = fBASE(1)
        dbFOLDERS(13).fISN = fBASE(1)
        dbFOLDERS(13).fSPEC = Left_Align(Get_Compname_DOCLOG(fISN), 16)&"CredO   ARMSOFT                       005  "
    Call CheckDB_FOLDERS(dbFOLDERS(13),1)
    
        dbFOLDERS(13).fNAME = "COTSDtUn"
        dbFOLDERS(13).fKEY = fBASE(2)
        dbFOLDERS(13).fISN = fBASE(2)
        dbFOLDERS(13).fSPEC = Left_Align(Get_Compname_DOCLOG(fISN), 16)&"CredO   ARMSOFT                       005  "
    Call CheckDB_FOLDERS(dbFOLDERS(13),1)
    
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
    Call CheckQueryRowCount("HIREST","fOBJECT",fOBJECT(8),0)
    Call CheckQueryRowCount("HIREST","fOBJECT",fOBJECT(9),0)
    Call CheckQueryRowCount("HIREST","fOBJECT",fOBJECT(10),0)
    Call CheckQueryRowCount("HIREST","fOBJECT",fOBJECT(11),0)
    Call CheckQueryRowCount("HIREST","fOBJECT",fOBJECT(12),0)
    Call CheckQueryRowCount("HIREST","fOBJECT",fOBJECT(13),0)
    Call CheckQueryRowCount("HIREST","fOBJECT",fOBJECT(14),0)

    Call Close_AsBank()
End Sub 

Sub Test_Overlimit_Schedule_Init()

    Set NewOverlimitDoc = New_OverlimitDoc()
        NewOverlimitDoc.DocType = "¶»ñ³Í³Ëë (·ñ³ýÇÏáí å³ÛÙ.)"

        'For Interests Tab (Ընդանուր)
        NewOverlimitDoc.GeneralTab.AgreementN = "00000113032"
        NewOverlimitDoc.GeneralTab.ExpectedClient = "00034852"
        NewOverlimitDoc.GeneralTab.SettlementAccount = "00000113032"
        NewOverlimitDoc.GeneralTab.ExpectedName = "ý²àôêî"
        NewOverlimitDoc.GeneralTab.ExpectedCurrency = "000"
        NewOverlimitDoc.GeneralTab.Comment = "MaxSinvolsSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSS1"
        NewOverlimitDoc.GeneralTab.GenerateButton = True
        NewOverlimitDoc.GeneralTab.ExpectedCreditCode = ""
        NewOverlimitDoc.GeneralTab.ExpectedUseOtherAccountRemainders = "3"
        NewOverlimitDoc.GeneralTab.UseClientsSchema = "1"
        NewOverlimitDoc.GeneralTab.AccountConnectionSchema = "002"
        NewOverlimitDoc.GeneralTab.SigningDate = "010120"
        NewOverlimitDoc.GeneralTab.DisbursementDate = "010120"
        NewOverlimitDoc.GeneralTab.Division = "01"
        NewOverlimitDoc.GeneralTab.Department = "4"
        NewOverlimitDoc.GeneralTab.AccessType = "CO1"

        'For ScheduleOrganizing Tab (Գրաֆիկի լրացման ձև)
        NewOverlimitDoc.ScheduleOrganizingTab.FillTab = True
        NewOverlimitDoc.ScheduleOrganizingTab.ScheduleDateOrgMode = "2"
        NewOverlimitDoc.ScheduleOrganizingTab.RepaymentDays = ""
        NewOverlimitDoc.ScheduleOrganizingTab.PerioducityMonts = "10"
        NewOverlimitDoc.ScheduleOrganizingTab.PerioducityDays = "29"
        NewOverlimitDoc.ScheduleOrganizingTab.DayOverpassingMethod = "2"
        NewOverlimitDoc.ScheduleOrganizingTab.OvrDays = "13"
        
        'For Interests Tab (Տոկոսներ)
        NewOverlimitDoc.InterestsTab.FillTab = True
        NewOverlimitDoc.InterestsTab.ExpectedKindOfScale = "1"
        NewOverlimitDoc.InterestsTab.KindOfScale = ""
        NewOverlimitDoc.InterestsTab.FineOnPastDueSum = "11.1111" 
        NewOverlimitDoc.InterestsTab.Div = "29"
        
        'For Additional Tab (Լրացուցիչ)
        NewOverlimitDoc.AdditionalTab.FillTab = True
        NewOverlimitDoc.AdditionalTab.Sector = "A"
        NewOverlimitDoc.AdditionalTab.UsageField = "03.006"
        NewOverlimitDoc.AdditionalTab.Aim = "04"
        NewOverlimitDoc.AdditionalTab.InternationalOrganization = "BSTDB"
        NewOverlimitDoc.AdditionalTab.ProjectName = "9"
        NewOverlimitDoc.AdditionalTab.Guarantee = "1"
        NewOverlimitDoc.AdditionalTab.Country = "CH"
        NewOverlimitDoc.AdditionalTab.Region = "004"
        NewOverlimitDoc.AdditionalTab.RegionNewLR = "050000009"
        NewOverlimitDoc.AdditionalTab.Note = "01"
        NewOverlimitDoc.AdditionalTab.Note2 = "022"
        NewOverlimitDoc.AdditionalTab.Note3 = "03"
        NewOverlimitDoc.AdditionalTab.AgreemPaperN = "88888888888888888881"
        NewOverlimitDoc.AdditionalTab.ClosingDate = ""
        NewOverlimitDoc.AdditionalTab.FullyClosed = ""
        NewOverlimitDoc.AdditionalTab.SubjectiveCategorized = "1"
        
        'For Notification Tab (Ծանուցում)       
        NewOverlimitDoc.NotificationTab.FillTab = True
        NewOverlimitDoc.NotificationTab.NotifyMode = "1"
        NewOverlimitDoc.NotificationTab.SendNotificationAddress = "2"
        
        'For LoanReg Tab (Վարկ.ռեգ)   
        NewOverlimitDoc.LoanRegTab.FillTab = True
        NewOverlimitDoc.LoanRegTab.AccumulateInLoanReg = "0"
        NewOverlimitDoc.LoanRegTab.PledgeCurrency = "000"
        NewOverlimitDoc.LoanRegTab.PladgeObjectArca = "02"
        NewOverlimitDoc.LoanRegTab.NotClassifiable = "1"
        NewOverlimitDoc.LoanRegTab.LRCodeNew = ""
        NewOverlimitDoc.LoanRegTab.RevisionReason = "1"
        NewOverlimitDoc.LoanRegTab.RepaymentSourse = "4"
        NewOverlimitDoc.LoanRegTab.PledgeObjectNew = "16"
        NewOverlimitDoc.LoanRegTab.GuaranteedByOtherCallateral = "1"
        
        '"Հաստատող փաստաթնթեր 1" Ֆիլտրի լրացման արժեքներ
    Set VerifyOverlimit1 = New_VerifyContract()
        VerifyOverlimit1.AgreementN = "00000113032"
        VerifyOverlimit1.Client = "00034852"
        VerifyOverlimit1.Curr = "000"
        VerifyOverlimit1.Note = "01"
        VerifyOverlimit1.Note2 = "022"
        VerifyOverlimit1.Note3 = "03"
        VerifyOverlimit1.Executors = ""
        VerifyOverlimit1.Division = "01"
        VerifyOverlimit1.Department = "4"
        VerifyOverlimit1.AccessType = "CO1"
    
        '"Պայմանագրեր" Ֆիլտրի լրացման արժեքներ
    Set ContractNew = New_ContractOverlimit()
        ContractNew.AgreementLevel = "1"
        ContractNew.AgreementSpecies = "8"
        ContractNew.AgreementN = "00000113032"
        ContractNew.Curr = "000"
        ContractNew.Client = "00034852"
        ContractNew.ShowClosed = "1"
        ContractNew.Note = "01"
        ContractNew.Note2 = "022"
        ContractNew.Note3 = "03"
        ContractNew.Division = "01"
        ContractNew.Department = "4"
        ContractNew.AccessType = "CO1"
        
        '"Աջ կլիկ/Գերածախս" Պատուհանի լրացման արժեքներ
    Set RcOptionOverlimit = New_RcOverlimit()
        RcOptionOverlimit.ExpectedAgreementN = "00000113032"
        RcOptionOverlimit.Date = "010120"
        RcOptionOverlimit.Sum = "100121500.01"
        RcOptionOverlimit.CashOrNo = "2"
        RcOptionOverlimit.CalcAcc = "00000113032"
        RcOptionOverlimit.Comment = "RcOptionOverlimitRcOptionOverlimitRcOptionOverlimitRcOptionOverlimitRcOptionOverlimitRcOptionOverli1"
        RcOptionOverlimit.Division = "01"
        RcOptionOverlimit.Department = "4"
        
        '"Աջ կլիկ/Տոկոսների հաշվարկ" Պատուհանի լրացման արժեքներ
    Set CalcPercents = New_RcCalculatePercents()
        CalcPercents.ExpectedAgreementN = "00000113032"
        CalcPercents.CalculationDate = "010120"
        CalcPercents.OperationDate = "010120" 
        CalcPercents.FineOnPastDueSum = "22222222.09"
        CalcPercents.FineOnPastDueSum2 = "11111111.09"
        CalcPercents.TotalPenalty = "22222222.10"
        CalcPercents.TotalPenalty2 = "11111111.10"
        CalcPercents.Comment =  "CalcPercentsCalcPercentsCalcPercentsCalcPercentsCalcPercentsCalcPercentsCalcPercentsCalcPercentsCal1"
        CalcPercents.Division =  "01"
        CalcPercents.Department =  "4"
    
    Set CalcPercents_2 = New_RcCalculatePercents()
        CalcPercents_2.ExpectedAgreementN = "00000113032"
        CalcPercents_2.CalculationDate = "010120"
        CalcPercents_2.OperationDate = "010120" 
        CalcPercents_2.FineOnPastDueSum = "987.09"
        CalcPercents_2.FineOnPastDueSum2 = "123.09"
        CalcPercents_2.TotalPenalty = "987.10"
        CalcPercents_2.TotalPenalty2 = "123.10"
        CalcPercents_2.Division =  "00"
        CalcPercents_2.Department =  "4" 
        
        '"Աջ կլիկ/Պարտքերի մարում" Պատուհանի լրացման արժեքներ
    Set NewOverlimitRepay = New_RcOverlimitRepay()   
        NewOverlimitRepay.ExpectedAgreementN = "00000113032"
        NewOverlimitRepay.RepaymentCurrency = ""
        NewOverlimitRepay.ExpectedBaseSum = "0.00"
        NewOverlimitRepay.BaseSum = "22200022.11"
        NewOverlimitRepay.AMD1 = "0.00"
        NewOverlimitRepay.ExpectedFineOnPastSum = "11,111,111.00"
        NewOverlimitRepay.FineOnPastSum = "11111111.00"
        NewOverlimitRepay.AMD2 = "0.00"
        NewOverlimitRepay.TotalAmount = "33,311,133.10"
        NewOverlimitRepay.CashCashles = "2"
        NewOverlimitRepay.ExchangeRate = "0"
        NewOverlimitRepay.Per = "0"
        NewOverlimitRepay.Account = "00000113032"
        NewOverlimitRepay.AccountComment = ""
        NewOverlimitRepay.AMDAccount = ""
        NewOverlimitRepay.Comment =  "NewOverlimitRepayNewOverlimitRepayNewOverlimitRepayNewOverlimitRepayNewOverlimitRepayNewOverlimitRe1"
        NewOverlimitRepay.RemittanceInfo1 = "Ø.·.`  100,121,500.00;" 
        NewOverlimitRepay.RemittanceInfo2 = "Ø.·.ïÅ.`  11,111,111.00"
        NewOverlimitRepay.Division = "01"
        NewOverlimitRepay.Department = "4"    
        
    Set NewOverlimitRepay2 = New_RcOverlimitRepay()   
        NewOverlimitRepay2.ExpectedAgreementN = "00000113032"
        NewOverlimitRepay2.Date = "010120"
        NewOverlimitRepay2.RepaymentCurrency = ""
        NewOverlimitRepay2.ExpectedBaseSum = "0.00"
        NewOverlimitRepay2.BaseSum = "77921477.9"
        NewOverlimitRepay2.AMD1 = "0.00"
        NewOverlimitRepay2.ExpectedFineOnPastSum = "11,112,098.20"
        NewOverlimitRepay2.FineOnPastSum = "11112098.2"
        NewOverlimitRepay2.AMD2 = "0.00"
        NewOverlimitRepay2.TotalAmount = "89,033,576.10"
        NewOverlimitRepay2.CashCashles = "2"
        NewOverlimitRepay2.ExchangeRate = "0"
        NewOverlimitRepay2.Per = "0"
        NewOverlimitRepay2.Account = "00000113032"
        NewOverlimitRepay2.AccountComment = ""
        NewOverlimitRepay2.AMDAccount = ""
        NewOverlimitRepay2.Comment =  "NewOverlimitRepayNewOverlimitRepayNewOverlimitRepayNewOverlimitRepayNewOverlimitRepayNewOverlimitRe1"
        NewOverlimitRepay2.RemittanceInfo1 = "Ø.·.`  77,921,477.90" 
        NewOverlimitRepay2.RemittanceInfo2 = "Ø.·.ïÅ.`  11,112,098.20"
        NewOverlimitRepay2.Division = "01"
        NewOverlimitRepay2.Department = "4"   
           
        '"Աջ կլիկ/Պահուստավորում" Պատուհանի լրացման արժեքներ
    Set NewStore = New_RcStore()    
        NewStore.ExpectedAgreementN = "00000113032"
        NewStore.Date = "010120"
        NewStore.Provision = "99999999.99"
        NewStore.UnProvision = ""
        NewStore.Comment =  "NewStoreNewStoreNewStoreNewStoreNewStoreNewStoreNewStoreNewStoreNewStoreNewStoreNewStoreNewStoreNew1"
        NewStore.Division = "01"
        NewStore.Department = "3"
        
        '"Աջ կլիկ/Դուրս գրում" Պատուհանի լրացման արժեքներ
    Set NewWriteOut = New_RcWriteOut()
        NewWriteOut.ExpectedAgreementN = "00000113032"
        NewWriteOut.Date = "010120"
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
        NewWriteOff.ExpectedAgreementN = "00000113032"
        NewWriteOff.Date = "010120"
        NewWriteOff.ExpectedBaseSum = "100.00"
        NewWriteOff.BaseSum = "100.00"
        NewWriteOff.ExpectedFineOnPastSum = "11,111,998.20"
        NewWriteOff.FineOnPastSum = "11111998.20"
        NewWriteOff.TotalSum = "11,112,098.20"
        NewWriteOff.Comment = "wWriteOffwWriteOffwWriteOffwWriteOffwWriteOffwWriteOffwWriteOffwWriteOffwWriteOffwWriteOffwWriteOff1"
        NewWriteOff.Division = "01"
        NewWriteOff.Department = "3"
End Sub   



Sub SQL_Init_Overlimit_Schedule(fISN,fISN2)
       
    Set dbFOLDERS(1) = New_DB_FOLDERS()
        dbFOLDERS(1).fFOLDERID = "AGROVERLIM"
        dbFOLDERS(1).fNAME = "COUniver"
        dbFOLDERS(1).fKEY = "00000113032"
        dbFOLDERS(1).fISN = fISN
        dbFOLDERS(1).fSTATUS = "1"
        dbFOLDERS(1).fCOM = "ý²àôêî"
        dbFOLDERS(1).fSPEC = "00000113032   CO1 0000"
        dbFOLDERS(1).fECOM = ""
        dbFOLDERS(1).fDCBRANCH = "01"
        dbFOLDERS(1).fDCDEPART = "4"

    Set dbFOLDERS(2) = New_DB_FOLDERS()
        dbFOLDERS(2).fFOLDERID = "Agr."&fISN
        dbFOLDERS(2).fNAME = "COUniver"
        dbFOLDERS(2).fKEY = fISN
        dbFOLDERS(2).fISN = fISN
        dbFOLDERS(2).fSTATUS = "1"
        dbFOLDERS(2).fCOM = "¶»ñ³Í³Ëë (·ñ³ýÇÏáí å³ÛÙ.)"
        dbFOLDERS(2).fSPEC = "1¶»ñ³Í³Ëë (·ñ³ýÇÏáí å³ÛÙ.)- 00000113032 {ý²àôêî}"
     
    Set dbFOLDERS(3) = New_DB_FOLDERS()
        dbFOLDERS(3).fFOLDERID = "C.903824400"
        dbFOLDERS(3).fNAME = "COUniver"
        dbFOLDERS(3).fKEY = fISN
        dbFOLDERS(3).fISN = fISN
        dbFOLDERS(3).fSTATUS = "1"
        dbFOLDERS(3).fCOM = " ¶»ñ³Í³Ëë (·ñ³ýÇÏáí å³ÛÙ.) (Ý³Ë³·ÇÍ)"
        dbFOLDERS(3).fSPEC = "00000113032 (ý²àôêî),     0 - Ð³ÛÏ³Ï³Ý ¹ñ³Ù"   
                
    Set dbFOLDERS(4) = New_DB_FOLDERS()
        dbFOLDERS(4).fFOLDERID = "NOTCLASSIFIABLE"
        dbFOLDERS(4).fNAME = "COUniver"
        dbFOLDERS(4).fKEY = fISN
        dbFOLDERS(4).fISN = fISN
        dbFOLDERS(4).fSTATUS = "0"
        dbFOLDERS(4).fCOM = "¶»ñ³Í³Ëë (·ñ³ýÇÏáí å³ÛÙ.)"
    
    Set dbFOLDERS(5) = New_DB_FOLDERS()
        dbFOLDERS(5).fFOLDERID = "SSWork.CRCO20200101"
        dbFOLDERS(5).fNAME = "COUniver"
        dbFOLDERS(5).fKEY = fISN
        dbFOLDERS(5).fISN = fISN
        dbFOLDERS(5).fSTATUS = "1"
        dbFOLDERS(5).fCOM = "¶»ñ³Í³Ëë (·ñ³ýÇÏáí å³ÛÙ.)"
        dbFOLDERS(5).fSPEC = "00000113032   CO1 20200101            0.0077  00034852Üáñ å³ÛÙ³Ý³·Çñ      "  
        dbFOLDERS(5).fECOM = "Overlimit (univer.)"
        dbFOLDERS(5).fDCBRANCH = "01"
        dbFOLDERS(5).fDCDEPART = "4"
    
    Set dbFOLDERS(6) = New_DB_FOLDERS()
        dbFOLDERS(6).fFOLDERID = "SUBJRISK"
        dbFOLDERS(6).fNAME = "COUniver"
        dbFOLDERS(6).fKEY = fISN
        dbFOLDERS(6).fISN = fISN
        dbFOLDERS(6).fSTATUS = "0"
        dbFOLDERS(6).fCOM = "¶»ñ³Í³Ëë (·ñ³ýÇÏáí å³ÛÙ.)"
        dbFOLDERS(6).fSPEC = "COUniver"  
               
    Set dbCONTRACT = New_DB_CONTRACTS()
        dbCONTRACT.fDGISN = fISN
        dbCONTRACT.fDGPARENTISN = fISN
        dbCONTRACT.fDGISN1 = fISN
        dbCONTRACT.fDGISN3 = fISN
        dbCONTRACT.fDGAGRKIND = "8 "
        dbCONTRACT.fDGSTATE = "1"
        dbCONTRACT.fDGTYPENAME = "COUniver"
        dbCONTRACT.fDGCODE = "00000113032   "
        dbCONTRACT.fDGPPRCODE = "88888888888888888881"
        dbCONTRACT.fDGCAPTION = "ý²àôêî"
        dbCONTRACT.fDGCLICODE = "00034852"
        dbCONTRACT.fDGCUR = "000"
        dbCONTRACT.fDGSUMMA = "0.00"
        dbCONTRACT.fDGALLSUMMA = "0.00"
        dbCONTRACT.fDGRISKDEGREE = "0.00"
        dbCONTRACT.fDGRISKDEGNB = "0.00"
        dbCONTRACT.fDGSCHEDULE = "9"
        dbCONTRACT.fDGNOTE = "01"
        dbCONTRACT.fDGNOTE2 = "022"
        dbCONTRACT.fDGNOTE3 = "03"
        dbCONTRACT.fDGDISTRICT = "004"
        dbCONTRACT.fDGACSBRANCH = "01"
        dbCONTRACT.fDGACSDEPART = "4"
        dbCONTRACT.fDGACSTYPE = "CO1 "
        dbCONTRACT.fDGACRANOTE = "02"
        dbCONTRACT.fDGCRDTCODE = ""
        dbCONTRACT.fDGAIM = "04"
        dbCONTRACT.fDGUSAGEFIELD = "03.006"
        dbCONTRACT.fDGCOUNTRY = "CH "
        dbCONTRACT.fDGREGION = "050000009"
        dbCONTRACT.fDGREVISIONREASON = "1"
        dbCONTRACT.fDGREPSOURCE = "4"
        dbCONTRACT.fDGMORTSUBJECT = "16"
        
    Set dbFOLDERS(7) = New_DB_FOLDERS()
        dbFOLDERS(7).fFOLDERID = "SSConf.CRCO001"
        dbFOLDERS(7).fNAME = "COUniver"
        dbFOLDERS(7).fKEY = fISN
        dbFOLDERS(7).fISN = fISN
        dbFOLDERS(7).fSTATUS = "4"
        dbFOLDERS(7).fCOM = "¶»ñ³Í³Ëë (·ñ³ýÇÏáí å³ÛÙ.)"
        dbFOLDERS(7).fSPEC = "00000113032   CO1 20200101            0.0077  00034852"
        dbFOLDERS(7).fECOM = "Overlimit (univer.)"
        dbFOLDERS(7).fDCBRANCH = "01"
        dbFOLDERS(7).fDCDEPART = "4"
        
    Set dbFOLDERS(8) = New_DB_FOLDERS()
        dbFOLDERS(8).fFOLDERID = "ALLACCSACC"
        dbFOLDERS(8).fNAME = "COAgrAcc"
        dbFOLDERS(8).fKEY = fISN
        dbFOLDERS(8).fISN = fISN2
        dbFOLDERS(8).fSTATUS = "1"
        dbFOLDERS(8).fCOM = "¶»ñ³Í³ËëÇ Ñ³ßí³å³Ñ³Ï³Ý Ñ³í»Éí³Í"
        dbFOLDERS(8).fSPEC = "000000113032                                            1"
    
    Set dbFOLDERS(9) = New_DB_FOLDERS()
        dbFOLDERS(9).fFOLDERID = "ALLACCSGEN"
        dbFOLDERS(9).fNAME = "COAgrAcc"
        dbFOLDERS(9).fKEY = fISN
        dbFOLDERS(9).fISN = fISN2
        dbFOLDERS(9).fSTATUS = "1"
        dbFOLDERS(9).fCOM = "¶»ñ³Í³ËëÇ Ñ³ßí³å³Ñ³Ï³Ý Ñ³í»Éí³Í"
        dbFOLDERS(9).fSPEC = "01080793012"
        
    Set dbFOLDERS(10) = New_DB_FOLDERS()
        dbFOLDERS(10).fFOLDERID = "ALLACCSRES"
        dbFOLDERS(10).fNAME = "COAgrAcc"
        dbFOLDERS(10).fKEY = fISN
        dbFOLDERS(10).fISN = fISN2
        dbFOLDERS(10).fSTATUS = "1"
        dbFOLDERS(10).fCOM = "¶»ñ³Í³ËëÇ Ñ³ßí³å³Ñ³Ï³Ý Ñ³í»Éí³Í"
        dbFOLDERS(10).fSPEC = "000004532017303038100072112153000"
    
    Set dbFOLDERS(11) = New_DB_FOLDERS()
        dbFOLDERS(11).fFOLDERID = "Agr."&fISN
        dbFOLDERS(11).fNAME = "COAgrAcc"
        dbFOLDERS(11).fKEY = fISN2
        dbFOLDERS(11).fISN = fISN2
        dbFOLDERS(11).fSTATUS = "1"
        dbFOLDERS(11).fCOM = "¶»ñ³Í³ËëÇ Ñ³ßí³å³Ñ³Ï³Ý Ñ³í»Éí³Í"
        dbFOLDERS(11).fSPEC = "1¶»ñ³Í³ËëÇ Ñ³ßí³å³Ñ³Ï³Ý Ñ³í»Éí³Í- 00000113032   "
        
    Set dbFOLDERS(12) = New_DB_FOLDERS()
        dbFOLDERS(12).fFOLDERID = "CAGRACCS"
        dbFOLDERS(12).fNAME = "COAgrAcc"
        dbFOLDERS(12).fKEY = "00000113032   "
        dbFOLDERS(12).fISN = fISN2
        dbFOLDERS(12).fSTATUS = "1"
        dbFOLDERS(12).fCOM = "¶»ñ³Í³ËëÇ Ñ³ßí³å³Ñ³Ï³Ý Ñ³í»Éí³Í"
        dbFOLDERS(12).fSPEC = ""   
     
    Set dbFOLDERS(13) = New_DB_FOLDERS()
        dbFOLDERS(13).fFOLDERID = "Agr."&fISN
        dbFOLDERS(13).fNAME = "COTSDtUn"
        dbFOLDERS(13).fKEY = "COTSDtUn"
        dbFOLDERS(13).fSTATUS = "1"
        dbFOLDERS(13).fCOM = "Ø³ñáõÙÝ»ñÇ ·ñ³ýÇÏ"
        dbFOLDERS(13).fSPEC = "1Ø³ñáõÙÝ»ñÇ ·ñ³ýÇÏ`  01/01/20"    
        dbFOLDERS(13).fECOM = "Ø³ñáõÙÝ»ñÇ ·ñ³ýÇÏ" 
End Sub