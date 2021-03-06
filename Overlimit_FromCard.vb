'USEUNIT Library_Common 
'USEUNIT Library_Colour
'USEUNIT Constants
'USEUNIT Overlimit_Library
'USEUNIT Subsystems_SQL_Library
'USEUNIT Library_CheckDB
'USEUNIT Library_Contracts
'USEUNIT Payment_Except_Library

Option Explicit
'Test Case Id - 146808

Dim fADB
Sub Check_OverlimitFromCard()
    
    Dim sDATE,fDATE,VerificationDoc
    Call Initialize_ForOverlimitFromCard()
     
    'Համակարգ մուտք գործել ARMSOFT օգտագործողով
    sDATE = "20140101"
    fDATE = "20201125"
    Call Initialize_AsBank("bank", sDATE, fDATE)
    Login("ARMSOFT")
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''-Մուտք գործել "Պլաստիկ քարտերի ԱՇՏ (SV)"-''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''       
    'Մուտք գործել "Պլաստիկ քարտերի ԱՇՏ (SV)"
    Call ChangeWorkspace(c_CardsSV) 
    Call wTreeView.DblClickItem("|äÉ³ëïÇÏ ù³ñï»ñÇ ²Þî (SV)|äÉ³ëïÇÏ ù³ñï»ñ")
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''-Îատարել Քարտային վճարում "Պլաստիկ քարտեր" թղթապանակից-'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''    
    Log.Message "Card Payments Action",,,DivideColor   
    
    If Check_CardExist_In_Carsds_Folder("9051190300000047") Then
        PaymentfISN = Card_PaymentAction("201120","2","5500","3")
        BuiltIn.Delay(2000)
        wMDIClient.VBObject("frmPttel").Close
    Else
        Log.Error "Can Not Open Պլաստիկ քարտեր Pttel",,,ErrorColor      
    End If  
    
    Log.Message "SQL Check After Card Payments Action",,,SqlDivideColor
    Log.Message "PaymentfISN = "& PaymentfISN,,,SqlDivideColor
    Call SQL_Initialize_ForOverlimitFromCard(PaymentfISN,"") 
    
    'SQL Ստուգում DOCS աղուսյակում
    fBODY = "  ACSBRANCH:00  ACSDEPART:2  DATE:20201120  CARDNUM:9051190300000047  CARDNUMMASK:905****30****0*7  ACCCODE:01046983311  CARDCURR:000  CRDFEETP:2  FEE:5500  LOCKFEE:0  FEECURR:000  MNTFEETP:3  USEOVERLIMIT:1  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",PaymentfISN,1)
    Call CheckDB_DOCS(PaymentfISN,"CardFee ","1",fBODY,1)
    
    'SQL Ստուգում DOCLOG աղուսյակում 
    Call CheckQueryRowCount("DOCLOG","fISN",PaymentfISN,1)
    Call CheckDB_DOCLOG(PaymentfISN,"77","N","1"," ",1)
    
    'SQL Ստուգում FOLDERS աղուսյակում 
    Call CheckQueryRowCount("FOLDERS","fISN",PaymentfISN,2)
    Call CheckDB_FOLDERS(dbFOLDERS(1),1)
    Call CheckDB_FOLDERS(dbFOLDERS(2),1)
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''-Քարտային վճարումներ թղթապանակից "կատարել քարտային հաշվից" գործողությունը-'''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''      
    Log.Message "Make From Card Account Action",,,DivideColor   
    
    Call Card_Payment("201120") 
    BuiltIn.Delay(2000)
    'Կատարում է ստուգում, եթե քաղվածքի պատուհանը հայտնվել է ,ապա փակում է, հակառակ դեպքում դուրս է բերում սխալ
    If wMDIClient.VBObject("FrmSpr").Exists Then
        wMDIClient.VBObject("FrmSpr").Close
    Else
        Log.Error "The window doesn't exist",,,ErrorColor
    End If
    
    'SQL Ստուգում DOCS աղուսյակում
    Call CheckQueryRowCount("DOCS","fISN",PaymentfISN,1)
    Call CheckDB_DOCS(PaymentfISN,"CardFee ","2",fBODY,1)
    
    'SQL Ստուգում DOCLOG աղուսյակում 
    Call CheckQueryRowCount("DOCLOG","fISN",PaymentfISN,4)
    Call CheckDB_DOCLOG(PaymentfISN,"77","N","1","",1)
    Call CheckDB_DOCLOG(PaymentfISN,"77","E","1","",1)
    Call CheckDB_DOCLOG(PaymentfISN,"77","M","1","CREATED",1)
    Call CheckDB_DOCLOG(PaymentfISN,"77","C","2","",1)
    
    'SQL Ստուգում FOLDERS աղուսյակում 
    dbFOLDERS(2).fSPEC = "905119030000004700001046983311            0.00         5500.00000            0.00    2²µ»É Îáµ»ÉÛ³Ý                                     20201120Øß³ÏíáÕ                         002"
    BuiltIn.Delay(500)
    Call CheckQueryRowCount("FOLDERS","fISN",PaymentfISN,2)
    Call CheckDB_FOLDERS(dbFOLDERS(1),1)
    Call CheckDB_FOLDERS(dbFOLDERS(2),1)   
'    
    Query = "Select fISN from DOCP where fPARENTISN = " & PaymentfISN
    fBASE(0) = my_Row_Count(Query) 
    'SQL Ստուգում DOCP աղուսյակում  
    Call CheckQueryRowCount("DOCP","fPARENTISN",PaymentfISN,1)
    Call CheckDB_DOCP(fBASE(0),"MemOrd  ",PaymentfISN,1)
    
    'SQL Ստուգում HI աղուսյակում  
    Call CheckQueryRowCount("HI","fBASE",fBASE(0),2)
    Call Check_HI_CE_accounting ("20201120",fBASE(0), "11",  "1630361", "5500.00", "000", "5500.00", "FEE", "D")
    Call Check_HI_CE_accounting ("20201120",fBASE(0), "11",  "798984076", "5500.00", "000", "5500.00", "FEE", "C")
    
    'SQL Ստուգում HIREST  աղուսյակում 
    Call CheckDB_HIREST("01", "1630361","-4131.60","000","-4131.60",6)    
    Call CheckDB_HIREST("11", "1630361","5500.00","000","5500.00",1)   
    Call CheckDB_HIREST("11", "798984076","-5500.00","000","-5500.00",1) 
 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''-"Գլխավոր հաշվապահ/Աշխատանքային փաստաթղթեր" թղթապանակից կատարել հաշվառել գործողությունը-'''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "To Count This Payment",,,DivideColor       
    
    'Մուտք "Գլխավոր հաշվապահի ԱՇՏ"   
    Call ChangeWorkspace(c_ChiefAcc)
    
    Call ToCountPayment(c_ToCount,"201120") 
    
    'SQL Ստուգում DOCS աղուսյակում
    Call CheckQueryRowCount("DOCS","fISN",PaymentfISN,1)
    Call CheckDB_DOCS(PaymentfISN,"CardFee ","4",fBODY,1)

    'SQL Ստուգում FOLDERS աղուսյակում 
    dbFOLDERS(2).fSPEC = "905119030000004700001046983311            0.00         5500.00000            0.00    4²µ»É Îáµ»ÉÛ³Ý                                     20201120ì»ñçÝ³Ï³Ý                       002"
    dbFOLDERS(2).fFOLDERID = "CardFee.20201120"
    BuiltIn.Delay(500)
    Call CheckQueryRowCount("FOLDERS","fISN",PaymentfISN,1)
    Call CheckDB_FOLDERS(dbFOLDERS(2),1)   
    
    'SQL Ստուգում HI աղուսյակում  
    Call CheckQueryRowCount("HI","fBASE",fBASE(0),2)
    Call Check_HI_CE_accounting ("20201120",fBASE(0), "01",  "1630361", "5500.00", "000", "5500.00", "FEE", "D")
    Call Check_HI_CE_accounting ("20201120",fBASE(0), "01",  "798984076", "5500.00", "000", "5500.00", "FEE", "C")
    
    'SQL Ստուգում MEMORDERS աղուսյակում  
    Call CheckDB_MEMORDERS(fBASE(0),"MemOrd  ","1","2020-11-20","5","5500.00","000",1)
    
    'SQL Ստուգում HIREST  աղուսյակում 
    Call CheckDB_HIREST("01", "1630361","-4131.60","000","-4131.60",5) 
    Call CheckDB_HIREST("01", "1630361","1368.40","000","1368.40",1)       
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''-Բացել "Գերածախս ունեցող հաշիվներ" թղթապանակը-'''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "Open Accounts with Overlimit Doc",,,DivideColor
    
    'Մուտք գործել "Գերածախս"
    Call ChangeWorkspace(c_Overlimit) 
    
    Call wTreeView.DblClickItem("|¶»ñ³Í³Ëë|¶»ñ³Í³Ëë áõÝ»óáÕ Ñ³ßÇíÝ»ñ|")
    BuiltIn.Delay(delay_middle)
    Call Fill_AccWithOverlimit(AccWithOverlimit)
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''-Հայտնցած տողի վրա կատարել աջ կլիկ - Գերածախսի բացում (խմբ.)-''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "Right Click - Open Overlimit Action",,,DivideColor    
    
    OverlimitAccountIsn = OpenOverimitFromAccount("201120")

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''-Պայմանագրեր թղթապանակում փաստատթղթի առկայության ստուգում-'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "Check (Existing Contract) Function",,,DivideColor
    
    Call ExistsContract_Filter_Fill("|¶»ñ³Í³Ëë|",ContractFillter,1)
    
    'Ստուգում "Մնացորդ", "Ժամկետանց գումարի տույժ", "Հաշվի Մնացորդ" սյուների արժեքները
    Call CompareFieldValue("frmPttel", "fAgrRem", "1,368.40")
    Call CompareFieldValue("frmPttel", "fPenRem", "0.00")
    Call CompareFieldValue("frmPttel", "fAccRem", "0.00")  
    AccountParentIsn = GetAccountIsnOverlimit()
    fISN = GetIsn()
    
    Log.Message "SQL Check After Right Click - Open Overlimit Action",,,SqlDivideColor
    Log.Message "fISN = "& fISN,,,SqlDivideColor 
    Log.Message "AccountParentIsn = "& AccountParentIsn,,,SqlDivideColor   
    Log.Message "OverlimitAccountIsn = "& OverlimitAccountIsn,,,SqlDivideColor   

    Call SQL_Initialize_ForOverlimitFromCard(fISN,AccountParentIsn) 
    'SQL Ստուգում DOCS աղուսյակում  
    fBODY = "  CODE:01046983311  CRDTCODE:777000000250L001  CLICOD:00000025  NAME:²µ»É Îáµ»ÉÛ³Ý  CURRENCY:000  ACCACC:01046983311  AUTODEBT:1  ACCCONNMODE:3  USECLICONNSCH:0  DATE:20201120  DATEGIVE:20201120  ACSBRANCH:00  ACSDEPART:2  ACSTYPE:CO1  KINDSCALE:1  PCPENAGR:0/1  CONSTPER:0  SECTOR:F  SCHEDULE:9  GUARANTEE:9  PERRES:1  PPRCODE:01046983311  SUBJRISK:0  CHRGFIRSTDAY:1  GIVEN:1  PUTINLR:0  NOTCLASS:0  OTHERCOLLATERAL:0  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",fISN,1)
    Call CheckDB_DOCS(fISN,"COSimpl","7",fBODY,1)

    'SQL Ստուգում DOCS աղուսյակում ¶»ñ³Í³ËëÇ Ñ³ßí³å³Ñ³Ï³Ý Ñ³í»Éí³Í-Ի համար
    fBODY = "  CODE:01046983311  CURRENCY:000  CLICOD:00000025  JURSTAT:21  VOLORT:7  PETBUJ:2  REZ:1  RELBANK:0  RABBANK:0  ACCAGR:01080793012  ACCACC:01046983311  FILLACCS:0  OPENACCS:0  TYPEPEN:0  "
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
    
    'SQL Ստուգում ACCOUNTS, HIREST, DOCP աղուսյակներում
    Call CheckQueryRowCount("DOCP","fPARENTISN",AccountParentIsn,5)
    'Row 1
    Query = "SELECT fISN FROM (SELECT *, ROW_NUMBER() OVER (ORDER BY fDCD ASC) AS rownum FROM DOCP where fPARENTISN = "&AccountParentIsn&" and fISN <> 1630361) AS foo WHERE  rownum > 0 AND rownum <= 1 "
    AccIsn(0) = my_Row_Count(Query) 
    Call CheckQueryRowCount("ACCOUNTS","fISN",AccIsn(0),1)
    Call CheckDB_DOCP(AccIsn(0),"Acc     ",AccountParentIsn,1)
    Call CheckDB_HIREST("01", AccIsn(0),"1368.40","000","1368.40",1)    
    'Row 2
    Query = "SELECT fISN FROM (SELECT *, ROW_NUMBER() OVER (ORDER BY fDCD ASC) AS rownum FROM DOCP where fPARENTISN = "&AccountParentIsn&" and fISN <> 1630361) AS foo WHERE  rownum > 1 AND rownum <= 2 "
    AccIsn(1) = my_Row_Count(Query) 
    Call CheckDB_DOCP(AccIsn(1),"Acc     ",AccountParentIsn,1)
    Call CheckQueryRowCount("HIREST","fOBJECT",AccIsn(1),1)
    'Row 3
    Query = "SELECT fISN FROM (SELECT *, ROW_NUMBER() OVER (ORDER BY fDCD ASC) AS rownum FROM DOCP where fPARENTISN = "&AccountParentIsn&" and fISN <> 1630361) AS foo WHERE  rownum > 2 AND rownum <= 3 "
    AccIsn(2) = my_Row_Count(Query) 
    Call CheckQueryRowCount("ACCOUNTS","fISN",AccIsn(2),1)
    Call CheckDB_DOCP(AccIsn(2),"Acc     ",AccountParentIsn,1)
    Call CheckQueryRowCount("HIREST","fOBJECT",AccIsn(2),1)
    'Row 4
    Query = "SELECT fISN FROM (SELECT *, ROW_NUMBER() OVER (ORDER BY fDCD ASC) AS rownum FROM DOCP where fPARENTISN = "&AccountParentIsn&" and fISN <> 1630361) AS foo WHERE  rownum > 3 AND rownum <= 4 "
    AccIsn(3) = my_Row_Count(Query) 
    Call CheckQueryRowCount("ACCOUNTS","fISN",AccIsn(3),1)
    Call CheckDB_DOCP(AccIsn(3),"Acc     ",AccountParentIsn,1)
    Call CheckQueryRowCount("HIREST","fOBJECT",AccIsn(3),1)
    
    'SQL Ստուգում CONTRACTS աղուսյակում 
    Call CheckQueryRowCount("CONTRACTS","fDGISN",fISN,1)
    Call CheckDB_CONTRACTS(dbCONTRACT,1)
    
    'SQL Ստուգում CAGRACCS աղուսյակում 
    Call CheckQueryRowCount("CAGRACCS","fAGRISN",fISN,1)
   
    'SQL Ստուգում HI աղուսյակում 
    Query = "Select fBASE from HI WHERE fSUM = '1368.40'"
    fBASE(1) = my_Row_Count(Query)
    Query = "Select fADB from HI WHERE fSUM = '1368.40'"
    fADB = my_Row_Count(Query)
    Call CheckQueryRowCount("HI","fBASE",fBASE(1),2)
    Call Check_HI_CE_accounting ("20201120",fBASE(1), "01", "1630361", "1368.40", "000", "1368.40", "OVD", "C") 
    Call Check_HI_CE_accounting ("20201120",fBASE(1), "01", fADB, "1368.40", "000", "1368.40", "OVD", "D") 
    
    'SQL Ստուգում HIF  աղուսյակում 
    Call Check_HIF("2020-11-20", "N0", fISN, "0.00", "1.00", "PPA", Null)
    Call Check_HIF("2020-11-20", "N0", fISN, "0.00", "0.00", "LIM", Null)
   
    'SQL Ստուգում HIR աղուսյակում 
    Call CheckQueryRowCount("HIR","fOBJECT",fISN,2)
    Call Check_HIR("20201120", "R1", fISN, "000", "1368.40", "AGR", "D")
    Call Check_HIR("20201120", "RÄ", fISN, "000", "1368.40", "AGJ", "D")
    
   'SQL Ստուգում RESNUMBERS  աղուսյակում 
    Call CheckDB_RESNUMBERS(fISN,"C","01046983311   ",1)
    
    'SQL Ստուգում HIRREST  աղուսյակում 
    Call CheckQueryRowCount("HIRREST","fOBJECT",fISN,2)
    Call CheckDB_HIRREST("R1",fISN,"1368.40","20201120",1)
    Call CheckDB_HIRREST("RÄ",fISN,"1368.40","20201120",1)    
    
    'SQL Ստուգում FOLDERS աղուսյակում 
    Call CheckQueryRowCount("FOLDERS","fISN",fISN,3)
    Call CheckDB_FOLDERS(dbFOLDERS(3),1)
    Call CheckDB_FOLDERS(dbFOLDERS(4),1)
    Call CheckDB_FOLDERS(dbFOLDERS(5),1)

    Call CheckQueryRowCount("FOLDERS","fISN",AccountParentIsn,2)
    Call CheckDB_FOLDERS(dbFOLDERS(9),1)
    Call CheckDB_FOLDERS(dbFOLDERS(10),1)
     
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''-"Պարտքերի մարում" գործողության կատարում-'''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''      
    Log.Message "Check RC (Overlimit Repay/Պարտքերի մարում) Function",,,DivideColor

    Call Overlimit_Repay(NewOverlimitRepay)
    
    'Կատարում է ստուգում, եթե քաղվածքի պատուհանը հայտնվել է ,ապա փակում է, հակառակ դեպքում դուրս է բերում սխալ
    If wMDIClient.VBObject("FrmSpr").Exists Then
        wMDIClient.VBObject("FrmSpr").Close
    Else
        Log.Error "The window doesn't exist",,,ErrorColor
    End If
    wMDIClient.VBObject("frmPttel").Close
    
    Log.Message "SQL Check After RC (Overlimit Repay/Պարտքերի մարում) Function",,,SqlDivideColor
    Log.Message "fISN = " & NewOverlimitRepay.Isn,,,SqlDivideColor
    Call SQL_Initialize_ForOverlimitFromCard(fISN,NewOverlimitRepay.Isn) 
    
    'SQL Ստուգում DOCS աղուսյակում  
    fBODY = "  CODE:01046983311  DATE:20201120  SUMAGR:1368.4  SUMMA:1368.4  CASHORNO:1  ISPUSA:0  COMMENT:Overlimit From Card  ACSBRANCH:00  ACSDEPART:1  ACSTYPE:CO1  USERID:  77  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",NewOverlimitRepay.Isn,1)
    Call CheckDB_DOCS(NewOverlimitRepay.Isn,"CODSDebt","1",fBODY,1)
    
    'SQL Ստուգում DOCLOG աղուսյակում 
    Call CheckQueryRowCount("DOCLOG","fISN",NewOverlimitRepay.Isn,5)
    Call CheckDB_DOCLOG(NewOverlimitRepay.Isn,"77","N","1","",1)
    Call CheckDB_DOCLOG(NewOverlimitRepay.Isn,"77","M","99","àõÕ³ñÏí»É ¿ Ñ³ëï³ïÙ³Ý",1)
    Call CheckDB_DOCLOG(NewOverlimitRepay.Isn,"77","M","1","CREATED",1)
    Call CheckDB_DOCLOG(NewOverlimitRepay.Isn,"77","C","1","",2)

    'SQL Ստուգում DOCP աղուսյակում  
    Call CheckQueryRowCount("DOCP","fPARENTISN",NewOverlimitRepay.Isn,1)
    
    'SQL Ստուգում FOLDERS աղուսյակում  
    Call CheckQueryRowCount("FOLDERS","fISN",NewOverlimitRepay.Isn,5)
    Call CheckDB_FOLDERS(dbFOLDERS(11),1)
    Call CheckDB_FOLDERS(dbFOLDERS(12),1)
    Call CheckDB_FOLDERS(dbFOLDERS(13),1)
    Call CheckDB_FOLDERS(dbFOLDERS(14),1)
    Call CheckDB_FOLDERS(dbFOLDERS(15),1)    
    
    Query = "Select fISN from DOCP where fPARENTISN = " & NewOverlimitRepay.Isn
    fBASE(2) = my_Row_Count(Query)  
    
    'SQL Ստուգում DOCP աղուսյակում  
    Call CheckQueryRowCount("DOCP","fPARENTISN",NewOverlimitRepay.Isn,1)
    Call CheckDB_DOCP(fBASE(2),"KasPrOrd",NewOverlimitRepay.Isn,1)
    
    'SQL Ստուգում HI աղուսյակում  
    Call CheckQueryRowCount("HI","fBASE",fBASE(2),2)
    Query = "Select fOBJECT from HI where fDBCR = 'D' and fBASE = " & fBASE(2)
    fOBJECT(1) = my_Row_Count(Query) 
    Call Check_HI_CE_accounting ("20201120",fBASE(2), "11",  fOBJECT(1), "1368.40", "000", "1368.40", "OVO", "D")
    
    Query = "Select fOBJECT from HI where fDBCR = 'C' and fBASE = " & fBASE(2)
    fOBJECT(2) = my_Row_Count(Query) 
    Call Check_HI_CE_accounting ("20201120",fBASE(2), "11",  fOBJECT(2), "1368.40", "000", "1368.40", "OVO", "C")
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''-Գլխավոր հաշվապահ/Աշխատանքային փաստաթղթեր թղթապանակից կատարել "Ուղարկել հաստատման" գործողությունը-''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "To Count This Payment",,,DivideColor       
    
    'Մուտք "Գլխավոր հաշվապահի ԱՇՏ"   
    Call ChangeWorkspace(c_ChiefAcc)
    DocNum = ToCountPayment(c_SendToVer,"201120")     
    
    Log.Message "SQL Check After Գլխավոր հաշվապահ/Աշխատանքային փաստաթղթեր թղթապանակից կատարել Ուղարկել հաստատման Function",,,SqlDivideColor
    
    'SQL Ստուգում DOCS աղուսյակում
    fBODY = "  ACSBRANCH:00  ACSDEPART:2  DATE:20201120  CARDNUM:9051190300000047  CARDNUMMASK:905****30****0*7  ACCCODE:01046983311  CARDCURR:000  CRDFEETP:2  FEE:5500  LOCKFEE:0  FEECURR:000  MNTFEETP:3  USEOVERLIMIT:1  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",PaymentfISN,1)
    Call CheckDB_DOCS(PaymentfISN,"CardFee ","4",fBODY,1)
    
    'SQL Ստուգում DOCLOG աղուսյակում 
    Call CheckQueryRowCount("DOCLOG","fISN",PaymentfISN,6)
    Call CheckDB_DOCLOG(PaymentfISN,"77","N","1","",1)
    Call CheckDB_DOCLOG(PaymentfISN,"77","E","1","",1)
    Call CheckDB_DOCLOG(PaymentfISN,"77","M","1","CREATED",1)
    Call CheckDB_DOCLOG(PaymentfISN,"77","C","2","",1)
    Call CheckDB_DOCLOG(PaymentfISN,"77","M","2","PROCESSED",1)
    Call CheckDB_DOCLOG(PaymentfISN,"77","C","4","",1)
    
    'SQL Ստուգում DOCLOG աղուսյակում 
    Call CheckQueryRowCount("DOCLOG","fISN",fBASE(2),2)
    Call CheckDB_DOCLOG(fBASE(2),"77","N","5","",1)
    
    'SQL Ստուգում FOLDERS աղուսյակում 
    Call CheckQueryRowCount("FOLDERS","fISN",fBASE(2),3)
    
    Set dbFOLDERS(16) = New_DB_FOLDERS()
        dbFOLDERS(16).fFOLDERID = "Ver.20201120001"
        dbFOLDERS(16).fNAME = "KasPrOrd"
        dbFOLDERS(16).fKEY = fBASE(2)
        dbFOLDERS(16).fISN = fBASE(2)
        dbFOLDERS(16).fSTATUS = "4"
        dbFOLDERS(16).fCOM = "Î³ÝËÇÏ Ùáõïù"
        dbFOLDERS(16).fSPEC = DocNum & "77700000001100  7770001080793012         1368.40000  77¶»ñ³Í³ËëÇ Ù³ñáõÙ                ä³ÛÙ³Ý³·Çñª 01046983311         ²µ»É Îáµ»ÉÛ³Ý                   "   
        dbFOLDERS(16).fECOM = "Cash Deposit Advice"
        dbFOLDERS(16).fDCBRANCH = "00 "
        dbFOLDERS(16).fDCDEPART = "1  "  
    Call CheckDB_FOLDERS(dbFOLDERS(16),1)   
        
    Set dbFOLDERS(17) = New_DB_FOLDERS()
        dbFOLDERS(17).fFOLDERID = "C.1628339"
        dbFOLDERS(17).fNAME = "KasPrOrd"
        dbFOLDERS(17).fKEY = fBASE(2)
        dbFOLDERS(17).fISN = fBASE(2)
        dbFOLDERS(17).fSTATUS = "0"
        dbFOLDERS(17).fCOM = "Î³ÝËÇÏ Ùáõïù"
        dbFOLDERS(17).fSPEC = "²Ùë³ÃÇí- 20/11/20 N- "&DocNum&" ¶áõÙ³ñ-             1,368.40 ²ñÅ.- 000 [àõÕ³ñÏí³Í I Ñ³ëï³ïÙ³Ý]"   
        dbFOLDERS(17).fECOM = "Cash Deposit Advice"
    Call CheckDB_FOLDERS(dbFOLDERS(17),1)  
       
    Set dbFOLDERS(18) = New_DB_FOLDERS()
        dbFOLDERS(18).fFOLDERID = "Oper.20201120"
        dbFOLDERS(18).fNAME = "KasPrOrd"
        dbFOLDERS(18).fKEY = fBASE(2)
        dbFOLDERS(18).fISN = fBASE(2)
        dbFOLDERS(18).fSTATUS = "0"
        dbFOLDERS(18).fCOM = "Î³ÝËÇÏ Ùáõïù"
        dbFOLDERS(18).fSPEC = DocNum & "77700000001100  7770001080793012         1368.40000àõÕ³ñÏí³Í I Ñ³ëï³ïÙ³Ý                                 77²µ»É Îáµ»ÉÛ³Ý                                                   001                             ¶»ñ³Í³ËëÇ Ù³ñáõÙ ä³ÛÙ³Ý³·Çñª 01046983311                                                                                                    "   
        dbFOLDERS(18).fECOM = "Cash Deposit Advice"
        dbFOLDERS(18).fDCBRANCH = "00 "
        dbFOLDERS(18).fDCDEPART = "1" 
    Call CheckDB_FOLDERS(dbFOLDERS(18),1)   

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''-Գլխավոր հաշվապահ/Հաստատվող փաստաթղթեր(|) թղթապանակից կատարել "Վավերացնել" գործողությունը-''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "To Count This Payment",,,DivideColor   

    Set VerificationDoc = New_VerificationDocument()
    
    Call GoToVerificationDocument("|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ (I)",VerificationDoc)
    
    If WaitForPttel("frmPttel") Then
        If SearchInPttel("frmPttel",7, "1368.4") Then
            Call wMainForm.MainMenu.Click(c_AllActions)
            Call wMainForm.PopupMenu.Click(c_ToConfirm)
            BuiltIn.Delay(3000)
            Call ClickCmdButton(1, "Ð³ëï³ï»É")
        Else 
            Log.Error "Տողը չի գտնվել Հաստատվող փաստաթղթեր(|) թղթապանակում" ,,,ErrorColor
        End If
        BuiltIn.Delay(delay_middle)
        wMDIClient.WaitVBObject("frmPttel",delay_middle).Close
     Else
        Log.Error "Can Not Open Հաստատվող փաստաթղթեր(|) Window",,,ErrorColor      
     End If     
     If DocForm.Exists Then
        Log.Error "Can Not Close Հաստատվող փաստաթղթեր(|) Window",,,ErrorColor
     End If    
     
     Log.Message "SQL Check After Գլխավոր հաշվապահ/Հաստատվող փաստաթղթեր(|) թղթապանակից կատարել վավերացնել Function",,,SqlDivideColor
     
    'SQL Ստուգում DOCS աղուսյակում
    fBODY = "  ACSBRANCH:00  ACSDEPART:1  BLREP:0  OPERTYPE:OVO  TYPECODE:20,21,27,90,99  USERID:  77  DOCNUM:"&DocNum&"  DATE:20201120  KASSA:001  ACCDB:000001100  CUR:000  ACCCR:01080793012  SUMMA:1368.4  BASE:ä³ÛÙ³Ý³·Çñª 01046983311  AIM:¶»ñ³Í³ËëÇ Ù³ñáõÙ  CLICODE:00000025  PAYER:²µ»É  PAYERLASTNAME:Îáµ»ÉÛ³Ý  PASSNUM:AA025836955KOGHMIC  PASTYPE:01  DATEPASS:19990516  CITIZENSHIP:1  COUNTRY:AM  ADDRESS:ù. ºñ¨³Ý, Øá½³ÙµÇÏÇ 5  ISTLLCREATED:1  ACSBRANCHINC:00  ACSDEPARTINC:2  VOLORT:7  NONREZ:0  JURSTAT:21  USEOVERLIMIT:0  NOTSENDABLE:0  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",fBASE(2),1)
    Call CheckDB_DOCS(fBASE(2),"KasPrOrd","15",fBODY,1)
     
    'SQL Ստուգում DOCLOG աղուսյակում 
    Call CheckQueryRowCount("DOCLOG","fISN",fBASE(2),4)
    Call CheckDB_DOCLOG(fBASE(2),"77","N","5","",1)
    Call CheckDB_DOCLOG(fBASE(2),"77","W","102","",1)
    Call CheckDB_DOCLOG(fBASE(2),"77","C","15","",1)
    
    'SQL Ստուգում FOLDERS աղուսյակում
    dbFOLDERS(17).fSTATUS = "4"
    dbFOLDERS(17).fSPEC = "²Ùë³ÃÇí- 20/11/20 N- "&DocNum&" ¶áõÙ³ñ-             1,368.40 ²ñÅ.- 000 [Ð³ëï³ïí³Í]"   
    dbFOLDERS(18).fSTATUS = "4"
    dbFOLDERS(18).fSPEC = DocNum & "77700000001100  7770001080793012         1368.40000Ð³ëï³ïí³Í                                             77²µ»É Îáµ»ÉÛ³Ý                   AA025836955KOGHMIC  16/05/1999                                  ¶»ñ³Í³ËëÇ Ù³ñáõÙ ä³ÛÙ³Ý³·Çñª 01046983311                                                                                                    "
	  BuiltIn.Delay(500)
    
    Call CheckDB_FOLDERS(dbFOLDERS(16),0) 
    Call CheckDB_FOLDERS(dbFOLDERS(17),1)   
    Call CheckDB_FOLDERS(dbFOLDERS(18),1)  
     
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''-Գլխավոր հաշվապահ/Աշխատանքային փաստաթղթեր թղթապանակից կատարել "Վավերացնել" գործողությունը-''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "To Count This Payment",,,DivideColor       
    
    'Մուտք Գլխավոր հաշվապահի ԱՇՏ   
    Call ChangeWorkspace(c_ChiefAcc)
    Call ToCountPayment(c_ToConfirm,"201120")     
    
    Log.Message "SQL Check After Գլխավոր հաշվապահ/Աշխատանքային փաստաթղթեր թղթապանակից կատարել Վավերացնել Function",,,SqlDivideColor
     
    'SQL Ստուգում DOCS աղուսյակում
    Call CheckQueryRowCount("DOCS","fISN",fBASE(2),1)
    Call CheckDB_DOCS(fBASE(2),"KasPrOrd","11",fBODY,1)
    
    'SQL Ստուգում DOCLOG աղուսյակում 
    Call CheckQueryRowCount("DOCLOG","fISN",fBASE(2),6)
    Call CheckDB_DOCLOG(fBASE(2),"77","W","16","",1)
    Call CheckDB_DOCLOG(fBASE(2),"77","M","11","¶ñ³Ýóí»É »Ý Ó¨³Ï»ñåáõÙÝ»ñÁ",1)
    
    'SQL Ստուգում FOLDERS աղուսյակում
    Call CheckDB_FOLDERS(dbFOLDERS(17),0)   
    Call CheckDB_FOLDERS(dbFOLDERS(18),0)  
    
    'SQL Ստուգում HI աղուսյակում  
    Call CheckQueryRowCount("HI","fBASE",fBASE(2),2)
    Call Check_HI_CE_accounting ("20201120",fBASE(2), "01",  fOBJECT(1), "1368.40", "000", "1368.40", "OVO", "D")
    Call Check_HI_CE_accounting ("20201120",fBASE(2), "01",  fOBJECT(2), "1368.40", "000", "1368.40", "OVO", "C")
    
    'SQL Ստուգում HIR աղուսյակում 
    Call CheckQueryRowCount("HIR","fOBJECT",fISN,4)
    Call Check_HIR("20201120", "R1", fISN, "000", "1368.40", "DBT", "C")
    Call Check_HIR("20201120", "RÄ", fISN, "000", "1368.40", "DBT", "C")
    
    'SQL Ստուգում HIREST  աղուսյակում 
    Call CheckDB_HIREST("01", "1630361","-4131.60","000","-4131.60",5) 
    Call CheckDB_HIREST("01", "1630361","0.00","000","0.00",2)  
    
    'SQL Ստուգում PAYMENTS աղուսյակում 
    Call CheckQueryRowCount("PAYMENTS","fISN",fBASE(2),1)
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''-Գերածախս/Պայմանագրեր թղթապանակում փաստատթղթի Փոփոխության ստուգում-'''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "Check (Existing Contract) Function",,,DivideColor
    
    'Մուտք գործել "Գերածախս"
    Call ChangeWorkspace(c_Overlimit) 
    
    Call ExistsContract_Filter_Fill("|¶»ñ³Í³Ëë|",ContractFillter,1)    
    'Ստուգում "Մնացորդ", "Ժամկետանց գումարի տույժ", "Հաշվի Մնացորդ" սյուների արժեքները
    Call CompareFieldValue("frmPttel", "fAgrRem", "0.00")
    Call CompareFieldValue("frmPttel", "fPenRem", "0.00")
    Call CompareFieldValue("frmPttel", "fAccRem", "0.00") 
    wMDIClient.VBObject("frmPttel").Close
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''-"Հաշվառված վճարային փաստաթղթերից" հեռացնել "Կանխիք մուտք" և "Պարտքերի մարում" գործողությունները-'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "Delete Cash deposit and Overlimit Repay Action",,,DivideColor     
    
    'Մուտք Գլխավոր հաշվապահի ԱՇՏ   
    Call ChangeWorkspace(c_ChiefAcc)
    
    wTreeView.DblClickItem("|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|Ð³ßí³éí³Í í×³ñ³ÛÇÝ ÷³ëï³ÃÕÃ»ñ")
    'Լրացնել "Ամսաթիվ" դաշտը
    Call Rekvizit_Fill("Dialog",1,"General","PERN", "201120")
    Call Rekvizit_Fill("Dialog",1,"General","PERK", "201120")
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    
    Set DocForm = wMDIClient.VBObject("frmPttel")
    If WaitForPttel("frmPttel") Then
        Call SearchAndDelete("frmPttel", 1, "Î³ÝËÇÏ Ùáõïù", "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ") 
        BuiltIn.Delay(2000)
        Call MessageExists(2,"Ð»é³óÝ»É Ó¨³Ï»ñåáõÙÝ»ñÁ")
        Call ClickCmdButton(5, "²Ûá") 
        Call SearchAndDelete("frmPttel", 1, "ÐÇß³ñ³ñ ûñ¹»ñ", "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ") 
        BuiltIn.Delay(2000)
        wMDIClient.WaitVBObject("frmPttel",delay_middle).Close
     Else
        Log.Error "Can Not Open Հաշվառված վճարային փաստաթղթեր Window",,,ErrorColor      
     End If     
     If DocForm.Exists Then
        Log.Error "Can Not Close Հաշվառված վճարային փաստաթղթեր Window",,,ErrorColor
     End If    
     
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''-Հեռացնել "Կարգադրություններ"-ից կատարված մարում գործողությունը-'''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "Delete Order Action",,,DivideColor     
    
    wTreeView.DblClickItem("|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|Î³ñ·³¹ñáõÃÛáõÝÝ»ñ")
     
    Set DocForm = wMDIClient.VBObject("frmPttel")
    If WaitForPttel("frmPttel") Then
        Call SearchAndDelete("frmPttel", 3, "1368.4", "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ") 
        BuiltIn.Delay(1500)
        wMDIClient.WaitVBObject("frmPttel",delay_middle).Close
     Else
        Log.Error "Can Not Open Կարգադրություններ Window",,,ErrorColor      
     End If     
     If DocForm.Exists Then
        Log.Error "Can Not Close Կարգադրություններ Window",,,ErrorColor
     End If  
     
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''-"Քարտային վճարումներ" թղթապանակից "հեռացնել" կատարված վճարում գործողությունը-'''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''      
    Log.Message "Delete Order Action",,,DivideColor   
    
    Call ChangeWorkspace(c_CardsSV) 
    Call wTreeView.DblClickItem(("|äÉ³ëïÇÏ ù³ñï»ñÇ ²Þî (SV)|ÂÕÃ³å³Ý³ÏÝ»ñ|ø³ñï³ÛÇÝ í×³ñáõÙÝ»ñ"))
    BuiltIn.Delay(1500)
    Set DocForm = wMDIClient.VBObject("frmPttel")
    If WaitForPttel("frmPttel") Then
        Call SearchAndDelete("frmPttel", 0, "20/11/20", "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ") 
        BuiltIn.Delay(1500)
        wMDIClient.WaitVBObject("frmPttel",delay_middle).Close
     Else
        Log.Error "Can Not Open Քարտային վճարումներ Window",,,ErrorColor      
     End If     
     If DocForm.Exists Then
        Log.Error "Can Not Close Քարտային վճարումներ Window",,,ErrorColor
     End If   
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''-Պայմանագրեր թղթապանակից հեռացնել փաստատթուղթը-''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "Delete Contract",,,DivideColor
    
    'Մուտք գործել "Գերածախս"
    Call ChangeWorkspace(c_Overlimit) 
    
    Call ExistsContract_Filter_Fill("|¶»ñ³Í³Ëë|",ContractFillter,1)
    'Ստուգում "Մնացորդ", "Ժամկետանց գումարի տույժ", "Հաշվի Մնացորդ" սյուների արժեքները
    Call CompareFieldValue("frmPttel", "fAgrRem", "1,368.40")
    Call CompareFieldValue("frmPttel", "fPenRem", "0.00")
    Call CompareFieldValue("frmPttel", "fAccRem", "5,500.00")   
    Call SearchAndDelete("frmPttel", 3, "1368.4", "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ") 
    wMDIClient.WaitVBObject("frmPttel",delay_middle).Close
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''-SQL Check After Deleted All Action-'''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 
    
    Log.Message "SQL Check After Deleted All Action",,,SqlDivideColor
    
    'SQL Ստուգում DOCS աղուսյակում
    fBODY = "  ACSBRANCH:00  ACSDEPART:2  DATE:20201120  CARDNUM:9051190300000047  CARDNUMMASK:905****30****0*7  ACCCODE:01046983311  CARDCURR:000  CRDFEETP:2  FEE:5500  LOCKFEE:0  FEECURR:000  MNTFEETP:3  USEOVERLIMIT:1  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",PaymentfISN,1)
    Call CheckDB_DOCS(PaymentfISN,"CardFee ","999",fBODY,1)
    
    fBODY = "  CODE:01046983311  CRDTCODE:777000000250L001  CLICOD:00000025  NAME:²µ»É Îáµ»ÉÛ³Ý  CURRENCY:000  ACCACC:01046983311  AUTODEBT:1  ACCCONNMODE:3  USECLICONNSCH:0  DATE:20201120  DATEGIVE:20201120  ACSBRANCH:00  ACSDEPART:2  ACSTYPE:CO1  KINDSCALE:1  PCPENAGR:0/1  CONSTPER:0  SECTOR:F  SCHEDULE:9  GUARANTEE:9  PERRES:1  PPRCODE:01046983311  SUBJRISK:0  CHRGFIRSTDAY:1  GIVEN:1  PUTINLR:0  NOTCLASS:0  OTHERCOLLATERAL:0  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",fISN,1)
    Call CheckDB_DOCS(fISN,"COSimpl","999",fBODY,1)
    
    fBODY = "  ACSBRANCH:00  ACSDEPART:1  BLREP:0  OPERTYPE:OVO  TYPECODE:20,21,27,90,99  USERID:  77  DOCNUM:"&DocNum&"  DATE:20201120  KASSA:001  ACCDB:000001100  CUR:000  ACCCR:01080793012  SUMMA:1368.4  BASE:ä³ÛÙ³Ý³·Çñª 01046983311  AIM:¶»ñ³Í³ËëÇ Ù³ñáõÙ  CLICODE:00000025  PAYER:²µ»É  PAYERLASTNAME:Îáµ»ÉÛ³Ý  PASSNUM:AA025836955KOGHMIC  PASTYPE:01  DATEPASS:19990516  CITIZENSHIP:1  COUNTRY:AM  ADDRESS:ù. ºñ¨³Ý, Øá½³ÙµÇÏÇ 5  ISTLLCREATED:1  ACSBRANCHINC:00  ACSDEPARTINC:2  VOLORT:7  NONREZ:0  JURSTAT:21  USEOVERLIMIT:0  NOTSENDABLE:0  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",fBASE(2),1)
    Call CheckDB_DOCS(fBASE(2),"KasPrOrd","999",fBODY,1)
    
    'SQL Ստուգում DOCLOG աղուսյակում 
    Call CheckDB_DOCLOG(fISN,"77","D","999","",1)
    
    'SQL Ստուգում FOLDERS աղուսյակում 
    Set dbFOLDERS_ForDelete = New_DB_FOLDERS()
        dbFOLDERS_ForDelete.fFOLDERID = ".R." & aqConvert.DateTimeToFormatStr(aqDateTime.Now(),"%Y%m%d")
        dbFOLDERS_ForDelete.fNAME = "CardFee "
        dbFOLDERS_ForDelete.fKEY = PaymentfISN
        dbFOLDERS_ForDelete.fISN = PaymentfISN
        dbFOLDERS_ForDelete.fSTATUS = "0"
        dbFOLDERS_ForDelete.fSPEC = Left_Align(Get_Compname_DOCLOG(fISN), 16)&"PCardSV ARMSOFT                       001  "
        dbFOLDERS_ForDelete.fDCBRANCH = "00"
        dbFOLDERS_ForDelete.fDCDEPART = "2"
    Call CheckDB_FOLDERS(dbFOLDERS_ForDelete,1)
    
        dbFOLDERS_ForDelete.fNAME = "MemOrd  "
        dbFOLDERS_ForDelete.fKEY = fBASE(0)
        dbFOLDERS_ForDelete.fISN = fBASE(0)
        dbFOLDERS_ForDelete.fSPEC = Left_Align(Get_Compname_DOCLOG(fISN), 16)&"GlavBux ARMSOFT                       115  "
        dbFOLDERS_ForDelete.fDCDEPART = "1"
    Call CheckDB_FOLDERS(dbFOLDERS_ForDelete,1)
    
        dbFOLDERS_ForDelete.fNAME = "Acc     "
        dbFOLDERS_ForDelete.fKEY = AccIsn(0) 
        dbFOLDERS_ForDelete.fISN = AccIsn(0) 
        dbFOLDERS_ForDelete.fSPEC = Left_Align(Get_Compname_DOCLOG(fISN), 16)&"CredO   ARMSOFT                       002  "
        dbFOLDERS_ForDelete.fDCDEPART = "2"
    Call CheckDB_FOLDERS(dbFOLDERS_ForDelete,1)
    
        dbFOLDERS_ForDelete.fKEY = AccIsn(1) 
        dbFOLDERS_ForDelete.fISN = AccIsn(1) 
    Call CheckDB_FOLDERS(dbFOLDERS_ForDelete,1)
    
        dbFOLDERS_ForDelete.fKEY = AccIsn(2) 
        dbFOLDERS_ForDelete.fISN = AccIsn(2) 
    Call CheckDB_FOLDERS(dbFOLDERS_ForDelete,1)
    
        dbFOLDERS_ForDelete.fKEY = AccIsn(3) 
        dbFOLDERS_ForDelete.fISN = AccIsn(3) 
    Call CheckDB_FOLDERS(dbFOLDERS_ForDelete,1)
    
        dbFOLDERS_ForDelete.fNAME = "COAgrAcc"
        dbFOLDERS_ForDelete.fKEY = AccountParentIsn
        dbFOLDERS_ForDelete.fISN = AccountParentIsn
        dbFOLDERS_ForDelete.fSPEC = Left_Align(Get_Compname_DOCLOG(fISN), 16)&"CredO   ARMSOFT                       0021 "
        dbFOLDERS_ForDelete.fDCBRANCH = ""
        dbFOLDERS_ForDelete.fDCDEPART = ""
    Call CheckDB_FOLDERS(dbFOLDERS_ForDelete,1)
    
        dbFOLDERS_ForDelete.fNAME = "COSimpl "
        dbFOLDERS_ForDelete.fKEY = fISN
        dbFOLDERS_ForDelete.fISN = fISN
        dbFOLDERS_ForDelete.fSPEC = Left_Align(Get_Compname_DOCLOG(fISN), 16)&"CredO   ARMSOFT                       007  "
        dbFOLDERS_ForDelete.fDCBRANCH = "00"
        dbFOLDERS_ForDelete.fDCDEPART = "2"
    Call CheckDB_FOLDERS(dbFOLDERS_ForDelete,1)
    
        dbFOLDERS_ForDelete.fNAME = "CODSAgr "
        dbFOLDERS_ForDelete.fKEY = fBASE(1) 
        dbFOLDERS_ForDelete.fISN = fBASE(1) 
        dbFOLDERS_ForDelete.fSPEC = Left_Align(Get_Compname_DOCLOG(fISN), 16)&"CredO   ARMSOFT                       115  "
    Call CheckDB_FOLDERS(dbFOLDERS_ForDelete,1)
    
        dbFOLDERS_ForDelete.fNAME = "CODSDebt"
        dbFOLDERS_ForDelete.fKEY = NewOverlimitRepay.Isn
        dbFOLDERS_ForDelete.fISN = NewOverlimitRepay.Isn
        dbFOLDERS_ForDelete.fSPEC = Left_Align(Get_Compname_DOCLOG(fISN), 16)&"GlavBux ARMSOFT                       001  "
        dbFOLDERS_ForDelete.fDCDEPART = "1"
    Call CheckDB_FOLDERS(dbFOLDERS_ForDelete,1)
    
        dbFOLDERS_ForDelete.fNAME = "KasPrOrd"
        dbFOLDERS_ForDelete.fKEY = fBASE(2)
        dbFOLDERS_ForDelete.fISN = fBASE(2)
        dbFOLDERS_ForDelete.fSPEC = Left_Align(Get_Compname_DOCLOG(fISN), 16)&"GlavBux ARMSOFT                       1111 "
    Call CheckDB_FOLDERS(dbFOLDERS_ForDelete,1)

    'SQL Ստուգում ACCOUNTS, HIREST, DOCP աղուսյակներում
    Call CheckQueryRowCount("DOCP","fPARENTISN",PaymentfISN,0)
    Call CheckQueryRowCount("DOCP","fPARENTISN",AccountParentIsn,0)
    Call CheckQueryRowCount("DOCP","fPARENTISN",NewOverlimitRepay.Isn,0)

    Call CheckQueryRowCount("ACCOUNTS","fISN",AccIsn(0),0)
    Call CheckQueryRowCount("ACCOUNTS","fISN",AccIsn(1),0)
    Call CheckQueryRowCount("ACCOUNTS","fISN",AccIsn(2),0)

    Call CheckQueryRowCount("HIREST","fOBJECT",AccIsn(1),0)
    Call CheckQueryRowCount("HIREST","fOBJECT",AccIsn(2),0)
    Call CheckQueryRowCount("HIREST","fOBJECT",AccIsn(3),0)
    
    'SQL Ստուգում CONTRACTS աղուսյակում 
    Call CheckQueryRowCount("CONTRACTS","fDGISN",fISN,0)
    
    'SQL Ստուգում CAGRACCS աղուսյակում 
    Call CheckQueryRowCount("CAGRACCS","fAGRISN",fISN,0)
    
    'SQL Ստուգում HI աղուսյակում 
    Call CheckQueryRowCount("HI","fBASE",fBASE(1),0)
    Call CheckQueryRowCount("HI","fBASE",fBASE(2),0)
    
    'SQL Ստուգում HIF աղուսյակում 
    Call CheckQueryRowCount("HIF","fOBJECT",fISN,0)

    'SQL Ստուգում HIR աղուսյակում 
    Call CheckQueryRowCount("HIR","fOBJECT",fISN,0)
    
    'SQL Ստուգում RESNUMBERS  աղուսյակում 
    Call CheckQueryRowCount("RESNUMBERS","fISN",fISN,0)
    
    'SQL Ստուգում HIRREST  աղուսյակում 
    Call CheckQueryRowCount("HIRREST","fOBJECT",fISN,0)
    
    'SQL Ստուգում HIREST  աղուսյակում 
    Call CheckDB_HIREST("01", "1630361","-4131.60","000","-4131.60",6) 
    Call CheckDB_HIREST("01", "1630361","0.00","000","0.00",1)  
    Call CheckDB_HIREST("01", "1630361","0.00","XXX","-999999999999.99",1)  
    Call CheckQueryRowCount("HIREST","fOBJECT",AccIsn(0),0)
    Call CheckQueryRowCount("HIREST","fOBJECT",AccIsn(1),0)
    Call CheckQueryRowCount("HIREST","fOBJECT",AccIsn(2),0)
    Call CheckQueryRowCount("HIREST","fOBJECT",AccIsn(3),0)
    
    'SQL Ստուգում PAYMENTS աղուսյակում 
    Call CheckQueryRowCount("PAYMENTS","fISN",fBASE(2),0)
    
    Call Close_AsBank()
End Sub  

Sub Initialize_ForOverlimitFromCard()
      
        '"Գերածախս ունեցող հաշիվներ" Ֆիլտրի լրացման արժեքներ
    Set AccWithOverlimit = New_AccountsWithOverlimit()
        AccWithOverlimit.Curr = "000"
        AccWithOverlimit.Client = "00000025"
        AccWithOverlimit.AccountMask = "01046983311"
            
        '"Պայմանագրեր" Ֆիլտրի լրացման արժեքներ
    Set ContractFillter = New_ContractOverlimit()
        ContractFillter.AgreementLevel = "1"
        ContractFillter.AgreementSpecies = "5"
        ContractFillter.Curr = "000"
        ContractFillter.Client = "00000025"
    
        '"Աջ կլիկ/Պարտքերի մարում" Պատուհանի լրացման արժեքներ
    Set NewOverlimitRepay = New_RcOverlimitRepay()   
        NewOverlimitRepay.ExpectedAgreementN = "01046983311"
        NewOverlimitRepay.Date = "201120"
        NewOverlimitRepay.ExpectedBaseSum = "1,368.40"
        NewOverlimitRepay.BaseSum = "1,368.40"
        NewOverlimitRepay.AMD1 = "0.00"
        NewOverlimitRepay.ExpectedFineOnPastSum = "0.00"
        NewOverlimitRepay.AMD2 = "0.00"
        NewOverlimitRepay.TotalAmount = "1,368.40"
        NewOverlimitRepay.CashCashles = "1"
        NewOverlimitRepay.Account = "01046983311"
        NewOverlimitRepay.Comment =  "Overlimit From Card"
        NewOverlimitRepay.RemittanceInfo1 = "Ø.·.`  1,368.40;" 
        NewOverlimitRepay.RemittanceInfo2 = "Ø.·.ïÅ.`  0.00"
        NewOverlimitRepay.Division = "00"
        NewOverlimitRepay.Department = "1"   
End Sub

Sub SQL_Initialize_ForOverlimitFromCard(fISN,fISN2)

    Set dbCONTRACT = New_DB_CONTRACTS()
        dbCONTRACT.fDGISN = fISN
        dbCONTRACT.fDGPARENTISN = fISN
        dbCONTRACT.fDGISN1 = fISN
        dbCONTRACT.fDGISN3 = fISN
        dbCONTRACT.fDGAGRKIND = "5 "
        dbCONTRACT.fDGSTATE = "7"
        dbCONTRACT.fDGTYPENAME = "COSimpl "
        dbCONTRACT.fDGCODE = "01046983311   "
        dbCONTRACT.fDGPPRCODE = "01046983311"
        dbCONTRACT.fDGCAPTION = "²µ»É Îáµ»ÉÛ³Ý"
        dbCONTRACT.fDGCLICODE = "00000025"
        dbCONTRACT.fDGCUR = "000"
        dbCONTRACT.fDGSUMMA = "0.00"
        dbCONTRACT.fDGALLSUMMA = "0.00"
        dbCONTRACT.fDGRISKDEGREE = "0.00"
        dbCONTRACT.fDGRISKDEGNB = "0.00"
        dbCONTRACT.fDGSCHEDULE = "9"
        dbCONTRACT.fDGACSBRANCH = "00 "
        dbCONTRACT.fDGACSDEPART = "2"
        dbCONTRACT.fDGACSTYPE = "CO1 "
        dbCONTRACT.fDGCRDTCODE = "777000000250L001"
        
    Set dbFOLDERS(1) = New_DB_FOLDERS()
        dbFOLDERS(1).fFOLDERID = "C.1628339"
        dbFOLDERS(1).fNAME = "CardFee "
        dbFOLDERS(1).fKEY = fISN
        dbFOLDERS(1).fISN = fISN
        dbFOLDERS(1).fSTATUS = "1"
        dbFOLDERS(1).fCOM = "ø³ñï³ÛÇÝ í×³ñáõÙ"
        dbFOLDERS(1).fSPEC = "î»ëï³ÛÇÝ ù³ñï³ÛÇÝ í×³ñÙ³Ý ïÇå   ²ñø³ ¹»µ»ï   5,500.00- - Ð³ÛÏ³Ï³Ý ¹ñ³Ù"
        dbFOLDERS(1).fECOM = "Card Fee"

    Set dbFOLDERS(2) = New_DB_FOLDERS()
        dbFOLDERS(2).fFOLDERID = "CardFeeNew.20201120"
        dbFOLDERS(2).fNAME = "CardFee "
        dbFOLDERS(2).fKEY = fISN
        dbFOLDERS(2).fISN = fISN
        dbFOLDERS(2).fSTATUS = "0"
        dbFOLDERS(2).fCOM = "î»ëï³ÛÇÝ ù³ñï³ÛÇÝ í×³ñÙ³Ý ïÇå"
        dbFOLDERS(2).fSPEC = "905119030000004700001046983311            0.00         5500.00000            0.00    1²µ»É Îáµ»ÉÛ³Ý                                     20201120Üáñ                             002"
        dbFOLDERS(2).fECOM = "Test Cart payment type"
        dbFOLDERS(2).fDCBRANCH = "00"
        dbFOLDERS(2).fDCDEPART = "2"
        
    Set dbFOLDERS(3) = New_DB_FOLDERS()
        dbFOLDERS(3).fFOLDERID = "AGROVERLIM"
        dbFOLDERS(3).fNAME = "COSimpl "
        dbFOLDERS(3).fKEY = "01046983311"
        dbFOLDERS(3).fISN = fISN
        dbFOLDERS(3).fSTATUS = "1"
        dbFOLDERS(3).fCOM = "²µ»É Îáµ»ÉÛ³Ý"
        dbFOLDERS(3).fSPEC = "01046983311   CO1 0000"  
        dbFOLDERS(3).fDCBRANCH = "00 "
        dbFOLDERS(3).fDCDEPART = "2  " 
                
    Set dbFOLDERS(4) = New_DB_FOLDERS()
        dbFOLDERS(4).fFOLDERID = "Agr."&fISN
        dbFOLDERS(4).fNAME = "COSimpl "
        dbFOLDERS(4).fKEY = fISN
        dbFOLDERS(4).fISN = fISN
        dbFOLDERS(4).fSTATUS = "1"
        dbFOLDERS(4).fCOM = "¶»ñ³Í³Ëë ³ÝÅ³ÙÏ»ï"
        dbFOLDERS(4).fSPEC = "1¶»ñ³Í³Ëë ³ÝÅ³ÙÏ»ï- 01046983311 {²µ»É Îáµ»ÉÛ³Ý}"  
    
    Set dbFOLDERS(5) = New_DB_FOLDERS()
        dbFOLDERS(5).fFOLDERID = "C.1628339"
        dbFOLDERS(5).fNAME = "COSimpl "
        dbFOLDERS(5).fKEY = fISN
        dbFOLDERS(5).fISN = fISN
        dbFOLDERS(5).fSTATUS = "1"
        dbFOLDERS(5).fCOM = " ¶»ñ³Í³Ëë ³ÝÅ³ÙÏ»ï"
        dbFOLDERS(5).fSPEC = "01046983311 (²µ»É Îáµ»ÉÛ³Ý),     0 - Ð³ÛÏ³Ï³Ý ¹ñ³Ù"  
        dbFOLDERS(5).fECOM = "1"
        
    Set dbFOLDERS(6) = New_DB_FOLDERS()
        dbFOLDERS(6).fFOLDERID = "ALLACCSACC"
        dbFOLDERS(6).fNAME = "COAgrAcc"
        dbFOLDERS(6).fKEY = fISN
        dbFOLDERS(6).fISN = fISN2
        dbFOLDERS(6).fSTATUS = "1"
        dbFOLDERS(6).fCOM = "¶»ñ³Í³ËëÇ Ñ³ßí³å³Ñ³Ï³Ý Ñ³í»Éí³Í"
        dbFOLDERS(6).fSPEC = "001046983311                                            1"  
        
    Set dbFOLDERS(7) = New_DB_FOLDERS()
        dbFOLDERS(7).fFOLDERID = "ALLACCSGEN"
        dbFOLDERS(7).fNAME = "COAgrAcc"
        dbFOLDERS(7).fKEY = fISN
        dbFOLDERS(7).fISN = fISN2
        dbFOLDERS(7).fSTATUS = "1"
        dbFOLDERS(7).fCOM = "¶»ñ³Í³ËëÇ Ñ³ßí³å³Ñ³Ï³Ý Ñ³í»Éí³Í"
        dbFOLDERS(7).fSPEC = "01080793012"
        
    Set dbFOLDERS(8) = New_DB_FOLDERS()
        dbFOLDERS(8).fFOLDERID = "ALLACCSRES"
        dbFOLDERS(8).fNAME = "COAgrAcc"
        dbFOLDERS(8).fKEY = fISN
        dbFOLDERS(8).fISN = fISN2
        dbFOLDERS(8).fSTATUS = "1"
        dbFOLDERS(8).fCOM = "¶»ñ³Í³ËëÇ Ñ³ßí³å³Ñ³Ï³Ý Ñ³í»Éí³Í"
        dbFOLDERS(8).fSPEC = "000004532017303038100072112153000"
    
    Set dbFOLDERS(9) = New_DB_FOLDERS()
        dbFOLDERS(9).fFOLDERID = "Agr."& fISN
        dbFOLDERS(9).fNAME = "COAgrAcc"
        dbFOLDERS(9).fKEY = fISN2
        dbFOLDERS(9).fISN = fISN2
        dbFOLDERS(9).fSTATUS = "1"
        dbFOLDERS(9).fCOM = "¶»ñ³Í³ËëÇ Ñ³ßí³å³Ñ³Ï³Ý Ñ³í»Éí³Í"
        dbFOLDERS(9).fSPEC = "1¶»ñ³Í³ËëÇ Ñ³ßí³å³Ñ³Ï³Ý Ñ³í»Éí³Í- 01046983311   "
        
    Set dbFOLDERS(10) = New_DB_FOLDERS()
        dbFOLDERS(10).fFOLDERID = "CAGRACCS"
        dbFOLDERS(10).fNAME = "COAgrAcc"
        dbFOLDERS(10).fKEY = "01046983311   "
        dbFOLDERS(10).fISN = fISN2
        dbFOLDERS(10).fSTATUS = "1"
        dbFOLDERS(10).fCOM = "¶»ñ³Í³ËëÇ Ñ³ßí³å³Ñ³Ï³Ý Ñ³í»Éí³Í"
        
    Set dbFOLDERS(11) = New_DB_FOLDERS()
        dbFOLDERS(11).fFOLDERID = "AGRORDERS"
        dbFOLDERS(11).fNAME = "CODSDebt"
        dbFOLDERS(11).fKEY = fISN2
        dbFOLDERS(11).fISN = fISN2
        dbFOLDERS(11).fSTATUS = "1"
        dbFOLDERS(11).fCOM = "ä³ñïù»ñÇ Ù³ñÙ³Ý Ñ³Ûï"
        dbFOLDERS(11).fSPEC = "2020112001046983311       "
        
    Set dbFOLDERS(12) = New_DB_FOLDERS()
        dbFOLDERS(12).fFOLDERID = "Agr." & fISN
        dbFOLDERS(12).fNAME = "CODSDebt"
        dbFOLDERS(12).fKEY = fISN2
        dbFOLDERS(12).fISN = fISN2
        dbFOLDERS(12).fSTATUS = "1"
        dbFOLDERS(12).fCOM = "ä³ñïù»ñÇ Ù³ñÙ³Ý Ñ³Ûï"
        dbFOLDERS(12).fSPEC = "1ä³ñïù»ñÇ Ù³ñÙ³Ý Ñ³Ûï, ¶áõÙ³ñÁª 1368.4 -Ð³ÛÏ³Ï³Ý ¹ñ³Ù"   
     
    Set dbFOLDERS(13) = New_DB_FOLDERS()
        dbFOLDERS(13).fFOLDERID = "AgrOrd." & fISN
        dbFOLDERS(13).fNAME = "CODSDebt"
        dbFOLDERS(13).fKEY = fISN2
        dbFOLDERS(13).fISN = fISN2
        dbFOLDERS(13).fSTATUS = "1"
        dbFOLDERS(13).fCOM = "ä³ñïù»ñÇ Ù³ñÙ³Ý Ñ³Ûï"
        dbFOLDERS(13).fSPEC = ""   
    
    Set dbFOLDERS(14) = New_DB_FOLDERS()
        dbFOLDERS(14).fFOLDERID = "C.1628339"
        dbFOLDERS(14).fNAME = "CODSDebt"
        dbFOLDERS(14).fKEY = fISN2
        dbFOLDERS(14).fISN = fISN2
        dbFOLDERS(14).fSTATUS = "1"
        dbFOLDERS(14).fCOM = "ä³ñïù»ñÇ Ù³ñÙ³Ý Ñ³Ûï"
        dbFOLDERS(14).fSPEC = "ä³ÛÙ³Ý³·Çñ N% 01046983311, ¶áõÙ³ñÁª 1368.4 -Ð³ÛÏ³Ï³Ý ¹ñ³Ù"    
    
    Set dbFOLDERS(15) = New_DB_FOLDERS()
        dbFOLDERS(15).fFOLDERID = "ORDGOWAY"
        dbFOLDERS(15).fNAME = "CODSDebt"
        dbFOLDERS(15).fKEY = fISN2
        dbFOLDERS(15).fISN = fISN2
        dbFOLDERS(15).fSTATUS = "1"
        dbFOLDERS(15).fCOM = "ä³ñïù»ñÇ Ù³ñÙ³Ý Ñ³Ûï"
        dbFOLDERS(15).fSPEC = "20201120000²µ»É Îáµ»ÉÛ³Ý                   01046983311       00000025CO1 "   
        dbFOLDERS(15).fDCBRANCH = "00 "
        dbFOLDERS(15).fDCDEPART = "1"  
End Sub