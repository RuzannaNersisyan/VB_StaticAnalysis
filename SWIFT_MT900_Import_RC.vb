'USEUNIT SWIFT_International_Payorder_Library
'USEUNIT Library_Common
'USEUNIT Library_Colour
'USEUNIT Library_CheckDB
'USEUNIT Constants
'USEUNIT Library_Contracts
'USEUNIT Main_Accountant_Filter_Library 
'USEUNIT Payment_Except_Library
'USEUNIT Mem_Order_Library
'USEUNIT Currency_Exchange_Confirmphases_Library
'USEUNIT International_PayOrder_ConfirmPhases_Library
Option Explicit
Dim importPath, docN, savePath, pathAct, pathExp, workingDocs, currExchange , verifDoc, memOrder, convOrder, query95, forPayOrder, recieved
Dim dbSW_MESSAGES, dbFOLDERS(3), dbHI(4), dbCUREXCHANGES(0), i, sending
Dim sDATE, fDATE, fileFrom, fileTo, what, fWith, max, min, rand, isn, regex, fBODY, folderDirect, stDate, enDate, wUser, docType, sPath, pathE
Dim wName, passNum, cliCode,paySysIn, paySysOut, acsBranch,acsDepart, docISN, selectedView, expExcel                                         

'Դիտել փաստաթուղթը, Արտարժույթի փոխանակում, Հիշարար օրդեր
'Test case ID 185554
Sub SWIFT_900_RC_Test_1()
    
    Call Test_Initialize_SWIFT_900_RC()
    'Համակարգ մուտք գործել ARMSOFT օգտագործողով
    
    Call Initialize_AsBank("bank", sDATE, fDATE)
    Login("ARMSOFT")
    
'-----------------------------------------------------------------------------
'------ "S.W.I.F.T. ԱՇՏ/Պարամետրեր"-ում կատարել համապատասխան փոփոխությունները--------
'-----------------------------------------------------------------------------
    Log.Message "-- S.W.I.F.T. ԱՇՏ/Պարամետրեր-ում կատարել համապատասխան փոփոխությունները --",,,DivideColor
        
    'Մուտք գործել "S.W.I.F.T. ԱՇՏ"
    Call ChangeWorkspace(c_SWIFT)
    Call wTreeView.DblClickItem("|S.W.I.F.T. ²Þî                  |ä³ñ³Ù»ïñ»ñ")
    BuiltIn.Delay(3000)
    Randomize
    rand = Int((min-max+1)*Rnd+max)
    fileFrom = Project.Path &"Stores\SWIFT\HT900\ImportFile\HT000900_1.RJE"
    fileTo = Project.Path &"Stores\SWIFT\HT900\ImportFile\Import\HT000900_1.RJE"
    what = "CITIDEFFXXCITIGB2"
    fWith = "CITIDEFFXXCITIGB2" & rand
    'Ջնջում է Import թղթապանակի պարունակությունը
    aqFileSystem.DeleteFile(Project.Path &"Stores\SWIFT\HT900\ImportFile\Import\*")
    'Վերարտագրում է նախօրոք տրված ֆայլը մեկ այլ ֆայլի մեջ` կատարելով փոփոխություն
    Call Read_Write_File_With_replace(fileFrom,fileTo,what,fWith)
    
    'SWOUT պարամետրի փոփոխում
    Call SetParameter_InPttel("SWOUT" ,importPath)
    'SWSPFSACKDIR պարամետրի փոփոխում
    Call SetParameter("SWSPFSACKDIR", "")
    'SWSPFSNAKDIR պարամետրի փոփոխում
    Call SetParameter("SWSPFSNAKDIR", "")
    'SWSPFSOUT պարամետրի փոփոխում
    Call SetParameter("SWSPFSOUT", "")
    'Փակել պարամետրեր թղթապանակը
    Call Close_Window(wMDIClient, "frmPttel")
    Login("ARMSOFT")
    
'-----------------------------------------------------------------------------
'----------------- Կատարել Ընդունել SWIFT համակարգից գործողությունը ------------------
'-----------------------------------------------------------------------------
    Log.Message "Կատարել Ընդունել SWIFT համակարգից գործողությունը",,,DivideColor
    
    'Մուտք գործել "S.W.I.F.T. ԱՇՏ"
    Call ChangeWorkspace(c_SWIFT)
    Call Recieve_From_SWIFT (1)
    
'-------------------------------------------------------------------
'------- Անցում "Ստացված խառը հաղորդագրություններ" թղթապանակ --------------
'-------------------------------------------------------------------
    Log.Message "Անցում 'Ստացված խառը հաղորդագրություններ' թղթապանակ",,,DivideColor       
    'Մուտք գործել Փոխանցումներ/Ստացված փոխանցումներ թղթապանակ
    Call GoTo_Recieved_Messages (recieved, "|S.W.I.F.T. ²Þî                  |Ê³éÁ Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñ|êï³óí³Í Ë³éÁ Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñ")
    BuiltIn.Delay(4000) 
    'Ստանում է փաստաթղթի isn-ը
    If SearchInPttel("frmPttel", 3, docN) Then
        isn = GetIsn()
    Else 
        Log.Error "Document Row Not found",,,ErrorColor    
    End If
    Log.Message "fISN = "& isn,,,SqlDivideColor
'--------------------------------------------------------------------------------        
'---------Կատարել Դիտել փաստաթուղթը գործողությունը ներմուծված փաստաթղթի համար------------
'--------------------------------------------------------------------------------
    Log.Message "Դիտել փաստաթուղթը",,,DivideColor2
    regex="([[].{10}])|(\d{2}[/]\d{2}[/]\d{2})|(\d{2}:\d{2})|(CITIGB\d{3}LX)"
    sPath = savePath & "Actual\"
    pathE = savePath & "Expected\Expected_MT900.txt"
    If SearchInPttel("frmPttel", 3, docN) Then
        Call View_Doc_Action (sPath, "Actual_MT900.txt", pathE, regex)   
    Else 
        Log.Error "Document Row Not found",,,ErrorColor        
    End If
'--------------------------------------------------------------------------------        
'--------Կատարել Արտարժույթի փոխանակում գործողությունը ներմուծված փաստաթղթի համար-----------
'--------------------------------------------------------------------------------
    Log.Message "Արտարժույթի փոխանակում",,,DivideColor2   
    
    If SearchInPttel("frmPttel", 3, docN) Then
        Call wMainForm.MainMenu.Click(c_AllActions)
        Call wMainForm.PopupMenu.Click(c_CurExch)
        If wMDIClient.WaitVBObject("frmASDocForm",4000).Exists Then
            'Հաշիվ Դեբետ դաշտի լրացում
            Call Rekvizit_Fill ("Document",1,"General","ACCDB",currExchange.commonTab.dAcc)
            currExchange.commonTab.isn = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
            currExchange.commonTab.docN = Get_Rekvizit_Value("Document",currExchange.commonTab.tabN,"General","DOCNUM")
            Call ClickCmdButton(1, "Î³ï³ñ»É")
            BuiltIn.Delay(2000)
        Else
            Log.Error "Exchange Window not found",,,ErrorColor
        End If       
        'Փակել "Ստացված խառը հաղորդագրություններ" թղթապանակը
        Call Close_Window(wMDIClient, "frmPttel")
        'Ստուգել փաստաթղթի համապատասխանությունը օրինակի հետ
        If wMDIClient.WaitVBObject("FrmSpr",2000).Exists Then
            Call SaveDoc(savePath & "Actual\", "Actual_Currency_Exchange")
            pathAct = savePath & "Actual\Actual_Currency_Exchange.txt"
            pathExp = savePath & "Expected\Expected_Currency_Exchange.txt"
            regex="(N \d{6})|([[].{10}])|(\d{2}[/]\d{2}[/]\d{2} \d{2}:\d{2})"
            Call Compare_Files(pathAct, pathExp, regex)
            'Փակել տպելու ձևը
            Call Close_Window(wMDIClient, "FrmSpr" )
        Else 
            'Փակել "Ստացված խառը հաղորդագրություններ" թղթապանակը
            Log.Error "Can't find document print view",,,ErrorColor
        End If
    Else 
        Log.Error "Document Row Not found",,,ErrorColor        
    End If
    'Տպում է փաստաթղթի isn-ը ու համարը
    Log.Message "Exchange fISN = "& currExchange.commonTab.isn,,,SqlDivideColor
    Log.Message "Exchange docN = "& currExchange.commonTab.docN
    
    'SQL
    Log.Message "SQL Ստուգումներ Արտարժույթի փոխանակում գործողությունը կատարելուց հետո",,,SqlDivideColor 
    Call DB_Initialize_MT900_RC_1(currExchange.commonTab.isn, currExchange.commonTab.docN)
    'DOCLOG
    Call CheckQueryRowCount("DOCLOG","fISN",isn,3)
    Call CheckDB_DOCLOG(isn,"77","N","10"," ",1)
    Call CheckDB_DOCLOG(isn,"77","M","10","Received",1) 
    Call CheckDB_DOCLOG(isn,"77","M","10","CREATED",1) 
    
    Call CheckQueryRowCount("DOCLOG","fISN",currExchange.commonTab.isn,2)
    Call CheckDB_DOCLOG(currExchange.commonTab.isn,"77","N","1"," ",1)
    Call CheckDB_DOCLOG(currExchange.commonTab.isn,"77","C","5"," ",1)
    'DOCS
    fBODY = "  TYPECODE1:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28  TYPECODE2:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28  USERID:  77  "_
            &"ACSBRANCH:00  ACSDEPART:1  BLREP:0  DOCNUM:"&currExchange.commonTab.docN&"  DATE:"&aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%Y%m%d")&"  "_
            &"ACCDB:00067110103  ACCCR:10330030101  CURDB:003  CURCR:001  CASH:1  COURSE:     1.4189/    1  CRSNAME:001 / 003  SUMDB:363.66  "_
            &"SUMCR:516  CUPUSA:1  CURTES:1  CURVAIR:3  VOLORT:7  NONREZ:0  JURSTAT:21  AIM:Additional info  AIMINCLUDESPLACE:0  CBCRS1:450.0000/1  "_
            &"CBCRS2:400.0000/1  TIME:"&currExchange.addTab.fTime&"  TYPECODE3:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28  CURCOMIS:000  "_
            &"COMCBCRS:1/1  INCACCCOM:001453300  PAYSYSIN:5  INCACCCUREX:000931900  EXPACCCUREX:001434300  SENT2SW:0  RECFROMSW:0  FRSHCASHAC:0  "_
            &"CANCELREQ:0  CBCONFIRMD:0  CURIN:000  CLICODE:00000671  PAYREC:äáÕáëÛ³Ý  PAYRECLASTNAME:ê³ñ·Çë  PASSNUM:AA48536633  PASTYPE:01  "_
            &"PASBY:001  DATEPASS:20040831  REGNUM:1234567890  DATEBIRTH:20050117  CITIZENSHIP:1  COUNTRY:AM  USEOVERLIMIT:0  SYSCASE:SWDCCONF  "_
            &"NOTSENDABLECR:0  NOTSENDABLEDB:0  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",currExchange.commonTab.isn,1)
    Call CheckDB_DOCS(currExchange.commonTab.isn,"CurChng ","5",fBODY,1)
    'DOCP
    Call CheckQueryRowCount("DOCP","fISN",currExchange.commonTab.isn,1)
    Call CheckDB_DOCP(currExchange.commonTab.isn,"CurChng ",isn,1)
    'FODLERS
    Call CheckQueryRowCount("FOLDERS","fISN",currExchange.commonTab.isn,2)
    Call CheckDB_FOLDERS(dbFOLDERS(0),1) 
    Call CheckDB_FOLDERS(dbFOLDERS(1),1) 
    'HI
    Call CheckQueryRowCount("HI", "fBASE", currExchange.commonTab.isn, 4)
    Call Check_DB_HI(dbHI(0) ,1)
    Call Check_DB_HI(dbHI(1) ,1)
    Call Check_DB_HI(dbHI(2) ,1)
    Call Check_DB_HI(dbHI(3) ,1)
    'HIREST
    Call CheckDB_HIREST("11", "341169779" , "-3783780.00" ,"001", "-9459.45", 1)
    Call CheckDB_HIREST("01", "341169779" , "180039800.00" ,"001", "450099.50", 1)
    
    Call CheckDB_HIREST("01", "860427540" , "67500.00" ,"003", "150.00", 1)
    Call CheckDB_HIREST("11", "860427540" , "163647.00" ,"003", "363.66", 1)
    
    Call CheckDB_HIREST("11", "1629708" , "51234.10" ,"000", "51234.10", 1)
    Call CheckDB_HIREST("01", "1629708" , "4056711.00" ,"000", "4056711.00", 1)

'--------------------------------------------------------------------------------        
'-------------Կատարել Վավերացնել գործողությունը Արտարժույթի փոխանակման համար---------------
'--------------------------------------------------------------------------------
    Log.Message "Արտարժույթի փոխանակման հաստատում",,,DivideColor    
    
    'Մուտք Գլխավոր Հաշվապահի ԱՇՏ/Աշխատանքային փաստաթղթեր
    Log.Message "Անցում Գլխավոր Հաշվապահի ԱՇՏ",,,DivideColor
    
    Call ChangeWorkspace(c_ChiefAcc) 
    Call GoTo_MainAccWorkingDocuments("|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|", workingDocs) 
    BuiltIn.Delay(3000)
    If SearchInPttel ("frmPttel", 2, currExchange.commonTab.docN) Then
        Call wMainForm.MainMenu.Click(c_AllActions)
        Call wMainForm.PopupMenu.Click(c_ToConfirm)
        If wMDIClient.WaitVBObject("frmASDocForm",4000).Exists Then 
            'Ստուգել պատուհանում լրացված տվյալների ճշտությունը
            Call Currency_Exchange_Check(currExchange)
            Call ClickCmdButton(1, "Ð³ëï³ï»É")
        Else    
            Log.Error "Currency Exchange window not found",,,ErrorColor
        End If
    Else 
        Log.Error "Document Row Not found",,,ErrorColor                  
    End If
    
    'SQL
    Log.Message "SQL Ստուգումներ Արտարժույթի փոխանակումը վավերացնելուց հետո հետո",,,SqlDivideColor 
    'DOCLOG
    Call CheckQueryRowCount("DOCLOG","fISN",currExchange.commonTab.isn,4)
    Call CheckDB_DOCLOG(currExchange.commonTab.isn,"77","N","1"," ",1)
    Call CheckDB_DOCLOG(currExchange.commonTab.isn,"77","C","5"," ",1)
    Call CheckDB_DOCLOG(currExchange.commonTab.isn,"77","W","6"," ",1)
    Call CheckDB_DOCLOG(currExchange.commonTab.isn,"77","C","101"," ",1)
    'DOCS
    Call CheckQueryRowCount("DOCS","fISN",currExchange.commonTab.isn,1)
    Call CheckDB_DOCS(currExchange.commonTab.isn,"CurChng ","101",fBODY,1)
    'FODLERS
    For i = 0 to 1
        dbFOLDERS(i).fSTATUS = "0"
    Next
    dbFOLDERS(0).fSPEC = Replace (dbFOLDERS(0).fSPEC , "Üáñ" , "àõÕ³ñÏí³Í I Ñ³ëï³ïÙ³Ý")
    dbFOLDERS(1).fSPEC = Replace (dbFOLDERS(1).fSPEC , "Üáñ                  " , "àõÕ³ñÏí³Í I Ñ³ëï³ïÙ³Ý")
    Call CheckQueryRowCount("FOLDERS","fISN",currExchange.commonTab.isn,3)
    Call CheckDB_FOLDERS(dbFOLDERS(0),1) 
    Call CheckDB_FOLDERS(dbFOLDERS(1),1)    
    Call CheckDB_FOLDERS(dbFOLDERS(2),1)
    
    'Փակել "Աշխատանքային փաստաթղթեր" թղթապանակը
    Call Close_Window(wMDIClient, "frmPttel") 
    
    'Մուտք Հաստատվող փաստաթղթեր թղթապանակ
    Call GoToVerificationDocument("|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ (I)",verifDoc)
    BuiltIn.Delay(2000)
    If SearchInPttel("frmPttel", 3, currExchange.commonTab.docN) Then
        'Վավերացնել Արտարժույթի փոխանակումը
        Call Confirm_Document()
    Else 
        Log.Error "Document Row Not found",,,ErrorColor        
    End If
    
    'SQL
    Log.Message "SQL Ստուգումներ Արտարժույթի փոխանակումը Հաստատվող փաստաթղթեր թղթապանակից վավերացնելուց հետո",,,SqlDivideColor 
    'DOCLOG
    Call CheckQueryRowCount("DOCLOG","fISN",currExchange.commonTab.isn,6)
    Call CheckDB_DOCLOG(currExchange.commonTab.isn,"77","N","1"," ",1)
    Call CheckDB_DOCLOG(currExchange.commonTab.isn,"77","C","5"," ",1)
    Call CheckDB_DOCLOG(currExchange.commonTab.isn,"77","W","6"," ",1)
    Call CheckDB_DOCLOG(currExchange.commonTab.isn,"77","C","101"," ",1)
    Call CheckDB_DOCLOG(currExchange.commonTab.isn,"77","W","102"," ",1)
    Call CheckDB_DOCLOG(currExchange.commonTab.isn,"77","C","101"," ",1)    
    'DOCS
    fBODY = "  OPERTYPE:CEX  TYPECODE1:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28  TYPECODE2:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28"_
            &"  USERID:  77  ACSBRANCH:00  ACSDEPART:1  BLREP:0  DOCNUM:"&currExchange.commonTab.docN&"  "_
            &"DATE:"&aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%Y%m%d")&"  ACCDB:00067110103  ACCCR:10330030101  CURDB:003  CURCR:001  CASH:1  "_
            &"COURSE:     1.4189/    1  CRSNAME:001 / 003  SUMDB:363.66  SUMCR:516  CUPUSA:1  CURTES:1  CURVAIR:3  VOLORT:7  NONREZ:0  JURSTAT:21  "_
            &"AIM:Additional info  AIMINCLUDESPLACE:0  CBCRS1:450.0000/1  CBCRS2:400.0000/1  TIME:"&currExchange.addTab.fTime&"  TYPECODE3:-20 21 22"_
            &" 23 24 30 31 32 25 26 92 93 11 27 33 28  CURCOMIS:000  COMCBCRS:1/1  INCACCCOM:001453300  PAYSYSIN:5  INCACCCUREX:000931900  "_
            &"EXPACCCUREX:001434300  SENT2SW:0  RECFROMSW:0  FRSHCASHAC:0  CANCELREQ:0  CBCONFIRMD:0  CURIN:000  CLICODE:00000671  PAYREC:äáÕáëÛ³Ý  "_
            &"PAYRECLASTNAME:ê³ñ·Çë  PASSNUM:AA48536633  PASTYPE:01  PASBY:001  DATEPASS:20040831  REGNUM:1234567890  DATEBIRTH:20050117  "_
            &"CITIZENSHIP:1  COUNTRY:AM  USEOVERLIMIT:0  SYSCASE:SWDCCONF  NOTSENDABLECR:0  NOTSENDABLEDB:0  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",currExchange.commonTab.isn,1)
    Call CheckDB_DOCS(currExchange.commonTab.isn,"CurChng ","11",fBODY,1)
    'FODLERS
    Call CheckQueryRowCount("FOLDERS","fISN",currExchange.commonTab.isn,0)
    'CUREXCHANGES
    Call CheckQueryRowCount("CUREXCHANGES","fISN",currExchange.commonTab.isn,1)
    Call CheckDB_CUREXCHANGES(dbCUREXCHANGES(0),1)
    'HI
    For i = 0 to 4
        With dbHI(i)
            .fTYPE = "01"
            .fBASEBRANCH = "00 "
            .fBASEDEPART = "1  "
        End With
    Next 
    dbHI(4).fTYPE = "CE"           
    Call CheckQueryRowCount("HI", "fBASE", currExchange.commonTab.isn, 5)
    Call Check_DB_HI(dbHI(0) ,1)
    Call Check_DB_HI(dbHI(1) ,1)
    Call Check_DB_HI(dbHI(2) ,1)
    Call Check_DB_HI(dbHI(3) ,1)
    Call Check_DB_HI(dbHI(4) ,1)
    
    'HIREST
    Call CheckDB_HIREST("11", "341169779" , "-3577380.00" ,"001", "-8943.45", 1)
    Call CheckDB_HIREST("01", "341169779" , "179833400.00" ,"001", "449583.50", 1)
    
    Call CheckDB_HIREST("01", "860427540" , "231147.00" ,"003", "513.66", 1)
    Call CheckDB_HIREST("11", "860427540" , "0.00" ,"003", "0.00", 2)
    
    Call CheckDB_HIREST("11", "1629708" , "8481.10" ,"000", "8481.10", 1)
    Call CheckDB_HIREST("01", "1629708" , "4099464.00" ,"000", "4099464.00", 1)
    
    'Փակել "Հաստատվող փաստաթղթեր" թղթապանակը
    Call Close_Window(wMDIClient, "frmPttel")
    
    'Մուտք Հաշվառված վճարային փաստաթղթեր
    Log.Message  "Մուտք Հաշվառված վճարային փաստաթղթեր"
    folderDirect = "|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|Ð³ßí³éí³Í í×³ñ³ÛÇÝ ÷³ëï³ÃÕÃ»ñ"
    stDate = aqDateTime.Today
    enDate = aqDateTime.Today
    wUser = "77"
    docType = "CurChng"
    wName = "" 
    passNum = ""
    cliCode = ""
    paySysIn = ""
    paySysOut = ""
    acsBranch = ""
    acsDepart = ""
    docISN = ""
    selectedView = "Payments"
    expExcel = "0"
    Call OpenAccPaymentDocFolder(folderDirect, stDate, enDate, wUser, docType,wName, passNum, cliCode,paySysIn, paySysOut, acsBranch,_
                                               acsDepart, docISN, selectedView, expExcel)
    BuiltIn.Delay(2000)
    'Ջնջել Արտարժույթի փոխանակումը
    Call SearchAndDelete ( "frmPttel", 2, currExchange.commonTab.docN, "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ")
    'Փակել "Հաշվառված վճարային փաստաթղթեր" թղթապանակը
    Call Close_Window(wMDIClient, "frmPttel")
    
    'SQL
    Log.Message "SQL Ստուգումներ Արտարժույթի փոխանակումը ջնջելուց հետո",,,SqlDivideColor 
    'DOCLOG
    Call CheckQueryRowCount("DOCLOG","fISN",currExchange.commonTab.isn,7)
    Call CheckDB_DOCLOG(currExchange.commonTab.isn,"77","N","1"," ",1)
    Call CheckDB_DOCLOG(currExchange.commonTab.isn,"77","C","5"," ",1)
    Call CheckDB_DOCLOG(currExchange.commonTab.isn,"77","W","6"," ",1)
    Call CheckDB_DOCLOG(currExchange.commonTab.isn,"77","C","101"," ",1)
    Call CheckDB_DOCLOG(currExchange.commonTab.isn,"77","W","102"," ",1)
    Call CheckDB_DOCLOG(currExchange.commonTab.isn,"77","M","11","¶ñ³Ýóí»É »Ý Ó¨³Ï»ñåáõÙÝ»ñÁ",1)
    Call CheckDB_DOCLOG(currExchange.commonTab.isn,"77","D","999"," ",1)
    'DOCS
    Call CheckQueryRowCount("DOCS","fISN",currExchange.commonTab.isn,1)
    Call CheckDB_DOCS(currExchange.commonTab.isn,"CurChng ","999",fBODY,1)   
    'HI
    Call CheckQueryRowCount("HI", "fBASE", currExchange.commonTab.isn, 0)
    'HIREST
    Call CheckDB_HIREST("01", "341169779" , "180039800.00" ,"001", "450099.50", 1)
    
    Call CheckDB_HIREST("01", "860427540" , "67500.00" ,"003", "150.00", 1)
    
    Call CheckDB_HIREST("01", "1629708" , "4056711.00" ,"000", "4056711.00", 1)
    
    'Անցնել "S.W.I.F.T. ԱՇՏ"
    Call ChangeWorkspace(c_SWIFT)

'-------------------------------------------------------------------
'------- Անցում "Ստացված խառը հաղորդագրություններ" թղթապանակ --------------
'-------------------------------------------------------------------
    Log.Message "Անցում 'Ստացված խառը հաղորդագրություններ' թղթապանակ",,,DivideColor       
    Call GoTo_Recieved_Messages (recieved, "|S.W.I.F.T. ²Þî                  |Ê³éÁ Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñ|êï³óí³Í Ë³éÁ Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñ")

'--------------------------------------------------------------------------------        
'-----------Կատարել Հիշարար օրդեր գործողությունը ներմուծված փաստաթղթի համար----------------
'--------------------------------------------------------------------------------
    Log.Message "Հիշարար օրդեր",,,DivideColor2        
    If SearchInPttel("frmPttel", 3, docN) Then
        Call wMainForm.MainMenu.Click(c_AllActions)
        Call wMainForm.PopupMenu.Click(c_MemOrd)
        If wMDIClient.WaitVBObject("frmASDocForm",2000).Exists Then 
            memOrder.DocN = Get_Rekvizit_Value("Document",1,"General","DOCNUM")
            memOrder.Isn = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
            'Լրացնել Հաշիվ Դեբետ դաշտը
            Call Rekvizit_Fill ("Document",1,"General","ACCDB",memOrder.AccD)
            Call ClickCmdButton(1, "Î³ï³ñ»É")  
        Else    
            Log.Error "memOrder window not found",,,ErrorColor
        End If
        'Ստուգել փաստաթղթի համապատասխանությունը օրինակի հետ
        If wMDIClient.WaitVBObject("FrmSpr",2000).Exists Then
            Call SaveDoc(savePath & "Actual\", "Actual_memOrder")
            pathAct = savePath & "Actual\Actual_memOrder.txt"
            pathExp = savePath & "Expected\Expected_memOrder.txt"
            regex="(N \d{6})|([[].*])|(\d{2}[/]\d{2}[/]\d{2} \d{2}:\d{2})"
            Call Compare_Files(pathAct, pathExp, regex)
            'Փակել տպելու ձևը
            Call Close_Window(wMDIClient, "FrmSpr" )
        Else
            Log.Error "Can't find document print view",,,ErrorColor
        End If
        'Տպում է փաստաթղթի isn-ը ու համարը
        Log.Message "memOrder fISN = "& memOrder.Isn ,,,SqlDivideColor
        Log.Message "memOrder docN = "& memOrder.DocN
    Else 
        Log.Error "Document Row Not found",,,ErrorColor
    End If
    'Փակել "Ստացված խառը հաղորդագրություններ" թղթապանակը
    Call Close_Window(wMDIClient, "frmPttel")
    'SQL
    Log.Message "SQL Ստուգումներ Հիշարար օրդեր ստեղծելուց հետո",,,SqlDivideColor 
    Call DB_Initialize_MT900_RC_2(memOrder.Isn, memOrder.DocN)
    'DOCLOG
    Call CheckQueryRowCount("DOCLOG","fISN",memOrder.Isn,2)
    Call CheckDB_DOCLOG(memOrder.Isn,"77","N","1"," ",1)
    Call CheckDB_DOCLOG(memOrder.Isn,"77","C","10"," ",1)
    'DOCP
    Call CheckQueryRowCount("DOCP","fISN",memOrder.Isn,1)
    Call CheckDB_DOCP(memOrder.Isn,"MemOrd  ",isn,1)
    'DOCS
    fBODY = "  TYPECODE1:-10 20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28  TYPECODE2:-10 20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28  "_
            &"USERID:  77  ACSBRANCH:00  ACSDEPART:1  DOCNUM:"&memOrder.DocN&"  DATE:"&aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%Y%m%d")&"  "_
            &"ACCDB:000048201  ACCCR:10330030101  CUR:001  SUMMA:516  AIM:Additional info  PAYSYSIN:5  SYSCASE:SWDCCONF  USEOVERLIMIT:0  "_
            &"NOTSENDABLECR:0  NOTSENDABLEDB:0  SBQENABLED:0  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",memOrder.Isn,1)
    Call CheckDB_DOCS(memOrder.Isn,"MemOrd  ","10",fBODY,1)
    'FODLERS
    Call CheckQueryRowCount("FOLDERS","fISN",memOrder.Isn,1)
    Call CheckDB_FOLDERS(dbFOLDERS(0),1) 
    'HI
    Call CheckQueryRowCount("HI", "fBASE", memOrder.Isn, 2)
    Call Check_DB_HI(dbHI(0) ,1)
    Call Check_DB_HI(dbHI(1) ,1)
    'HIREST
    Call CheckDB_HIREST("11", "341169779" , "-3783780.00" ,"001", "-9459.45", 1)
    Call CheckDB_HIREST("01", "341169779" , "180039800.00" ,"001", "450099.50", 1)
    
    Call CheckDB_HIREST("01", "1714456" , "38081.10" ,"001", "90.73", 1)
    Call CheckDB_HIREST("11", "1714456" , "206400.00" ,"001", "516.00", 1)
    
'--------Մուտք Գլխավոր Հաշվապահի ԱՇՏ/Աշխատանքային փաստաթղթեր---------------------------------
    Log.Message "Անցում Գլխավոր Հաշվապահի ԱՇՏ",,,DivideColor
    
    Call ChangeWorkspace(c_ChiefAcc) 
    Call GoTo_MainAccWorkingDocuments("|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|", workingDocs) 
    BuiltIn.Delay(3000)
    Log.Message "Հիշարար օրդերի հաշվառում"
    If SearchInPttel("frmPttel", 2, memOrder.DocN) Then
       'Կատարել Դիտել գործողությունը հիշարար օրդերի համար
       Call View_memOrder (memOrder,"frmPttel")
       BuiltIn.Delay(2000)
       'Հաշվառել Հիշարար օրդերը
       Call Register_Payment()
    End If
    
    'SQL
    Log.Message "SQL Ստուգումներ Հիշարար օրդերը հաշվառելուց հետո",,,SqlDivideColor 
    'DOCLOG
    Call CheckQueryRowCount("DOCLOG","fISN",memOrder.Isn,3)
    Call CheckDB_DOCLOG(memOrder.Isn,"77","N","1"," ",1)
    Call CheckDB_DOCLOG(memOrder.Isn,"77","C","10"," ",1)
    Call CheckDB_DOCLOG(memOrder.Isn,"77","M","5","¶ñ³Ýóí»É »Ý Ó¨³Ï»ñåáõÙÝ»ñÁ",1)
    'DOCS
    fBODY = "  OPERTYPE:MSC  TYPECODE1:-10 20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28  TYPECODE2:-10 20 21 22 23 24 30 31 32 25 26 92 93 11 27"_
            &" 33 28  USERID:  77  ACSBRANCH:00  ACSDEPART:1  DOCNUM:"&memOrder.DocN&"  DATE:"&aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%Y%m%d")&"  "_
            &"ACCDB:000048201  ACCCR:10330030101  CUR:001  SUMMA:516  AIM:Additional info  PAYSYSIN:5  SYSCASE:SWDCCONF  USEOVERLIMIT:0  "_
            &"NOTSENDABLECR:0  NOTSENDABLEDB:0  SBQENABLED:0  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",memOrder.Isn,1)
    Call CheckDB_DOCS(memOrder.Isn,"MemOrd  ","5",fBODY,1)
    'HIREST
    Call CheckDB_HIREST("01", "341169779" , "179833400.00" ,"001", "449583.50", 1)
    Call CheckDB_HIREST("01", "1714456" , "244481.10" ,"001", "606.73", 1)
    'memOrderS
    Call CheckQueryRowCount("MEMORDERS","fISN",memOrder.Isn,1)
    Call CheckDB_memOrderS(memOrder.Isn,"MemOrd  ","1",aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%Y%m%d"),"5","516.00","001",1)
    
    'Փակել "Աշխատանքային փաստաթղթեր" թղթապանակը
    Call Close_Window(wMDIClient, "frmPttel")
    'Մուտք Հաշվառված վճարային փաստաթղթեր
    Log.Message  "Մուտք Հաշվառված վճարային փաստաթղթեր"
    docType = "MemOrd "
    Call OpenAccPaymentDocFolder(folderDirect, stDate, enDate, wUser, docType,wName, passNum, cliCode,paySysIn, paySysOut, acsBranch,_
                                               acsDepart, docISN, selectedView, expExcel)
    BuiltIn.Delay(2000)
    'Ջնջել Հիշարար օրդերը
    Call SearchAndDelete ( "frmPttel", 2, memOrder.DocN, "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ")
    'Փակել "Հաշվառված վճարային փաստաթղթեր" թղթապանակը
    Call Close_Window(wMDIClient, "frmPttel")
    
    'SQL
    Log.Message "SQL Ստուգումներ Հիշարար օրդերը ջնջելուց հետո",,,SqlDivideColor 
    'DOCLOG
    Call CheckQueryRowCount("DOCLOG","fISN",memOrder.Isn,4)
    Call CheckDB_DOCLOG(memOrder.Isn,"77","N","1"," ",1)
    Call CheckDB_DOCLOG(memOrder.Isn,"77","C","10"," ",1)
    Call CheckDB_DOCLOG(memOrder.Isn,"77","M","5","¶ñ³Ýóí»É »Ý Ó¨³Ï»ñåáõÙÝ»ñÁ",1)
    Call CheckDB_DOCLOG(memOrder.Isn,"77","D","999"," ",1)
    'DOCP 
    Call CheckQueryRowCount("DOCP","fISN",memOrder.Isn,0)
    'HIREST
    Call CheckDB_HIREST("01", "341169779" , "180039800.00" ,"001", "450099.50", 1)
    Call CheckDB_HIREST("01", "1714456" , "38081.10" ,"001", "90.73", 1)
    
'------------------------------------------------------------------
'----------------- Ջնջում է ներմուծված փաստաթուղթը ---------------------
'------------------------------------------------------------------     
    'Մուտք Ադմինիստրարտոր ԱՇՏ
    Call ChangeWorkspace(c_Admin40)
    folderDirect = "|²¹ÙÇÝÇëïñ³ïáñÇ ²Þî 4.0|ÂÕÃ³å³Ý³ÏÝ»ñ|êï»ÕÍí³Í ÷³ëï³ÃÕÃ»ñ"
    stDate = aqDateTime.Today
    enDate = aqDateTime.Today
    wUser = 77
    docType = ""
    'Մուտք Ստեղծված փաստաթղթեր թղթապանակ
    Call OpenCreatedDocFolder(folderDirect, stDate, enDate, wUser, docType)
    
    'Ջնջում է փաստաթղթի հետ կապակցված Տարանցիկ վճարային փոխանցման փաստաթուղթը
    Call SearchAndDelete( "frmPttel", 2, isn, "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ" ) 
    Call Close_Window(wMDIClient, "frmPttel")
    Call Close_AsBank()    
End Sub

'Փոխանակման օրդեր, Ճշտել (n95), Վճարման հանձնարարագիր
'Test case ID 187210
Sub SWIFT_900_RC_Test_2()
    
    Call Test_Initialize_SWIFT_900_RC()
    'Համակարգ մուտք գործել ARMSOFT օգտագործողով
    Call Initialize_AsBank("bank", sDATE, fDATE)
    Login("ARMSOFT")
    
'-----------------------------------------------------------------------------
'------ "S.W.I.F.T. ԱՇՏ/Պարամետրեր"-ում կատարել համապատասխան փոփոխությունները--------
'-----------------------------------------------------------------------------
    Log.Message "-- S.W.I.F.T. ԱՇՏ/Պարամետրեր-ում կատարել համապատասխան փոփոխությունները --",,,DivideColor
        
    'Մուտք գործել "S.W.I.F.T. ԱՇՏ"
    Call ChangeWorkspace(c_SWIFT)
    Call wTreeView.DblClickItem("|S.W.I.F.T. ²Þî                  |ä³ñ³Ù»ïñ»ñ")
    BuiltIn.Delay(3000)

    Randomize
    rand = Int((min-max+1)*Rnd+max)
    fileFrom = Project.Path &"Stores\SWIFT\HT900\ImportFile\HT000900_1.RJE"
    fileTo = Project.Path &"Stores\SWIFT\HT900\ImportFile\Import\HT000900_1.RJE"
    what = "CITIDEFFXXCITIGB2"
    fWith = "CITIDEFFXXCITIGB2" & rand
    'Ջնջում է Import թղթապանակի պարունակությունը
    aqFileSystem.DeleteFile(Project.Path &"Stores\SWIFT\HT900\ImportFile\Import\*")
    'Վերարտագրում է նախօրոք տրված ֆայլը մեկ այլ ֆայլի մեջ` կատարելով փոփոխություն
    Call Read_Write_File_With_replace(fileFrom,fileTo,what,fWith)
    
    'SWOUT պարամետրի փոփոխում
    Call SetParameter_InPttel("SWOUT" ,importPath)
    'SWSPFSACKDIR պարամետրի փոփոխում
    Call SetParameter("SWSPFSACKDIR", "")
    'SWSPFSNAKDIR պարամետրի փոփոխում
    Call SetParameter("SWSPFSNAKDIR", "")
    'SWSPFSOUT պարամետրի փոփոխում
    Call SetParameter("SWSPFSOUT", "")
    'Փակել պարամետրեր թղթապանակը
    Call Close_Window(wMDIClient, "frmPttel")
    Login("ARMSOFT")
    
'-----------------------------------------------------------------------------
'----------------- Կատարել Ընդունել SWIFT համակարգից գործողությունը ------------------
'-----------------------------------------------------------------------------
    Log.Message "Կատարել Ընդունել SWIFT համակարգից գործողությունը",,,DivideColor
    
    'Մուտք գործել "S.W.I.F.T. ԱՇՏ"
    Call ChangeWorkspace(c_SWIFT)
    Call Recieve_From_SWIFT (1)
    
'-------------------------------------------------------------------
'------- Անցում "Ստացված խառը հաղորդագրություններ" թղթապանակ --------------
'-------------------------------------------------------------------
    Log.Message "Անցում 'Ստացված խառը հաղորդագրություններ' թղթապանակ",,,DivideColor       
    'Մուտք գործել Փոխանցումներ/Ստացված փոխանցումներ թղթապանակ
    Call GoTo_Recieved_Messages (recieved, "|S.W.I.F.T. ²Þî                  |Ê³éÁ Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñ|êï³óí³Í Ë³éÁ Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñ")
    BuiltIn.Delay(4000) 
    'Ստանում է փաստաթղթի isn-ը
    If SearchInPttel("frmPttel", 3, docN) Then
        isn = GetIsn()
    Else 
        Log.Error "Document Row Not found",,,ErrorColor    
    End If
    Log.Message "fISN = "& isn,,,SqlDivideColor
   
'--------------------------------------------------------------------------------        
'---------Կատարել Փոխանակման օրդեր գործողությունը ներմուծված փաստաթղթի համար--------------
'--------------------------------------------------------------------------------
    Log.Message "Փոխանակման օրդեր",,,DivideColor2       
    
    If SearchInPttel("frmPttel", 3, docN) Then
        Call wMainForm.MainMenu.Click(c_AllActions)
        Call wMainForm.PopupMenu.Click(c_ConvOrd)
        If wMDIClient.WaitVBObject("frmASDocForm",4000).Exists Then 
            convOrder.commonTab.isn = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
            convOrder.commonTab.docN = Get_Rekvizit_Value("Document",convOrder.commonTab.tabN,"General","DOCNUM")       
            'Լրացնել Հաշիվ Դեբետ դաշտը
            Call Rekvizit_Fill ("Document",1,"General","ACCDB",convOrder.commonTab.dAcc)
            Call ClickCmdButton(1, "Î³ï³ñ»É")
        Else    
            Log.Error "Conversion Order window not found",,,ErrorColor
        End If    
        'Ստուգել փաստաթղթի համապատասխանությունը օրինակի հետ
        If wMDIClient.WaitVBObject("FrmSpr",2000).Exists Then
            Call SaveDoc(savePath & "Actual\", "Actual_Conversion_Order")
            pathAct = savePath & "Actual\Actual_Conversion_Order.txt"
            pathExp = savePath & "Expected\Expected_Conversion_Order.txt"
            regex="(N \d{6})|([[].*])|(\d{2}[/]\d{2}[/]\d{2} \d{2}:\d{2})"
            Call Compare_Files(pathAct, pathExp, regex)
            'Փակել տպելու ձևը
            Call Close_Window(wMDIClient, "FrmSpr" )
        Else
            Log.Error "Can't find document print view",,,ErrorColor
        End If
        'Տպում է փաստաթղթի isn-ը ու համարը
        Log.Message "Conversion Order fISN = "& convOrder.commonTab.isn ,,,SqlDivideColor
        Log.Message "Conversion Order docN = "& convOrder.commonTab.docN  

    Else 
        Log.Error "Document Row Not found",,,ErrorColor
    End If
    'Փակել "Ստացված խառը հաղորդագրություններ" թղթապանակը
    Call Close_Window(wMDIClient, "frmPttel")     
    
    'SQL
    Log.Message "SQL Ստուգումներ Փոխանակման օրդեր ստեղծելուց հետո",,,SqlDivideColor  
    Call DB_Initialize_MT900_RC_3(convOrder.commonTab.isn, convOrder.commonTab.docN)  
    'DOCLOG
    Call CheckQueryRowCount("DOCLOG","fISN",convOrder.commonTab.isn,2)
    Call CheckDB_DOCLOG(convOrder.commonTab.isn,"77","N","1"," ",1)
    Call CheckDB_DOCLOG(convOrder.commonTab.isn,"77","C","10"," ",1) 
    'DOCP
    Call CheckQueryRowCount("DOCP","fISN",convOrder.commonTab.isn,1)
    Call CheckDB_DOCP(convOrder.commonTab.isn,"ConvOrd ",isn,1)
    'DOCS
    fBODY = "  TYPECODE1:-10 20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28  TYPECODE2:-10 20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28  "_
            &"USERID:  77  ACSBRANCH:00  ACSDEPART:1  DOCNUM:"&convOrder.commonTab.docN&"  DATE:"&aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%Y%m%d")&"  "_
            &"ACCDB:000084700  ACCCR:10330030101  CURDB:000  CURCR:001  CBCRS:400.0000/1  CRSNAME:000 / 001  SUMDB:206400  SUMCR:516  PAYSYSIN:5  "_
            &"AIM:Additional info  CUPUSA:2  TIME:"&convOrder.psTab.fTime&"  CURTES:1  CURVAIR:3  NONREZ:0  SYSCASE:SWDCCONF  USEOVERLIMIT:0  "_
            &"NOTSENDABLECR:0  NOTSENDABLEDB:0  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",convOrder.commonTab.isn,1)
    Call CheckDB_DOCS(convOrder.commonTab.isn,"ConvOrd ","10",fBODY,1)
    'FODLERS
    Call CheckQueryRowCount("FOLDERS","fISN",convOrder.commonTab.isn,1)
    Call CheckDB_FOLDERS(dbFOLDERS(0),1) 
    'HI
    Call CheckQueryRowCount("HI", "fBASE", convOrder.commonTab.isn, 2)
    Call Check_DB_HI(dbHI(0) ,1)
    Call Check_DB_HI(dbHI(1) ,1)
    'HIREST
    Call CheckDB_HIREST("11", "341169779" , "-3783780.00" ,"001", "-9459.45", 1)
    Call CheckDB_HIREST("01", "341169779" , "180039800.00" ,"001", "450099.50", 1)
    
    Call CheckDB_HIREST("11", "1630226" , "206400.00" ,"000", "206400.00", 1)
    
    
'-----Մուտք Գլխավոր Հաշվապահի ԱՇՏ/Աշխատանքային փաստաթղթեր------------------------------------
    Log.Message "Մուտք Գլխավոր Հաշվապահի ԱՇՏ",,,DivideColor
    Call ChangeWorkspace(c_ChiefAcc) 
    Call GoTo_MainAccWorkingDocuments("|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|", workingDocs) 
    BuiltIn.Delay(3000)
    If SearchInPttel("frmPttel", 2, convOrder.commonTab.docN) Then
        'Կատարել Դիտել գործողությունը Փոխանակման օրդերի համար
        Call wMainForm.MainMenu.Click(c_AllActions)
        Call wMainForm.PopupMenu.Click(c_View)       
        If wMDIClient.WaitVBObject("frmASDocForm",2000).Exists Then
            Call Conversion_Order_Check(convOrder)
            Call ClickCmdButton(1,"OK")     
        Else    
            Log.Error "Conversion Order window not found",,,ErrorColor
        End If
        BuiltIn.Delay(2000)
        'Հաշվառել Փոխանակման օրդերը
        Call Register_Payment()
    End If
    'Փակել "Աշխատանքային փաստաթղթեր" թղթապանակը
    Call Close_Window(wMDIClient, "frmPttel")
    
    'SQL
    Log.Message "SQL Ստուգումներ Փոխանակման օրդերը հաշվառելուց հետո",,,SqlDivideColor  
    'DOCLOG
    Call CheckQueryRowCount("DOCLOG","fISN",convOrder.commonTab.isn,3)
    Call CheckDB_DOCLOG(convOrder.commonTab.isn,"77","N","1"," ",1)
    Call CheckDB_DOCLOG(convOrder.commonTab.isn,"77","C","10"," ",1) 
    Call CheckDB_DOCLOG(convOrder.commonTab.isn,"77","M","2","¶ñ³Ýóí»É »Ý Ó¨³Ï»ñåáõÙÝ»ñÁ",1)
    'DOCS
    fBODY = "  TYPECODE1:-10 20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28  TYPECODE2:-10 20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28  "_
            &"USERID:  77  OPERTYPE:CEX  ACSBRANCH:00  ACSDEPART:1  DOCNUM:"&convOrder.commonTab.docN&"  "_
            &"DATE:"&aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%Y%m%d")&"  ACCDB:000084700  ACCCR:10330030101  CURDB:000  CURCR:001  "_
            &"CBCRS:400.0000/1  CRSNAME:000 / 001  SUMDB:206400  SUMCR:516  PAYSYSIN:5  AIM:Additional info  CUPUSA:2  TIME:"&convOrder.psTab.fTime&"  "_
            &"CURTES:1  CURVAIR:3  NONREZ:0  SYSCASE:SWDCCONF  USEOVERLIMIT:0  NOTSENDABLECR:0  NOTSENDABLEDB:0  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",convOrder.commonTab.isn,1)
    Call CheckDB_DOCS(convOrder.commonTab.isn,"ConvOrd ","2",fBODY,1)
    'CUREXCHANGES
    Call CheckQueryRowCount("CUREXCHANGES","fISN",convOrder.commonTab.isn,1)
    Call CheckDB_CUREXCHANGES(dbCUREXCHANGES(0),1)
    'HI
    For i = 0 to 1
        With dbHI(i)
            .fTYPE = "01"
            .fBASEBRANCH = "00 "
            .fBASEDEPART = "1  "
        End With
    Next 
    Call CheckQueryRowCount("HI", "fBASE", convOrder.commonTab.isn, 3)
    Call Check_DB_HI(dbHI(0) ,1)
    Call Check_DB_HI(dbHI(1) ,1)
    Call Check_DB_HI(dbHI(2) ,1)
    'HIREST
    Call CheckDB_HIREST("11", "341169779" , "-3577380.00" ,"001", "-8943.45", 1)
    Call CheckDB_HIREST("01", "341169779" , "179833400.00" ,"001", "449583.50", 1)
    
    Call CheckDB_HIREST("01", "1630226" , "206400.00" ,"000", "206400.00", 1)
           
    'Մուտք Հաշվառված վճարային փաստաթղթեր
    Log.Message  "Մուտք Հաշվառված վճարային փաստաթղթեր"
    folderDirect = "|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|Ð³ßí³éí³Í í×³ñ³ÛÇÝ ÷³ëï³ÃÕÃ»ñ"
    stDate = aqDateTime.Today
    enDate = aqDateTime.Today
    wUser = "77"
    docType = "ConvOrd"
    wName = "" 
    passNum = ""
    cliCode = ""
    paySysIn = ""
    paySysOut = ""
    acsBranch = ""
    acsDepart = ""
    docISN = ""
    selectedView = "Payments"
    expExcel = "0"
    Call OpenAccPaymentDocFolder(folderDirect, stDate, enDate, wUser, docType,wName, passNum, cliCode,paySysIn, paySysOut, acsBranch,_
                                               acsDepart, docISN, selectedView, expExcel)
    BuiltIn.Delay(2000)
    'Ջնջել Փոխանակման օրդերը
    Call SearchAndDelete ( "frmPttel", 2, convOrder.commonTab.docN, "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ")
    'Փակել "Հաշվառված վճարային փաստաթղթեր" թղթապանակը
    Call Close_Window(wMDIClient, "frmPttel")
    
    'SQL
    Log.Message "SQL Ստուգումներ Փոխանակման օրդերը Ջնջելուց հետո",,,SqlDivideColor  
    'DOCLOG
    Call CheckQueryRowCount("DOCLOG","fISN",convOrder.commonTab.isn,4)
    Call CheckDB_DOCLOG(convOrder.commonTab.isn,"77","N","1"," ",1)
    Call CheckDB_DOCLOG(convOrder.commonTab.isn,"77","C","10"," ",1) 
    Call CheckDB_DOCLOG(convOrder.commonTab.isn,"77","M","2","¶ñ³Ýóí»É »Ý Ó¨³Ï»ñåáõÙÝ»ñÁ",1)
    Call CheckDB_DOCLOG(convOrder.commonTab.isn,"77","D","999"," ",1)
    'DOCS
    Call CheckQueryRowCount("DOCS","fISN",convOrder.commonTab.isn,1)
    Call CheckDB_DOCS(convOrder.commonTab.isn,"ConvOrd ","999",fBODY,1)
    'HIREST
    Call CheckDB_HIREST("11", "341169779" , "-3577380.00" ,"001", "-8943.45", 1)
    Call CheckDB_HIREST("01", "341169779" , "180039800.00" ,"001", "450099.50", 1)
    
    Call CheckDB_HIREST("11", "1630226" , "0.00" ,"000", "0.00", 2)
    Call CheckDB_HIREST("01", "1630226" , "0.00" ,"000", "0.00", 7)
    'DOCP
    Call CheckQueryRowCount("DOCP","fISN",convOrder.commonTab.isn,0)
    
    'Անցնել "S.W.I.F.T. ԱՇՏ"
    Call ChangeWorkspace(c_SWIFT)
'-------------------------------------------------------------------
'------- Անցում "Ստացված խառը հաղորդագրություններ" թղթապանակ --------------
'-------------------------------------------------------------------
    Log.Message "Անցում 'Ստացված խառը հաղորդագրություններ' թղթապանակ",,,DivideColor  
    Call GoTo_Recieved_Messages (recieved, "|S.W.I.F.T. ²Þî                  |Ê³éÁ Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñ|êï³óí³Í Ë³éÁ Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñ")     
'--------------------------------------------------------------------------------        
'------------Կատարել Ճշտել (n95) գործողությունը ներմուծված փաստաթղթի համար----------------
'--------------------------------------------------------------------------------
    Log.Message "Ճշտել (n95)",,,DivideColor2           
    If SearchInPttel("frmPttel", 3, docN) Then
        Call wMainForm.MainMenu.Click(c_AllActions)
        Call wMainForm.PopupMenu.Click(c_ToQuestion95)    
        If wMDIClient.WaitVBObject("frmASDocForm",4000).Exists Then
            query95.commonTab.isn = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
            query95.commonTab.docN = Get_Rekvizit_Value("Document",query95.commonTab.tabN,"General","BMDOCNUM")
            'Լրացնել Հարց դաշտը
            Call Rekvizit_Fill ("Document",1,"General","QUESTION",query95.commonTab.queries)
            BuiltIn.Delay(2000)
            Call ClickCmdButton(1, "Î³ï³ñ»É")
            'Տպում է փաստաթղթի isn-ը ու համարը
            Log.Message "Query docN = "& query95.commonTab.docN
            Log.Message "Query fISN = "&  query95.commonTab.isn ,,,SqlDivideColor
        Else    
            Log.Error "Query window not found",,,ErrorColor
        End If
    Else 
        Log.Error "Document Row Not found",,,ErrorColor
    End If
    'Փակել "Ստացված խառը հաղորդագրություններ" թղթապանակը
    Call Close_Window(wMDIClient, "frmPttel")
    
    'SQL
    Log.Message "SQL Ստուգումներ Ճշտում (n95) ստեղծելուց հետո",,,SqlDivideColor  
    'DOCLOG
    Call CheckQueryRowCount("DOCLOG","fISN",query95.commonTab.isn,2)
    Call CheckDB_DOCLOG(query95.commonTab.isn,"77","N","1"," ",1)
    Call CheckDB_DOCLOG(query95.commonTab.isn,"77","C","9"," ",1) 
    'DOCS
    fBODY = "  CATEGORY:9  BMDOCNUM:"&query95.commonTab.docN&"  DATE:"&aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%Y%m%d")&"  TYPE:900  "_
            &"OSESNUM:1  OSESISN:123456  REFERENCE:"&docN&"  QUESTION:/12/  RS:2  VERIFIED:0  USERID:  77  SNDREC:CITIATWXXXX  "_
            &"BMIODATE:"&aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%Y%m%d")&"  RSBKMAIL:0  DELIV:0  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",query95.commonTab.isn,1)
    Call CheckDB_DOCS(query95.commonTab.isn,"MTN95   ","9",fBODY,1)
    'DOCSIM
    Call CheckQueryRowCount("DOCSIM","fISN",query95.commonTab.isn,1)
    'SW_MESSAGES
    Call CheckQueryRowCount("SW_MESSAGES","fISN",query95.commonTab.isn,1)
    With dbSW_MESSAGES
        .fDOCNUM = query95.commonTab.docN 
        .fISN = query95.commonTab.isn
        .fUNIQUEID = "ISN"&query95.commonTab.isn
    End With 
    Call CheckDB_SW_MESSAGES(dbSW_MESSAGES,1)
    
'------Մուտք Ուղարկվող Հաղորդագրություններ/Ուղարկվող խառը հաղորդագրություններ թղթապանակ----------------
    Call GoTo_Sending_Messages(sending, "|S.W.I.F.T. ²Þî                  |àõÕ³ñÏíáÕ Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñ|àõÕ³ñÏíáÕ Ë³éÁ Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñ")
    If SearchInPttel("frmPttel", 2, query95.commonTab.docN) Then
        'Կատարել Դիտել գործողությունը n995 հաղորդագրության համար
        Call wMainForm.MainMenu.Click(c_AllActions)
        Call wMainForm.PopupMenu.Click(c_View)    
        BuiltIn.Delay(2000)
        If wMDIClient.WaitVBObject("frmASDocForm",2000).Exists Then
            Call swQuery_Check(query95)
            Call ClickCmdButton(1, "OK")
        Else    
            Log.Error "Query window not found",,,ErrorColor
        End If
    Else    
        Log.Error "Documet row not found",,,ErrorColor
    End If     
    'Ջնջել n995 հաղորդագրությունը
    Call SearchAndDelete ( "frmPttel", 2, query95.commonTab.docN, "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ")
    'Փակել Ուղարկվող խառը հաղորդագրություններ թղթապանակը
    Call Close_Window(wMDIClient, "frmPttel")
    
    'SQL
    Log.Message "SQL Ստուգումներ Ճշտում (n95) ջնջելուց հետո",,,SqlDivideColor  
    'DOCLOG
    Call CheckQueryRowCount("DOCLOG","fISN",query95.commonTab.isn,3)
    Call CheckDB_DOCLOG(query95.commonTab.isn,"77","N","1"," ",1)
    Call CheckDB_DOCLOG(query95.commonTab.isn,"77","C","9"," ",1)
    Call CheckDB_DOCLOG(query95.commonTab.isn,"77","D","999"," ",1)
    'DOCS
    Call CheckQueryRowCount("DOCS","fISN",query95.commonTab.isn,1)
    Call CheckDB_DOCS(query95.commonTab.isn,"MTN95   ","999",fBODY,1)
    
'-------------------------------------------------------------------
'------- Անցում "Ստացված խառը հաղորդագրություններ" թղթապանակ --------------
'-------------------------------------------------------------------
    Log.Message "Անցում 'Ստացված խառը հաղորդագրություններ' թղթապանակ",,,DivideColor       
    Call GoTo_Recieved_Messages (recieved, "|S.W.I.F.T. ²Þî                  |Ê³éÁ Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñ|êï³óí³Í Ë³éÁ Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñ")
'--------------------------------------------------------------------------------        
'--------Կատարել Վճարման հանձնարարագիր գործողությունը ներմուծված փաստաթղթի համար-----------
'--------------------------------------------------------------------------------
    Log.Message "Վճարման հանձնարարագիր",,,DivideColor2      
    If SearchInPttel("frmPttel", 3, docN) Then
        Call wMainForm.MainMenu.Click(c_AllActions)
        Call wMainForm.PopupMenu.Click(c_PayOrder)    
        If wMDIClient.WaitVBObject("frmASDocForm",4000).Exists Then
            forPayOrder.commonTab.isn = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
            forPayOrder.commonTab.docN = Get_Rekvizit_Value("Document",forPayOrder.commonTab.tabN,"General","DOCNUM")
            'Լրացնել Վճարողի հաշիվ դաշտը
            Call Rekvizit_Fill ("Document",1,"General","ACCDB","[Home]![End][Del]" & forPayOrder.commonTab.payerAcc)
            'Լրացնել Ստացող դաշտը
            Call Rekvizit_Fill ("Document",1,"General","RECEIVER",forPayOrder.commonTab.receiver)
            'Տպում է փաստաթղթի isn-ը ու համարը
            Log.Message "Payorder fISN = "& forPayOrder.commonTab.isn ,,,SqlDivideColor
            Log.Message "Payorder docN = "& forPayOrder.commonTab.docN 
            Call ClickCmdButton(1, "Î³ï³ñ»É")
        Else    
            Log.Error "PayOrder window not found",,,ErrorColor
        End If
    Else    
        Log.Error "Documet row not found",,,ErrorColor
    End If
    'Ստուգել փաստաթղթի համապատասխանությունը օրինակի հետ
    If wMDIClient.WaitVBObject("FrmSpr",4000).Exists Then
        Call SaveDoc(savePath & "Actual\", "Actual_PayOrder")
        pathAct = savePath & "Actual\Actual_PayOrder.txt"
        pathExp = savePath & "Expected\Expected_PayOrder.txt"
        regex="(N \d{6})|([[].*])|(\d{2}[/]\d{2}[/]\d{2} \d{2}:\d{2})"
        Call Compare_Files(pathAct, pathExp, regex)
        'Փակել տպելու ձևը
        Call Close_Window(wMDIClient, "FrmSpr" )
    Else
        Log.Error "Can't find document print view",,,ErrorColor
    End If  
    'Փակել "Ստացված խառը հաղորդագրություններ" թղթապանակը
    Call Close_Window(wMDIClient, "frmPttel")
    
    'SQL
    Log.Message "SQL Ստուգումներ Վճարման հանձնարարագիր ստեղծելուց հետո",,,SqlDivideColor  
    Call DB_Initialize_MT900_RC_4(forPayOrder.commonTab.isn, forPayOrder.commonTab.docN )
    'DOCLOG
    Call CheckQueryRowCount("DOCLOG","fISN",forPayOrder.commonTab.isn,3)
    Call CheckDB_DOCLOG(forPayOrder.commonTab.isn,"77","N","1"," ",1)
    Call CheckDB_DOCLOG(forPayOrder.commonTab.isn,"77","M","99","àõÕ³ñÏí»É ¿ Ñ³ëï³ïÙ³Ý",1)
    Call CheckDB_DOCLOG(forPayOrder.commonTab.isn,"77","C","101"," ",1) 
    'DOCS
    fBODY = "  USERID:  77  ACSBRANCH:00  ACSDEPART:1  BLREP:0  DATE:"&aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%Y%m%d")&"  "_
            &"DOCNUM:"&forPayOrder.commonTab.docN&"  CLITRANS:2  CLICODE:00000015  RES:1  ACCDB:7770000003183311  PAYER:Client 00000015  RECOP:D  "_
            &"RECEIVER:0457RU45000  ISFINCOMPR:0  COUNTRY:AT  SUMMA:516  CUR:001  REPAY:0  SYSCASE:SWDCCONF  BMDOCNUM:"&docN&"  "_
            &"ADDINFO:Additional info  TCORRACC:000548101  CORRACC:10330030101  PAYSYSIN:5  ONORDER:0  FORTRADE:0  COVER:1  DUPLICATE:0  ACC2ACC:0  "_
            &"CANCELREQ:0  REF:SWIFT CHARGES  PINSTOP:D  PAYINST:SWIFT LA HULPE BELGIUM             INV0013368088 INV3009097822  PCORBANK:CITIATWXXXX  "_
            &"TYPECODE1:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28  TYPECODE2:-10 20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28  "_
            &"JURSTAT:21  VOLORT:7  PAYER1:Client 00000015  ACCTYPE:B  CORTYPE:3  SNDREC:CITIATWXXXX  MT:900  USEOVERLIMIT:0  NOTSENDABLE:0  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",forPayOrder.commonTab.isn,1) 
    Call CheckDB_DOCS(forPayOrder.commonTab.isn,"CrPayFor","101",fBODY,1)
    'DOCP
    Call CheckQueryRowCount("DOCP","fISN",forPayOrder.commonTab.isn,1)
    Call CheckDB_DOCP(forPayOrder.commonTab.isn,"CrPayFor",isn,1)
    'FODLERS
    Call CheckQueryRowCount("FOLDERS","fISN",forPayOrder.commonTab.isn,4)
    Call CheckDB_FOLDERS(dbFOLDERS(0),1) 
    Call CheckDB_FOLDERS(dbFOLDERS(1),1)
    Call CheckDB_FOLDERS(dbFOLDERS(2),1) 
    Call CheckDB_FOLDERS(dbFOLDERS(3),1)
    'HI
    Call CheckQueryRowCount("HI", "fBASE", forPayOrder.commonTab.isn, 2)
    Call Check_DB_HI(dbHI(0) ,1)
    Call Check_DB_HI(dbHI(1) ,1)
    'HIREST
    Call CheckDB_HIREST("11", "1706816" , "788833.70" ,"001", "2511.00", 1)
    Call CheckDB_HIREST("01", "1706816" , "-80599659.00" ,"001", "-191903.95", 1)
    
    Call CheckDB_HIREST("11", "1630510" , "57847521.00" ,"001", "100932.00", 1)
    Call CheckDB_HIREST("01", "1630510" , "30269163.05" ,"001", "83836.37", 1)
    
'-----Մուտք Գլխավոր Հաշվապահի ԱՇՏ/Աշխատանքային փաստաթղթեր------------------------------------   
    Log.Message "Մուտք Գլխավոր Հաշվապահի ԱՇՏ",,,DivideColor
    Call ChangeWorkspace(c_ChiefAcc) 
    Call GoTo_MainAccWorkingDocuments("|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|", workingDocs) 
    BuiltIn.Delay(3000)     
    If SearchInPttel("frmPttel", 2, forPayOrder.commonTab.docN) Then
        'Կատարել Դիտել գործողությունը Վճարման հանձնարարագրի համար
        Call wMainForm.MainMenu.Click(c_AllActions)
        Call wMainForm.PopupMenu.Click(c_View)    
        If wMDIClient.WaitVBObject("frmASDocForm",4000).Exists Then
            Call Foreign_Payment_Order_Sent_Check(forPayOrder)
            Call ClickCmdButton(1, "OK")
        Else    
            Log.Error "Query window not found",,,ErrorColor
        End If
    Else    
        Log.Error "Documet row not found",,,ErrorColor
    End If     
    Call SearchAndDelete ( "frmPttel", 2, forPayOrder.commonTab.docN, "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ")
    Call Close_Window(wMDIClient, "frmPttel")
    
    'SQL
    Log.Message "SQL Ստուգումներ Վճարման հանձնարարագիրը ջնջելուց հետո",,,SqlDivideColor  
    'DOCLOG
    Call CheckQueryRowCount("DOCLOG","fISN",forPayOrder.commonTab.isn,4)
    Call CheckDB_DOCLOG(forPayOrder.commonTab.isn,"77","N","1"," ",1)
    Call CheckDB_DOCLOG(forPayOrder.commonTab.isn,"77","M","99","àõÕ³ñÏí»É ¿ Ñ³ëï³ïÙ³Ý",1)
    Call CheckDB_DOCLOG(forPayOrder.commonTab.isn,"77","C","101"," ",1) 
    Call CheckDB_DOCLOG(forPayOrder.commonTab.isn,"77","D","999"," ",1)
    'DOCP
    Call CheckQueryRowCount("DOCP","fISN",forPayOrder.commonTab.isn,0)
    'DOCS
    Call CheckQueryRowCount("DOCS","fISN",forPayOrder.commonTab.isn,1)
    Call CheckDB_DOCS(forPayOrder.commonTab.isn,"CrPayFor","999",fBODY,1)
    'HIREST
    Call CheckDB_HIREST("11", "1706816" , "582433.70" ,"001", "1995.00", 1)
    
    Call CheckDB_HIREST("11", "1630510" , "58053921.00" ,"001", "101448.00", 1)

'------------------------------------------------------------------
'----------------- Ջնջում է ներմուծված փաստաթուղթը ---------------------
'------------------------------------------------------------------     
    'Մուտք Ադմինիստրարտոր ԱՇՏ
    Call ChangeWorkspace(c_Admin40)
    folderDirect = "|²¹ÙÇÝÇëïñ³ïáñÇ ²Þî 4.0|ÂÕÃ³å³Ý³ÏÝ»ñ|êï»ÕÍí³Í ÷³ëï³ÃÕÃ»ñ"
    stDate = aqDateTime.Today
    enDate = aqDateTime.Today
    wUser = 77
    docType = ""
    'Մուտք Ստեղծված փաստաթղթեր թղթապանակ
    Call OpenCreatedDocFolder(folderDirect, stDate, enDate, wUser, docType)
    
    'Ջնջում է փաստաթղթի հետ կապակցված Տարանցիկ վճարային փոխանցման փաստաթուղթը
    Call SearchAndDelete( "frmPttel", 2, isn, "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ" ) 
    Call Close_Window(wMDIClient, "frmPttel")
    Call Close_AsBank()    
End Sub


Sub Test_Initialize_SWIFT_900_RC()
    sDATE = "20020101"
    fDATE = "20260101"
    
    min=100
    max=999
    importPath = Project.Path & "Stores\SWIFT\HT900\ImportFile\Import\"
    docN = "FGVM0012047905"
    savePath =  Project.Path & "Stores\SWIFT\HT900\PrintView\"
    
    Set workingDocs = New_MainAccWorkingDocuments()
    With workingDocs
        .startDate = aqDateTime.Today
		.endDate = aqDateTime.Today
    End With
    Set recieved = New_Recieved()
    With recieved
        .sDate = aqDateTime.Today
        .eDate = aqDateTime.Today
    End With
    
    Set sending = New_Sending()
    With sending
        .sDate = "010120"
        .eDate = aqDateTime.Today
    End With
    
    Set currExchange = New_currExchange(0, 0, 0)
    With currExchange
'        1 Ընդհանուր
        .commonTab.div = "00"
        .commonTab.dep = "1"
        .commonTab.fDate = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%d/%m/%y")
        .commonTab.dAcc = "00067110103"
        .commonTab.cAcc = "10330030101"
        .commonTab.cur1 = "003"
        .commonTab.cur2 = "001"
        .commonTab.way = "1"
        .commonTab.course = "1.4189/1"
        .commonTab.sum1 = "363.66"
        .commonTab.sum2 = "516.00"
        .commonTab.buySell = "1"
        .commonTab.opType = "1"
        .commonTab.opPlace = "3"
        .commonTab.busField = "7"
        .commonTab.legalPos = "21"
        .commonTab.aim = "Additional info"  
'        2 Լրացուցիչ
        .addTab.CBCourse1 = "450.0000/1"
        .addTab.CBCourse2 = "400.0000/1"
        If aqDateTime.Compare(aqConvert.DateTimeToFormatStr(aqDateTime.Time, "%H:%M"), "16:00") < 0 Then
            .addTab.fTime = "1"
        Else
            .addTab.fTime = "2"
        End If
        .addTab.comAccCur = "000"
        .addTab.incAccCom = "001453300  "
        .addTab.recPaySys = "5"
        .addTab.incCurrExch = "000931900"
        .addTab.expenseCurrExch = "001434300"
'        3 Դրամարկղ
        .cashDeskTab.totInputCur = "000"
        .cashDeskTab.clientCode = "00000671"
        .cashDeskTab.fName = "äáÕáëÛ³Ý"
        .cashDeskTab.lName = "ê³ñ·Çë"
        .cashDeskTab.idNum = "AA48536633"
        .cashDeskTab.idType = "01"
        .cashDeskTab.issuedBy = "001"
        .cashDeskTab.issueDate = "31/08/2004"
        .cashDeskTab.socCard = "1234567890"
        .cashDeskTab.birthDate = "17/01/2005"
        .cashDeskTab.citizen = "1"
        .cashDeskTab.country = "AM"
    End With
    
    Set verifDoc = New_VerificationDocument() 
    With verifDoc
        .DocType = "CurChng"
        .User = "77"
    End With
    Set memOrder = New_memOrder()
    With memOrder
        .Div = "00"
        .Dep = "1"
        .MDate = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%d/%m/%y")
        .AccD = "000048201  "
        .AccC = "10330030101"
        .Curr = "001"
        .Sum = "516.00"
        .Aim = "Additional info"
        .Paysys = "5"
    End With
    
    Set convOrder = New_ConversionOrder(0, 0, 0)
    With convOrder
        .commonTab.div = "00"
        .commonTab.dep = "1"
        .commonTab.fDate = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%d/%m/%y")
        .commonTab.dAcc = "000084700  "
        .commonTab.cAcc = "10330030101"
        .commonTab.cur1 = "000"
        .commonTab.cur2 = "001"
        .commonTab.CBCourse = "400.0000/1"
        .commonTab.sum1 = "206,400.00"
        .commonTab.sum2 = "516.00"
        .commonTab.paySys = "5"
        .commonTab.aim = "Additional info"  
        .psTab.purSale = "2"
        If aqDateTime.Compare(aqConvert.DateTimeToFormatStr(aqDateTime.Time, "%H:%M"), "16:00") < 0 Then
            .psTab.fTime = "1"
        Else
            .psTab.fTime = "2"
        End If
        .psTab.opType = "1"
        .psTab.opPlace = "3"
        .psTab.busField = ""
        .psTab.legalPos = ""
    End With  
    
    Set query95 = New_swQuery()
    With query95
        .commonTab.category = "9"
        .commonTab.origMessageDate = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%d/%m/%y")
        .commonTab.origMessageMT = "900"
        .commonTab.sesN = "1"
        .commonTab.sesIsn = "123456"
        .commonTab.reference = "FGVM0012047905"
        .commonTab.queries = "/12/"
        .commonTab.sentRec = "2"
        .commonTab.origMessage = ":20:FGVM0012047905"& vbNewLine _
                                &":21:SWIFT CHARGES"& vbNewLine _
                                &":25:400886573500USD"& vbNewLine _
                                &":32A:220322USD516,"& vbNewLine _
                                &":52D:SWIFT LA HULPE BELGIUM"& vbNewLine _
                                &"INV0013368088 INV3009097822"& vbNewLine _
                                &":72:Additional info"
        .addTab.sendRec = "CITIATWXXXX"
        .addTab.sendRecDate = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%d/%m/%y")                                 
    End With
    Set forPayOrder = New_ForeignPaymentOrderSent(0, 0 ,0 , 0)
    With forPayOrder
        'Ընդհանուր
        .commonTab.div = "00"
        .commonTab.dep = "1"
        .commonTab.fDate = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%d/%m/%y")
        .commonTab.clientTransfer = "2"
        .commonTab.payerCode = "00000015"
        .commonTab.residence = "1"
        .commonTab.payerAcc = "77700/00003183311"
        .commonTab.payer = "Client 00000015"
        .commonTab.benefClientType = "D"
        .commonTab.receiver = "0457RU45000"
        .commonTab.receiptCountry = "AT"
        .commonTab.amount = "516.00"
        .commonTab.cur = "001"
        'Լրացուցիչ
        .addTab.docN = "FGVM0012047905"
        .addTab.addInfo = "Additional info"
        .addTab.transitAcc = "000548101  "
        .addTab.correspondentAcc = "10330030101"
        .addTab.recPaySys = "5"
        'Ֆին. կազմակերպ.
        .finOrgTab.reference = "SWIFT CHARGES"
        .finOrgTab.ordInstType = "D"
        .finOrgTab.ordInst = "SWIFT LA HULPE BELGIUM             INV0013368088 INV3009097822"
        .finOrgTab.payBankCorr = "CITIATWXXXX"
        'Գանձում փոխանցումից
        .tChargeTab.legPosition = "21"
        .tChargeTab.busField = "7"
        'Դրամարկղ
        .cDeskTab.depositor = "Client 00000015"
        'Այլ
        .otherTab.accType = "B"
        .otherTab.fType = "3"
        .otherTab.sendRec = "CITIATWXXXX"
        .otherTab.msgType = "900"
    End With
End Sub

Sub DB_Initialize_MT900_RC_1(fISN, fDOCN)
     Set dbFOLDERS(0) = New_DB_FOLDERS()
    With dbFOLDERS(0) 
        .fFOLDERID = "C.566034471"
        .fNAME = "CurChng "
        .fKEY = fISN
        .fISN = fISN
        .fSTATUS = "1"
        .fCOM = "²ñï³ñÅáõÛÃÇ ÷áË³Ý³ÏáõÙ"
        .fSPEC = "²Ùë³ÃÇí- "&aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%d/%m/%y")&" N- "&fDOCN&" ¶áõÙ³ñ-               363.66 ²ñÅ.- 003 [Üáñ]"
        .fECOM = ""
        .fDCBRANCH = "   "
        .fDCDEPART = "   "
    End With
    
    Set dbFOLDERS(1) = New_DB_FOLDERS()
    With dbFOLDERS(1) 
        .fFOLDERID = "Oper."&aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%Y%m%d")
        .fNAME = "CurChng "
        .fKEY = fISN
        .fISN = fISN
        .fSTATUS = "5"
        .fCOM = "²ñï³ñÅáõÛÃÇ ÷áË³Ý³ÏáõÙ"
        .fSPEC = fDOCN&"77700000671101037770010330030101          363.66003Üáñ                                                   77äáÕáëÛ³Ý ê³ñ·Çë"_
                &"                 AA48536633 001 31/08/2004                              5        Additional info"_
                &"                                                                                                                             "
        .fECOM = "Currency Exchange"
        .fDCBRANCH = "00 "
        .fDCDEPART = "1  "
    End With
    
    Set dbFOLDERS(2) = New_DB_FOLDERS()
    With dbFOLDERS(2) 
        .fFOLDERID = "Ver."&aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%Y%m%d")&"001"
        .fNAME = "CurChng "
        .fKEY = fISN
        .fISN = fISN
        .fSTATUS = "4"
        .fCOM = "²ñï³ñÅáõÛÃÇ ÷áË³Ý³ÏáõÙ"
        .fSPEC = fDOCN&"77700000671101037770010330030101          363.66003  77Additional info                 äáÕáëÛ³Ý ê³ñ·Çë"_
                &"                                                        5 "
        .fECOM = "Currency Exchange"
        .fDCBRANCH = "00 "
        .fDCDEPART = "1  "
    End With 
    
    Set dbHI(0) = New_DB_HI()
    With dbHI(0)  
        .fBASE = fISN
        .fDATE = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%Y-%m-%d")
        .fTYPE = "11"
        .fSUM = "206400.00"
        .fCUR = "001"
        .fCURSUM = "516.00"
        .fOP = "CEX"
        .fDBCR = "C"
        .fADB = "860427540"
        .fACR = "341169779"
        .fSPEC = fDOCN&"                   Additional info                   0   400.0000    1"
        .fBASEBRANCH = ""
        .fBASEDEPART = ""
    End With
    
    Set dbHI(1) = New_DB_HI()
    With dbHI(1)  
        .fBASE = fISN
        .fDATE = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%Y-%m-%d")
        .fTYPE = "11"
        .fSUM = "206400.00"
        .fCUR = "003"
        .fCURSUM = "363.66"
        .fOP = "CEX"
        .fDBCR = "D"
        .fADB = "860427540"
        .fACR = "341169779"
        .fSPEC = fDOCN&"                   Additional info                   1   567.5631    1"
        .fBASEBRANCH = ""
        .fBASEDEPART = ""
    End With
    
    Set dbHI(2) = New_DB_HI()
    With dbHI(2)  
        .fBASE = fISN
        .fDATE = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%Y-%m-%d")
        .fTYPE = "11"
        .fSUM = "42753.00"
        .fCUR = "000"
        .fCURSUM = "42753.00"
        .fOP = "MSC"
        .fDBCR = "D"
        .fADB = "1629708"
        .fACR = "860427540"
        .fSPEC = fDOCN&"                   ìÝ³ëÝ»ñ ³ñï. ÷áË³Ý³ÏáõÙÇó         1     1.0000    1"
        .fBASEBRANCH = ""
        .fBASEDEPART = ""
    End With
    
    Set dbHI(3) = New_DB_HI()
    With dbHI(3)  
        .fBASE = fISN
        .fDATE = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%Y-%m-%d")
        .fTYPE = "11"
        .fSUM = "42753.00"
        .fCUR = "003"
        .fCURSUM = "0.00"
        .fOP = "MSC"
        .fDBCR = "C"
        .fADB = "1629708"
        .fACR = "860427540"
        .fSPEC = fDOCN&"                   ìÝ³ëÝ»ñ ³ñï. ÷áË³Ý³ÏáõÙÇó         0   450.0000    1"
        .fBASEBRANCH = ""
        .fBASEDEPART = ""
    End With 
    
    Set dbHI(4) = New_DB_HI()
    With dbHI(4)  
        .fBASE = fISN
        .fDATE = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%Y-%m-%d")
        .fTYPE = "CE"
        .fSUM = "516.00"
        .fCUR = "003"
        .fCURSUM = "363.66"
        .fOP = "PUR"
        .fDBCR = "D"
        .fADB = "-1"
        .fACR = "-1"
        .fSPEC = "%.4189/1       0"&fDOCN&"7 "
        .fBASEBRANCH = ""
        .fBASEDEPART = ""
    End With
    
    Set dbCUREXCHANGES(0) = New_DB_CUREXCHANGES()
    With dbCUREXCHANGES(0)
        .fISN = fISN
        .fDOCTYPE = "CurChng "
        .fCOMPLETED = "1"
        .fEXPORTED = "0"
        .fDATE = aqConvert.DateTimeToFormatStr(aqDateTime.Now(),"%Y%m%d")
        .fSTATE = "11"
        .fDOCNUM = fDOCN
        .fCLIENT = "00000671"
        .fNAME = "äáÕáëÛ³Ý ê³ñ·Çë"
        .fACCDB = "7770000067110103"
        .fCURDB = "003"
        .fSUMDB = "363.66"
        .fACCCR = "7770010330030101"
        .fCURCR = "001"
        .fSUMCR = "516.00"
        .fCOM = "Additional info                                                                                                                             "
        .fPASSPORT = "AA48536633                      "
        .fKASCODE = "   "
        .fCURCOMIS = "000"
        .fSUMCOMIS = "0.00"
        .fSUMCOMISAMD = "0.00"
        .fACSBRANCH = "00"
        .fACSDEPART = "1"
    End With
    
End Sub

Sub DB_Initialize_MT900_RC_2(fISN, fDOCN)

    With dbFOLDERS(0) 
        .fFOLDERID = "Oper."&aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%Y%m%d")
        .fNAME = "MemOrd  "
        .fKEY = fISN
        .fISN = fISN
        .fSTATUS = "1"
        .fCOM = "ÐÇß³ñ³ñ ûñ¹»ñ"
        .fSPEC = fDOCN&"77700000048201  7770010330030101          516.00001Üáñ                                                   77"_
                &"                                                                                       5        Additional info"_
                &"                                                                                                                             "
        .fECOM = "Memorial Order"
        .fDCBRANCH = "00 "
        .fDCDEPART = "1  "
    End With
    
    With dbHI(0)  
        .fBASE = fISN
        .fDATE = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%Y-%m-%d")
        .fTYPE = "11"
        .fSUM = "206400.00"
        .fCUR = "001"
        .fCURSUM = "516.00"
        .fOP = "MSC"
        .fDBCR = "D"
        .fADB = "1714456"
        .fACR = "341169779"
        .fSPEC = fDOCN&"                   Additional info                   0   400.0000    1"
        .fBASEBRANCH = ""
        .fBASEDEPART = ""
    End With       
     
    With dbHI(1)  
        .fBASE = fISN
        .fDATE = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%Y-%m-%d")
        .fTYPE = "11"
        .fSUM = "206400.00"
        .fCUR = "001"
        .fCURSUM = "516.00"
        .fOP = "MSC"
        .fDBCR = "C"
        .fADB = "1714456"
        .fACR = "341169779"
        .fSPEC = fDOCN&"                   Additional info                   1   400.0000    1"
        .fBASEBRANCH = ""
        .fBASEDEPART = ""
    End With           
End Sub

Sub DB_Initialize_MT900_RC_3(fISN, fDOCN)
    Set dbFOLDERS(0) = New_DB_FOLDERS()
    With dbFOLDERS(0) 
        .fFOLDERID = "Oper."&aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%Y%m%d")
        .fNAME = "ConvOrd "
        .fKEY = fISN
        .fISN = fISN
        .fSTATUS = "1"
        .fCOM = "öáË³Ý³ÏÙ³Ý ûñ¹»ñ"
        .fSPEC = fDOCN&"77700000084700  7770010330030101          516.00001Üáñ                                                   77"_
                &"                                                                                       5        Additional info"_
                &"                                                                                                                             "
        .fECOM = "Conversion Order"
        .fDCBRANCH = "00 "
        .fDCDEPART = "1  "
    End With 
    Set dbHI(0) = New_DB_HI()
    With dbHI(0)  
        .fBASE = fISN
        .fDATE = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%Y-%m-%d")
        .fTYPE = "11"
        .fSUM = "206400.00"
        .fCUR = "000"
        .fCURSUM = "206400.00"
        .fOP = "CEX"
        .fDBCR = "D"
        .fADB = "1630226"
        .fACR = "341169779"
        .fSPEC = fDOCN&"                   Additional info                   0     1.0000    1"
        .fBASEBRANCH = ""
        .fBASEDEPART = ""
    End With
    Set dbHI(1) = New_DB_HI()
    With dbHI(1)  
        .fBASE = fISN
        .fDATE = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%Y-%m-%d")
        .fTYPE = "11"
        .fSUM = "206400.00"
        .fCUR = "001"
        .fCURSUM = "516.00"
        .fOP = "CEX"
        .fDBCR = "C"
        .fADB = "1630226"
        .fACR = "341169779"
        .fSPEC = fDOCN&"                   Additional info                   1   400.0000    1"
        .fBASEBRANCH = ""
        .fBASEDEPART = ""
    End With
    Set dbHI(2) = New_DB_HI()
    With dbHI(2)  
        .fBASE = fISN
        .fDATE = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%Y-%m-%d")
        .fTYPE = "CE"
        .fSUM = "516.00"
        .fCUR = "000"
        .fCURSUM = "206400.00"
        .fOP = "SAL"
        .fDBCR = "D"
        .fADB = "-1"
        .fACR = "-1"
        .fSPEC = "%  400.0000/1     0"&fDOCN&"  "
        .fBASEBRANCH = "00 "
        .fBASEDEPART = "1  "
    End With
    
    Set dbCUREXCHANGES(0) = New_DB_CUREXCHANGES()
    With dbCUREXCHANGES(0)
        .fISN = fISN
        .fDOCTYPE = "ConvOrd "
        .fCOMPLETED = "1"
        .fEXPORTED = "0"
        .fDATE = aqConvert.DateTimeToFormatStr(aqDateTime.Now(),"%Y%m%d")
        .fSTATE = "2"
        .fDOCNUM = fDOCN
        .fCLIENT = "        "
        .fNAME = ""
        .fACCDB = "77700000084700  "
        .fCURDB = "000"
        .fSUMDB = "206400.00"
        .fACCCR = "7770010330030101"
        .fCURCR = "001"
        .fSUMCR = "516.00"
        .fCOM = "Additional info                                                                                                                             "
        .fPASSPORT = ""
        .fKASCODE = "   "
        .fCURCOMIS = "   "
        .fSUMCOMIS = "0.00"
        .fSUMCOMISAMD = "0.00"
        .fACSBRANCH = "00"
        .fACSDEPART = "1"
    End With 
    
    Set dbSW_MESSAGES = New_SW_MESSAGES()
    With dbSW_MESSAGES
        .fDOCNUM = fDOCN
        .fISN = fISN
        .fUNIQUEID = "ISN"&fISN
        .fDATE = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%Y-%m-%d")
        .fMT = "995"
        .fCATEGORY = "2"
        .fSR = "0"
        .fSRBANK = "CITIATWXXXX"
        .fSYS = "1"
        .fSTATE = "  "
        .fUSER = "77"
        .fACCDB = ""
        .fACCCR = ""
        .fAMOUNT = "0.00"
        .fCURR = "   "
        .fPAYER = ""
        .fRECEIVER = ""
        .fAIM = ""
        .fBRANCH = Null
        .fDEPART = Null
    End With 
End Sub

Sub DB_Initialize_MT900_RC_4(fISN, fDOCN)               

    With dbFOLDERS(0) 
        .fFOLDERID = "C.1628330"
        .fNAME = "CrPayFor"
        .fKEY = fISN
        .fISN = fISN
        .fSTATUS = "0"
        .fCOM = "ØÇç³½·. í×³ñÙ³Ý Ñ³ÝÓÝ³ñ³ñ³·Çñ (áõÕ.)"
        .fSPEC = "²Ùë³ÃÇí- "&aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%d/%m/%y")&" N- "&fDOCN&" ¶áõÙ³ñ-               "_
                &"516.00 ²ñÅ.- 001 [àõÕ³ñÏí³Í I Ñ³ëï³ïÙ³]"
        .fECOM = "Foreign Payment Order (to be sent)"
        .fDCBRANCH = "   "
        .fDCDEPART = "   "
    End With
    Set dbFOLDERS(1) = New_DB_FOLDERS()
    With dbFOLDERS(1) 
        .fFOLDERID = "EPS."&fISN
        .fNAME = "CrPayFor"
        .fKEY = "FGVM0012047905  "
        .fISN = fISN
        .fSTATUS = "0"
        .fCOM = "ØÇç³½·. í×³ñÙ³Ý Ñ³ÝÓÝ³ñ³ñ³·Çñ (áõÕ.)"
        .fSPEC = "900B  1 "&aqConvert.DateTimeToFormatStr(aqDateTime.Now(),"%Y%m%d")&"          516.0000100000000     CITIATWXXXX7770000003183311Client"_
                &" 00000015                    10330030101                                                  AT            "_
                &"0.00àõÕ³ñÏí³Í I Ñ³ëï³ïÙ³       77700            0.00                                 SWIFT CHARGES                               "_
                &"0.00      77                                              5"
        .fECOM = "Foreign Payment Order (to be sent)"
        .fDCBRANCH = "00 "
        .fDCDEPART = "1  "
    End With 
    Set dbFOLDERS(2) = New_DB_FOLDERS()
    With dbFOLDERS(2) 
        .fFOLDERID = "Oper."&aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%Y%m%d")
        .fNAME = "CrPayFor"
        .fKEY = fISN
        .fISN = fISN
        .fSTATUS = "0"
        .fCOM = "ØÇç³½·. í×³ñÙ³Ý Ñ³ÝÓÝ³ñ³ñ³·Çñ (áõÕ.)"
        .fSPEC = fDOCN&"77700000031833117770010330030101          516.00001àõÕ³ñÏí³Í I Ñ³ëï³ïÙ³Ý                                 77Client 00000015"_
                &"                                                                        5                                                        "_
                &"                                                                                            "
        .fECOM = "Foreign Payment Order (to be sent)"
        .fDCBRANCH = "00 "
        .fDCDEPART = "1  "
    End With 
    Set dbFOLDERS(3) = New_DB_FOLDERS()
    With dbFOLDERS(3) 
        .fFOLDERID = "Ver."&aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%Y%m%d")&"001"
        .fNAME = "CrPayFor"
        .fKEY = fISN
        .fISN = fISN
        .fSTATUS = "4"
        .fCOM = "ØÇç³½·. í×³ñÙ³Ý Ñ³ÝÓÝ³ñ³ñ³·Çñ (áõÕ.)"
        .fSPEC = fDOCN&"77700000031833117770010330030101          516.00001  77                                Client 00000015                 "_
                &"0457RU45000                            5 "
        .fECOM = "Foreign Payment Order (to be sent)"
        .fDCBRANCH = "00 "
        .fDCDEPART = "1  "
    End With                   

    With dbHI(0)  
        .fBASE = fISN
        .fDATE = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%Y-%m-%d")
        .fTYPE = "11"
        .fSUM = "206400.00"
        .fCUR = "001"
        .fCURSUM = "516.00"
        .fOP = "TRF"
        .fDBCR = "C"
        .fADB = "1706816"
        .fACR = "1630510"
        .fSPEC = fDOCN&"                                                     1   400.0000    1"
        .fBASEBRANCH = ""
        .fBASEDEPART = ""
    End With
    
    With dbHI(1)  
        .fBASE = fISN
        .fDATE = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%Y-%m-%d")
        .fTYPE = "11"
        .fSUM = "206400.00"
        .fCUR = "001"
        .fCURSUM = "516.00"
        .fOP = "TRF"
        .fDBCR = "D"
        .fADB = "1706816"
        .fACR = "1630510"
        .fSPEC = fDOCN&"                                                     0   400.0000    1"
        .fBASEBRANCH = ""
        .fBASEDEPART = ""
    End With
 
End Sub

