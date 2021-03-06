'USEUNIT SWIFT_International_Payorder_Library
'USEUNIT Library_Common
'USEUNIT Library_Colour
'USEUNIT Library_CheckDB
'USEUNIT Constants
'USEUNIT Library_Contracts
'USEUNIT Main_Accountant_Filter_Library 
Option Explicit
'Test Case ID 183736 sysType = 1 -SWIFT
'Test Case ID 185519 sysType = 2 -SPFS
Dim sDATE, fDATE, fileFrom, fileTo, what, fWith, rand, max, min, for_Ex_Con(1), dbFOLDERS(3), sumSQL , recieved
Dim folderDirect, stDate, enDate, wUser, docType, importPath, fBODY, i, dbSW_MESSAGES(1), settingsPath
Dim SortArr(1), regex, Path1, Path2, sending
Sub SWIFT_300_International_Import_Test(sysType)
    'Համակարգ մուտք գործել ARMSOFT օգտագործողով
    Call Test_Initialize_SWIFT_300 ()
    Call Initialize_AsBank("bank", sDATE, fDATE)
    Login("ARMSOFT")
'----------------------------------------------------
'--------------- Կարգավորումների ներմուծում --------------
'----------------------------------------------------
    Log.Message "Կարգավորումների ներմուծում ",,,DivideColor
    settingsPath = Project.Path & "Stores\SWIFT\HT300\Settings\Setting_1.txt"'SWSMG - Պարամետրում '300' հաղորդագրությունը առկա չէ
    folderDirect = "|²¹ÙÇÝÇëïñ³ïáñÇ ²Þî 4.0|Ð³Ù³Ï³ñ·³ÛÇÝ ³ßË³ï³ÝùÝ»ñ|îíÛ³ÉÝ»ñÇ Ý»ñÙáõÍáõÙ|ö³ëï³ÃÕÃ»ñÇ Ý»ñÙáõÍáõÙ"
    Call ChangeWorkspace(c_Admin40)
    BuiltIn.Delay(3000)
    Call Settings_Import(settingsPath,folderDirect)
    Login("ARMSOFT")
    Call CheckQueryRowCount("PARAMS","fVALUE","101,110,200,201,203,102,N99,103,N98,410,N96",1)
'-----------------------------------------------------------------------------
'------ "S.W.I.F.T. ԱՇՏ/Պարամետրեր"-ում կատարել համապատասխան փոփոխությունները-------
'-----------------------------------------------------------------------------
    Log.Message "-- S.W.I.F.T. ԱՇՏ/Պարամետրեր-ում կատարել համապատասխան փոփոխությունները --",,,DivideColor  
    'Մուտք գործել "S.W.I.F.T. ԱՇՏ"
    Call ChangeWorkspace(c_SWIFT)
    Call wTreeView.DblClickItem("|S.W.I.F.T. ²Þî                  |ä³ñ³Ù»ïñ»ñ")
    BuiltIn.Delay(3000)
    'Նոր փաստաթղթի համարի գեներացում
    max=100
    min=999
    Randomize
    rand = Int((max-min+1)*Rnd+min)
    fileFrom = Project.Path &"Stores\SWIFT\HT300\ImportFile\IA000390.RJE"
    fileTo = Project.Path &"Stores\SWIFT\HT300\ImportFile\Import\IA000391.RJE"
    what = "CITI2111089856"
    fWith = "CITI2111089" & rand
    'Ջնջում է Import թղթապանակի պարունակությունը
    aqFileSystem.DeleteFile(Project.Path &"Stores\SWIFT\HT300\ImportFile\Import\*")
    Log.Message(fWith)
    
    'SWGPI Պարամետրի փոփոխում SQL հարցման միջոցով
    Call SetParameter("SWGPI", "1")
    
    'Վերարտագրում է նախօրոք տրված ֆայլը մեկ այլ ֆայլի մեջ` կատարելով փոփոխություն
    Call Read_Write_File_With_replace(fileFrom,fileTo,what,fWith)
    importPath = Project.Path & "Stores\SWIFT\HT300\ImportFile\Import\"
    Select Case sysType
    Case 1
        'SWOUT պարամետրի փոփոխում
        Call SetParameter_InPttel("SWOUT" ,importPath)
        'SWSPFSACKDIR պարամետրի փոփոխում
        Call SetParameter("SWSPFSACKDIR", "")
        'SWSPFSNAKDIR պարամետրի փոփոխում
        Call SetParameter("SWSPFSNAKDIR", "")
        'SWSPFSOUT պարամետրի փոփոխում
        Call SetParameter("SWSPFSOUT", "")
    Case 2 
        'SWSPFSACKDIR պարամետրի փոփոխում
        Call SetParameter_InPttel("SWSPFSACKDIR" ,importPath)
        'SWSPFSNAKDIR պարամետրի փոփոխում
        Call SetParameter_InPttel("SWSPFSNAKDIR" ,importPath)
        'SWSPFSOUT պարամետրի փոփոխում
        Call SetParameter_InPttel("SWSPFSOUT" ,importPath)      
        'SWOUT պարամետրի փոփոխում
        Call SetParameter("SWOUT", "")  
    End Select        
    Call Close_Window(wMDIClient, "frmPttel")
    Login("ARMSOFT")

'-----------------------------------------------------------------------------
'----------------- Կատարել Ընդունել SWIFT համակարգից գործողությունը ------------------
'----------------------------------------------------------------------------- 
    Log.Message "Կատարել Ընդունել SWIFT համակարգից գործողությունը",,,DivideColor
    
    'Մուտք գործել "S.W.I.F.T. ԱՇՏ"
    Call ChangeWorkspace(c_SWIFT)
    Call Recieve_From_SWIFT(1)
'-------------------------------------------------------------------
'---------------- Ստուգում է փաստաթղթի առկայությունը --------------------
'-------------------------------------------------------------------
    Log.Message " Ստուգում է փաստաթղթի առկայությունը ",,,DivideColor       
    
    'Մուտք գործել Փոխանցումներ/Ստացված փաստաթղթեր թղթապանակ
    Call GoTo_Recieved_Messages (recieved, "|S.W.I.F.T. ²Þî                  |öáË³ÝóáõÙÝ»ñ|êï³óí³Í ÷áË³ÝóáõÙÝ»ñ")
    'Աղյուսակի տեսքի համեմատում
    Call ColumnSorting(SortArr, 2, "frmPttel")
    
    Path1 = Project.Path & "Stores\SWIFT\HT300\Actual_Pttel.txt"
    Path2 = Project.Path & "Stores\SWIFT\HT300\Expected_Pttel.txt"
    regex = "(\d{2}[/]\d{2}[/]\d{2})|(\d{2}:\d{2})|(CITI\d{10})"
    
    'Արտահանել և Ð³Ù»Ù³ï»É »ñÏáõ txt ý³ÛÉ»ñ
    Call ExportToTXTFromPttel("frmPttel",Path1)
    Call Compare_Files(Path2, Path1, regex)
    'Ստուգում է փաստաթղթի առկայությունը
    for_Ex_Con(0).docN = fWith
    for_Ex_Con(1).reference = fWith
    If SearchInPttel("frmPttel",3, for_Ex_Con(0).docN) Then
        Call wMainForm.MainMenu.Click(c_AllActions)
        Call wMainForm.PopupMenu.Click(c_View)
        'Ստուգում է հաղորդագրության պատուհանի բովանդակությունը
        Call Check_Foreign_Exchange_Confirm_Window(for_Ex_Con(0))
        Call ClickCmdButton(1, "OK")    
        Log.Message "SQL ստուգում (Ընդունել SWIFT համակարգից գործողություն)ից հետո",,,SqlDivideColor
        Log.Message "fISN = "& for_Ex_Con(0).isn,,,SqlDivideColor
     Else 
        Log.Error "Document row not found",,,ErrorColor
     End If      
    'SQL
    Call Initialize_DB_SWIFT_300(for_Ex_Con(0).isn, for_Ex_Con(0).docN)
    'DOCLOG
    Call CheckQueryRowCount("DOCLOG","fISN",for_Ex_Con(0).isn,1)
    Call CheckDB_DOCLOG(for_Ex_Con(0).isn,"77","N","10","",1)
    
    'DOCS
    Call CheckQueryRowCount("DOCS","fISN",for_Ex_Con(0).isn,1)

    fBODY = "  BMDOCNUM:" & for_Ex_Con(0).docN & "  REF:11111  OPERTYPE:CANC  COMREF:PRML223858  OPERSCOPE:AGNT  "_
          & "BLOCKIND:N  SPLITIND:Y  PARTYAOP:A  PARTYAID:123546988  PARTYA:CITIUS33XXX  PARTYBOP:A  PARTYBID:125645656asd  "_
          & "PARTYB:CITIUS33XXX  FUNDOP:J  FUND:CITIUS33XXX  VERIFIED:0  DATE:20211108  PAYDATE:20211108  CURB:001  CURS:003  SUMMAB:100  "_
          &"SUMMAS:72.16  COURSE:1.3858/1  RATE:1,3858  SNDREC:XXN}{4:{4:  DAGENTBOP:A  DAGENTBID:464asd64s  DAGENTB:CITIUS33XXX  "_
          &"MEDBOP:J  MEDBBANK:CITIUS33XXX  RAGENTBOP:A  RAGENTBID:45646asd6564  RAGENTB:AMASJPJZXXX  DAGENTSOP:J  DAGENTS:CITIUS33XXX  MEDSOP:A  "_
          &"MEDSID:645646  MEDSBANK:DGPBDE3MBRA  RAGENTSOP:J  RAGENTS:CITIUS33XXX  BENINSTSOP:A  BENINSTSID:asdasd  BENINSTS:CITIUS33XXX  "_
          &"CNTREF:1125s52248s5  BRKREF:21asd84646  TERMS:sadf12541a  ADDINFO:12556s54das  BMNAME:IA000391#001  "_
          &"BMIODATE:" & aqConvert.DateTimeToFormatStr(aqDateTime.Now(),"%Y%m%d") & "  UNIQUEID:US33XXXXN}{4:  RSBKMAIL:1  DELIV:0  USERID:  77  "_
          &"PROCESSED:0  CONTACT:55d445s8  DEALMETH:BROK  DEALADD:3218655s85  DEALAOP:A  DEALAID:125598556asd  DEALA:CITIUS33XXX  DEALBOP:A  "_
          &"DEALBID:s99d56a9662  DEALB:CITIUS33XXX  BROKEROP:J  BROKER:BRAJINBBBYF  BRCHGCUR:001  BRCHGSUM:10  COUNT:2  EVNTTYPE:EAMT  "_
          &"EVNTREF:1356132  ULREF:131sd6s5  SETTDATE:20211108  SETCUR:001  SETSUMMA:-7  TAXCUR:001  TAXSUMMA:0.25  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckDB_DOCS(for_Ex_Con(0).isn,"MT300   ","10",fBODY,1)
    
    'FOLDERS
    Call CheckQueryRowCount("FOLDERS","fISN",for_Ex_Con(0).isn,1)
    Call CheckDB_FOLDERS(dbFOLDERS(0),1) 
    
    'DOCSG
    Call CheckQueryRowCount("DOCSG", "fISN", for_Ex_Con(0).isn, 25)
    For i = 0 to for_Ex_Con(0).splitCount - 1
        Call CheckDB_DOCSG(for_Ex_Con(0).isn,"SPLIT",i ,"BENINSTP",for_Ex_Con(0).splitBenInst(i),1)
        Call CheckDB_DOCSG(for_Ex_Con(0).isn,"SPLIT",i ,"BENINSTPOP",for_Ex_Con(0).splitDataTypeBenInst(i),1)
        Call CheckDB_DOCSG(for_Ex_Con(0).isn,"SPLIT",i ,"BS",for_Ex_Con(0).splitPurSale(i),1)
        Call CheckDB_DOCSG(for_Ex_Con(0).isn,"SPLIT",i ,"CUR",for_Ex_Con(0).splitCurB(i),1)
        Call CheckDB_DOCSG(for_Ex_Con(0).isn,"SPLIT",i ,"DAGENTP",for_Ex_Con(0).splitDelAgent(i),1)
        Call CheckDB_DOCSG(for_Ex_Con(0).isn,"SPLIT",i ,"DAGENTPOP",for_Ex_Con(0).splitDataTypeDelAgent(i),1)
        Call CheckDB_DOCSG(for_Ex_Con(0).isn,"SPLIT",i ,"MEDPBANK",for_Ex_Con(0).splitIntermInst(i),1)
        Call CheckDB_DOCSG(for_Ex_Con(0).isn,"SPLIT",i ,"MEDPOP",for_Ex_Con(0).splitTypeOfInterm(i),1)
        Call CheckDB_DOCSG(for_Ex_Con(0).isn,"SPLIT",i ,"RAGENTP",for_Ex_Con(0).splitRecAgent(i),1)
        Call CheckDB_DOCSG(for_Ex_Con(0).isn,"SPLIT",i ,"RAGENTPOP",for_Ex_Con(0).splitDataTypeRecAgent(i),1)
        sumSQL = Replace (for_Ex_Con(0).splitSumB(i), ".00", "")
        Call CheckDB_DOCSG(for_Ex_Con(0).isn,"SPLIT",i ,"SUMMA",sumSQL,1)
    Next
    Call CheckDB_DOCSG(for_Ex_Con(0).isn,"SPLIT",0 ,"DAGENTPID",for_Ex_Con(0).splitPIDDelAgent(0),1)
    Call CheckDB_DOCSG(for_Ex_Con(0).isn,"SPLIT",0 ,"MEDPID",for_Ex_Con(0).splitPIDIntermInst(0),1)    
    Call CheckDB_DOCSG(for_Ex_Con(0).isn,"SPLIT",0 ,"RAGENTPID",for_Ex_Con(0).splitPIDRecAgent(0),1)
        
    'SW_MESSAGES
    dbSW_MESSAGES(0).fISN = for_Ex_Con(0).isn
           
    Call CheckQueryRowCount("SW_MESSAGES","fISN",for_Ex_Con(0).isn,1)
    Call CheckDB_SW_MESSAGES(dbSW_MESSAGES(0),1) 
'-------------------------------------------------------------------
'--------------- Կատարում է "Հաստատել" գործողությունը -------------------
'-------------------------------------------------------------------
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_Confirm)
    If wMDIClient.waitVBObject("frmASDocForm", 2000).exists Then
       Call Rekvizit_Fill("Document",1,"General", "OPERTYPE",for_Ex_Con(1).opType)
       for_Ex_Con(1).docN = Get_Rekvizit_Value("Document",1,"General","BMDOCNUM")
       Call ClickCmdButton(1, "Î³ï³ñ»É")
    End If
    Call Close_Window(wMDIClient, "frmPttel")
    'Մուտք Ուղարկվող Հաղորդագրություններ/Ուղարկվող փոխանցումներ թղթապանակ
    Call GoTo_Sending_Transfer(sending, "|S.W.I.F.T. ²Þî                  |àõÕ³ñÏíáÕ Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñ|àõÕ³ñÏíáÕ ÷áË³ÝóáõÙÝ»ñ")
    'Կատարում է Դիտել գործողությունը և ստուգում է հաղորդագրության պատուհանի բովանդակությունը
    If wMDIClient.waitVBObject("frmPttel", 4000).exists  Then
        If SearchInPttel("frmPttel", 2, for_Ex_Con(1).docN) Then
            Call wMainForm.MainMenu.Click(c_AllActions)
            Call wMainForm.PopupMenu.Click(c_View)
            for_Ex_Con(1).isn = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
            Call Check_Foreign_Exchange_Confirm_Window(for_Ex_Con(1))
            Call ClickCmdButton(1, "OK")  
            Log.Message "fISN = "& for_Ex_Con(1).isn,,,SqlDivideColor 
        Else 
            Log.Error "Document row not found",,,ErrorColor
        End If          
    Else 
        Log.Error "Pttel form not found",,, ErrorColor
    End If          
  
    Log.Message "SQL ստուգում Հաստատել գործողությունից հետո",,,SqlDivideColor
    Call Initialize_DB_SWIFT_300(for_Ex_Con(1).isn, for_Ex_Con(1).docN) 
    'DOCLOG
    Call CheckQueryRowCount("DOCLOG","fISN",for_Ex_Con(0).isn,2)
    Call CheckDB_DOCLOG(for_Ex_Con(0).isn,"77","N","10","",1)
    Call CheckDB_DOCLOG(for_Ex_Con(0).isn,"77","M","10","CREATED",1)
    Call CheckQueryRowCount("DOCLOG","fISN",for_Ex_Con(1).isn,2)
    Call CheckDB_DOCLOG(for_Ex_Con(1).isn,"77","N","1","",1)
    Call CheckDB_DOCLOG(for_Ex_Con(1).isn,"77","C","9","",1)
        
    'DOCP     
    Call CheckQueryRowCount("DOCP","fISN",for_Ex_Con(1).isn,1)
    Call CheckDB_DOCP(for_Ex_Con(1).isn,"MT300   ",for_Ex_Con(0).isn,1)
    'DOCS    
    Call CheckQueryRowCount("DOCS","fISN",for_Ex_Con(1).isn,1)
    fBODY = "  BMDOCNUM:" & for_Ex_Con(1).docN & "  REF:" & for_Ex_Con(0).docN & "  OPERTYPE:NEWT  COMREF:BEST223858XXN}:{  OPERSCOPE:AGNT  BLOCKIND:N  "_
          & "SPLITIND:Y  PARTYAOP:A  PARTYAID:125645656asd  PARTYA:CITIUS33XXX  PARTYBOP:A  PARTYBID:123546988  PARTYB:CITIUS33XXX  FUNDOP:J  "_          
          & "FUND:CITIUS33XXX  VERIFIED:0  DATE:20211108  PAYDATE:20211108  CURB:003  CURS:001  SUMMAB:72.16  SUMMAS:100  COURSE:1.3858/1  "_
          &"RATE:1,3858  SNDREC:XXN}{4:{4:  DAGENTBOP:J  DAGENTB:CITIUS33XXX  MEDBOP:A  MEDBID:645646  MEDBBANK:DGPBDE3MBRA  RAGENTBOP:J  "_
          &"RAGENTB:CITIUS33XXX  DAGENTSOP:A  DAGENTSID:464asd64s  DAGENTS:CITIUS33XXX  MEDSOP:J  MEDSBANK:CITIUS33XXX  RAGENTSOP:A  "_
          &"RAGENTSID:45646asd6564  RAGENTS:AMASJPJZXXX  CNTREF:1125s52248s5  BRKREF:21asd84646  TERMS:sadf12541a  ADDINFO:12556s54das  "_
          &"BMIODATE:" & aqConvert.DateTimeToFormatStr(aqDateTime.Now(),"%Y%m%d") & "  RSBKMAIL:0  DELIV:0  USERID:  77  PROCESSED:0  "_
          &"CONTACT:55d445s8  DEALMETH:BROK  DEALADD:3218655s85  DEALAOP:A  DEALAID:s99d56a9662  DEALA:CITIUS33XXX  DEALBOP:A  "_
          &"DEALBID:125598556asd  DEALB:CITIUS33XXX  BROKEROP:J  BROKER:BRAJINBBBYF  COUNT:2  "

    fBODY = Replace(fBODY, "  ", "%")
    Call CheckDB_DOCS(for_Ex_Con(1).isn,"MT300   ","9",fBODY,1)
    
    'DOCSG
    Call CheckQueryRowCount("DOCSG", "fISN", for_Ex_Con(1).isn, 21)
    For i = 0 to for_Ex_Con(1).splitCount - 1
        Call CheckDB_DOCSG(for_Ex_Con(1).isn,"SPLIT",i ,"BS",for_Ex_Con(1).splitPurSale(i),1)
        Call CheckDB_DOCSG(for_Ex_Con(1).isn,"SPLIT",i ,"CUR",for_Ex_Con(1).splitCurB(i),1)
        Call CheckDB_DOCSG(for_Ex_Con(1).isn,"SPLIT",i ,"DAGENTP",for_Ex_Con(1).splitDelAgent(i),1)
        Call CheckDB_DOCSG(for_Ex_Con(1).isn,"SPLIT",i ,"DAGENTPOP",for_Ex_Con(1).splitDataTypeDelAgent(i),1)
        Call CheckDB_DOCSG(for_Ex_Con(1).isn,"SPLIT",i ,"MEDPBANK",for_Ex_Con(1).splitIntermInst(i),1)
        Call CheckDB_DOCSG(for_Ex_Con(1).isn,"SPLIT",i ,"MEDPOP",for_Ex_Con(1).splitTypeOfInterm(i),1)
        Call CheckDB_DOCSG(for_Ex_Con(1).isn,"SPLIT",i ,"RAGENTP",for_Ex_Con(1).splitRecAgent(i),1)
        Call CheckDB_DOCSG(for_Ex_Con(1).isn,"SPLIT",i ,"RAGENTPOP",for_Ex_Con(1).splitDataTypeRecAgent(i),1)
        sumSQL = Replace (for_Ex_Con(1).splitSumB(i), ".00", "")
        Call CheckDB_DOCSG(for_Ex_Con(1).isn,"SPLIT",i ,"SUMMA",sumSQL,1)
    Next
    Call CheckDB_DOCSG(for_Ex_Con(1).isn,"SPLIT",0 ,"DAGENTPID",for_Ex_Con(1).splitPIDDelAgent(0),1)
    Call CheckDB_DOCSG(for_Ex_Con(1).isn,"SPLIT",0 ,"MEDPID",for_Ex_Con(1).splitPIDIntermInst(0),1)    
    Call CheckDB_DOCSG(for_Ex_Con(1).isn,"SPLIT",0 ,"RAGENTPID",for_Ex_Con(1).splitPIDRecAgent(0),1)    
    
    'SW_MESSAGES
    dbSW_MESSAGES(1).fISN = for_Ex_Con(1).isn      
    Call CheckQueryRowCount("SW_MESSAGES","fISN",for_Ex_Con(1).isn,1)
    Call CheckDB_SW_MESSAGES(dbSW_MESSAGES(1),1)
'------------------------------------------------------------------
'----------------- Ջնջում է հաստատման փաստաթուղթը --------------------
'------------------------------------------------------------------     
    Call SearchAndDelete( "frmPttel", 2, for_Ex_Con(1).isn , "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ" )
    Call Close_Window(wMDIClient, "frmPttel")

    'Մուտք Ադմինիստրարտոր ԱՇՏ
    Call ChangeWorkspace(c_Admin40)
    folderDirect = "|²¹ÙÇÝÇëïñ³ïáñÇ ²Þî 4.0|ÂÕÃ³å³Ý³ÏÝ»ñ|êï»ÕÍí³Í ÷³ëï³ÃÕÃ»ñ"
    stDate = aqDateTime.Today
    enDate = aqDateTime.Today
    wUser = 77
    docType = ""
    'Մուտք Ստեղծված փաստաթղթեր թղթապանակ
    Call OpenCreatedDocFolder(folderDirect, stDate, enDate, wUser, docType)
'------------------------------------------------------------------
'----------------- Ջնջում է ներմուծված փաստաթուղթը ---------------------
'------------------------------------------------------------------     
    Call SearchAndDelete( "frmPttel", 2, for_Ex_Con(0).isn , "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ" ) 
    Log.Message "SQL ստուգում երկու փաստաթղթերը ջնջելուց հետո",,,SqlDivideColor
    'DOCLOG
    Call CheckQueryRowCount("DOCLOG","fISN",for_Ex_Con(0).isn,4)
    Call CheckDB_DOCLOG(for_Ex_Con(0).isn,"77","N","10","",1)
    Call CheckDB_DOCLOG(for_Ex_Con(0).isn,"77","M","10","CREATED",1)
    Call CheckDB_DOCLOG(for_Ex_Con(0).isn,"77","D","999","",1)
    Call CheckDB_DOCLOG(for_Ex_Con(0).isn,"77","M","10","DELETED",1)
    
    Call CheckQueryRowCount("DOCLOG","fISN",for_Ex_Con(1).isn,3)
    Call CheckDB_DOCLOG(for_Ex_Con(1).isn,"77","N","1","",1)
    Call CheckDB_DOCLOG(for_Ex_Con(1).isn,"77","C","9","",1)
    Call CheckDB_DOCLOG(for_Ex_Con(1).isn,"77","D","999","",1)
    
    'SW_MESSAGES
    Call CheckQueryRowCount("SW_MESSAGES","fISN",for_Ex_Con(0).isn,0)
    Call CheckQueryRowCount("SW_MESSAGES","fISN",for_Ex_Con(1).isn,0)   
    
    'DOCP
    Call CheckDB_DOCP(for_Ex_Con(1).isn,"MT300   ",for_Ex_Con(0).isn,0)

                                                                                                                                                                                                                                                                                                                       Call Close_Window(wMDIClient, "frmPttel")   
'---------------------------------------------------
'------------- Կարգավորումների ներմուծում-2 -------------
'---------------------------------------------------
    Log.Message "Կարգավորումների ներմուծում-2 (300 Հաղորդագրության ավելացում SWSM պարամետրում) ",,,DivideColor
    settingsPath = Project.Path & "Stores\SWIFT\HT300\Settings\Setting_2.txt"'SWSMG - Պարամետրում ՛300՛ հաղորդագրությունը առկա է
    folderDirect = "|²¹ÙÇÝÇëïñ³ïáñÇ ²Þî 4.0|Ð³Ù³Ï³ñ·³ÛÇÝ ³ßË³ï³ÝùÝ»ñ|îíÛ³ÉÝ»ñÇ Ý»ñÙáõÍáõÙ|ö³ëï³ÃÕÃ»ñÇ Ý»ñÙáõÍáõÙ"
    Call ChangeWorkspace(c_Admin40)
    BuiltIn.Delay(3000)
    Call Settings_Import(settingsPath,folderDirect)
    Login("ARMSOFT")
    Call CheckQueryRowCount("PARAMS","fVALUE","101,110,200,201,203,102,N99,103,N98,300,410,N96",1)
'-----------------------------------------------------------------------------
'----------------- Կատարել Ընդունել SWIFT համակարգից գործողությունը ------------------
'-----------------------------------------------------------------------------
    Log.Message "Կատարել Ընդունել SWIFT համակարգից գործողությունը",,,DivideColor     
    'Մուտք գործել "S.W.I.F.T. ԱՇՏ"
    Call ChangeWorkspace(c_SWIFT)
    BuiltIn.Delay(3000)
    for_Ex_Con(1).opType = "AMND"
    max=100
    min=999
    Randomize
    rand = Int((max-min+1)*Rnd+min)
    fileFrom = Project.Path &"Stores\SWIFT\HT300\ImportFile\IA000390.RJE"
    fileTo = Project.Path &"Stores\SWIFT\HT300\ImportFile\Import\IA000391.RJE"
    what = "CITI2111089856"
    fWith = "CITI2111089" & rand
    Log.Message(fWith)
    
    'Վերարտագրում է նախօրոք տրված ֆայլը մեկ այլ ֆայլի մեջ` կատարելով փոփոխություն
    Call Read_Write_File_With_replace(fileFrom,fileTo,what,fWith)
    Call Recieve_From_SWIFT(1)        
'------------------------------------------------------------------
'---------------- Ստուգում է փաստաթղթի առկայությունը --------------------
'------------------------------------------------------------------
    Log.Message " Ստուգում է փաստաթղթի առկայությունը ",,,DivideColor       
    
    'Մուտք գործել Փոխանցումներ/Ստացված փաստաթղթեր թղթապանակ
    Call GoTo_Recieved_Messages (recieved, "|S.W.I.F.T. ²Þî                  |öáË³ÝóáõÙÝ»ñ|êï³óí³Í ÷áË³ÝóáõÙÝ»ñ")
    'Աղյուսակի տեսքի համեմատում
    Call ColumnSorting(SortArr, 2, "frmPttel")
    
    Path1 = Project.Path & "Stores\SWIFT\HT300\Actual_Pttel.txt"
    Path2 = Project.Path & "Stores\SWIFT\HT300\Expected_Pttel.txt"
    regex = "(\d{2}[/]\d{2}[/]\d{2})|(\d{2}:\d{2})|(CITI\d{10})"
    
    'Արտահանել և Ð³Ù»Ù³ï»É »ñÏáõ txt ý³ÛÉ»ñ
    Call ExportToTXTFromPttel("frmPttel",Path1)
    Call Compare_Files(Path2, Path1, regex)
    'Ստուգում է փաստաթղթի առկայությունը
    for_Ex_Con(0).docN = fWith
    for_Ex_Con(1).reference = fWith
    If SearchInPttel("frmPttel",3, for_Ex_Con(0).docN) Then
        Call wMainForm.MainMenu.Click(c_AllActions)
        Call wMainForm.PopupMenu.Click(c_View)
        'Ստուգում է հաղորդագրության պատուհանի բովանդակությունը
        Call Check_Foreign_Exchange_Confirm_Window(for_Ex_Con(0))
        Call ClickCmdButton(1, "OK")    
    Else 
        Log.Error "Document row not found",,,ErrorColor
    End If      
    Log.Message "SQL ստուգում (Ընդունել SWIFT համակարգից գործողություն)ից հետո",,,SqlDivideColor
    Log.Message "fISN = "& for_Ex_Con(0).isn ,,,SqlDivideColor
    'SQL
    Call Initialize_DB_SWIFT_300(for_Ex_Con(0).isn, for_Ex_Con(0).docN)
    'DOCLOG
    Call CheckQueryRowCount("DOCLOG","fISN",for_Ex_Con(0).isn,1)
    Call CheckDB_DOCLOG(for_Ex_Con(0).isn,"77","N","10","",1)
    
    'DOCS
    Call CheckQueryRowCount("DOCS","fISN",for_Ex_Con(0).isn,1)

    fBODY = "  BMDOCNUM:" & for_Ex_Con(0).docN & "  REF:11111  OPERTYPE:CANC  COMREF:PRML223858  OPERSCOPE:AGNT  "_
          & "BLOCKIND:N  SPLITIND:Y  PARTYAOP:A  PARTYAID:123546988  PARTYA:CITIUS33XXX  PARTYBOP:A  PARTYBID:125645656asd  "_
          & "PARTYB:CITIUS33XXX  FUNDOP:J  FUND:CITIUS33XXX  VERIFIED:0  DATE:20211108  PAYDATE:20211108  CURB:001  CURS:003  SUMMAB:100  "_
          &"SUMMAS:72.16  COURSE:1.3858/1  RATE:1,3858  SNDREC:XXN}{4:{4:  DAGENTBOP:A  DAGENTBID:464asd64s  DAGENTB:CITIUS33XXX  "_
          &"MEDBOP:J  MEDBBANK:CITIUS33XXX  RAGENTBOP:A  RAGENTBID:45646asd6564  RAGENTB:AMASJPJZXXX  DAGENTSOP:J  DAGENTS:CITIUS33XXX  MEDSOP:A  "_
          &"MEDSID:645646  MEDSBANK:DGPBDE3MBRA  RAGENTSOP:J  RAGENTS:CITIUS33XXX  BENINSTSOP:A  BENINSTSID:asdasd  BENINSTS:CITIUS33XXX  "_
          &"CNTREF:1125s52248s5  BRKREF:21asd84646  TERMS:sadf12541a  ADDINFO:12556s54das  BMNAME:IA000391#001  "_
          &"BMIODATE:" & aqConvert.DateTimeToFormatStr(aqDateTime.Now(),"%Y%m%d") & "  UNIQUEID:US33XXXXN}{4:  RSBKMAIL:1  DELIV:0  USERID:  77  "_
          &"PROCESSED:0  CONTACT:55d445s8  DEALMETH:BROK  DEALADD:3218655s85  DEALAOP:A  DEALAID:125598556asd  DEALA:CITIUS33XXX  DEALBOP:A  "_
          &"DEALBID:s99d56a9662  DEALB:CITIUS33XXX  BROKEROP:J  BROKER:BRAJINBBBYF  BRCHGCUR:001  BRCHGSUM:10  COUNT:2  EVNTTYPE:EAMT  "_
          &"EVNTREF:1356132  ULREF:131sd6s5  SETTDATE:20211108  SETCUR:001  SETSUMMA:-7  TAXCUR:001  TAXSUMMA:0.25  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckDB_DOCS(for_Ex_Con(0).isn,"MT300   ","10",fBODY,1)
    
    'FOLDERS
    Call CheckQueryRowCount("FOLDERS","fISN",for_Ex_Con(0).isn,1)
    Call CheckDB_FOLDERS(dbFOLDERS(0),1) 
    
    'DOCSG
    Call CheckQueryRowCount("DOCSG", "fISN", for_Ex_Con(0).isn, 25)
    For i = 0 to for_Ex_Con(0).splitCount - 1
        Call CheckDB_DOCSG(for_Ex_Con(0).isn,"SPLIT",i ,"BENINSTP",for_Ex_Con(0).splitBenInst(i),1)
        Call CheckDB_DOCSG(for_Ex_Con(0).isn,"SPLIT",i ,"BENINSTPOP",for_Ex_Con(0).splitDataTypeBenInst(i),1)
        Call CheckDB_DOCSG(for_Ex_Con(0).isn,"SPLIT",i ,"BS",for_Ex_Con(0).splitPurSale(i),1)
        Call CheckDB_DOCSG(for_Ex_Con(0).isn,"SPLIT",i ,"CUR",for_Ex_Con(0).splitCurB(i),1)
        Call CheckDB_DOCSG(for_Ex_Con(0).isn,"SPLIT",i ,"DAGENTP",for_Ex_Con(0).splitDelAgent(i),1)
        Call CheckDB_DOCSG(for_Ex_Con(0).isn,"SPLIT",i ,"DAGENTPOP",for_Ex_Con(0).splitDataTypeDelAgent(i),1)
        Call CheckDB_DOCSG(for_Ex_Con(0).isn,"SPLIT",i ,"MEDPBANK",for_Ex_Con(0).splitIntermInst(i),1)
        Call CheckDB_DOCSG(for_Ex_Con(0).isn,"SPLIT",i ,"MEDPOP",for_Ex_Con(0).splitTypeOfInterm(i),1)
        Call CheckDB_DOCSG(for_Ex_Con(0).isn,"SPLIT",i ,"RAGENTP",for_Ex_Con(0).splitRecAgent(i),1)
        Call CheckDB_DOCSG(for_Ex_Con(0).isn,"SPLIT",i ,"RAGENTPOP",for_Ex_Con(0).splitDataTypeRecAgent(i),1)
        sumSQL = Replace (for_Ex_Con(0).splitSumB(i), ".00", "")
        Call CheckDB_DOCSG(for_Ex_Con(0).isn,"SPLIT",i ,"SUMMA",sumSQL,1)
    Next
    Call CheckDB_DOCSG(for_Ex_Con(0).isn,"SPLIT",0 ,"DAGENTPID",for_Ex_Con(0).splitPIDDelAgent(0),1)
    Call CheckDB_DOCSG(for_Ex_Con(0).isn,"SPLIT",0 ,"MEDPID",for_Ex_Con(0).splitPIDIntermInst(0),1)    
    Call CheckDB_DOCSG(for_Ex_Con(0).isn,"SPLIT",0 ,"RAGENTPID",for_Ex_Con(0).splitPIDRecAgent(0),1)
        
    'SW_MESSAGES
    dbSW_MESSAGES(0).fISN = for_Ex_Con(0).isn         
    Call CheckQueryRowCount("SW_MESSAGES","fISN",for_Ex_Con(0).isn,1)
    Call CheckDB_SW_MESSAGES(dbSW_MESSAGES(0),1) 
'-------------------------------------------------------------------
'--------------- Կատարում է "Հաստատել" գործողությունը -------------------
'-------------------------------------------------------------------    
    
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_Confirm)
    If wMDIClient.waitVBObject("frmASDocForm", 2000).exists Then
       Call Rekvizit_Fill("Document",1,"General", "OPERTYPE",for_Ex_Con(1).opType)
       for_Ex_Con(1).docN = Get_Rekvizit_Value("Document",1,"General","BMDOCNUM")
       Call ClickCmdButton(1, "Î³ï³ñ»É")
    End If
    Call Close_Window(wMDIClient, "frmPttel")
    'Մուտք Ուղարկվող Հաղորդագրություններ/Ուղարկվող փոխանցումներ թղթապանակ
    Call GoTo_Sending_Transfer(sending, "|S.W.I.F.T. ²Þî                  |àõÕ³ñÏíáÕ Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñ|àõÕ³ñÏíáÕ ÷áË³ÝóáõÙÝ»ñ")
    'Կատարում է Դիտել գործողությունը և ստուգում հաղորդագրության պատուհանի բովանդակությունը
    If wMDIClient.waitVBObject("frmPttel", 4000).exists  Then
        If SearchInPttel("frmPttel", 2, for_Ex_Con(1).docN) Then
            Call wMainForm.MainMenu.Click(c_AllActions)
            Call wMainForm.PopupMenu.Click(c_View)
            
'            for_Ex_Con(1).isn = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
            Call Check_Foreign_Exchange_Confirm_Window(for_Ex_Con(1))
            Call ClickCmdButton(1, "OK")  
            Log.Message "fISN = "& for_Ex_Con(1).isn,,,SqlDivideColor 
        Else 
            Log.Error "Document row not found",,,ErrorColor
        End If          
    Else 
        Log.Error "Pttel form not found",,, ErrorColor
    End If          
    
    Log.Message "SQL ստուգում Հաստատել գործողությունից հետո",,,SqlDivideColor
    Call Initialize_DB_SWIFT_300(for_Ex_Con(1).isn, for_Ex_Con(1).docN) 
    'DOCLOG
    Call CheckQueryRowCount("DOCLOG","fISN",for_Ex_Con(0).isn,2)
    Call CheckDB_DOCLOG(for_Ex_Con(0).isn,"77","N","10","",1)
    Call CheckDB_DOCLOG(for_Ex_Con(0).isn,"77","M","10","CREATED",1)
    Call CheckQueryRowCount("DOCLOG","fISN",for_Ex_Con(1).isn,2)
    Call CheckDB_DOCLOG(for_Ex_Con(1).isn,"77","N","1","",1)
    Call CheckDB_DOCLOG(for_Ex_Con(1).isn,"77","C","9","",1)
        
    'DOCP     
    Call CheckQueryRowCount("DOCP","fISN",for_Ex_Con(1).isn,1)
    Call CheckDB_DOCP(for_Ex_Con(1).isn,"MT300   ",for_Ex_Con(0).isn,1)
    'DOCS    
    Call CheckQueryRowCount("DOCS","fISN",for_Ex_Con(1).isn,1)
    fBODY = "  BMDOCNUM:" & for_Ex_Con(1).docN & "  REF:" & for_Ex_Con(0).docN & "  OPERTYPE:AMND  COMREF:BEST223858XXN}:{  OPERSCOPE:AGNT  BLOCKIND:N  "_
          & "SPLITIND:Y  PARTYAOP:A  PARTYAID:125645656asd  PARTYA:CITIUS33XXX  PARTYBOP:A  PARTYBID:123546988  PARTYB:CITIUS33XXX  FUNDOP:J  "_          
          & "FUND:CITIUS33XXX  VERIFIED:0  DATE:20211108  PAYDATE:20211108  CURB:003  CURS:001  SUMMAB:72.16  SUMMAS:100  COURSE:1.3858/1  "_
          &"RATE:1,3858  SNDREC:XXN}{4:{4:  DAGENTBOP:J  DAGENTB:CITIUS33XXX  MEDBOP:A  MEDBID:645646  MEDBBANK:DGPBDE3MBRA  RAGENTBOP:J  "_
          &"RAGENTB:CITIUS33XXX  DAGENTSOP:A  DAGENTSID:464asd64s  DAGENTS:CITIUS33XXX  MEDSOP:J  MEDSBANK:CITIUS33XXX  RAGENTSOP:A  "_
          &"RAGENTSID:45646asd6564  RAGENTS:AMASJPJZXXX  CNTREF:1125s52248s5  BRKREF:21asd84646  TERMS:sadf12541a  ADDINFO:12556s54das  "_
          &"BMIODATE:" & aqConvert.DateTimeToFormatStr(aqDateTime.Now(),"%Y%m%d") & "  RSBKMAIL:0  DELIV:0  USERID:  77  PROCESSED:0  "_
          &"CONTACT:55d445s8  DEALMETH:BROK  DEALADD:3218655s85  DEALAOP:A  DEALAID:s99d56a9662  DEALA:CITIUS33XXX  DEALBOP:A  "_
          &"DEALBID:125598556asd  DEALB:CITIUS33XXX  BROKEROP:J  BROKER:BRAJINBBBYF  COUNT:2  "

    fBODY = Replace(fBODY, "  ", "%")
    Call CheckDB_DOCS(for_Ex_Con(1).isn,"MT300   ","9",fBODY,1)
    
    'DOCSG
    Call CheckQueryRowCount("DOCSG", "fISN", for_Ex_Con(1).isn, 21)
    For i = 0 to for_Ex_Con(1).splitCount - 1
        Call CheckDB_DOCSG(for_Ex_Con(1).isn,"SPLIT",i ,"BS",for_Ex_Con(1).splitPurSale(i),1)
        Call CheckDB_DOCSG(for_Ex_Con(1).isn,"SPLIT",i ,"CUR",for_Ex_Con(1).splitCurB(i),1)
        Call CheckDB_DOCSG(for_Ex_Con(1).isn,"SPLIT",i ,"DAGENTP",for_Ex_Con(1).splitDelAgent(i),1)
        Call CheckDB_DOCSG(for_Ex_Con(1).isn,"SPLIT",i ,"DAGENTPOP",for_Ex_Con(1).splitDataTypeDelAgent(i),1)
        Call CheckDB_DOCSG(for_Ex_Con(1).isn,"SPLIT",i ,"MEDPBANK",for_Ex_Con(1).splitIntermInst(i),1)
        Call CheckDB_DOCSG(for_Ex_Con(1).isn,"SPLIT",i ,"MEDPOP",for_Ex_Con(1).splitTypeOfInterm(i),1)
        Call CheckDB_DOCSG(for_Ex_Con(1).isn,"SPLIT",i ,"RAGENTP",for_Ex_Con(1).splitRecAgent(i),1)
        Call CheckDB_DOCSG(for_Ex_Con(1).isn,"SPLIT",i ,"RAGENTPOP",for_Ex_Con(1).splitDataTypeRecAgent(i),1)
        sumSQL = Replace (for_Ex_Con(1).splitSumB(i), ".00", "")
        Call CheckDB_DOCSG(for_Ex_Con(1).isn,"SPLIT",i ,"SUMMA",sumSQL,1)
    Next
    Call CheckDB_DOCSG(for_Ex_Con(1).isn,"SPLIT",0 ,"DAGENTPID",for_Ex_Con(1).splitPIDDelAgent(0),1)
    Call CheckDB_DOCSG(for_Ex_Con(1).isn,"SPLIT",0 ,"MEDPID",for_Ex_Con(1).splitPIDIntermInst(0),1)    
    Call CheckDB_DOCSG(for_Ex_Con(1).isn,"SPLIT",0 ,"RAGENTPID",for_Ex_Con(1).splitPIDRecAgent(0),1)    
    
    'SW_MESSAGES
    dbSW_MESSAGES(1).fISN = for_Ex_Con(1).isn      
    Call CheckQueryRowCount("SW_MESSAGES","fISN",for_Ex_Con(1).isn,1)
    Call CheckDB_SW_MESSAGES(dbSW_MESSAGES(1),1)
'------------------------------------------------------------------
'----------------- Ջնջում է հաստատման փաստաթուղթը --------------------
'------------------------------------------------------------------      
    Call SearchAndDelete( "frmPttel", 2, for_Ex_Con(1).isn , "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ" )
    Call Close_Window(wMDIClient, "frmPttel")
    
    'Մուտք Ադմինիստրարտոր ԱՇՏ
    Call ChangeWorkspace(c_Admin40)
    folderDirect = "|²¹ÙÇÝÇëïñ³ïáñÇ ²Þî 4.0|ÂÕÃ³å³Ý³ÏÝ»ñ|êï»ÕÍí³Í ÷³ëï³ÃÕÃ»ñ"
    stDate = aqDateTime.Today
    enDate = aqDateTime.Today
    wUser = 77
    docType = ""
    'Մուտք Ստեղծված փաստաթղթեր թղթապանակ
    Call OpenCreatedDocFolder(folderDirect, stDate, enDate, wUser, docType)
'------------------------------------------------------------------
'----------------- Ջնջում է ներմուծված փաստաթուղթը ---------------------
'------------------------------------------------------------------     
    Call SearchAndDelete( "frmPttel", 2, for_Ex_Con(0).isn , "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ" ) 

    Log.Message "SQL ստուգում երկու փաստաթղթերը ջնջելուց հետո",,,SqlDivideColor
    'DOCLOG
    Call CheckQueryRowCount("DOCLOG","fISN",for_Ex_Con(0).isn,4)
    Call CheckDB_DOCLOG(for_Ex_Con(0).isn,"77","N","10","",1)
    Call CheckDB_DOCLOG(for_Ex_Con(0).isn,"77","M","10","CREATED",1)
    Call CheckDB_DOCLOG(for_Ex_Con(0).isn,"77","D","999","",1)
    Call CheckDB_DOCLOG(for_Ex_Con(0).isn,"77","M","10","DELETED",1)
    
    Call CheckQueryRowCount("DOCLOG","fISN",for_Ex_Con(1).isn,3)
    Call CheckDB_DOCLOG(for_Ex_Con(1).isn,"77","N","1","",1)
    Call CheckDB_DOCLOG(for_Ex_Con(1).isn,"77","C","9","",1)
    Call CheckDB_DOCLOG(for_Ex_Con(1).isn,"77","D","999","",1)
    
    'SW_MESSAGES
    Call CheckQueryRowCount("SW_MESSAGES","fISN",for_Ex_Con(0).isn,0)
    Call CheckQueryRowCount("SW_MESSAGES","fISN",for_Ex_Con(1).isn,0)   
    
    'DOCP
    Call CheckDB_DOCP(for_Ex_Con(1).isn,"MT300   ",for_Ex_Con(0).isn,0)

    Call Close_Window(wMDIClient, "frmPttel")       
    Call Close_AsBank()
End Sub


Sub Test_Initialize_SWIFT_300 ()
    SortArr(0) = "BMNAME"
    SortArr(1) = "DOCNUM"
    
    Set recieved = New_Recieved()
    With recieved
        .sDate = aqDateTime.Today
        .eDate = aqDateTime.Today
        .mt = "300"
    End With
    
    Set sending = New_Sending()
    With sending
        .sDate = "010120"
        .eDate = aqDateTime.Today
        .mt = "300"
    End With
    Set for_Ex_Con(0) = New_Foreign_Exch_Confirm
    With for_Ex_Con(0)
        '1
        .reference = "11111"
        .opType = "CANC"
        .comReference = "PRML223858"
        .opScope = "AGNT"
        .blockTrade = "N"
        .splitSettlemnt = "Y"
        .dataTypePartyA = "A"
        .PIDPartyA = "123546988"
        .partyA = "CITIUS33XXX"
        .dataTypePartyB = "A"
        .PIDPartyB = "125645656asd"
        .partyB = "CITIUS33XXX"
        .dataTypeFund = "J"
        .fund = "CITIUS33XXX"
        '2
        .date = "08/11/21"
        .valDate = "08/11/21"
        .curB = "001"
        .curS = "003"
        .sumB = "100.00"
        .sumS = "72.16"
        .exCourse = "1.3858/1"
        .exRate = "1,3858      "
        .senderReciever = ""
        '3
        .bDataTypeDelAgent  = "A"
        .bPIDDelAgent = "464asd64s"
        .bDelAgent = "CITIUS33XXX"
        .bDataTypeIntInst = "J"
        .bIntInst = "CITIUS33XXX"
        .bDataTypeRecieveAgent = "A"
        .bPIDRecieveAgent = "45646asd6564"
        .bRecieveAgent = "AMASJPJZXXX"
        '4
        .sDataTypeDelAgent  = "J"
        .sDelAgent = "CITIUS33XXX"
        .sTypeIntInst = "A"
        .sPIDIntInst = "645646"
        .sIntInst = "DGPBDE3MBRA"
        .sDataTypeRecieveAgent = "J"
        .sRecieveAgent = "CITIUS33XXX"
        .sDataTypeBenefInst = "A"
        .sPIDBenefInst = "asdasd"
        .sBenefInst = "CITIUS33XXX"
        '5
        .counterpartyRef = "1125s52248s5"
        .brokRef = "21asd84646"
        .terms = "sadf12541a"
        .senderToRecInfo = "12556s54das"
        .fileName = "IA000391#001"
        .sendRecDate = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%d/%m/%y")
        '6
        .contInfo = "55d445s8"
        .dealMethod = "BROK"
        .dealMethodAdd = "3218655s85"
        .dataTypeDealingPartyA = "A"
        .PIDDealingPartyA = "125598556asd"
        .dealingPartyA = "CITIUS33XXX"
        .dataTypeDealingPartyB = "A"
        .PIDDealingPartyB = "s99d56a9662"
        .dealingPartyB = "CITIUS33XXX"
        .dataTypeBrok = "J"
        .broker = "BRAJINBBBYF"
        .brokerComissionCur = "001"
        .brokerComission = "10.00"
        '7
        .splitCount = "2"
        .splitPurSale(0) = "N"
        .splitCurB (0) = "001"
        .splitSumB (0) = "10.00"
        .splitDataTypeDelAgent(0) = "A"
        .splitPIDDelAgent(0) = "15315sd"
        .splitDelAgent(0) = "/ABIC/CITIUS33XXX"
        .splitTypeOfInterm (0) = "A"
        .splitPIDIntermInst (0) = "12564564asdf"
        .splitIntermInst (0) = "BROMITRDXXX"
        .splitDataTypeRecAgent (0) = "A"
        .splitPIDRecAgent (0) = "548sdf48sd4f6"
        .splitRecAgent (0) = "CITIUS33XXX"
        .splitDataTypeBenInst (0) = "J"
        .splitPIDBenInst (0) = ""
        .splitBenInst (0) = "CITIUS33XXX"   
        .splitPurSale(1) = "Y"
        .splitCurB (1) = "003"
        .splitSumB (1) = "100.00"
        .splitDataTypeDelAgent(1) = "J"
        .splitPIDDelAgent(1) = ""
        .splitDelAgent(1) = "/ABIC/CITIUS33XXX"
        .splitTypeOfInterm (1) = "J"
        .splitPIDIntermInst (1) = ""
        .splitIntermInst (1) = "CITIUS33XXX"
        .splitDataTypeRecAgent (1) = "J"
        .splitPIDRecAgent (1) = ""
        .splitRecAgent (1) = "CITIUS33XXX"
        .splitDataTypeBenInst (1) = "J"
        .splitPIDBenInst (1) = ""
        .splitBenInst (1) = "CITIUS33XXX"
        '8
        .eventType = "EAMT"
        .eventReference = "1356132"
        .eventReference21F = "131sd6s5"
        .profLossSettDate = "08/11/21"
        .currToBeSettled = "001"
        .sumToBeSetelled = "-7.00"
        .reportCur = "001"
        .taxSum = "0.25"
    End With
    
    Set for_Ex_Con(1) = New_Foreign_Exch_Confirm
    With for_Ex_Con(1)
        '1
        .opType = "NEWT"
        .comReference = "BEST223858XXN}:{"
        .opScope = "AGNT"
        .blockTrade = "N"
        .splitSettlemnt = "Y"
        .dataTypePartyA = "A"
        .PIDPartyA = "125645656asd"
        .partyA = "CITIUS33XXX"
        .dataTypePartyB = "A"
        .PIDPartyB = "123546988"
        .partyB = "CITIUS33XXX"
        .dataTypeFund = "J"
        .fund = "CITIUS33XXX"
        '2
        .date = "08/11/21"
        .valDate = "08/11/21"
        .curB = "003"
        .curS = "001"
        .sumB = "72.16"
        .sumS = "100.00"
        .exCourse = "1.3858/1"
        .exRate = "1,3858      "
        .senderReciever = ""
        '3
        .bDataTypeDelAgent  = "J"
        .bDelAgent = "CITIUS33XXX"
        .bDataTypeIntInst = "A"
        .bPIDIntInst = "645646"
        .bIntInst = "DGPBDE3MBRA"
        .bDataTypeRecieveAgent = "J"
        .bRecieveAgent = "CITIUS33XXX"
        
        '4
        .sDataTypeDelAgent  = "A"
        .sPIDDelAgent = "464asd64s"
        .sDelAgent = "CITIUS33XXX"
        .sTypeIntInst = "J"
        .sIntInst = "CITIUS33XXX"
        .sDataTypeRecieveAgent = "A"
        .sPIDRecieveAgent = "45646asd6564"
        .sRecieveAgent = "AMASJPJZXXX"
        '5
        .counterpartyRef = "1125s52248s5"
        .brokRef = "21asd84646"
        .terms = "sadf12541a"
        .senderToRecInfo = "12556s54das"
        .fileName = ""
        .sendRecDate = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%d/%m/%y")
        '6
        .contInfo = "55d445s8"
        .dealMethod = "BROK"
        .dealMethodAdd = "3218655s85"
        .dataTypeDealingPartyA = "A"
        .PIDDealingPartyA = "s99d56a9662"
        .dealingPartyA = "CITIUS33XXX"
        .dataTypeDealingPartyB = "A"
        .PIDDealingPartyB = "125598556asd"
        .dealingPartyB = "CITIUS33XXX"
        .dataTypeBrok = "J"
        .broker = "BRAJINBBBYF"
        '7
        .splitCount = "2"
        .splitPurSale(0) = "Y"
        .splitCurB (0) = "001"
        .splitSumB (0) = "10.00"
        .splitDataTypeDelAgent(0) = "A"
        .splitPIDDelAgent(0) = "548sdf48sd4f6"
        .splitDelAgent(0) = "CITIUS33XXX"
        .splitTypeOfInterm (0) = "A"
        .splitPIDIntermInst (0) = "12564564asdf"
        .splitIntermInst (0) = "BROMITRDXXX"
        .splitDataTypeRecAgent (0) = "A"
        .splitPIDRecAgent (0) = "15315sd"
        .splitRecAgent (0) = "/ABIC/CITIUS33XXX"
        .splitDataTypeBenInst (0) = ""
        .splitPIDBenInst (0) = ""
        .splitBenInst (0) = ""   
        
        .splitPurSale(1) = "N"
        .splitCurB (1) = "003"
        .splitSumB (1) = "100.00"
        .splitDataTypeDelAgent(1) = "J"
        .splitPIDDelAgent(1) = ""
        .splitDelAgent(1) = "CITIUS33XXX"
        .splitTypeOfInterm (1) = "J"
        .splitPIDIntermInst (1) = ""
        .splitIntermInst (1) = "CITIUS33XXX"
        .splitDataTypeRecAgent (1) = "J"
        .splitPIDRecAgent (1) = ""
        .splitRecAgent (1) = "/ABIC/CITIUS33XXX"
        .splitDataTypeBenInst (1) = ""
        .splitPIDBenInst (1) = ""
        .splitBenInst (1) = ""
    End With


    sDATE = "20020101"
    fDATE = "20260101"
End Sub

Sub Initialize_DB_SWIFT_300(fISN, docN)

    Set dbFOLDERS(0) = New_DB_FOLDERS()
    With dbFOLDERS(0)
      .fFOLDERID = "EPSFOREX.20211108"
      .fNAME = "MT300   "
      .fKEY = fISN
      .fISN = fISN
      .fSTATUS = "0"
      .fCOM = "²ñï³ñÅáõÛÃÇ ÷áË³Ý³ÏÙ³Ý Ñ³ëï³ïáõÙ"
      .fSPEC = "20211108" & docN & "  001          100.00003           72.162"
      .fECOM = ""
      .fDCBRANCH = "   "
      .fDCDEPART = "   "
    End With
    
    
    Set dbSW_MESSAGES(0) = New_SW_MESSAGES()
    With dbSW_MESSAGES(0)
       .fUNIQUEID = "US33XXXXN}{4:               "
       .fDATE = "20211108"
       .fMT = "300"
       .fCATEGORY = "1"
       .fDOCNUM = docN & "  "
       .fSR = "2"
       .fSRBANK = "XXN}{4:{4: "
       .fSYS = "1"
       .fSTATE = "10"
       .fUSER = "77"
       .fACCDB = "                                  "
       .fACCCR = "                                  "
       .fAMOUNT = "72.16"
       .fCURR = "003"
       .fPAYER = "                                "
       .fRECEIVER = "                                "
       .fAIM = "12556s54das                     "
       .fBRANCH = "   "
       .fDEPART = "   "
    End With
    
    Set dbSW_MESSAGES(1) = New_SW_MESSAGES()
    With dbSW_MESSAGES(1)
       .fUNIQUEID = "ISN" & fISN
       .fDATE = "20211108"
       .fMT = "300"
       .fCATEGORY = "1"
       .fDOCNUM = docN & "      "
       .fSR = "0"
       .fSRBANK = "XXN}{4:{4: "
       .fSYS = "1"
       .fSTATE = "  "
       .fUSER = "77"
       .fACCDB = "                                  "
       .fACCCR = "                                  "
       .fAMOUNT = "72.16"
       .fCURR = "003"
       .fPAYER = "                                "
       .fRECEIVER = "                                "
       .fAIM = "12556s54das                     "
       .fBRANCH = "   "
       .fDEPART = "   "
    End With        

End Sub
