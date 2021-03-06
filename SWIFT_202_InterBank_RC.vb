'USEUNIT SWIFT_International_Payorder_Library
'USEUNIT Library_Common
'USEUNIT Library_Colour
'USEUNIT Library_CheckDB
'USEUNIT Constants
'USEUNIT Library_Contracts
'USEUNIT Main_Accountant_Filter_Library
'USEUNIT Payment_Except_Library
Option Explicit
'Test Case ID 185233
Dim sDATE, fDATE, max, min, rand, fileFrom, fileTo,what, fWith, isn, docN, folderDirect, stDate, enDate, wUser, docType, messageTime, recieved
Dim confIsn, regex, pathAct, pathExp, savePath, messageIsn, messageDocN, fBODY, dbSW_MESSAGES(1), query, confirmation, messageDesc, toSend

Sub Swift_202_InterBank_RC_Test()

    Call Test_Initialize_SWIFT_202_RC()
    Call Initialize_AsBank("bank", sDATE, fDATE)
    'Համակարգ մուտք գործել ARMSOFT օգտագործողով
    Login("ARMSOFT")

'-----------------------------------------------------------------------------
'------ "S.W.I.F.T. ԱՇՏ/Պարամետրեր"-ում կատարել համապատասխան փոփոխությունները-------
'-----------------------------------------------------------------------------
    Log.Message "-- S.W.I.F.T. ԱՇՏ/Պարամետրեր-ում կատարել համապատասխան փոփոխությունները --",,,DivideColor  
    'Մուտք գործել "S.W.I.F.T. ԱՇՏ"
    Call ChangeWorkspace(c_SWIFT)
    Call wTreeView.DblClickItem("|S.W.I.F.T. ²Þî                  |ä³ñ³Ù»ïñ»ñ")
    BuiltIn.Delay(3000)
    'Նոր փաստաթղթի համարի գեներացում
    min=100
    max=999
    Randomize
    rand = Int((min-max+1)*Rnd+max)
    fileFrom = Project.Path &"Stores\SWIFT\HT202\ImportFile\IA000392.RJE"
    fileTo = Project.Path &"Stores\SWIFT\HT202\ImportFile\Import\IA000393.RJE"
    what = "PRCB2109079860"
    fWith = "PRCB2109079" & rand
    'SWGPI Պարամետրի փոփոխում SQL հարցման միջոցով
    Call SetParameter("SWGPI", "1")
    'SWSPFSACKDIR պարամետրի փոփոխում
    Call SetParameter("SWSPFSACKDIR", "")
    'SWSPFSNAKDIR պարամետրի փոփոխում
    Call SetParameter("SWSPFSNAKDIR", "")
    'SWSPFSOUT պարամետրի փոփոխում
    Call SetParameter("SWSPFSOUT", "")
    'Ջնջում է Import թղթապանակի պարունակությունը
    aqFileSystem.DeleteFile(Project.Path &"Stores\SWIFT\HT202\ImportFile\Import\*")
    Log.Message(fWith)
    'Վերարտագրում է նախօրոք տրված ֆայլը մեկ այլ ֆայլի մեջ` կատարելով փոփոխություն
    Call Read_Write_File_With_replace(fileFrom,fileTo,what,fWith)    
    'SWOUT Պարամետրի խմբագրում
    Call SetParameter_InPttel("SWOUT",Project.Path & "Stores\SWIFT\HT202\ImportFile\Import\")
    'Փակել Պարամետրեր թղթապանակը
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
'---------- Կատարել Հաստատում(910/Կրեդիտ) գործողությունը -----------------
'-------------------------------------------------------------------
    Log.Message "Կատարել Հաստատում (910/Կրեդիտ) գործողությունը",,,DivideColor       
    
    'Մուտք գործել Փոխանցումներ/Ստացված փոխանցումներ թղթապանակ
    Call GoTo_Recieved_Messages (recieved, "|S.W.I.F.T. ²Þî                  |öáË³ÝóáõÙÝ»ñ|êï³óí³Í ÷áË³ÝóáõÙÝ»ñ")
    BuiltIn.Delay(4000) 
    docN = fWith
    'Ստուգում է փաստաթղթի առկայությունը
    If SearchInPttel("frmPttel",3, docN) Then
        'Ստանում է փաստաթղթի isn-ը
        isn = GetIsn()
        Log.Message "fISN = "& isn,,,SqlDivideColor
        'Կատարել Հաստատում (Դ/Կ) գործողությունը
        Call wMainForm.MainMenu.Click(c_AllActions)
        Call wMainForm.PopupMenu.Click(c_ConfirmDK)
        If asbank.waitVBObject("frmAsUstPar",1500).Exists Then
            Call Rekvizit_Fill("Dialog", 1, "General", "MT" , "910")
            Call ClickCmdButton(2, "Î³ï³ñ»É") 
        Else
            Log.Error "Confirmation Dialog not found",,,ErrorColor
        End If 
    Else 
        Log.Error"Document "& docN &" Not Found"
    End If    
    
    If wMDIClient.WaitVBObject("frmASDocForm",4000).Exists Then
        'Հաստատման հաղորդագրության Համարի և Isn-ի ստացում
        confirmation.NumberOfDocument = Get_Rekvizit_Value("Document",1,"General","BMDOCNUM")
        confIsn = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
        Log.Message "Confirmation fISN = "& confIsn,,,SqlDivideColor
        Call ClickCmdButton(1, "Î³ï³ñ»É")        
    Else 
        Log.Error "Document window not found",,,ErrorColor
    End If
    'Փակել Ստացված փոխանցումներ թղթապանակը
    Call Close_Window(wMDIClient, "frmPttel")
    
    'SQL
    Call Initialize_DB_SW202_RC(confIsn, confirmation.NumberOfDocument)
    Log.Message "SQL Ստուգումներ Հաստատում (Դ/Կ) գործողությունը կատարելուց հետո",,,SqlDivideColor
    'DOCLOG
    Call CheckQueryRowCount("DOCLOG","fISN",confIsn,2)
    Call CheckDB_DOCLOG(confIsn,"77","C","9"," ",1)
    Call CheckDB_DOCLOG(confIsn,"77","N","1"," ",1)
    'DOCS                                       
    fBODY = "  MT:910  BMDOCNUM:" & confirmation.NumberOfDocument & "  REFERENCE:VTBR2109079860  ACC:777000000003  DATE:20210907  CUR:001  SUMMA:300  PINSTOP:A  "_
            &"PAYINST:POALILITXXX  VERIFIED:0  USERID:  77  BMIODATE:"& aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%Y%m%d") &"  RSBKMAIL:0  DELIV:0  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",confIsn,1)
    Call CheckDB_DOCS(confIsn,"MT900   ","9",fBODY,1)
    'DOCP
    Call CheckQueryRowCount("DOCP","fISN",confIsn,1)
    Call CheckDB_DOCP(confIsn,"MT900   ",isn,1)
    'SW_MESSAGES
    dbSW_MESSAGES(0).fDOCNUM = confirmation.NumberOfDocument
    dbSW_MESSAGES(0).fISN = confIsn      
    Call CheckQueryRowCount("SW_MESSAGES","fISN",confIsn,1)
    Call CheckDB_SW_MESSAGES(dbSW_MESSAGES(0),1)
    
    'Անցնել Ուղարկվող հաղորդագրություններ/Ուղարկվող խառը հաղորդագրություններ թղթապանակ   
    Call GoTo_Sending_Messages(toSend, "|S.W.I.F.T. ²Þî                  |àõÕ³ñÏíáÕ Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñ|àõÕ³ñÏíáÕ Ë³éÁ Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñ")
    'Ստուգել Հաստատման հաղորդագրության "910" առկայությունը
    If SearchInPttel("frmPttel", 2, confirmation.NumberOfDocument) Then
        Call wMainForm.MainMenu.Click(c_AllActions)
        Call wMainForm.PopupMenu.Click(c_View)    
        If wMDIClient.WaitVBObject("frmASDocForm",4000).Exists Then
            Call Check_Confirmation(confirmation)
            Call ClickCmdButton(1, "OK") 
        Else 
            Log.Error "Document window not found",,,ErrorColor
        End If
    Else 
        Log.Error"Document Row Not Found"
    End If       
    'Ջնջել Հաստատման հաղորդագրութունը "910" 
    Call SearchAndDelete ( "frmPttel", 2, confirmation.NumberOfDocument, "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ" )
    'Փակել Ուղարկվող խառը հաղորդագրություններ թղթապանակը
    Call Close_Window(wMDIClient, "frmPttel")
    
    'SQL
    Log.Message "SQL Ստուգումներ MT900 հաղորդագրությունը ջնջելուց հետո հետո",,,SqlDivideColor
    'DOCLOG
    Call CheckQueryRowCount("DOCLOG","fISN",confIsn,3)
    Call CheckDB_DOCLOG(confIsn,"77","C","9"," ",1)
    Call CheckDB_DOCLOG(confIsn,"77","N","1"," ",1)
    Call CheckDB_DOCLOG(confIsn,"77","D","999"," ",1)
    
    Call CheckQueryRowCount("DOCLOG","fISN",isn,3)
    Call CheckDB_DOCLOG(isn,"77","N","33"," ",1)
    Call CheckDB_DOCLOG(isn,"77","M","10","Received",1)
    Call CheckDB_DOCLOG(isn,"77","M","10","DELETED",1)
    
    'DOCS                                       
    fBODY = "  MT:910  BMDOCNUM:" & confirmation.NumberOfDocument & "  REFERENCE:VTBR2109079860  ACC:777000000003  DATE:20210907  CUR:001  SUMMA:300  PINSTOP:A  "_
            &"PAYINST:POALILITXXX  VERIFIED:0  USERID:  77  BMIODATE:"& aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%Y%m%d") &"  RSBKMAIL:0  DELIV:0  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",confIsn,1)
    Call CheckDB_DOCS(confIsn,"MT900   ","999",fBODY,1)
    'DOCP
    Call CheckQueryRowCount("DOCP","fISN",confIsn,0)

'-------------------------------------------------------------------
'------------- Կատարել Դիտել փաստաթուղթը գործողությունը -------------------
'-------------------------------------------------------------------
    Log.Message "Կատարել Դիտել փաստաթուղթը գործողությունը",,,DivideColor      
    'Մուտք գործել Փոխանցումներ/Ստացված փոխանցումներ թղթապանակ
    Call GoTo_Recieved_Messages (recieved, "|S.W.I.F.T. ²Þî                  |öáË³ÝóáõÙÝ»ñ|êï³óí³Í ÷áË³ÝóáõÙÝ»ñ")
    regex="(PRCB2109079\d{3})|([[].{10}])|(\d{2}[/]\d{2}[/]\d{2})|(\d{2}:\d{2})"
    If SearchInPttel("frmPttel", 3, docN) Then
        Call View_Doc_Action (savePath, "Actual.txt", pathExp, regex)   
    Else 
        Log.Error "Document Row Not found",,,ErrorColor        
    End If
'-------------------------------------------------------------------
'-------- Կատարել Պատասխանել (Gpi) (Մերժել) գործողությունը ---------------
'-------------------------------------------------------------------   
    If SearchInPttel("frmPttel", 3, docN ) Then
        Call wMainForm.MainMenu.Click(c_AllActions)    
        Call wMainForm.PopupMenu.Click(c_AnswerGpi)     
        If asbank.waitVBObject("frmAsUstPar",1500).Exists Then
            Call Rekvizit_Fill("Dialog" , 1, "General", "RESPONSE" , "RJCT")
            Call ClickCmdButton(2, "Î³ï³ñ»É")
            If wMDIClient.WaitVBObject("frmASDocForm",4000).Exists Then 
                'Պատասխան հաղորդագրության Համարի և Isn-ի ստացում
                messageDocN = Get_Rekvizit_Value("Document",1,"General","BMDOCNUM")
                messageIsn = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn 
                messageDesc = Get_Rekvizit_Value("Document",1,"Comment","DESCRIPT")
                Call ClickCmdButton(1, "Î³ï³ñ»É")
            Else 
                Log.Error "Document window not found",,,ErrorColor
            End If
        Else
            Log.Error "Response Dialog not found",,,ErrorColor
        End If
    Else 
        Log.Error"Document Row Not Found" 
    End If 
    
    Log.Message "Message fISN = "& messageIsn,,,SqlDivideColor
    
    'SQL
    Call Initialize_DB_SW202_RC(messageIsn, messageDocN)
    Log.Message "SQL Ստուգումներ Պատասխանել(Gpi) գործողությունը կատարելուց հետո",,,SqlDivideColor
    'DOCLOG
    Call CheckQueryRowCount("DOCLOG","fISN",messageIsn,2)
    Call CheckDB_DOCLOG(messageIsn,"77","C","9"," ",1)
    Call CheckDB_DOCLOG(messageIsn,"77","N","1"," ",1)
    'DOCS                                       
    fBODY = "  CATEGORY:2  BMDOCNUM:" & messageDocN & "  REFERENCE:" & docN & "  "_
            &"DESCRIPT://"& aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%y%m%d") &"  +0400  //RJCT  //BESTAM22XXX  //USD300,"_
            &"  VERIFIED:0  USERID:  77  BMIODATE:"& aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%Y%m%d") &"  RSBKMAIL:0  DELIV:0  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",messageIsn,1)
    Call CheckDB_DOCS(messageIsn,"MTN99   ","9",fBODY,1)
    'DOCP
    Call CheckQueryRowCount("DOCP","fISN",messageIsn,1)
    Call CheckDB_DOCP(messageIsn,"MTN99   ",isn,1)
    'SW_MESSAGES
    'Ստանում է հաղորդագրության ստեղծման ժամը հաղորդագրության Նկարագրություն դաշտից
    messageTime = Left(messageDesc , 12)
    messageTime = Right(messageTime , 4)
    dbSW_MESSAGES(1).fDOCNUM = messageDocN
    dbSW_MESSAGES(1).fISN = messageIsn      
    Call CheckQueryRowCount("SW_MESSAGES","fISN",messageIsn,1)
    dbSW_MESSAGES(1).fAIM = "//" & aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%y%m%d") & messageTime &"+0400               "
    Call CheckDB_SW_MESSAGES(dbSW_MESSAGES(1),1)
    'Փակել Ստացված փոխանցումներ թղթապանակը
    Call Close_Window(wMDIClient, "frmPttel")
    
    'Անցնել Ուղարկվող հաղորդագրություններ/Ուղարկվող խառը հաղորդագրություններ թղթապանակ
    Call GoTo_Sending_Messages(toSend, "|S.W.I.F.T. ²Þî                  |àõÕ³ñÏíáÕ Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñ|àõÕ³ñÏíáÕ Ë³éÁ Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñ")
    'Ջնջել Պատասխան հաղորդագրութունը
    Call SearchAndDelete ( "frmPttel", 2, messageDocN, "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ" )
    'Փակել Ուղարկվող խառը հաղորդագրություններ թղթապանակը
    Call Close_Window(wMDIClient, "frmPttel")
    
    'SQL
    Log.Message "SQL Ստուգումներ Պատասխանի հաղորդագրությունը ջնջելուց հետո հետո",,,SqlDivideColor
    'DOCLOG
    Call CheckQueryRowCount("DOCLOG","fISN",messageIsn,3)
    Call CheckDB_DOCLOG(messageIsn,"77","C","9"," ",1)
    Call CheckDB_DOCLOG(messageIsn,"77","N","1"," ",1)
    Call CheckDB_DOCLOG(messageIsn,"77","D","999"," ",1)
    
    Call CheckQueryRowCount("DOCLOG","fISN",isn,4)
    Call CheckDB_DOCLOG(isn,"77","N","33"," ",1)
    Call CheckDB_DOCLOG(isn,"77","M","10","Received",1)
    Call CheckDB_DOCLOG(isn,"77","M","10","DELETED",2)
    
    'DOCS                                       
    Call CheckQueryRowCount("DOCS","fISN",messageIsn,1)
    Call CheckDB_DOCS(messageIsn,"MTN99   ","999",fBODY,1)
    'DOCP
    Call CheckQueryRowCount("DOCP","fISN",messageIsn,0)
    
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
    'Ջնջում է փաստաթղթի հետ կապակցված Տարանցիկ վճարային փոխանցման փաստաթուղթը
    Call SearchAndDelete( "frmPttel", 1, "TransPay" , "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ" ) 
    Call SearchAndDelete( "frmPttel", 2, isn , "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ" ) 
    Call Close_Window(wMDIClient, "frmPttel")
    
    'SQL
    Log.Message "SQL Ստուգումներ Ներմուծված հաղորդագրությունը ջնջելուց հետո հետո",,,SqlDivideColor
    'DOCLOG    
    Call CheckQueryRowCount("DOCLOG","fISN",isn,6)
    Call CheckDB_DOCLOG(isn,"77","N","33"," ",1)
    Call CheckDB_DOCLOG(isn,"77","M","10","Received",1)
    Call CheckDB_DOCLOG(isn,"77","M","10","DELETED",3)
    Call CheckDB_DOCLOG(isn,"77","D","999"," ",1)
    
    'DOCS                                       
    Call CheckQueryRowCount("DOCS","fISN",isn,1)
    
    Call Close_AsBank()
    
End Sub

Sub Test_Initialize_SWIFT_202_RC()
    sDATE = "20020101"
    fDATE = "20260101"
    savePath =  Project.Path & "Stores\SWIFT\HT202\"
    pathAct = savePath & "Actual.txt"  
    pathExp = savePath & "Expected.txt"
    
    Set confirmation = New_ConfirmationAgreement()
    With confirmation
        .MsgType = "910"
        .Reference = "VTBR2109079860"
        .Account = "777000000003"
        .Date = "07/09/21"
        .Curr = "001"
        .Amount = "300.00"
        .TypeOfOrderingInstitution = "A"
        .OrderingInstitution = "POALILITXXX"
        .TypeOfIntermediaryInstitution = "A"
        .IntermediaryInstitution = "13N}{3:{19:"
        .DateSend = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%d/%m/%y")
    End With
    
    Set recieved = New_Recieved()
    With recieved
        .sDate = aqDateTime.Today
        .eDate = aqDateTime.Today
    End With
    Set toSend = New_Sending()
    With toSend
        .sDate = aqDateTime.Today
        .eDate = aqDateTime.Today
    End With
End Sub

Sub Initialize_DB_SW202_RC(fIsn, docN)
    Set dbSW_MESSAGES(0) = New_SW_MESSAGES()
    With dbSW_MESSAGES(0)
       .fUNIQUEID = "ISN" & fIsn
       .fDATE = "20210907"
       .fMT = "910"
       .fCATEGORY = "2"
       .fSR = "0"
       .fSRBANK = ""
       .fSYS = "1"
       .fSTATE = "  "
       .fUSER = "77"
       .fACCDB = ""
       .fACCCR = ""
       .fAMOUNT = "300.00"
       .fCURR = "001"
       .fPAYER = ""
       .fRECEIVER = ""
       .fAIM = ""
    End With
    
    
    Set dbSW_MESSAGES(1) = New_SW_MESSAGES()
    With dbSW_MESSAGES(1)
       .fUNIQUEID = "ISN" & fIsn
       .fDATE = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%Y%m%d")
       .fMT = "299"
       .fCATEGORY = "2"
       .fSR = "0"
       .fSYS = "1"
       .fUSER = "77"
       .fAMOUNT = "00.00"
    End With
End Sub    
    
    
    
    