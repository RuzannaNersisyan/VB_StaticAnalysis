'USEUNIT SWIFT_International_Payorder_Library
'USEUNIT Library_Common
'USEUNIT Library_Colour
'USEUNIT Library_CheckDB
'USEUNIT Constants
'USEUNIT Payment_Except_Library
'USEUNIT Library_Contracts
'USEUNIT SWIFT_202_InterBank_Transfer
'USEUNIT Main_Accountant_Filter_Library
Option Explicit 

'Test case ID 184735
Dim sDATE, fDATE, settingsPath, max, min, rand, fileFrom, fileTo, what, fWith, savePath, regex, pathAct, pathExp, docN, isn, confDocN, confIsn
Dim messageN,folderDirect, proprietMessage, interBankTransfer, stDate, enDate, wUser, docType, fBODY, dbSW_MESSAGES(2), toSend, recieved

Sub SWIFT_300_Inernational_RightClick_Test()
    Call Test_Initialize_SWIFT_300_RC("","")
    
    'Համակարգ մուտք գործել ARMSOFT օգտագործողով
    Call Initialize_AsBank("bank", sDATE, fDATE)
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
    
    'Վերարտագրում է նախօրոք տրված ֆայլը մեկ այլ ֆայլի մեջ` կատարելով փոփոխություն
    Call Read_Write_File_With_replace(fileFrom,fileTo,what,fWith)
    
    Call SetParameter_InPttel("SWOUT" ,Project.Path & "Stores\SWIFT\HT300\ImportFile\Import\" )
    Call Close_Window(wMDIClient, "frmPttel" )
'-----------------------------------------------------------------------------
'----------------- Կատարել Ընդունել SWIFT համակարգից գործողությունը ------------------
'-----------------------------------------------------------------------------
    Log.Message "Կատարել Ընդունել SWIFT համակարգից գործողությունը",,,DivideColor
    
    'Մուտք գործել "S.W.I.F.T. ԱՇՏ"
    Call ChangeWorkspace(c_SWIFT)
    Call Recieve_From_SWIFT(1)
'--------------------------------------------------------------------------------        
'---------Կատարել Դիտել փաստաթուղթը գործողությունը ներմուծված փաստաթղթի համար--------------
'--------------------------------------------------------------------------------
    Log.Message "Դիտել փաստաթուղթը",,,DivideColor       
    
    'Մուտք գործել Փոխանցումներ/Ստացված փաստաթղթեր թղթապանակ
    Call GoTo_Recieved_Messages(recieved, "|S.W.I.F.T. ²Þî                  |öáË³ÝóáõÙÝ»ñ|êï³óí³Í ÷áË³ÝóáõÙÝ»ñ")
    docN = fWith
    If SearchInPttel("frmPttel", 3, docN) Then
        isn = GetIsn()
        Call wMainForm.MainMenu.Click(c_AllActions)
        Call wMainForm.PopupMenu.Click(c_ViewDoc)    
        BuiltIn.Delay(2000)
        If wMDIClient.WaitVBObject("FrmSpr",2000).exists Then
            Call SaveDoc(savePath, "Actual")
            regex="([[].{10}])|(\d{2}[/]\d{2}[/]\d{2})|(\d{2}:\d{2})"
            Call Compare_Files(pathAct, pathExp, regex)
            'Փակել Փաստաթղթի տպելու ձևը
            Call Close_Window(wMDIClient, "FrmSpr" )
        Else
            Log.Error "Can't find document print view",,,ErrorColor
        End If    
    Else 
        Log.Error "Document row not found",,, ErrorColor
    End If
    Log.Message "fISN = "& isn,,,SqlDivideColor

'-------------------------------------------------------------------
'--------------- Կատարել "Հաստատում" գործողությունը -----------------------
'-------------------------------------------------------------------
    Log.Message "Հաստատում",,,DivideColor 
          
    If SearchInPttel("frmPttel", 3, docN) Then
        Call wMainForm.MainMenu.Click(c_AllActions) 
        Call wMainForm.PopupMenu.Click(c_Confirm)
        If wMDIClient.waitVBObject("frmASDocForm", 2000).exists Then
           Call Rekvizit_Fill("Document",1,"General", "OPERTYPE","EXOP")
           confDocN = Get_Rekvizit_Value("Document",1,"General","BMDOCNUM")
           confIsn = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
           Call ClickCmdButton(1, "Î³ï³ñ»É")
        End If
    Else 
        Log.Error "Document row not found",,, ErrorColor
    End If    
    Log.Message "Confirmed fISN = "& confIsn,,,SqlDivideColor
        
    'Փակել Ստացված փաստաթղթեր թղթապանակը
    Call Close_Window(wMDIClient, "frmPttel")
    
'-------------------------------------------------------------------
'--------- Կատարել "Հաճախորդի թղթապանակ" գործողությունը ------------------
'-------------------------------------------------------------------
    Log.Message "Կատարել Հաճախորդի թղթապանակ գործողությունը",,,DivideColor
    
    'Մուտք Ուղարկվող Հաղորդագրություններ/Ուղարկվող փոխանցումներ թղթապանակ
    Call GoTo_Sending_Transfer(toSend, "|S.W.I.F.T. ²Þî                  |àõÕ³ñÏíáÕ Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñ|àõÕ³ñÏíáÕ ÷áË³ÝóáõÙÝ»ñ")
    
    Log.Message "Հաճախորդի թղթապանակ",,,DivideColor       
    If SearchInPttel("frmPttel",2, confDocN) Then
        Call wMainForm.MainMenu.Click(c_AllActions)
        Call wMainForm.PopupMenu.Click(c_ClFolder)
        Call MessageExists(2,"Ð³×³Ëáñ¹Ç ÃÕÃ³å³Ý³ÏÁ Ñ³ë³Ý»ÉÇ ã¿")
        Call ClickCmdButton(5, "OK")
    Else 
        Log.Error "Document not found"
    End If
    
'-------------------------------------------------------------------
'---- Կատարել "Ստեղծել անհատական հաղորդագրություն(SWIFT)" գործողությունը -----
'-------------------------------------------------------------------
    Log.Message "Ստեղծել անհատական հաղորդագրություն(SWIFT)",,,DivideColor
           
    If SearchInPttel("frmPttel",2, confDocN) Then
        Call wMainForm.MainMenu.Click(c_AllActions)
        Call wMainForm.PopupMenu.Click(c_CrSWProp)
        If wMDIClient.WaitVBObject("frmASDocForm",4000).exists Then
            Call Test_Initialize_SWIFT_300_RC(docN, confDocN)
            proprietMessage.docN = Get_Rekvizit_Value("Document",1,"General","BMDOCNUM")
            proprietMessage.isn = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
            Call ClickCmdButton(1, "Î³ï³ñ»É")
            
            'SQL
            Call SQL_Initialize_DB_SW300_RC(proprietMessage.isn, proprietMessage.docN)
        
        Log.Message "'SQL Ստուգումներ Ստեղծել անհատական հաղորդագրություն(SWIFT) գործողությունը կատարելուց հետո",,,SqlDivideColor
        Log.Message "Message fISN = "& proprietMessage.isn,,,SqlDivideColor
            
            'DOCLOG
            Call CheckQueryRowCount("DOCLOG","fISN",proprietMessage.isn,2)
            Call CheckDB_DOCLOG(proprietMessage.isn,"77","C","9"," ",1)
            Call CheckDB_DOCLOG(proprietMessage.isn,"77","N","1"," ",1)
            'DOCS                                       
            fBODY = "  CATEGORY:3  BMDOCNUM:" & proprietMessage.docN & "  TYPE:300  AIM:²ÝÑ³ï³Ï³Ý Ñ³Õáñ¹³·ñáõÃÛáõÝ  VERIFIED:0  USERID:  77  "_
                    &"BMIODATE:"& aqConvert.DateTimeToFormatStr(aqDateTime.Today,"20%y%m%d") &"  RSBKMAIL:0  DELIV:0  "
            fBODY = Replace(fBODY, "  ", "%")
            Call CheckQueryRowCount("DOCS","fISN",proprietMessage.isn,1)
            Call CheckDB_DOCS(proprietMessage.isn,"MTN98   ","9",fBODY,1)
            'DOCP
            Call CheckQueryRowCount("DOCP","fISN",proprietMessage.isn,1)
            Call CheckDB_DOCP(proprietMessage.isn,"MTN98   ",confIsn,1)
            'SW_MESSAGES
            dbSW_MESSAGES(0).fDOCNUM = proprietMessage.docN
            dbSW_MESSAGES(0).fISN = proprietMessage.isn      
            Call CheckQueryRowCount("SW_MESSAGES","fISN",proprietMessage.isn,1)
            Call CheckDB_SW_MESSAGES(dbSW_MESSAGES(0),1)
            
            'Փակել Ուղարկվող փոխանցումներ թղթապանակը
            Call Close_Window (wMDIClient, "frmPttel")
            
            'Անցնել Ուղարկվող հաղորդագրություններ/Ուղարկվող խառը հաղորդագրություններ թղթապանակ   
            Call GoTo_Sending_Messages(toSend, "|S.W.I.F.T. ²Þî                  |àõÕ³ñÏíáÕ Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñ|àõÕ³ñÏíáÕ Ë³éÁ Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñ")
            If WaitForPttel("frmPttel") Then
                'Կատարել Դիտել գործողությունը Անհատական հաղորդագրության համար
                If SearchInPttel("frmPttel",2, proprietMessage.docN) Then
                    Call wMainForm.MainMenu.Click(c_AllActions)
                    Call wMainForm.PopupMenu.Click(c_View)
                    If wMDIClient.WaitVBObject("frmASDocForm",4000).exists Then
                        Call Personal_Message_Window_Check (proprietMessage)
                        Call ClickCmdButton(1, "OK")
                        'Ջնջել Անհատական հաղորդագրությունը
                        Call SearchAndDelete("frmPttel", 2,proprietMessage.docN,"Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ")
                    End If
                Else 
                    Log.Error "Document row not found",,, ErrorColor
                End If       
                'Փակել Ուղարկվող խառը հաղորդագրություններ թղթապանակը
                Call Close_Window(wMDIClient, "frmPttel")
            Else 
                Log.Error "Pttel not found",,,ErrorColor 
            End If              
        Else
            Log.Error "Document window not found",,,ErrorColor  
        End If      
    Else 
        Log.Error "Document not found",,,ErrorColor
    End If      
    
    'SQL
     Log.Message "'SQL Ստուգումներ անհատական հաղորդագրությունը ջնջելուց հետո",,,SqlDivideColor
    'DOCLOG
    Call CheckQueryRowCount("DOCLOG","fISN",proprietMessage.isn,3)
    Call CheckDB_DOCLOG(proprietMessage.isn,"77","C","9"," ",1)
    Call CheckDB_DOCLOG(proprietMessage.isn,"77","N","1"," ",1)
    Call CheckDB_DOCLOG(proprietMessage.isn,"77","D","999"," ",1)
    
    'DOCS
    Call CheckDB_DOCS(proprietMessage.isn,"MTN98   ","999",fBODY,1)
'-------------------------------------------------------------------
'-------------- Կատարել "Ստեղծել MT202" գործողությունը -------------------
'-------------------------------------------------------------------    
    Log.Message "Ստեղծել MT202",,,DivideColor       
    
    'Մուտք Ուղարկվող Հաղորդագրություններ/Ուղարկվող փոխանցումներ թղթապանակ
    Call GoTo_Sending_Transfer(toSend, "|S.W.I.F.T. ²Þî                  |àõÕ³ñÏíáÕ Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñ|àõÕ³ñÏíáÕ ÷áË³ÝóáõÙÝ»ñ")      
    
    If SearchInPttel("frmPttel",2, confDocN) Then
        Call wMainForm.MainMenu.Click(c_AllActions)
        Call wMainForm.PopupMenu.Click(c_SWCreate202)
        If wMDIClient.WaitVBObject("frmASDocForm",4000).exists Then
            interBankTransfer.common.docN = confDocN
            interBankTransfer.common.reference = confDocN
            Call Rekvizit_Fill("Document",2,"General","ADDINFO","^A[Del]") 
            Call ClickCmdButton(1, "Î³ï³ñ»É")
            Call MessageExists(2, "CITIUS33XXX µ³ÝÏÇ Ñ³Ù³ñ Áëï 001 ³ñÅáõÛÃÇ Ý³ËÁÝïñ»ÉÇ ÃÕÃ³Ïó³ÛÇÝ Ñ³ßÇí " & vbNewLine & "Ýß³Ý³Ïí³Í ã¿")
            Call ClickCmdButton(5, "²Ûá")
        Else
            Log.Error "Document window not found",,,ErrorColor  
        End If
        'Կատարել Դիտել գործողությունը MT202 հաղորդագրության համար
        If SearchInPttel("frmPttel",1, "202") Then
            Call wMainForm.MainMenu.Click(c_AllActions)
            Call wMainForm.PopupMenu.Click(c_View)
            Call MessageExists(2, "Ð³Õáñ¹³·ñáõÃÛáõÝÁ áõÝÇ Ï³å³Ïóí³Í ÷³ëï³ÃáõÕÃ")
            Call ClickCmdButton(5, "OK")
            Call InterBank_Transfer_Check(interBankTransfer)
            interBankTransfer.common.isn = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
            Call ClickCmdButton(1, "OK")
        Else
            Log.Error "Document not found",,,ErrorColor
        End If             
    Else 
        Log.Error "Document not found"
    End If
    
    Log.Message "SQL ստուգումներ MT202 Հաղորդագրությունը ստեղծելուց հետո ",,,SqlDivideColor
    Log.Message "MT202 fISN = "& interBankTransfer.common.isn,,,SqlDivideColor
    'SQL
    Call SQL_Initialize_DB_SW300_RC(interBankTransfer.common.isn, interBankTransfer.common.docN)
    'DOCLOG
    Call CheckQueryRowCount("DOCLOG","fISN",confIsn,4)
    Call CheckDB_DOCLOG(confIsn,"77","C","9"," ",1)
    Call CheckDB_DOCLOG(confIsn,"77","N","1"," ",1)
    Call CheckDB_DOCLOG(confIsn,"77","M","9","DELETED",1)
    Call CheckDB_DOCLOG(confIsn,"77","M","9","CREATED",1)
    
    Call CheckQueryRowCount("DOCLOG","fISN",interBankTransfer.common.isn,1)
    Call CheckDB_DOCLOG(confIsn,"77","N","1"," ",1)
    
    'DOCS
    fBODY = "  MT:202  BMDOCNUM:"& interBankTransfer.common.docN &"  REF:" & interBankTransfer.common.reference & "  DATE:20211108  RINSTOP:A  "_
            &"RINSTID:45646asd6564  RECINST:AMASJPJZXXX  RECOP:A  RECEIVER:XXN}{4:{4:  SUMMA:100  CUR:001  VERIFIED:0  "_
            &"BMIODATE:"& aqConvert.DateTimeToFormatStr(aqDateTime.Today,"20%y%m%d") &"  RSBKMAIL:0  DELIV:0  "_
            &"USERID:  77  SNDREC:CITIUS33XXX  ISCOVER:0  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckDB_DOCS(interBankTransfer.common.isn,"MT202   ","9",fBODY,1)
    Call CheckQueryRowCount("DOCS","fISN",interBankTransfer.common.isn,1)
        
    'DOCP
    Call CheckQueryRowCount("DOCP","fISN",interBankTransfer.common.isn,1)
    Call CheckDB_DOCP(interBankTransfer.common.isn,"MT202   ",confIsn,1)
    'SW_MESSAGES
    dbSW_MESSAGES(1).fDOCNUM = interBankTransfer.common.docN
    dbSW_MESSAGES(1).fISN = interBankTransfer.common.isn      
    Call CheckQueryRowCount("SW_MESSAGES","fISN",interBankTransfer.common.isn,1)
    Call CheckDB_SW_MESSAGES(dbSW_MESSAGES(1),1)
    
    'Ջնջել MT202 հաղորդագրությունը
    Call SearchAndDelete("frmPttel",1,"202","Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ") 
    
    'SQL
    Log.Message "SQL ստուգումներ MT202 Հաղորդագրությունը ջնջելուց հետո ",,,SqlDivideColor
    'DOCLOG
    Call CheckQueryRowCount("DOCLOG","fISN",interBankTransfer.common.isn,2)
    Call CheckDB_DOCLOG(interBankTransfer.common.isn,"77","N","9"," ",1)
    Call CheckDB_DOCLOG(interBankTransfer.common.isn,"77","D","999"," ",1)
    
    'DOCLOG
    Call CheckQueryRowCount("DOCLOG","fISN",confIsn,5)
    Call CheckDB_DOCLOG(confIsn,"77","C","9"," ",1)
    Call CheckDB_DOCLOG(confIsn,"77","N","1"," ",1)
    Call CheckDB_DOCLOG(confIsn,"77","M","9","DELETED",2)
    Call CheckDB_DOCLOG(confIsn,"77","M","9","CREATED",1)
    
    'DOCS
    Call CheckDB_DOCS(interBankTransfer.common.isn,"MT202   ","999",fBODY,1)
    
'-------------------------------------------------------------------
'----------------- Կատարել "Խմբագրել" գործողությունը ----------------------
'-------------------------------------------------------------------    
    Log.Message "Խմբագրել",,,DivideColor     
    
    If SearchInPttel("frmPttel", 2, confDocN) Then
        Call wMainForm.MainMenu.Click(c_AllActions)
        Call wMainForm.PopupMenu.Click(c_ToEdit)
        If wMDIClient.WaitVBObject("frmASDocForm", 4000).exists Then
            Call GoTo_ChoosedTab(2)
            Call Rekvizit_Fill ("Document", 2, "General", "SNDREC", "####RU##")   
            Call ClickCmdButton(1, "Î³ï³ñ»É")
        Else
            Log.Error "Document window not found",,,ErrorColor  
        End If
    Else 
        Log.Error "Document not found",,,ErrorColor        
    End If
    
    'SQL
    Log.Message "'SQL Ստուգումներ հաստատված MT300 հաղորդագրությունը խմբագրելուց հետո հետո",,,SqlDivideColor  
    
    Call SQL_Initialize_DB_SW300_RC(confIsn, interBankTransfer.common.docN)
    'DOCLOG
    Call CheckQueryRowCount("DOCLOG","fISN",confIsn,6)
    Call CheckDB_DOCLOG(confIsn,"77","C","9"," ",1)
    Call CheckDB_DOCLOG(confIsn,"77","N","1"," ",1)
    Call CheckDB_DOCLOG(confIsn,"77","M","9","DELETED",2)
    Call CheckDB_DOCLOG(confIsn,"77","M","9","CREATED",1)   
    Call CheckDB_DOCLOG(confIsn,"77","E","9"," ",1)    
    'DOCS    
    Call CheckQueryRowCount("DOCS","fISN",confIsn,1)
    fBODY = "  BMDOCNUM:" & confDocN & "  REF:" & docN & "  OPERTYPE:EXOP  COMREF:BEST223858XXN}:{  OPERSCOPE:AGNT  BLOCKIND:N  "_
          & "SPLITIND:Y  PARTYAOP:A  PARTYAID:125645656asd  PARTYA:CITIUS33XXX  PARTYBOP:A  PARTYBID:123546988  PARTYB:CITIUS33XXX  FUNDOP:J  "_          
          & "FUND:CITIUS33XXX  VERIFIED:0  DATE:20211108  PAYDATE:20211108  CURB:003  CURS:001  SUMMAB:72.16  SUMMAS:100  COURSE:1.3858/1  "_
          &"RATE:1,3858  SNDREC:####RU##  DAGENTBOP:J  DAGENTB:CITIUS33XXX  MEDBOP:A  MEDBID:645646  MEDBBANK:DGPBDE3MBRA  RAGENTBOP:J  "_
          &"RAGENTB:CITIUS33XXX  DAGENTSOP:A  DAGENTSID:464asd64s  DAGENTS:CITIUS33XXX  MEDSOP:J  MEDSBANK:CITIUS33XXX  RAGENTSOP:A  "_
          &"RAGENTSID:45646asd6564  RAGENTS:AMASJPJZXXX  CNTREF:1125s52248s5  BRKREF:21asd84646  TERMS:sadf12541a  ADDINFO:12556s54das  "_
          &"BMIODATE:" & aqConvert.DateTimeToFormatStr(aqDateTime.Now(),"%Y%m%d") & "  RSBKMAIL:0  DELIV:0  USERID:  77  PROCESSED:0  "_
          &"CONTACT:55d445s8  DEALMETH:BROK  DEALADD:3218655s85  DEALAOP:A  DEALAID:s99d56a9662  DEALA:CITIUS33XXX  DEALBOP:A  "_
          &"DEALBID:125598556asd  DEALB:CITIUS33XXX  BROKEROP:J  BROKER:BRAJINBBBYF  COUNT:2  "

    fBODY = Replace(fBODY, "  ", "%")
    Call CheckDB_DOCS(confIsn,"MT300   ","9",fBODY,1)    
    'SW_MESSAGES
    dbSW_MESSAGES(2).fDOCNUM = confDocN
    dbSW_MESSAGES(2).fISN = confIsn      
    Call CheckQueryRowCount("SW_MESSAGES","fISN",confIsn,1)
    Call CheckDB_SW_MESSAGES(dbSW_MESSAGES(2),1)
    
    'Ջնջել հաստատված փաստաթուղթը
    Call SearchAndDelete("frmPttel", 2, confDocN, "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ")
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
    Log.Message "Ջնջում է ներմուծված փաստաթուղթը",,,DivideColor

    Call SearchAndDelete( "frmPttel", 2, isn , "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ" ) 
    'Փակել Ստեղծված փաստաթղթեր թղթապանակը
    Call Close_Window(wMDIClient, "frmPttel")
    
    Call Close_AsBank()
End Sub
 
Sub Test_Initialize_SWIFT_300_RC(docN, confDocN)
    sDATE = "20050101"
    fDATE = "20260101"
    
    savePath =  Project.Path & "Stores\SWIFT\HT300\"
    pathAct = savePath & "Actual.txt"  
    pathExp = savePath & "Expected.txt"
    
    Set recieved = New_Recieved()
    With recieved
        .sDate = aqDateTime.Today
        .eDate = aqDateTime.Today
    End With
    
    Set toSend = New_Sending()
    With toSend
        .sDate = "010120"
        .eDate = aqDateTime.Today
    End With
    
    Set proprietMessage = New_Proprietary_Message
    With proprietMessage
        .category = "3"
        .messageType = "300"
        .descript = "²ÝÑ³ï³Ï³Ý Ñ³Õáñ¹³·ñáõÃÛáõÝ"
        '2
        .sendRecDate = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%d/%m/%y") 

    End With
    
    Set interBankTransfer = New_InterBank_Transfer
    With interBankTransfer
        .common.msgType = "202"
        .common.date = "08/11/21"
        .common.accWithInstType = "A"
        .common.accWithInstPID = "45646asd6564"
        .common.accWithInst = "AMASJPJZXXX"
        .common.benClientType = "A"
        .common.benClient = "XXN}{4:{4:"
        .common.sum = "100.00"
        .common.cur = "001"
        '2
        .add.sendRecDate = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%d/%m/%y")
        '3
        .finOrg.sendRec = "CITIUS33XXX"  
        '4
        .preTransfer.check = False      
    End With
    
    proprietMessage.subMessage = ":15A:" & vbCrLf _ 
                                        & ":20:"& confDocN & vbNewLine _
                                        &":21:"& docN & vbNewLine _
                                        &":22A:EXOP" & vbNewLine _
                                        &":94A:AGNT" & vbNewLine _
                                        &":22C:BEST223858XXN}:{" & vbNewLine _
                                        &":17T:N" & vbNewLine _
                                        &":17U:Y" & vbNewLine _
                                        &":82A:/125645656asd" & vbNewLine _
                                        &"CITIUS33XXX" & vbNewLine _
                                        &":87A:/123546988" & vbNewLine _
                                        &"CITIUS33XXX" & vbNewLine _
                                        &":83J:CITIUS33XXX" & vbNewLine _
                                        &":77D:sadf12541a" & vbNewLine _
                                        &":15B:" & vbNewLine _
                                        &":30T:20211108" & vbNewLine _
                                        &":30V:20211108" & vbNewLine _
                                        &":36:1,3858" & vbNewLine _
                                        &":32B:EUR72,16" & vbNewLine _
                                        &":53J:CITIUS33XXX" & vbNewLine _
                                        &":56A:/645646" & vbNewLine _
                                        &"DGPBDE3MBRA" & vbNewLine _
                                        &":57J:CITIUS33XXX" & vbNewLine _
                                        &":33B:USD100," & vbNewLine _
                                        &":53A:/464asd64s" & vbNewLine _
                                        &"CITIUS33XXX" & vbNewLine _ 
                                        &":56J:CITIUS33XXX" & vbNewLine _
                                        &":57A:/45646asd6564" & vbNewLine _
                                        &"AMASJPJZXXX" & vbNewLine _
                                        &":15C:" & vbNewLine _
                                        &":29A:55d445s8" & vbNewLine _
                                        &":24D:BROK/3218655s85" & vbNewLine _
                                        &":84A:/s99d56a9662" & vbNewLine _
                                        &"CITIUS33XXX" & vbNewLine _
                                        &":85A:/125598556asd" & vbNewLine _
                                        &"CITIUS33XXX" & vbNewLine _
                                        &":88J:BRAJINBBBYF" & vbNewLine _
                                        &":26H:1125s52248s5" & vbNewLine _
                                        &":21G:21asd84646" & vbNewLine _
                                        &":72:12556s54das" & vbNewLine _
                                        &":15D:" & vbNewLine _
                                        &":17A:Y" & vbNewLine _
                                        &":32B:USD10," & vbNewLine _
                                        &":53A:/548sdf48sd4f6" & vbNewLine _
                                        &"CITIUS33XXX" & vbNewLine _
                                        &":56A:/12564564asdf" & vbNewLine _
                                        &"BROMITRDXXX" & vbNewLine _
                                        &":57A:/15315sd" & vbNewLine _
                                        &"/ABIC/CITIUS33XXX" & vbNewLine _
                                        &":17A:N" & vbNewLine _
                                        &":32B:EUR100," & vbNewLine _
                                        &":53J:CITIUS33XXX" & vbNewLine _
                                        &":56J:CITIUS33XXX" & vbNewLine _
                                        &":57J:/ABIC/CITIUS33XXX" & vbNewLine _
                                        &":16A:2"
End Sub

Sub SQL_Initialize_DB_SW300_RC(fIsn, docN)
    Set dbSW_MESSAGES(0) = New_SW_MESSAGES()
    With dbSW_MESSAGES(0)
       .fUNIQUEID = "ISN" & fIsn
       .fDATE = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"20%y%m%d")
       .fMT = "398"
       .fCATEGORY = "2"
       .fSR = "0"
       .fSRBANK = ""
       .fSYS = "1"
       .fSTATE = "  "
       .fUSER = "77"
       .fACCDB = ""
       .fACCCR = ""
       .fAMOUNT = "0.00"
       .fCURR = ""
       .fPAYER = ""
       .fRECEIVER = ""
       .fAIM = "²ÝÑ³ï³Ï³Ý Ñ³Õáñ¹³·ñáõÃÛáõÝ      "
       .fBRANCH = Null
       .fDEPART = Null
    End With
    
    Set dbSW_MESSAGES(1) = New_SW_MESSAGES()
    With dbSW_MESSAGES(1)
       .fUNIQUEID = "ISN" & fIsn
       .fDATE = "20211108"
       .fMT = "202"
       .fCATEGORY = "1"
       .fSR = "0"
       .fSRBANK = "CITIUS33XXX"
       .fSYS = "1"
       .fSTATE = "  "
       .fUSER = "77"
       .fACCDB = ""
       .fACCCR = ""
       .fAMOUNT = "100.00"
       .fCURR = "001"
       .fPAYER = ""
       .fRECEIVER = "XXN}{4:{4:                      "
       .fAIM = ""
       .fBRANCH = ""
       .fDEPART = ""
    End With 
    
    Set dbSW_MESSAGES(2) = New_SW_MESSAGES()
    With dbSW_MESSAGES(2)
       .fUNIQUEID = "ISN" & fIsn
       .fDATE = "20211108"
       .fMT = "300"
       .fCATEGORY = "1"
       .fSR = "0"
       .fSRBANK = "####RU##   "
       .fSYS = "1"
       .fSTATE = "  "
       .fUSER = "77"
       .fACCDB = ""
       .fACCCR = ""
       .fAMOUNT = "72.16"
       .fCURR = "003"
       .fPAYER = ""
       .fRECEIVER = ""
       .fAIM = "12556s54das                     "
       .fBRANCH = ""
       .fDEPART = ""
    End With    

End Sub
