'USEUNIT Library_Common
'USEUNIT Constants
'USEUNIT Mem_Order_Library
'USEUNIT Library_Colour
'USEUNIT Library_CheckDB
'USEUNIT Library_Contracts
'USEUNIT Main_Accountant_Filter_Library

'Test Case ID 182231
Option Explicit

Dim grCredOrd (1), sDATE, eDATE, cAccIsn(2), dAccIsn, cAccEditIsn(1), dAccEditIsn, pathExp, i, workingDocs, docGrid
Dim folderDirect, stDate, enDate, wUser, docType , colReadOnlyArray, fBODY, sumSQL, sumCur(1),  dbFOLDERS(3), expMessage
    
Sub Group_Credit_Order_With_Curr_Change_Test () 
    aCount = 3
    
    Call Test_Initialize_Group_Cred_Order_Curr
    
    'Մուտք ծրագիր ARMSOFT Օգտագործողով
    Call Initialize_AsBank("bank", sDATE, eDATE)
    Call Login ("ARMSOFT")
    
    'Հաշիվների ISN-ների ստացում SQL աղյուսակներից
    For i = 0 to grCredOrd(0).commonTab.cAccsCount - 1
        cAccIsn(i) = GetAccountISN(grCredOrd(0).commonTab.accC(i))
    Next    
    dAccIsn = GetAccountISN(grCredOrd(0).commonTab.accD)
    For i = 0 to grCredOrd(1).commonTab.cAccsCount - 1
        cAccEditIsn(i) = GetAccountISN(grCredOrd(1).commonTab.accC(i))
    Next  
    dAccEditIsn = GetAccountISN(grCredOrd(1).commonTab.accD)
    
    'Մուտք Գլխավոր հաշվապահի ԱՇՏ
    Log.Message  "Մուտք Գլխավոր հաշվապահի ԱՇՏ",,,DivideColor
    Call ChangeWorkspace(c_ChiefAcc)


'----------------------------------------------------------------------------------   
'---------------------Ստեղծել Խմբային Կրեդիտի Օրդեր--------------------------------------
'------Նոր փաստաթղթեր/Վճարային փաստաթղթեր/Ներքին գործարքներ/Խմբային կրեդիտի օրդեր ճանապարհից--- 

    Call wTreeView.DblClickItem("|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|Üáñ ÷³ëï³ÃÕÃ»ñ|ì×³ñ³ÛÇÝ ÷³ëï³ÃÕÃ»ñ|Ü»ñùÇÝ ·áñÍ³ñùÝ»ñ|ÊÙµ³ÛÇÝ Ïñ»¹ÇïÇ ûñ¹»ñ")
    Call Fill_Group_Cred_Order_Common (grCredOrd(0).commonTab)
    Call ClickCmdButton(1, "Î³ï³ñ»É")    
    
    'SQL Ստուգումներ
    Log.Message "'SQL Ստուգումներ Խմբային կրեդիտի օրդեր ստեղծելուց հետո",,,SqlDivideColor
    Call Intitialize_DB_Group_Cred_Order (grCredOrd(0).commonTab.isn , grCredOrd(0).commonTab.docN)
    'DOCLOG
    Call CheckQueryRowCount("DOCLOG","fISN",grCredOrd(0).commonTab.isn, 1)
    Call CheckDB_DOCLOG(grCredOrd(0).commonTab.isn,"77","N","1"," ",1)
    
    'DOCS   
    fBODY = "  TYPECODE:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28 10  USERID:  77  ACSBRANCH:00  ACSDEPART:1  "_
               &"DOCNUM:" & grCredOrd(0).commonTab.docN & "  DATE:20220228  ACCDB:77781633311  CUR:001  SUMMA:9999999999  AIM:1  "   
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS", "fISN", grCredOrd(0).commonTab.isn, 1)
    Call CheckDB_DOCS(grCredOrd(0).commonTab.isn, "CrPkOrd ", "1", fBODY,1)

    'DOCSG
    Call CheckQueryRowCount("DOCSG", "fISN", grCredOrd(0).commonTab.isn, 6)
    For i = 0 to grCredOrd(0).commonTab.cAccsCount - 1
         Call CheckDB_DOCSG(grCredOrd(0).commonTab.isn,"SUBSUMS",grCredOrd(0).commonTab.cAccRowN(i) ,"ACCCR",grCredOrd(0).commonTab.accC(i),1)
    Next
    For i = 0 to grCredOrd(0).commonTab.cAccsCount - 1
         Call CheckDB_DOCSG(grCredOrd(0).commonTab.isn,"SUBSUMS",grCredOrd(0).commonTab.cAccRowN(i) ,"ACCCRNAME",grCredOrd(0).commonTab.cAccName(i),1)
    Next
    For i = 0 to grCredOrd(0).commonTab.cAccsCount - 1
         sumSQL = Replace (grCredOrd(0).commonTab.sum(i),"," , "")
         sumSQL = Replace (sumSQL, ".00", "")
         Call CheckDB_DOCSG(grCredOrd(0).commonTab.isn,"SUBSUMS", grCredOrd(0).commonTab.cAccRowN(i) , "SUMMA", sumSQL, 1)
    Next
    
    'FODLERS
    Call CheckQueryRowCount("FOLDERS","fISN",grCredOrd(0).commonTab.isn,1) 
    Call CheckDB_FOLDERS(dbFOLDERS(0),1)
    
    'HI
    sumCur(0) = "1777777777600.00"
    sumCur(1) = "2222222222000.00"
    Call CheckQueryRowCount("HI", "fBASE", grCredOrd(0).commonTab.isn, 4)
    For i = 0 to grCredOrd(0).commonTab.cAccsCount - 1
         sumSQL = Replace (grCredOrd(0).commonTab.sum(i),"," , "")
         Call Check_HI_CE_accounting ("20220228",grCredOrd(0).commonTab.isn , "11", dAccIsn ,sumCur(i), "001", sumSQL, "MSC", "D")
         Call Check_HI_CE_accounting ("20220228",grCredOrd(0).commonTab.isn , "11", cAccIsn(i) ,sumCur(i), "001", sumSQL, "MSC", "C")
    Next
     
    'HIREST
    sumCur(0) = "1777777777600.00"
    sumCur(1) = "2222222222000.00"
    For i = 0 to grCredOrd(0).commonTab.cAccsCount - 1
        sumSQL = Replace (grCredOrd(0).commonTab.sum(i),"," , "")
        Call CheckDB_HIREST("11", cAccIsn(i) , "-" & sumCur(i) ,"001", "-" & sumSQL, 1)
    Next
    Call CheckDB_HIREST("11", dAccIsn , "3999999999600.00" ,"001", "9999999999.00", 1)
    
    'Ստուգել տպելու ձևը
    Call Group_Cred_Order_DocCheck (pathExp, "Curr_Act")
    
    'Փաստաթղթի համարի և isn-ի լոգավորում
    Log.Message "Document ISN = " & grCredOrd(0).commonTab.isn,,, SqlDivideColor
    Log.Message "Document Number = " & grCredOrd(0).commonTab.DocN,,, DivideColor
    
    Call Close_Window(wMDIClient, "FrmSpr" )


'--------------------------------------------------------------------------------    
'--------------------Մուտք Աշխատանքային փաստաթղթեր թղթապանակ------------------------
'--------------------------------------------------------------------------------
    Log.Message  "Մուտք Աշխատանքային փաստաթղթեր թղթապանակ",,,DivideColor
    Call GoTo_MainAccWorkingDocuments("|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|", workingDocs)
    'Կատարել Դիտել գործողությունը
    grCredOrd(0).commonTab.mDate = "28/02/22"
    Call View_Group_Cred_Order (grCredOrd(0), "frmPttel")
'--------------------------------------------------------------------------------    
'---------Արժույթի փոփոխման հետ կապված Error-ների ստուգում փաստաթուղթը խմբագրելիս---------
'--------------------------------------------------------------------------------
    If SearchInPttel("frmPttel",2, grCredOrd(0).commonTab.docN) Then
        BuiltIn.Delay(2000)
        Call wMainForm.MainMenu.Click (c_Allactions)
        Call wMainForm.PopupMenu.Click (c_ToEdit)
        BuiltIn.Delay(delay_middle)
        Set docGrid = wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame").VBObject("DocGrid")
        
        If  wMDIClient.WaitVBObject("frmASDocForm",3000).exists Then
            'Հաշիվ Դեբետ դաշտի լրացում 000 արժույթ ունեցող հաշվեհամարով
            wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame").VBObject("AsTypeFolder").VBObject("TDBMask").Keys("00005650100")
            Call ClickCmdButton(1,"Î³ï³ñ»É")
            expMessage = "Ð³ßíÇ ³ñï³ñÅáõÃ³ÛÝáõÃÛ³Ý ëË³É " & vbNewLine & "   Ð³ßÇí  -  00068360101 (³ñï³ñÅáõÛÃÇ Ïá¹ª 001) " & vbNewLine _ 
                        &"   Ð³ßíÇ ³ñï³ñÅáõÛÃÁ å»ïù ¿ ÉÇÝÇ  000 "
            Call MessageExists(2,expMessage )
            Call ClickCmdButton(5, "OK" )
            For i = 0 to docGrid.ApproxCount
                With docGrid
                    .row = 0
                    .Keys ("^d")
                End With     
            Next
            'Գումարներ աղյուսակի 003 արժույթ ունեցող Հաշվեհամարով լրացում
            With docGrid
                .Col = 0
                .row = 0
                .Keys("01078573313" & "[Enter]")                 
            End With
            Call MessageExists(2,"01078573313  Ñ³ßíÇ ³ñÅáõÛÃÇ ³ÝÑ³Ù³å³ï³ëË³ÝáõÃÛáõÝ" )
            Call ClickCmdButton(5, "OK" )
            Call ClickCmdButton(1,"¸³¹³ñ»óÝ»É") 
        End If
    End If
'--------------------------------------------------------------------------------------------------------------- 
'------------Խմբագրել Խմբային կրեդիտի օրդերը-------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------- 
    Call Edit_Group_Cred_Order (grCredOrd(0), grCredOrd(1), 0, "frmPttel")
    
    'SQL Ստուգումներ
    Log.Message "'SQL Ստուգումներ Խմբային կրեդիտի օրդերը խմբագրելուց հետո",,,SqlDivideColor
    Call Intitialize_DB_Group_Cred_Order (grCredOrd(1).commonTab.isn , grCredOrd(1).commonTab.docN)
    'DOCLOG
    Call CheckQueryRowCount("DOCLOG","fISN",grCredOrd(1).commonTab.isn, 2)
    Call CheckDB_DOCLOG(grCredOrd(1).commonTab.isn,"77","N","1"," ",1)
    Call CheckDB_DOCLOG(grCredOrd(1).commonTab.isn,"77","E","1"," ",1)
    'DOCS                                                     
    fBODY = "  TYPECODE:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28 10  USERID:  77  ACSBRANCH:00  ACSDEPART:1  "_
               &"DOCNUM:" & grCredOrd(1).commonTab.docN & "  DATE:20220901  ACCDB:77809533317  CUR:003  SUMMA:5000000  "_
               &"AIM:Ð³Ù³Ó³ÛÝ ³ñï. ³éùáõí³×³éùÇ å³ÛÙ³Ý³·ñÇ0000102030405060708091011121314151617181920212223242526272829303132333435363738394041424344454647484950  "   
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS", "fISN", grCredOrd(1).commonTab.isn, 1)
    Call CheckDB_DOCS(grCredOrd(1).commonTab.isn, "CrPkOrd ", "1", fBODY,1)

    'DOCSG
    Call CheckQueryRowCount("DOCSG", "fISN", grCredOrd(1).commonTab.isn, 3)
    For i = 0 to grCredOrd(1).commonTab.cAccsCount - 1
        Call CheckDB_DOCSG(grCredOrd(1).commonTab.isn,"SUBSUMS",grCredOrd(1).commonTab.cAccRowN(i) ,"ACCCR",grCredOrd(1).commonTab.accC(i),1)
    Next
    For i = 0 to grCredOrd(1).commonTab.cAccsCount - 1
        Call CheckDB_DOCSG(grCredOrd(1).commonTab.isn,"SUBSUMS",grCredOrd(1).commonTab.cAccRowN(i) ,"ACCCRNAME",grCredOrd(1).commonTab.cAccName(i),1)
    Next
    For i = 0 to grCredOrd(1).commonTab.cAccsCount - 1
        sumSQL = Replace (grCredOrd(1).commonTab.sum(i),"," , "")
        sumSQL = Replace (sumSQL, ".00", "")
        Call CheckDB_DOCSG(grCredOrd(1).commonTab.isn,"SUBSUMS", grCredOrd(1).commonTab.cAccRowN(i) , "SUMMA", sumSQL, 1)
    Next
    
    'FODLERS
    Call CheckQueryRowCount("FOLDERS","fISN",grCredOrd(1).commonTab.isn,1) 
    Call CheckDB_FOLDERS(dbFOLDERS(1),1)
    
    'HI
    sumCur(0) = "2250000000.00"
    Call CheckQueryRowCount("HI", "fBASE", grCredOrd(1).commonTab.isn, 2)
    For i = 0 to grCredOrd(1).commonTab.cAccsCount - 1
         sumSQL = Replace (grCredOrd(1).commonTab.sum(i),"," , "")
         Call Check_HI_CE_accounting ("20220901",grCredOrd(1).commonTab.isn , "11", dAccEditIsn ,sumCur(i), "003", sumSQL, "MSC", "D")
         Call Check_HI_CE_accounting ("20220901",grCredOrd(1).commonTab.isn , "11", cAccEditIsn(i) ,sumCur(i), "003", sumSQL, "MSC", "C")
    Next
     
    'HIREST
    sumCur(0) = "2250000000.00"
    For i = 0 to grCredOrd(1).commonTab.cAccsCount - 1
        sumSQL = Replace (grCredOrd(1).commonTab.sum(i),"," , "")
        Call CheckDB_HIREST("11", cAccEditIsn(i) , "-" & sumCur(i) ,"003", "-" & sumSQL, 1)
    Next
    Call CheckDB_HIREST("11", dAccEditIsn , "2249935356.30" ,"003", "4999902.00", 1)

    'Կատարել Դիտել գործողությունը
    grCredOrd(1).commonTab.mDate = "01/09/22"
    Call View_Group_Cred_Order (grCredOrd(1), "frmPttel")
'-----------------------------------------------------------------------    
'------------------Հաշվառել փասթտաթուղթը----------------------------------
'-----------------------------------------------------------------------    
    If SearchInPttel("frmPttel", 2, grCredOrd(1).commonTab.docN) Then
       Call Register_Payment()
    End If
    
    'SQL Ստուգումներ
    Log.Message "'SQL Ստուգումներ Խմբային կրեդիտի օրդերը հաշվառելուց հետո",,,SqlDivideColor
    Call Intitialize_DB_Group_Cred_Order (grCredOrd(1).commonTab.isn , grCredOrd(1).commonTab.docN)
    'DOCLOG
    Call CheckQueryRowCount("DOCLOG","fISN",grCredOrd(1).commonTab.isn, 3)
    Call CheckDB_DOCLOG(grCredOrd(1).commonTab.isn,"77","N","1"," ",1)
    Call CheckDB_DOCLOG(grCredOrd(1).commonTab.isn,"77","E","1"," ",1)
    Call CheckDB_DOCLOG(grCredOrd(1).commonTab.isn,"77","M","3","¶ñ³Ýóí»É »Ý Ó¨³Ï»ñåáõÙÝ»ñÁ",1)
    
    'DOCS                                                     
    fBODY = "  TYPECODE:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28 10  USERID:  77  ACSBRANCH:00  ACSDEPART:1  "_
               &"DOCNUM:" & grCredOrd(1).commonTab.docN & "  DATE:20220901  ACCDB:77809533317  CUR:003  SUMMA:5000000  "_
               &"AIM:Ð³Ù³Ó³ÛÝ ³ñï. ³éùáõí³×³éùÇ å³ÛÙ³Ý³·ñÇ0000102030405060708091011121314151617181920212223242526272829303132333435363738394041424344454647484950  "    
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS", "fISN", grCredOrd(1).commonTab.isn, 1)
    Call CheckDB_DOCS(grCredOrd(1).commonTab.isn, "CrPkOrd ", "3", fBODY,1)
     
    'HIREST
    sumCur(0) = "2250000000.00"
    For i = 0 to grCredOrd(1).commonTab.cAccsCount - 1
        sumSQL = Replace (grCredOrd(1).commonTab.sum(i),"," , "")
        Call CheckDB_HIREST("01", cAccEditIsn(i) , "-" & sumCur(i) ,"003", "-" & sumSQL, 1)
    Next
    Call CheckDB_HIREST("01", dAccEditIsn , "2249953097.00" ,"003", "4999911.00", 1)
    
    'MEMORDERS
    Call CheckDB_MEMORDERS(grCredOrd(1).commonTab.isn,"CrPkOrd ","1","20220901","3","5000000.00","003",1)    
    
    Call Close_Window(wMDIClient, "frmPttel" )


'------------------------------------------------------------------
'------------Բացել Ստեղծված փաստաթղթեր թղթապանակը---------------------
'------------------------------------------------------------------
    folderDirect = "|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|êï»ÕÍí³Í ÷³ëï³ÃÕÃ»ñ"
    stDate = "010122"
    enDate = "010125"
    wUser = 77
    docType = "CrPkOrd "
    Call OpenCreatedDocFolder(folderDirect, stDate, enDate, wUser, docType)
    'Կատարել Դիտել գործողությունը
    If SearchInPttel("frmPttel" , 2, grCredOrd(1).commonTab.isn) Then
        BuiltIn.Delay(2000)
        Call wMainForm.MainMenu.Click (c_Allactions)
        Call wMainForm.PopupMenu.Click (c_View)
        BuiltIn.Delay(delay_middle)
        colReadOnlyArray = Array (True, True, True)
        Call Group_Cred_Order_Window_Check (grCredOrd(1), colReadOnlyArray)
        Call ClickCmdButton(1,"OK")
    Else 
        Log.Error  "Փաստաթուղթը չի գտնվել ստեղծված փաստաթղթեր աղյուսակում"
    End If    

'--------------------------------------------------------------------------    
'--------------------Ջնջել փաստաթուղթը---------------------------------------
'--------------------------------------------------------------------------   
    If SearchInPttel("frmPttel", 2, grCredOrd(1).commonTab.isn) Then
        Call wMainForm.MainMenu.Click(c_AllActions)
        Call wMainForm.PopupMenu.Click(c_Delete )
        If MessageExists(2, "ö³ëï³ÃáõÕÃÁ çÝç»ÉÇë` ÏÑ»é³óí»Ý Ýñ³ Ñ»ï Ï³åí³Í ËÙµ³ÛÇÝ " & vbCrLf &"Ó¨³Ï»ñåáõÙÝ»ñÁ") Then
           ' Սեղմել "Կատարել" կոճակը
           Call ClickCmdButton(5, "Î³ï³ñ»É")  
           If  MessageExists(1, "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ") Then
               ' Սեղմել "Այո" կոճակը
               Call ClickCmdButton(3, "²Ûá")  
           Else
               Log.Error"Հաղորդագրության պատուհանը չի բացվել" ,,,ErrorColor
           End If
        Else
            Log.Error"Հաղորդագրության պատուհանը չի բացվել" ,,,ErrorColor
        End If
    Else
        Log.Error "Document with N " & grCredOrd(1).commonTab.docN & " not found"
    End If 
    
    
    'SQL ստուգումներ
    Log.Message "'SQL Ստուգումներ հաշվառված Խմբային կրեդիտի օրդերը ջնջելուց հետո",,,SqlDivideColor
    
    'DOCLOG
    Call CheckQueryRowCount("DOCLOG","fISN",grCredOrd(1).commonTab.isn, 4)
    Call CheckDB_DOCLOG(grCredOrd(1).commonTab.isn,"77","N","1"," ",1)
    Call CheckDB_DOCLOG(grCredOrd(1).commonTab.isn,"77","E","1"," ",1)
    Call CheckDB_DOCLOG(grCredOrd(1).commonTab.isn,"77","M","3","¶ñ³Ýóí»É »Ý Ó¨³Ï»ñåáõÙÝ»ñÁ",1)
    Call CheckDB_DOCLOG(grCredOrd(1).commonTab.isn,"77","D","999"," ",1)
    
    'DOCS                                                     
    fBODY = "  TYPECODE:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28 10  USERID:  77  ACSBRANCH:00  ACSDEPART:1  "_
               &"DOCNUM:" & grCredOrd(1).commonTab.docN & "  DATE:20220901  ACCDB:77809533317  CUR:003  SUMMA:5000000  "_
               &"AIM:Ð³Ù³Ó³ÛÝ ³ñï. ³éùáõí³×³éùÇ å³ÛÙ³Ý³·ñÇ0000102030405060708091011121314151617181920212223242526272829303132333435363738394041424344454647484950  "    
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS", "fISN", grCredOrd(1).commonTab.isn, 1)
    Call CheckDB_DOCS(grCredOrd(1).commonTab.isn, "CrPkOrd ", "999", fBODY,1)

    'FODLERS
    Call CheckQueryRowCount("FOLDERS","fISN",grCredOrd(1).commonTab.isn,1)
    Call CheckDB_FOLDERS(dbFOLDERS(2),1) 
    
    'MEMORDERS
    Call CheckDB_MEMORDERS(grCredOrd(1).commonTab.isn,"CrPkOrd ","1","20220901","3","5000000.00","003",0)    
    
    Call Close_Window(wMDIClient, "frmPttel" )
    Call Close_AsBank()

End Sub

Sub Test_Initialize_Group_Cred_Order_Curr()
        
    sDate = "20050101"
    eDate = "20250101"
    
    'Ստեղծվող Խմբային կրեդիտի օրդերի արժեքներ
    Set grCredOrd(0) = New_Group_Credit_Order(0, 0, 0)
    With grCredOrd(0)
        .commonTab.branch = "00"
        .commonTab.dep = "1"
        .commonTab.mDate = "280222"
        .commonTab.accD = "77781633311"              
        .commonTab.cAccsCount = 2
        .commonTab.accC(0) = "00068360101"
        .commonTab.accC(1) = "00066610301"
        .commonTab.curr = "001"
        .commonTab.sum(0) = "4,444,444,444.00"
        .commonTab.sum(1) = "5,555,555,555.00"
        .commonTab.fullSum = "9,999,999,999.00"
        .commonTab.aim = "1"
        .attachTab.tabN = 2
    End With 
    
    'Խմբագրվող տվյալներ
    Set grCredOrd(1) = New_Group_Credit_Order(0, 0, 0)
    With grCredOrd(1)
        .commonTab.branch = "00"
        .commonTab.dep = "1"
        .commonTab.mDate = "010922"
        .commonTab.accD = "77809533317"              
        .commonTab.cAccsCount = 1
        .commonTab.accC(0) = "000008603  "
        .commonTab.curr = "003"
        .commonTab.sum(0) = "5,000,000.00"
        .commonTab.fullSum = "5,000,000.00"
        .commonTab.aim = "Ð³Ù³Ó³ÛÝ ³ñï. ³éùáõí³×³éùÇ å³ÛÙ³Ý³·ñÇ0000102030405060708091011121314151617181920212223242526272829303132333435363738394041424344454647484950"
        .attachTab.tabN = 2
    End With 
    'Աշխատանքային փաստաթղթերի ֆիլտրի տվյալներ
    Set workingDocs = New_MainAccWorkingDocuments()
    With workingDocs
        .startDate = grCredOrd(0).commonTab.mDate
        .endDate = grCredOrd(1).commonTab.mDate
    End With
    
    'Փաստաթղթի տպելու ձևի օրինակի ճանապարհ
    pathExp = Project.path & "Stores\MemorialOrder\Group_Cred_Order_With_Curr_Change_Exp.txt"

End Sub

Sub Intitialize_DB_Group_Cred_Order (fISN,fDOCN)
    Dim tday
    
    tday = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"20%y%m%d")
    Set dbFOLDERS(0) = New_DB_FOLDERS()
    With dbFOLDERS(0) 
        .fFOLDERID = "Oper.20220228"
        .fNAME = "CrPkOrd "
        .fKEY = fISN
        .fISN = fISN
        .fSTATUS = "1"
        .fCOM = "ÊÙµ³ÛÇÝ Ïñ»¹ÇïÇ ûñ¹»ñ"
                          
        .fSPEC = fDOCN & "7770077781633311                   9999999999.00001Üáñ                                                   "_
                 &"77                                                                                                1                 "_
                 &"                                                                                                                          "
        .fECOM = "Group Credit Order"
        .fDCBRANCH = "00 "
        .fDCDEPART = "1  "
    End With
    
    Set dbFOLDERS(1) = New_DB_FOLDERS()
    With dbFOLDERS(1) 
        .fFOLDERID = "Oper.20220901"
        .fNAME = "CrPkOrd "
        .fKEY = fISN
        .fISN = fISN
        .fSTATUS = "1"
        .fCOM = "ÊÙµ³ÛÇÝ Ïñ»¹ÇïÇ ûñ¹»ñ"       
        .fSPEC = fDOCN & "7770077809533317                      5000000.00003ÊÙµ³·ñíáÕ                                             "_
                 &"77                                                                                                "_
                 &"Ð³Ù³Ó³ÛÝ ³ñï. ³éùáõí³×³éùÇ å³ÛÙ³Ý³·ñÇ0000102030405060708091011121314151617181920212223242526272829303132333435363738394041424344454647484950"
        .fECOM = "Group Credit Order"
        .fDCBRANCH = "00 "
        .fDCDEPART = "1  "
    End With
    
    Set dbFOLDERS(2) = New_DB_FOLDERS()
    With dbFOLDERS(2) 
        .fFOLDERID = ".R."&tday
        .fNAME = "CrPkOrd "
        .fKEY = fISN
        .fISN = fISN
        .fSTATUS = "0"
        .fCOM = ""
        .fSPEC = Left_Align(Get_Compname_DOCLOG(fISN), 16) & "GlavBux ARMSOFT                       113  "
        .fECOM = ""
        .fDCBRANCH = "00 "
        .fDCDEPART = "1  "
    End With
    
End Sub
