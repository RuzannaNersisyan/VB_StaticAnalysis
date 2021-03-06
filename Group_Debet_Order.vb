'USEUNIT Library_Common
'USEUNIT Constants
'USEUNIT Mem_Order_Library
'USEUNIT Library_Colour
'USEUNIT Library_CheckDB
'USEUNIT Library_Contracts
'USEUNIT DAHK_Library_Filter
'USEUNIT Main_Accountant_Filter_Library
Option Explicit
'Test Case ID 182690
Dim sDATE, eDate, grDebOrd(2), workingDocs, filter_Pttel, pathExp, folderDirect, tday, sumSQL, curSumSQL(2), expMessage
Dim stDate, enDate, wUser, docType, wName, passNum, cliCode, paySysIn, paySysOut, acsBranch, i, fBODY, docGrid
Dim acsDepart, docISN, selectedView, expExcel, dAccIsn(3), cAccIsn, dAccEditIsn(3), cAccEditIsn, dbFOLDERS (3) 

Sub Group_Debet_Order_Test() 
    aCount = 3
    Call Test_Initialize_Group_Deb_Order()
    
    'Մուտք ծրագիր ARMSOFT Օգտագործողով
    Call Initialize_AsBank("bank", sDATE, eDATE)
    Call Login ("ARMSOFT")
    'Հաշիվների ISN-ների ստացում SQL աղյուսակներից
    For i = 0 to grDebOrd(0).commonTab.dAccsCount 
        dAccIsn(i) = GetAccountISN(grDebOrd(0).commonTab.accD(i))
    Next    
    cAccIsn = GetAccountISN(grDebOrd(0).commonTab.accC)
    For i = 0 to grDebOrd(1).commonTab.dAccsCount 
        dAccEditIsn(i) = GetAccountISN(grDebOrd(1).commonTab.accD(i))
    Next 
    cAccEditIsn = GetAccountISN(grDebOrd(1).commonTab.accC)

    'Մուտք Գլխավոր հաշվապահի ԱՇՏ| Հաշիվներ
    Log.Message  "Մուտք Գլխավոր հաշվապահի ԱՇՏ| Հաշիվներ",,,DivideColor
    Call ChangeWorkspace(c_ChiefAcc)
    Call OpenAccauntsFolder ("|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|Ð³ßÇíÝ»ñ","1","","","000","","","","","",0,"","","","","",0,0,0,"","","","","","ACCS","0")
    
    'Ֆիլտրել հաշիվներ թղթապանակը ըստ տիպ սյան
    Call Pttel_Filtering (filter_Pttel, "frmPttel")

'--------------------------------------------------------------------    
'----------------Ստեղծել Խմբային Դեբետի Օրդեր----------------------------
'--------------------------------------------------------------------   
    Log.Message  "Ստեղծել Խմբային Դեբետի Օրդեր",,,DivideColor
    Call Create_Group_Deb_Order(grDebOrd(0), "Î³ï³ñ»É", "frmPttel")
    'Ստուգել փաստաթղթի տպելու ձևը
    Call Group_Deb_Order_DocCheck (pathExp, 0)

    'SQL
    Log.Message "'SQL Ստուգումներ Խմբային Դեբետի օրդեր ստեղծելուց հետո",,,SqlDivideColor
    Call Intitialize_DB_Group_Deb_Order (grDebOrd(0).commonTab.isn , grDebOrd(0).commonTab.docN)
    
    'DOCLOG
    Call CheckQueryRowCount("DOCLOG","fISN",grDebOrd(0).commonTab.isn, 1)
    Call CheckDB_DOCLOG(grDebOrd(0).commonTab.isn,"77","N","1"," ",1)
    
    'DOCS   
    fBODY = "  TYPECODE:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28 10  USERID:  77  ACSBRANCH:01  ACSDEPART:2  "_
               &"DOCNUM:" & grDebOrd(0).commonTab.docN & "  DATE:20210214  ACCCR:00005650100  CUR:000  SUMMA:33000000  "_
               &"AIM:Ð³Ù³Ó³ÛÝ ³ñï. ³éùáõí³×³éùÇ å³ÛÙ³Ý³·ñÇ0000102030405060708091011121314151617181920212223242526272829303132333435363738394041424344454647484950  "   
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS", "fISN", grDebOrd(0).commonTab.isn, 1)
    Call CheckDB_DOCS(grDebOrd(0).commonTab.isn, "DbPkOrd ", "1", fBODY,1)
    
    'DOCSATTACH
    Call CheckQueryRowCount("DOCSATTACH", "fISN", grDebOrd(0).commonTab.isn, 1)
    Call CheckDB_DOCSATTACH(grDebOrd(0).commonTab.isn, grDebOrd(0).attachTab.fileName(0), 0, "", 1)
    
    'DOCSG
     Call CheckQueryRowCount("DOCSG", "fISN", grDebOrd(0).commonTab.isn, 9)
     For i = 0 to grDebOrd(0).commonTab.dAccsCount - 1
         Call CheckDB_DOCSG(grDebOrd(0).commonTab.isn,"SUBSUMS",grDebOrd(0).commonTab.dAccRowN(i) ,"ACCDB",grDebOrd(0).commonTab.accD(i),1)
     Next
     For i = 0 to grDebOrd(0).commonTab.dAccsCount - 1
         Call CheckDB_DOCSG(grDebOrd(0).commonTab.isn,"SUBSUMS",grDebOrd(0).commonTab.dAccRowN(i) ,"ACCDBNAME",grDebOrd(0).commonTab.dAccName(i),1)
     Next
     For i = 0 to grDebOrd(0).commonTab.dAccsCount - 1
         sumSQL = Replace (grDebOrd(0).commonTab.sum(i),"," , "")
         sumSQL = Replace (sumSQL, ".00", "")
         Call CheckDB_DOCSG(grDebOrd(0).commonTab.isn,"SUBSUMS", grDebOrd(0).commonTab.dAccRowN(i) , "SUMMA", sumSQL, 1)
    Next
    
    'FODLERS
    Call CheckQueryRowCount("FOLDERS","fISN",grDebOrd(0).commonTab.isn,1) 
    Call CheckDB_FOLDERS(dbFOLDERS(0),1)
    
    'HI
    Call CheckQueryRowCount("HI", "fBASE", grDebOrd(0).commonTab.isn, 6)
    For i = 0 to grDebOrd(0).commonTab.dAccsCount - 1
        sumSQL = Replace (grDebOrd(0).commonTab.sum(i),"," , "")
        Call Check_HI_CE_accounting ("20210214",grDebOrd(0).commonTab.isn , "11", cAccIsn ,sumSQL, "000", sumSQL, "MSC", "C")
        Call Check_HI_CE_accounting ("20210214",grDebOrd(0).commonTab.isn , "11", dAccIsn(i) ,sumSQL, "000", sumSQL, "MSC", "D")
    Next
     
    'HIREST
    For i = 0 to grDebOrd(0).commonTab.dAccsCount - 1
        sumSQL = Replace (grDebOrd(0).commonTab.sum(i),"," , "")
        Call CheckDB_HIREST("11", dAccIsn(i) , sumSQL ,"000", sumSQL, 1)
    Next
    Call CheckDB_HIREST("11", cAccIsn , "-33000000.00" ,"000", "-33000000.00", 1)
    
    Log.Message "Document ISN = " & grDebOrd(0).commonTab.isn,,, SqlDivideColor
    Log.Message "Document Number = " & grDebOrd(0).commonTab.DocN,,, DivideColor
    
    Call Close_Window(wMDIClient, "FrmSpr" )
    Call Close_Window(wMDIClient, "frmPttel" )
    
    'Բացել Աշխատանքային փաստաթղթեր թղթապանակը
    Log.Message  "Բացել Աշխատանքային փաստաթղթեր թղթապանակը",,,DivideColor
    Call GoTo_MainAccWorkingDocuments("|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|", workingDocs)
    grDebOrd(0).commonTab.mDate = "14/02/21"
    'Կատարել Դիտել գործողությունը
    Call View_Group_Deb_Order (grDebOrd(0), "frmPttel")

'--------------------------------------------------------------------------------    
'---------Արժույթի փոփոխման հետ կապված Error-ների ստուգում փաստաթուղթը խմբագրելիս---------
'--------------------------------------------------------------------------------
    If SearchInPttel("frmPttel",2, grDebOrd(0).commonTab.docN) Then
        BuiltIn.Delay(2000)
        Call wMainForm.MainMenu.Click (c_Allactions)
        Call wMainForm.PopupMenu.Click (c_ToEdit)
        BuiltIn.Delay(delay_middle)
        Set docGrid = wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame").VBObject("DocGrid")
        
        If  wMDIClient.WaitVBObject("frmASDocForm",3000).exists Then
            'Հաշիվ Դեբետ դաշտի լրացում 002 արժույթ ունեցող հաշվեհամարով
            wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame").VBObject("AsTypeFolder").VBObject("TDBMask").Keys("[Home]![End][Del]" & "000004502")
            Call ClickCmdButton(1,"Î³ï³ñ»É")
            expMessage = "Ð³ßíÇ ³ñï³ñÅáõÃ³ÛÝáõÃÛ³Ý ëË³É " & vbNewLine & "   Ð³ßÇí  -  01046563311 (³ñï³ñÅáõÛÃÇ Ïá¹ª 000) " & vbNewLine _ 
                        &"   Ð³ßíÇ ³ñï³ñÅáõÛÃÁ å»ïù ¿ ÉÇÝÇ  002 "
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

'--------------------------------------------------------------------    
'----------------Խմբագրել Խմբային Դեբետի Օրդերը----------------------------
'-------------------------------------------------------------------- 
    Call Edit_Group_Deb_Order (grDebOrd(0), grDebOrd(1), 0, "frmPttel")
    
    'SQL
    Log.Message "'SQL Ստուգումներ Խմբային Դեբետի օրդերը խմբագրելուց հետո",,,SqlDivideColor
    Call Intitialize_DB_Group_Deb_Order (grDebOrd(1).commonTab.isn , grDebOrd(1).commonTab.docN)
    'DOCLOG
    Call CheckQueryRowCount("DOCLOG","fISN",grDebOrd(1).commonTab.isn, 2)
    Call CheckDB_DOCLOG(grDebOrd(1).commonTab.isn,"77","N","1"," ",1)
    Call CheckDB_DOCLOG(grDebOrd(1).commonTab.isn,"77","E","1"," ",1)
    
    'DOCS   
    fBODY = "  TYPECODE:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28 10  USERID:  77  ACSBRANCH:01  ACSDEPART:2  "_
               &"DOCNUM:" & grDebOrd(1).commonTab.docN & "  DATE:20220308  ACCCR:00011830101  CUR:001  SUMMA:4100  AIM:b  "   
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS", "fISN", grDebOrd(1).commonTab.isn, 1)
    Call CheckDB_DOCS(grDebOrd(0).commonTab.isn, "DbPkOrd ", "1", fBODY,1)
    
    'DOCSG
    Call CheckQueryRowCount("DOCSG", "fISN", grDebOrd(1).commonTab.isn, 9)
    For i = 0 to grDebOrd(1).commonTab.dAccsCount - 1
        Call CheckDB_DOCSG(grDebOrd(1).commonTab.isn,"SUBSUMS",grDebOrd(1).commonTab.dAccRowN(i) ,"ACCDB",grDebOrd(1).commonTab.accD(i),1)
    Next
    For i = 0 to grDebOrd(1).commonTab.dAccsCount - 1
        Call CheckDB_DOCSG(grDebOrd(1).commonTab.isn,"SUBSUMS",grDebOrd(1).commonTab.dAccRowN(i) ,"ACCDBNAME",grDebOrd(1).commonTab.dAccName(i),1)
    Next
    For i = 0 to grDebOrd(1).commonTab.dAccsCount - 1
        sumSQL = Replace (grDebOrd(1).commonTab.sum(i),"," , "")
        sumSQL = Replace (sumSQL, ".00", "")
        Call CheckDB_DOCSG(grDebOrd(1).commonTab.isn,"SUBSUMS", grDebOrd(1).commonTab.dAccRowN(i) , "SUMMA", sumSQL, 1)
    Next
    
    'FODLERS
    Call CheckQueryRowCount("FOLDERS","fISN",grDebOrd(1).commonTab.isn,1) 
    Call CheckDB_FOLDERS(dbFOLDERS(1),1)

    'HI
    Call CheckQueryRowCount("HI", "fBASE", grDebOrd(1).commonTab.isn, 6)
    curSumSQL(0) = "600000.00" 
    curSumSQL(1) = "480000.00"
    curSumSQL(2) = "560000.00"
    For i = 0 to grDebOrd(1).commonTab.dAccsCount - 1
         sumSQL = Replace (grDebOrd(1).commonTab.sum(i),"," , "")
         Call Check_HI_CE_accounting ("20220308",grDebOrd(1).commonTab.isn , "11", cAccEditIsn ,curSumSQL(i), "001", sumSQL, "MSC", "C")
         Call Check_HI_CE_accounting ("20220308",grDebOrd(1).commonTab.isn , "11", dAccEditIsn(i) ,curSumSQL(i), "001", sumSQL, "MSC", "D")
    Next
     
    'HIREST
    For i = 0 to grDebOrd(1).commonTab.dAccsCount - 1
        sumSQL = Replace (grDebOrd(1).commonTab.sum(i),"," , "")
        Call CheckDB_HIREST("11", dAccEditIsn(i) , curSumSQL(i) ,"001", sumSQL, 1)
    Next
    Call CheckDB_HIREST("11", cAccEditIsn , "-1640000.00" ,"001", "-4100.00", 1)
    
    'Կատարել Դիտել գործողությունը
    grDebOrd(1).commonTab.mDate = "08/03/22"
    Call View_Group_Deb_Order (grDebOrd(1), "frmPttel")

'--------------------------------------------------------------------    
'----------------Հաշվառել Խմբային Դեբետի Օրդերը----------------------------
'-------------------------------------------------------------------- 
    Log.Message  "Փաստաթղթի հաշվառում",,,DivideColor
    If SearchInPttel("frmPttel", 2, grDebOrd(1).commonTab.docN) Then
       Call Register_Payment()
    End If
    
    'SQL
    Log.Message "'SQL Ստուգումներ Խմբային Դեբետի օրդերը հաշվառելուց հետո",,,SqlDivideColor
    Call Intitialize_DB_Group_Deb_Order (grDebOrd(1).commonTab.isn , grDebOrd(1).commonTab.docN)
    'DOCLOG
    Call CheckQueryRowCount("DOCLOG","fISN",grDebOrd(1).commonTab.isn, 3)
    Call CheckDB_DOCLOG(grDebOrd(1).commonTab.isn,"77","N","1"," ",1)
    Call CheckDB_DOCLOG(grDebOrd(1).commonTab.isn,"77","E","1"," ",1)
    Call CheckDB_DOCLOG(grDebOrd(1).commonTab.isn,"77","M","3","¶ñ³Ýóí»É »Ý Ó¨³Ï»ñåáõÙÝ»ñÁ",1)
    
    'DOCS        
    fBODY = "  TYPECODE:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28 10  USERID:  77  ACSBRANCH:01  ACSDEPART:2  "_
               &"DOCNUM:" & grDebOrd(1).commonTab.docN & "  DATE:20220308  ACCCR:00011830101  CUR:001  SUMMA:4100  AIM:b  "   
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS", "fISN", grDebOrd(1).commonTab.isn, 1)
    Call CheckDB_DOCS(grDebOrd(1).commonTab.isn, "DbPkOrd ", "3", fBODY,1)
    
    'DOCSG
    Call CheckQueryRowCount("DOCSG", "fISN", grDebOrd(1).commonTab.isn, 0)
    
    'FODLERS
    Call CheckQueryRowCount("FOLDERS","fISN",grDebOrd(1).commonTab.isn,0) 
    
    'HIREST
    Call CheckDB_HIREST("01", cAccEditIsn , "-1640000.00" ,"001", "-4100.00", 1)
    Call CheckDB_HIREST("01", dAccEditIsn (0) , "-104176969.10" ,"001", "-248994.81", 1)
    Call CheckDB_HIREST("01", dAccEditIsn (1) , "-356945.80" ,"001", "-822.43", 1)
    Call CheckDB_HIREST("01", dAccEditIsn (2) , "-95200000.00" ,"001", "-226600.00", 1)
    
    'MEMORDERS
    Call CheckDB_MEMORDERS(grDebOrd(1).commonTab.isn,"DbPkOrd  ","1","20220308","3","4100.00","001",1)
    
    Call Close_Window(wMDIClient, "frmPttel" )
    
    'Մուտք Հաշվառված վճարային փաստաթղթեր
    Log.Message  "Մուտք Հաշվառված վճարային փաստաթղթեր",,,DivideColor
    folderDirect = "|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|Ð³ßí³éí³Í í×³ñ³ÛÇÝ ÷³ëï³ÃÕÃ»ñ"
    stDate = "010120"
    enDate = "010125"
    wUser = ""
    docType = "DbPkOrd "
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
'--------------------------------------------------------------------    
'-----------Խմբագրել հաշվառված Խմբային Դեբետի Օրդերը-----------------------
'--------------------------------------------------------------------     
    
    Log.Message  "Խմբագրել հաշվառված խմբային դեբետի օրդերը",,,DivideColor                                                                                      
    Call Edit_Group_Deb_Order (grDebOrd(1), grDebOrd(2), 1, "frmPttel")
    
    'SQL
    Log.Message "SQL Ստուգումներ հաշվառված Խմբային Դեբետի օրդերը խմբագրելուց հետո",,,SqlDivideColor
    Call Intitialize_DB_Group_Deb_Order (grDebOrd(2).commonTab.isn , grDebOrd(2).commonTab.docN)
    'DOCLOG
    Call CheckQueryRowCount("DOCLOG","fISN",grDebOrd(2).commonTab.isn, 4)
    Call CheckDB_DOCLOG(grDebOrd(2).commonTab.isn,"77","N","1"," ",1)
    Call CheckDB_DOCLOG(grDebOrd(2).commonTab.isn,"77","E","1"," ",1)
    Call CheckDB_DOCLOG(grDebOrd(2).commonTab.isn,"77","M","3","¶ñ³Ýóí»É »Ý Ó¨³Ï»ñåáõÙÝ»ñÁ",1)
    Call CheckDB_DOCLOG(grDebOrd(2).commonTab.isn,"77","E","3"," ",1)
    'DOCS   
    fBODY = "  TYPECODE:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28 10  USERID:  77  ACSBRANCH:01  ACSDEPART:2  "_
               &"DOCNUM:" & grDebOrd(2).commonTab.docN & "  DATE:20220308  ACCCR:00011830101  CUR:001  SUMMA:4100  AIM:123456  "   
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS", "fISN", grDebOrd(2).commonTab.isn, 1)
    Call CheckDB_DOCS(grDebOrd(0).commonTab.isn, "DbPkOrd ", "3", fBODY,1)
       
    'Կատարել Դիտել գործողությունը
    Call View_Group_Deb_Order (grDebOrd(2), "frmPttel")
    
'--------------------------------------------------------------------    
'------------------Ջնջել Խմբային Դեբետի Օրդերը----------------------------
'--------------------------------------------------------------------  
    Log.Message  "Ջնջել հաշվառված խմբային Դեբետի օրդերը",,,DivideColor
    If SearchInPttel("frmPttel", 2, grDebOrd(1).commonTab.docN) Then
        Call wMainForm.MainMenu.Click(c_AllActions)
        Call wMainForm.PopupMenu.Click(c_Delete )
        If  MessageExists(2, "ö³ëï³ÃáõÕÃÁ çÝç»ÉÇë` ÏÑ»é³óí»Ý Ýñ³ Ñ»ï Ï³åí³Í ËÙµ³ÛÇÝ " & vbCrLf &"Ó¨³Ï»ñåáõÙÝ»ñÁ") Then
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
         Log.Error "Document with N " & grDebOrd(1).commonTab.docN & " not found"
     End If  

    'SQL
    Log.Message "SQL Ստուգումներ հաշվառված Խմբային Դեբետի օրդերը ջնջելուց հետո",,,SqlDivideColor
    Call Intitialize_DB_Group_Deb_Order (grDebOrd(2).commonTab.isn , grDebOrd(2).commonTab.docN)
    'DOCLOG
    Call CheckQueryRowCount("DOCLOG","fISN",grDebOrd(2).commonTab.isn, 5)
    Call CheckDB_DOCLOG(grDebOrd(2).commonTab.isn,"77","N","1"," ",1)
    Call CheckDB_DOCLOG(grDebOrd(2).commonTab.isn,"77","E","1"," ",1)
    Call CheckDB_DOCLOG(grDebOrd(2).commonTab.isn,"77","M","3","¶ñ³Ýóí»É »Ý Ó¨³Ï»ñåáõÙÝ»ñÁ",1)
    Call CheckDB_DOCLOG(grDebOrd(2).commonTab.isn,"77","E","3"," ",1)
    Call CheckDB_DOCLOG(grDebOrd(2).commonTab.isn,"77","D","999"," ",1)
  
    'DOCS 
    fBODY = "  TYPECODE:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28 10  USERID:  77  ACSBRANCH:01  ACSDEPART:2  "_
               &"DOCNUM:" & grDebOrd(2).commonTab.docN & "  DATE:20220308  ACCCR:00011830101  CUR:001  SUMMA:4100  AIM:123456    "   
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS", "fISN", grDebOrd(2).commonTab.isn, 1)
    Call CheckDB_DOCS(grDebOrd(2).commonTab.isn, "DbPkOrd ", "999", fBODY,1)

    'FODLERS
    Call CheckQueryRowCount("FOLDERS","fISN",grDebOrd(2).commonTab.isn,1)
    Call CheckDB_FOLDERS(dbFOLDERS(2),1) 
    
    'MEMORDERS
    Call CheckDB_MEMORDERS(grDebOrd(1).commonTab.isn,"DbPkOrd  ","1","20220308","3","4100.00","001",0)
     
    Call Close_Window(wMDIClient, "frmPttel" )
    Call Close_AsBank() 
End Sub

Sub Test_Initialize_Group_Deb_Order()

    'Ստեղծվող օրդեր
    Set grDebOrd(0) = New_Group_Debet_Order(1, 0, 0)
    With grDebOrd(0)
        .commonTab.branch = "01"
        .commonTab.dep = "2"
        .commonTab.mDate = "140221"
        .commonTab.accC = "00005650100"              
        .commonTab.dAccsCount = 3
        .commonTab.accD(0) = "77797163331"
        .commonTab.accD(1) = "01046563311"        
        .commonTab.accD(2) = "01063963311"
        .commonTab.curr = "000"
        .commonTab.sum(0) = "15,000,000.00"
        .commonTab.sum(1) = "10,000,000.00"
        .commonTab.sum(2) = "8,000,000.00"
        .commonTab.fullSum = "33,000,000.00"
        .commonTab.aim = "Ð³Ù³Ó³ÛÝ ³ñï. ³éùáõí³×³éùÇ å³ÛÙ³Ý³·ñÇ0000102030405060708091011121314151617181920212223242526272829303132333435363738394041424344454647484950"
        .attachTab.tabN = 2
        .attachTab.fileName(0) = "Group_Deb_Order_Exp.txt"
        .attachTab.addFiles(0) =  Project.Path & "Stores\MemorialOrder\" & .attachTab.fileName(0)
        .attachTab.tabN = 2
    End With
    'Առաջին խմբագրվող տվյալներ
    Set grDebOrd(1) = New_Group_Debet_Order(1, 0, 0)
    With grDebOrd(1)
        .commonTab.branch = "01"
        .commonTab.dep = "2"
        .commonTab.mDate = "080322"
        .commonTab.accC = "00011830101"
        .commonTab.dAccsCount = 3
        .commonTab.accD(0) = "00001770101"
        .commonTab.accD(1) = "01062143321"
        .commonTab.accD(2) = "01070223321"
        .commonTab.curr = "001"
        .commonTab.sum(0) = "1,500.00"
        .commonTab.sum(1) = "1,200.00"
        .commonTab.sum(2) = "1,400.00"
        .commonTab.fullSum = "4,100.00"
        .commonTab.aim = "b"
        .attachTab.tabN = 2
        .attachTab.fileName(0) = "Group_Deb_Order_Exp.txt"
        .attachTab.addFiles(0) = Project.Path & "Stores\MemorialOrder\" & .attachTab.fileName(0)
        .attachTab.tabN = 2
    End With
    'Հաշվառումից հետո խմբագրվող տվյալներ
    Set grDebOrd(2) = New_Group_Debet_Order(1, 0, 0)
    With grDebOrd(2)
         .commonTab.branch = "01"
        .commonTab.dep = "2"
        .commonTab.mDate = "08/03/22"
        .commonTab.accC = "00011830101"
        .commonTab.dAccsCount = 3
        .commonTab.accD(0) = "00001770101"
        .commonTab.accD(1) = "01062143321"
        .commonTab.accD(2) = "01070223321"
        .commonTab.curr = "001"
        .commonTab.sum(0) = "1,500.00"
        .commonTab.sum(1) = "1,200.00"
        .commonTab.sum(2) = "1,400.00"
        .commonTab.fullSum = "4,100.00"
        .commonTab.aim = "123456"
        .attachTab.tabN = 2
        .attachTab.fileName(0) = "Group_Deb_Order_Exp.txt"
        .attachTab.addFiles(0) = Project.Path & "Stores\MemorialOrder\" & .attachTab.fileName(0)
        .attachTab.tabN = 2
    End With
    'Հաշիվներ թղթապանակի ֆիլտր թղթապանակ մուտք գործելուց հետո ֆիլտրը բացելու դեպքում
    Set filter_Pttel = New_Filter_Pttel (2)
    With filter_Pttel
        .andOr (1) = 1 
        .colName (0) = 5
        .colName (1) = 5
        .cond (0) = 0
        .cond (1) = 0 
        .val (0) = "01"
        .val (1) = "03"
    End With
    'Աշխատանքային փաստաթղթեր թղթապանակի ֆիլտր
    Set workingDocs = New_MainAccWorkingDocuments()
    With workingDocs
         .startDate = "140221"
				     .endDate = "080322"
    End With
    'Տպելու ձևը ստուգելու օրինակի ճանապարհ
    pathExp = Project.Path & "Stores\MemorialOrder\Group_Deb_Order_Exp.txt"
    tday = aqConvert.DateTimeToStr(aqDateTime.Today)
    sDate = "20050101"
    eDate = "20260101"

End Sub


Sub Intitialize_DB_Group_Deb_Order (fISN,fDOCN)
    Dim tday
    
    tday = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"20%y%m%d")
    Set dbFOLDERS(0) = New_DB_FOLDERS()
    With dbFOLDERS(0) 
        .fFOLDERID = "Oper.20210214"
        .fNAME = "DbPkOrd "
        .fKEY = fISN
        .fISN = fISN
        .fSTATUS = "1"
        .fCOM = "ÊÙµ³ÛÇÝ ¹»µ»ïÇ ûñ¹»ñ"
        .fSPEC = fDOCN & "                7770000005650100     33000000.00000Üáñ                                                   "_
                 &"77                                                                                                "_
                 &"Ð³Ù³Ó³ÛÝ ³ñï. ³éùáõí³×³éùÇ å³ÛÙ³Ý³·ñÇ0000102030405060708091011121314151617181920212223242526272829303132333435363738394041424344454647484950"
        .fECOM = "Group Debit Order"
        .fDCBRANCH = "01 "
        .fDCDEPART = "2  "
    End With
    
    Set dbFOLDERS(1) = New_DB_FOLDERS()
    With dbFOLDERS(1) 
        .fFOLDERID = "Oper.20220308"
        .fNAME = "DbPkOrd "
        .fKEY = fISN
        .fISN = fISN
        .fSTATUS = "1"
        .fCOM = "ÊÙµ³ÛÇÝ ¹»µ»ïÇ ûñ¹»ñ"
        .fSPEC = fDOCN & "                7770000011830101         4100.00001ÊÙµ³·ñíáÕ                                             "_
                 &"77                                                                                                b                        "_
                 &"                                                                                                                   "
        .fECOM = "Group Debit Order"
        .fDCBRANCH = "01 "
        .fDCDEPART = "2  "
    End With
    
    Set dbFOLDERS(2) = New_DB_FOLDERS()
    With dbFOLDERS(2) 
        .fFOLDERID = ".R."&tday
        .fNAME = "DbPkOrd "
        .fKEY = fISN
        .fISN = fISN
        .fSTATUS = "0"
        .fCOM = ""
        .fSPEC = Left_Align(Get_Compname_DOCLOG(fISN), 16) & "GlavBux ARMSOFT                       113  "
        .fECOM = ""
        .fDCBRANCH = "01 "
        .fDCDEPART = "2  "
    End With
    
End Sub
