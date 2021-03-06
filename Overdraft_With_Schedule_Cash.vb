Option Explicit
'USEUNIT Library_Common  
'USEUNIT Subsystems_SQL_Library 
'USEUNIT Constants
'USEUNIT Overdraft_NewCases_Library
'USEUNIT Akreditiv_Library
'USEUNIT Library_Contracts
'USEUNIT Library_CheckDB

'Test Case ID 165843

Sub Overdraft_With_Schedule_Cash_Test()
  Dim fDATE, sDATE, my_vbObj
  Dim CalcAcc, FolderName, Exists, opDate, MesBox, Data, Num, CashOrNo, Name, NameLen, ColNum, Pttel, DateS, DateF,_
      DocType, Typ, Date, Key, Paid, ExpMoney, accNum, accCredit, Client
  Dim CalculatePercents_ISN, FadeDebt_ISN, RepaySchedule_ISN, GiveOverdradt_ISN 
  Dim QueryString, ExpSQLValue, SQL_IsEqual,attr,dbFOLDERS(2)
  Dim Overdraft
  
  DateS = "010218"
  DateF = "010218"
  
  ''1, Համակարգ մուտք գործել ARMSOFT օգտագործողով
  fDATE = "20260101"
  sDATE = "20140101"
  Call Initialize_AsBank("bank", sDATE, fDATE)
  Call Create_Connection()
  
'--------------------------------------
  Set attr = Log.CreateNewAttributes
  attr.BackColor = RGB(0, 255, 255)
  attr.Bold = True
  attr.Italic = True
'--------------------------------------   
  
  ''2, Անցում "Օվերդրաֆտ (տեղաբաշխված)" ԱՇՏ
  Call ChangeWorkspace(c_Overdraft)
  
  'Կատարել "Նոր պայմանագրի ստեղծում"
  CalcAcc = "00001850100"
  
'---------------------------------------------------------------------------------    
'  'Ջնջել Փաստաթղթերը
  'Ջնջել Տոկոսների կուտակում/պարտքերի մարում -ները
  'Մուտք գործել "Օվերդրաֆտ (տեղաբաշխված)/Օվերդրաֆտ ունեցող հաշիվներ"
  FolderName = "|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|"
  Call wTreeView.DblClickItem(FolderName & "úí»ñ¹ñ³ýï áõÝ»óáÕ Ñ³ßÇíÝ»ñ")
  Call Rekvizit_Fill("Dialog", 1, "General", "ACCMASK", CalcAcc) 
  Call ClickCmdButton(2, "Î³ï³ñ»É")
  '"Գործողություններ/Բոլոր գործողություններ/Թղթապանակներ/Պայմանագրի թղթապանակ"
  If wMDIClient.VBObject("frmPttel").VBObject("tdbgView").ApproxCount <> 0 Then
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_Folders & "|" & c_AgrFolder)
    Name = "¶ñ³ýÇÏáí ûí»ñ¹ñ³ýïÇ å³ÛÙ³Ý³·Çñ"
    NameLen = 30
    ColNum = 0
    Pttel = "_2"
    Exists = Find_Doc_By(Name, NameLen,ColNum, Pttel)
  
    If Exists Then
      'Ջնջել "050218" ասմաթվով Օվերդրաֆտի մարումը
      Date = "050218"
      Typ = "22"
      Key = "0"
      Call DeleteD(Date, Typ, Key)
  
      'Ջնջել "050218" ասմաթվով Օվերդրաֆտի մարումը
      Date = "040218"
      Typ = "51"
      Key = "0"
      Call DeleteD(Date, Typ, Key)
  
      'Ջնջել "050218" ասմաթվով Օվերդրաֆտի մարումը
      Date = "030218"
      Typ = "22"
      Key = "0"
      Call DeleteD(Date, Typ, Key)
  
      'Ջնջել "050218" ասմաթվով Օվերդրաֆտի մարումը
      Date = "020218"
      Typ = "51"
      Key = "0"
      Call DeleteD(Date, Typ, Key)
    End If
    
    wMDIClient.VBObject("frmPttel_2").Close
    wMDIClient.VBObject("frmPttel").Close
  
    'Ջնջել "Կանխիկ ելք"-ը
    'Մուտք գործել "Հաճախորդի սպասարկում և դրամարկղ / Հաշվառված վճարային փաստաթղթեր"
    Call ChangeWorkspace(c_CustomerService)
    Call wTreeView.DblClickItem("|Ð³×³Ëáñ¹Ç ëå³ë³ñÏáõÙ ¨ ¹ñ³Ù³ñÏÕ |Ð³ßí³éí³Í í×³ñ³ÛÇÝ ÷³ëï³ÃÕÃ»ñ")
    
    Client = "00000018"
    Call Rekvizit_Fill("Dialog", 1, "General", "PERN", "^A[Del]" & DateS) 
    Call Rekvizit_Fill("Dialog", 1, "General", "PERK", "^A[Del]" & DateF) 
    Call Rekvizit_Fill("Dialog", 1, "General", "DOCTYPE", DocType)
    Call Rekvizit_Fill("Dialog", 1, "General", "CLICODE", Client)
    Call ClickCmdButton(2, "Î³ï³ñ»É")

    If wMDIClient.VBObject("frmPttel").VBObject("tdbgView").ApproxCount = 1 Then
      Name = "úí»ñ¹ñ³ýïÇ ïñ³Ù³¹ñáõÙ"
      NameLen = 21
      ColNum = 9
      Pttel = ""
      Exists = Find_Doc_By(Name, NameLen,ColNum, Pttel)
    
      If Exists Then
        Call wMainForm.MainMenu.Click(c_Opers & "|" & c_Delete)
        Call ClickCmdButton(3, "²Ûá")
        Call ClickCmdButton(5, "²Ûá")
      End If
    End If 
    wMDIClient.VBObject("frmPttel").Close
  
    'Ջնջել "Աշխատանքային փաստաթղթեր"-ից
    'Մուտք գործել "Հաճախորդի սպասարկում և դրամարկղ / Աշխատանքային փաստաթղթեր"
    Call wTreeView.DblClickItem("|Ð³×³Ëáñ¹Ç ëå³ë³ñÏáõÙ ¨ ¹ñ³Ù³ñÏÕ |²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ")
  
    Call Rekvizit_Fill("Dialog", 1, "General", "PERN", "^A[Del]" & DateS) 
    Call Rekvizit_Fill("Dialog", 1, "General", "PERK", "^A[Del]" & DateF)  
    Call Rekvizit_Fill("Dialog", 1, "General", "DOCTYPE", DocType)
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    If wMDIClient.VBObject("frmPttel").VBObject("tdbgView").ApproxCount = 1 Then
      Name = "úí»ñ¹ñ³ýïÇ ïñ³Ù³¹ñáõÙ"
      NameLen = 21
      ColNum = 12
      Pttel = ""
      Exists = Find_Doc_By(Name, NameLen,ColNum, Pttel)
    
      If Exists Then
        Call wMainForm.MainMenu.Click(c_Opers & "|" & c_Delete)
        Call ClickCmdButton(3, "²Ûá")
        Call ClickCmdButton(5, "²Ûá")
      End If
    End If 
    wMDIClient.VBObject("frmPttel").Close

  
    'Պայմանագրի թղթապանակից Ջնջել "Գումարի տրամադրում...", "Գրաֆիկով օվերդրաֆտ" պայմանագրերը 
    Call ChangeWorkspace(c_Overdraft)
    'Մուտք գործել "Օվերդրաֆտ (տեղաբաշխված)/Օվերդրաֆտ ունեցող հաշիվներ"
    FolderName = "|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|"
    Call wTreeView.DblClickItem(FolderName & "úí»ñ¹ñ³ýï áõÝ»óáÕ Ñ³ßÇíÝ»ñ")
    Call Rekvizit_Fill("Dialog", 1, "General", "ACCMASK", CalcAcc) 
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    '"Գործողություններ/Բոլոր գործողություններ/Թղթապանակներ/Պայմանագրի թղթապանակ"
    If wMDIClient.VBObject("frmPttel").VBObject("tdbgView").ApproxCount <> 0 Then  
      Call wMainForm.MainMenu.Click(c_AllActions)
      Call wMainForm.PopupMenu.Click(c_Folders & "|" & c_AgrFolder)
   
      '''''Պայմանագրի թղթապանակից Ջնջել "Գումարի տրամադրում...", Գրաֆիկոց օվերդրաֆտի պայմանագիրը
  
      Name = "¶áõÙ³ñÇ ïñ³Ù³¹ñáõÙ"
      NameLen = 18
      ColNum = 0
      Pttel = "_2"
      Exists = Find_Doc_By(Name, NameLen,ColNum, Pttel)
      If Exists Then
        Call wMainForm.MainMenu.Click(c_Opers & "|" & c_Delete)
        Call ClickCmdButton(3, "²Ûá")
      End If
      wMDIClient.VBObject("frmPttel_2").Close
    End If   
  
    If wMDIClient.VBObject("frmPttel").VBObject("tdbgView").ApproxCount = 1 Then
      Call wMainForm.MainMenu.Click(c_Opers & "|" & c_Delete)
      Call ClickCmdButton(3, "²Ûá")
    End If 
  End If 
  wMDIClient.VBObject("frmPttel").Close

'---------------------------------------------------------------------------------  
  Call Log.Message("Գրաֆիկով Օվերդրաֆտ պայմանագրի ստեղծում",,,attr)
  Set Overdraft = New_Overdraft()
  With Overdraft
    .CreditCard = 0
    .DocType = "¶ñ³ýÇÏáí ûí»ñ¹ñ³ýïÇ å³ÛÙ³Ý³·Çñ" 
    .CalcAcc = "00001850100"                                    
    .Limit = 100000
    .Date = "010218" 
    .GiveDate = "010218"
    .Term = "010219"
    .DateFillType = 1
    .Percent = 18
    .NonUsedPercent = 0
    .PayDates = 5
    .PaperCode = 123
    Call .CreatePlOverdraft(FolderName & "Üáñ å³ÛÙ³Ý³·ñÇ ëï»ÕÍáõÙ")
  
    Log.Message(.DocNum)
    ''SQL ստուգում պայամանգիր ստեղցելուց հետո: 
      ''CONTRACTS
      QueryString = "select count(*) from CONTRACTS where fDGISN = " & .fBASE &_
                      "and fDGAGRTYPE = 'C' and fDGMODTYPE = 3 and fDGAGRKIND = '8L'" &_
                      "and fDGSTATE = 206 and fDGSUMMA = 100000.00 and fDGALLSUMMA = 0.00"
      ExpSQLValue = 1
      colNum = 0
      SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, colNum)
      If Not SQL_IsEqual Then
        Log.Error("QueryString = " & QueryString & ":  Expected result = " & ExpSQLValue)
      End If  
                                
      ''FOLDERS
      QueryString = "select count(*) from FOLDERS where fISN = " & .fBASE 
      ExpSQLValue = 3
      colNum = 0
      SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, colNum)
      If Not SQL_IsEqual Then
        Log.Error("QueryString = " & QueryString & ":  Expected result = " & ExpSQLValue)
      End If                          
                 
    Set MesBox = AsBank.WaitVBObject("frmAsMsgBox", 2000)
    If MesBox.Exists Then   
      Sys.Process("Asbank").VBObject("frmAsMsgBox").VBObject("cmdButton").ClickButton     
    End If
                                
    ''4.Մարման գրաֆիկի նշանակում
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_RepaySchedule)  
    
    Name = "Ø³ñáõÙÝ»ñÇ ·ñ³ýÇÏ"
    NameLen = 17
    ColNum = 0
    Pttel = ""
    Exists = Find_Doc_By(Name, NameLen,ColNum, Pttel)
    If Exists Then 
      Call wMainForm.MainMenu.Click(c_AllActions)
      Call wMainForm.PopupMenu.Click(c_View)
      RepaySchedule_ISN = wMDIClient.VBObject("frmASDocForm").DocFormCommon.Doc.isn
      wMDIClient.VBObject("frmASDocForm").Close
    End If
  
    ''SQL ստուգում Մարման գրաֆիկ ստեղցելուց հետո: 
      ''AGRSCHEDULE
      QueryString = "select count(*) from AGRSCHEDULE where fBASE= '" & RepaySchedule_ISN & "'" &_
                     "and fKIND = 9 and fTYPE = 0 and fINC = 1"
      ExpSQLValue = 1
      colNum = 0
      SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, colNum)
      If Not SQL_IsEqual Then
        Log.Error("QueryString = " & QueryString & ":  Expected result = " & ExpSQLValue)
      End If  
    
      ''AGRSCHEDULEVALUES
       QueryString = "select count(*) from AGRSCHEDULEVALUES where fAGRISN = " & .fBASE &_
                      "and fSUM = 0.00 and (fVALUETYPE = 1 or fVALUETYPE = 2)"
      ExpSQLValue = 2
      colNum = 0
      SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, colNum)
      If Not SQL_IsEqual Then
        Log.Error("QueryString = " & QueryString & ":  Expected result = " & ExpSQLValue)
      End If     
       
      ''FOLDERS
      QueryString = "select count(*) from FOLDERS where fISN = " & RepaySchedule_ISN
      ExpSQLValue = 1
      colNum = 0
      SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, colNum)
      If Not SQL_IsEqual Then
        Log.Error("QueryString = " & QueryString & ":  Expected result = " & ExpSQLValue)
      End If
  
    Set MesBox = AsBank.WaitVBObject("frmAsMsgBox", 2000)
    If MesBox.Exists Then   
      Sys.Process("Asbank").VBObject("frmAsMsgBox").VBObject("cmdButton").ClickButton     
    End if  
  
    ''5.Այլ վճարումների գրաֆիկի նշանակում
    Data = Find_Data ("¶ñ³ýÇÏáí ûí»ñ¹ñ³ýïÇ å³ÛÙ³Ý³·Çñ- "& Trim(.DocNum) &" {ä»ïñáëÛ³Ý ä»ïñáë}",0)
    If Not Data then
       Call Log.Error("Փաստաթուղթը չի գտնվել") 
       Exit Sub
     End If
     
    Call ContractAction (c_OtherPaySchedule)
    wMDIClient.VBObject("frmASDocForm").VBObject("CmdOk_2").ClickButton

    Set MesBox = AsBank.WaitVBObject("frmAsMsgBox", 2000)
    If MesBox.Exists Then   
      Sys.Process("Asbank").VBObject("frmAsMsgBox").VBObject("cmdButton").ClickButton     
    End if   
  
    ''6."Գրաֆիկով օվերդրաֆտային պայմանագրի" համար կատարել "Գործողություններ/Բոլոր գործողություններ/Ուղարկել հաստատման " գործողությունը 
    'կանգնել պայմանագրի վրա
    Data = Find_Data ("¶ñ³ýÇÏáí ûí»ñ¹ñ³ýïÇ å³ÛÙ³Ý³·Çñ- "& Trim(.DocNum) &" {ä»ïñáëÛ³Ý ä»ïñáë}",0)
    If Not Data then
       call Log.Error("Փաստաթուղթը չի գտնվել") 
       exit Sub
     End If
    call ContractAction(c_SendToVer)
    Sys.Process("Asbank").VBObject("frmAsMsgBox").VBObject("cmdButton").ClickButton 
    BuiltIn.Delay(2000)
    wMDIClient.VBObject("frmPttel").Close
  
    ''7.Մուտք գործել Հաստատվող փաստաթղթեր 1 թղթապանակ - Պայմանագիրը պետք է առկա լինի
    Call wTreeView.DblClickItem("|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I")
    Sys.Process("Asbank").VBObject("frmAsUstPar").VBObject("TabFrame").VBObject("AsTpComment").VBObject("TDBComment").Keys(.DocNum & "[Tab]") 
    Sys.Process("Asbank").VBObject("frmAsUstPar").VBObject("CmdOK").ClickButton
    Builtin.Delay(2000)
    Set my_vbObj = wMDIClient.VBObject("frmPttel").VBObject("tdbgView")
    If my_vbObj.ApproxCount <> 1 Then
      Call Log.Error("Պայմանագիրը առկա չէ Հաստատվող փաստաթղթեր 1 թղթապանակում:")
      Exit Sub
    End If
  
    ''8.Վավերացնել պայմանագիրը    
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_ToConfirm)
    wMDIClient.vbObject("frmASDocForm").vbObject("CmdOk_2").Click()
  
    wMDIClient.VBObject("frmPttel").Close
   
    ''9.Մուտք գործել "Օվերդրաֆտ (Տեղաբաշխված )/ Պայմանագրեր" թղթապանակ - Պայմանագիրը պետք է առկա լինի:
    FolderName = "|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|"
    Exists = LetterOfCredit_Filter_Fill(FolderName, .DocLevel, .DocNum)
    If (Not Exists) Then
      call Log.Error("Պայմանագիրը առկա չէ")
      exit sub
    End If  
  
    Call Log.Message("Օվերդրաֆտի տրամադրում",,,attr)
    CashOrNo = "1"
    accNum = ""
    accCredit = ""
    GiveOverdradt_ISN= Give_Overdradt(.GiveDate, .Limit, CashOrNo, Num, accNum, accCredit)
   
    ''SQL ստուգում Օվերդրաֆտ տրամադրելուց հետո:   
      BuiltIn.Delay(delay_small) 
      ''CAGRACCS
      QueryString = "select count(*) from CAGRACCS where fAGRISN = " & .fBASE 
      ExpSQLValue = 1
      colNum = 0
      SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, colNum)
      If Not SQL_IsEqual Then
        Log.Error("QueryString = " & QueryString & ":  Expected result = " & ExpSQLValue)
      End If
    
      ''FOLDERS
      QueryString = "select count(*) from FOLDERS where fISN = " & GiveOverdradt_ISN 
      ExpSQLValue = 5
      colNum = 0
      SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, colNum)
      If Not SQL_IsEqual Then
        Log.Error("QueryString = " & QueryString & ":  Expected result = " & ExpSQLValue)
      End If 
    
      ''FOLDERS
      QueryString = "select count(*) from FOLDERS where fISN= " & .fBASE
      ExpSQLValue = 5
      colNum = 0
      SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, colNum)
      If Not SQL_IsEqual Then
        Log.Error("QueryString = " & QueryString & ":  Expected result = " & ExpSQLValue)
      End If 
      
      Set dbFOLDERS(1) = New_DB_FOLDERS()
          dbFOLDERS(1).fFOLDERID = "LOANREGISTER"
          dbFOLDERS(1).fNAME = "C3Univer"
          dbFOLDERS(1).fKEY = .fBASE 
          dbFOLDERS(1).fISN = .fBASE 
          dbFOLDERS(1).fSTATUS = 1
          dbFOLDERS(1).fCOM = "ä»ïñáëÛ³Ý ä»ïñáë"
          dbFOLDERS(1).fSPEC = "C38"& Trim(.DocNum) &"          123                               0                                                                                                                                                             0.00                                                                                                                                                                                                                                                                                               "

      Set dbFOLDERS(2) = New_DB_FOLDERS()
          dbFOLDERS(2).fFOLDERID = "LOANREGISTER2"
          dbFOLDERS(2).fNAME = "C3Univer"
          dbFOLDERS(2).fKEY = .fBASE 
          dbFOLDERS(2).fISN = .fBASE 
          dbFOLDERS(2).fSTATUS = "1"
          dbFOLDERS(2).fCOM = "ä»ïñáëÛ³Ý ä»ïñáë"
          dbFOLDERS(2).fSPEC = "0"
        
      Call CheckDB_FOLDERS(dbFOLDERS(1), 1)
      Call CheckDB_FOLDERS(dbFOLDERS(2), 1)

      ''HI
      QueryString = "select count(*) from HI where fBASE = " & .fBASE &_
                     "and fSUM = 100000.00 and fCURSUM = 100000.00 and fTYPE = 02"
      ExpSQLValue = 2
      colNum = 0
      SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, colNum)
      If Not SQL_IsEqual Then
        Log.Error("QueryString = " & QueryString & ":  Expected result = " & ExpSQLValue)
      End If

      ''HIF
      QueryString = "select count(*) from HIF where fBASE= " & .fBASE 
      ExpSQLValue = 19
      colNum = 0
      SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, colNum)
      If Not SQL_IsEqual Then
        Log.Error("QueryString = " & QueryString & ":  Expected result = " & ExpSQLValue)
      End If
    
    BuiltIn.Delay(2000)
    wMDIClient.VBObject("frmPttel").Close
  
    ''11.Մուտք գործել "Հաճախորդի սպասարկում և դրամարկղ / Աշխատանքային փաստաթղթեր " թղթապանակ - "Կանխիկ ելք " փաստաթուղթը պետք է առկա լինի :
    Call ChangeWorkspace(c_CustomerService)
    Call wTreeView.DblClickItem("|Ð³×³Ëáñ¹Ç ëå³ë³ñÏáõÙ ¨ ¹ñ³Ù³ñÏÕ |²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ")
    Sys.Process("Asbank").VBObject("frmAsUstPar").VBObject("TabFrame").VBObject("TDBDate").Keys(.Date & "[Tab]")
    Sys.Process("Asbank").VBObject("frmAsUstPar").VBObject("TabFrame").VBObject("TDBDate_2").Keys(.Date & "[Tab]")
    Sys.Process("Asbank").VBObject("frmAsUstPar").VBObject("CmdOK").ClickButton
  
    BuiltIn.Delay(2000)
    Data = Find_Data (Num,2)
    If Not Data then
       call Log.Error("Փաստաթուղթը չի գտնվել") 
       exit Sub
    End If
   
    ''12."Կանխիկ ելք" փաստաթուղթը ուղարկել հաստատման:
    call ContractAction(c_SendToVer)
    Sys.Process("Asbank").VBObject("frmAsUstPar").VBObject("TabFrame").VBObject("ASTypeTree").VBObject("TDBMask").Keys("001" & "[Tab]")
    Sys.Process("Asbank").VBObject("frmAsUstPar").VBObject("CmdOK").ClickButton
  
    'Փակել "Աշխատանքային փաստաթղթեր" պատուհանը 
    BuiltIn.Delay(2000)
    wMDIClient.VBObject("frmPttel").Close
   
    ''13.Մուտք գործել "Հաստատող 1 ԱՇՏ/Հաստատվող վճարային փաստաթղթեր " թղթապանակ - "Կանխիկ ելք " փաստաթուղթը պետք է առկա լինի : 
    Call ChangeWorkspace(c_Verifier1)
    
    Dim VerificationDoc
    Set VerificationDoc = New_VerificationDocument()
        VerificationDoc.User = "77"
        
    Call GoToVerificationDocument("|Ð³ëï³ïáÕ I ²Þî|Ð³ëï³ïíáÕ í×³ñ³ÛÇÝ ÷³ëï³ÃÕÃ»ñ",VerificationDoc)
    
    BuiltIn.Delay(2000)
    Data = Find_Data (Num, 3)
    If Not Data then
      Call Log.Error("Փաստաթուղթը չի գտնվել Հաստատվող վճարային փաստաթղթեր թղթապանակում:") 
      Exit Sub
    End If
  
    ''14.Վավերացնել փաստաթուղթը 
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_ToConfirm)
    wMDIClient.vbObject("frmASDocForm").vbObject("CmdOk_2").Click()

    BuiltIn.Delay(2000)
    wMDIClient.VBObject("frmPttel").Close
  
    ''15.'Մուտք գործել "Օվերդրաֆտ (տեղաբաշխված)/ Պայմանգրեր" թղթապանակ - Պայմանագիրը պետք է առկա լինի:
    Call ChangeWorkspace(c_Overdraft)
    FolderName = "|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|"
    Exists = LetterOfCredit_Filter_Fill(FolderName, .DocLevel, .DocNum)
    If Not Exists Then
      Call Log.Error("Պայմանագիրը առկա չէ Պայմանգրեր թղթապանակում:")
      Exit Sub
    End If  
  
    Call Log.Message("Օվերդրաֆտի տոկոսների հաշվարկ",,,attr)
    BuiltIn.Delay(2000)
    wMDIClient.VBObject("frmPttel").Close
    opDate = "020218"
    CalculatePercents_ISN = Overdraft_Percent_Accounting(.DocNum,opDate)
  
    ''SQL ստուգում Օվերդրաֆտի տոկոսների հաշվարկից հետո:      
        ''AGRSCHEDULEVALUES
        QueryString = "select count(*) from AGRSCHEDULEVALUES where fAGRISN = " & .fBASE 
        ExpSQLValue = 28
        colNum = 0
        SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, colNum)
        If Not SQL_IsEqual Then
          Log.Error("QueryString = " & QueryString & ":  Expected result = " & ExpSQLValue)
        End If
      
        ''HI
        QueryString = "select count(*) from HI where fBASE= " & CalculatePercents_ISN &_
                       "and fSUM = 98.70  and fCURSUM = 98.70"
        ExpSQLValue = 4
        colNum = 0
        SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, colNum)
        If Not SQL_IsEqual Then
          Log.Error("QueryString = " & QueryString & ":  Expected result = " & ExpSQLValue)
        End If
    
        ''HIF
        QueryString = "select count(*) from HIF where fBASE = " & CalculatePercents_ISN &_
                       "and fOBJECT = " & .fBASE & " and fSUM = 0.00 and fCURSUM = 0.00"
        ExpSQLValue = 1
        colNum = 0
        SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, colNum)
        If Not SQL_IsEqual Then
          Log.Error("QueryString = " & QueryString & ":  Expected result = " & ExpSQLValue)
        End If
    
        ''HIR
        QueryString = "select count(*) from HIR where fBASE = " & CalculatePercents_ISN &_
                       "and fOBJECT = "& .fBASE & " and fTYPE = 'R2' and fCURSUM = 98.70"
        ExpSQLValue = 1
        colNum = 0
        SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, colNum)
        If Not SQL_IsEqual Then
          Log.Error("QueryString = " & QueryString & ":  Expected result = " & ExpSQLValue)
        End If
      
        ''HIRREST
        QueryString = "select count(*) from HIRREST where fOBJECT = " & .fBASE &_
                       "and (fLASTREM = 100000.00 or fLASTREM = 98.70) and fSTARTREM = 0.00 and (fTYPE = 'R1' or fTYPE = 'R2')"
        ExpSQLValue = 2
        colNum = 0
        SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, colNum)
        If Not SQL_IsEqual Then
          Log.Error("QueryString = " & QueryString & ":  Expected result = " & ExpSQLValue)
        End If
    
        ''HIT
        QueryString = "select count(*) from HIT where fOBJECT = " & .fBASE &_ 
                       "and fCURSUM = 98.70 and fTYPE = 'N2'"
        ExpSQLValue = 1
        colNum = 0
        SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, colNum)
        If Not SQL_IsEqual Then
          Log.Error("QueryString = " & QueryString & ":  Expected result = " & ExpSQLValue)
        End If  
  
    Call Log.Message("Օվերդրաֆտի պարտքերի մարում",,,attr)
    wMDIClient.VBObject("frmPttel").Close
    opDate = "030218"
    Paid = "1000"
    FadeDebt_ISN = Overdraft_Repayment_Operation(.DocNum, opDate, Paid, Null, Null)
  
    ''SQL ստուգում Օվերդրաֆտի պարտքերի մարումից հետո:
        BuiltIn.Delay(delay_small)
      
        ''AGRSCHEDULEVALUES
        QueryString = "select count(*) from AGRSCHEDULEVALUES where fAGRISN = " & .fBASE 
        ExpSQLValue = 54
        colNum = 0
        SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, colNum)
        If Not SQL_IsEqual Then
          Log.Error("QueryString = " & QueryString & ":  Expected result = " & ExpSQLValue)
        End If
      
        ''FOLDERS
        QueryString = "select count(*) from FOLDERS where fISN = " & .fBASE 
        ExpSQLValue = 5
        colNum = 0
        SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, colNum)
        If Not SQL_IsEqual Then
          Log.Error("QueryString = " & QueryString & ":  Expected result = " & ExpSQLValue)
        End If
      
        ''HIR
        QueryString = "select count(*) from HIR where fBASE= " & FadeDebt_ISN &_
                       "and fCURSUM = 1000.00 and fTYPE = 'R1'"
        ExpSQLValue = 1
        colNum = 0
        SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, colNum)
        If Not SQL_IsEqual Then
          Log.Error("QueryString = " & QueryString & ":  Expected result = " & ExpSQLValue)
        End If
      
        ''HIRREST
        QueryString = "select count(*) from HIRREST where fOBJECT = " & .fBASE &_
                       "and (fLASTREM = 99000.00 or fLASTREM = 98.70) and fSTARTREM = 0.00 and (fTYPE = 'R1' or fTYPE = 'R2')"
        ExpSQLValue = 2
        colNum = 0
        SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, colNum)
        If Not SQL_IsEqual Then
          Log.Error("QueryString = " & QueryString & ":  Expected result = " & ExpSQLValue)
        End If
   
    Call Log.Message("Օվերդրաֆտի տոկոսների հաշվարկ",,,attr)
    wMDIClient.VBObject("frmPttel").Close
    opDate = "040218"
    CalculatePercents_ISN = Overdraft_Percent_Accounting(.DocNum,opDate)
  
      ''SQL ստուգում Օվերդրաֆտի տոկոսների հաշվարկից հետո:      
        ''AGRSCHEDULEVALUES
        QueryString = "select count(*) from AGRSCHEDULEVALUES where fAGRISN = " & .fBASE
        ExpSQLValue = 54
        colNum = 0
        SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, colNum)
        If Not SQL_IsEqual Then
          Log.Error("QueryString = " & QueryString & ":  Expected result = " & ExpSQLValue)
        End If
 
        ''HI
        QueryString = "select count(*) from HI where fBASE = " & CalculatePercents_ISN &_
                       "and  fSUM = 97.60  and fCURSUM = 97.60"
        ExpSQLValue = 4
        colNum = 0
        SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, colNum)
        If Not SQL_IsEqual Then
          Log.Error("QueryString = " & QueryString & ":  Expected result = " & ExpSQLValue)
        End If
    
        ''HIR
        QueryString = "select count(*) from HIR where fBASE = " & CalculatePercents_ISN &_
                       "and fOBJECT = " & .fBASE & " and (fCURSUM = 97.60 or fCURSUM = 196.30 or fCURSUM = 6692.30)"
        ExpSQLValue = 3
        colNum = 0
        SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, colNum)
        If Not SQL_IsEqual Then
          Log.Error("QueryString = " & QueryString & ":  Expected result = " & ExpSQLValue)
        End If
      
        ''HIRREST
        QueryString = "select count(*) from HIRREST where fOBJECT = " & .fBASE &_
                       "and (fLASTREM = 99000.00 or fLASTREM = 196.30 or fLASTREM = 6692.30) and fSTARTREM = 0.00 and (fPENULTREM = 100000.00 or fPENULTREM = 98.70 or fPENULTREM = 0.00 )"
        ExpSQLValue = 4
        colNum = 0
        SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, colNum)
        If Not SQL_IsEqual Then
          Log.Error("QueryString = " & QueryString & ":  Expected result = " & ExpSQLValue)
        End If
    
        ''HIT
        QueryString = "select count(*) from HIT where fOBJECT = " & .fBASE &_ 
                       "and (fCURSUM = 98.70 or fCURSUM = 97.60)and fTYPE = 'N2'"
        ExpSQLValue = 2
        colNum = 0
        SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, colNum)
        If Not SQL_IsEqual Then
          Log.Error("QueryString = " & QueryString & ":  Expected result = " & ExpSQLValue)
        End If  
    
    Call Log.Message("Օվերդրաֆտի պարտքերի մարում",,,attr)
  
    wMDIClient.VBObject("frmPttel").Close
    ExpMoney = "6,888.60"
    opDate = "050218"
    FadeDebt_ISN = Overdraft_Repayment_Check(.DocNum,opDate, ExpMoney)
  
      ''SQL ստուգում Օվերդրաֆտի պարտքերի մարումից հետո: 
      BuiltIn.Delay(delay_small)
    
        ''AGRSCHEDULEVALUES
        QueryString = "select count(*) from AGRSCHEDULEVALUES where fAGRISN = " & .fBASE
        ExpSQLValue = 78
        colNum = 0
        SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, colNum)
        If Not SQL_IsEqual Then
          Log.Error("QueryString = " & QueryString & ":  Expected result = " & ExpSQLValue)
        End If
      
                
        ''HIR
        QueryString = "select count(*) from HIR where fBASE = " & FadeDebt_ISN &_
                       "and (fCURSUM = 6692.30 or fCURSUM = 196.30)"
        ExpSQLValue = 4
        colNum = 0
        SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, colNum)
        If Not SQL_IsEqual Then
          Log.Error("QueryString = " & QueryString & ":  Expected result = " & ExpSQLValue)
        End If
      
        BuiltIn.Delay(delay_small)
      
        ''HIRREST
        QueryString = "select count(*) from HIRREST where fOBJECT = " & .fBASE &_
                       "and (fLASTREM = 92307.70 or fLASTREM = 0.00) and fSTARTREM = 0.00 and (fPENULTREM = 99000.00 or fPENULTREM = 196.30 or fPENULTREM = 0.00 )"
        ExpSQLValue = 4
        colNum = 0
        SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, colNum)
        If Not SQL_IsEqual Then
          Log.Error("QueryString = " & QueryString & ":  Expected result = " & ExpSQLValue)
        End If
  
    wMDIClient.VBObject("frmPttel").Close
 
  '---------------------------------------------------------------------------------    
    Call Log.Message("Ջնջել բոլոր փաստաթղթերը",,,attr)
    DocType = "KasRsOrd"

    'Ջնջել Տոկոսների կուտակում/պարտքերի մարում -ները
    'Մուտք գործել "Օվերդրաֆտ (տեղաբաշխված)/Օվերդրաֆտ ունեցող հաշիվներ"
    FolderName = "|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|"
    Call wTreeView.DblClickItem(FolderName & "úí»ñ¹ñ³ýï áõÝ»óáÕ Ñ³ßÇíÝ»ñ")
    Call Rekvizit_Fill("Dialog", 1, "General", "ACCMASK", .CalcAcc) 
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    '"Գործողություններ/Բոլոր գործողություններ/Թղթապանակներ/Պայմանագրի թղթապանակ"
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_Folders & "|" & c_AgrFolder)
    Name = "¶ñ³ýÇÏáí ûí»ñ¹ñ³ýïÇ å³ÛÙ³Ý³·Çñ"
    NameLen = 30
    ColNum = 0
    Pttel = "_2"
    Exists = Find_Doc_By(Name, NameLen,ColNum, Pttel)
  
    If Not Exists Then
      Call Log.Error("Գրաֆիկով օվերդրաֆտ պայմանագիրը չի գտնվել:")  
      Exit Sub
    End If    

    'Ջնջել "050218" ասմաթվով Օվերդրաֆտի մարումը
    Date = "050218"
    Typ = "22"
    Key = "0"
    Call DeleteD(Date, Typ, Key)
  
    'Ջնջել "050218" ասմաթվով Օվերդրաֆտի մարումը
    Date = "040218"
    Typ = "51"
    Key = "0"
    Call DeleteD(Date, Typ, Key)
  
    'Ջնջել "050218" ասմաթվով Օվերդրաֆտի մարումը
    Date = "030218"
    Typ = "22"
    Key = "0"
    Call DeleteD(Date, Typ, Key)
  
    'Ջնջել "050218" ասմաթվով Օվերդրաֆտի մարումը
    Date = "020218"
    Typ = "51"
    Key = "0"
    Call DeleteD(Date, Typ, Key)
    
    wMDIClient.VBObject("frmPttel_2").Close
  
    wMDIClient.VBObject("frmPttel").Close
  
    'Ջնջել "Կանխիկ ելք"-ը
    'Մուտք գործել "Հաճախորդի սպասարկում և դրամարկղ / Հաշվառված վճարային փաստաթղթեր"
    Call ChangeWorkspace(c_CustomerService)
    Call wTreeView.DblClickItem("|Ð³×³Ëáñ¹Ç ëå³ë³ñÏáõÙ ¨ ¹ñ³Ù³ñÏÕ |Ð³ßí³éí³Í í×³ñ³ÛÇÝ ÷³ëï³ÃÕÃ»ñ")
  
    Call Rekvizit_Fill("Dialog", 1, "General", "PERN", "^A[Del]" & DateS) 
    Call Rekvizit_Fill("Dialog", 1, "General", "PERK", "^A[Del]" & DateF) 
    Call Rekvizit_Fill("Dialog", 1, "General", "DOCTYPE", DocType) 
    Call Rekvizit_Fill("Dialog", 1, "General", "CLICODE", .Client ) 
    Call ClickCmdButton(2, "Î³ï³ñ»É")

    Name = "úí»ñ¹ñ³ýïÇ ïñ³Ù³¹ñáõÙ"
    NameLen = 21
    ColNum = 9
    Pttel = ""
    Exists = Find_Doc_By(Name, NameLen,ColNum, Pttel)
    
    If Not Exists Then
      Call Log.Error("Կանխիկ ելք փաստաթուղթը չի գտնվել:")  
      Exit Sub
    End If
 
    Call wMainForm.MainMenu.Click(c_Opers & "|" & c_Delete)
    Call ClickCmdButton(3, "²Ûá")
    Call ClickCmdButton(5, "²Ûá")

    wMDIClient.VBObject("frmPttel").Close

    'Պայմանագրի թղթապանակից Ջնջել "Գումարի տրամադրում...", "Գրաֆիկոց օվերդրաֆտ" պայմանագրերը 
    Call ChangeWorkspace(c_Overdraft)
    'Մուտք գործել "Օվերդրաֆտ (տեղաբաշխված)/Օվերդրաֆտ ունեցող հաշիվներ"
    FolderName = "|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|"
    Call wTreeView.DblClickItem(FolderName & "úí»ñ¹ñ³ýï áõÝ»óáÕ Ñ³ßÇíÝ»ñ")
    Call Rekvizit_Fill("Dialog", 1, "General", "ACCMASK", .CalcAcc) 
    Call ClickCmdButton(2, "Î³ï³ñ»É")

    '"Գործողություններ/Բոլոր գործողություններ/Թղթապանակներ/Պայմանագրի թղթապանակ"
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_Folders & "|" & c_AgrFolder)
   
    '''''Պայմանագրի թղթապանակից Ջնջել "Գումարի տրամադրում..."-ը 
  
    Name = "¶áõÙ³ñÇ ïñ³Ù³¹ñáõÙ"
    NameLen = 18
    ColNum = 0
    Pttel = "_2"
    Exists = Find_Doc_By(Name, NameLen,ColNum, Pttel)
    If Not Exists Then
      Call Log.Error("Գումարի տրամադրում փաստաթուղթը չի գտնվել:")  
      Exit Sub
    End If
  
    Call wMainForm.MainMenu.Click(c_Opers & "|" & c_Delete)
    Call ClickCmdButton(3, "²Ûá")

    wMDIClient.VBObject("frmPttel_2").Close
  
    ''Ջնջել Գրաֆիկով օվերդրաֆտ պայմանագիրը
    If wMDIClient.VBObject("frmPttel").VBObject("tdbgView").ApproxCount = 1 Then
      Call wMainForm.MainMenu.Click(c_Opers & "|" & c_Delete)
      Call ClickCmdButton(3, "²Ûá")
    End If  
    wMDIClient.VBObject("frmPttel").Close

  '---------------------------------------------------------------------------------  
  End With   
  Call Close_AsBank()  
End Sub