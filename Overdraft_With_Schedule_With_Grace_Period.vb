Option Explicit
'USEUNIT Library_Common  
'USEUNIT Subsystems_SQL_Library 
'USEUNIT Constants
'USEUNIT Overdraft_NewCases_Library
'USEUNIT Akreditiv_Library
'USEUNIT Library_CheckDB

'Test Case Id 165846

Sub Overdraft_With_Schedule_With_Grace_Period_Test()
  Dim fDATE, sDATE, my_vbObj
  Dim fBASE, RepaySchedule_ISN, GiveOverdradt_ISN, CalcDoc_ISN, RepayDoc_isn
  Dim OverdraftType, CreditCard, AutoDebt, PayerCode,TemplateType,CurCode,CalcAcc, Sum, Renewable, Limit, aDate, pDate, tDate, RepayBy, _
      DateType,SumsDateType,PCAGR, Baj, SumsFillType, PaymentDate, BranchCode, UsageField, Aim, Schedule,_
      Guarant, Country, District, RegionLR, Number,  sDOCNUM, PledgeCode, PledgeCur, PledgeValue, PledgeCount
  Dim str, AdditionalDays, Direction
  Dim name, name_len, ColNum, Pttel, IfExists, Data, docType, FolderName
  Dim GiveDate, Money, oType, Num, accNum, accCredit
  Dim opDate, PerSum, opPerSum, opUnUsePerSum , expectedMoney, opSum
  Dim DateS, DateF, date, typ, Key, attr
  Dim sql_Value, sql_isEqual, queryString,dbFOLDERS(2)
  
  ''1, Համակարգ մուտք գործել ARMSOFT օգտագործողով
  fDATE = "20260101"
  sDATE = "20140101"
  Call Initialize_AsBank("bank", sDATE, fDATE)
  Login("ARMSOFT")
  Call Create_Connection()
  
 '--------------------------------------
  Set attr = Log.CreateNewAttributes
  attr.BackColor = RGB(0, 255, 255)
  attr.Bold = True
  attr.Italic = True
'--------------------------------------  
  
  ''2, Անցում "Օվերդրաֆտ (տեղաբաշխված)" ԱՇՏ
  Call ChangeWorkspace(c_Overdraft)
  
'------------------------------------------------------------------------------  
  ''25. Ջնջել բոլոր փաստաթղթերը
  'Մուտք գործել "Օվերդրաֆտ (տեղաբաշխված)/Օվերդրաֆտ ունեցող հաշիվներ"
  calcAcc = "00001850100"
  DateS = "010117" 
  DateF = "010119"
  payerCode  = "00000018"
  
  FolderName = "|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|"
  Call wTreeView.DblClickItem(FolderName & "úí»ñ¹ñ³ýï áõÝ»óáÕ Ñ³ßÇíÝ»ñ")
  Call Rekvizit_Fill("Dialog", 1, "General", "ACCMASK", calcAcc) 
  Call ClickCmdButton(2, "Î³ï³ñ»É")
  
  '"Գործողություններ/Բոլոր գործողություններ/Թղթապանակներ/Պայմանագրի թղթապանակ"
  If Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").VBObject("tdbgView").ApproxCount <> 0 Then
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_Folders & "|" & c_AgrFolder)
  
    'Ջնջել "Գրավի վերադարձը" և "Գրավի տրամադրումը"
    name = "¶ñ³íÇ å³ÛÙ³Ý³·Çñ"
    name_len = 16
    ColNum = 0
    Pttel = "_2"
    IfExists = Find_Doc_By(name, name_len, ColNum, Pttel)
    If IfExists Then 
      Call wMainForm.MainMenu.Click(c_AllActions)
      Call wMainForm.PopupMenu.Click(c_OpersView)
  
      Sys.Process("Asbank").VBObject("frmAsUstPar").VBObject("TabFrame").VBObject("TDBDate").Keys("^A[Del]" & "[Tab]")
      Sys.Process("Asbank").VBObject("frmAsUstPar").VBObject("TabFrame").VBObject("TDBDate_2").Keys("^A[Del]" & "[Tab]")
      Sys.Process("Asbank").VBObject("frmAsUstPar").VBObject("CmdOK").ClickButton
      Set my_vbObj = Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel_3").VBObject("tdbgView")
  
      If my_vbObj.ApproxCount <> 0 Then
        my_vbObj.MoveLast
        Do While my_vbObj.ApproxCount <> 0  
          Call wMainForm.MainMenu.Click(c_AllActions)
          Call wMainForm.PopupMenu.Click(c_Delete)
          Sys.Process("Asbank").VBObject("frmDeleteDoc").VBObject("YesButton").ClickButton
        Loop
      End If
  
      Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel_3").Close
    
      'Ջնջել "Գրավի պայմանագիրը"
      Call wMainForm.MainMenu.Click(c_AllActions)
      Call wMainForm.PopupMenu.Click(c_Delete)
      Sys.Process("Asbank").VBObject("frmDeleteDoc").VBObject("YesButton").ClickButton
    End If
      
    name = "¶ñ³ýÇÏáí ûí»ñ¹ñ³ýïÇ å³ÛÙ³Ý³·Çñ"
    name_len = 30
    ColNum = 0
    Pttel = "_2"
    IfExists = Find_Doc_By(name, name_len, ColNum, Pttel)
    If IfExists Then
      'Ջնջել օվերդրաֆտի "Տոկոսի մարումը"
      date = "110418"
      typ = "53"
      Key = "2"
      Call DeleteD(date, typ, Key)
    
      'Ջնջել Օվերդրաֆտի "Տոկոսի հաշվարկումը"
      date = "100418"
      typ = "51"
      Key = "2"
      Call DeleteD(date, typ, Key)
  
      'Ջնջել  "Օվերդրաֆտի մարումը"
      date = "100418"
      typ = "22"
      Key = "2"
      Call DeleteD(date, typ, Key)
  
      'Ջնջել  "Ժամկետնանց գումարի ձևավորումը"
      date = "100418"
      typ = "g1"
      Key = "2"
      Call DeleteD(date, typ, Key)
  
      'Ջնջել "Արտոնյալ օվերդրաֆտի հաշվարկման ամսաթվերը"
      BuiltIn.delay(2000)
      Call wMainForm.MainMenu.Click(c_AllActions)
      Call wMainForm.PopupMenu.Click(c_ViewEdit & "|" & c_Other & "|" & c_OvGrPerCalcDates)
  
      Sys.Process("Asbank").VBObject("frmAsUstPar").VBObject("TabFrame").VBObject("TDBDate").Keys("^A[Del]" & "[Tab]")
      Sys.Process("Asbank").VBObject("frmAsUstPar").VBObject("TabFrame").VBObject("TDBDate_2").Keys("^A[Del]" & "[Tab]")
      Sys.Process("Asbank").VBObject("frmAsUstPar").VBObject("CmdOK").ClickButton
  
      Set my_vbObj = Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel_3").VBObject("tdbgView")

      If my_vbObj.ApproxCount <> 0 Then
        my_vbObj.MoveLast
        Do While my_vbObj.ApproxCount <> 0
          Call wMainForm.MainMenu.Click(c_AllActions)
          Call wMainForm.PopupMenu.Click(c_Delete)
          Sys.Process("Asbank").VBObject("frmDeleteDoc").VBObject("YesButton").ClickButton
        Loop
      End If
  
      Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel_3").Close
  
      'Ջնջել "Օվերդրաֆտի տրամադրումը"
      name = "¶ñ³ýÇÏáí ûí»ñ¹ñ³ýïÇ å³ÛÙ³Ý³·Çñ"
      name_len = 30
      ColNum = 0
      Pttel = "_2"
      IfExists = Find_Doc_By(name, name_len, ColNum, Pttel)
      If Not IfExists Then
        Call Log.Error("Փաստաթուղթը չի գտնվել") 
        Exit Sub
      End If
  
      'Ջնջել "Օվերդրաֆտի տրամադրումը"
      date = "040418"
      typ = "21"
      Key = "2"
      Call DeleteD(date, typ, Key)
  
      'Ջնջել "Գրաֆիկով օվերդրաֆտ պայամանգիրը"
      name = "¶ñ³ýÇÏáí ûí»ñ¹ñ³ýïÇ å³ÛÙ³Ý³·Çñ"
      name_len = 30
      ColNum = 0
      Pttel = "_2"
      IfExists = Find_Doc_By(name, name_len, ColNum, Pttel)
      If Not IfExists Then
        Call Log.Error("Փաստաթուղթը չի գտնվել") 
        Exit Sub
      End If
  
      Call wMainForm.MainMenu.Click(c_AllActions)
      Call wMainForm.PopupMenu.Click(c_Delete)
      
      Sys.Process("Asbank").VBObject("frmDeleteDoc").VBObject("YesButton").ClickButton
    End If
    
    Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel_2").Close
  End If
  
  Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").Close
'------------------------------------------------------------------------------
  
   Call Log.Message("Գրաֆիկով օվերդրաֆտի պայմանագիր (արտոնյալ ժամկետով) պայմանագրի ստեղծում",,,attr)
  'Կատարել "Նոր պայմանագրի ստեղծում"
  OverdraftType = "¶ñ³ýÇÏáí ûí»ñ¹ñ³ýïÇ å³ÛÙ³Ý³·Çñ (³ñïáÝÛ³É Å³ÙÏ»ïáí)"
  PayerCode  = "00000018"
  TemplateType = ""
  CurCode = "000"
  CalcAcc = "00001850100"
  Sum = "100000"
  Renewable = 1
  Limit = 0
  aDate = "040418"
  pDate = "040418"
  tDate = "040419"
  DateType= "1"
  sumsDateType = "1"
  PCAGR = "18"
  Baj = "365"
  SumsFillType = "01"
  PaymentDate = "5"
  BranchCode = "U2"
  Aim = "00"
  Schedule = "9"
  Guarant = "9"
  Country = "AM"
  District = "001"
  RegionLR = "010000008"
  Number = "333"
  PledgeCode = "00001"
  PledgeCur = "000"
  PledgeValue = "200000"
  PledgeCount= "1"
  Call Letter_Of_Overdraft_Doc_Fill(OverdraftType, CreditCard, PayerCode,TemplateType,CurCode,CalcAcc, Sum, Renewable,_
                                    Limit, aDate, pDate, tDate, RepayBy, DateType,SumsDateType,PCAGR,_
                                    Baj, SumsFillType, PaymentDate, BranchCode, UsageField, Aim, Schedule,_
                                    Guarant, Country, District, RegionLR, Number, fBASE, sDOCNUM, AutoDebt, PledgeCode, PledgeCur, PledgeValue, PledgeCount)
  
    ''SQL ստուգում պայամանգիր ստեղցելուց հետո: 
    ''CONTRACTS
    queryString = "select count(*) from CONTRACTS where fDGISN= " & fBASE &_
                    "and fDGAGRTYPE = 'C' and fDGMODTYPE = 3 and fDGSTATE = 206" &_
                    "and fDGSUMMA = 100000.00 and fDGALLSUMMA = 0.00"
    sql_Value = 1
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
      Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If  
                                
    ''FOLDERS
    queryString = "select count(*) from FOLDERS where fISN= " & fBASE
    sql_Value = 3
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
      Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If                                  
                                    
  ''4.Մարման գրաֆիկի նշանակում
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_RepaySchedule)     
  
  str = GetVBObject ("AUTODATEUN", wMDIClient.VBObject("frmASDocForm"))
  wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame").VBObject(str).Click
  
  AdditionalDays = "5"
  Sys.Process("Asbank").VBObject("frmAsUstPar").VBObject("TabFrame").VBObject("TDBNumber_5").Keys(AdditionalDays & "[Tab]")
  
  Direction = "2"
  Sys.Process("Asbank").VBObject("frmAsUstPar").VBObject("TabFrame").VBObject("ASTypeTree_4").VBObject("TDBMask").Keys(Direction & "[Tab]")
  
  Sys.Process("Asbank").VBObject("frmAsUstPar").VBObject("CmdOK").ClickButton
  Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmASDocForm").VBObject("CmdOk_2").ClickButton
  
  name = "Ø³ñáõÙÝ»ñÇ ·ñ³ýÇÏ"
  name_len = 17
  ColNum = 0
  Pttel = ""
  IfExists = Find_Doc_By(name, name_len, ColNum, Pttel)
  If IfExists Then 
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_View)
    RepaySchedule_ISN = Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmASDocForm").DocFormCommon.Doc.isn
    Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmASDocForm").Close
  End If
  
    ''SQL ստուգում Մարման գրաֆիկ ստեղցելուց հետո: 
    ''AGRSCHEDULE
    queryString = "select count(*) from AGRSCHEDULE where fBASE = " & RepaySchedule_ISN &_
                    "and fTYPE = 0 and fKIND = 9"
    sql_Value = 1
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
      Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If  
    
    ''AGRSCHEDULEVALUES
    queryString = "select count(*) from AGRSCHEDULEVALUES where fAGRISN = " & fBASE 
    sql_Value = 26
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
      Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    ''CONTRACTS
    queryString = "select fDGSTATE from CONTRACTS where fDGISN = " & fBASE &_
                     "and fDGAGRTYPE = 'C' and fDGMODTYPE = 3 and fDGSTATE = 1" &_
                     "and fDGSUMMA = 100000.00 and fDGALLSUMMA = 0.00"
    sql_Value = 1
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
      Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If 
       
    ''FOLDERS
    queryString = "select count(*) from FOLDERS where fISN = '" & RepaySchedule_ISN & "'"
    sql_Value = 1
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
      Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
   
  ''5.Այլ վճարումների գրաֆիկի նշանակում
  name = "¶ñ³ýÇÏáí ûí»ñ¹ñ³ýïÇ å³ÛÙ³Ý³·Çñ"
  name_len = 30
  ColNum = 0
  Pttel = ""
  IfExists = Find_Doc_By(name, name_len, ColNum, Pttel)
  If Not IfExists then
     call Log.Error("Փաստաթուղթը չի գտնվել") 
     exit Sub
   End If
     
  call ContractAction (c_OtherPaySchedule)
  Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmASDocForm").VBObject("CmdOk_2").ClickButton

   ''6."Գրաֆիկով օվերդրաֆտային պայմանագրի" համար կատարել "Գործողություններ/Բոլոր գործողություններ/Ուղարկել հաստատման " գործողությունը 
  'կանգնել պայմանագրի վրա
  name = "¶ñ³ýÇÏáí ûí»ñ¹ñ³ýïÇ å³ÛÙ³Ý³·Çñ"
  name_len = 30
  ColNum = 0
  Pttel = ""
  IfExists = Find_Doc_By(name, name_len, ColNum, Pttel)
  If Not IfExists then
     call Log.Error("Փաստաթուղթը չի գտնվել") 
     exit Sub
  End If
   
  call ContractAction(c_SendToVer)
  Sys.Process("Asbank").VBObject("frmAsMsgBox").VBObject("cmdButton").ClickButton 
  
  Builtin.Delay(2000)
  Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").Close
  
  ''7.Մուտք գործել Հաստատվող փաստաթղթեր 1 թղթապանակ - Պայմանագիրը պետք է առկա լինի
  Call wTreeView.DblClickItem("|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I")
  Sys.Process("Asbank").VBObject("frmAsUstPar").VBObject("TabFrame").VBObject("AsTpComment").VBObject("TDBComment").Keys(sDOCNUM & "[Tab]") 
  Sys.Process("Asbank").VBObject("frmAsUstPar").VBObject("CmdOK").ClickButton
  Builtin.Delay(1000)
  Set my_vbObj = Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").VBObject("tdbgView")
  If my_vbObj.ApproxCount <> 1 Then
    Call Log.Error("Պայմանագիրը առկա չէ Հաստատվող փաստաթղթեր 1 թղթապանակում:")
    Exit Sub
  End If
  
  ''8.Վավերացնել պայմանագիրը    
  Builtin.Delay(1000)
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_ToConfirm)
  Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmASDocForm").VBObject("CmdOk_2").ClickButton
  
  Builtin.Delay(1000)
  Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").Close
   
  ''9.Մուտք գործել "Պայմանագրեր" թղթապանակ - Պայմանագիրը պետք է առկա լինի:
  docType = "1"
  FolderName = "|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|"
  IfExists = LetterOfCredit_Filter_Fill(FolderName, docType, sDOCNUM)
  If (Not IfExists) Then
    call Log.Error("Պայմանագիրը առկա չէ")
    exit sub
  End If          
  
  '10.Ստուգել գրավի և պայմանագրի կապակցվածությունը:
  
  Call Log.Message("Օվերդրաֆտի տրամադրում",,,attr)
  GiveDate = aDate
  Money = "100000"
  oType = "2"
  Num = ""
  accNum = payerCode
  accCredit = ""
  GiveOverdradt_ISN = Give_Overdradt(GiveDate, Money, oType, Num, accNum, accCredit)
  
    ''SQL ստուգում Օվերդրաֆտ տրամադրելուց հետո:   
    BuiltIn.Delay(delay_small)
  
    ''AGRSCHEDULE
    queryString = "select count(*) from AGRSCHEDULE where fAGRISN = " & fBASE &_
                       " and (fINC = 1 or  fINC = 2) and (fKIND = 9 or fKIND = 3)" 
    sql_Value = 2
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
      Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If  
  
    BuiltIn.Delay(delay_small)
    
    ''AGRSCHEDULEVALUES
    queryString = "select count(*) from AGRSCHEDULEVALUES where fAGRISN = " & fBASE 
    sql_Value = 52
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
      Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If  
       
    BuiltIn.Delay(delay_small)
    
    ''CONTRACTS
    queryString = "select count(*) from CONTRACTS where fDGISN = " & fBASE &_
                      "and fDGAGRTYPE = 'C' and fDGMODTYPE = 3 and fDGSTATE = 7" &_
                      "and fDGSUMMA = 100000.00 and fDGALLSUMMA = 0.00"
    sql_Value = 1
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
      Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If 
    
    ''FOLDERS
    queryString = "select count(*) from FOLDERS where fISN = " & fBASE 
    sql_Value = 5
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
      Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If 
    
    Set dbFOLDERS(1) = New_DB_FOLDERS()
        dbFOLDERS(1).fFOLDERID = "LOANREGISTER"
        dbFOLDERS(1).fNAME = "C3Univer"
        dbFOLDERS(1).fKEY = fBASE 
        dbFOLDERS(1).fISN = fBASE 
        dbFOLDERS(1).fSTATUS = 1
        dbFOLDERS(1).fCOM = "ä»ïñáëÛ³Ý ä»ïñáë"
        dbFOLDERS(1).fSPEC = "C38"& Trim(sDOCNUM) &"          333                               0                                                                                                                                                             0.00                                                                                                                                                                                                                                                                                               "

    Set dbFOLDERS(2) = New_DB_FOLDERS()
        dbFOLDERS(2).fFOLDERID = "LOANREGISTER2"
        dbFOLDERS(2).fNAME = "C3Univer"
        dbFOLDERS(2).fKEY = fBASE 
        dbFOLDERS(2).fISN = fBASE 
        dbFOLDERS(2).fSTATUS = "1"
        dbFOLDERS(2).fCOM = "ä»ïñáëÛ³Ý ä»ïñáë"
        dbFOLDERS(2).fSPEC = "0"
        
    Call CheckDB_FOLDERS(dbFOLDERS(1), 1)
    Call CheckDB_FOLDERS(dbFOLDERS(2), 1)
    
    ''HI
    queryString = "select count(*) from HI where fBASE = " & GiveOverdradt_ISN &_ 
                      "and fSUM = 100000.00 and fCURSUM = 100000.00" &_
                      "and (fTYPE = 01 or fTYPE = 02)"
    sql_Value = 3
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
      Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    ''HIF
    queryString = "select count(*) from HIF where fBASE = " & fBASE 
    sql_Value = 20
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
      Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    ''HIR
    queryString = "select count(*) from HIR where fBASE= " & GiveOverdradt_ISN &_ 
                      "and fCURSUM = 100000.00"
    sql_Value = 1
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
      Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    ''HIRREST
    queryString = "select count(*) from HIRREST where fOBJECT = " & fBASE &_
                    "and fLASTREM = 100000.00 and fPENULTREM = 0.00" &_
                    "and fSTARTREM = 0.00"
    sql_Value = 1
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
      Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If

  
  Call Log.Message("Օվերդրաֆտի տոկոսների հաշվարկ",,,attr)
  Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").Close
  opDate = "040418"
  Call Overdraft_Percent_Calculation(CalcDoc_ISN, sDOCNUM, opDate, PerSum)
  
      ''SQL ստուգում Օվերդրաֆտի տոկոսների հաշվարկից հետո:      
      ''HIF
      queryString = "select count(*) from HIF where fBASE= " & CalcDoc_ISN &_
                       "and fSUM = 0.00 and fCURSUM = 0.00"
      sql_Value = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If  
  
  ''13.Ստուգել, որ "Տոկոսագումար" դաշտի արժեքը զրոյական լինի:
  If PerSum <> 0.00 Then
    Call Log.Error("Տոկոսագումար դաշտի արժեքը զրո չէ:")
  End If 
  
  Call Log.Message("Օվերդրաֆտի տոկոսների հաշվարկ",,,attr)
  Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").Close
  opDate = "050418"
  Call Overdraft_Percent_Calculation(CalcDoc_ISN, sDOCNUM, opDate, PerSum)
  
      ''SQL ստուգում Օվերդրաֆտի տոկոսների հաշվարկից հետո:      
      ''HIF
      queryString = "select count(*) from HIF where fBASE= " & CalcDoc_ISN &_
                       "and fSUM = 0.00 and fCURSUM = 0.00"
      sql_Value = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If
  
  ''15.Ստուգել, որ "Տոկոսագումար" դաշտի արժեքը զրոյական լինի:
  If PerSum <> 0.00 Then
    Call Log.Error("Տոկոսագումար դաշտի արժեքը զրո չէ:")
  End If      
  
  Call Log.Message("Օվերդրաֆտի տոկոսների հաշվարկ",,,attr)
  Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").Close
  opDate = "090418"
  Call Overdraft_Percent_Calculation(CalcDoc_ISN, sDOCNUM, opDate, PerSum)
  
      ''SQL ստուգում Օվերդրաֆտի տոկոսների հաշվարկից հետո:         
      ''HIF
      queryString = "select count(*) from HIF where fBASE = " & CalcDoc_ISN &_
                       "and fSUM = 0.00 and fCURSUM = 0.00"
      sql_Value = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If
    
      ''HIR
      queryString = "select count(*) from HIR where fBASE= " & CalcDoc_ISN &_
                        "and fCURSUM = 7692.30"
      sql_Value = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If
    
      ''HIRREST
      queryString = "select count(*) from HIRREST where fOBJECT = " & fBASE &_ 
                        "and fPENULTREM = 0.00 and fSTARTREM = 0.00 and (fLASTREM = 100000.00 or fLASTREM = 7692.30)"
      sql_Value = 2
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If

  
  ''17.Ստուգել, որ "Տոկոսագումար" դաշտի արժեքը զրոյական լինի:
  If PerSum <> 0.00 Then
    Call Log.Error("Տոկոսագումար դաշտի արժեքը զրո չէ:")
  End If 
  
  Call Log.Message("Օվերդրաֆտի պարտքերի մարում",,,attr)
  Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").Close
  opDate = "100418"
  opSum = ""
  opPerSum = ""
  opUnUsePerSum = ""
  RepayDoc_isn = Overdraft_Repayment_Operation(sDOCNUM, opDate, opSum, opPerSum, opUnUsePerSum)
  
    ''SQL ստուգում Օվերդրաֆտի պարտքերի մարումից հետո: 
    BuiltIn.Delay(delay_small)      
    ''HI
    queryString = "select count(*) from HI where fBASE= " & RepayDoc_isn &_
                      "and fSUM = 7692.30 and fCURSUM = 7692.30" &_
                      "and (fTYPE = 01 or fTYPE = 02)"
    sql_Value = 3
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
      Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If

    ''HIR
    queryString = "select count(*) from HIR where fBASE= " & RepayDoc_isn &_
                      "and fCURSUM = 7692.30"
    sql_Value = 2
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
      Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    ''HIRREST
    queryString = "select count(*) from HIRREST where fOBJECT = " & fBASE &_ 
                      "and (fPENULTREM = 0.00 or fPENULTREM = 100000.00)" &_
                      "and fSTARTREM = 0.00 and (fLASTREM = 0.00 or fLASTREM = 92307.70)"
    sql_Value = 2
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
      Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If

  ''19.Օվերդրաֆտի տոկոսների հաշվարկ
  Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").Close
  opDate = "100418"
  Call Overdraft_Percent_Calculation(CalcDoc_ISN, sDOCNUM, opDate, PerSum)
  
      ''SQL ստուգում Օվերդրաֆտի տոկոսների հաշվարկից հետո:    
      ''HI
      queryString = "select count(*) from HIF where fBASE = " & CalcDoc_ISN &_
                       "and fSUM = 0.00 and fCURSUM = 0.00"
      sql_Value = 2
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If
           
      ''HIF
      queryString = "select count(*) from HIF where fBASE = " & CalcDoc_ISN &_
                       "and fSUM = 0.00 and fCURSUM = 0.00"
      sql_Value = 2
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If
    
      ''HIR
      queryString = "select count(*) from HIR where fBASE= " & CalcDoc_ISN &_
                        "and fCURSUM = 318.60"
      sql_Value = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If
    
      ''HIRREST
      queryString = "select count(*) from HIRREST where fOBJECT = " & fBASE &_ 
                        "and (fLASTREM = 92307.70 or fLASTREM = 318.60 or fLASTREM = 0.00)" &_
                        "and (fPENULTREM = 0.00 or fPENULTREM = 100000.00)" &_
                        "and fSTARTREM = 0.00"
      sql_Value = 3
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If
  
  
  ''20.Ստուգել, որ "Տոկոսագումար" դաշտի արժեքը զրոյական լինի:
  If PerSum = 0.00 Then
    Call Log.Error("Տոկոսագումարը արժեքը զրո է:")
  End If          
  
  Call Log.Message("Օվերդրաֆտի պարտքերի մարում ժամկետից շուտ մարելով ամբողջ պարտքը",,,attr)
  Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").Close
  opDate = "110418"
  opSum = 92307.70
  opPerSum = 318.60
  opUnUsePerSum = ""
  RepayDoc_isn = Overdraft_Repayment_Operation(sDOCNUM, opDate, opSum, opPerSum, opUnUsePerSum)
  
    ''SQL ստուգում Օվերդրաֆտի պարտքերի մարումից հետո: 
    BuiltIn.Delay(delay_small)      
   ''AGRSCHEDULEVALUES
    queryString = "select count(*) from AGRSCHEDULEVALUES where fAGRISN = " & fBASE
    sql_Value = 82
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
      Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    BuiltIn.Delay(delay_small) 
         
    ''HI
    queryString = "select count(*) from HI where fBASE = " & RepayDoc_isn &_
                      "and (fSUM = 92307.70 or fSUM = 318.60)" &_
                      "and (fCURSUM = 92307.70 or fCURSUM = 318.60)"
    sql_Value = 5
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
      Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If

    ''HIR
    queryString = "select count(*) from HIR where fBASE = " & RepayDoc_isn &_
                      "and (fCURSUM = 92307.70 or fCURSUM = 318.60)"
    sql_Value = 2
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
      Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    ''HIRREST
    queryString = "select count(*) from HIRREST where fOBJECT = " & fBASE &_ 
                      "and (fLASTREM = 92307.70 or fLASTREM = 318.60 or fLASTREM = 0.00)" &_
                      "and fSTARTREM = 0.00 and fLASTREM = 0.00"
    sql_Value = 3
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
      Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If

  
  ''23.Ստուգել, որ մնացորդը զրոյացած լինի:
  Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").Close
  
  'Մուտք գործել "Պայմանագրեր" թղթապանակ
  docType = "1"
  FolderName = "|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|"
  IfExists = LetterOfCredit_Filter_Fill(FolderName, docType, sDOCNUM)
  If Not IfExists Then
    call Log.Error("Պայմանագիրը առկա չէ")
    exit sub
  End If 
  
  expectedMoney = "0.00"
  Data = Find_Data (expectedMoney, 3)
  If Not Data then
     call Log.Error("Հաշվի մնացորդը չի զրոյացել:") 
     exit Sub
  End If 
  
  Call Log.Message("Գրավի վերադարձ",,,attr)
  Builtin.Delay(1000)
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_Folders & "|" & c_AgrFolder)
  
  name = "¶ñ³íÇ å³ÛÙ³Ý³·Çñ"
  name_len = 16
  ColNum = 0
  Pttel = "_2"
  IfExists = Find_Doc_By(name, name_len, ColNum, Pttel)
  If Not IfExists Then 
    Call Log.Error("Գրավի պայմանագիրը առկա չէ:")
  End If
  
  opDate = "110418"
  Sum = ""
  Call ReturnPledge(opDate, Sum)
  
  Builtin.Delay(1000)
  Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel_2").Close
  Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").Close

'------------------------------------------------------------------------------  
  Call Log.Message("Ջնջել բոլոր փաստաթղթերը",,,attr)
  'Մուտք գործել "Օվերդրաֆտ (տեղաբաշխված)/Օվերդրաֆտ ունեցող հաշիվներ"
  calcAcc = "00001850100"
  DateS = "010117" 
  DateF = "010119"
  payerCode  = "00000018"
  
  FolderName = "|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|"
  Call wTreeView.DblClickItem(FolderName & "úí»ñ¹ñ³ýï áõÝ»óáÕ Ñ³ßÇíÝ»ñ")
  Call Rekvizit_Fill("Dialog", 1, "General", "ACCMASK", calcAcc) 
  Call ClickCmdButton(2, "Î³ï³ñ»É")
  
  '"Գործողություններ/Բոլոր գործողություններ/Թղթապանակներ/Պայմանագրի թղթապանակ"
  If Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").VBObject("tdbgView").ApproxCount <> 0 Then
   Builtin.Delay(1000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_Folders & "|" & c_AgrFolder)
  
    'Ջնջել "Գրավի վերադարձը" և "Գրավի տրամադրումը"
    name = "¶ñ³íÇ å³ÛÙ³Ý³·Çñ"
    name_len = 16
    ColNum = 0
    Pttel = "_2"
    IfExists = Find_Doc_By(name, name_len, ColNum, Pttel)
    If IfExists Then 
        Builtin.Delay(1000)
      Call wMainForm.MainMenu.Click(c_AllActions)
      Call wMainForm.PopupMenu.Click(c_OpersView)
  
      Sys.Process("Asbank").VBObject("frmAsUstPar").VBObject("TabFrame").VBObject("TDBDate").Keys("^A[Del]" & "[Tab]")
      Sys.Process("Asbank").VBObject("frmAsUstPar").VBObject("TabFrame").VBObject("TDBDate_2").Keys("^A[Del]" & "[Tab]")
      Sys.Process("Asbank").VBObject("frmAsUstPar").VBObject("CmdOK").ClickButton
      Set my_vbObj = Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel_3").VBObject("tdbgView")
  
      If my_vbObj.ApproxCount <> 0 Then
        my_vbObj.MoveLast
        Do While my_vbObj.ApproxCount <> 0  
          Builtin.Delay(1000)
          Call wMainForm.MainMenu.Click(c_AllActions)
          Call wMainForm.PopupMenu.Click(c_Delete)
          Sys.Process("Asbank").VBObject("frmDeleteDoc").VBObject("YesButton").ClickButton
        Loop
      End If
  
      Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel_3").Close
    
      'Ջնջել "Գրավի պայմանագիրը"
      Builtin.Delay(1000)
      Call wMainForm.MainMenu.Click(c_AllActions)
      Call wMainForm.PopupMenu.Click(c_Delete)
      Sys.Process("Asbank").VBObject("frmDeleteDoc").VBObject("YesButton").ClickButton
    End If
      
    name = "¶ñ³ýÇÏáí ûí»ñ¹ñ³ýïÇ å³ÛÙ³Ý³·Çñ"
    name_len = 30
    ColNum = 0
    Pttel = "_2"
    IfExists = Find_Doc_By(name, name_len, ColNum, Pttel)
    If IfExists Then
      'Ջնջել օվերդրաֆտի "Տոկոսի մարումը"
      date = "110418"
      typ = "53"
      Key = "2"
      Call DeleteD(date, typ, Key)
   
      'Ջնջել Օվերդրաֆտի "Տոկոսի հաշվարկումը"
      date = "100418"
      typ = "51"
      Key = "2"
      Call DeleteD(date, typ, Key)
  
      'Ջնջել  "Օվերդրաֆտի մարումը"
      date = "100418"
      typ = "22"
      Key = "2"
      Call DeleteD(date, typ, Key)
  
      'Ջնջել  "Ժամկետնանց գումարի ձևավորումը"
      date = "100418"
      typ = "g1"
      Key = "2"
      Call DeleteD(date, typ, Key)
  
      'Ջնջել "Արտոնյալ օվերդրաֆտի հաշվարկման ամսաթվերը"
      Builtin.Delay(1000)
      Call wMainForm.MainMenu.Click(c_AllActions)
      Call wMainForm.PopupMenu.Click(c_ViewEdit & "|" & c_Other & "|" & c_OvGrPerCalcDates)
  
      Sys.Process("Asbank").VBObject("frmAsUstPar").VBObject("TabFrame").VBObject("TDBDate").Keys("^A[Del]" & "[Tab]")
      Sys.Process("Asbank").VBObject("frmAsUstPar").VBObject("TabFrame").VBObject("TDBDate_2").Keys("^A[Del]" & "[Tab]")
      Sys.Process("Asbank").VBObject("frmAsUstPar").VBObject("CmdOK").ClickButton
  
      Set my_vbObj = Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel_3").VBObject("tdbgView")

      If my_vbObj.ApproxCount <> 0 Then
        my_vbObj.MoveLast
        Do While my_vbObj.ApproxCount <> 0
          Builtin.Delay(1000)
          Call wMainForm.MainMenu.Click(c_AllActions)
          Call wMainForm.PopupMenu.Click(c_Delete)
          Sys.Process("Asbank").VBObject("frmDeleteDoc").VBObject("YesButton").ClickButton
        Loop
      End If
  
      Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel_3").Close
  
      'Ջնջել "Օվերդրաֆտի տրամադրումը"
      name = "¶ñ³ýÇÏáí ûí»ñ¹ñ³ýïÇ å³ÛÙ³Ý³·Çñ"
      name_len = 30
      ColNum = 0
      Pttel = "_2"
      IfExists = Find_Doc_By(name, name_len, ColNum, Pttel)
      If Not IfExists Then
        Call Log.Error("Փաստաթուղթը չի գտնվել") 
        Exit Sub
      End If
  
      'Ջնջել "Օվերդրաֆտի տրամադրումը"
      date = "040418"
      typ = "21"
      Key = "2"
      Call DeleteD(date, typ, Key)
  
      'Ջնջել "Գրաֆիկով օվերդրաֆտ պայամանգիրը"
      name = "¶ñ³ýÇÏáí ûí»ñ¹ñ³ýïÇ å³ÛÙ³Ý³·Çñ"
      name_len = 30
      ColNum = 0
      Pttel = "_2"
      IfExists = Find_Doc_By(name, name_len, ColNum, Pttel)
      If Not IfExists Then
        Call Log.Error("Փաստաթուղթը չի գտնվել") 
        Exit Sub
      End If
  
      Builtin.Delay(1000)
      Call wMainForm.MainMenu.Click(c_AllActions)
      Call wMainForm.PopupMenu.Click(c_Delete)
      Sys.Process("Asbank").VBObject("frmDeleteDoc").VBObject("YesButton").ClickButton     
    End If
    Builtin.Delay(1000)
   Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel_2").Close
  End If
  
  Builtin.Delay(1000)
  Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").Close
'----------------------------------------------------------------------------------
  Call Close_AsBank()
End Sub