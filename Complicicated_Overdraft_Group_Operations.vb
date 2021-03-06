Option Explicit
'USEUNIT Library_Common  
'USEUNIT Subsystems_SQL_Library 
'USEUNIT Constants
'USEUNIT Overdraft_NewCases_Library
'USEUNIT Akreditiv_Library
'USEUNIT Group_Operations_Library
'USEUNIT Library_CheckDB

'Test case Id 165850

Sub Complicicated_Overdraft_Group_Operations_Test()
  Dim fDATE, sDATE, my_vbObj
  Dim queryString, sql_Value, sql_isEqual, colNum, name, name_len
  Dim SubAgr_ISN, MemOrd_ISN1, MemOrd_ISN2, MemOrd_ISN3, MemOrd_ISN4, RepaySchedule_ISN,_
      GroupCalc_ISN, GroupCalcSubAgr_ISN, GroupRepay_ISN
  Dim OverdraftType, CalcAcc1, Summa1, Limit, opDate1, Term1, RepayBy, DateType,_
      SumsDateType, SumsFillType, payerCode, temlateType, curCode, PaperCode1,_
      fBASE1, DocNum1, PledgeCode, PledgeCur, PledgeValue, PledgeCount
  Dim DocNum, fBASE, CreditCard, ClientCode, Curr, CalcAcc, Summa, Renewable, opDate, Term, _
      OverdraftPercent, NonUsedPercent, Baj, PastSum, PastPerSum, NonUsedPerSum, _
      DateFill, Paragraph, CheckPayDates, PayDates, Direction, AutoDebt, _
      UseOtherAccounts, Scheme, AutoDateChild, TypeAutoDate, AgrPeriod, DefineSchedule, _
      PerSumPayDate, StartDate, Sector, UsageField, Aim, Schedule, _
      Guarantee, Country, District, RegionLR, PaperCode
  Dim CalcDate, FormDate, DocType, FolderName, IfExists, Date, Typ, Client, Pttel,_
      MesBox, GiveDate, FirstDate, PayDatesCheck, SubAgrDocNum1, SubAgrDocNum2,_
      SubAgrDocNum3
  Dim CreditAcc, DebitAcc, OrderSum, LastDate, Workspace, Action, Count,_
      DelDocNum, SubDocCount, Operation
  Dim arrayDocNum, arrayCalcAcc, arrCheckbox
  Dim attr,dbFOLDERS(2)
  
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

  ''2, Անցում կատարել "Օվերդրաֆտ (տեղաբաշխված)" ԱՇՏ
  Call ChangeWorkspace(c_Overdraft)
 
  CreditCard = 1
  CalcAcc = "00001103022"      
  Summa = "100000"
  Renewable = 1
  opDate = "080518"
  Term = "080519"
  OverdraftPercent = 12
  NonUsedPercent = 12
  Baj = "365"
  PastSum = ""
  PastPerSum = ""
  NonUsedPerSum = ""
  DateFill = 1
  Paragraph = 1
  CheckPayDates = 0
  AutoDebt = 1
  AutoDateChild = 1
  TypeAutoDate = 2
  AgrPeriod = 12
  DefineSchedule = 1
  PerSumPayDate = 1
  PayDates = 15
  Direction = 2
  Sector = "U2"
  Aim = "00"
  Schedule = "9"
  Guarantee = "9"
  Country = "AM"
  RegionLR = "010000008"
  District = "001"
  PaperCode = "110"
  
  OverdraftType = "¶ñ³ýÇÏáí ûí»ñ¹ñ³ýïÇ å³ÛÙ³Ý³·Çñ"
  CreditCard = 1 
  CalcAcc1 = "33120090600"''' ???
  Summa1 = "200000"
  Limit = 0
  opDate1 = opDate
  Term1 = Term
  AutoDebt = 1 
  DateType = 1
  SumsDateType = 1
  SumsFillType = "01"
  UsageField = "01.001"
  PaperCode1 = "111"
'-------------------------------------------------------------------------------      
  ''Ջնջել բոլոր փաստաթղթերը
  'Մուտք գործել "Օվերդրաֆտ ունեցող հաշիվներ" թղթապանակ
  FolderName = "|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|"
  Call wTreeView.DblClickItem(FolderName & "úí»ñ¹ñ³ýï áõÝ»óáÕ Ñ³ßÇíÝ»ñ")
  Call Rekvizit_Fill("Dialog", 1, "General", "ACCMASK", CalcAcc)
  Call ClickCmdButton(2, "Î³ï³ñ»É")
  
  If wMDIClient.VBObject("frmPttel").VBObject("tdbgView").ApproxCount <> 0 Then
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_View)
    DocNum = wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame").VBObject("TextC").Text
    wMDIClient.VBObject("frmASDocForm").Close
    wMDIClient.VBObject("frmPttel").Close
    
    Workspace = "|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|"
    DocType = 2
    FirstDate = "091018"
    LastDate = "161018"
    Call GroupDelete(Workspace, DocType, DocNum, FirstDate, LastDate, Action)
  
    'Ջնջել ենթապայմանագրերը
    Workspace = "úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)"
    
    DocType = 1
    Call wTreeView.DblClickItem("|" & Workspace & "|ä³ÛÙ³Ý³·ñ»ñ")
    With Asbank.VBObject("frmAsUstPar")
    	.VBObject("TabFrame").VBObject("ASTypeTree").VBObject("TDBMask").Keys(DocType & "[Tab]")
      .VBObject("TabFrame").VBObject("AsTpComment").VBObject("TDBComment").Keys(DocNum & "[Tab]")
    	.VBObject("CmdOK").Click()
    End With
      
    SubDocCount = wMDIClient.VBObject("frmPttel").VBObject("tdbgView").ApproxCount
    wMDIClient.VBObject("frmPttel").Close
    Select Case SubDocCount
    Case 3     
      Call docDelete(Workspace, DocType, Trim(DocNum) & "_001")
      Call docDelete(Workspace, DocType, Trim(DocNum) & "_002")
    Case 2
      Call docDelete(Workspace, DocType, Trim(DocNum))
    End select  
      
    'Ջնջել օվերդրաֆտի տրամադրումը
    DocType = 2
    Call wTreeView.DblClickItem("|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|ä³ÛÙ³Ý³·ñ»ñ")
    Call Rekvizit_Fill("Dialog", 1, "General", "LEVEL", DocType) 
    Call Rekvizit_Fill("Dialog", 1, "General", "NUM", DocNum)
    Call ClickCmdButton(2, "Î³ï³ñ»É") 
    
    Date = "080518"
    Pttel = "_2"
    Typ = ""
    Call DeleteActions(Date, Pttel, Typ, MesBox)
   
    'Ջնջել մայր պայամնագիրը
    Call wMainForm.MainMenu.Click(c_Opers & "|" & c_Delete)
  
    Call ClickCmdButton(3, "²Ûá") 
  End If  
  wMDIClient.VBObject("frmPttel").Close
  
  'Ջնջել "Հիշարար օրդեր"-ները
  'Մուտք գործել "Հաճախորդի սպասարկում և դրամարկղ"
  Call ChangeWorkspace(c_CustomerService) 
  
  Date = "080518"
  Typ = "MemOrd"
  Client = "00000018"
  Call DeletePayDoc(Date, MemOrd_ISN1, Typ, Client)
   
  Date = "151018"
  Typ = "MemOrd"
  Client = "00000018"
  Call DeletePayDoc(Date, MemOrd_ISN2, Typ, Client)
  
  ''Ջնջել Գրաֆիկով օվերդրաֆտ պայամանագրի հետ կապված բոլոր փաստաթղթերը
  Call ChangeWorkspace(c_Overdraft)
  FolderName = "|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|"
  Call wTreeView.DblClickItem(FolderName & "úí»ñ¹ñ³ýï áõÝ»óáÕ Ñ³ßÇíÝ»ñ")
  Call Rekvizit_Fill("Dialog", 1, "General", "ACCMASK", CalcAcc1) 
  Call ClickCmdButton(2, "Î³ï³ñ»É")

  If wMDIClient.VBObject("frmPttel").VBObject("tdbgView").ApproxCount <> 0 Then
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_View)
    With wMDIClient
      DocNum1 = .VBObject("frmASDocForm").VBObject("TabFrame").VBObject("TextC").Text
      .VBObject("frmASDocForm").Close
      .VBObject("frmPttel").Close
    End With
    
    Call ChangeWorkspace(c_Overdraft)
    Workspace = "|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|"
    DocType = 2
    FirstDate = "^A[Del]"
    LastDate = "^A[Del]"
    Call GroupDelete(Workspace, DocType, DocNum1, FirstDate, LastDate, Action)
  
    'Ջնջել Գրաֆիկով օվերդրաֆտ պայամնագիրը
    DocType = 1
    Call wTreeView.DblClickItem("|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|ä³ÛÙ³Ý³·ñ»ñ")
    Call Rekvizit_Fill("Dialog", 1, "General", "LEVEL", DocType) 
    Call Rekvizit_Fill("Dialog", 1, "General", "NUM", DocNum1)
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    
    Call wMainForm.MainMenu.Click(c_Opers & "|" & c_Delete)
  
    Call ClickCmdButton(3, "²Ûá")
  End If
  wMDIClient.VBObject("frmPttel").Close
    
  'Ջնջել "Հիշարար օրդեր"-ները
  Call ChangeWorkspace(c_CustomerService) 
  Date = "080518"
  Typ = "MemOrd"
  Client = "00001003"
  Call DeletePayDoc(Date, MemOrd_ISN3, Typ, Client)
  
  Date = "151018"
  Typ = "MemOrd"
  Client = "00000233"
  Call DeletePayDoc(Date, MemOrd_ISN4, Typ, Client)
'-------------------------------------------------------------------------------  
  Call ChangeWorkspace(c_Overdraft)
  
  Call Log.Message("3.Բարդ օվերդրաֆտ(գծային) պայմանագրի ստեղծում",,,attr)
  Call Letter_Of_Complicicated_Overdraft_Doc_Fill(DocNum, fBASE, CreditCard, ClientCode, _
                           Curr, CalcAcc, Summa, Renewable, opDate, Term, _
                           OverdraftPercent, NonUsedPercent, Baj, PastSum, PastPerSum, _
                           NonUsedPerSum, DateFill, Paragraph, CheckPayDates, _
                           PayDates, Direction, AutoDebt, UseOtherAccounts, Scheme, AutoDateChild, _
                           TypeAutoDate, AgrPeriod, DefineSchedule, _
                              PerSumPayDate, StartDate, Sector, UsageField, Aim, _
                           Schedule, Guarantee, Country, District, RegionLR, PaperCode)
  Log.Message("Բարդ օվերդրաֆտ(գծային) պայմանագրի համարը` " & DocNum) 
                           
      ''SQL ստուգում պայամանգիր ստեղցելուց հետո: 
      ''CONTRACTS
      queryString = "select count(*) from CONTRACTS where fDGISN = " & fBASE &_
                      "and fDGAGRTYPE = 'C' and fDGMODTYPE = 3 and fDGAGRKIND = 'XL'" &_
                      "and fDGSTATE = 1 and fDGSUMMA = 100000.00 and fDGALLSUMMA = 0.00"
      sql_Value = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If  
                                
      ''FOLDERS
      queryString = "select count(*) from FOLDERS where fISN = " & fBASE 
      sql_Value = 3
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If                       
                                          
  ''4.Պայմանագիրը ուղարկել հաստատման
  Builtin.Delay(2000)
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_SendToVer)
  Call ClickCmdButton(5, "²Ûá")  
    
  Builtin.Delay(2000)
  wMDIClient.VBObject("frmPttel").Close
  
  ''5.Մուտք գործել "Հաստատվող փաստաթղթեր 1" թղթապանակ 
  Call wTreeView.DblClickItem("|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I")
  With Asbank.VBObject("frmAsUstPar")
    .VBObject("TabFrame").VBObject("AsTpComment").VBObject("TDBComment").Keys(DocNum & "[Tab]") 
    .VBObject("CmdOK").ClickButton
  End With  
  Builtin.Delay(2000)
  If wMDIClient.VBObject("frmPttel").VBObject("tdbgView").ApproxCount <> 1 Then
    Call Log.Error("Պայմանագիրը առկա չէ Հաստատվող փաստաթղթեր 1 թղթապանակում:")
    Exit Sub
  End If
  
  ''6.Վավերացնել պայմանագիրը
  Builtin.Delay(2000)
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_ToConfirm)
  
  wMDIClient.VBObject("frmASDocForm").VBObject("CmdOk_2").ClickButton

  Builtin.Delay(2000)
  wMDIClient.VBObject("frmPttel").Close
  
  Call Log.Message("7.Գրաֆիկօվ օվերդրաֆտ պայմանագրի ստեղծում",,,attr)   
  Call Letter_Of_Overdraft_Doc_Fill(OverdraftType, CreditCard, payerCode,temlateType,curCode,_
                                      CalcAcc1, Summa1, Renewable, Limit, opDate1, opDate1, Term1, RepayBy,_
                                      DateType, SumsDateType, OverdraftPercent, Baj, SumsFillType,_
                                      PayDates, Sector, UsageField, Aim, Schedule,Guarantee, Country, District, RegionLR, _
                                      PaperCode1, fBASE1, DocNum1, AutoDebt, PledgeCode, PledgeCur, PledgeValue, PledgeCount)
  Log.Message("Գրաֆիկօվ օվերդրաֆտ պայմանագրի համարը` " & DocNum1)                
  ''SQL ստուգում պայամանգիր ստեղցելուց հետո: 
      ''CONTRACTS
      queryString = "select count(*) from CONTRACTS where fDGISN = " & fBASE1 &_
                      "and fDGAGRTYPE = 'C' and fDGMODTYPE = 3 and fDGAGRKIND = '8L'" &_
                      "and fDGSTATE = 206 and fDGSUMMA = 200000.00 and fDGALLSUMMA = 0.00"
      sql_Value = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If  
                                
      ''FOLDERS
      queryString = "select count(*) from FOLDERS where fISN = " & fBASE1 
      sql_Value = 3
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If                          
                                     
  ''8.Մարման գրաֆիկի նշանակում
  Builtin.Delay(2000)
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_RepaySchedule)  
    
  name = "Ø³ñáõÙÝ»ñÇ ·ñ³ýÇÏ"
  name_len = 17
  ColNum = 0
  Pttel = ""
  IfExists = Find_Doc_By(name, name_len,ColNum, Pttel)
  If IfExists Then 
    Builtin.Delay(2000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_View)
    RepaySchedule_ISN = wMDIClient.VBObject("frmASDocForm").DocFormCommon.Doc.isn
    Builtin.Delay(2000)
    wMDIClient.VBObject("frmASDocForm").Close
  End If
  
  ''SQL ստուգում Մարման գրաֆիկ ստեղցելուց հետո: 
      ''AGRSCHEDULE
      queryString = "select count(*) from AGRSCHEDULE where fBASE = " & RepaySchedule_ISN &_
                      "and fKIND = 9 and fTYPE = 0 and fINC = 1"
      sql_Value = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If  
    
      ''CONTRACTS
      queryString = "select count(*) from CONTRACTS where fDGISN = " & fBASE1 &_
                      "and fDGAGRTYPE = 'C' and fDGMODTYPE = 3 and fDGAGRKIND = '8L'" &_
                      "and fDGSTATE = 1 and fDGSUMMA = 200000.00 and fDGALLSUMMA = 0.00"
      sql_Value = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If    
       
      ''FOLDERS
      queryString = "select count(*) from FOLDERS where fISN= '" & RepaySchedule_ISN & "'"
      sql_Value = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If
                  
  ''9.Գրաֆիկով օվերդրաֆտի պայմանագիրը ուղարկել հաստատման:
  name = "¶ñ³ýÇÏáí ûí»ñ¹ñ³ýïÇ å³ÛÙ³Ý³·Çñ"
  name_len = 30
  ColNum = 0
  Pttel = ""
  IfExists = Find_Doc_By(name, name_len,ColNum, Pttel)
  If Not IfExists then
     Call Log.Error("Գրաֆիկով օվերդրաֆտի պայմանագիրը փաստաթուղթը չի գտնվել") 
     Exit Sub
   End If
   
   Builtin.Delay(2000)
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_SendToVer)
  Call ClickCmdButton(5, "²Ûá")   
    
  Builtin.Delay(2000)
  wMDIClient.VBObject("frmPttel").Close
  
  ''10.Մուտք գործել "Հաստատվող փաստաթղթեր 1" թղթապանակ - Գրաֆիկով օվերդրաֆտ պայամանգիրը պետք է առկա լինի:
  Call wTreeView.DblClickItem("|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I")
  Call Rekvizit_Fill("Dialog", 1, "General", "NUM", DocNum1) 
  Call ClickCmdButton(2, "Î³ï³ñ»É")
  Builtin.Delay(2000)
  Set my_vbObj = wMDIClient.VBObject("frmPttel").VBObject("tdbgView")
  If my_vbObj.ApproxCount <> 1 Then
    Call Log.Error("Պայմանագիրը առկա չէ Հաստատվող փաստաթղթեր 1 թղթապանակում:")
    Exit Sub
  End If
  
  ''11.Վավերացնել պայմանագիրը
  Builtin.Delay(2000)
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_ToConfirm)
  Call ClickCmdButton(1, "Ð³ëï³ï»É")
  Builtin.Delay(2000)
  wMDIClient.VBObject("frmPttel").Close
   
  ''12.Մուտք գործել "Պայմանագրեր" թղթապանակ - Փաստաթուղթը պետք է առկա լինի:
  DocType = "2"
  FolderName = "|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|"
  IfExists = LetterOfCredit_Filter_Fill(FolderName, DocType, DocNum)
  If Not IfExists Then
    Call Log.Error("Պայմանագիրը առկա չէ Պայմանագրեր թղթապանակում")
    Exit Sub
  End If
  
  Call Log.Message("13.Ենթապայամանագրի բացում",,,attr)
  Summa = "10000"
  FirstDate = "080718"
  PayDatesCheck = 1    
  PayDates = 15  
  Direction = 0                 
  Call OpenSubagreement(SubAgrDocNum1, SubAgr_ISN, Summa, opDate, opDate, DateFill, FirstDate, Paragraph, PayDatesCheck, PayDates, Direction)
  
  wMDIClient.VBObject("frmPttel").Close
  
      ''SQL ստուգում Ենթապայամանագրի բացումից հետո
      ''CONTRACTS
      queryString = "select count(*) from CONTRACTS where fDGISN = " & SubAgr_ISN &_
                      "and fDGAGRTYPE = 'C' and fDGMODTYPE = 3 and fDGAGRKIND = 2 " &_
                      "and fDGSTATE = 1 and fDGSUMMA = 10000.00 and fDGALLSUMMA = 0.00"
      sql_Value = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If  
                                
      ''FOLDERS
      queryString = "select count(*) from FOLDERS where fISN = " & SubAgr_ISN 
      sql_Value = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If     
      
      ''HI
      queryString = "select count(*) from HI where fBASE= " & fBASE &_
                     "and fSUM = 100000.00 and fCURSUM = 100000.00"
      sql_Value = 2
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If
      
      ''HIF
      queryString = "select count(*) from HIF where fBASE = " & fBASE 
      sql_Value = 28
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If
      
  ''14.Մուտք գործել "Աշխատանքային փաստաթղթեր" թղթապանակ - Ենթապայմանագիրը պետք է առկա լինի:
  Call wTreeView.DblClickItem("|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ")
  Call Rekvizit_Fill("Dialog", 1, "General", "NUM", DocNum) 
  Call ClickCmdButton(2, "Î³ï³ñ»É")
  Builtin.Delay(2000)
  Set my_vbObj = wMDIClient.VBObject("frmPttel").VBObject("tdbgView")
  If my_vbObj.ApproxCount = 0 Then
    Call Log.Error("Ենթապայմանագիրը առկա չէ Աշխատանքային փաստաթղթեր թղթապանակում:")
    Exit Sub
  End If
  
  ''15.Ենթապայամանագիրը ուղարկել հաստատման
  Builtin.Delay(2000)
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_SendToVer)
  Call ClickCmdButton(5, "²Ûá")   
  Builtin.Delay(2000)
  wMDIClient.VBObject("frmPttel").Close
  
  ''16.Մուտք գործել "Հաստատվող փաստաթղթեր 1" թղթապանակ 
  Call wTreeView.DblClickItem("|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I")
  Sys.Process("Asbank").VBObject("frmAsUstPar").VBObject("TabFrame").VBObject("AsTpComment").VBObject("TDBComment").Keys(DocNum & "[Tab]") 
  Sys.Process("Asbank").VBObject("frmAsUstPar").VBObject("CmdOK").ClickButton
  Set my_vbObj = Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").VBObject("tdbgView")
  If my_vbObj.ApproxCount = 0 Then
    Call Log.Error("Պայմանագիրը առկա չէ Հաստատվող փաստաթղթեր 1 թղթապանակում:")
    Exit Sub
  End If
  
  ''17.Վավերացնել պայմանագիրը
  Builtin.Delay(2000)
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_ToConfirm)
  Call ClickCmdButton(1, "Ð³ëï³ï»É")
  Builtin.Delay(2000)
  wMDIClient.VBObject("frmPttel").Close
  
  ''18.Մուտք գործել "Հաճախորդի սպասարկում և դրամարկղ"
  Call ChangeWorkspace(c_CustomerService)
 
  Call Log.Message("19.Ստեղծել 'Հիշարար օրդեր' (Հաշիվ դեբետ = 00001103022)",,,attr)
  CreditAcc = "00001850100"
  OrderSum = "10000"
  MemOrd_ISN1 = Mem_Order_Create_Order(opDate, CalcAcc, CreditAcc, OrderSum)
  
  Builtin.Delay(2000)
  wMDIClient.VBObject("FrmSpr").Close
  
  ''20.Հաշվառել "Հիշարար օրդեր" փաստաթուղթը
  Builtin.Delay(2000)
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_DoTrans)
  Call ClickCmdButton(5, "²Ûá")
  Builtin.Delay(2000)
  wMDIClient.VBObject("frmPttel").Close
 
  Call Log.Message("21.Ստեղծել 'Հիշարար օրդեր' (Հաշիվ դեբետ = 33120090600)",,,attr)
  CreditAcc = "64110069600"
  OrderSum = "20000"
  MemOrd_ISN3 = Mem_Order_Create_Order(opDate, CalcAcc1, CreditAcc, OrderSum)
  
  Builtin.Delay(2000)
  wMDIClient.VBObject("FrmSpr").Close
  
  ''22.Հաշվառել "Հիշարար օրդեր" փաստաթուղթը
  Builtin.Delay(2000)
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_DoTrans)
  Call ClickCmdButton(5, "²Ûá")
  Builtin.Delay(2000)
  wMDIClient.VBObject("frmPttel").Close
 
  ''23.Մուտք գործել "Օվերդրաֆտ(տեղաբաշխված)" 
  Call ChangeWorkspace(c_Overdraft)
  
  ''24.Մուտք գործել "Պայմանագրեր" թղթապանակ - Փաստաթուղթը պետք է առկա լինի:
  DocType = "2"
  FolderName = "|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|"
  IfExists = LetterOfCredit_Filter_Fill(FolderName, DocType, DocNum)
  If Not IfExists Then
    Call Log.Error("Պայմանագիրը առկա չէ Պայմանագրեր թղթապանակում")
    Exit Sub
  End If
  
  Call Log.Message("25.Ենթապայամանագրի բացում",,,attr)
  Summa = "5000"
  Date = "100518"
  GiveDate = "151018"
  FirstDate = GiveDate
  PayDatesCheck = 0    
  Paragraph = 2  
  Direction = 0                 
  Call OpenSubagreement(SubAgrDocNum2, SubAgr_ISN, Summa, Date, GiveDate, DateFill, FirstDate, Paragraph, PayDatesCheck, PayDates, Direction)
  
  Builtin.Delay(2000)
  wMDIClient.VBObject("frmPttel").Close
  
    ''SQL ստուգում Ենթապայամանագրի բացումից հետո
      ''FOLDERS
      queryString = "select count(*) from FOLDERS where fISN = " & SubAgr_ISN 
      sql_Value = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If     
  
  ''26.Մուտք գործել "Աշխատանքային փաստաթղթեր" թղթապանակ - Ենթապայմանագիրը պետք է առկա լինի:
  Call wTreeView.DblClickItem("|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ")
  Call Rekvizit_Fill("Dialog", 1, "General", "NUM", DocNum) 
  Call ClickCmdButton(2, "Î³ï³ñ»É")
  Builtin.Delay(2000)
  Set my_vbObj = wMDIClient.VBObject("frmPttel").VBObject("tdbgView")
  If my_vbObj.ApproxCount = 0 Then
    Call Log.Error("Ենթապայմանագիրը առկա չէ Աշխատանքային փաստաթղթեր թղթապանակում:")
    Exit Sub
  End If
  
  ''27.Ենթապայամանագիրը ուղարկել հաստատման
  Builtin.Delay(2000)
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_SendToVer)
  Call ClickCmdButton(5, "²Ûá")  
  Builtin.Delay(2000)
  wMDIClient.VBObject("frmPttel").Close
  
  ''28.Մուտք գործել "Հաստատվող փաստաթղթեր 1" թղթապանակ 
  Call wTreeView.DblClickItem("|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I")
  Call Rekvizit_Fill("Dialog", 1, "General", "NUM", DocNum) 
  Call ClickCmdButton(2, "Î³ï³ñ»É")
  Builtin.Delay(2000)
  Set my_vbObj = wMDIClient.VBObject("frmPttel").VBObject("tdbgView")
  If my_vbObj.ApproxCount = 0 Then
    Call Log.Error("Պայմանագիրը առկա չէ Հաստատվող փաստաթղթեր 1 թղթապանակում:")
    Exit Sub
  End If
  
  ''29.Վավերացնել ենթապայմանագիրը
  Builtin.Delay(2000)
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_ToConfirm)
  Call ClickCmdButton(1, "Ð³ëï³ï»É")
  Builtin.Delay(2000)
  wMDIClient.VBObject("frmPttel").Close
  
  Call Log.Message("30.Օվերդրաֆտի խմբային տրամադրում",,,attr) 
  ReDim arrayCalcAcc(1)
  arrayCalcAcc(0) = CalcAcc
  arrayCalcAcc(1) = CalcAcc1
  Count = 2
  Operation = "Give"
  Call OverdraftGroupOperation(arrayCalcAcc, Count, opDate, Operation)
     
    ''SQL ստուգում Օվերդրաֆտ տրամադրելուց հետո:  
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
          dbFOLDERS(1).fNAME = "C3Compl"
          dbFOLDERS(1).fKEY = fBASE 
          dbFOLDERS(1).fISN = fBASE 
          dbFOLDERS(1).fSTATUS = 1
          dbFOLDERS(1).fCOM = "ä»ïñáëÛ³Ý ä»ïñáë"
          dbFOLDERS(1).fSPEC = "C3X"& Trim(DocNum) &"          110                               0                                                                                                                                                             0.00                                                                                                                                                                                                                                                                                               "

      Set dbFOLDERS(2) = New_DB_FOLDERS()
          dbFOLDERS(2).fFOLDERID = "LOANREGISTER2"
          dbFOLDERS(2).fNAME = "C3Compl"
          dbFOLDERS(2).fKEY = fBASE 
          dbFOLDERS(2).fISN = fBASE 
          dbFOLDERS(2).fSTATUS = "1"
          dbFOLDERS(2).fCOM = "ä»ïñáëÛ³Ý ä»ïñáë"
          dbFOLDERS(2).fSPEC = "0"
        
      Call CheckDB_FOLDERS(dbFOLDERS(1), 1)
      Call CheckDB_FOLDERS(dbFOLDERS(2), 1)
  
  ''31.Մուտք գործել "Հաճախորդի սպասարկում և դրամարկղ"
  Call ChangeWorkspace(c_CustomerService)
  
  Call Log.Message("32.Ստեղծել 'Հիշարար օրդեր' (Հաշիվ կրեդիտ = 00001103022)",,,attr)
  DebitAcc = "00001850100"
  OrderSum = "20000"
  opDate = "151018"
  MemOrd_ISN2 = Mem_Order_Create_Order(opDate, DebitAcc, CalcAcc, OrderSum)
  
  Builtin.Delay(2000)
  wMDIClient.VBObject("FrmSpr").Close
  
  ''33.Հաշվառել "Հիշարար օրդեր" փաստաթուղթը
  Builtin.Delay(2000)
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_DoTrans)
  Call ClickCmdButton(5, "²Ûá")  
  Builtin.Delay(2000)
  wMDIClient.VBObject("frmPttel").Close
  
  Call Log.Message("32.Ստեղծել 'Հիշարար օրդեր' (Հաշիվ կրեդիտ = 33120090600)",,,attr)
  DebitAcc = "64110069600"
  OrderSum = "30000"
  opDate = "151018"
  MemOrd_ISN4 = Mem_Order_Create_Order(opDate, DebitAcc, CalcAcc1, OrderSum)
  
  Builtin.Delay(2000)
  wMDIClient.VBObject("FrmSpr").Close
  
  ''33.Հաշվառել "Հիշարար օրդեր" փաստաթուղթը
  Builtin.Delay(2000)
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_DoTrans)
  Call ClickCmdButton(5, "²Ûá")  
  Builtin.Delay(2000)
  wMDIClient.VBObject("frmPttel").Close
  
  ''34.Մուտք գերծել "Օվերդրաֆտ(տեղաբաշխված)"
  Call ChangeWorkspace(c_Overdraft)
  
  Call Log.Message("35.Խմբային տոկոսների հաշվարկ Մայր պայմանագրերի համար",,,attr)
  ReDim arrayDocNum(1)
  arrayDocNum(0) = DocNum
  arrayDocNum(1) = DocNum1
  DocType = "2"
  CalcDate = "151018"
  FormDate = "151018" 
  GroupCalc_ISN = OverdraftGroupCalculation(arrayDocNum, Count, DocType, CalcDate, FormDate)
   
  Call Log.Message("36.Խմբային տոկոսների հաշվարկ ենթապայմանագրերի համար",,,attr)
  DocType = "1" 
  Call wTreeView.DblClickItem("|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|ä³ÛÙ³Ý³·ñ»ñ")
  Call Rekvizit_Fill("Dialog", 1, "General", "LEVEL", DocType) 
  Call Rekvizit_Fill("Dialog", 1, "General", "NUM", DocNum)
	Call ClickCmdButton(2, "Î³ï³ñ»É")
  BuiltIn.Delay(2000)
  
  opDate = "151018"
  ReDim arrCheckbox(2)          
  arrCheckbox = Array("CHG", "OPX")
  Call Group_Calculation(opDate, arrCheckbox)
  
  'Վերցնել Խմբային հաշվարկի ISN-ը
  BuiltIn.Delay(2000)
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_OpersView)
  Call Rekvizit_Fill("Dialog", 1, "General", "START", FirstDate)
  Call Rekvizit_Fill("Dialog", 1, "General", "END", LastDate)
  Call ClickCmdButton(2, "Î³ï³ñ»É")
  
  BuiltIn.Delay(2000)
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_View)
  GroupCalcSubAgr_ISN = wMDIClient.VBObject("frmASDocForm").DocFormCommon.Doc.isn
  wMDIClient.VBObject("frmASDocForm").Close
  Builtin.Delay(2000)
  wMDIClient.VBObject("frmPttel_2").Close
  Builtin.Delay(2000)
  wMDIClient.VBObject("frmPttel").Close
      ''SQL ստուգում Խմբային տոկոսների հաշվարկից հետո
      ''HI
      queryString = "select count(*) from HI where fBASE = " & GroupCalc_ISN
      sql_Value = 92
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If
      
      ''HI
      queryString = "select count(*) from HI where fBASE = " & GroupCalcSubAgr_ISN
      sql_Value = 66
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If
      
      ''HIF
      queryString = "select count(*) from HIF where fBASE = " & GroupCalc_ISN 
      sql_Value = 60
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If
      
      ''HIF
      queryString = "select count(*) from HIF where fBASE = " & GroupCalcSubAgr_ISN 
      sql_Value = 31
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If
      
      ''HIR
      queryString = "select count(*) from HIR where fBASE = " & GroupCalc_ISN 
      sql_Value = 51
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If
      
      ''HIRREST
      queryString = "select count(*) from HIRREST where fOBJECT = " & fBASE1
      sql_Value = 5
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If
      
      ''HIT
      queryString = "select count(*) from HIT where fOBJECT = " & fBASE1
      sql_Value = 23
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If
      
      ''HIRREST
      queryString = "select count(*) from HIRREST where fOBJECT = " & fBASE &_
                     "and ((fLASTREM = 4763.90 and fPENULTREM = 0.00) or (fLASTREM = 4527.20 and fPENULTREM = 3698.70))"
      sql_Value = 2
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If
    
      ''HIT
      queryString = "select count(*) from HIT where fOBJECT = " & fBASE 
      sql_Value = 11
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If  

  Call Log.Message("37.Խմբային մարում",,,attr)
  opDate = "161018"
  Operation = "Repayment"
  GroupRepay_ISN = OverdraftGroupOperation(arrayCalcAcc, Count, opDate, Operation)
      
      ''SQL ստուգում Խմբային մարումից հետո
      ''HI
      queryString = "select count(*) from HI where fBASE = " & GroupRepay_ISN
      sql_Value = 14
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If
       
      ''HIR
      queryString = "select count(*) from HIR where fBASE = " & GroupRepay_ISN 
      sql_Value = 10
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If
      
      ''HIRREST
      queryString = "select count(*) from HIRREST where fOBJECT = " & fBASE &_
                     "and ((fLASTREM = 0.00 and fPENULTREM = 4763.90) or (fLASTREM = 0.00 and fPENULTREM = 4527.20))"
      sql_Value = 2
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If
    
      ''HIR
      queryString = "select count(*) from HIR where fOBJECT = " & fBASE                  
      sql_Value = 18
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If  
      
'-------------------------------------------------------------------------------      
  Call Log.Message("40.Ջնջել Բարդ օվերդրաֆտ պայամանագրի հետ կապված բոլոր փաստաթղթերը",,,attr)
  Workspace = "|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|"
  DocType = 2
  FirstDate = "090518"
  LastDate = "161018"
  Call GroupDelete(Workspace, DocType, DocNum, FirstDate, LastDate, Action)
  
  'Ջնջել ենթապայմանագրերը
  Workspace = "úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)"
  
  DocType = 1
  Call docDelete(Workspace, DocType, SubAgrDocNum1)
  
  DocType = 1
  Call docDelete(Workspace, DocType, SubAgrDocNum2)
  
  'Ջնջել օվերդրաֆտի տրամադրումը
  DocType = 2
  Call wTreeView.DblClickItem("|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|ä³ÛÙ³Ý³·ñ»ñ")
	Call Rekvizit_Fill("Dialog", 1, "General", "LEVEL", DocType) 
  Call Rekvizit_Fill("Dialog", 1, "General", "NUM", DocNum)
	Asbank.VBObject("frmAsUstPar").VBObject("CmdOK").Click()
 
  Date = "080518"
  Pttel = "_2"
  Typ = ""
  Call DeleteActions(Date, Pttel, Typ, MesBox)
   
  'Ջնջել մայր պայամնագիրը
  Call wMainForm.MainMenu.Click(c_Opers & "|" & c_Delete)
  
  Call ClickCmdButton(3, "²Ûá")
  Builtin.Delay(2000)
  wMDIClient.VBObject("frmPttel").Close
  
  'Ջնջել "Հիշարար օրդեր"-ները
  Call ChangeWorkspace(c_CustomerService) 
  Client = ""
  
  Date = "080518"
  Call DeletePayDoc(Date, MemOrd_ISN1, Typ, Client)
   
  Date = "151018"
  Call DeletePayDoc(Date, MemOrd_ISN2, Typ, Client)
  
  ''Ջնջել Գրաֆիկով օվերդրաֆտ պայամանագրի հետ կապված բոլոր փաստաթղթերը
  Call ChangeWorkspace(c_Overdraft)
  Workspace = "|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|"
  DocType = 2
  FirstDate = "^A[Del]"
  LastDate = "^A[Del]"
  Call GroupDelete(Workspace, DocType, DocNum1, FirstDate, LastDate, Action)
  
  'Ջնջել Գրաֆիկով օվերդրաֆտ պայամնագիրը
  DocType = 1
  Call wTreeView.DblClickItem("|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|ä³ÛÙ³Ý³·ñ»ñ")
	Call Rekvizit_Fill("Dialog", 1, "General", "LEVEL", DocType) 
  Call Rekvizit_Fill("Dialog", 1, "General", "NUM", DocNum1)
	Call ClickCmdButton(2, "Î³ï³ñ»É")

  Builtin.Delay(2000)
  Call wMainForm.MainMenu.Click(c_Opers & "|" & c_Delete)
  
  Call ClickCmdButton(3, "²Ûá")
  Builtin.Delay(2000)
  wMDIClient.VBObject("frmPttel").Close

  'Ջնջել "Հիշարար օրդեր"-ները
  Call ChangeWorkspace(c_CustomerService) 
  Date = "080518"
  Call DeletePayDoc(Date, MemOrd_ISN3, Typ, Client)
  
  Date = "151018"
  Call DeletePayDoc(Date, MemOrd_ISN4, Typ, Client)
'-------------------------------------------------------------------------------  
  Call Close_AsBank()
End Sub