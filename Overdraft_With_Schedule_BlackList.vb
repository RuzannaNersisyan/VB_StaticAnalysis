Option Explicit
'USEUNIT Library_Common  
'USEUNIT Subsystems_SQL_Library  
'USEUNIT Constants
'USEUNIT Overdraft_NewCases_Library
'USEUNIT Akreditiv_Library
'USEUNIT Loan_Agreements_Library
'USEUNIT Library_CheckDB

'Test Case ID 165842

Sub Overdraft_With_Schedule_BlackListTest()
  Dim fDATE, sDATE, my_vbObj
  Dim QueryString, ExpSQLValue, SQL_IsEqual
  Dim GiveOverdradt_ISN, RepaySchedule_ISN, CalcDoc_ISN, RepayDoc_ISN
  Dim CalcAcc, Data, FolderName, Date, EndDate, ReNew, dateFill, AddDates, IfExists, opDate, opPerSum, NewPrMoney, NewMoney,_
      OldPrMoney, OldMoney, Name, NameLen, ColNum, Pttel, Typ, Key, attr
  Dim Overdraft,dbFOLDERS(2)
    
  ''1.Համակարգ մուտք գործել ARMSOFT օգտագործողով:
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

  ''2.Անցում կատարել "Օվերդրաֆտ (տեղաբաշխված)" ԱՇՏ:
  Call ChangeWorkspace(c_Overdraft)

'----------------------------------------------------------------------------------------------   
  ''.Ջնջել բոլոր փաստաթղթերը:
  'Մուտք գործել "Օվերդրաֆտ (տեղաբաշխված)/Օվերդրաֆտ ունեցող հաշիվներ"
  CalcAcc = "33170160500"
  FolderName = "|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|"
  Call wTreeView.DblClickItem(FolderName & "úí»ñ¹ñ³ýï áõÝ»óáÕ Ñ³ßÇíÝ»ñ")
  Call Rekvizit_Fill("Dialog", 1, "General", "ACCMASK", CalcAcc)
  Call ClickCmdButton(2, "Î³ï³ñ»É")
   Builtin.Delay(2000)
  '"Գործողություններ/Բոլոր գործողություններ/Թղթապանակներ/Պայմանագրի թղթապանակ"
  If wMDIClient.VBObject("frmPttel").VBObject("tdbgView").ApproxCount <> 0 Then
     Builtin.Delay(2000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_Folders & "|" & c_AgrFolder)
    
   'Վերականգնել (ջնջել) "Մարումների գրաֆիկը"
    Name = "Ø³ñáõÙÝ»ñÇ ·ñ³ýÇÏ`  07/03/18"
    NameLen = 28
    ColNum = 0
    Pttel = "_2"
    IfExists = Find_Doc_By(Name, NameLen,ColNum, Pttel)
    If IfExists Then
      Call wMainForm.MainMenu.Click(c_Opers & "|" & c_Delete)
      Call ClickCmdButton(3, "²Ûá")
    End If
       
    Name = "¶ñ³ýÇÏáí ûí»ñ¹ñ³ýïÇ å³ÛÙ³Ý³·Çñ"
    NameLen = 30
    ColNum = 0
    Pttel = "_2"
    Call Find_Doc_By(Name, NameLen,ColNum, Pttel)
    
    'Ջնջել "Կանխավ վճարված տոկոսների վերադարձ"-ը 
    Date = "080318"
    Typ = "57"
    Key = "1"
    Call DeleteD(Date, Typ, Key)
    
    'Ջնջել "Տոկոսի մարում"-ը
    Date = "070318"
    Typ = "53"
    Key = "0"
    Call DeleteD(Date, Typ, Key)
    
    'Ջնջել "Տոկոսի հաշվարկում"-ը
    Date = "060318"
    Typ = "51"
    Key = "0"
    Call DeleteD(Date, Typ, Key)
    
    'Ջնջել "Օվերդրաֆտի տրամադրում"-ը
    Date = "070218"
    Typ = "21"
    Key = "0"
    Call DeleteD(Date, Typ, Key)
    
    wMDIClient.VBObject("frmPttel_2").Close
  
    Call wMainForm.MainMenu.Click(c_Opers & "|" & c_Delete)
    Call ClickCmdButton(3, "²Ûá")
  End If
   Builtin.Delay(2000)
  wMDIClient.VBObject("frmPttel").Close
'----------------------------------------------------------------------------------------------
  
  Call Log.Message("Գրաֆիկով Օվերդրաֆտ պայմանագրի ստեղծում",,,attr)
  Set Overdraft = New_Overdraft()
  With Overdraft
    .DocType = "¶ñ³ýÇÏáí ûí»ñ¹ñ³ýïÇ å³ÛÙ³Ý³·Çñ" 
    .Template = "0001"
    .CalcAcc = "33170160500"                                    
    .Limit = 100000
    .Date = "070218" 
    .GiveDate = "070218"
    .Term = "070219"
    .Percent = ""
    .NonUsedPercent = ""
    .Baj = ""
    .DateFill = ""
    .Paragraph = ""
    .SumsDateFillType = ""
    .PayDates = ""
    .PaperCode = 111
    Call .CreatePlOverdraft(FolderName & "Üáñ å³ÛÙ³Ý³·ñÇ ëï»ÕÍáõÙ")
  
    ''SQL ստուգում պայամանգիր ստեղցելուց հետո: 
      ''CONTRACTS
      QueryString = "select count(*) from CONTRACTS where fDGISN= '" & .fBASE & "'" &_
                      "and fDGAGRTYPE = 'C' and fDGMODTYPE = 3 and fDGAGRKIND = '8L'" &_
                      "and fDGSTATE = 206 and fDGSUMMA = 100000.00 and fDGALLSUMMA = 0.00"
      ExpSQLValue = 1
      ColNum = 0
      SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, ColNum)
      If Not SQL_IsEqual Then
        Log.Error("QueryString = " & QueryString & ":  Expected result = " & ExpSQLValue)
      End If  
                                
      ''FOLDERS
      QueryString = "select count(*) from FOLDERS where fISN= '" & .fBASE & "'"
      ExpSQLValue = 3
      ColNum = 0
      SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, ColNum)
      If Not SQL_IsEqual Then
        Log.Error("QueryString = " & QueryString & ":  Expected result = " & ExpSQLValue)
      End If                                   
  
    ''4.Մարման գրաֆիկի նշանակում
     Builtin.Delay(2000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_RepaySchedule)     
  
    Name = "Ø³ñáõÙÝ»ñÇ ·ñ³ýÇÏ"
    NameLen = 17
    ColNum = 0
    Pttel = ""
    IfExists = Find_Doc_By(Name, NameLen,ColNum, Pttel)
    If IfExists Then 
      Builtin.Delay(2000)
      Call wMainForm.MainMenu.Click(c_AllActions)
      Call wMainForm.PopupMenu.Click(c_View)
      RepaySchedule_ISN = wMDIClient.VBObject("frmASDocForm").DocFormCommon.Doc.isn
      wMDIClient.VBObject("frmASDocForm").Close
    End If
  
      ''SQL ստուգում Մարման գրաֆիկ ստեղցելուց հետո: 
      ''AGRSCHEDULE
      QueryString = "select count(*) from AGRSCHEDULE where fBASE= '" & RepaySchedule_ISN & "'" &_
                      "and fKIND = 9"
      ExpSQLValue = 1
      ColNum = 0
      SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, ColNum)
      If Not SQL_IsEqual Then
        Log.Error("QueryString = " & QueryString & ":  Expected result = " & ExpSQLValue)
      End If  
    
      ''CONTRACTS
      QueryString = "select fDGSTATE from CONTRACTS where fDGISN= '" & .fBASE & "'"
      ExpSQLValue = 1
      ColNum = 0
      SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, ColNum)
      If Not SQL_IsEqual Then
        Log.Error("QueryString = " & QueryString & ":  Expected result = " & ExpSQLValue)
      End If 
       
      ''FOLDERS
      QueryString = "select count(*) from FOLDERS where fISN= '" & RepaySchedule_ISN & "'"
      ExpSQLValue = 1
      ColNum = 0
      SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, ColNum)
      If Not SQL_IsEqual Then
        Log.Error("QueryString = " & QueryString & ":  Expected result = " & ExpSQLValue)
      End If
  
    ''5.Այլ վճարումների գրաֆիկի նշանակում
    Data = Find_Data ("¶ñ³ýÇÏáí ûí»ñ¹ñ³ýïÇ å³ÛÙ³Ý³·Çñ- "& Trim(.DocNum) &" {²ÝáõÝ ²½·³ÝáõÝÛ³Ý}",0)
    If Not Data then
       call Log.Error("Փաստաթուղթը չի գտնվել") 
       exit Sub
     End If
     
    call ContractAction (c_OtherPaySchedule)
    Call ClickCmdButton(1, "Î³ï³ñ»É")

     ''6."Գրաֆիկով օվերդրաֆտային պայմանագրի" համար կատարել "Գործողություններ/Բոլոր գործողություններ/Ուղարկել հաստատման " գործողությունը 
    'կանգնել պայմանագրի վրա
    Data = Find_Data ("¶ñ³ýÇÏáí ûí»ñ¹ñ³ýïÇ å³ÛÙ³Ý³·Çñ- "& Trim(.DocNum) &" {²ÝáõÝ ²½·³ÝáõÝÛ³Ý}",0)
    If Not Data then
       call Log.Error("Փաստաթուղթը չի գտնվել") 
       exit Sub
     End If
    call ContractAction(c_SendToVer)
    Call ClickCmdButton(5, "²Ûá")
     Builtin.Delay(2000)
    wMDIClient.VBObject("frmPttel").Close

    ''7.Անցում կատարել "<<Սև ցուցակ>> հաստատողի ԱՇՏ" :
    Call ChangeWorkspace(c_BLVerifyer)
  
    ''8.Մուտք գործել "Հաստատվող տեղաբաշխված միջոցներ" թղթապանակ:
    Call wTreeView.DblClickItem("|§ê¨ óáõó³Ï¦ Ñ³ëï³ïáÕÇ ²Þî|Ð³ëï³ïíáÕ ï»Õ³µ³ßËí³Í ÙÇçáóÝ»ñ ¨ »ñ³ßË³íáñáõÃÛáõÝÝ»ñ")
	   Call Rekvizit_Fill("Dialog", 1, "General", "SUBSYS", "C3") 
    Call Rekvizit_Fill("Dialog", 1, "General", "NUM", .DocNum) 
    Call ClickCmdButton(2, "Î³ï³ñ»É")
  
    ''9.Վավերացնել պայմանագիրը
     Builtin.Delay(2000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_ToConfirm)
    Call ClickCmdButton(1, "Ð³ëï³ï»É")
     Builtin.Delay(2000)
    wMDIClient.VBObject("frmPttel").Close
  
    ''10.Մուտք գործել "Օվերդրաֆտ (տեղաբաշխված)/Հաստատվող փաստաթղթեր 1" թղթապանակ:
    Call ChangeWorkspace(c_Overdraft)
    Call wTreeView.DblClickItem("|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I")
    Call ClickCmdButton(2, "Î³ï³ñ»É")
  
    ''11.Վավերացնել պայմանագիրը:
    Data = Find_Data (.DocNum, 2)
    If Not Data Then
       Call Log.Error("Փաստաթուղթը չի գտնվել") 
       Exit Sub
     End If
     Builtin.Delay(2000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_ToConfirm)
    Call ClickCmdButton(1, "Ð³ëï³ï»É")
    wMDIClient.VBObject("frmPttel").Close
  
    ''12.Մուտք գործել "Օվերդրաֆտ (Տեղաբաշխված )/ Պայմանագրեր" թղթապանակ: - Պայմանագիրը պետք է առկա լինի:
    IfExists = LetterOfCredit_Filter_Fill(FolderName, .DocLevel, .DocNum)
    If (Not IfExists) Then
      Call Log.Error("Պայմանագիրը առկա չէ")
      Exit sub
    End If
  
    Call Log.Message("Օվերդրաֆտի տրամադրում",,,attr)
    GiveOverdradt_ISN = Give_Overdradt(.GiveDate, .Limit, "2", Null, "", "2")
  
    ''SQL ստուգում Օվերդրաֆտ տրամադրելուց հետո:   
    BuiltIn.Delay(delay_small)
  
      ''AGRSCHEDULE
      QueryString = "select count(*) from AGRSCHEDULE where fAGRISN = '" & .fBASE &_
                         "' and ((fINC = 1 and fKIND = 9) or (fINC = 2 and fKIND = 3))" 
      ExpSQLValue = 2
      ColNum = 0
      SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, ColNum)
      If Not SQL_IsEqual Then
        Log.Error("QueryString = " & QueryString & ":  Expected result = " & ExpSQLValue)
      End If  
  
      ''AGRSCHEDULEVALUES
      QueryString = "select count(*) from AGRSCHEDULEVALUES where fAGRISN = '" & .fBASE & "'" 
      ExpSQLValue = 48
      ColNum = 0
      SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, ColNum)
      If Not SQL_IsEqual Then
        Log.Error("QueryString = " & QueryString & ":  Expected result = " & ExpSQLValue)
      End If  
       
      BuiltIn.Delay(delay_small)
    
      ''FOLDERS
      QueryString = "select count(*) from FOLDERS where fISN= '" & .fBASE & "'"
      ExpSQLValue = 5
      ColNum = 0
      SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, ColNum)
      If Not SQL_IsEqual Then
        Log.Error("QueryString = " & QueryString & ":  Expected result = " & ExpSQLValue)
      End If 
      
      Set dbFOLDERS(1) = New_DB_FOLDERS()
          dbFOLDERS(1).fFOLDERID = "LOANREGISTER"
          dbFOLDERS(1).fNAME = "C3Univer"
          dbFOLDERS(1).fKEY = .fBASE 
          dbFOLDERS(1).fISN = .fBASE 
          dbFOLDERS(1).fSTATUS = 1
          dbFOLDERS(1).fCOM = "²ÝáõÝ ²½·³ÝáõÝÛ³Ý"
          dbFOLDERS(1).fSPEC = "C38"& Trim(.DocNum) &"          111                               0                                                                                                                                                             0.00                                                                                                                                                                                                                                                                                               "

      Set dbFOLDERS(2) = New_DB_FOLDERS()
          dbFOLDERS(2).fFOLDERID = "LOANREGISTER2"
          dbFOLDERS(2).fNAME = "C3Univer"
          dbFOLDERS(2).fKEY = .fBASE 
          dbFOLDERS(2).fISN = .fBASE 
          dbFOLDERS(2).fSTATUS = "1"
          dbFOLDERS(2).fCOM = "²ÝáõÝ ²½·³ÝáõÝÛ³Ý"
          dbFOLDERS(2).fSPEC = "0"
        
      Call CheckDB_FOLDERS(dbFOLDERS(1), 1)
      Call CheckDB_FOLDERS(dbFOLDERS(2), 1)
    
      ''HI
      QueryString = "select count(*) from HI where fBASE= '" & .fBASE & "' and fSUM = 100000.00"
      ExpSQLValue = 2
      ColNum = 0
      SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, ColNum)
      If Not SQL_IsEqual Then
        Log.Error("QueryString = " & QueryString & ":  Expected result = " & ExpSQLValue)
      End If
    
      ''HI
      QueryString = "select count(*) from HI where fBASE= '" & GiveOverdradt_ISN & "'" &_ 
                        "and fSUM = 100000.00 and fCURSUM = 100000.00"
      ExpSQLValue = 3
      ColNum = 0
      SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, ColNum)
      If Not SQL_IsEqual Then
        Log.Error("QueryString = " & QueryString & ":  Expected result = " & ExpSQLValue)
      End If
    
      ''HIF
      QueryString = "select count(*) from HIF where fBASE= '" & .fBASE & "'"
      ExpSQLValue = 19
      ColNum = 0
      SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, ColNum)
      If Not SQL_IsEqual Then
        Log.Error("QueryString = " & QueryString & ":  Expected result = " & ExpSQLValue)
      End If
    
      ''HIR
      QueryString = "select count(*) from HIR where fBASE= '" & GiveOverdradt_ISN &_ 
                        "' and fOBJECT = '" & .fBASE & "' and fCURSUM = 100000.00 and fTYPE = 'R1'"
      ExpSQLValue = 1
      ColNum = 0
      SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, ColNum)
      If Not SQL_IsEqual Then
        Log.Error("QueryString = " & QueryString & ":  Expected result = " & ExpSQLValue)
      End If
    
      ''HIRREST
      QueryString = "select fLASTREM from HIRREST where fOBJECT= '" & .fBASE & "'"
      ExpSQLValue = 100000.00
      ColNum = 0
      SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, ColNum)
      If Not SQL_IsEqual Then
        Log.Error("QueryString = " & QueryString & ":  Expected result = " & ExpSQLValue)
      End If
   
    Call Log.Message("Օվերդրաֆտի տոկոսների հաշվարկ",,,attr)
    wMDIClient.VBObject("frmPttel").Close
    opDate = "060318"
    CalcDoc_ISN = Overdraft_Percent_Accounting(.DocNum, opDate)
  
      ''SQL ստուգում Օվերդրաֆտի տոկոսների հաշվարկից հետո:      
        ''HI
        QueryString = "select count(*) from HI where fBASE= '" & CalcDoc_ISN &_
                         "' and fSUM = 767.10 and fCURSUM = 767.10"
        ExpSQLValue = 4
        ColNum = 0
        SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, ColNum)
        If Not SQL_IsEqual Then
          Log.Error("QueryString = " & QueryString & ":  Expected result = " & ExpSQLValue)
        End If
    
        ''HIF
        QueryString = "select count(*) from HIF where fBASE= '" & CalcDoc_ISN &_
                         "' and fSUM = 0.00 and fCURSUM = 0.00"
        ExpSQLValue = 1
        ColNum = 0
        SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, ColNum)
        If Not SQL_IsEqual Then
          Log.Error("QueryString = " & QueryString & ":  Expected result = " & ExpSQLValue)
        End If
    
        ''HIR
        QueryString = "select count(*) from HIR where fBASE= '" & CalcDoc_ISN &_
                          "'  and fOBJECT = '" & .fBASE & "' and (fCURSUM = 767.10 or fCURSUM = 8333.30)"
        ExpSQLValue = 3
        ColNum = 0
        SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, ColNum)
        If Not SQL_IsEqual Then
          Log.Error("QueryString = " & QueryString & ":  Expected result = " & ExpSQLValue)
        End If
    
        ''HIT
        QueryString = "select count(*) from HIT where fOBJECT= '" & .fBASE &_ 
                          "' and fOBJECT = '" & .fBASE & "' and fCURSUM = 767.10 and fTYPE = 'N2'"
        ExpSQLValue = 1
        ColNum = 0
        SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, ColNum)
        If Not SQL_IsEqual Then
          Log.Error("QueryString = " & QueryString & ":  Expected result = " & ExpSQLValue)
        End If
  
    Call Log.Message("Օվերդրաֆտի պարտքերի մարում` նախատեսվածից ավելի Տոկոսագումար մարելով",,,attr)
    wMDIClient.VBObject("frmPttel").Close
    opDate = "070318"
    opPerSum = "1000"
    RepayDoc_ISN = Overdraft_Repayment_Operation(.DocNum, opDate, "", opPerSum, "")
 
    ''SQL ստուգում Օվերդրաֆտի պարտքերի մարումից հետո: 
  
    BuiltIn.Delay(delay_small)
  
      ''AGRSCHEDULE
      QueryString = "select count(*) from AGRSCHEDULE where fAGRISN = '" & .fBASE &_
                        "' and ((fINC = 1 and fKIND = 9) or (fINC = 2 and fKIND = 3) or (fINC = 3 and fKIND = 2))" 
      ExpSQLValue = 3
      ColNum = 0
      SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, ColNum)
      If Not SQL_IsEqual Then
        Log.Error("QueryString = " & QueryString & ":  Expected result = " & ExpSQLValue)
      End If  
         
      ''HI
      QueryString = "select count(*) from HI where fBASE= '" & RepayDoc_ISN & "'"
      ExpSQLValue = 5
      ColNum = 0
      SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, ColNum)
      If Not SQL_IsEqual Then
        Log.Error("QueryString = " & QueryString & ":  Expected result = " & ExpSQLValue)
      End If

      ''HIR
      QueryString = "select count(*) from HIR where fBASE= '" & RepayDoc_ISN &_
                        "' and (fCURSUM = 8333.30 or fCURSUM = 1000.00 or fCURSUM = 767.10)"
      ExpSQLValue = 4
      ColNum = 0
      SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, ColNum)
      If Not SQL_IsEqual Then
        Log.Error("QueryString = " & QueryString & ":  Expected result = " & ExpSQLValue)
      End If
      
    ''17.Ստուգել "Հաշվ.% մնացորդ" դաշտի արժեքը բացասական լինի:
    my_vbObj = wMDIClient.VBObject("frmPttel").VBObject("tdbgView").Columns.Item(4) 
    If my_vbObj >= 0 Then
      Call Log_Error("Հաշվ.% մնացորդ դաշտի արժեքը բացասական չէ:")
    End If
 
    Call Log.Message("Կանխավ վճարված տոկոսների վերադարձ",,,attr)
    Date = "080318"
    Call ReturnPrepaidRates(Date, "")
  
    ''19.Ստուգել, որ "Հաշվ.% մնացորդ" դաշտի արժեքը զրոյացած լինի:
    my_vbObj = wMDIClient.VBObject("frmPttel").VBObject("tdbgView").Columns.Item(4) 
    If my_vbObj <> 0 Then
      Call Log.Error("Հաշվ.% մնացորդ դաշտի արժեքը չի զրոյացել:")
    End If
  
    Call Log.Message("Մարման գրաֆիկի վերանայում",,,attr)
    Date = "070318"
    EndDate = "070220"
    ReNew = "1"
    dateFill = 1
    Call PaymentScheduleReview(Date, EndDate, ReNew, dateFill, NewMoney, NewPrMoney)
  
    ''21.Ստուգել, որ Մարման գրաֆիկի Գումարը և Տոկոսագումարը փոխված լինեն:
    OldPrMoney = "5,390.70"
    OldMoney = "100,000.00"
    If OldPrMoney = NewPrMoney Then
      Call Log.Error("Մարման գրաֆիկի Գումարը չի փոխվել:")
    End If
  
    If OldMoney = NewMoney Then
      Call Log.Error("Մարման գրաֆիկի Տոկոսագումարը չի փոխվել:")
    End If
    wMDIClient.VBObject("frmPttel").Close
  
  '----------------------------------------------------------------------------------------------   
     Call Log.Message("Ջնջել բոլոր փաստաթղթերը",,,attr)
    'Մուտք գործել "Օվերդրաֆտ (տեղաբաշխված)/Օվերդրաֆտ ունեցող հաշիվներ"
    Call wTreeView.DblClickItem(FolderName & "úí»ñ¹ñ³ýï áõÝ»óáÕ Ñ³ßÇíÝ»ñ")
    Call Rekvizit_Fill("Dialog", 1, "General", "ACCMASK", .CalcAcc)
    Call ClickCmdButton(2, "Î³ï³ñ»É")
  
    '"Գործողություններ/Բոլոր գործողություններ/Թղթապանակներ/Պայմանագրի թղթապանակ"
    If wMDIClient.VBObject("frmPttel").VBObject("tdbgView").ApproxCount <> 0 Then
      Call wMainForm.MainMenu.Click(c_AllActions)
      Call wMainForm.PopupMenu.Click(c_Folders & "|" & c_AgrFolder)
    
     'Վերականգնել (ջնջել) "Մարումների գրաֆիկը"
      Name = "Ø³ñáõÙÝ»ñÇ ·ñ³ýÇÏ`  07/03/18"
      NameLen = 30
      ColNum = 0
      Pttel = "_2"
      Call Find_Doc_By(Name, NameLen,ColNum, Pttel)
      Call wMainForm.MainMenu.Click(c_Opers & "|" & c_Delete)
      Call ClickCmdButton(3, "²Ûá")
    
      Name = "¶ñ³ýÇÏáí ûí»ñ¹ñ³ýïÇ å³ÛÙ³Ý³·Çñ"
      NameLen = 30
      ColNum = 0
      Pttel = "_2"
      Call Find_Doc_By(Name, NameLen,ColNum, Pttel)
    
      'Ջնջել "Կանխավ վճարված տոկոսների վերադարձ"-ը 
      Date = "080318"
      Typ = "57"
      Key = "1"
      Call DeleteD(Date, Typ, Key)
    
      'Ջնջել "Տոկոսի մարում"-ը
      Date = "070318"
      Typ = "53"
      Key = "0"
      Call DeleteD(Date, Typ, Key)
    
      'Ջնջել "Տոկոսի հաշվարկում"-ը
      Date = "060318"
      Typ = "51"
      Key = "0"
      Call DeleteD(Date, Typ, Key)
    
      'Ջնջել "Օվերդրաֆտի տրամադրում"-ը
      Date = "070218"
      Typ = "21"
      Key = "0"
      Call DeleteD(Date, Typ, Key)
    
      wMDIClient.VBObject("frmPttel_2").Close
  
      Call wMainForm.MainMenu.Click(c_Opers & "|" & c_Delete)
      Call ClickCmdButton(3, "²Ûá")
    End If
     Builtin.Delay(2000)
    wMDIClient.VBObject("frmPttel").Close
  '----------------------------------------------------------------------------------------------
  End With
  Call Close_AsBank()
End Sub
