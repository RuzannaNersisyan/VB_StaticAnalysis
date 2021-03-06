Option Explicit

'USEUNIT Library_Common 
'USEUNIT Subsystems_SQL_Library  
'USEUNIT Akreditiv_Library
'USEUNIT Loan_Agreements_Library 
'USEUNIT Loan_Agreements_With_Schedule_Linear_Library
'USEUNIT Loan_Agreemnts_With_Schedule_Library
'USEUNIT Credit_Line_Library
'USEUNIT Group_Operations_Library
'USEUNIT Constants

'Test Case Id 165858
'Test Case Id 165859

Sub Complicated_Loan_Agreement_Linear_Unused_Part_Interests_Test(Renewable)
  Dim fDATE, sDATE, frmAsMsgBox, FrmSpr, attr
  Dim queryString, sqlValue, colNum, sql_isEqual
  Dim CollectFromProvision_ISN, GiveCredit_ISN, IntCalc_ISN, ChangeLim_ISN,_
      Repay_ISN, Agr_ISN, SubAgr_ISN, GroupCalc_ISN, summa
  Dim Loan, LoanWithSchedule, FolderName, opDate, Sum, calcDate, exTerm, MainSum, PerSum, Prc,_
      NonUsedPrc, EffRete, ActRete, ExpectedSum
  Dim arrCheckbox     
      
'--------------------------------------
  Set attr = Log.CreateNewAttributes
  attr.BackColor = RGB(0, 255, 255)
  attr.Bold = True
  attr.Italic = True
'--------------------------------------
    
  ''1.Համակարգ մուտք գործել ARMSOFT օգտագործողով
  fDATE = "20260101"
  sDATE = "20140101"
  Call Initialize_AsBank("bank", sDATE, fDATE)
  Login("ARMSOFT")
  Call Create_Connection()
  
  ''2.Մուտք գործել "Վարկեր (տեղաբաշխված)"
  Call ChangeWorkspace(c_Loans)  
  
  Call Log.Message("Բարդ վարկ (գծային) պայմանագրի ստեղծում",,,attr)
  FolderName = "|ì³ñÏ»ñ (ï»Õ³µ³ßËí³Í)|"
  Set Loan = New_LoanDocument()
  With Loan
    .DocType = "´³ñ¹ í³ñÏ (·Í³ÛÇÝ)"
    .CalcAcc = "00000113032"                                    
    .Limit = 1000000
    .Date = "050219"
    .Percent = 0 
    .GiveDate = "050219"
    .Term = "050220"
    .Renewable = Renewable
    .PaperCode = 123
  
    Call .CreatePlLoan(FolderName & "Üáñ å³ÛÙ³Ý³·ñÇ ëï»ÕÍáõÙ")
    
    Log.Message(.DocNum)
    Agr_ISN = .fBase
    
        ''SQL ստուգում պայամանգիր ստեղցելուց հետո: 
          ''CONTRACTS
          queryString = "SELECT COUNT(*) FROM CONTRACTS WHERE fDGISN = " & .fBASE &_
                          "AND fDGAGRTYPE = 'C' AND fDGAGRCHILDREN = 0 " &_
                          "AND fDGSUMMA = 1000000.00 and fDGALLSUMMA = 0.00"
          sqlValue = 1
          colNum = 0
          sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
          If Not sql_isEqual Then
            Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
          End If  
                                
          ''FOLDERS
          queryString = "SELECT COUNT(*) FROM FOLDERS WHERE fISN = " & .fBASE 
          sqlValue = 3
          colNum = 0
          sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
          If Not sql_isEqual Then
            Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
          End If   

    'Պայմանագրին ուղղարկել հաստատման
    .SendToVerify(Null)
    'Հաստատել
    .Verify("|ì³ñÏ»ñ (ï»Õ³µ³ßËí³Í)|Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I")
    
    Call LetterOfCredit_Filter_Fill(FolderName, .DocLevel, .DocNum)
    
    Call Log.Message("Ենթապայմանագրի բացում",,,attr)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_OpenSubAgr)
    Set LoanWithSchedule = New_LoanDocument()
    With LoanWithSchedule
      .DocType = "¶ñ³ýÇÏáí í³ñÏ³ÛÇÝ å³ÛÙ³Ý³·Çñ"
      .Limit = 100000
      .Date = "050219"
      .Percent = 0 
      .GiveDate = "050219"
      .Term = "050220"
      .PaperCode = 124
      
      Call .CreatePlLoan(Null)
    
      Log.Message(.DocNum)
      SubAgr_ISN = .fBase
      
         ''SQL ստուգում պայամանգիր ստեղցելուց հետո: 
          ''CONTRACTS
          queryString = "SELECT COUNT(*) FROM CONTRACTS WHERE fDGISN = " & SubAgr_ISN &_
                          "AND fDGAGRTYPE = 'C' AND fDGAGRCHILDREN = 1 AND fDGSTATE = 206 " &_
                          "AND fDGSUMMA = 100000.00 AND fDGALLSUMMA = 0.00"
          sqlValue = 1
          colNum = 0
          sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
          If Not sql_isEqual Then
            Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
          End If  
                                
          ''FOLDERS
          queryString = "SELECT COUNT(*) FROM FOLDERS WHERE fISN = " & SubAgr_ISN 
          sqlValue = 2
          colNum = 0
          sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
          If Not sql_isEqual Then
            Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
          End If                          
      
      'Ենթապայմանագրին ուղղարկել հաստատման
      wMDIClient.VBObject("frmPttel").Close
      .SendToVerify("|ì³ñÏ»ñ (ï»Õ³µ³ßËí³Í)|²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ")
      'Հաստատել
      .Verify("|ì³ñÏ»ñ (ï»Õ³µ³ßËí³Í)|Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I")
    
      Call LetterOfCredit_Filter_Fill(FolderName, .DocLevel, .DocNum)
    
      Call Log.Message("Գանձում տրամադրումից ենթապայմանագրի համար",,,attr)
      Call Collect_From_Provision(.Date, summa, 2, Loan.CalcAcc, CollectFromProvision_ISN)
      
      Call Log.Message("Վարկի տրամադրում ենթապայմանագրի համար",,,attr)
      Call Give_Credit(.Date, .Limit, 2, Loan.CalcAcc, GiveCredit_ISN)
      BuiltIn.Delay(2000)
      wMDIClient.VBObject("frmPttel").Close
    End With 
    
    Call LetterOfCredit_Filter_Fill(FolderName, .DocLevel, .DocNum)
    
    Call Log.Message("Տոկոսների հաշվարկ մայր պայմանագրի համար` 06/02/19 ամսաթվով",,,attr)
    calcDate = "060219"
    IntCalc_ISN = Calculate_Percents(calcDate, calcDate, False)
    
    'Ստուգել, որ հաշվարկված գումարը հավասար լինի սպասվածին
    ExpectedSum = "394.50"
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_OpersView)
    Call Rekvizit_Fill("Dialog", 1, "General", "START", calcDate) 
    Call Rekvizit_Fill("Dialog", 1, "General", "END", calcDate)
    Call Rekvizit_Fill("Dialog", 1, "General", "DEALTYPE", "^A[Del]" & "[Tab]")
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    BuiltIn.Delay(2000)
    
    If Trim(wMDIClient.VBObject("frmPttel_2").VBObject("tdbgView").Columns.Item(4).Text) <> ExpectedSum Then
         Call Log.Error("Չօգտ. մասի տոկոսի հաշվարկը սխալ է:")
    End If
    wMDIClient.VBObject("frmPttel_2").Close
    wMDIClient.VBObject("frmPttel").Close

        ''SQL ստուգում Տոկոսների հաշվարկից հետո: 
          ''HI
          queryString = "SELECT COUNT(*) FROM HI WHERE fBASE = " & IntCalc_ISN &_
                         "AND fSUM = 394.50 AND fCURSUM = 394.50" 
          sqlValue = 2
          colNum = 0
          sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
          If Not sql_isEqual Then
            Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
          End If  
        
          ''HIF
          queryString = "SELECT COUNT(*) FROM HIF WHERE fBASE = " & .fBASE 
          sqlValue = 29
          colNum = 0
          sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
          If Not sql_isEqual Then
            Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
          End If  
          
          ''HIF
          queryString = "SELECT COUNT(*) FROM HIF WHERE fBASE = " & IntCalc_ISN &_
                         "AND fSUM = 0.00 AND fCURSUM = 0.00" 
          sqlValue = 1
          colNum = 0
          sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
          If Not sql_isEqual Then
            Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
          End If  
          
          ''HIR
          queryString = "SELECT COUNT(*) FROM HIR WHERE fBASE = " & IntCalc_ISN &_
                         "AND fCURSUM = 394.50" 
          sqlValue = 1
          colNum = 0
          sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
          If Not sql_isEqual Then
            Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
          End If  
          
          ''HIRREST
          queryString = "SELECT COUNT(*) FROM HIRREST WHERE fOBJECT = " & .fBASE &_
                         "AND fLASTREM = 394.50 AND fPENULTREM = 0.00 AND fSTARTREM = 0.00" 
          sqlValue = 1
          colNum = 0
          sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
          If Not sql_isEqual Then
            Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
          End If  
          
          ''HIT
          queryString = "SELECT COUNT(*) FROM HIT WHERE fBASE = " & IntCalc_ISN &_
                         "AND fCURSUM = 394.50" 
          sqlValue = 1
          colNum = 0
          sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
          If Not sql_isEqual Then
            Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
          End If  
          
    Call LetterOfCredit_Filter_Fill(FolderName, LoanWithSchedule.DocLevel, LoanWithSchedule.DocNum)
    
    Call Log.Message("Տոկոսների հաշվարկ  ենթապայմանագրի համար` 06/02/19 ամսաթվով",,,attr)
    IntCalc_ISN = Calculate_Percents(calcDate, calcDate, False)
    BuiltIn.Delay(2000)
    wMDIClient.VBObject("frmPttel").Close

        ''SQL ստուգում Տոկոսների հաշվարկից հետո: 
          ''HI
          queryString = "SELECT COUNT(*) FROM HI WHERE fBASE = " & IntCalc_ISN &_
                         "AND fSUM = 1.00 AND fCURSUM = 1.00" 
          sqlValue = 2
          colNum = 0
          sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
          If Not sql_isEqual Then
            Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
          End If  
               
          ''HIF
          queryString = "SELECT COUNT(*) FROM HIF WHERE fBASE = " & IntCalc_ISN &_
                         "AND fSUM = 0.00 AND fCURSUM = 0.00" 
          sqlValue = 1
          colNum = 0
          sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
          If Not sql_isEqual Then
            Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
          End If  
          
          ''HIR
          queryString = "SELECT COUNT(*) FROM HIR WHERE fBASE = " & IntCalc_ISN &_
                         "AND fCURSUM = 1.00" 
          sqlValue = 1
          colNum = 0
          sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
          If Not sql_isEqual Then
            Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
          End If  
          
          ''HIT
          queryString = "SELECT COUNT(*) FROM HIT WHERE fBASE = " & IntCalc_ISN &_
                         "AND fCURSUM = 1.00" 
          sqlValue = 1
          colNum = 0
          sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
          If Not sql_isEqual Then
            Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
          End If  
        
    Call LetterOfCredit_Filter_Fill(FolderName, .DocLevel, .DocNum)
    
    Call Log.Message("Վարկի սահմանաչափի փոփոխում:",,,attr)
    opDate = "150219"
    Sum = 2000000
    ChangeLim_ISN = Change_Limit(opDate , Sum)    

        ''SQL ստուգում սահմանաչափի փոփոխումից հետո: 
          ''FOLDERS
          queryString = "SELECT COUNT(*) FROM FOLDERS WHERE fISN = " & ChangeLim_ISN 
          sqlValue = 1
          colNum = 0
          sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
          If Not sql_isEqual Then
            Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
          End If 
          
          ''HI
          queryString = "SELECT COUNT(*) FROM HI WHERE fBASE = " & ChangeLim_ISN &_
                         "AND fSUM = 1000000.00 AND fCURSUM = 1000000.00" 
          sqlValue = 2
          colNum = 0
          sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
          If Not sql_isEqual Then
            Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
          End If  
               
          ''HIF
          queryString = "SELECT COUNT(*) FROM HIF WHERE fBASE = " & ChangeLim_ISN &_
                         "AND ((fSUM = 2000000.00 AND fCURSUM = 0.00) OR (fSUM = 0.00 AND fCURSUM = 365.00))" 
          sqlValue = 3
          colNum = 0
          sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
          If Not sql_isEqual Then
            Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
          End If  
        
    Call Log.Message("Տոկոսների հաշվարկ մայր պայմանագրի համար` 04/03/19 ամսաթվով",,,attr)
    calcDate = "040319"
    IntCalc_ISN = Calculate_Percents(calcDate, calcDate, False)
    
        ''SQL ստուգում Տոկոսների հաշվարկից հետո: 
          ''HI
          queryString = "SELECT COUNT(*) FROM HI WHERE fBASE = " & IntCalc_ISN &_
                         "AND fSUM = 9074.00 AND fCURSUM = 9074.00" 
          sqlValue = 2
          colNum = 0
          sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
          If Not sql_isEqual Then
            Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
          End If  
               
          ''HIF
          queryString = "SELECT COUNT(*) FROM HIF WHERE fBASE = " & IntCalc_ISN &_
                         "AND fSUM = 0.00 AND fCURSUM = 0.00" 
          sqlValue = 1
          colNum = 0
          sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
          If Not sql_isEqual Then
            Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
          End If  
          
          ''HIR
          queryString = "SELECT COUNT(*) FROM HIR WHERE fBASE = " & IntCalc_ISN &_
                         "AND (fCURSUM = 9074.00 OR fCURSUM = 9468.50)" 
          sqlValue = 2
          colNum = 0
          sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
          If Not sql_isEqual Then
            Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
          End If  
          
          ''HIT
          queryString = "SELECT COUNT(*) FROM HIT WHERE fBASE = " & IntCalc_ISN &_
                         "AND fCURSUM = 9074.00" 
          sqlValue = 1
          colNum = 0
          sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
          If Not sql_isEqual Then
            Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
          End If  
    
    'Ստուգել, որ հաշվարկված գումարը հավասար լինի սպասվածին
    ExpectedSum = "9,074.00"
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_OpersView)
    Call Rekvizit_Fill("Dialog", 1, "General", "START", calcDate) 
    Call Rekvizit_Fill("Dialog", 1, "General", "END", calcDate)
    Call Rekvizit_Fill("Dialog", 1, "General", "DEALTYPE", "^A[Del]" & "[Tab]")
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    
    BuiltIn.Delay(2000)
    If Trim(wMDIClient.VBObject("frmPttel_2").VBObject("tdbgView").Columns.Item(4).Text) <> ExpectedSum Then
         Call Log.Error("Չօգտ. մասի տոկոսի հաշվարկը սխալ է:")
    End If
    
    BuiltIn.Delay(2000)
    wMDIClient.VBObject("frmPttel_2").Close
    BuiltIn.Delay(2000)
    wMDIClient.VBObject("frmPttel").Close
    
    Call LetterOfCredit_Filter_Fill(FolderName, LoanWithSchedule.DocLevel, LoanWithSchedule.DocNum)
    
    Call Log.Message("Տոկոսների հաշվարկ  ենթապայմանագրի համար` 04/03/19 ամսաթվով",,,attr)
    IntCalc_ISN = Calculate_Percents(calcDate, calcDate, False)
    
        ''SQL ստուգում Տոկոսների հաշվարկից հետո: 
          ''HI
          queryString = "SELECT COUNT(*) FROM HI WHERE fBASE = " & IntCalc_ISN &_
                         "AND fSUM = 13.20 AND fCURSUM = 13.20" 
          sqlValue = 2
          colNum = 0
          sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
          If Not sql_isEqual Then
            Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
          End If  
               
          ''HIF
          queryString = "SELECT COUNT(*) FROM HIF WHERE fBASE = " & IntCalc_ISN &_
                         "AND fSUM = 0.00 AND fCURSUM = 0.00" 
          sqlValue = 1
          colNum = 0
          sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
          If Not sql_isEqual Then
            Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
          End If  
          
          ''HIR
          queryString = "SELECT COUNT(*) FROM HIR WHERE fBASE = " & IntCalc_ISN &_
                         "AND (fCURSUM = 13.20 OR fCURSUM = 8333.30)" 
          sqlValue = 2
          colNum = 0
          sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
          If Not sql_isEqual Then
            Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
          End If  
          
          ''HIT
          queryString = "SELECT COUNT(*) FROM HIT WHERE fBASE = " & IntCalc_ISN &_
                         "AND (fCURSUM = 13.20 OR fCURSUM = 0.00)" 
          sqlValue = 2
          colNum = 0
          sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
          If Not sql_isEqual Then
            Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
          End If 
              
    Call Log.Message("Պարտքերի մարում ենթապայմանագրի համար",,,attr)
    opDate = "050319"
    Call Fade_Debt(opDate, Repay_ISN, "", "", "", False)
    
          BuiltIn.Delay(2000)
        ''SQL ստուգում Պարտքերի մարումից հետո:
         ''AGRSCHEDULEVALUES
          queryString = "SELECT COUNT(*) FROM AGRSCHEDULEVALUES WHERE fAGRISN = " & SubAgr_ISN &_
                         "AND (fSUM = 8333.30 OR fSUM = 0.00 OR fSUM = 8333.70)" 
          sqlValue = 46
          colNum = 0
          sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
          If Not sql_isEqual Then
            Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
          End If  
         
          ''HI
          queryString = "SELECT COUNT(*) FROM HI WHERE fBASE = " & Repay_ISN &_
                         "AND fSUM = 8333.30 AND fCURSUM = 8333.30" 
          sqlValue = 3
          colNum = 0
          sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
          If Not sql_isEqual Then
            Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
          End If  
               
          ''HIF
          queryString = "SELECT COUNT(*) FROM HIF WHERE fBASE = " & Repay_ISN &_
                         "AND fSUM = 1991666.70" 
          If Renewable = 0 Then                         
            sqlValue = 1
          Else   
            sqlValue = 0
          End If  
          colNum = 0
          sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
          If Not sql_isEqual Then
            Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
          End If  
          
          ''HIR
          queryString = "SELECT COUNT(*) FROM HIR WHERE fBASE = " & Repay_ISN &_
                         "AND fCURSUM = 8333.30" 
          sqlValue = 2
          colNum = 0
          sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
          If Not sql_isEqual Then
            Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
          End If  
          
    BuiltIn.Delay(2000)      
    wMDIClient.VBObject("frmPttel").Close
    
    Call LetterOfCredit_Filter_Fill(FolderName, .DocLevel, .DocNum)
    
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_ViewEdit & "|" & c_Other & "|" & c_Limits)
    If Renewable = 0 Then
      'Ստուգել, որ սահմանաչափը նվազած լինի մարված գումարի չափով
      Call Rekvizit_Fill("Dialog", 1, "General", "START", opDate) 
      Call Rekvizit_Fill("Dialog", 1, "General", "END", opDate)
      Call ClickCmdButton(2, "Î³ï³ñ»É")
      ExpectedSum = "1,991,666.70"
      If Trim(wMDIClient.VBObject("frmPttel_2").VBObject("tdbgView").Columns.Item(3).Text) <> ExpectedSum Then
           Call Log.Error("Սահմանաչափը չի նվազել մարված գումարի չափով")
      End If
    Else
      'Ստուգել, որ սահմանաչափը պահպանված լինի
      Call Rekvizit_Fill("Dialog", 1, "General", "START", "^A[Del]" ) 
      Call Rekvizit_Fill("Dialog", 1, "General", "END", "^A[Del]")
      Call Rekvizit_Fill("Dialog", 1, "CheckBox", "ONLYCH", 1)
      Call ClickCmdButton(2, "Î³ï³ñ»É")
      BuiltIn.Delay(2000)
      ExpectedSum = "2,000,000.00"
      If Trim(wMDIClient.VBObject("frmPttel_2").VBObject("tdbgView").Columns.Item(3).Text) <> ExpectedSum Then
        Call Log.Error("Սահմանաչափը փոփոխվել է")
      End If
    End If   
    wMDIClient.VBObject("frmPttel_2").Close
    
    Call Log.Message("Խմբային տոկոսների հաշվարկ մայր պայմանագրի համար` 20/03/19 ամսաթվով",,,attr)
    opDate = "200319"
    ReDim arrCheckbox(1)          
    arrCheckbox = Array("CHG")
    Call Group_Calculation(opDate, arrCheckbox)
    
    'Վերցնել խմբային հաշվարկի ISN-ը
    BuiltIn.Delay(2000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_OpersView)
    Call Rekvizit_Fill("Dialog", 1, "General", "START", opDate)
    Call Rekvizit_Fill("Dialog", 1, "General", "END", opDate)
    Call ClickCmdButton(2, "Î³ï³ñ»É")
  
    BuiltIn.Delay(2000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_View)
    GroupCalc_ISN = wMDIClient.VBObject("frmASDocForm").DocFormCommon.Doc.isn
    BuiltIn.Delay(2000)
    wMDIClient.VBObject("frmASDocForm").Close
  
    BuiltIn.Delay(2000)
    wMDIClient.VBObject("frmPttel_2").Close
    
        ''SQL ստուգում Տոկոսների հաշվարկից հետո: 
          ''HIF
          queryString = "SELECT COUNT(*) FROM HIF WHERE fBASE = " & GroupCalc_ISN &_
                         "AND fSUM = 0.00 AND fCURSUM = 0.00" 
          sqlValue = 3
          colNum = 0
          sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
          If Not sql_isEqual Then
            Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
          End If  
          
          If Renewable = 0 Then
            ''HIR
            queryString = "SELECT COUNT(*) FROM HIR WHERE fBASE = " & GroupCalc_ISN &_
                           "AND (fCURSUM = 416.40 OR fCURSUM = 6246.60)" 
            sqlValue = 2
            colNum = 0
            sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
            If Not sql_isEqual Then
              Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
            End If  
          
            ''HIT
            queryString = "SELECT COUNT(*) FROM HIT WHERE fBASE = " & GroupCalc_ISN &_
                           "AND (fCURSUM = 416.40 OR fCURSUM = 6246.60)" 
            sqlValue = 2
            colNum = 0
            sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
            If Not sql_isEqual Then
              Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
            End If 
          Else
            ''HIR
            queryString = "SELECT COUNT(*) FROM HIR WHERE fBASE = " & GroupCalc_ISN &_
                           "AND (fCURSUM = 418.30 OR fCURSUM = 6273.90)" 
            sqlValue = 2
            colNum = 0
            sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
            If Not sql_isEqual Then
              Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
            End If  
          
            ''HIT
            queryString = "SELECT COUNT(*) FROM HIT WHERE fBASE = " & GroupCalc_ISN &_
                           "AND (fCURSUM = 418.30 OR fCURSUM = 6273.90)" 
            sqlValue = 2
            colNum = 0
            sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
            If Not sql_isEqual Then
              Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
            End If 
          End If
    
    'Ստուգել, որ հաշվարկված գումարը հավասար լինի սպասվածին
    If Renewable = 0 Then
      ExpectedSum = "416.40" 
    Else 
      ExpectedSum = "418.30"     
    End If
    
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_OpersView)
    Call Rekvizit_Fill("Dialog", 1, "General", "START", opDate) 
    Call Rekvizit_Fill("Dialog", 1, "General", "END", opDate)
    Call Rekvizit_Fill("Dialog", 1, "General", "DEALTYPE", "^A[Del]" & "[Tab]")
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    
    BuiltIn.Delay(2000)
    If Trim(wMDIClient.VBObject("frmPttel_2").VBObject("tdbgView").Columns.Item(4).Text) <> ExpectedSum Then
       Call Log.Error("Չօգտ. մասի տոկոսի հաշվարկը սխալ է:")
    End If
    wMDIClient.VBObject("frmPttel_2").VBObject("tdbgView").MoveNext
    
    If Renewable = 0 Then
      ExpectedSum = "6,246.60" 
    Else 
      ExpectedSum = "6,273.90"     
    End If
    If Trim(wMDIClient.VBObject("frmPttel_2").VBObject("tdbgView").Columns.Item(4).Text) <> ExpectedSum Then
       Call Log.Error("Չօգտ. մասի տոկոսի հաշվարկը սխալ է:")
    End If
    
    BuiltIn.Delay(2000)
    wMDIClient.VBObject("frmPttel_2").Close
    BuiltIn.Delay(2000)
    wMDIClient.VBObject("frmPttel").Close
    
    Call Log.Message("Բոլոր փաստաթղթերի ջնջում",,,attr)
    
    Call DeleteAllActions("|ì³ñÏ»ñ (ï»Õ³µ³ßËí³Í)|Üáñ ÷³ëï³Ã., ÃÕÃ³å³Ý³ÏÝ»ñ, Ñ³ßí»ïíáõÃÛáõÝÝ»ñ",.DocNum,"010118","010121")
    
    Call GroupDelete(FolderName, .DocLevel, .DocNum, "^A[Del]", "^A[Del]", "^A[Del]")
    Call LetterOfCredit_Filter_Fill(FolderName, .DocLevel, .DocNum)
    Call Delete_ViewEdit("^A[Del]", "^A[Del]", c_Other & "|" & c_Limits)
    
    BuiltIn.Delay(2000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_Delete)
    Call ClickCmdButton(3, "²Ûá")
    
    Call Close_AsBank()
  End With  
End Sub