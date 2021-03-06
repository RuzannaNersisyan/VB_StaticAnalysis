Option Explicit

'USEUNIT Library_Common  
'USEUNIT Factoring_Library
'USEUNIT Payment_Order_ConfirmPhases_Library
'USEUNIT Akreditiv_Library
'USEUNIT Loan_Agreements_Library
'USEUNIT Overdraft_NewCases_Library
'USEUNIT Group_Operations_Library
'USEUNIT Subsystems_SQL_Library 
'USEUNIT Constants

'Test Case Id 165855

Sub Factoring_With_Schedule_BlackList_Test()
  Dim attr, fDATE, sDATE, FolderName, queryString, sqlValue, colNum, sql_isEqual
  Dim CollectFromProvision_ISN, GiveFactoring_ISN, CalculatePercents_ISN, FadeDebt_ISN       
  Dim Agr, PerAcc, calcDate, opDate, ExpectedSum, Exists, Dissum, AccumRem, EndDate,_
      dateFill, OldPerSum, NewPerSum, summa
  Dim name, name_len, Pttel    
'--------------------------------------
  Set attr = Log.CreateNewAttributes
  attr.BackColor = RGB(0, 255, 255)
  attr.Bold = True
  attr.Italic = True
'--------------------------------------  

  'Համակարգ մուտք գործել ARMSOFT օգտագործողով
  fDATE = "20260101"
  sDATE = "20140101"
  Call Initialize_AsBank("bank", sDATE, fDATE)
  Login("ARMSOFT")
  Call Create_Connection()
  
  Call ChangeWorkspace(c_Factoring)
  FolderName = "|ØØÄä ü³ÏïáñÇÝ·|"
  
  Call Log.Message("Ֆակտորինգի պայմանագրի ստեղծում",,,attr)
  Set Agr = New_FactoringDoc()
  With Agr
    .PayerAcc = "33170160500"
    .LenderAcc = "00000113032"
    .Amount = 100000
    .Date = "220419" 
    .GiveDate = "220419"
    .Term = "220420"
    .DocLevel = 1
    .PaidAmount = 100000
    .PaperCode = 333
    .DocType = "¶ñ³ýÇÏáí ý³ÏïáñÇÝ·Ç å³ÛÙ³Ý³·Çñ"
    .DocTypeNum = "8"
    
    Call .CreateFactoring(FolderName & "Üáñ å³ÛÙ³Ý³·ñÇ ëï»ÕÍáõÙ")
    
    Log.Message("Գրաֆիկով ֆակտորինգի պայմանագրի համարը  և ISN-ը`")
    Log.Message(.DocNum)
    Log.Message(.fBASE)
    
    ''SQL ստուգում պայամանգիր ստեղցելուց հետո: 
          ''CONTRACTS
          queryString = "SELECT COUNT(*) FROM CONTRACTS WHERE fDGISN = " & .fBASE &_
                          "AND fDGAGRTYPE = 'C' AND fDGMODTYPE = 5 " &_
                          "AND fDGAGRKIND = '8' AND fDGSTATE = 206 " &_
                          "AND fDGSUMMA = 100000.00 AND fDGALLSUMMA = 100000.00"
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
          
    'Մարման գրաֆիկի նշանակում
    Builtin.Delay(2000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_RepaySchedule)          
    
    wMDIClient.VBObject("frmPttel").VBObject("tdbgView").MoveNext      
    'Այլ վճարումների գրաֆիկի նշանակում
    Builtin.Delay(2000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_OtherPaySchedule)
    Call ClickCmdButton(1, "Î³ï³ñ»É")
    wMDIClient.VBObject("frmPttel").VBObject("tdbgView").MovePrevious
    
    'Պայմանագրը ուղարկել հաստատման                               
    Call PaySys_Send_To_Verify()
    Builtin.Delay(2000)
    wMDIClient.VBObject("frmPttel").Close
    
    'Պայմանագիրը հաստատել "<<Սև ցուցակ>> հաստատողի ԱՇՏ" - ում:
    Call ChangeWorkspace(c_BLVerifyer)
    Call wTreeView.DblClickItem("|§ê¨ óáõó³Ï¦ Ñ³ëï³ïáÕÇ ²Þî|Ð³ëï³ïíáÕ ï»Õ³µ³ßËí³Í ÙÇçáóÝ»ñ ¨ »ñ³ßË³íáñáõÃÛáõÝÝ»ñ")
	Call Rekvizit_Fill("Dialog", 1, "General", "SUBSYS", "C5") 
    Call Rekvizit_Fill("Dialog", 1, "General", "NUM", .DocNum) 
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    Builtin.Delay(2000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_ToConfirm)
    Call ClickCmdButton(1, "Ð³ëï³ï»É")
    Builtin.Delay(2000)
    wMDIClient.VBObject("frmPttel").Close
    
    Call ChangeWorkspace(c_Factoring)
    Call LetterOfCredit_Filter_Fill(FolderName, .DocLevel, .DocNum)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_Folders & "|" & c_AgrFolder)
     'Վերցնել "Տոկոսային եկամուտների հաշիվ" ռեկվիզիտի արժեքը`
     Builtin.Delay(2000)
      Call wMainForm.MainMenu.Click(c_AllActions)
      Call wMainForm.PopupMenu.Click(c_View)
  
      Set wMDIClient.VBObject("frmASDocForm").VBObject("TabStrip").SelectedItem = wMDIClient.VBObject("frmASDocForm").VBObject("TabStrip").Tabs(2)
      PerAcc = wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame_2").VBObject("AsTypeFolder_13").VBObject("TDBMask").Text
      wMDIClient.VBObject("frmASDocForm").Close
      Builtin.Delay(2000)
      wMDIClient.VBObject("frmPttel_2").Close
      Builtin.Delay(2000)
      wMDIClient.VBObject("frmPttel").Close
    
      'Մուտք գործել "Հաճախորդի սպասարկում և դրամարկղ" 
      Call ChangeWorkspace(c_CustomerService)
      Call wTreeView.DblClickItem("|Ð³×³Ëáñ¹Ç ëå³ë³ñÏáõÙ ¨ ¹ñ³Ù³ñÏÕ |Ð³ßÇíÝ»ñ")
      Call Rekvizit_Fill("Dialog", 1, "General", "ACCMASK", PerAcc)
      Call ClickCmdButton(2, "Î³ï³ñ»É") 
      'Փոխել Տոկոսային եկամուտների հաշիվի սոտորին սահմանը
      Builtin.Delay(2000)
      Call wMainForm.MainMenu.Click(c_AllActions)
      Call wMainForm.PopupMenu.Click(c_ChangeLowerBound)
      Call Rekvizit_Fill("Document", 1, "General", "CHGDATE", "220419")
      Call Rekvizit_Fill("Document", 1, "General", "LLIMIT", "-999999999")
      Call ClickCmdButton(1, "Î³ï³ñ»É") 
      Builtin.Delay(2000)
      wMDIClient.VBObject("frmPttel").Close
    
    Call ChangeWorkspace(c_Factoring)
    Call LetterOfCredit_Filter_Fill(FolderName, .DocLevel, .DocNum)
          
    Call Log.Message("Գանձում տրամադրումից",,,attr)
    Call Collect_From_Provision(.Date, summa, 2, .PayerAcc, CollectFromProvision_ISN)
    
        ''SQL ստուգում Գանձում տրամադրումից հետո: 
            ''CONTRACTS
            queryString = "SELECT fDGSTATE FROM CONTRACTS WHERE fDGISN = " & .fBASE 
            sqlValue = 7
            colNum = 0
            sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
            If Not sql_isEqual Then
              Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
            End If
          
            ''HI
            queryString = "SELECT COUNT(*) FROM HI WHERE fBASE = " & CollectFromProvision_ISN &_
                           "AND fSUM = 100.00 AND fCURSUM = 100.00"  
            sqlValue = 2
            colNum = 0
            sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
            If Not sql_isEqual Then
              Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
            End If
           
            
            ''HIR
            queryString = "SELECT fCURSUM FROM HIR WHERE fOBJECT = " & .fBASE  
            sqlValue = 100.00
            colNum = 0
            sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
            If Not sql_isEqual Then
              Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
            End If
            
            ''HIRREST
            queryString = "SELECT fLASTREM FROM HIRREST WHERE fOBJECT = " & .fBASE 
            sqlValue = 100.00
            colNum = 0
            sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
            If Not sql_isEqual Then
              Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
            End If

    Call Log.Message("ՄՄԺՊ ֆակտորինգի տրամադրում",,,attr)
    GiveFactoring_ISN = GiveFactoring(.Date, 2, Null)  

          ''SQL ստուգում ֆակտորինգի տրամադրումից հետո: 
            ''HI
            BuiltIn.Delay(2000)
            queryString = "SELECT COUNT(*) FROM HI WHERE fBASE = " & GiveFactoring_ISN &_
                           "AND ((fSUM = 100.00 AND fCURSUM = 100.00) OR (fSUM = 100000.00 AND fCURSUM = 100000.00))" 	&_		
													 "AND fTYPE = '01'"  
            sqlValue = 6
            colNum = 0
            sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
            If Not sql_isEqual Then
              Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
            End If  
            
            ''HI2
            queryString = "SELECT COUNT(*) FROM HI2 WHERE fBASE = " & GiveFactoring_ISN &_
                           "AND fSUM = 100000.00 AND fCURSUM = 100000.00"  
            sqlValue = 1
            colNum = 0
            sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
            If Not sql_isEqual Then
              Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
            End If 
            
            ''HIF
            queryString = "SELECT COUNT(*) FROM HIF WHERE fBASE = " & GiveFactoring_ISN &_
                           "AND fSUM = 0.00 AND fCURSUM = 0.00"  
            sqlValue = 1
            colNum = 0
            sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
            If Not sql_isEqual Then
              Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
            End If 
            
            ''HIR
            queryString = "SELECT COUNT(*) FROM HIR WHERE fBASE = " & GiveFactoring_ISN &_
                           "AND ((fCURSUM = 100000.00 AND fOP = 'AGR')" &_
													 "OR (fCURSUM = 100.00 AND fOP = 'PAY'))"  
            sqlValue = 2
            colNum = 0
            sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
            If Not sql_isEqual Then
              Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
            End If
            
            ''HIREST2
            queryString = "SELECT COUNT(*) FROM HIREST2 WHERE fOBJECT = " & .fBASE &_
                           "AND ((fREM = 0.00 AND fCURREM = 0.00)" &_
													 "OR (fREM = 100000.00 AND fCURREM = 100000.00))"  
            sqlValue = 2
            colNum = 0
            sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
            If Not sql_isEqual Then
              Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
            End If 
            
            ''HIRREST
            queryString = "SELECT COUNT(*) FROM HIRREST WHERE fOBJECT = " & .fBASE &_
                           "AND (fLASTREM = 100000.00 OR fLASTREM = 100.00" &_
                           "OR fLASTREM = -100.00) AND fPENULTREM = 0.00 AND fSTARTREM = 0.00"  
            sqlValue = 3
            colNum = 0
            sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
            If Not sql_isEqual Then
              Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
            End If 
        
    Call Log.Message("Տոկոսների հաշվարկ մարման նախորդ օրով",,,attr)
    calcDate = "210519"
    CalculatePercents_ISN = Calculate_Percents(calcDate, calcDate, False)
    
          ''SQL ստուգում Տոկոսների հաշվարկից հետո: 
            ''HI
            queryString = "SELECT COUNT(*) FROM HI WHERE fBASE = " & CalculatePercents_ISN &_
                           "AND ((fSUM = 986.30 AND fCURSUM = 986.30)" 	&_		
													 "OR (fSUM = 1001.00 AND fCURSUM = 1001.00))"  
            sqlValue = 4
            colNum = 0
            sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
            If Not sql_isEqual Then
              Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
            End If  
        
            ''HIF
            queryString = "SELECT COUNT(*) FROM HIF WHERE fBASE = " & CalculatePercents_ISN &_
                           "AND fSUM = 0.00 AND fCURSUM = 0.00"  
            sqlValue = 1
            colNum = 0
            sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
            If Not sql_isEqual Then
              Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
            End If  
            
            ''HIR
            queryString = "SELECT COUNT(*) FROM HIR WHERE fBASE = " & CalculatePercents_ISN &_
                           "AND (fCURSUM = 986.30 OR fCURSUM = 14.70 OR fCURSUM = 8333.30)"  
            sqlValue = 4
            colNum = 0
            sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
            If Not sql_isEqual Then
              Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
            End If  
            
            ''HIRREST
            queryString = "SELECT COUNT(*) FROM HIRREST WHERE fOBJECT = " & .fBASE
            sqlValue = 6
            colNum = 0
            sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
            If Not sql_isEqual Then
              Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
            End If  
            
            ''HIT
            queryString = "SELECT COUNT(*) FROM HIT WHERE fBASE = " & CalculatePercents_ISN &_
                           "AND (fCURSUM = 986.30 OR fCURSUM = 14.70)"
            sqlValue = 2
            colNum = 0
            sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
            If Not sql_isEqual Then
              Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
            End If
    
    'Կատարել պարտքերի մարում` նախատեսվածից ավելի Տոկոսագումար մարելով:
    Call Log.Message("Պարտքերի մարում",,,attr)
    opDate = "220519"
    Call Fade_Debt(opDate, FadeDebt_ISN, "", "", 1000, False)
    
          ''SQL ստուգում Պարտքերի մարումից հետո: 
            ''HI
            BuiltIn.Delay(1000)
            queryString = "SELECT COUNT(*) FROM HI WHERE fBASE = " & FadeDebt_ISN &_
                           "AND ((fSUM = 8333.30 AND fCURSUM = 8333.30) " &_
                           "OR (fSUM = 1000.00 AND fCURSUM = 1000.00))" 
            sqlValue = 4
            colNum = 0
            sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
            If Not sql_isEqual Then
              Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
            End If  
        
            ''HI2
            queryString = "SELECT COUNT(*) FROM HI2 WHERE fBASE = " & FadeDebt_ISN &_
                           "AND fSUM = 8333.30 AND fCURSUM = 8333.30"
            sqlValue = 1
            colNum = 0
            sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
            If Not sql_isEqual Then
              Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
            End If 
            
            ''HIR
            queryString = "SELECT COUNT(*) FROM HIR WHERE fBASE = " & FadeDebt_ISN &_
                           "AND (fCURSUM = 8333.30 OR fCURSUM = 1000.00 OR fCURSUM = 986.30)"
            sqlValue = 4
            colNum = 0
            sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
            If Not sql_isEqual Then
              Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
            End If 
            
            ''HIRREST
            queryString = "SELECT SUM(fLASTREM) FROM HIRREST WHERE fOBJECT = " & .fBASE 
            sqlValue = 91667.70
            colNum = 0
            sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
            If Not sql_isEqual Then
              Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
            End If
            
            ''HIREST2
            queryString = "SELECT COUNT(*) FROM HIREST2 WHERE fOBJECT = " & .fBASE &_
                           "AND ((fREM = 0.00 AND fCURREM = 0.00) " &_
                           "OR (fREM = 91666.70 AND fCURREM = 91666.70))"
            sqlValue = 2
            colNum = 0
            sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
            If Not sql_isEqual Then
              Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
            End If
    
    'Ստուգել "Հաշվ.% մնացորդ" դաշտի արժեքը բացասական լինի
    AccumRem = wMDIClient.VBObject("frmPttel").VBObject("tdbgView").Columns.Item(4) 
    If AccumRem >= 0 Then
      Call Log_Error("Հաշվ.% մնացորդ դաշտի արժեքը բացասական չէ:")
    End If
    
    'Կանխավ վճարված տոկոսների վերադարձ
    Call ReturnPrepaidRates(opDate, "")
    
    'Ստուգել "Հաշվ.% մնացորդ" դաշտի արժեքը 0 լինի
    AccumRem = wMDIClient.VBObject("frmPttel").VBObject("tdbgView").Columns.Item(4) 
    If AccumRem <> 0 Then
      Call Log_Error("Հաշվ.% մնացորդ դաշտի արժեքը 0 չէ:")
    End If
    
    'Մարման գրաֆիկի վերանայում
    EndDate = "220421"
    dateFill = 1
    Call PaymentScheduleReview(opDate, EndDate, "", dateFill, "", "")

    'Ստուգել, որ Տոկոսագումարը մեծացած լինի լինի:
    OldPerSum = 5553.50
    
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_Folders & "|" & c_CurrentSchedules)
    NewPerSum = wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame").VBObject("DocGrid").Columns.Item(4).FooterText
    
    If CInt(NewPerSum) <= OldPerSum Then
      Log.Error("Հաշվարկված տոկոսագումարը սխալ է:")
    End If
    Builtin.Delay(2000)
    wMDIClient.VBObject("frmASDocForm").Close
    
    Call Log.Message("Բոլոր փաստաթղթերի ջնջում",,,attr)
    'Ջնջել "Գրաֆիկի վերանայումը"
    Builtin.Delay(2000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_Folders & "|" & c_AgrFolder)
    name = "Ø³ñáõÙÝ»ñÇ ·ñ³ýÇÏ`  22/05/19"
    name_len = 28
    ColNum = 0
    Pttel = "_2"
    Call Find_Doc_By(name, name_len,ColNum, Pttel)
    
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_Delete)
    Call ClickCmdButton(3, "²Ûá")
    Builtin.Delay(2000)
    wMDIClient.VBObject("frmPttel_2").Close

    'Ջնջում "Գործողությունների դիտումից"
    Builtin.Delay(2000)
    wMDIClient.VBObject("frmPttel").Close
    Call GroupDelete(FolderName, 1, .DocNum, "^A[Del]", "^A[Del]", "^A[Del]") 
    
    Call LetterOfCredit_Filter_Fill(FolderName, .DocLevel, .DocNum)
    Builtin.Delay(2000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_Delete)
    Call ClickCmdButton(3, "²Ûá")
  End With    
  Call Close_AsBank()
End Sub