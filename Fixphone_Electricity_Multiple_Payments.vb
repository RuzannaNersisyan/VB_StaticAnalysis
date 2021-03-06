Option Explicit

'USEUNIT Comunal_Library
'USEUNIT Library_Common  
'USEUNIT SWIFT_International_Payorder_Library
'USEUNIT Subsystems_SQL_Library
'USEUNIT OLAP_Library
'USEUNIT Constants

'Test Case Id 165872

Sub Fixphone_Payments_Test()
  Dim fDATE, sDATE, attr, FolderName, FrmSpr, FrmMsgBox, StrToFind, ActualMessage, ExpectedMessage
  Dim QueryString, ExpSQLValue, ColNum, SQL_IsEqual, Amount_1, Debt_1, AmountUSD_1, Amount_2, Debt_2, AmountUSD_2      
  Dim CommunalPayment
  Dim arrIgnore
  
  'Համակարգ մուտք գործել ADMIN օգտագործողով
  fDATE = "20240101"
  sDATE = "20140101"
  Call Initialize_AsBankQA(sDATE, fDATE)
  Login("ADMIN")
  Call Create_Connection()
  
'--------------------------------------
  Set attr = Log.CreateNewAttributes
  attr.BackColor = RGB(0, 255, 255)
  attr.Bold = True
  attr.Italic = True
'--------------------------------------
   
  ' Ջնջել գոյություն ունեցող ֆայլերը 
  If aqFileSystem.Exists(Project.Path & "Stores\Communal\Actual\Fixphone\Actual Message.txt") Then
    Call aqFileSystem.DeleteFile(Project.Path & "Stores\Communal\Actual\Fixphone\Actual Message.txt")
  End If

  FolderName = "|Ð³×³Ëáñ¹Ç ëå³ë³ñÏáõÙ ¨ ¹ñ³Ù³ñÏÕ |"
  Call ChangeWorkspace(c_CustomerService)
  
  Call Log.Message("Կոմունալ վճարումներ փաստաթղթի ստեղծում",,,attr)
  Set CommunalPayment = New_CommunalPaymentDoc()
  With CommunalPayment
    .Date = aqConvert.DateTimeToStr(aqDateTime.Today)
    .arrayServicesToBePaid = Array(1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    .Tel = "111111"
    
    Call .CreateComPay(FolderName & "²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ")

    Amount_1 = 1470.00
    Debt_1 = 1461.00
    AmountUSD_1 = 3.5427
    Amount_2 = 500.00
    Debt_2 = 500.00
    AmountUSD_2 = 1.205
    
       'SQL ստուգում Կոմունալ վճարումների հանձնարարագիրը ստեղցելուց հետո: 
          'HI
          QueryString = "SELECT COUNT(*) FROM HI WHERE fBASE = " & .fBASE &_
                          "AND (fSUM = " & Amount_1 & " OR fSUM = " & Amount_2 & ") AND (fCURSUM = " & Amount_1 & " OR fCURSUM = " &_
                          Amount_2 & ")"
          ExpSQLValue = 4
          ColNum = 0
          SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, ColNum)
          If Not SQL_IsEqual Then
            Log.Error("Expected result = " & ExpSQLValue)
          End If  
         
    'Վավերացնել Կոմունալ վճարումների հանձնարարագիրը
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_ToConfirm)
    Call ClickCmdButton(1, "Ð³ëï³ï»É")
    wMDIClient.VBObject("frmPttel").Close
    
       'SQL ստուգում Կոմունալ վճարումների հանձնարարագիրը Վավերացնելուց հետո: 
          'COM_PAYMENTS
          QueryString = "SELECT COUNT(*) FROM COM_PAYMENTS WHERE fISN = " & .fBASE &_
                         "AND (fAMOUNT = " & Amount_1 & " OR fAMOUNT = " & Amount_2 & ") AND (fDEBT = " & Debt_1 & " OR fDEBT = " &_
                         Debt_2 &  ") AND fEXPDATE IS NULL"
          ExpSQLValue = 2
          ColNum = 0
          SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, ColNum)
          If Not SQL_IsEqual Then
            Log.Error("Expected result = " & ExpSQLValue)
          End If  
          
          'PAYMENTS
          QueryString = "SELECT COUNT(*) FROM PAYMENTS WHERE fISN = " & .fBASE &_
                         "AND (fSUMMA = " & Amount_1 & " OR fSUMMA = " & Amount_2 & ") AND (fSUMMAAMD = " & Amount_1 &_
                         " OR fSUMMAAMD = " & Amount_2 & ") AND (fSUMMAUSD = " & AmountUSD_1 & " OR fSUMMAUSD = " & AmountUSD_2 & ")"
          ExpSQLValue = 2
          ColNum = 0
          SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, ColNum)
          If Not SQL_IsEqual Then
            Log.Error("Expected result = " & ExpSQLValue)
          End If  
    
    Call ChangeWorkspace(c_ComPay)
    wTreeView.DblClickItem("|ÎáÙáõÝ³É í×³ñáõÙÝ»ñÇ ²Þî|ÎáÙáõÝ³É í×³ñáõÙÝ»ñ")
    Call Rekvizit_Fill("Dialog", 1, "General", "DSDATE", .Date)
    Call Rekvizit_Fill("Dialog", 1, "General", "DEDATE", .Date)
    Call Rekvizit_Fill("Dialog", 1, "General", "DISN", .fBASE)
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    
    If wMDIClient.VBObject("frmPttel").VBObject("tdbgView").ApproxCount <> 2 Then
      Log.Error("Կոմունալ վճարումների հանձնարարագրերը առկա չեն Կոմունալ վճարումներ թղթապանակում")
      Exit Sub
    End If
    
    Call Log.Message("Խմբային արտահանել կոմունալ վճարումների հանձնարարագրերը",,,attr)
    Call wMainForm.MainMenu.Click(c_Editor)
    Call wMainForm.PopupMenu.Click(c_MarkAll)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_ExportData)
    BuiltIn.Delay(5000)

    Call wMainForm.MainMenu.Click("Պատուհաններ|2  Տվյալների արտահանման սխալներ")

    Set FrmSpr = wMDIClient.WaitVbObject("FrmSpr", 2000)
    
    If FrmSpr.Exists Then 
      FrmSpr.SetFocus
      'Սեղմել "Հիշել որպես"
      Call wMainForm.MainMenu.Click(c_SaveAs)
      ActualMessage = Project.Path & "Stores\Communal\Actual\Fixphone\Actual Message.txt"
      ExpectedMessage = Project.Path & "Stores\Communal\Expected\Fixphone Electricity Expected Message.txt"
      Sys.Process("Asbank").Window("#32770", "ÐÇß»É áñå»ë", 1).Window("DUIViewWndClassName", "", 1).Window("DirectUIHWND", "", 1).Window("FloatNotifySink", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1).Keys(ActualMessage)
      Sys.Process("Asbank").Window("#32770", "ÐÇß»É áñå»ë", 1).Window("Button", "&Save", 1).Click()
    
      Call Compare_Files(ExpectedMessage, ActualMessage, "")
      FrmSpr.Close
    Else  
      Call Log.Error("Մշակված/Չմշակված տողերի մասին հաղորդագրությունը չի հայտնվել") 
    End If
    
       'SQL ստուգում Կոմունալ վճարումների հանձնարարագիրը Արտահանելուց հետո: 
          'COM_PAYMENTS
          QueryString = "SELECT COUNT(*) FROM COM_PAYMENTS WHERE fISN = " & .fBASE &_
                         "AND (fAMOUNT = " & Amount_1 & " OR fAMOUNT = " & Amount_2 & ") AND (fDEBT = " & Debt_1 & " OR fDEBT = " &_
                         Debt_2 & ") AND NOT fEXPDATE IS NULL"
          ExpSQLValue = 2
          ColNum = 0
          SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, ColNum)
          If Not SQL_IsEqual Then
            Log.Error("Expected result = " & ExpSQLValue)
          End If  
          
          'HI
          QueryString = "SELECT COUNT(*) FROM HI WHERE fBASE = " & .fBASE &_
                         "AND (fSUM = " & Amount_1 & " OR fSUM = " & Amount_2 & ") AND (fCURSUM = " & Amount_1 & " OR fCURSUM = " &_
                          Amount_2 & ")"
          ExpSQLValue = 4
          ColNum = 0
          SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, ColNum)
          If Not SQL_IsEqual Then
            Log.Error("Expected result = " & ExpSQLValue)
          End If  
    
    Call Log.Message("Մարել (հավաքական) գործողության կատարում",,,attr)
    Call ComunalGroupRepay(.Date)    
    
    Call Log.Message("Բոլոր փաստաթղթերի ջնջում",,,attr)
    
    Call aqFileSystem.DeleteFile(ActualMessage)
    'Ջնջել փաստաթուղթը COM_PAYMENTS աղյուսակից
    queryString = "DELETE FROM COM_PAYMENTS WHERE fISN = " & .fBASE
    Call Execute_SLQ_Query(queryString)
    BuiltIn.Delay(2000)
    Call wMainForm.MainMenu.Click(c_ToRefresh)
    BuiltIn.Delay(2000)
    If wMDIClient.VBObject("frmPttel").VBObject("tdbgView").ApproxCount <> 0 Then
      Log.Error("Փաստաթուղթը չի ջնջվել COM_PAYMENTS աղյուսակից")
    End If
    
    wMDIClient.VBObject("frmPttel").Close
    Call ChangeWorkspace(c_CustomerService)
    wTreeView.DblClickItem(FolderName & "Ð³ßí³éí³Í í×³ñ³ÛÇÝ ÷³ëï³ÃÕÃ»ñ")
    
    'Ջնջել Խմբային հիշարար օրդեր փաստաթուղթը
    Call Rekvizit_Fill("Dialog", 1, "General", "PERN", "^A[Del]" &.Date) 
    Call Rekvizit_Fill("Dialog", 1, "General", "PERK", "^A[Del]" &.Date) 
    Call Rekvizit_Fill("Dialog", 1, "General", "DOCTYPE", "^A[Del]" & "CmMOrdPk")
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    
    Builtin.Delay(2000)
    If wMDIClient.VBObject("frmPttel").VBObject("tdbgView").ApproxCount <> 1 Then
      Log.Error("'Հաշվառված վճարային փաստաթղթեր' թղթապանակում 'Խմբային հիշարար օրդեր' տեսակի փաստաթղթերի քանակը 1 չէ:")
    Else  
      Call wMainForm.MainMenu.Click(c_AllActions)
      Call wMainForm.PopupMenu.Click(c_Delete)
      Call ClickCmdButton(5, "Î³ï³ñ»É")
      Call ClickCmdButton(3, "²Ûá")
    End If
    
    wMDIClient.VBObject("frmPttel").Close
    
    wTreeView.DblClickItem(FolderName & "Ð³ßí³éí³Í í×³ñ³ÛÇÝ ÷³ëï³ÃÕÃ»ñ")
    
    'Ջնջել Կոմունալ վճարումների հանձնարարագիրը
    Call Rekvizit_Fill("Dialog", 1, "General", "PERN", "^A[Del]" &.Date) 
    Call Rekvizit_Fill("Dialog", 1, "General", "PERK", "^A[Del]" &.Date) 
    Call Rekvizit_Fill("Dialog", 1, "General", "DOCISN", .fBASE)
    Call ClickCmdButton(2, "Î³ï³ñ»É")    

    BuiltIn.Delay(2000)
    If wMDIClient.VBObject("frmPttel").VBObject("tdbgView").ApproxCount <> 1 Then
      Log.Error("Կոմունալ վճարումների հանձնարարագիրը առկա չէ 'Հաշվառված վճարային փաստաթղթեր' թղթապանակում:")
      Exit Sub
    Else  
      Call wMainForm.MainMenu.Click(c_AllActions)
      Call wMainForm.PopupMenu.Click(c_Delete)
      Call ClickCmdButton(5, "Î³ï³ñ»É")
      Call ClickCmdButton(3, "²Ûá")
    End If
  End With
    
  Call Close_AsBank()  
End Sub