Option Explicit

'USEUNIT Comunal_Library
'USEUNIT Library_Common  
'USEUNIT SWIFT_International_Payorder_Library
'USEUNIT Subsystems_SQL_Library
'USEUNIT OLAP_Library
'USEUNIT Constants

'Test Case Id 165870
'Test Case Id 165878

Sub Gas_Multiple_Payments_Test(Online)
  Dim fDATE, sDATE, attr, FolderName, FrmSpr, FrmMsgBox, DBFForCompare, ExcelForCompare, ExportedFile, ExportedFileToTxt, ExportedExcel,_
      year, day, month, param, StrToFind, Service, TXTFileName, ActualMessage, ExpectedMessage
  Dim QueryString, ExpSQLValue, ColNum, SQL_IsEqual, Amount_1, Amount_2, Debt_1, Debt_2, AmountUSD_1, AmountUSD_2      
  Dim CommunalPayment_1, CommunalPayment_2
  Dim arrIgnore
  Dim wndNotepad, progress, comboBox, edit, ExportPath
  Dim fso, objFolder, obj
  
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

  DBFForCompare = Project.Path & "Stores\Communal\Expected\Gas.txt"
  ExcelForCompare = Project.Path & "Stores\Communal\Expected\Gas.xls"
  ExportPath = Project.Path & "Stores\Communal\Actual\Gas"
  
  ' Ջնջել գոյություն ունեցող ֆայլերը  
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set objFolder = fso.GetFolder(ExportPath)
  For Each obj in objFolder.Files
      If aqString.StrMatches("11.*000000.dbf", obj.Name) Then
         Call aqFileSystem.DeleteFile(ExportPath & "\" & obj.Name)
      ElseIf aqString.StrMatches("11.*000000.txt", obj.Name) Then
         Call aqFileSystem.DeleteFile(ExportPath & "\" & obj.Name)
      ElseIf aqString.StrMatches("11.*000000.xls", obj.Name) Then
         Call aqFileSystem.DeleteFile(ExportPath & "\" & obj.Name)
      End If
  Next
  If aqFileSystem.Exists(ExportPath & "\Actual Message.txt") Then
    Call aqFileSystem.DeleteFile(ExportPath & "\Actual Message.txt")
  End If

  'Կոմունալ ծառայությունների կարգավորումներում "G ծառայության օպերատորի արժեքը offline դեպքում պետք է լինի "O" 
  Call ChangeWorkspace(c_ComPay)
  wTreeView.DblClickItem("|ÎáÙáõÝ³É í×³ñáõÙÝ»ñÇ ²Þî|ÎáÙáõÝ³É Í³é³ÛáõÃÛáõÝÝ»ñÇ Ï³ñ·³íáñáõÙ")
  With wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame").VBObject("DocGrid")
    .Row = 3
    .Col = 1
    If Online Then
      '"G" ծառայության օպերատորը լրացնել "I" 
      .Keys("I" & "[Tab]")
      Amount_1 = 3220.00
      Amount_2 = 3220.00
      Debt_1 = 3218.00
      Debt_2 = 3218.00
      AmountUSD_1 = 7.7602
      AmountUSD_2 = 7.7602
      Call ClickCmdButton(1, "Î³ï³ñ»É")
    Else   
      .Keys("O" & "[Tab]")
      Amount_1 = 8210.00
      Amount_2 = 12930.00
      Debt_1 = 8200.14
      Debt_2 = 12926.92
      AmountUSD_1 = 19.786
      AmountUSD_2 = 31.1611
      Call ClickCmdButton(1, "Î³ï³ñ»É")
      
      'Արտահնման ճանապարհի կարգավորում
      wTreeView.DblClickItem("|ÎáÙáõÝ³É í×³ñáõÙÝ»ñÇ ²Þî|¶³½Ç í×³ñáõÙÝ»ñ|¶³½Ç í×³ñáõÙÝ»ñÇ Ï³ñ·³íáñáõÙÝ»ñ")
      Call Rekvizit_Fill("Document", 1, "General", "OUTDIR", "^A[Del]" & ExportPath) 
      Call ClickCmdButton(1, "Î³ï³ñ»É")
    End If
  End With

  FolderName = "|Ð³×³Ëáñ¹Ç ëå³ë³ñÏáõÙ ¨ ¹ñ³Ù³ñÏÕ |"
  Call ChangeWorkspace(c_CustomerService)

  Call Log.Message("Առաջին Կոմունալ վճարումներ փաստաթղթի ստեղծում",,,attr)
  Set CommunalPayment_1 = New_CommunalPaymentDoc()
  With CommunalPayment_1
    .Date = aqConvert.DateTimeToStr(aqDateTime.Today)
    .arrayServicesToBePaid = Array(0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0)
    .Tel = "327463"
    Call .CreateComPay(FolderName & "²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ")
  
       'SQL ստուգում Կոմունալ վճարումների հանձնարարագիրը ստեղցելուց հետո: 
          'HI
          QueryString = "SELECT COUNT(*) FROM HI WHERE fBASE = " & .fBASE &_
                          "AND fSUM = " & Amount_1 & " AND fCURSUM = " & Amount_1 & " AND fTRANS =0"
          ExpSQLValue = 2
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
                         "AND fAMOUNT = " & Amount_1 & " AND fDEBT = " & Debt_1 & " AND fEXPDATE IS NULL"
          ExpSQLValue = 1
          ColNum = 0
          SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, ColNum)
          If Not SQL_IsEqual Then
            Log.Error("Expected result = " & ExpSQLValue)
          End If  
          
          'PAYMENTS
          QueryString = "SELECT COUNT(*) FROM PAYMENTS WHERE fISN = " & .fBASE &_
                         "AND fSUMMA = " & Amount_1 & " AND fSUMMAAMD = " & Amount_1 & " AND fSUMMAUSD = " & AmountUSD_1
          ExpSQLValue = 1
          ColNum = 0
          SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, ColNum)
          If Not SQL_IsEqual Then
            Log.Error("Expected result = " & ExpSQLValue)
          End If  
  End With 
      
  Call Log.Message("Երկրորդ Կոմունալ վճարումներ փաստաթղթի ստեղծում",,,attr)
  Set CommunalPayment_2 = New_CommunalPaymentDoc()
  With CommunalPayment_2
    .Date = aqConvert.DateTimeToStr(aqDateTime.Today)
    .arrayServicesToBePaid = Array(0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0)
    .Tel = "326820"
    Call .CreateComPay(FolderName & "²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ")
  
       'SQL ստուգում Կոմունալ վճարումների հանձնարարագիրը ստեղցելուց հետո: 
          'HI
          QueryString = "SELECT COUNT(*) FROM HI WHERE fBASE = " & .fBASE &_
                          "AND fSUM = " & Amount_2 & " AND fCURSUM = " & Amount_2 & " AND fTRANS = 0"
          ExpSQLValue = 2
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
                         "AND fAMOUNT = " & Amount_2 & " AND fDEBT = " & Debt_2 & " AND fEXPDATE IS NULL"
          ExpSQLValue = 1
          ColNum = 0
          SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, ColNum)
          If Not SQL_IsEqual Then
            Log.Error("Expected result = " & ExpSQLValue)
          End If  
          
          'PAYMENTS
          QueryString = "SELECT COUNT(*) FROM PAYMENTS WHERE fISN = " & .fBASE &_
                         "AND fSUMMA = " & Amount_2 & " AND fSUMMAAMD = " & Amount_2 & " AND fSUMMAUSD = " & AmountUSD_2
          ExpSQLValue = 1
          ColNum = 0
          SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, ColNum)
          If Not SQL_IsEqual Then
            Log.Error("Expected result = " & ExpSQLValue)
          End If  
  End With 
  
  If Online Then
    Service = "I004"
  Else
    Service = "G"  
  End If
  
  Call ChangeWorkspace(c_ComPay)
  wTreeView.DblClickItem("|ÎáÙáõÝ³É í×³ñáõÙÝ»ñÇ ²Þî|ÎáÙáõÝ³É í×³ñáõÙÝ»ñ")
  Call Rekvizit_Fill("Dialog", 1, "General", "DSDATE", CommunalPayment_1.Date)
  Call Rekvizit_Fill("Dialog", 1, "General", "DEDATE", CommunalPayment_1.Date)
  Call Rekvizit_Fill("Dialog", 1, "General", "TYPE", Service)
  Call ClickCmdButton(2, "Î³ï³ñ»É")
    
  If wMDIClient.VBObject("frmPttel").VBObject("tdbgView").ApproxCount <> 2 Then
    Log.Error("Կոմունալ վճարումների հանձնարարագրերը առկա չեն Կոմունալ վճարումներ թղթապանակում")
  '  Exit Sub
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
    ActualMessage = Project.Path & "Stores\Communal\Actual\Gas\Actual Message.txt"
    ExpectedMessage = Project.Path & "Stores\Communal\Expected\Gas Expected Message.txt"
    Sys.Process("Asbank").Window("#32770", "ÐÇß»É áñå»ë", 1).Window("DUIViewWndClassName", "", 1).Window("DirectUIHWND", "", 1).Window("FloatNotifySink", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1).Keys(ActualMessage)
    Sys.Process("Asbank").Window("#32770", "ÐÇß»É áñå»ë", 1).Window("Button", "&Save", 1).Click()
    
    Call Compare_Files(ExpectedMessage, ActualMessage, "")
    FrmSpr.Close
  Else  
    Call Log.Error("Մշակված/Չմշակված տողերի մասին հաղորդագրությունը չի հայտնվել") 
  End If

       'SQL ստուգում Կոմունալ վճարումների հանձնարարագիրը Արտահանելուց հետո: 
          'COM_PAYMENTS
          QueryString = "SELECT COUNT(*) FROM COM_PAYMENTS WHERE fISN = " & CommunalPayment_1.fBASE &_
                          "AND fAMOUNT = " & Amount_1 & " AND fDEBT = " & Debt_1 & " AND NOT fEXPDATE IS NULL"
          ExpSQLValue = 1
          ColNum = 0
          SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, ColNum)
          If Not SQL_IsEqual Then
            Log.Error("Expected result = " & ExpSQLValue)
          End If  
          
          'HI
          QueryString = "SELECT COUNT(*) FROM HI WHERE fBASE = " & CommunalPayment_1.fBASE &_
                          "AND fSUM = " & Amount_1 & " AND fCURSUM = " & Amount_1 & " AND fTRANS = 0"
          ExpSQLValue = 2
          ColNum = 0
          SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, ColNum)
          If Not SQL_IsEqual Then
            Log.Error("Expected result = " & ExpSQLValue)
          End If  

       'SQL ստուգում Կոմունալ վճարումների հանձնարարագիրը Արտահանելուց հետո: 
          'COM_PAYMENTS
          QueryString = "SELECT COUNT(*) FROM COM_PAYMENTS WHERE fISN = " & CommunalPayment_2.fBASE &_
                          "AND fAMOUNT = " & Amount_2 & " AND fDEBT = " & Debt_2 & " AND NOT fEXPDATE IS NULL"
          ExpSQLValue = 1
          ColNum = 0
          SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, ColNum)
          If Not SQL_IsEqual Then
            Log.Error("Expected result = " & ExpSQLValue)
          End If  
          
          'HI
          QueryString = "SELECT COUNT(*) FROM HI WHERE fBASE = " & CommunalPayment_2.fBASE &_
                          "AND fSUM = " & Amount_2 & " AND fCURSUM = " & Amount_2 & " AND fTRANS = 0"
          ExpSQLValue = 2
          ColNum = 0
          SQL_IsEqual = CheckDB_Value(QueryString, ExpSQLValue, ColNum)
          If Not SQL_IsEqual Then
            Log.Error("Expected result = " & ExpSQLValue)
          End If  
                    
    
  If Not Online Then
    year = aqDateTime.GetYear(aqDateTime.Today)
    If aqDateTime.GetMonth(aqDateTime.Today) < 10 Then
      month = "0" & aqDateTime.GetMonth(aqDateTime.Today)
    Else
      month = aqDateTime.GetMonth(aqDateTime.Today)  
    End If
    If aqDateTime.GetDay(aqDateTime.Today) < 10 Then
      day = "0" & aqDateTime.GetDay(aqDateTime.Today)
    Else
      day = aqDateTime.GetDay(aqDateTime.Today)  
    End If

    ExportedFile = Project.Path & "Stores\Communal\Actual\Gas\11" & year & month & day & "000000.dbf"   
    ExportedFileToTxt = Project.Path & "Stores\Communal\Actual\Gas\11" & year & month & day & "000000.txt" 
    ExportedExcel = Project.Path & "Stores\Communal\Actual\Gas\11" & year & month & day & "000000.xls" 
    TXTFileName = "11" & year & month & day & "000000.txt"
    Call aqFileSystem.CopyFile(ExportedFile, ExportedFileToTxt)
    TestedApps.Notepad.Run
    Set wndNotepad = Sys.Process("notepad").Window("Notepad")
    Call wndNotepad.MainMenu.Click("File|Open...")
    Set progress = Sys.Process("notepad").Window("#32770", "Open").Window("WorkerW").Window("ReBarWindow32").Window("Address Band Root").Window("msctls_progress32")
    Call Sys.Process("notepad").Window("#32770", "Open").Window("WorkerW").Window("ReBarWindow32").Window("Address Band Root").Window("msctls_progress32").Window("ToolbarWindow32", "Address band toolbar").ClickItem(202, False)
    Sys.Process("notepad").Window("#32770", "Open", 1).Window("ComboBoxEx32", "", 1).Window("ComboBox", "", 1).Keys(ExportedFileToTxt)
    Sys.Process("notepad").Window("#32770", "Open", 1).Window("Button", "&Open", 1).ClickButton
    Call wndNotepad.MainMenu.Click("File|Save As...")
    Sys.Process("notepad").Window("#32770", "Save As", 1).Window("Button", "&Save", 1).ClickButton
    Sys.Process("notepad").Window("#32770", "Confirm Save As", 1).UIAObject("Confirm_Save_As").Window("CtrlNotifySink", "", 7).Window("Button", "&Yes", 1).ClickButton
    Sys.Process("notepad").Window("Notepad", TXTFileName & " - Notepad", 1).Close
    
    'dbf ֆայլերի համեմատում   
   param = "(........A)|(.......A)|(202.....)|(.........[/])"
    Call Compare_Files(DBFForCompare, ExportedFileToTxt, param) 
    
    'xls ֆայլերի համեմատում
    ReDim arrIgnore(1)          
    arrIgnore = Array("$E$5","$F$5")
    Call ComparisonTwoExcelFilesWithCheck(ExcelForCompare, ExportedExcel, arrIgnore)
  End If
  
  Call Log.Message("Մարել (հավաքական) գործողության կատարում",,,attr)
  Call ComunalGroupRepay(CommunalPayment_1.Date)
  
  Call Log.Message("Բոլոր փաստաթղթերի ջնջում",,,attr)
  'Ջնջել փաստաթուղթը COM_PAYMENTS աղյուսակից
  queryString = "DELETE FROM COM_PAYMENTS WHERE fISN = " & CommunalPayment_1.fBASE
  Call Execute_SLQ_Query(queryString)
  queryString = "DELETE FROM COM_PAYMENTS WHERE fISN = " & CommunalPayment_2.fBASE
  Call Execute_SLQ_Query(queryString)
  BuiltIn.Delay(2000)
  Call wMainForm.MainMenu.Click(c_ToRefresh)
  BuiltIn.Delay(2000)
  If wMDIClient.VBObject("frmPttel").VBObject("tdbgView").ApproxCount <> 0 Then
    Log.Error("Փաստաթուղթը չի ջնջվել COM_PAYMENTS աղյուսակից")
  End If
  
  If Not Online Then
    'Ջնջել արտահանված փաստաթուղթը (.dbf, .txt, .xls ֆայլերը)
    Call aqFileSystem.DeleteFile(ExportedFile)
    Call aqFileSystem.DeleteFile(ExportedFileToTxt)
    Call aqFileSystem.DeleteFile(ExportedExcel)
  End If
  
  Call aqFileSystem.DeleteFile(ActualMessage)
  
  wMDIClient.VBObject("frmPttel").Close
  Call ChangeWorkspace(c_CustomerService)
  wTreeView.DblClickItem(FolderName & "Ð³ßí³éí³Í í×³ñ³ÛÇÝ ÷³ëï³ÃÕÃ»ñ")
    
  'Ջնջել Խմբային հիշարար օրդեր փաստաթուղթը
  Call Rekvizit_Fill("Dialog", 1, "General", "PERN", "^A[Del]" & CommunalPayment_1.Date) 
  Call Rekvizit_Fill("Dialog", 1, "General", "PERK", "^A[Del]" & CommunalPayment_1.Date) 
  Call Rekvizit_Fill("Dialog", 1, "General", "DOCTYPE", "^A[Del]" & "CmMOrdPk")
  Call ClickCmdButton(2, "Î³ï³ñ»É")
    
  BuiltIn.Delay(2000)
  If wMDIClient.VBObject("frmPttel").VBObject("tdbgView").ApproxCount <> 1 Then
    Log.Error("'Հաշվառված վճարային փաստաթղթեր' թղթապանակում 'Խմբային հիշարար օրդեր' տեսակի փաստաթղթերի քանակը 1 չէ:")
  Else  
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_Delete)
    If Not Online Then
      Call ClickCmdButton(5, "Î³ï³ñ»É")    
    End If
    Call ClickCmdButton(3, "²Ûá")
  End If
    
  wMDIClient.VBObject("frmPttel").Close
    
  wTreeView.DblClickItem(FolderName & "Ð³ßí³éí³Í í×³ñ³ÛÇÝ ÷³ëï³ÃÕÃ»ñ")
  'Ջնջել առաջին Կոմունալ վճարումների հանձնարարագիրը
  Call Rekvizit_Fill("Dialog", 1, "General", "PERN", "^A[Del]" & CommunalPayment_1.Date) 
  Call Rekvizit_Fill("Dialog", 1, "General", "PERK", "^A[Del]" & CommunalPayment_1.Date) 
  Call Rekvizit_Fill("Dialog", 1, "General", "DOCISN", CommunalPayment_1.fBASE)
  Call ClickCmdButton(2, "Î³ï³ñ»É")    

  BuiltIn.Delay(2000)
  If wMDIClient.VBObject("frmPttel").VBObject("tdbgView").ApproxCount <> 1 Then
    Log.Error("Առաջին կոմունալ վճարումների հանձնարարագիրը առկա չէ 'Հաշվառված վճարային փաստաթղթեր' թղթապանակում:")
    Exit Sub
  Else  
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_Delete)
    Call ClickCmdButton(3, "²Ûá")
  End If
  wMDIClient.VBObject("frmPttel").Close
  
  BuiltIn.Delay(2000)
  wTreeView.DblClickItem(FolderName & "Ð³ßí³éí³Í í×³ñ³ÛÇÝ ÷³ëï³ÃÕÃ»ñ")
  'Ջնջել երկրորդ Կոմունալ վճարումների հանձնարարագիրը
  Call Rekvizit_Fill("Dialog", 1, "General", "PERN", "^A[Del]" & CommunalPayment_1.Date) 
  Call Rekvizit_Fill("Dialog", 1, "General", "PERK", "^A[Del]" & CommunalPayment_1.Date) 
  Call Rekvizit_Fill("Dialog", 1, "General", "DOCISN", CommunalPayment_2.fBASE)
  Call ClickCmdButton(2, "Î³ï³ñ»É")    

  BuiltIn.Delay(2000)
  If wMDIClient.VBObject("frmPttel").VBObject("tdbgView").ApproxCount <> 1 Then
    Log.Error("Երկրորդ կոմունալ վճարումների հանձնարարագիրը առկա չէ 'Հաշվառված վճարային փաստաթղթեր' թղթապանակում:")
    Exit Sub
  Else  
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_Delete)
    Call ClickCmdButton(3, "²Ûá")
  End If

  Call Close_AsBank()
End Sub
