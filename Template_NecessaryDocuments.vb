'USEUNIT Library_Common

Public GroupCashInputISN
Public CreditContractISN

'---------------------------------------------------------------------------------------------
Sub CreateCashInputBatchOrder (DocNum, DocDate, CashAccount, CreditAccount1, CreditAccount2, CreditAccount3, Sum1, Sum2, Sum3, Purpose1, Purpose2, Purpose3, CashMark, WithVerify, Name, fBASE)
  BuiltIn.Delay(3000)
  Call wTreeView.DblClickItem("|Ð³×³Ëáñ¹Ç ëå³ë³ñÏáõÙ ¨ ¹ñ³Ù³ñÏÕ |²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ")
  Call ClickCmdButton(2, "Î³ï³ñ»É")
  
  
  BuiltIn.Delay(3000)
  Call wMainForm.MainMenu.Click(c_AllActions)
  
  Call wMainForm.PopupMenu.Click("?????? ??????????")
  Call wMainForm.PopupMenu.Click("??????? ?????? ?????")

  Set frmASDocForm = wMDIClient.VBObject("frmASDocForm")
  Set TabFrame = frmASDocForm.VBObject("TabFrame")
  fBASE = frmASDocForm.DocFormCommon.Doc.ISN   
  
  Call Rekvizit_Fill("Document", 1, "General", "DOCNUM", DocNum)
  Call Rekvizit_Fill("Document", 1, "General", "DATE", DocDate)
  Call Rekvizit_Fill("Document", 1, "General", "ACCDB", CashAccount)
  
  TabFrame.VBObject("AsTypeFolder").VBObject("TDBMask").Keys(CashAccount & "[Enter][Enter]")

  accountColNum = frmASDocForm.DocFormCommon.Doc.Grid("SubSums").NumFromName("ACCCR")
  summaColNum = frmASDocForm.DocFormCommon.Doc.Grid("SubSums").NumFromName("SUMMA")
  aimColNum = frmASDocForm.DocFormCommon.Doc.Grid("SubSums").NumFromName("AIM")  
  
  TabFrame.VBObject("DocGrid").Col= accountColNum
  TabFrame.VBObject("DocGrid").Keys(CreditAccount1 & "[Right]")
  TabFrame.VBObject("DocGrid").Col= summaColNum
  TabFrame.VBObject("DocGrid").Keys(Sum1 & "[Right]")
  TabFrame.VBObject("DocGrid").Col= aimColNum
  TabFrame.VBObject("DocGrid").Text = Purpose1
  TabFrame.VBObject("DocGrid").Keys("[Enter]")  
  
  TabFrame.VBObject("DocGrid").Col= accountColNum
  TabFrame.VBObject("DocGrid").Keys(CreditAccount2 & "[Right]")
  TabFrame.VBObject("DocGrid").Col= summaColNum
  TabFrame.VBObject("DocGrid").Keys(Sum2 & "[Right]")
  TabFrame.VBObject("DocGrid").Col= aimColNum   
  TabFrame.VBObject("DocGrid").Text = Purpose2
  TabFrame.VBObject("DocGrid").Keys("[Right]")
  
  TabFrame.VBObject("DocGrid").Col= accountColNum
  TabFrame.VBObject("DocGrid").Keys(CreditAccount3 & "[Right]")
  TabFrame.VBObject("DocGrid").Col= summaColNum
  TabFrame.VBObject("DocGrid").Keys(Sum3 & "[Right]")
  TabFrame.VBObject("DocGrid").Col= aimColNum
  TabFrame.VBObject("DocGrid").Text = Purpose3
  TabFrame.VBObject("DocGrid").Keys("[Right]") 
  TabFrame.VBObject("DocGrid").MovePrevious()
  
  str = GetVBObject ("KASSIMV", frmASDocForm)   
  TabFrame.VBObject(str).VBObject("TDBMask").Keys(CashMark & "[Tab]")  
  Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmASDocForm").VBObject("TabFrame").VBObject("AsTpComment").VBObject("TDBComment").keys("Master" & "[Tab]")
  
  DocNum = TabFrame.VBObject("TextC").text
  
  Call Rekvizit_Fill("Document", 1, "General", "PAYER", Name) 
  
  Call ClickCmdButton(1, "Î³ï³ñ»É")
'  Call frmASDocForm.VBObject("CmdOk_2").ClickButton
  'Sys.Process("Asbank").VBObject("frmAsMsgBox").VBObject("cmdButton").Click()
  
  If Not IsNull(WithVerify) Then

    Do Until Trim(wMDIClient.VBObject("frmPttel").VBObject("tdbgView").Columns.Item(2).text) = DocNum or wMDIClient.VBObject("frmPttel").VBObject("tdbgView").Row = wMDIClient.VBObject("frmPttel").VBObject("tdbgView").VisibleRows-1 
      wMDIClient.VBObject("frmPttel").VBObject("tdbgView").MoveNext  
    Loop
    If Trim(wMDIClient.VBObject("frmPttel").VBObject("tdbgView").Columns.Item(2).text) = DocNum  Then
      BuiltIn.Delay(3000)
      Call wMainForm.MainMenu.Click(c_AllActions)        
      Call wMainForm.PopupMenu.Click(c_ToConfirm)       
      Call ClickCmdButton(1, "Î³ï³ñ»É")
    End If
  End If

End Sub

Sub CreateCreditContract_Graph_ (CreditContractISN, CreditContractNumber, MarGraf, SahmGraf, ClientCode, TaxAccount, CurCode,  CreditContractType,  Sum, DistributionType, ExpiryDate, Course, CourseDividor, ExpiredSumPercent, ExpiredSumPercentDividor, ExpiredPercentPercent, ExpiredPercentPercentDividor, UnusedPartPercent, UnusedPartPercentDividor, Branching, Program, Guarantee, Region, FillMethodDate, FillMethodSum, FType )

  BuiltIn.Delay(2000)   
  wMDIClient.VBObject("frmExplorer").SetFocus()
  BuiltIn.Delay(1000)                                                                                                                                                                                                                                                                             
  Call wTreeView.DblClickItem("|ì³ñÏ»ñ (ï»Õ³µ³ßËí³Í)|Üáñ å³ÛÙ³Ý³·ñÇ ëï»ÕÍáõÙ")

  Do Until p1.frmModalBrowser.VBObject("tdbgView").EOF
    If p1.frmModalBrowser.VBObject("tdbgView").Columns.Item(1).Text = "¶ñ³ýÇÏáí í³ñÏ³ÛÇÝ å³ÛÙ³Ý³·Çñ" Then
      Call p1.frmModalBrowser.VBObject("tdbgView").Keys("[Enter]")      
      Exit Do        
    Else
      Call p1.frmModalBrowser.VBObject("tdbgView").MoveNext
    End If    
  Loop   

  Set wfrmASDocForm = wMDIClient.VBObject("frmASDocForm")
  Set wTabFrame = wfrmASDocForm.VBObject("TabFrame")
  
  CreditContractISN= wMDIClient.VBObject("frmASDocForm").DocFormCommon.Doc.ISN
  CreditContractNumber= wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame").VBObject("TextC").Text
 
  expDateStr = "010607"  ' Ø³ñÙ³Ý Å³ÙÏ»ï  
  expDateStr1 = "010507"  ' í³ñÏ³ÛÇÝ ·ÍÇ Å³ÙÏ»ï
  
  Select Case FType
  
  Case 0  'í³ñÏÇ ï»ë³Ïáí --------------------------------------------- 
     
    Call wTabFrame.VBObject("AsTpComment").VBObject("TDBMask").Keys(ClientCode & "[Tab]")
     
    str = GetVBObject ("AGRTYPE", wfrmASDocForm)
    wTabFrame.VBObject(str).VBObject("TDBMask").Keys(CreditContractType & "[Tab]")
    
    Call wTabFrame.VBObject("TDBNumber").Keys(Sum & "[Tab]")

    Set wTabStrip = wfrmASDocForm.VBObject("TabStrip")  

    wTabStrip.SelectedItem = wTabStrip.Tabs(6)
    Set wTabFrame_6 = wfrmASDocForm.VBObject("TabFrame_6")  

    str = GetVBObject ("PPRCODE", wfrmASDocForm)   
    wTabFrame_6.VBObject(str).Keys("P1" & "[Tab]")   
      
    Call wfrmASDocForm.VBObject("CmdOk_2").ClickButton
  
    wMDIClient.VBObject("frmPttel").SetFocus()
  
    BuiltIn.Delay(1000)
    Call wMainForm.MainMenu.Click("¶áñÍáÕáõÃÛáõÝÝ»ñ|´áÉáñ ·áñÍáÕáõÃÛáõÝÝ»ñÁ . . .")
    BuiltIn.Delay(1000)
    Call wMainForm.PopupMenu.Click("Ø³ñÙ³Ý ·ñ³ýÇÏÇ Ýß³Ý³ÏáõÙ")
    BuiltIn.Delay(1000)  
    wfrmASDocForm.VBObject("CmdOk_2").ClickButton 

  Case 1  '³é³Ýó í³ñÏÇ ï»ë³ÏÇ ----------------------------------------
 
    Call wTabFrame.VBObject("AsTpComment").VBObject("TDBMask").Keys(ClientCode & "[Tab]")
    Call wTabFrame.VBObject("AsTypeFolder_2").VBObject("TDBMask").Keys(CurCode & "[Tab]")
    Call wTabFrame.VBObject("TDBNumber").Keys(Sum & "[Tab]")

    Call wTabFrame.VBObject("TDBDate_3").Keys(expDateStr & "[Tab]")
    
    Set wTabStrip = wfrmASDocForm.VBObject("TabStrip")
    wTabStrip.SelectedItem = wTabStrip.Tabs(4)
    Set wTabFrame_4 = wfrmASDocForm.VBObject("TabFrame_4")
    
    wTabFrame_4.VBObject("AsCourse_2").VBObject("TDBNumber1").Keys(Course & "[Tab]")  
    wTabFrame_4.VBObject("AsCourse_2").VBObject("TDBNumber2").Keys(CourseDividor & "[Tab]")
    
    wTabFrame_4.VBObject("AsCourse_3").VBObject("TDBNumber1").Keys(UnusedPartPercent & "[Tab]")
    wTabFrame_4.VBObject("AsCourse_3").VBObject("TDBNumber2").Keys(UnusedPartPercentDividor & "[Tab]") 


    wTabStrip.SelectedItem = wTabStrip.Tabs(5)
    Set wTabFrame_5 = wfrmASDocForm.VBObject("TabFrame_5")

    wTabFrame_5.VBObject("AsCourse_5").VBObject("TDBNumber1").Keys(ExpiredSumPercent & "[Tab]") 
    wTabFrame_5.VBObject("AsCourse_5").VBObject("TDBNumber2").Keys(ExpiredSumPercentDividor & "[Tab]")  
    wTabFrame_5.VBObject("AsCourse_6").VBObject("TDBNumber1").Keys(ExpiredPercentPercent & "[Tab]") 
    wTabFrame_5.VBObject("AsCourse_6").VBObject("TDBNumber2").Keys(ExpiredPercentPercentDividor & "[Tab]")  

  
    wTabStrip.SelectedItem = wTabStrip.Tabs(6)
    Set wTabFrame_6 = wfrmASDocForm.VBObject("TabFrame_6")  
  
    str = GetVBObject ("BRANCH", wfrmASDocForm)   
    wTabFrame_6.VBObject(str).VBObject("TDBMask").Keys(Branching & "[Tab]")    
  
    str = GetVBObject ("SCHEDULE", wfrmASDocForm)   
    wTabFrame_6.VBObject(str).VBObject("TDBMask").Keys(Program & "[Tab]")   
 
    str = GetVBObject ("GUARANTEE", wfrmASDocForm)   
    wTabFrame_6.VBObject(str).VBObject("TDBMask").Keys(Guarantee & "[Tab]")     
  
    str = GetVBObject ("LRDISTR", wfrmASDocForm)   
    wTabFrame_6.VBObject(str).VBObject("TDBMask").Keys(Region & "[Tab]")    

    str = GetVBObject ("PPRCODE", wfrmASDocForm)   
    wTabFrame_6.VBObject(str).Keys("P1" & "[Tab]")
           
    wfrmASDocForm.VBObject("CmdOk_2").ClickButton
  
    wMDIClient.VBObject("frmPttel").SetFocus()
  
    BuiltIn.Delay(1000)
    Call wMainForm.MainMenu.Click("¶áñÍáÕáõÃÛáõÝÝ»ñ|´áÉáñ ·áñÍáÕáõÃÛáõÝÝ»ñÁ . . .")
    BuiltIn.Delay(1000)
    Call wMainForm.PopupMenu.Click("Ø³ñÙ³Ý ·ñ³ýÇÏÇ Ýß³Ý³ÏáõÙ")
    BuiltIn.Delay(1000)  
  
    str = GetVBObject ("AUTODATEUN", wMDIClient.VBObject("frmASDocForm"))   
    wfrmASDocForm.VBObject("TabFrame").VBObject(str).ClickButton(cbChecked)
    'wfrmASDocForm.VBObject("TabFrame").VBObject("CheckBox_2").ClickButton(cbChecked)
    Set wfrmAsUstPar = Sys.Process("Asbank").VBObject("frmAsUstPar")
    

    str = GetVBObject_Dialog ("DATESFILLTYPE", wfrmAsUstPar)                           
    wfrmAsUstPar.VBObject("TabFrame").VBObject(str).Keys(FillMethodDate & "[Tab]")

    str = GetVBObject_Dialog ("AGRPERIOD", wfrmAsUstPar)                           
    wfrmAsUstPar.VBObject("TabFrame").VBObject(str).Keys("1" & "[Tab][Tab]")

    str = GetVBObject_Dialog ("SUMSFILLTYPE", wfrmAsUstPar)                           
    wfrmAsUstPar.VBObject("TabFrame").VBObject(str).Keys(FillMethodSum & "[Tab]")
  
    str = GetVBObject_Dialog ("OVERDAY", wfrmAsUstPar)                           
    wfrmAsUstPar.VBObject("TabFrame").VBObject(str).Keys("1" & "[Tab]")

    wfrmAsUstPar.VBObject("CmdOK").ClickButton 
    
    
    wfrmASDocForm.VBObject("CmdOk_2").ClickButton 
 
  Case 2  'í³ñÏ³ÛÇÝ ·Íáí ---------------------------------------------
  
    Call wTabFrame.VBObject("AsTpComment").VBObject("TDBMask").Keys(ClientCode & "[Tab]")
    Call wTabFrame.VBObject("AsTypeFolder_2").VBObject("TDBMask").Keys(CurCode & "[Tab]")
    If DistributionType = 0 Then
      Call wTabFrame.VBObject("AsTypeFolder_3").VBObject("TDBMask").Keys("77782963313" & "[Tab]")
    End If
    Call wTabFrame.VBObject("TDBNumber").Keys(Sum & "[Tab]")  
    Call wTabFrame.VBObject("TDBDate_3").Keys(expDateStr & "[Tab]")   ' Ù³ñÙ³Ý Å³ÙÏ»ï 
    
    str = GetVBObject ("ISLINE", wfrmASDocForm)   
    wTabFrame.VBObject(str).ClickButton(cbChecked)    

    str = GetVBObject ("MARPERTYPE", wfrmASDocForm)   
    wTabFrame.VBObject(str).VBObject("TDBMask").Keys(DistributionType & "[Tab]") 

    str = GetVBObject ("DATELNGEND", wfrmASDocForm)   
    wTabFrame.VBObject(str).Keys(expDateStr1 & "[Tab]") 

     Set wTabStrip = wfrmASDocForm.VBObject("TabStrip")
    
    If DistributionType = 1 or DistributionType = 2 Then

      wTabStrip.SelectedItem = wTabStrip.Tabs(3)          
      Set wTabFrame_3 = wfrmASDocForm.VBObject("TabFrame_3")

      str = GetVBObject ("DATESFILLTYPE", wfrmASDocForm)   
      wTabFrame_3.VBObject(str).Keys("2" & "[Tab]") 

      str = GetVBObject ("SUMSFILLTYPE", wfrmASDocForm)   
      wTabFrame_3.VBObject(str).Keys("11" & "[Tab]") 

      str = GetVBObject ("DBTMINPER", wfrmASDocForm)   
      wTabFrame_3.VBObject(str).Keys("10" & "[Tab]") 
      
      str = GetVBObject ("AGRPERIOD", wfrmASDocForm)   
      wTabFrame_3.VBObject(str).Keys("1" & "[Tab][Tab]") 

      str = GetVBObject ("OVERDAY", wfrmASDocForm)   
      wTabFrame_3.VBObject(str).Keys("1" & "[Tab]") 

    End If
       
    wTabStrip.SelectedItem = wTabStrip.Tabs(4)              ' í³ñÏÇ ïáÏáë³¹ñáõÛùÁ, µ³Å. .... 
    Set wTabFrame_4 = wfrmASDocForm.VBObject("TabFrame_4")
 
    wTabFrame_4.VBObject("AsCourse_2").VBObject("TDBNumber1").Keys(Course & "[Tab]")  
    wTabFrame_4.VBObject("AsCourse_2").VBObject("TDBNumber2").Keys(CourseDividor & "[Tab]")
    
    wTabFrame_4.VBObject("AsCourse_3").VBObject("TDBNumber1").Keys(UnusedPartPercent & "[Tab]")
    wTabFrame_4.VBObject("AsCourse_3").VBObject("TDBNumber2").Keys(UnusedPartPercentDividor & "[Tab]") 


    wTabStrip.SelectedItem = wTabStrip.Tabs(5)
    Set wTabFrame_5 = wfrmASDocForm.VBObject("TabFrame_5")

    wTabFrame_5.VBObject("AsCourse_5").VBObject("TDBNumber1").Keys(ExpiredSumPercent & "[Tab]") 
    wTabFrame_5.VBObject("AsCourse_5").VBObject("TDBNumber2").Keys(ExpiredSumPercentDividor & "[Tab]")  
    wTabFrame_5.VBObject("AsCourse_6").VBObject("TDBNumber1").Keys(ExpiredPercentPercent & "[Tab]") 
    wTabFrame_5.VBObject("AsCourse_6").VBObject("TDBNumber2").Keys(ExpiredPercentPercentDividor & "[Tab]")  
  
    wTabStrip.SelectedItem = wTabStrip.Tabs(6)
    Set wTabFrame_6 = wfrmASDocForm.VBObject("TabFrame_6")  
  
    str = GetVBObject ("BRANCH", wfrmASDocForm)   
    wTabFrame_6.VBObject(str).VBObject("TDBMask").Keys(Branching & "[Tab]")    
  
    str = GetVBObject ("SCHEDULE", wfrmASDocForm)   
    wTabFrame_6.VBObject(str).VBObject("TDBMask").Keys(Program & "[Tab]")   
 
    str = GetVBObject ("GUARANTEE", wfrmASDocForm)   
    wTabFrame_6.VBObject(str).VBObject("TDBMask").Keys(Guarantee & "[Tab]")     
  
    str = GetVBObject ("LRDISTR", wfrmASDocForm)   
    wTabFrame_6.VBObject(str).VBObject("TDBMask").Keys(Region & "[Tab]")    

    str = GetVBObject ("PPRCODE", wfrmASDocForm)   
    wTabFrame_6.VBObject(str).Keys("P1" & "[Tab]")
               
    wfrmASDocForm.VBObject("CmdOk_2").ClickButton
    
    wMDIClient.VBObject("frmPttel").SetFocus()

    If DistributionType = 3 Then   'ë³ÑÙ³Ý³ã³÷»ñáí µ³ßËíáÕ 
  
      BuiltIn.Delay(1000)
      Call wMainForm.MainMenu.Click("¶áñÍáÕáõÃÛáõÝÝ»ñ|´áÉáñ ·áñÍáÕáõÃÛáõÝÝ»ñÁ . . .")
      BuiltIn.Delay(1000)
      Call wMainForm.PopupMenu.Click("ê³ÑÙ³Ý³ã³÷Ç ·ñ³ýÇÏÇ Ýß³Ý³ÏáõÙ")
      BuiltIn.Delay(1000) 
  
      str = GetVBObject ("AUTODATELM", wMDIClient.VBObject("frmASDocForm"))   
      wfrmASDocForm.VBObject("TabFrame").VBObject(str).ClickButton(cbChecked)
     ' wfrmASDocForm.VBObject("TabFrame").VBObject("CheckBox").ClickButton(cbChecked) 

      Set wfrmAsDialog2 = Sys.Process("Asbank").VBObject("frmAsDialog2")
  
      wfrmAsDialog2.VBObject("TabFrame").VBObject("TDBNumber").Keys(1 & "[Tab]")

      str = GetVBObject_Dialog ("OVERDAY", wfrmAsDialog2)                           
      wfrmAsDialog2.VBObject("TabFrame").VBObject(str).Keys("1" & "[Tab]")


      wfrmAsDialog2.VBObject("CmdOK").ClickButton    
  
      wfrmASDocForm.VBObject("CmdOk_2").ClickButton  

      Call wMDIClient.VBObject("frmPttel").VBObject("tdbgView").MoveFirst
    
    End If    
  
    BuiltIn.Delay(1000)
    Call wMainForm.MainMenu.Click("¶áñÍáÕáõÃÛáõÝÝ»ñ|´áÉáñ ·áñÍáÕáõÃÛáõÝÝ»ñÁ . . .")
    BuiltIn.Delay(1000)
    Call wMainForm.PopupMenu.Click("Ø³ñÙ³Ý ·ñ³ýÇÏÇ Ýß³Ý³ÏáõÙ")
    BuiltIn.Delay(1000)  
    
    If Not DistributionType = 0 Then
 
    str = GetVBObject ("AUTODATEUN", wMDIClient.VBObject("frmASDocForm"))   
    wfrmASDocForm.VBObject("TabFrame").VBObject(str).ClickButton(cbChecked)
    'wfrmASDocForm.VBObject("TabFrame").VBObject("CheckBox_2").ClickButton(cbChecked)
    Set wfrmAsUstPar = Sys.Process("Asbank").VBObject("frmAsUstPar")
   
    str = GetVBObject_Dialog ("DATESFILLTYPE", wfrmAsUstPar)                           
    wfrmAsUstPar.VBObject("TabFrame").VBObject(str).Keys(FillMethodDate & "[Tab]")

    str = GetVBObject_Dialog ("AGRPERIOD", wfrmAsUstPar)                           
    wfrmAsUstPar.VBObject("TabFrame").VBObject(str).Keys("1" & "[Tab][Tab]")

    str = GetVBObject_Dialog ("SUMSFILLTYPE", wfrmAsUstPar)                           
    wfrmAsUstPar.VBObject("TabFrame").VBObject(str).Keys(FillMethodSum & "[Tab]")
  
    str = GetVBObject_Dialog ("OVERDAY", wfrmAsUstPar)                           
    wfrmAsUstPar.VBObject("TabFrame").VBObject(str).Keys("1" & "[Tab]")

    wfrmAsUstPar.VBObject("CmdOK").ClickButton 
    

    
      wfrmASDocForm.VBObject("CmdOk_2").ClickButton 
    End If  
      
  Case 3    'í³ñÏ³ÛÇÝ ù³ñïáí -----------------------------------------
  
    Call wTabFrame.VBObject("AsTpComment").VBObject("TDBMask").Keys(ClientCode & "[Tab]")
    Call wTabFrame.VBObject("AsTypeFolder_2").VBObject("TDBMask").Keys(CurCode & "[Tab]")    
    Call wTabFrame.VBObject("AsTypeFolder_3").VBObject("TDBMask").Keys(TaxAccount & "[Tab]")'Ñ³ñÏ³ÛÇÝ Ñ³ßÇí
    Call wTabFrame.VBObject("TDBNumber").Keys(Sum & "[Tab]") 

    Call wTabFrame.VBObject("TDBDate_3").Keys(expDateStr & "[Tab]")   ' Ù³ñÙ³Ý Å³ÙÏ»ï
    
    str = GetVBObject ("ISLINE", wfrmASDocForm)   
    wTabFrame.VBObject(str).ClickButton(cbChecked)  
    
    str = GetVBObject ("MARPERTYPE", wfrmASDocForm)   
    wTabFrame.VBObject(str).VBObject("TDBMask").Keys(DistributionType & "[Tab]") 

    
    str = GetVBObject ("DATELNGEND", wfrmASDocForm)   
    wTabFrame.VBObject(str).Keys(expDateStr1 & "[Tab]") 

    str = GetVBObject ("ISCRCARD", wfrmASDocForm)   
    wTabFrame.VBObject(str).ClickButton(cbChecked)          
  
    str = GetVBObject ("CARDAGRTYPE", wfrmASDocForm)   
    wTabFrame.VBObject(str).VBObject("TDBMask").Keys(2 & "[Tab]") 
 
    Set wTabStrip = wfrmASDocForm.VBObject("TabStrip")

    wTabStrip.SelectedItem = wTabStrip.Tabs(3)          
    Set wTabFrame_3 = wfrmASDocForm.VBObject("TabFrame_3")

      str = GetVBObject ("DATESFILLTYPE", wfrmASDocForm)   
      wTabFrame_3.VBObject(str).Keys("2" & "[Tab]") 

      str = GetVBObject ("SUMSFILLTYPE", wfrmASDocForm)   
      wTabFrame_3.VBObject(str).Keys("11" & "[Tab]") 

      str = GetVBObject ("DBTMINPER", wfrmASDocForm)   
      wTabFrame_3.VBObject(str).Keys("10" & "[Tab]") 
      
      str = GetVBObject ("AGRPERIOD", wfrmASDocForm)   
      wTabFrame_3.VBObject(str).Keys("1" & "[Tab][Tab]") 

      str = GetVBObject ("OVERDAY", wfrmASDocForm)   
      wTabFrame_3.VBObject(str).Keys("1" & "[Tab]") 


    wTabStrip.SelectedItem = wTabStrip.Tabs(4)              
    Set wTabFrame_4 = wfrmASDocForm.VBObject("TabFrame_4")
 
    wTabFrame_4.VBObject("AsCourse_2").VBObject("TDBNumber1").Keys(Course & "[Tab]")  
    wTabFrame_4.VBObject("AsCourse_2").VBObject("TDBNumber2").Keys(CourseDividor & "[Tab]")
    
    wTabFrame_4.VBObject("AsCourse_3").VBObject("TDBNumber1").Keys(UnusedPartPercent & "[Tab]")
    wTabFrame_4.VBObject("AsCourse_3").VBObject("TDBNumber2").Keys(UnusedPartPercentDividor & "[Tab]") 


    wTabStrip.SelectedItem = wTabStrip.Tabs(5)
    Set wTabFrame_5 = wfrmASDocForm.VBObject("TabFrame_5")

    wTabFrame_5.VBObject("AsCourse_5").VBObject("TDBNumber1").Keys(ExpiredSumPercent & "[Tab]") 
    wTabFrame_5.VBObject("AsCourse_5").VBObject("TDBNumber2").Keys(ExpiredSumPercentDividor & "[Tab]")  
    wTabFrame_5.VBObject("AsCourse_6").VBObject("TDBNumber1").Keys(ExpiredPercentPercent & "[Tab]") 
    wTabFrame_5.VBObject("AsCourse_6").VBObject("TDBNumber2").Keys(ExpiredPercentPercentDividor & "[Tab]")  
  
    wTabStrip.SelectedItem = wTabStrip.Tabs(6)
    Set wTabFrame_6 = wfrmASDocForm.VBObject("TabFrame_6")  
    
    str = GetVBObject ("BRANCH", wfrmASDocForm)   
    wTabFrame_6.VBObject(str).VBObject("TDBMask").Keys(Branching & "[Tab]")    
  
    str = GetVBObject ("SCHEDULE", wfrmASDocForm)   
    wTabFrame_6.VBObject(str).VBObject("TDBMask").Keys(Program & "[Tab]")   
 
    str = GetVBObject ("GUARANTEE", wfrmASDocForm)   
    wTabFrame_6.VBObject(str).VBObject("TDBMask").Keys(Guarantee & "[Tab]")     
  
    str = GetVBObject ("LRDISTR", wfrmASDocForm)   
    wTabFrame_6.VBObject(str).VBObject("TDBMask").Keys(Region & "[Tab]")    

    str = GetVBObject ("PPRCODE", wfrmASDocForm)   
    wTabFrame_6.VBObject(str).Keys("P1" & "[Tab]")
              
    wfrmASDocForm.VBObject("CmdOk_2").ClickButton
    
    wMDIClient.VBObject("frmPttel").SetFocus()

    BuiltIn.Delay(1000)
    Call wMainForm.MainMenu.Click("¶áñÍáÕáõÃÛáõÝÝ»ñ|´áÉáñ ·áñÍáÕáõÃÛáõÝÝ»ñÁ . . .")
    BuiltIn.Delay(1000)
    Call wMainForm.PopupMenu.Click("Ø³ñÙ³Ý ·ñ³ýÇÏÇ Ýß³Ý³ÏáõÙ")
    BuiltIn.Delay(1000)  
  
    str = GetVBObject ("AUTODATEUN", wMDIClient.VBObject("frmASDocForm"))   
    wfrmASDocForm.VBObject("TabFrame").VBObject(str).ClickButton(cbChecked)
    'wfrmASDocForm.VBObject("TabFrame").VBObject("CheckBox_2").ClickButton(cbChecked)
    Set wfrmAsUstPar = Sys.Process("Asbank").VBObject("frmAsUstPar")
    
    str = GetVBObject_Dialog ("DATESFILLTYPE", wfrmAsUstPar)                           
    wfrmAsUstPar.VBObject("TabFrame").VBObject(str).Keys(FillMethodDate & "[Tab]")

    str = GetVBObject_Dialog ("AGRPERIOD", wfrmAsUstPar)                           
    wfrmAsUstPar.VBObject("TabFrame").VBObject(str).Keys("1" & "[Tab][Tab]")

    str = GetVBObject_Dialog ("SUMSFILLTYPE", wfrmAsUstPar)                           
    wfrmAsUstPar.VBObject("TabFrame").VBObject(str).Keys(FillMethodSum & "[Tab]")
  
    str = GetVBObject_Dialog ("OVERDAY", wfrmAsUstPar)                           
    wfrmAsUstPar.VBObject("TabFrame").VBObject(str).Keys("1" & "[Tab]")

    wfrmAsUstPar.VBObject("CmdOK").ClickButton 
    
    wfrmASDocForm.VBObject("CmdOk_2").ClickButton 

  End Select   
   
End Sub

'---------------------------------------------------------------------------------------------
Function SerializeDate (DateWithSlashes)

  dd = Left(DateWithSlashes,2)
  mm = Mid(DateWithSlashes,4,2)
  yy = Right(DateWithSlashes,2)
    
  Utilities.ShortDateFormat = "yyyymmdd" 
  SerializeDate = Utilities.DateToStr(DateSerial(yy, mm, dd)) 
  
End Function

Sub FillMarGraf(MarGraf, FType, FillMethod, DistributionType)                 
   
  MarGraf(1,1) = SerializeDate("03/07/06")         
  MarGraf(2,1) = SerializeDate("01/08/06")         
  MarGraf(3,1) = SerializeDate("01/09/06")        
  MarGraf(4,1) = SerializeDate("02/10/06")         
  MarGraf(5,1) = SerializeDate("01/11/06")         
  MarGraf(6,1) = SerializeDate("01/12/06")         
  MarGraf(7,1) = SerializeDate("01/01/07")         
  MarGraf(8,1) = SerializeDate("01/02/07")          
  MarGraf(9,1) = SerializeDate("01/03/07")            
  MarGraf(10,1) = SerializeDate("02/04/07")         
  MarGraf(11,1) = SerializeDate("01/05/07")           
  MarGraf(12,1) = SerializeDate("01/06/07")        
                                                   
  Select Case FType
  
  Case 0      ' í³ñÏÇ ï»ë³Ïáí  
   
    MarGraf(1,2) = "36779.00"     
    MarGraf(2,2) = "38466.60"    
    MarGraf(3,2) = "38641.50"    
    MarGraf(4,2) = "39429.10"     
    MarGraf(5,2) = "40460.80"   
    MarGraf(6,2) = "41258.90"   
    MarGraf(7,2) = "41898.60"   
    MarGraf(8,2) = "42752.60"    
    MarGraf(9,2) = "43979.80"     
    MarGraf(10,2) = "44430.90"    
    MarGraf(11,2) = "45547.10"     
    MarGraf(12,2) = "46355.10"  
    
    MarGraf(1,3) = "10520.50"          
    MarGraf(2,3) = "8832.90" 
    MarGraf(3,3) = "8658.00" 
    MarGraf(4,3) = "7870.40" 
    MarGraf(5,3) = "6838.70" 
    MarGraf(6,3) = "6040.60" 
    MarGraf(7,3) = "5400.90" 
    MarGraf(8,3) = "4546.90"      
    MarGraf(9,3) = "3319.70"
    MarGraf(10,3) = "2868.60" 
    MarGraf(11,3) = "1752.40"
    MarGraf(12,3) = "944.90"
  
  Case 1         '³é³Ýó í³ñÏÇ ï»ë³ÏÇ 
  
    If FillMethod = 1 Then
     
      MarGraf(1,2) = "41666.70"     
      MarGraf(2,2) = "41666.70"    
      MarGraf(3,2) = "41666.70"    
      MarGraf(4,2) = "41666.70"     
      MarGraf(5,2) = "41666.70"   
      MarGraf(6,2) = "41666.70"   
      MarGraf(7,2) = "41666.70"   
      MarGraf(8,2) = "41666.70"    
      MarGraf(9,2) = "41666.70"     
      MarGraf(10,2) = "41666.70"    
      MarGraf(11,2) = "41666.70"     
      MarGraf(12,2) = "41666.30"  
    
      MarGraf(1,3) = "10520.50"          
      MarGraf(2,3) = "8739.70" 
      MarGraf(3,3) = "8493.10" 
      MarGraf(4,3) = "7643.80" 
      MarGraf(5,3) = "6575.30" 
      MarGraf(6,3) = "5753.40" 
      MarGraf(7,3) = "5095.90" 
      MarGraf(8,3) = "4246.60"      
      MarGraf(9,3) = "3068.50"
      MarGraf(10,3) = "2630.10" 
      MarGraf(11,3) = "1589.00"
      MarGraf(12,3) = "849.30"
   
    End If
  
    If FillMethod = 21 Then
    
      MarGraf(1,2) = "36779.00"     
      MarGraf(2,2) = "38466.60"    
      MarGraf(3,2) = "38641.50"    
      MarGraf(4,2) = "39429.10"     
      MarGraf(5,2) = "40460.80"   
      MarGraf(6,2) = "41258.90"   
      MarGraf(7,2) = "41898.60"   
      MarGraf(8,2) = "42752.60"    
      MarGraf(9,2) = "43979.80"     
      MarGraf(10,2) = "44430.90"    
      MarGraf(11,2) = "45547.10"     
      MarGraf(12,2) = "46355.10"  
    
      MarGraf(1,3) = "10520.50"          
      MarGraf(2,3) = "8832.90" 
      MarGraf(3,3) = "8658.00" 
      MarGraf(4,3) = "7870.40" 
      MarGraf(5,3) = "6838.70" 
      MarGraf(6,3) = "6040.60" 
      MarGraf(7,3) = "5400.90" 
      MarGraf(8,3) = "4546.90"      
      MarGraf(9,3) = "3319.70"
      MarGraf(10,3) = "2868.60" 
      MarGraf(11,3) = "1752.40"
      MarGraf(12,3) = "944.80"
    
    End If
    
  Case 2           'í³ñÏ³ÛÇÝ ·Íáí 
  
    If DistributionType = 0 Then
      MarGraf(1,2) = "41666.70"     
      MarGraf(2,2) = "41666.70"    
      MarGraf(3,2) = "41666.70"    
      MarGraf(4,2) = "41666.70"     
      MarGraf(5,2) = "41666.70"   
      MarGraf(6,2) = "41666.70"   
      MarGraf(7,2) = "41666.70"   
      MarGraf(8,2) = "41666.70"    
      MarGraf(9,2) = "41666.70"     
      MarGraf(10,2) = "41666.70"    
      MarGraf(11,2) = "41666.70"     
      MarGraf(12,2) = "41666.30"  
    
      MarGraf(1,3) = "10520.50"          
      MarGraf(2,3) = "8739.70" 
      MarGraf(3,3) = "8493.10" 
      MarGraf(4,3) = "7643.80" 
      MarGraf(5,3) = "6575.30" 
      MarGraf(6,3) = "5753.40" 
      MarGraf(7,3) = "5095.90" 
      MarGraf(8,3) = "4246.60"      
      MarGraf(9,3) = "3068.50"
      MarGraf(10,3) = "2630.10" 
      MarGraf(11,3) = "1589.00"
      MarGraf(12,3) = "849.30"
    Else
       For i = 1 to 12
          MarGraf(i,2) = "0" 
          MarGraf(i,3) = "0" 
       Next
    End If
    
  Case 3        ' í³ñÏ³ÛÇÝ ù³ñïáí 
  
    For i = 1 to 12
      MarGraf(i,2) = "0" 
      MarGraf(i,3) = "0" 
    Next
    
  End Select 
End Sub

Public Sub CreateNecessaryDocuments
  Utilities.ShortDateFormat = "yyyymmdd" 
  endDATE = Utilities.DateToStr(Utilities.Date())   
  startDATE = "20100101"   'startDATE = Utilities.DateToStr(Utilities.IncMonth(Utilities.Now, -124)) 
    
  Call Initialize_AsBank("bank", startDATE, endDATE)  
  Call login("Armsoft")  
        
    GroupCashInputISN = ""
    CreditContractISN = ""
    
    'creating group cash input for testing template printing
    DocNum = "000888"
    DocDate = "240310"
    CashAccount = "000001100"
    CreditAccount1 = "10330030101"
    CreditAccount2 = "00068020101"
    CreditAccount3 = "33170080101"
    Sum1 = 150.15
    Sum2 = 301.20
    Sum3 = 3548.65
    Purpose1 = " Ð³Ù³Ó³ÛÝ ³ñï. ³éùáõí³×³éùÇ å³ÛÙ³Ý³·ñÇ"
    Purpose2 = " Ð³ñÏ»ñÇ Ù³ñáõÙ"
    Purpose3 = "Custom Purpose"
    CashMark = "022"
    Name = "Ð³Û³·ñáµ³ÝÏ"
    
    Call Login ("teller")
    Call CreateCashInputBatchOrder (Docnum, Docdate, CashAccount, CreditAccount1, CreditAccount2, CreditAccount3, _
                                    Sum1, Sum2, Sum3, Purpose1, Purpose2, Purpose3, CashMark, _
                                    WithVerify, Name, fBASE)
    
    GroupCashInputISN = fBASE
    
    Call Login("creditoperator")
    AllocateWithLim = False
    IsCard = False
    FType = 2
    FillMethod = 0
    FillMethodDate = ""
    FillMethodSumDate = ""
    FillMethodSum = ""
    
    ClientCode = "00000233"
    TaxAccount = "77782963313" 'Ñ³ñÏ³ÛÇÝ Ñ³ßÇí
    CreditContractType = "0004" 'êå³éáÕ³Ï³Ý 1 ï³ñÇ AMD
    Course = "24" ' ì³ñÏÇ ïáÏáë³¹ñáõÛù
    CourseDividor = "365"
    ExpiredSumPercent = "0.15" ' Ä³ÙÏ»ï³Ýó ·áõÙ³ñÇ ïáÏáë³¹ñáõÛù      PPA
    ExpiredSumPercentDividor = "1" ' ´³Å³Ý³ñ³ñ
    ExpiredPercentPercent = "0.15" ' Ä³ÙÏ»ï³Ýó ïáÏáëÇ ïáÏáë³¹ñáõÛù    PPP
    ExpiredPercentPercentDividor = "1" ' ´³Å³Ý³ñ³ñ
    UnusedPartPercentDividor = "1"
    UnusedPartPercent = "0.15"
    Branching = "9"
    Program = "9"
    Guarantee = "9"
    Region = "001"
    
    
    Sum = 500000
    
    CurCode = "000"
    ExpiryDate = "12"
    
    
    Utilities.ShortDateFormat = "ddmmyy"
    createDate = Utilities.DateToStr(Utilities.Now)
    startDate = Utilities.DateToStr(Utilities.Now)
    
    expDateStr = Utilities.DateToStr(Utilities.IncMonth(Utilities.Now, 12))
    expDateStr1 = Utilities.DateToStr(Utilities.IncMonth(Utilities.Now, 12))
    
    
    Utilities.ShortDateFormat = "dd/mm/yy"
    
    MarGraf = CreateVariantArray2(1, ExpiryDate, 1, 3)
    SahmGraf = CreateVariantArray2(1, ExpiryDate, 1, 2)
    
    Call FillMarGraf( MarGraf, FType, FillMethod)
    
    Call CreateCreditContract_Graph_ (CreditContractISN, CreditContractNumber, MarGraf, SahmGraf, ClientCode, _
                                      TaxAccount, CurCode, CreditContractType, Sum, AllocateWithLim, IsCard, _
                                      ExpiryDate, Course, CourseDividor, ExpiredSumPercent, ExpiredSumPercentDividor, _
                                      ExpiredPercentPercent, ExpiredPercentPercentDividor, UnusedPartPercent, UnusedPartPercentDividor, _
                                      Branching, Program, Guarantee, Region, FillMethod, FillMethodDate, FillMethodSumDate, _
                                      FillMethodSum, "010607", "010507", "010606", "010606", FType )
    CreditContractISN = CreditContractISN
    
End Sub

'-------------------------------------------------------------------------------------------------------
Public Sub DeleteNecessaryDocuments
    
    If GroupCashInputISN <> "" Then
        DeleteDoc(GroupCashInputISN)
    End If
    
    If CreditContractISN <> "" Then
        DeleteDoc(CreditContractISN)
    End If
End Sub

'-------------------------------------------------------------------------------------------------------
Public Sub DeleteDoc(fISN)
    Set asbank = p1
    Call LocateDocument(fISN)
    
    Call asbank.MainForm.MainMenu.Click("Գործողություններ|Ջնջել")
    
    Set frmAsMsgBox = asbank.WaitVBObject("frmAsMsgBox", 1000)
    If frmAsMsgBox.Exists Then
        Call frmAsMsgBox.vbObject("cmdButton").clickButton
    End If
    
    Set frmDeleteDoc = asbank.WaitVBObject("frmDeleteDoc", 1000)
    If frmDeleteDoc.Exists Then
        Call frmDeleteDoc.vbObject("YesButton").ClickButton
    Else
        Log.Error("Problem while deleteing doc " & fISN)
    End If
    
    wMDIClient.frmPttel.Close()
End Sub