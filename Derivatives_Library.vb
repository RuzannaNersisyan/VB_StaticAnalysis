Option Explicit

'USEUNIT Library_Common  
'USEUNIT Constants

Class DerivativeDoc
  Public DocNum, Client, BuyCurr, RepayCurr, BuyAcc, RepayAcc, SaleAmount,_
          PurAmount, Date, Term, ForwardExchg, AutoDebt, RecRate, PayRate, Baj, Sector,_
          UsageField, Aim, Schedule, Guarantee, Country, District, RegionLR,_
          Time, PurSale, OpPlace, PaperCode, fBASE
  Public BaseSum, DateFill, FirstDate, CheckPayDates, Paragraph, Direction, PayDates, OpFinalPay
         
  Private Sub Class_Initialize()
    Client = Null
    BuyCurr = "000"
    RepayCurr = "001"
    AutoDebt = 1
    RecRate = 12
    PayRate = 5
    Baj = 365
    Sector = "U2"
    UsageField = "01.001"
    Aim = "00"
    Schedule = 9
    Guarantee = 9
    Country = "AM"
    District = "001"
    RegionLR = "010000008"
    Time = 1
    PurSale = 1
    OpPlace = 4
    DateFill = 1
    CheckPayDates = 1
    Direction = 2
    PayDates = 15
    OpFinalPay = 1
  End Sub
  
  Public Sub CreateDerivative(FolderName, DocType) 
   Dim frmModalBrowser, wTabStrip, TabN, Rekv
    
    Call wTreeView.DblClickItem(FolderName)
    
    Set frmModalBrowser  = Sys.Process("Asbank").WaitVBObject("frmModalBrowser", 500)	
	  If frmModalBrowser.Exists	Then
			Do Until p1.frmModalBrowser.VBObject("tdbgView").EOF
  			If RTrim(p1.frmModalBrowser.VBObject("tdbgView").Columns.Item(col_item).Text) = DocType  Then
    			Call p1.frmModalBrowser.VBObject("tdbgView").Keys("[Enter]")
    			Exit do
  			Else
    			Call p1.frmModalBrowser.VBObject("tdbgView").MoveNext
  			End If
  		Loop 
	  Else
		  Log.Error("frmModalBrowser does not exists.")
		  Exit Sub
	  End If
    
    'Վերցնել "Պայմանագրի համար" դաշտի արշժեքը
    DocNum = wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame").VBObject("TextC").Text 
    'Լրացնել "Հաճախորդ" դաշտը 
    Call Rekvizit_Fill("Document", 1, "General", "CLICOD", Client)
    'Լրացնել "Ֆորվարդի գնման արժույթ" դաշտը 
    Call Rekvizit_Fill("Document", 1, "General", "CURRENCY", BuyCurr)
    
    If DocType <> "îáÏáë³¹ñáõÛù³ÛÇÝ ëíá÷" Then
      'Լրացնել "Ֆորվարդի վաճառքի արժույթ" դաշտը 
      Call Rekvizit_Fill("Document", 1, "General", "CURRENCY2", RepayCurr)
    End If  
    
    'Լրացնել "Գնման արժույթի հաշիվ" դաշտը 
    Call Rekvizit_Fill("Document", 1, "General", "ACCACC", BuyAcc)
    'Լրացնել "Վաճառքի արժույթի հաշիվ" դաշտը 
    Call Rekvizit_Fill("Document", 1, "General", "ACCACC1", RepayAcc)
    'Լրացնել "Կնքման ամսաթիվ" դաշտը 
    Call Rekvizit_Fill("Document", 1, "General", "DATE", Date)
    
    Select Case DocType
      Case "öáË³ñÅ»ù³ÛÇÝ ëíá÷", "²ñÅáõÃ³ÛÇÝ ëíá÷"
        'Լրացնել "Ձևակերպման ամսաթիվ" դաշտը 
        Call Rekvizit_Fill("Document", 1, "General", "DATEGIVE", Date)
        'Լրացնել "Վաճառվող գումար" դաշտը 
        Call Rekvizit_Fill("Document", 1, "General", "SSUMDB", SaleAmount)
    End Select  
    
    If DocType = "îáÏáë³¹ñáõÛù³ÛÇÝ ëíá÷" Then
      Call Rekvizit_Fill("Document", 1, "General", "SUMMA", BaseSum)
    Else
      'Լրացնել "Մարման ժամկետ" դաշտը 
      Call Rekvizit_Fill("Document", 1, "General", "DATEAGR", Term)
      'Լրացնել "Ֆորվարդի փոխարժեք" դաշտը
      Call Rekvizit_Fill("Document", 1, "General", "FCOURSE", ForwardExchg)
      'Լրացնել "Գնվող գումար" դաշտը 
      Call Rekvizit_Fill("Document", 1, "General", "FSUMDB", PurAmount)
    End If 
       
    'Լրացնել "Պարտքերի ավտոմատ մարում" նշիչը
   wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame").VBObject("CheckBox").Value = AutoDebt
      
   Select Case DocType
        Case "²ñÅáõÃ³ÛÇÝ ëíá÷", "îáÏáë³¹ñáõÛù³ÛÇÝ ëíá÷"
          'Անցել 2.Տոկոսներ
          Set wTabStrip = wMDIClient.VBObject("frmASDocForm").VBObject("TabStrip")    
          wTabStrip.SelectedItem = wTabStrip.Tabs(2)
          'Լրացնել "Ստացվելիք տոկոսադրույք" դաշտը 
          Call Rekvizit_Fill("Document", 2, "General", "PCAGR", RecRate & "[Tab]" & Baj)
          'Լրացնել "Վճարվելիք տոկոսադրույք" դաշտը 
          Call Rekvizit_Fill("Document", 2, "General", "PCAGRCR", PayRate & "[Tab]" & Baj)
          
          TabN = 3
        Case "öáË³ñÅ»ù³ÛÇÝ ëíá÷", "üÛáõã»ñë", "üáñí³ñ¹"
          TabN = 2
    End Select
    
    If DocType = "îáÏáë³¹ñáõÛù³ÛÇÝ ëíá÷" Then
      'Անցնել 3.Ժամկետներ
      Set wTabStrip = wMDIClient.VBObject("frmASDocForm").VBObject("TabStrip")    
      wTabStrip.SelectedItem = wTabStrip.Tabs(3)
      'Լրացնել "Մարման Ժամկետ" դաշտը
      Call Rekvizit_Fill("Document", 3, "General", "DATEAGR", Term)
      'Լրացնել "Ամսաթվերի լրացում" նշիչը
      If DateFill = 1 Then
        wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame_3").VBObject("CheckBox_2").Click
          With Asbank.VBObject("frmAsUstPar").VBObject("TabFrame")
            ' Լրացնել "Սկզբի ամսաթիվ" դաշտը
            .VBObject("TDBDate").Keys(FirstDate & "[Tab]")
            ' Լրացնել "Նշ." նշիչը
            .VBObject("Checkbox_2").Value = CheckPayDates
            If CheckPayDates = 0 Then 
              ' Լրացնել "Պարբերություն" դաշտը
              .VBObject("AsCourse").VBObject("TDBNumber1").Keys(Paragraph & "[Tab]" & "[Tab]")
            Else  
              ' Լրացնել "Մարման օրեր" դաշտը
              .VBObject("TextC").Keys(PayDates & "[Tab]")
            End If   
            ' Լրացնել "Շրջանցման ուղղություն" դաշտը
            .VBObject("ASTypeTree").VBObject("TDBMask").Keys(Direction & "[Tab]")
            ' Սեղմել "Կատարել"
            Asbank.VBObject("frmAsUstPar").VBObject("CmdOK").ClickButton
          End With 
      End If
      TabN = 4
    End If  
      
    'Անցել 3(2).Լրացուցիչ
    Set wTabStrip = wMDIClient.VBObject("frmASDocForm").VBObject("TabStrip")    
    wTabStrip.SelectedItem = wTabStrip.Tabs(TabN)
    'Լրացնել "Ճյուղայնություն" դաշտը
    Call Rekvizit_Fill("Document", TabN, "General", "SECTOR", Sector)
    'Լրացնել "Օգտագործման ոլորտ(նոր ՎՌ)" դաշտը
    Call Rekvizit_Fill("Document", TabN, "General", "USAGEFIELD", UsageField)
    'Լրացնել "Նպատակ" դաշտը
    Call Rekvizit_Fill("Document", TabN, "General", "AIM", Aim)
    'Լրացնել "Ժամանակ" դաշտը
    Call Rekvizit_Fill("Document", TabN, "General", "TIME", Time)
    'Լրացնել "Առք/Վաճառք" դաշտը
    Call Rekvizit_Fill("Document", TabN, "General", "CUPUSA", PurSale)
    'Լրացնել "Գործողության վայր" դաշտը
    Call Rekvizit_Fill("Document", TabN, "General", "CURVAIR", OpPlace)
    Set Rekv = wMDIClient.VBObject("frmASDocForm").WaitVBObject("AS_LABELOPERFINPAY", delay_small)
    If Rekv.Exists Then
      'Լրացնել "Գործարքի վերջնահաշվարկի բնույթ" դաշտը
      Call Rekvizit_Fill("Document", TabN, "General", "OPERFINPAY", OpFinalPay)
    End If  
    'Լրացնել "Երկիր" դաշտը
    Call Rekvizit_Fill("Document", TabN, "General", "COUNTRY", Country)
    'Լրացնել "Մարզ" դաշտը
    Call Rekvizit_Fill("Document", TabN, "General", "LRDISTR", District)
    'Լրացնել "Մարզ(նոր ՎՌ)" դաշտը
    Call Rekvizit_Fill("Document", TabN, "General", "REGION", RegionLR)
    'Լրացնել "Պայմանագրի թղթային համար" դաշտը
    Call Rekvizit_Fill("Document", TabN, "General", "PPRCODE", PaperCode)
    'Վերցնել պայմանագրի ISN-ը
    fBASE = Asbank.VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmASDocForm").DocFormCommon.Doc.isn
    
    'Սեղմել "Կատարել"
    Asbank.VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmASDocForm").VBObject("CmdOk_2").Click
  End Sub

 'Պայմանագիրը ուղարկում է հաստատման
  Public Function SendToVerify(FolderPath)
    Call wTreeView.DblClickItem(FolderPath)
    With Asbank
      .VBObject("frmAsUstPar").VBObject("TabFrame").VBObject("AsTpComment").VBObject("TDBComment").Keys(DocNum & "[Tab]")
      .VBObject("frmAsUstPar").VBObject("CmdOK").ClickButton
    End With
    BuiltIn.Delay(2000) 
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_SendToVer)
    Asbank.VBObject("frmAsMsgBox").VBObject("cmdButton").ClickButton  
    wMDIClient.VBObject("frmPttel").Close
  End Function
  
  'Հաստատում է պայմանագիրը
  Public Function Verify(FolderPath) 
    BuiltIn.Delay(3000) 
    Call wTreeView.DblClickItem(FolderPath)
    With Asbank
      .VBObject("frmAsUstPar").VBObject("TabFrame").VBObject("AsTpComment").VBObject("TDBComment").Keys(DocNum & "[Tab]") 
      .VBObject("frmAsUstPar").VBObject("CmdOK").ClickButton
    End With
    BuiltIn.Delay(4000) 
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_ToConfirm)
    BuiltIn.Delay(4000) 
    With wMDIClient
      .VBObject("frmASDocForm").VBObject("CmdOk_2").ClickButton
       BuiltIn.Delay(3000) 
      .VBObject("frmPttel").Close
       BuiltIn.Delay(3000) 
    End With
  End Function 
  
  Function OpenInFolder(FolderPath)
    Call wTreeView.DblClickItem(FolderPath)
    With Asbank.VBObject("frmAsUstPar")
      .VBObject("TabFrame").VBObject("AsTpComment").VBObject("TDBComment").Keys(DocNum & "[Tab]")
      .VBObject("CmdOK").ClickButton
    End With
  End Function
End Class

Public Function New_DerivativeDoc()
  Set New_DerivativeDoc = New DerivativeDoc
End Function