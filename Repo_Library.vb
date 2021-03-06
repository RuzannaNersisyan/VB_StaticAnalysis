Option Explicit

'USEUNIT Library_Common
'USEUNIT Library_CheckDB
'USEUNIT Constants
'USEUNIT Mortgage_Library

'----------------------------------------------------------------------
' ՌԵՊՈ տեսակի պայմանագրի ստեղծում
'----------------------------------------------------------------------

'Client -Հաճախորդի կոդ
'Curr - Արժույթ
'CalcAcc - Հաշվաչկային հաշիվ
'Summa - Գումար
'Date - կնքման ամսաթիվ
'SecState - Արժ. պետականություն
'SecName - Արժեթղթի անվանում
'Nominal - Անվանական արժեք
'Price - Արժեք
'Kindscale  - Օրացույցի հաշվարկան ձև
'Percent  - Ռեպոյի տոկոսադրույք
'GiveDate - Հատկացման ամսաթիվ
'Term - Մարման ժամկետ
'DateFill - Ամսաթվերի լրացում
'CheckPayDates - Նշ. նշիչ
'PayDates - Մարման օրեր
'Paragraph - Պարբերություն
'Direction - Շրջանցման ուղղություն
'DocNum - Պայմնագրի N
'fBASE - Պայմնագրի ISN
 
Sub Repo_Create(Client, Curr, CalcAcc, Summa, Date, SecState, SecName, Nominal, Price,_
                 Kindscale, Percent, GiveDate, Term, DateFill, CheckPayDates, PayDates,_
                 Paragraph, Direction, Sector, Aim, Country, District, Region, PaperCode,_
                 fBASE, DocNum)
    Dim state
    BuiltIn.Delay(delay_middle)
    Call wTreeView.DblClickItem("|è»åá Ñ³Ù³Ó³ÛÝ³·ñ»ñ|Üáñ å³ÛÙ³Ý³·ñÇ ëï»ÕÍáõÙ")
    BuiltIn.Delay(delay_middle)
    
    fBASE = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
    
    'Պայմանագրի համարի վերագրում փոփոխականին
    DocNum = Get_Rekvizit_Value("Document",1,"General","CODE")
    
    Log.Message(DocNum)
    'Լրացնել "Հաճախորդ" դաշտը
    Call Rekvizit_Fill("Document", 1, "General", "CLICOD", Client) 
    'Լրացնել "Արժույթ" դաշտը  
    Call Rekvizit_Fill("Document", 1, "General", "CURRENCY", Curr)
    'Լրացնել "Հաշիվ" դաշտը  
    Call Rekvizit_Fill("Document", 1, "General", "ACCACC", CalcAcc)
    'Լրացնել "Գումար" դաշտը      
    Call Rekvizit_Fill("Document", 1, "General", "SUMMA", Summa)
    'Լրացնել "Կնքման ամսաթիվ" դաշտը          
    Call Rekvizit_Fill("Document", 1, "General", "DATE", Date)
    'Լրացնել "Հատկացման ամսաթիվ"
    Call Rekvizit_Fill("Document", 5, "General", "DATEGIVE", GiveDate) 
    'Լրացնել "Արժ.պետականություն" դաշտը          
    Call Rekvizit_Fill("Document", 2, "General", "SECSTATE", SecState)
    
    state = False
    With wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame_2").VBObject("DocGrid")
        .Row = 0
        .Col = 0
        .Keys("^[Down]")
    
        Do Until p1.frmModalBrowser.vbObject("tdbgView").EOF
            If SecName = p1.frmModalBrowser.vbObject("tdbgView").Columns.Item(1).Text Then
                Call p1.frmModalBrowser.vbObject("tdbgView").Keys("[Enter]")
                state = True
                Exit Do
            Else
                Call p1.frmModalBrowser.vbObject("tdbgView").MoveNext
            End If
        Loop

        If state Then
             BuiltIn.Delay(2000)
            .Row = 0
            .Col = 2
            .Keys(Nominal & "[Enter]")
      
            .Row = 0
            .Col = 3
            .Keys(Price & "[Enter]" )
        Else  
            Log.Error("²Ýí³ÝáõÙ ¹³ßïÁ ãÇ Éñ³óí»É")
            p1.VBObject("frmModalBrowser").Close
        End If
    End With 
    
    'Լրացնել "Օրացույցի հաշվարկման ձև" դաշտը      
    Call Rekvizit_Fill("Document", 3, "General", "KINDSCALE", Kindscale) 
    'Լրացնել "Ռեպոյի տոկոսադրույք" դաշտը      
    Call Rekvizit_Fill("Document", 3, "General", "PCAGR", Percent & "[Tab]" & "365") 
    
    'Լրացնել "Մարման ժամկետ"
    Call Rekvizit_Fill("Document", 5, "General", "DATEAGR", Term) 

    'Լրացնել "Ամսաթվերի լրացում" նշիչը
      If DateFill = 1 Then
          wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame_5").VBObject("CheckBox_4").Click
          'Լրացնել "Նշ." նշիչը
          Call Rekvizit_Fill("Dialog", 1, "CheckBox", "INCLFIXD", CheckPayDates) 
          If CheckPayDates = 1 Then
              'Լրացնել "Մարման օրեր" դաշտը
              Call Rekvizit_Fill("Dialog", 1, "General", "FIXEDDAYS", PayDates)
          Else
              'Լրացնել "Պարպերություն" դաշտը
              Call Rekvizit_Fill("Dialog", 1, "General", "PERIODICITY", Paragraph & "[Tab]" & "[Tab]")
          End If
          'Լրացնել "Շրջանցման ուղղություն" դաշտը
          Call Rekvizit_Fill("Dialog", 1, "General", "PASSOVDIRECTION", Direction)
          'Սեղմել "Կատարել"
          Call ClickCmdButton(2, "Î³ï³ñ»É")
      End If

    'Լրացնել "Ճյուղայնություն"  դաշտը
    Call Rekvizit_Fill("Document", 6, "General", "SECTOR", Sector) 
    'Լրացնել "Նպատակ"  դաշտը
    Call Rekvizit_Fill("Document", 6, "General", "AIM", Aim) 
    'Լրացնել "Երկիր"  դաշտը
    Call Rekvizit_Fill("Document", 6, "General", "COUNTRY", Country) 
    'Լրացնել "Մարզ"  դաշտը
    Call Rekvizit_Fill("Document", 6, "General", "LRDISTR", District) 
    'Լրացնել "Մարզ(նոր ՎՌ)"  դաշտը
    Call Rekvizit_Fill("Document", 6, "General", "REGION", Region) 
    'Լրացնել "Պայմ.թղթային N"  դաշտը
    Call Rekvizit_Fill("Document", 6, "General", "PPRCODE", PaperCode) 
    
    Call ClickCmdButton(1, "Î³ï³ñ»É")
End Sub

'----------------------------------------------------------------------
' ä³ÛÙ³Ý³·ñÇ àõÕ³ñÏáõÙ Ñ³ëï³ïÙ³Ý
'----------------------------------------------------------------------
Function Repo_Send_To_Verify()
    BuiltIn.Delay(1000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_SendToVer)
    p1.vbObject("frmAsMsgBox").vbObject("cmdButton").Click()
    wMDIClient.vbObject("frmPttel").Close()
End Function

'----------------------------------
'Ռեպո- ի տրամադրում
'----------------------------------
Sub Repo_Provide(provide_date)
    BuiltIn.Delay(2000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_Opers & "|" & c_GiveAndBack & "|" & c_GiveRepo)
    BuiltIn.Delay(2000)
    Call Rekvizit_Fill("Document", 1, "General", "DATE", provide_date) 
    Call ClickCmdButton(1, "Î³ï³ñ»É")
    Call ClickCmdButton(5, "²Ûá")
    Call Close_Pttel("frmPttel")
End Sub

'----------------------------------
'Արժեթղթերի վաճառք գործողության կատարում
'op_date - վաճառքի ամսաթիվ
'op_sum - Արժեթղթի անվանական արժեք
'cash_or_no - Կանխիկ/անկանխիկ
'acc - հաշիվ
'sec_name - Արժեթղթի անվանում
'----------------------------------
Sub Repo_Sell_Security(op_date, op_sum, cash_or_no, acc)
    Dim wTabFrame
    
    BuiltIn.Delay(2000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_Opers & "|" & c_GiveAndBack & "|" & c_SecSell)
    BuiltIn.Delay(2000)
    
    'Ամսաթիվ դաշտի լրացում
    Call Rekvizit_Fill("Document",1,"General","DATE","![End]" & "[Del]" & op_date)
     
    Set wTabFrame = wMDIClient.vbObject("frmASDocForm").vbObject("TabFrame")
    
    'Գրիդի լրացում
    wTabFrame.vbObject("DocGrid").Row = 0
    wTabFrame.vbObject("DocGrid").Col = 1
    Call wTabFrame.vbObject("DocGrid").Keys(op_sum & "[Enter]" )

    'Կանխիկ/Անկանխիկ դաշտի լրացում
    Call Rekvizit_Fill("Document", 1, "General", "CASHORNO", cash_or_no)   
    'Հաշիվ դաշտի լրացում
    Call Rekvizit_Fill("Document", 1, "General", "ACCCORR", acc)
    
    If p1.WaitVBObject("frmAsMsgBox",1000).Exists Then
        Call ClickCmdButton(5, "Î³ï³ñ»É")
    End If
    
    'Կատարել կոճակի սեղմում
    Call ClickCmdButton(1, "Î³ï³ñ»É")
    Call ClickCmdButton(5, "²Ûá")

    Call Close_Pttel("frmPttel")
End Sub

'----------------------------------
'Պարտքերի մարում գործողության կատարում
'----------------------------------
'Date - մարման ամսաթիվ
'Sum - Հիմնական Գումար
'PerSum - Տոկոսագումար
Sub Repo_Repayment(Date, Sum, PerSum)
    BuiltIn.Delay(2000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_Opers & "|" & c_GiveAndBack & "|" & c_PayOffDebt)
    'Լրացնել Ամսաթիվ
    Call Rekvizit_Fill("Document", 1, "General", "DATE", Date)
    'Լրացնել Հիմնական գումար
    Call Rekvizit_Fill("Document", 1, "General", "SUMAGR", Sum)
    'Լրացնել Տոկոսագումար դաշտը
    Call Rekvizit_Fill("Document", 1, "General", "SUMPER", PerSum)
    'Լրացնել Մարման աղբյուր դաշտը
    Call Rekvizit_Fill("Document", 2, "General", "REPSOURCE", 1)
    Call ClickCmdButton(1, "Î³ï³ñ»É")
    Call ClickCmdButton(5, "²Ûá")
    
    Call Close_Pttel("frmPttel")
End Sub

'----------------------------------
'²ñÅ»ÃÕÃ»ñÇ ³éù ·áñÍáÕáõÃÛ³Ý Ï³ï³ñáõÙ
'op_date - í³×³éùÇ ³Ùë³ÃÇí
'op_sum - ²ñÅ»ÃÕÃÇ ³Ýí³Ý³Ï³Ý ³ñÅ»ù
'cash_or_no - Î³ÝËÇÏ/³ÝÏ³ÝËÇÏ
'acc - Ñ³ßÇí
'sec_name - ²ñÅ»ÃÕÃÇ ³Ýí³ÝáõÙ
'----------------------------------                 
Sub Repo_Buy_Security(op_date, op_sum, cash_or_no, acc)
    Dim wTabFrame
    BuiltIn.Delay(2000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_Opers & "|" & c_GiveAndBack & "|" & c_SecBuy)
    BuiltIn.Delay(2000)
    
    '²Ùë³ÃÇí ¹³ßïÇ Éñ³óáõÙ
    Call Rekvizit_Fill("Document", 1, "General", "DATE", "!" & "[End]" & "[Del]" & op_date)
    
    'Գրիդի լրացում
    Set wTabFrame = wMDIClient.vbObject("frmASDocForm").vbObject("TabFrame")
    wTabFrame.vbObject("DocGrid").Row = 0
    wTabFrame.vbObject("DocGrid").Col = 1
    Call wTabFrame.vbObject("DocGrid").Keys(op_sum & "[Enter]")
    
    'Կանխիկ/Անկանխիկ դաշտի լրացում
    Call Rekvizit_Fill("Document", 1, "General", "CASHORNO", cash_or_no)   
    'Հաշիվ դաշտի լրացում
    Call Rekvizit_Fill("Document", 1, "General", "ACCCORR", acc)
    
    If p1.WaitVBObject("frmAsMsgBox",1000).Exists Then
        Call ClickCmdButton(5, "Î³ï³ñ»É")
    End If
    
    'Կատարել կոճակի սեղմում
    Call ClickCmdButton(1, "Î³ï³ñ»É")
    Call ClickCmdButton(5, "²Ûá")

    Call Close_Pttel("frmPttel")
End Sub

'----------------------------------------------------------------
' Ð³Ï³¹³ñÓ èºäà ï»ë³ÏÇ å³ÛÙ³Ý³·ñÇ ëï»ÕÍáõÙ
'----------------------------------------------------------------
'client - Հաճախորդ
'curr - Արժույթ
'acc - Հաշիվ
'summa - Գումար
'date - Կնքման ամսաթիվ
'kindscale - Օրացույցի Հաշվարկման ձև
'per, 'baj - Ռեպոյի տոկոսադրույք
'dateGive - Հատկացման ամսաթիվ
'dateAgr - Մարման ժամկետ
'DateFill - "Ամսաթվերի լրացում" նշիչը
'CheckPayDates - "Նշ." նշիչ
'PayDates - Մարման օրեր
'Paragraph - Պարպերություն
'Direction - Շրջանցման ուղղություն
'secState - Արժ. պետականություն
'secClass - Արժեթղթի դաս
Sub Inverse_Repo_Create(client, curr, acc, summa, date, kindscale,per, baj, dateGive, dateAgr, DateFill,startDate,CheckPayDates, _
                          PayDates, Paragraph, Direction ,secState, secClass, security,Price,fBASE, docNum)
  
  'Ստանում է պայմանագրի ISN 
  fBASE = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
  'Ստանում է պայմանագրի համարը
  
  docNum = Get_Rekvizit_Value("Document", 1, "General", "CODE")
  
  'Լրացնում է Հաճախորդ դաշտը
  Call Rekvizit_Fill("Document",1,"General","CLICOD",client)
  'Լրացնում է արժույթ դաշտը
  Call Rekvizit_Fill("Document",1,"General","CURRENCY",curr)
  'Լրացնում է Հաշիվ դաշտը
  Call Rekvizit_Fill("Document",1,"General","ACCACC",acc)
  'Լրացնում է Գումար դաշտը
  Call Rekvizit_Fill("Document",1,"General","SUMMA",summa & "[Tab]")
  'Լրացնում է Կնքման ամսաթիվ դաշտը
  Call Rekvizit_Fill("Document",1,"General","DATE","^!A[Del]" & date)
  'Լրացնում է Օրացույցի Հաշվարկման ձև
  Call Rekvizit_Fill("Document",2,"General","KINDSCALE",kindscale)
 'Լրացնում է Ռեպոյի տոկոսադրույք դաշտը
  Call Rekvizit_Fill("Document",2,"General","PCAGR",per & "[Tab]" & baj)

  'Լրացնում է Հատկացման ամսաթիվ դաշտը
  Call Rekvizit_Fill("Document",4,"General","DATEGIVE","^!A[Del]" & dateGive)
  'Լրացնում է Մարման ժամկետ դաշտը
  Call Rekvizit_Fill("Document",4,"General","DATEAGR","^!A[Del]" & dateAgr)
  
  'Լրացնել "Ամսաթվերի լրացում" նշիչը
  If DateFill = 1 Then
    wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame_4").VBObject("CheckBox_4").Click
    'Լրացնել "Նշ." նշիչը
    p1.VBObject("frmAsUstPar").VBObject("TabFrame").VBObject("TDBDate").Keys(startDate)
    'Call Rekvizit_Fill("Dialog", 1, "CheckBox", "INCLFIXD", CheckPayDates) 
    If CheckPayDates = 1 Then
      'Լրացնել "Մարման օրեր" դաշտը
      Call Rekvizit_Fill("Dialog", 1, "General", "FIXEDDAYS", PayDates)
    Else
      'Լրացնել "Պարպերություն" դաշտը
      Call Rekvizit_Fill("Dialog", 1, "General", "PERIODICITY", Paragraph & "[Tab]" & "[Tab]")
    End If
    'Լրացնել "Շրջանցման ուղղություն" դաշտը
    Call Rekvizit_Fill("Dialog", 1, "General", "PASSOVDIRECTION", Direction)
    'Սեղմել "Կատարել"
    Call ClickCmdButton(2, "Î³ï³ñ»É")
  End If
  'Լրացնում է Արժ. պետականություն դաշտը
  Call Rekvizit_Fill("Document",5,"General","SECSTATE",secState)
  'Լրացնում է Արժեթղթի դաս դաշտը
  Call Rekvizit_Fill("Document",5,"General","SECFUNC",secClass)

  If secClass <> "6" Then
    With wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame_5").VBObject("DocGrid_2")
      .Row = 0
      .Col = 0
      .Keys("^[Down]")
    
      Do Until p1.frmModalBrowser.vbObject("tdbgView").EOF
          If Trim(security) = Trim(p1.frmModalBrowser.vbObject("tdbgView").Columns.Item(1).Text) Then
              Call p1.VBObject("frmModalBrowser").vbObject("tdbgView").Keys("[Enter]")
              Exit Do
          Else
              Call p1.frmModalBrowser.vbObject("tdbgView").MoveNext
          End If
      Loop
      If secState = 2 Then
          .Row = 0
          .Col = 4
          .Keys("15" & "[Enter]" )
      End If
      .Row = 0
      .Col = 3
      .Keys(Price & "[Enter]" )
    End With 
    
  Else  
    Set wMDIClient.VBObject("frmASDocForm").VBObject("TabStrip").SelectedItem = wMDIClient.VBObject("frmASDocForm").VBObject("TabStrip").Tabs(6)
    With wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame_6").VBObject("DocGrid_3")
      .Row = 0
      .Col = 0
      .Keys("^[Down]")
      
       Do Until p1.VBObject("frmModalBrowser").VBObject("tdbgView").EOF
          If Trim(security) = Trim(p1.frmModalBrowser.vbObject("tdbgView").Columns.Item(1).Text) Then
              Call p1.VBObject("frmModalBrowser").vbObject("tdbgView").Keys("[Enter]")
              Exit Do
          Else
              Call p1.frmModalBrowser.vbObject("tdbgView").MoveNext
          End If
      Loop
      
    End With  
  End If
  
  wMDIClient.VBObject("frmASDocForm").VBObject("CmdOk_2").Click()
End Sub    

'----------------------------------------------------------------------------------------
'Հակադարձ Ռեպոյի ներգրավում գործողություն
'Date - Ամսաթիվ
'----------------------------------------------------------------------------------------
Function InverseRepoAttraction(Date)
    BuiltIn.Delay(2000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_Opers & "|" & c_GiveAndBack & "|" & c_InvolveRepo)
    
    InverseRepoAttraction = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
    Call Rekvizit_Fill("Document", 1, "General", "DATE", "^!A[Del]" & Date) 
        
    Call ClickCmdButton(1, "Î³ï³ñ»É")
    Call ClickCmdButton(5, "²Ûá")
End Function

'--------------------------------------------------------------------------
'Հակադարձ Ռեպոյի Մարում գործողություն
'Date - Ամսաթիվ
'Sum - Հիմնական գումար
'PerSum - Տոկոսագումար
'CashOrNo - Կանխիկ/Անկանխիկ
'--------------------------------------------------------------------------
Function InverseRepoRepay(Date, Sum, PerSum, CashOrNo)
    BuiltIn.Delay(2000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_Opers & "|" & c_GiveAndBack & "|" & c_PayOffDebt)
    
    InverseRepoRepay = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
    
    'Լրացնել "Ամսաթիվ" դաշտը
    Call Rekvizit_Fill("Document", 1, "General", "DATE", Date) 
    'Լրացնել "Հիմնական գումար" դաշտը
    Call Rekvizit_Fill("Document", 1, "General", "SUMAGR", Sum)
    'Լրացնել "Տոկոսագումար" դաշտը
    Call Rekvizit_Fill("Document", 1, "General", "SUMPER", PerSum)
    'Լրացնել "Կանխիկ/Անկանխիկ" դաշտը
    Call Rekvizit_Fill("Document", 1, "General", "CASHORNO", CashOrNo)
 
    Call ClickCmdButton(1, "Î³ï³ñ»É")
    Call ClickCmdButton(5, "²Ûá")
End Function                                                                                

''------------------------------------------------------------------------------------------------------------
'"Թղթապանակներ/Ռեպոյով գնված արժեթղթեր" թղթապանակում դաշտերի ստուգում
'------------------------------------------------------------------------------------------------------------
'allCost - Անվանական արժեք
'secCost - Վաճառքի մասի անվ. արժեք
'repoSold_cost - Հակ.ռեպոյով վաճառված մաս անվ. արժեք
'cost - Արժեք
'soldCost - Վաճառված մասի արժեք
'partSoldCost - Հակ. Ռեպոյի վաճ. մասի արժեք
Sub Repo_Check_Col_Value(allCost, secCost, repoSoldCost, cost, soldCost, partSoldCost)
      
    Dim colNum
    BuiltIn.Delay(2000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_Folders & "|" & c_SecsFromRepo)
    BuiltIn.Delay(4000)
    
    If wMDIClient.VBObject("frmPttel_2").VBObject("tdbgView").ApproxCount = 1 Then
          'Ստուգում է "Անվանական արժեք" սյան արժեքը
          colNum =	wMDIClient.VBObject("frmPttel_2").GetColumnIndex("fNOMINAL")
          If Not Trim(wMDIClient.VBObject("frmPttel_2").VBObject("tdbgView").Columns.Item(colNum)) = Trim(allCost) Then
              Log.Message(wMDIClient.VBObject("frmPttel_2").VBObject("tdbgView").Columns.Item(colNum))
              Log.Message(allCost)
              Log.Error("Անվանական արժեքը սխալ է")
          End If
          'Ստուգում է "Վաճառքի մասի անվ. արժեք: սյան արժեքը
          colNum =	wMDIClient.VBObject("frmPttel_2").GetColumnIndex("fSELLNOMINAL")
          If Not Trim(wMDIClient.VBObject("frmPttel_2").VBObject("tdbgView").Columns.Item(colNum)) = Trim(secCost) Then
              Log.Message(wMDIClient.VBObject("frmPttel_2").VBObject("tdbgView").Columns.Item(colNum))
              Log.Message(sec_cost)
              Log.Error("Արժեթղթի արժեք սխալ է")
          End If
          'Ստուգում է "Հակ.ռեպոյով վաճառված մաս անվ. արժեք" սյան արժեքը
          colNum =	wMDIClient.VBObject("frmPttel_2").GetColumnIndex("fHRNOMINAL")
          If Not Trim(wMDIClient.VBObject("frmPttel_2").VBObject("tdbgView").Columns.Item(colNum)) = Trim(repoSoldCost) Then
              Log.Message(wMDIClient.VBObject("frmPttel_2").VBObject("tdbgView").Columns.Item(colNum))
              Log.Message(repo_sold_cost)
              Log.Error("Հակ.ռեպոյով վաճառված մաս արժեքը սխալ է")
          End If
            
          'Ստուգում է "Արժեք" սյան արժեքը
          colNum =	wMDIClient.VBObject("frmPttel_2").GetColumnIndex("fPRICE")
          If Not Trim(wMDIClient.VBObject("frmPttel_2").VBObject("tdbgView").Columns.Item(colNum)) = Trim(cost) Then
              Log.Message(wMDIClient.VBObject("frmPttel_2").VBObject("tdbgView").Columns.Item(colNum))
              Log.Message(cost)
              Log.Error("Վաճառված մաս արժեքը սխալ է")
          End If
            
          'Ստուգում է "Վաճառված մասի արժեք" սյան արժեքը
          colNum =	wMDIClient.VBObject("frmPttel_2").GetColumnIndex("fSELLPRICE")
          If Not Trim(wMDIClient.VBObject("frmPttel_2").VBObject("tdbgView").Columns.Item(colNum)) = Trim(soldCost) Then
              Log.Message(wMDIClient.VBObject("frmPttel_2").VBObject("tdbgView").Columns.Item(colNum))
              Log.Message(sold_cost)
              Log.Error("Վաճառված մաս արժեքը սխալ է")
          End If
            
          'Ստուգում է "Հակ. Ռեպոյի վաճ. մասի արժեք" սյան արժեքը
          colNum =	wMDIClient.VBObject("frmPttel_2").GetColumnIndex("fHRPRICE")
          If Not Trim(wMDIClient.VBObject("frmPttel_2").VBObject("tdbgView").Columns.Item(colNum)) = Trim(partSoldCost) Then
              Log.Message(wMDIClient.VBObject("frmPttel_2").VBObject("tdbgView").Columns.Item(colNum))
              Log.Message(part_sold_cost)
              Log.Error("Վաճառված մաս արժեքը սխալ է")
          End If
      Else 
            Log.Error("Ռեպոյով գնված արժեթղթեր թղթապանակը դատարկ է")
      End If        
    wMDIClient.VBObject("frmPttel_2").Close()
End Sub

'----------------------------------------------------------------------------------------
'Ռեպո համ-ով գնված արժ.գնի ճշտում
'----------------------------------------------------------------------------------------
Sub Repo_Secur_Cost_Correction(cDate,Date)
    BuiltIn.Delay(2000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(  c_Opers & "|" &  c_RepSecOvervalue)
    
    Call Rekvizit_Fill("Document",1,"General","DATECHARGE",cDate)
    Call Rekvizit_Fill("Document",1,"General","DATE",Date )
    Call ClickCmdButton(1, "Î³ï³ñ»É")
End Sub

'-----------------------------------------------------------------------------------------------------------------
' " Տոկոսադրույք" գործողության կատարում
'-----------------------------------------------------------------------------------------------------------------
'date - Անսաթիվ
'per - Ստացվելիք տոկոսադրոյք
'part -  բաժ`
'sPer - Վճարվելիք տոկոսադրույք
'sPart - բաժ`
Sub Set_Persentage_Repo(fBase,date,per,part)
    BuiltIn.Delay(2000)
    Call p1.VBObject("MainForm").MainMenu.Click(c_AllActions)
    Call p1.VBObject("MainForm").PopupMenu.Click(c_TermsStates & "|" & c_Percentages & "|" & c_Percentages)
   
    fBase = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
    'Լրացնել Ամսաթիվ դաշտը
    Call Rekvizit_Fill("Document",1,"General","DATE","^!A[Del]" & date)
    'Լրացնել Ստացվելիք տոկոսադրոյք դաշտը
    Call Rekvizit_Fill("Document",1,"General","PCAGR",per)
    'Լրացնել բաժ` դաշտը
    Call Rekvizit_Fill("Document",1,"General","PCAGR",part)
    'Սեղմել Կատարել կոճակը
    Call ClickCmdButton(1,"Î³ï³ñ»É")
End Sub

'-----------------------------------------------
'Արժեթղթի փոփոխություն գործողության կատարում
'-----------------------------------------------
'date - Ամսաթիվ
'oldSec - Հին արժեթուղթ
'newSec - Նոր արժեթուղթ
'oldSecCost - Հին արժեթղթի անվ.արժեք
'newSecCost - Նոր արժեթղթի անվ.արժեք
Sub Change_Security(date,oldSec,newSec,oldSecCost,newSecCost)
    
    BuiltIn.Delay(2000)
    Call wMainForm.MainMenu.Click(c_AllActions )
    Call wMainForm.PopupMenu.Click(c_Opers & "|" & c_SecChng)
          
    'Լրացնում է ամսաթիվ դաշտը
    Call Rekvizit_Fill("Document",1,"General","DATE","^!A[Del]" & date)
    'Լրացնում է Հին արժեթուղթ դաշտը
    wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame").VBObject("AsTpComment").VBObject("CmdViewHelp").Click()
    If oldSec = Trim(p1.VBObject("frmModalBrowser").VBObject("tdbgView").Columns.Item(1).Text) then
        Call p1.VBObject("frmModalBrowser").vbObject("tdbgView").Keys("[Enter]")
    Else
        Log.Error("Old security not found")
    End If
    'Լրացնում է նոր արժեթուղթ դաշտը
    wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame").VBObject("AsTpComment_3").VBObject("CmdViewHelp").Click()
    With p1.VBObject("frmModalBrowser").VBObject("tdbgView")
      .Row = 0
      .Col = 0
      .Keys("^[Down]")
    
      Do Until p1.frmModalBrowser.vbObject("tdbgView").EOF
          If Trim(newSec) = Trim(p1.frmModalBrowser.vbObject("tdbgView").Columns.Item(1).Text) Then
              Call p1.VBObject("frmModalBrowser").vbObject("tdbgView").Keys("[Enter]")
              Exit Do
          Else
              Call p1.frmModalBrowser.vbObject("tdbgView").MoveNext
          End If
      Loop
    End With
    Call Rekvizit_Fill("Document",1,"General","NEWSECCODE","[Tab]" & "[Tab]")
    
    'Լրացնում է Հին արժեթղթի անվ. արժեք դաշտը
    Call Rekvizit_Fill("Document",1,"General","OLDSUM",oldSecCost)
    'Լրացնում է Նոր արժեթղթի անվ. արժեք դաշտը
    Call Rekvizit_Fill("Document",1,"General","NEWSUM",newSecCost)
    'Սեղմում է Կատարել կոճակը
    Call ClickCmdButton(1,"Î³ï³ñ»É")
End Sub

'-----------------------------------------------------------------------------------
'Գումարի տեղափոխում գործողության կատարում
'------------------------------------------------------------------------------------
'docNum - Նոր պայմանագրի N 
'date - Ամսաթիվ
'sumAgr -  Հիմնական գումար 
'sumPer - Տոկոսագումար
Sub Repo_Sum_Transfer(fBASE,docNum,date,sumAgr,sumPer)

    BuiltIn.Delay(2000)
    Call wMainForm.MainMenu.Click(c_AllActions )
    Call wMainForm.PopupMenu.Click(c_Opers & "|" & c_PassSums)
    fBASE = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
    
    'Լրացնում է Նոր պայմանագրի N դաշտը
    Call Rekvizit_Fill("Document",1,"General","CODE2",docNum)
    'Լրացնում է Ամսաթիվ դաշտը
    Call Rekvizit_Fill("Document",1,"General","DATE",date)
     'Լրացնում է Հիմնական գումար դաշտը
    Call Rekvizit_Fill("Document",1,"General","SUMAGR",sumAgr)
    'Լրացնում է Տոկոսագումար դաշտը
    Call Rekvizit_Fill("Document",1,"General","SUMPER",sumPer)
    'Սեղմում է Կատարել կոճակը
    Call ClickCmdButton(1,"Î³ï³ñ»É")
End Sub