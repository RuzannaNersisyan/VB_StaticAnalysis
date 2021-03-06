'USEUNIT Library_Common
'USEUNIT Payment_Order_ConfirmPhases_Library
'USEUNIT Online_PaySys_Library
'USEUNIT Constants
'USEUNIT Credit_Line_Library
'USEUNIT Akreditiv_Library
'USEUNIT Mortgage_Library

'--------------------------------------------------------------------------------------
'Üáñ å³ÛÙ³Ý³·ñÇ ëï»ÕÍáõÙ ó³ÝÏÇó í³ñÏ³ÛÇÝ å³ÛÝ³·ñÇ ÁÝïñáõÙ
'--------------------------------------------------------------------------------------
'creditType - ì³ñÏ³ÛÇÝ å³ÛÙ³Ý³·ñÇ ï»ë³Ï

Sub Select_Credit_Type(creditType)
    
    Call ChangeWorkspace("Վարկեր (տեղաբաշխված)")
    Call wTreeView.DblClickItem("|ì³ñÏ»ñ (ï»Õ³µ³ßËí³Í)|Üáñ å³ÛÙ³Ý³·ñÇ ëï»ÕÍáõÙ")
    
    Do Until p1.frmModalBrowser.vbObject("tdbgView").EOF
        If RTrim(p1.frmModalBrowser.vbObject("tdbgView").Columns.Item(1).Text) = creditType Then
            Call p1.frmModalBrowser.vbObject("tdbgView").Keys("[Enter]")
            Exit Do
        Else
            Call p1.frmModalBrowser.vbObject("tdbgView").MoveNext
        End If
    Loop
    
End Sub

'--------------------------------------------------------------------------------------
'Պարտքերի մարում գործողույան կատարում
'--------------------------------------------------------------------------------------
'fadeDate - Պարտքերի մարման ամսաթիվ դաշտի արժեք
'fadeBase - Պարտքերի մարման փաստաթղթի ISN
'date1 - Հետ. պարտք դասշի արժեք
'percSumma - Տոկոսագումար դաշտի արժեք
'beforeTerm - true` մարումը կատարվում է ժամկետից շուտ
Sub Fade_Debt(fadeDate, fadeBase, date1,mainSumma, percSumma, beforeTerm) 
  Dim Str,Rekv

  BuiltIn.Delay(6000)
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_Opers & "|" & c_GiveAndBack & "|" & c_PayOffDebt)
   
  fadeBase = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
  If wMDIClient.WaitVBObject("frmASDocForm", 3000).Exists Then
    'Ամսաթիվ դաշտի լրացում
    Call Rekvizit_Fill("Document", 1, "General", "DATE", fadeDate)
    Set Rekv = wMDIClient.VBObject("frmASDocForm").WaitVBObject("AS_LABELDATEFOROFF", delay_small)
    If Rekv.Exists Then
      'Պարտքի ամսաթիվ դաշտի լրացում
      Call Rekvizit_Fill("Document", 1, "General", "DATEFOROFF", date1)
    End If  
    'Հիմնական գումար դաշտի լրացում
    Call Rekvizit_Fill("Document", 1, "General", "SUMAGR", mainSumma)   
    Set Rekv = wMDIClient.VBObject("frmASDocForm").WaitVBObject("AS_LABELSUMPER", delay_small)
    If Rekv.Exists Then
      'Տոկոսագումար դաշտի լրացում
      Call Rekvizit_Fill("Document", 1, "General", "SUMPER", percSumma)
    End If  
    'Անցում 3.Այլ էջին
    '"Մարման աղբյուր դաշտի լրացում"
    Set Rekv = wMDIClient.VBObject("frmASDocForm").WaitVBObject("AS_LABELREPSOURCE", delay_small)
    If Rekv.Exists Then
      Call Rekvizit_Fill("Document",3,"General","REPSOURCE", 1)
    End If  
    'Կատարել կոճակի սեղմում
    Call ClickCmdButton(1, "Î³ï³ñ»É")
    If beforeTerm Then
      Call ClickCmdButton(5, "Î³ï³ñ»É")
    End If 
    Call ClickCmdButton(1, "Î³ï³ñ»É")
    Call ClickCmdButton(5, "²Ûá")
  Else 
    Log.Error "Can't open frmASDocForm window", "", pmNormal, ErrorColor
  End If
End Sub

'--------------------------------------------------------------------------------------
'Ø³ñáõÙÝ»ñÇ ·ñ³ýÇÏÇ Ýß³Ý³ÏáõÙ ä³ÛÙ³Ý³·ñÇ ÃÕÃ³å³Ý³ÏÇó : üáõÝÏóÇ³Ý  í»ñ³¹³ñÓÝáõÙ ¿ true,
'»Ã» Ù³ñáõÙÝ»ñÇ ·ñ³ýÇÏÁ Ýß³Ý³ÏíáõÙ ¿ , ¨ false ` Ñ³Ï³é³Ï ¹»åùáõÙ :
'--------------------------------------------------------------------------------------
Function Fade_Schedule()
  Dim isExists : isExists = False
    
  BuiltIn.Delay(3000)
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_RepaySchedule)
    
  BuiltIn.Delay(1000)
  If wMDIClient.WaitVBObject("frmPttel", 3000).Exists Then
    Do Until wMDIClient.vbObject("frmPttel").vbObject("tdbgView").EOF
      If Left(Trim(wMDIClient.vbObject("frmPttel").vbObject("tdbgView").Text), 17) = "Ø³ñáõÙÝ»ñÇ ·ñ³ýÇÏ" Then
        isExists = True
        Exit Do
      Else 
        Call wMDIClient.vbObject("frmPttel").vbObject("tdbgView").MoveNext
      End If
    Loop
  Else
    Log.Message "The sending documnet frmPttel doesn't exist", "", pmNormal, ErrorColor
  End If

  Fade_Schedule = isExists
End Function

'--------------------------------------------------------------------------------------
'²ÛÉ í×³ñáõÙÝ»ñÏ ·ñ³ýÇÏÇ Ýß³Ý³ÏáõÙ ä³ÛÙ³Ý³·ñÇ ÃÕÃ³å³Ý³ÏÇó: ºÃ» ·ñ³ýÇÏÁ ëï»ÕÍíáõÙ ¿ ýáõÝÏóÇ³Ý
' í»ñ³¹³ñÓÝáõÙ ¿ true, »Ã» áã` false
'--------------------------------------------------------------------------------------
'griddate - ¶ñÇ¹Ç ³Ùë³ÃÇí ¹³ßïÇ ³ñÅ»ù
'summ - ¶ñÇ¹Ç ·áõÙ³ñ ¹³ßïÇ ³ñÅ»ù
Function Other_Payment_Schedule(griddate, summ)
  Dim gridData, gridSumma, isExist
  isExist = False
    
  Do Until wMDIClient.vbObject("frmPttel").vbObject("tdbgView").EOF
    If Left(Trim(wMDIClient.vbObject("frmPttel").vbObject("tdbgView").Text), 28) = "¶ñ³ýÇÏáí í³ñÏ³ÛÇÝ å³ÛÙ³Ý³·Çñ" Then
      BuiltIn.Delay(3000)
      Call wMainForm.MainMenu.Click(c_AllActions)
      Call wMainForm.PopupMenu.Click(c_OtherPaySchedule)
      Exit Do
    Else
      Call wMDIClient.vbObject("frmPttel").vbObject("tdbgView").MoveNext
    End If
  Loop
    
  BuiltIn.Delay(2000)
  gridData = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.Grid("DATES").NumFromName("DATEAGR")
  gridSumma = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.Grid("DATES").NumFromName("OTHERSUM")
  With wMDIClient.vbObject("frmASDocForm").vbObject("TabFrame").vbObject("DocGrid")
    '¶ñÇ¹Ç ³Ùë³ÃÇí ¹³ßïÇ Éñ³óáõÙ
    .Col = gridData
    .Keys(griddate & "[Enter]")
    '¶ñÇ¹Ç ·áõÙ³ñ ¹³ßïÇ Éñ³óáõÙ
    .Col = gridSumma
    .Keys(summ & "[Enter]")
  End With
  'Î³ï³ñ»É Ïá×³ÏÇ ë»ÕÙáõÙ
  Call ClickCmdButton(1, "Î³ï³ñ»É")
  '²ÛÉ í×³ñáõÙÝ»ñÇ ·ñ³ýÇÏÇ ³éÏ³ÛáõÃÛ³Ý ëïáõ·áõÙ
  If wMDIClient.WaitVBObject("frmPttel", 3000).Exists Then
    Do Until wMDIClient.vbObject("frmPttel").vbObject("tdbgView").EOF
      If Left(Trim(wMDIClient.vbObject("frmPttel").vbObject("tdbgView").Text), 22) = "²ÛÉ í×³ñáõÙÝ»ñÇ ·ñ³ýÇÏ" Then
        isExist = True
        Exit Do
      Else 
        Call wMDIClient.vbObject("frmPttel").vbObject("tdbgView").MoveNext
      End If
    Loop
  Else
    Log.Message "The sending documnet frmPttel doesn't exist", "", pmNormal, ErrorColor
  End If
    
  Other_Payment_Schedule = isExist
End Function

'--------------------------------------------------------------------------------------
'Տոկոսների հաշվարկում գործողության կատարում
'--------------------------------------------------------------------------------------
'calcDate - Տոկոսշների հաշվարկման ամսաթիվ դաշտի արժեք
'actionDate - Տոկոսների գործողության ամսաթիվ դաշտի արժեք
'beforeTerm - true եթե կատարվել է ժամկետից շուտ մարում
Function Calculate_Percents(calcDate , actionDate, beforeTerm)
  Dim calcPRBase
    
  BuiltIn.Delay(3000)
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_Opers & "|" & c_Interests & "|" & c_PrcAccruing)
    
  If wMDIClient.WaitVBObject("frmASDocForm", 3000).Exists Then
    calcPRBase = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
    'Հաշվարկման ամսաթիվ դաշտի լրացում
    Call Rekvizit_Fill("Document", 1, "General", "DATECHARGE", calcDate)
    'Գործողության ամսաթիվ դաշտի լրացում
    Call Rekvizit_Fill("Document", 1, "General", "DATE", actionDate)
    'Կատարել կոճակի սեղմում
    Call ClickCmdButton(1, "Î³ï³ñ»É")
    If beforeTerm Then
      Call ClickCmdButton(5, "²Ûá")
    End If
  Else 
    Log.Error "Can't open frmASDocForm window", "", pmNormal, ErrorColor
  End If

  Calculate_Percents = calcPRBase
End Function

'--------------------------------------------------------------------------------------
'Տրամադրում գանձումից գործողության կատարում
'--------------------------------------------------------------------------------------
'data - Տրամադրում գանձումից փաստաթղթի ամսաթիվ դաշտի արժեք
'cash - Կանխիկ/անկանխիկ դաշտի արժեք
'acc - Հաշիվ դաշտի արժեք
'fBaseCP - Տրամադրում գանձումից փաստաթղթի ISN
Function Collect_From_Provision(data, summ, cash, acc, fBaseCP)
  Dim Str, cashNumber
    
  BuiltIn.Delay(3000)  
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_Opers & "|" & c_GiveAndBack & "|" & c_GiveCharge)

  If wMDIClient.WaitVBObject("frmASDocForm", 3000).Exists Then
    'Ամսաթիվ դաշտի լրացում
    Call Rekvizit_Fill("Document",1,"General","DATE","![End]" & "[Del]" & data)
    'ISN-ի վերագրում փոփոխականին
    fBaseCP = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
    'Լրացնում է "Գումար" դաշտը
    Call  Rekvizit_Fill("Document",1,"General","SUMMA",summ)
    'Կանխիկ/Անկանխիկ դաշտի լրացում
    Call Rekvizit_Fill("Document", 1, "General", "CASHORNO", cash)
    
	'Փաստաթղթի N դաշտի արժեքի վերագրում փոփոխականին
    cashNumber = Get_Rekvizit_Value("Document",1,"Mask","CODE")
    If cash = "2" Then
      'Հաշիվ դաշտի լրացում
      Call Rekvizit_Fill("Document", 1, "General", "ACCCORR", acc)
      'Կատարել կոճակի սեղմում
      Call ClickCmdButton(1, "Î³ï³ñ»É")
      Call ClickCmdButton(5, "²Ûá")
    Else
	  'Կատարել կոճակի սեղմում
      Call ClickCmdButton(1, "Î³ï³ñ»É")
	  BuiltIn.Delay(1000)
       'Կատարել կոճակի սեղմում
      Call ClickCmdButton(1, "Î³ï³ñ»É")
          
      'Եթե քաղվածքի պատուհանը հայտնվել է, ապա փակում է
        If wMDIClient.WaitVBObject("FrmSpr",1000).Exists Then
            wMDIClient.VBObject("FrmSpr").Close
        Else
            Log.Error "Statement window doesn't exist!",,,ErrorColor
        End If
    End If
  Else 
    Log.Error "Can't open frmASDocForm window", "", pmNormal, ErrorColor
  End If
    
  Collect_From_Provision = cashNumber
End Function

'Գանձում ներգրավումից գործողության կատարում
'Date - Ամսաթիվ
'Sum - Գումար
'CashOrNo - Կանխիկ/Անկանխիկ
'--------------------------------------------------------------------------------------
Sub ChargeForAttraction(fBASE,Date, Sum, CashOrNo,acc)

  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_Opers & "|" & c_GiveAndBack & "|" & c_GetCharge)
  
  'ISN-ի վերագրում փոփոխականին
  fBASE = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
  
  'Ամսաթիվ դաշտի լրացում
  Call Rekvizit_Fill("Document",1,"General","DATE",Date)
  'Գումար դաշտի լրացում
  Call Rekvizit_Fill("Document",1,"General","SUMMA",Sum)
  'Կանխիկ/Անկանխիկ դաշտի լրացում
  Call Rekvizit_Fill("Document",1,"General","CASHORNO",CashOrNo)
  'Հաշիվ դաշտի լրացում 
  Call Rekvizit_Fill("Document",1,"General","ACCCORR",acc)
  
  Call ClickCmdButton(1, "Î³ï³ñ»É")
  Call ClickCmdButton(5, "²Ûá")
End Sub

'--------------------------------------------------------------------------------------
'Վարկի տրամադրում գործողության կատարում
'--------------------------------------------------------------------------------------
'data - Վարկի տրամադարում փաստաթղթի ամսաթիվ դաշտի արժեք
'cash - Կանխիկ/Անականպիկ դաշտի արժեք
'acc - Հաշիվ դաշտի արժեք
'giveCrBase - Փաստաթղթի ISN
Function Give_Credit (data, Sum, cash, acc, giveCrBase)
  Dim cashNumber, rekvName
    
  BuiltIn.Delay(3000) 
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_Opers & "|" & c_GiveAndBack & "|" & c_CredGrant)
     
  If wMDIClient.WaitVBObject("frmASDocForm", 3000).Exists Then
    'ISN-ի վերագրում փոփոխականին
    giveCrBase = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
    'Ամսաթիվ դաշտի լրացում
    Call Rekvizit_Fill("Document", 1, "General", "DATE", data)
    'Գումար դաշտի լրացում
    Call Rekvizit_Fill("Document", 1, "General", "SUMMA", Sum)
    'Կանխիկ/Անկանխիկ դաշտի լրացում
    Call Rekvizit_Fill("Document", 1, "General", "CASHORNO", cash)
	
    If cash = "2" Then
      'Հաշիվ դաշտի լրացում
      Call Rekvizit_Fill("Document", 1, "General", "ACCCORR", acc)
      'Կատարել կոճակի սեղմում
      Call ClickCmdButton(1, "Î³ï³ñ»É")
	  'Այո կոճակի սեղմում
      Call ClickCmdButton(5, "²Ûá")
    Else   
      Call ClickCmdButton(1, "Î³ï³ñ»É")
      
      'Փաստաթղթի N դաշտի արժեքի վերագրում փոփոխականին
      cashNumber = Get_Rekvizit_Value("Document",1,"General","DOCNUM")
      
      Call ClickCmdButton(1, "Î³ï³ñ»É")
        
      BuiltIn.Delay(1000)
      wMDIClient.vbObject("FrmSpr").Close        
    End If
  Else 
    Log.Error "Can't open frmASDocForm window", "", pmNormal, ErrorColor
  End If
  
  Give_Credit = cashNumber 
End Function

'--------------------------------------------------------------------------------------
'Ներգրավում գործողության կատարում
'Date - Ամսաթիվ
'Sum - Գումար
'CashOrNo - Կանխիկ/Անկանխիկ
'Action - Ներգրավվող պայմանագրի տեսակին համապատասխանող գործողությունը
'CalcAcc - Հաշիվ
'--------------------------------------------------------------------------------------
Function Attraction(Action, Date, Sum, CashOrNo, CalcAcc)
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_Opers & "|" & c_GiveAndBack & "|" & Action)
  
  'ISN-ի վերագրում փոփոխականին
  Attraction = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
  'Ամսաթիվ դաշտի լրացում
  Call Rekvizit_Fill("Document",1,"General","DATE",Date)
  'Գումար դաշտի լրացում
  Call Rekvizit_Fill("Document",1,"General","SUMMA",Sum)
  'Կանխիկ/Անկանխիկ դաշտի լրացում
  Call Rekvizit_Fill("Document",1,"General","CASHORNO",CashOrNo)
  'Հաշիվ դաշտի լրացում
  Call Rekvizit_Fill("Document",1,"General","ACCCORR", CalcAcc)

  Call ClickCmdButton(1, "Î³ï³ñ»É")
  Call ClickCmdButton(5, "²Ûá")
End Function

'--------------------------------------------------------------------------------------
'Խմբային տոկոսների հաշվարկ գործողության կատարում
'--------------------------------------------------------------------------------------
'calcDate - Խմբային տոկոսների հաշվարկ փաստաթղթի ամսաթիվ դաշտի արժեք
'givenDate - Հատկացման ամսաթիվ դաշտի արժեք
'rent - true արժեքի դեպքում հաշվարկվում է վարձավճար,false` կատարվում է տոկոսների հաշվարկ
Sub Percent_Group_Calculate(calcDate, givenDate, rent,changeLimit)
  wMDIClient.vbObject("frmPttel").vbObject("tdbgView").Keys("[Ins]")
  BuiltIn.Delay(3000)
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_GroupCalc)
  If p1.WaitVBObject("frmAsUstPar", 3000).Exists Then
    'Հաշվարկման ամսաթիվ դաշտի լրացում
    Call Rekvizit_Fill("Dialog", 1, "General", "CloseDate", calcDate)
    'Հատկացման ամսաթիվ դաշտի լրացում
    Call Rekvizit_Fill("Dialog", 1, "General", "SetDate", givenDate)
    'Վարձավճարի հաշվարկ դաշտի նշում
    If rent Then
      Call Rekvizit_Fill("Dialog", 1, "CheckBox", "GNZ", 1)
    Else
      If changeLimit then                                                                  
        'Սահմանաչափերի փոփոխում ըստ գրաֆիկների գործողության կատարում
        Call Rekvizit_Fill("Dialog", 1, "CheckBox", "LMS", 1)
      Else
        'Տոկոսների հաշվարկում դաշտի նշում
        Call Rekvizit_Fill("Dialog", 1, "CheckBox", "CHG", 1)
      End If 
    End If
    'Կատարել կոճակի սեղմում
    Call ClickCmdButton(2, "Î³ï³ñ»É")
	 BuiltIn.Delay(1000) 
    Call ClickCmdButton(5, "²Ûá")
  Else 
    Log.Error "Can't open frmAsUstPar window", "", pmNormal, ErrorColor
  End If
End Sub

'--------------------------------------------------------------------------------------
'ä³ÛÙ³Ý³·ñÇ ¹ÇïáõÙ Ñ³ßí»ïíáõÃÛ³Ý Ï³Ýã: üáõÝÏóÇ³Ý í»ñ³¹³ñÓÝáõÙ ¿ true , »Ã» ä³ÛÙ³Ý·ñÇ
'¹ÇïáõÙ Ñ³ßí»ïíáõÃÛáõÝÁ ³éÏ³ , false`  »Ã» ³ÛÝ µ³ó³Ï³ÛáõÙ ¿ :
'--------------------------------------------------------------------------------------
Function View_Contract()
  Dim isExists : isExists = True
    
  BuiltIn.Delay(3000)
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click("Տեղեկանքներ|Ընդհանուր դիտում")
  If p1.WaitVBObject("frmAsUstPar", 3000).Exists Then
    '²Ùë³ÃÇí ¹³ßïÇ Éñ³óáõÙ
    Call Rekvizit_Fill("Dialog", 1, "General", "LASTDATE", "!" & "[End]" & "[Del]")
    Call ClickCmdButton(2, "Î³ï³ñ»É")
  Else 
    Log.Error "Can't open frmAsUstPar window", "", pmNormal, ErrorColor
  End If
  If wMDIClient.WaitVBObject("FrmSpr", 2000).Exists Then
    wMDIClient.vbObject("FrmSpr").Close
  Else
    isExists = False
  End If
    
  View_Contract = isExists
End Function

'--------------------------------------------------------------------------------------
'ì³ñÏ³ÛÇÝ å³ÛÙ³Ý³·ñÇ ³éÏ³ÛáõÃÛ³Ý ëïáõ·áõÙ Ñ³ëï³ïí³Õ ÷³ëï³ÃÕÃ»ñ 1 ÃÕÃ³å³Ý³ÏáõÙ :
'üáõÝÏóÇ³Ý í»ñ¹³ñÓÝáõÙ ¿ true, »Ã» å³ÛÙ³Ý³·ÇñÁ ³éÏ³ ¿ , false` »Ã» ³ÛÝ µ³ó³Ï³ÛáõÙ ¿ :
'--------------------------------------------------------------------------------------
'docNum - ö³ëï³ÃÕÃÇ Ñ³Ù³ñ
Function Verify_Credit(docNum)
  Dim is_exists : is_exists = False
  Dim colN
   
  Call wTreeView.DblClickItem("|ì³ñÏ»ñ (ï»Õ³µ³ßËí³Í)|Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I")
  If p1.WaitVBObject("frmAsUstPar", 3000).Exists Then
    Call ClickCmdButton(2, "Î³ï³ñ»É")
  Else 
    Log.Error "Can't open frmAsUstPar window", "", pmNormal, ErrorColor
  End If
  If  wMDIClient.WaitVBObject("frmPttel", 3000).Exists Then
    colN = wMDIClient.vbObject("frmPttel").GetColumnIndex("CODE")
    If SearchInPttel("frmPttel", colN, docNum) Then
    is_exists = true
    End If
  Else
    Log.Message "The sending documnet frmPttel doesn't exist", "", pmNormal, ErrorColor
  End If
    
  Verify_Credit = is_exists
End Function

'--------------------------------------------------------------------------------------
'Î³ï³ñí³Í ·áñÍáÕáõÃÛáõÝÝ»ñÇ çÝçáõÙ ¶áñÍáÕáõÃÛáõÝÝ»ñÇ ¹ÇïáõÙ ÃÕÃ³å³Ý³ÏÇó : üáõÝÏóÇ³Ý
'í»ñ³¹³ñÓÝáõÙ ¿ true, »Ã» ·áñÍáÕáõÃÛáõÝÝ»ñÇ ¹ÇïáõÙ ÃÕÃ³å³Ý³ÏáõÙ ³éÏ³ ¿ ×Çßï ù³Ý³ÏáõÃÛ³Ùµ
'Ï³ï³ñí³Í ·áñÍáÕáõÃÛáõÝÝ»ñ ¨ false` Ñ³Ï³é³Ï ¹»åùáõÙ
'--------------------------------------------------------------------------------------
'actionCount - Î³ï³ñí³Í ·áñÍáÕáõÃÛáõÝÝ»ñÇ ù³Ý³Ï
Function Delete_Operations_From_OperationsView_Folder(actionCount)
  Dim actions : actions = True
    
  BuiltIn.Delay(3000)
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_OpersView)
  If p1.WaitVBObject("frmAsUstPar", 3000).Exists Then
    'Լրացնում է սկզբնաժամկետ դաշտը
    Call Rekvizit_Fill("Dialog", 1, "General", "START", "!" & "[End]" & "[Del]")
    'Լրացնում է վերջնաժամկետ դաշտը
	Call Rekvizit_Fill("Dialog", 1, "General", "END", "!" & "[End]" & "[Del]")
    '՚Լրացնում է Գործողության տեսակ դաշտը
	Call Rekvizit_Fill("Dialog", 1, "General", "DEALTYPE", "!" & "[End]" & "[Del]")
    Call ClickCmdButton(2, "Î³ï³ñ»É")
  Else
    Log.Error "Can't open frmAsUstPar window", "", pmNormal, ErrorColor
  End If
  If wMDIClient.WaitVBObject("frmPttel_2", 2000).Exists Then
    If wMDIClient.vbObject("frmPttel_2").vbObject("tdbgView").VisibleRows <> actionCount Then
      actions = False
    Else
		If wMDIClient.vbObject("frmPttel_2").vbObject("tdbgView").VisibleRows <> 0 Then
			Call wMDIClient.vbObject("frmPttel_2").vbObject("tdbgView").MoveLast
			Do Until wMDIClient.vbObject("frmPttel_2").vbObject("tdbgView").VisibleRows = 0
			BuiltIn.Delay(3000)
			Call wMainForm.MainMenu.Click(c_AllActions)
			Call wMainForm.PopupMenu.Click(c_Delete)
			BuiltIn.Delay(1000)
			Call ClickCmdButton(3, "²Ûá")
			If p1.WaitVBObject("frmAsMsgBox", 2000).Exists Then 
				Call ClickCmdButton(5, "Î³ï³ñ»É")
			End If
			Loop
		End If
    End If
    Call Close_Pttel("frmPttel_2")
  Else
    Log.Error "Can't open frmPttel_2 window", "", pmNormal, ErrorColor
  End If
    
  Delete_Operations_From_OperationsView_Folder = actions
End Function

'--------------------------------------------------------------------------------------
'Î³ï³ñí³Í ·áñÍáÕáõÃÛáõÝÝ»ñÇ çÝçáõÙ Ð³ßí³ñÏÙ³Ý ³Ùë³Ãí»ñ ÃÕÃ³å³Ý³ÏÇó
'--------------------------------------------------------------------------------------
Sub Delete_Operations_From_Calculation_Days_Folder()
  BuiltIn.Delay(3000)
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_ViewEdit & "|" & c_Other & "|" & c_CalcDates)
  If p1.WaitvbObject("frmAsUstPar", 3000).Exists Then
    'Լրացնում է սկզբնաժամկետ դաշտը
	Call Rekvizit_Fill("Dialog", 1, "General", "START", "!" & "[End]" & "[Del]")
    'Լրացնում է վերջնաժամկետ դաշտը 
	Call Rekvizit_Fill("Dialog", 1, "General", "END", "!" & "[End]" & "[Del]")
    Call ClickCmdButton(2, "Î³ï³ñ»É")
  Else
    Log.Error "Can't open frmAsUstPar window", "", pmNormal, ErrorColor
  End If
  If wMDIClient.WaitvbObject("frmPttel_2", 3000).Exists Then
    If wMDIClient.vbObject("frmPttel_2").vbObject("tdbgView").VisibleRows<>0 Then
      Call wMDIClient.vbObject("frmPttel_2").vbObject("tdbgView").MoveLast
      Do Until wMDIClient.vbObject("frmPttel_2").vbObject("tdbgView").VisibleRows = 0
        BuiltIn.Delay(2000)
        Call wMainForm.MainMenu.Click(c_AllActions)
        Call wMainForm.PopupMenu.Click(c_Delete)
        BuiltIn.Delay(1000)
        Call ClickCmdButton(3, "²Ûá")
      Loop
    End If
    Call Close_Pttel("frmPttel_2")
  Else
    Log.Error "Can't open frmPttel_2 window", "", pmNormal, ErrorColor
  End If
End Sub

'--------------------------------------------------------------------------------------
'ê¨ óáõó³ÏÇ "Ð³ëï³ïíáÕ ï»Õ³µ³ßËí³Í ÙÇçáóÝ»ñ" ÃÕÃ³å³Ý³ÏáõÙ å³ÛÙ³Ý³·ñÇ ³éÏ³ÛáõÃÛ³Ý ëïáõ·áõÙ:
'üáõÝÏóÇ³Ý í»ñ³¹³ñÓÝáõÙ ¿ True, »Ã» ³ÛÝ ³éÏ³ ¿ ¨ false` Ñ³Ï³é³Ï ¹»åùáõÙ :
'--------------------------------------------------------------------------------------
'docNum - ö³ëï³ÃÕÃÇ Ñ³Ù³ñ

Function Check_Doc_In_BlackList_Verifier (docNum)
    Dim is_exists : is_exists = False

    Call wTreeView.DblClickItem("|§ê¨ óáõó³Ï¦ Ñ³ëï³ïáÕÇ ²Þî|Ð³ëï³ïíáÕ ï»Õ³µ³ßËí³Í ÙÇçáóÝ»ñ ¨ »ñ³ßË³íáñáõÃÛáõÝÝ»ñ")
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    
    If wMDIClient.WaitVBObject("frmPttel", 3000).Exists Then
        Do Until wMDIClient.vbObject("frmPttel").vbObject("tdbgView").EOF
            If Trim(wMDIClient.vbObject("frmPttel").vbObject("tdbgView").Columns.Item(2).Text) = Trim(docNum) Then
                is_exists = True
                Exit Do
            Else
                Call wMDIClient.vbObject("frmPttel").vbObject("tdbgView").MoveNext
            End If
        Loop
    Else
        Log.Message("The double input frmPttel does't exist")
    End If
    Check_Doc_In_BlackList_Verifier = is_exists
    
End Function

'--------------------------------------------------------------------------------------
'ì³ñÓ³í×³ñÇ ·ñ³ýÇÏÇ Ýß³Ý³ÏáõÙ ä³ÛÙ³Ý³·ñÇ ÃÕÃ³å³Ý³ÏÇó : üáõÝÏóÇ³Ý  í»ñ³¹³ñÓÝáõÙ ¿ true,
'»Ã» Ù³ñáõÙÝ»ñÇ ·ñ³ýÇÏÁ Ýß³Ý³ÏíáõÙ ¿ , ¨ false ` Ñ³Ï³é³Ï ¹»åùáõÙ :
'--------------------------------------------------------------------------------------
Function Rent_Schedule()
    Dim isExists : isExists = False
    
    BuiltIn.Delay(3000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_ChargeSumSchedule)
    BuiltIn.Delay(1000)
    Call ClickCmdButton(1, "Î³ï³ñ»É")
    
    BuiltIn.Delay(2000)
    Do Until wMDIClient.vbObject("frmPttel").vbObject("tdbgView").EOF
        If Left(Trim(wMDIClient.vbObject("frmPttel").vbObject("tdbgView").Text), 20) = "ì³ñÓ³í×³ñÝ»ñÇ ·ñ³ýÇÏ" Then
            isExists = True
            Exit Do
        Else
            Call wMDIClient.vbObject("frmPttel").vbObject("tdbgView").MoveNext
        End If
    Loop
    
    Rent_Schedule = isExists
End Function

'-------------------------------------------------------------------------------------
' Վարկային պայմանագրերի դաս
'-------------------------------------------------------------------------------------
Class LoanDocument
  Public DocNum, CreditCode, Client, Curr, RepayCurr, CalcAcc, CalcAccPer,_
         Limit, Renewable,ArgType, AutoCap,  Date, AutoDebt, UseOtherAcc, Percent, Baj, PercPenAgr,_
         NonUsedPercent, GiveDate, Term, CloseDate, DateFill, FirstDate, Paragraph, PayDates,_
         CheckPayDates, Direction, Sector, UsageField, Aim, Schedule, Guarantee,_
         Country, District, RegionLR, WeightAMDRisk, PaperCode, LinkedAgr,_
         Note, PledgeCode, PledgeCurr, PladgeSum, PladgeCount, Time, fBASE,_
         DocLevel
  Public TaxRate, SubsidyRate, PermAsAcc, SumsDateFillType, SumsFillType, FillRoundPr
  Public DocType
            
  Private Sub Class_Initialize()
    CreditCode = Null
    Client = Null
    Curr = Null
    RepayCurr = 1
    CalcAccPer = Null
    Renewable = 1
    ArgType = ""
    AutoCap = 0
    AutoDebt = 1
    UseOtherAcc = Null
    Percent = 12
    Baj = 365
    PercPenAgr = 0
    NonUsedPercent = 8
    FillRoundPr = 2
    CloseDate = Null
    DateFill = 1
    Paragraph = 1
    PayDates = Null
    CheckPayDates = 0
    Direction = 2
    Sector = "U2"
    UsageField = "01.001"
    Aim = "00"
    Schedule = 9
    Guarantee = 9
    Country = "AM"
    District = "001"
    RegionLR = "010000008"
    WeightAMDRisk = 0
    LinkedAgr = Null
    Note = Null
    PledgeCode = Null
    PledgeCurr = Null
    PladgeSum = Null
    PladgeCount = Null
    Time = "111111"
    SumsDateFillType = 1
    SumsFillType = "01"
  End Sub
  
 '-------------------------------------------------------------------------------------
  'Ներգրավված վարկի ստեղծում
 '-------------------------------------------------------------------------------------
  Public Sub CreateAttrLoan(FolderName)
    Dim frmModalBrowser, wTabStrip, TabN
    
    Call wTreeView.DblClickItem(FolderName)
    
    Set frmModalBrowser  = Asbank.WaitVBObject("frmModalBrowser", 500)	
		Do Until p1.frmModalBrowser.VBObject("tdbgView").EOF
			If RTrim(p1.frmModalBrowser.VBObject("tdbgView").Columns.Item(col_item).Text) = DocType  Then
  			Call p1.frmModalBrowser.VBObject("tdbgView").Keys("[Enter]")
  			Exit do
			Else
  			Call p1.frmModalBrowser.VBObject("tdbgView").MoveNext
			End If
		Loop 
    
    'Վերցնել "Պայմանագրի համար" դաշտի արշժեքը
    DocNum = Get_Rekvizit_Value("Document",1,"General","CODE")
    'Լրացնել "Մարման արժույթ" դաշտը 
    Call Rekvizit_Fill("Document", 1, "General", "REPAYCURR", RepayCurr)
    'Լրացնել "Հաշվարկային հաշիվ" դաշտը
    Call Rekvizit_Fill("Document", 1, "General", "ACCACC", CalcAcc)   
    'Լրացնել "Տոկոսների վճարման հաշիվ" դաշտը
    Call Rekvizit_Fill("Document", 1, "General", "ACCACCPR", PermAsAcc)
    'Լրացնել "Սահմանաչափ"("Գումար") դաշտը
    Call Rekvizit_Fill("Document", 1, "General", "SUMMA", Limit)
    'Լրացնել "Պարտքերի ավտոմատ մարում" նշիչը
    wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame").VBObject("CheckBox_2").Value = AutoDebt
    'Լրացնել "Կնքման ամսաթիվ" դաշտը       
    Call Rekvizit_Fill("Document", 1, "General", "DATE", Date)
        
    If Left(DocType,8) = "¶ñ³ýÇÏáí" Then
      'Լրացնել "Հատկացման ամսաթիվ" դաշտը
      Call Rekvizit_Fill("Document", 1, "General", "DATEGIVE", GiveDate)
    End If
    
    If DocType = "´³ñ¹ í³ñÏ (·Í³ÛÇÝ)" or Left(DocType,8) = "¶ñ³ýÇÏáí" Then
      'Լրացնել "Մարման ժամկետ" դաշտը  
      Call Rekvizit_Fill("Document", 1, "General", "DATEAGR", Term)
      If DocType = "¶ñ³ýÇÏáí í³ñÏ³ÛÇÝ å³ÛÙ³Ý³·Çñ (·Í³ÛÇÝ)" Then
        'Լրացնել "Վարկային գծի գործելու ժամկետ" 
        Call Rekvizit_Fill("Document", 1, "General", "DATELNGEND", Term)
      End If
    End If
    
    If Left(DocType,8) = "¶ñ³ýÇÏáí" Then
      'Անցնել 2.Գրաֆիկի լրացման ձև
      If CheckPayDates = 1 Then
        'Լրացնել "Ամսաթվերի լրացման ձև" դաշտը 1 արժեքով
        Call Rekvizit_Fill("Document", 2, "General", "DATESFILLTYPE", 1)
        'Լրացնել "Մարման օրեր" դաշտը
        Call Rekvizit_Fill("Document", 2, "General", "FIXEDDAYS", PayDates)
      Else
        'Լրացնել "Ամսաթվերի լրացման ձև" դաշտը 2 արժեքով
        Call Rekvizit_Fill("Document", 2, "General", "DATESFILLTYPE", 2)
        'Լրացնել "Պարպերություն" դաշտը
        Call Rekvizit_Fill("Document", 2, "General", "AGRPERIOD", Paragraph & "[Tab]")
      End If
      'Լրացնել "Շրջանցման ուղղություն" դաշտը
      Call Rekvizit_Fill("Document", 2, "General", "PASSOVDIRECTION", Direction)
      'Լրացնել "Գումարների ամսաթվերի ընտրություն" դաշտը
      Call Rekvizit_Fill("Document", 2, "General", "SUMSDATESFILLTYPE", SumsDateFillType)
      'Լրացնել "Գումարների բաշխման ձև" դաշտը
      Call Rekvizit_Fill("Document", 2, "General", "SUMSFILLTYPE", SumsFillType)
      
      TabN = 4
    Else 
      TabN = 2  
    End If
    
    'Անցնել 2.Տոկոսներ
   'Լրացնել "Վարկի տոկոսադրույք" դաշտը 
    Call Rekvizit_Fill("Document", TabN, "General", "PCAGR", Percent & "[Tab]" & Baj)
    If DocType <> "ØÇ³Ý·³ÙÛ³ í³ñÏ" Then
      'Լրացնել "Չօգտագործված մասի տոկոսադրույք" դաշտը
      Call Rekvizit_Fill("Document", TabN, "General", "PCNOCHOOSE", NonUsedPercent & "[Tab]" & Baj)
    End If  
    'Լրացնել "Հարկի տոկոս" դաշտը
    If DocType <> "´³ñ¹ í³ñÏ (·Í³ÛÇÝ)" and DocType <> "¶ñ³ýÇÏáí í³ñÏ³ÛÇÝ å³ÛÙ³Ý³·Çñ (·Í³ÛÇÝ)" Then
      Call Rekvizit_Fill("Document", TabN, "General", "TAXVALUE", TaxRate)
    End If 
    
    If Left(DocType,8) <> "¶ñ³ýÇÏáí" Then  
      'Անցնել 4.Ժամկետներ
      Set wTabStrip = wMDIClient.VBObject("frmASDocForm").VBObject("TabStrip")    
      wTabStrip.SelectedItem = wTabStrip.Tabs(4)

      If DocType <> "´³ñ¹ í³ñÏ (·Í³ÛÇÝ)" Then
        'Լրացնել "Հատկացման ամսաթիվ" դաշտը
        Call Rekvizit_Fill("Document", 4, "General", "DATEGIVE", GiveDate)
        'Լրացնել"Մարման ժամկետ" դաշտը
        Call Rekvizit_Fill("Document", 4, "General", "DATEAGR", Term)
      End If  
      'Լրացնել "Ամսաթվերի լրացում" նշիչը
      If DateFill = 1 Then
        Select Case DocType
          Case "ì³ñÏ³ÛÇÝ ·ÇÍ"
            wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame_4").VBObject("CheckBox_4").Click
          Case "´³ñ¹ í³ñÏ (·Í³ÛÇÝ)"
            wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame_4").VBObject("CheckBox_3").Click
          Case "ØÇ³Ý·³ÙÛ³ í³ñÏ"
            wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame_4").VBObject("CheckBox_5").Click  
        End Select
        'Լրացնել "Նշ." նշիչը
        Call Rekvizit_Fill("Dialog", 1, "CheckBox", "INCLFIXD", CheckPayDates)
          If CheckPayDates = 1 Then
            'Լրացնել "Մարման օրեր" դաշտը
            Call Rekvizit_Fill("Dialog", 1, "General", "FIXEDDAYS", PayDates)
          Else
            'Լրացնել "Պարպերություն" դաշտը
            Call Rekvizit_Fill("Dialog", 1, "General", "PERIODICITY", Paragraph & "[Tab]")
          End If
          'Լրացնել "Շրջանցման ուղղություն" դաշտը
          Call Rekvizit_Fill("Dialog", 1, "General", "PASSOVDIRECTION", Direction)
          'Սեղմել "Կատարել"
          Call ClickCmdButton(2, "Î³ï³ñ»É")  
      End If
      TabN = 5
     Else
      TabN = 6 
     End If
     
    'Անցնել 5.Լրացուցիչ
   'Լրացնել "Գործարքի ժամ" դաշտը
    Call Rekvizit_Fill("Document", TabN, "General", "TIMEOP", Time)
    
    'Վերցմել պայմանագրի ISN-ը
    fBASE = wMDIClient.VBObject("frmASDocForm").DocFormCommon.Doc.isn
    
    'Սեղմել "Կատարել"
    Call ClickCmdButton(1, "Î³ï³ñ»É")  
  End Sub
  
'-------------------------------------------------------------------------------------
  'Տեղաբշծխված վարկի ստեղծում
'------------------------------------------------------------------------------------- 
  Public Sub CreatePlLoan(FolderName)
    Dim frmModalBrowser, wTabStrip, TabN, Rekv
    
    If Not IsNull(FolderName) Then
      Call wTreeView.DblClickItem(FolderName)      
    End If
    
    Set frmModalBrowser  = Asbank.WaitVBObject("frmModalBrowser", 500)	
    If frmModalBrowser.Exists Then
  		Do Until p1.frmModalBrowser.VBObject("tdbgView").EOF
  			If RTrim(p1.frmModalBrowser.VBObject("tdbgView").Columns.Item(col_item).Text) = DocType  Then
    			Call p1.frmModalBrowser.VBObject("tdbgView").Keys("[Enter]")
    			Exit do
  			Else
    			Call p1.frmModalBrowser.VBObject("tdbgView").MoveNext
  			End If
  		Loop 
    End If
    
    'Պայմանագրի մակարդակի վերագրում DocLevel օբյեկտին
    If Left(DocType, 4) <> "´³ñ¹" Then
      DocLevel = 1
    Else 
      DocLevel = 2  
    End If
    
    'Վերցնել "Պայմանագրի համար" դաշտի արշժեքը
    DocNum = wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame").VBObject("TextC").Text 
    
    'Լրացնել "Մարման արժույթ" դաշտը 
    Call Rekvizit_Fill("Document", 1, "General", "REPAYCURR", RepayCurr)
    Set Rekv = wMDIClient.VBObject("frmASDocForm").WaitVBObject("AS_LABELACCACC", delay_small)
    If Rekv.Exists Then
      'Լրացնել "Հաշվարկային հաշիվ" դաշտը
      Call Rekvizit_Fill("Document", 1, "General", "ACCACC", CalcAcc)   
    End If  
    Set Rekv = wMDIClient.VBObject("frmASDocForm").WaitVBObject("AS_LABELACCACCPR", delay_small)
    If Rekv.Exists Then
      'Լրացնել "Տոկոսների վճարման հաշիվ" դաշտը
      Call Rekvizit_Fill("Document", 1, "General", "ACCACCPR", PermAsAcc)
    End If  
    
    'Լրացնել "Սահմանաչափ"("Գումար") դաշտը
    Call Rekvizit_Fill("Document", 1, "General", "SUMMA", Limit)
    If Right(DocType, 8) = "(·Í³ÛÇÝ)" or DocType = "ì³ñÏ³ÛÇÝ ·ÇÍ" Then
      Call Rekvizit_Fill("Document", 1, "CheckBox", "ISREGENERATIVE", Renewable)
    End If
    
    'Լրացնել "Կնքման ամսաթիվ" դաշտը       
    Call Rekvizit_Fill("Document", 1, "General", "DATE", Date)
    
    If Left(DocType, 28) = "¶ñ³ýÇÏáí í³ñÏ³ÛÇÝ å³ÛÙ³Ý³·Çñ" Then
      'Լրացնել "Հատկացման ամսաթիվ" դաշտը
      Call Rekvizit_Fill("Document", 1, "General", "DATEGIVE", GiveDate)
    End If   
    
    If DocType <> "ØÇ³Ý·³ÙÛ³ í³ñÏ" And DocType <> "ì³ñÏ³ÛÇÝ ·ÇÍ" Then
      'Լրացնել "Մարման ժամկետ" դաշտը  
      Call Rekvizit_Fill("Document", 1, "General", "DATEAGR", Term)
      If DocType = "¶ñ³ýÇÏáí í³ñÏ³ÛÇÝ å³ÛÙ³Ý³·Çñ (·Í³ÛÇÝ)" Then
        Call Rekvizit_Fill("Document", 1, "General", "DATELNGEND", Term)
      End If
    End If
    
    'Լրացնել "Ձևանմուշի N" դաշտը   
    If ArgType = "0005" Then 
      Call Rekvizit_Fill("Document", 1, "General", "AGRTYPE", ArgType) 
    End If
    
    'Անցնել 2(3).Պարտքերի մարման ձև
    If DocType = "ì³ñÏ³ÛÇÝ ·ÇÍ" Or DocType = "ØÇ³Ý·³ÙÛ³ í³ñÏ" Then
      TabN = 2
    Else 
      TabN = 3
    End If
    
    Set wTabStrip = wMDIClient.VBObject("frmASDocForm").VBObject("TabStrip")    
    wTabStrip.SelectedItem = wTabStrip.Tabs(TabN)
    
    'Լրացնել "Պարտքերի ավտոմատ մարում" նշիչը
    Rekv = GetVBObject("AUTODEBT", wMDIClient.vbObject("frmASDocForm"))
    If wMDIClient.vbObject("frmASDocForm").vbObject("TabFrame" & "_" & TabN).vbObject(Rekv).Enabled <> False Then
      Call Rekvizit_Fill("Document", TabN, "CheckBox", "AUTODEBT", AutoDebt) 
    End If  
    'Լրացնել "Այլ հաշիվների մնացորդների օգտագործում" դաշտը
    Call Rekvizit_Fill("Document", TabN, "General", "ACCCONNMODE", UseOtherAcc)   
  
    If Left(DocType, 28) = "¶ñ³ýÇÏáí í³ñÏ³ÛÇÝ å³ÛÙ³Ý³·Çñ" Then
      'Անցնել 4.Գրաֆիկի լրացման ձև
      TabN = 4
      Set wTabStrip = wMDIClient.VBObject("frmASDocForm").VBObject("TabStrip")    
      wTabStrip.SelectedItem = wTabStrip.Tabs(TabN)
      
      If CheckPayDates = 1 Then
        'Լրացնել "Ամսաթվերի լրացման ձև" դաշտը 1 արժեքով
        Call Rekvizit_Fill("Document", TabN, "General", "DATESFILLTYPE", 1)
        'Լրացնել "Մարման օրեր" դաշտը
        Call Rekvizit_Fill("Document", TabN, "General", "FIXEDDAYS", PayDates)
      Else
        'Լրացնել "Ամսաթվերի լրացման ձև" դաշտը 1 արժեքով
        Call Rekvizit_Fill("Document", TabN, "General", "DATESFILLTYPE", 2)
        'Լրացնել "Պարպերություն" դաշտը
        Call Rekvizit_Fill("Document", TabN, "General", "AGRPERIOD", Paragraph & "[Tab]")
      End If
      'Լրացնել "Շրջանցման ուղղություն" դաշտը
      Call Rekvizit_Fill("Document", TabN, "General", "PASSOVDIRECTION", Direction)
      'Լրացնել "Գումարների ամսաթվերի ընտրություն" դաշտը
      Call Rekvizit_Fill("Document", TabN, "General", "SUMSDATESFILLTYPE", SumsDateFillType)
      'Լրացնել "Գումարների բաշխման ձև" դաշտը
      Call Rekvizit_Fill("Document", TabN, "General", "SUMSFILLTYPE", SumsFillType)
      
      TabN = 6
    Else 
      TabN = 4
    End If
    
    If DocType = "ì³ñÏ³ÛÇÝ ·ÇÍ" Or DocType = "ØÇ³Ý·³ÙÛ³ í³ñÏ" Then
      TabN = 3
    End If
    
    
    'Անցնել 3(6).Տոկոսներ
    Set wTabStrip = wMDIClient.VBObject("frmASDocForm").VBObject("TabStrip")    
    wTabStrip.SelectedItem = wTabStrip.Tabs(TabN)
    'Լրացնել "Վարկի տոկոսադրույք" դաշտը 
    Call Rekvizit_Fill("Document", TabN, "General", "PCAGR", Percent & "[Tab]" & Baj)
    'Լրացնել "Չօգտագործված մասի տոկոսադրույք" դաշտը
    Call Rekvizit_Fill("Document", TabN, "General", "PCNOCHOOSE", NonUsedPercent & "[Tab]" & Baj)
    
    If DocType <> "ØÇ³Ý·³ÙÛ³ í³ñÏ" And DocType <> "ì³ñÏ³ÛÇÝ ·ÇÍ" Then
      'Լրացնել "Սուբսիդավորման տոկոս" դաշտը
      Call Rekvizit_Fill("Document", TabN, "General", "PCGRANT", SubsidyRate)
    End If 
    
    If Left(DocType, 28) = "¶ñ³ýÇÏáí í³ñÏ³ÛÇÝ å³ÛÙ³Ý³·Çñ" Then
      Call Rekvizit_Fill("Document", TabN, "General", "FILLROUNDPR", FillRoundPr)
    End If
    
    If Left(DocType, 28) <> "¶ñ³ýÇÏáí í³ñÏ³ÛÇÝ å³ÛÙ³Ý³·Çñ" And DocType <> "´³ñ¹ í³ñÏ" Then  
      TabN = 6
           
    If DocType = "ì³ñÏ³ÛÇÝ ·ÇÍ" Or DocType = "ØÇ³Ý·³ÙÛ³ í³ñÏ" Then
      TabN = 5
    End If
      'Անցնել 5.Ժամկետներ
      Set wTabStrip = wMDIClient.VBObject("frmASDocForm").VBObject("TabStrip")    
      wTabStrip.SelectedItem = wTabStrip.Tabs(TabN)
      If DocType <> "´³ñ¹ í³ñÏ (·Í³ÛÇÝ)" Then
        'Լրացնել "Հատկացման ամսաթիվ" դաշտը
        If DocType = "ì³ñÏ³ÛÇÝ ·ÇÍ" or DocType = "ØÇ³Ý·³ÙÛ³ í³ñÏ" Then
          Call Rekvizit_Fill("Document", TabN, "General", "DATEGIVE", GiveDate)
        End If
        'Լրացնել "Մարման ժամկետ" դաշտը
        Call Rekvizit_Fill("Document", TabN, "General", "DATEAGR", Term)
      End If  
      
      'Լրացնել "Ամսաթվերի լրացում" նշիչը
       Set wTabStrip = wMDIClient.VBObject("frmASDocForm").VBObject("TabStrip")    
      wTabStrip.SelectedItem = wTabStrip.Tabs(TabN)
      If DateFill = 1 Then
        With wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame_" & TabN)
          Select Case DocType
            Case "ì³ñÏ³ÛÇÝ ·ÇÍ"
              .VBObject("CheckBox_11").Click
            Case "´³ñ¹ í³ñÏ (·Í³ÛÇÝ)"
              .VBObject("CheckBox_10").Click
            Case "ØÇ³Ý·³ÙÛ³ í³ñÏ"
              .VBObject("CheckBox_10").Click
          End Select
       End With   
        'Լրացնել "Նշ." նշիչը
        Call Rekvizit_Fill("Dialog", 1, "CheckBox", "INCLFIXD", CheckPayDates)
        If CheckPayDates = 1 Then
          'Լրացնել "Մարման օրեր" դաշտը
          Call Rekvizit_Fill("Dialog", 1, "General", "FIXEDDAYS", PayDates)
        Else
          'Լրացնել "Պարպերություն" դաշտը
          Call Rekvizit_Fill("Dialog", 1, "General", "PERIODICITY", Paragraph & "[Tab]")
        End If
        'Լրացնել "Շրջանցման ուղղություն" դաշտը
        Call Rekvizit_Fill("Dialog", 1, "General", "PASSOVDIRECTION", Direction)
        'Սեղմել "Կատարել"
        Call ClickCmdButton(2, "Î³ï³ñ»É")
      End If
      TabN = 7
     ElseIf DocType = "´³ñ¹ í³ñÏ" Then
      TabN = 5
     Else 
      TabN = 8
     End If

	 If DocType = "ì³ñÏ³ÛÇÝ ·ÇÍ" Or DocType = "ØÇ³Ý·³ÙÛ³ í³ñÏ" Then
       TabN = 6
	 End If
     
    'Անցնել 5(6, 8).Լրացուցիչ
    Set wTabStrip = wMDIClient.VBObject("frmASDocForm").VBObject("TabStrip")    
    wTabStrip.SelectedItem = wTabStrip.Tabs(TabN)
    'Լրացնել "Ճյուղայնություն" դաշտը
    Call Rekvizit_Fill("Document", TabN, "General", "SECTOR", Sector)
    'Լրացնել "Օգտագործման ոլորտ(նոր ՎՌ)" դաշտը
    Call Rekvizit_Fill("Document", TabN, "General", "USAGEFIELD", UsageField)
    'Լրացնել "Նպատակ" դաշտը
    Call Rekvizit_Fill("Document", TabN, "General", "AIM", Aim)
    'Լրացնել "Ծրագիր" դաշտը
    Call Rekvizit_Fill("Document", TabN, "General", "SCHEDULE", Schedule)
    'Լրացնել "Երաշխավորություն" դաշտը
    Call Rekvizit_Fill("Document", TabN, "General", "GUARANTEE", Guarantee)
    'Լրացնել "Երկիր" դաշտը
    Call Rekvizit_Fill("Document", TabN, "General", "COUNTRY", Country)
    'Լրացնել "Մարզ" դաշտը
    Call Rekvizit_Fill("Document", TabN, "General", "LRDISTR", District)
    'Լրացնել "Մարզ(նոր ՎՌ)" դաշտը
    Call Rekvizit_Fill("Document", TabN, "General", "REGION", RegionLR)
    'Լրացնել "Պայմանագրի թղթային համար" դաշտը
    Call Rekvizit_Fill("Document", TabN, "General", "PPRCODE", PaperCode)
    'Լրացնել "Գործարքի ժամ" դաշտը
    Call Rekvizit_Fill("Document", TabN, "General", "TIMEOP", Time)
    
    'Վերցմել պայմանագրի ISN-ը
    fBASE = wMDIClient.VBObject("frmASDocForm").DocFormCommon.Doc.isn
    
    'Սեղմել "Կատարել"
    Call ClickCmdButton(1, "Î³ï³ñ»É") 
  End Sub
  
  Public Sub CreateCreditLine()
   Call Credit_Line_Doc_Fill (Client, Curr, CalcAcc, Limit, Renewable, AutoCap,_
                              Date, Percent, Baj, PercPenAgr, Baj, GiveDate, Term,_
                              Paragraph, Direction, Sector, UsageField, Aim,_
                              Schedule, Guarantee, District, Note, paperCode,_
                              Time, fBASE, DocNum)
    DocLevel = 1                              
  End Sub
  
  'Պայմանագիրը ուղարկում է հաստատման
  Public Function SendToVerify(FolderPath)
    Dim i
    If Not IsNull(FolderPath) Then
      Call wTreeView.DblClickItem(FolderPath)
      Call Rekvizit_Fill("Dialog", 1, "General", "NUM", DocNum) 
      Call ClickCmdButton(2, "Î³ï³ñ»É")
    End If
    
    With wMainForm
      If Left(DocType, 28) = "¶ñ³ýÇÏáí í³ñÏ³ÛÇÝ å³ÛÙ³Ý³·Çñ" Then
        'Կատարել Մարման գրաֆիկի նշանակում
        Builtin.Delay(2000)
        Call .MainMenu.Click(c_AllActions)
        Call .PopupMenu.Click(c_RepaySchedule)
        Builtin.Delay(2000)
        If Right(DocType, 19) = "(³ñïáÝÛ³É Å³ÙÏ»ïáí)" Then
          For i = 0 To wTabFrame.vbObject("DocGrid").RowCount-1
            'Վերցնել Հիմնական ամսաթիվը
            wTabFrame.vbObject("DocGrid").Row = i
            wTabFrame.vbObject("DocGrid").Col = 3
            Call wTabFrame.vbObject("DocGrid").Keys(op_sum & "[Enter]")
          Next
        End If
        If IsNull(FolderPath) Then
          wMDIClient.VBObject("frmPttel").VBObject("tdbgView").MoveNext
        End If
      End If
      
      BuiltIn.Delay(1000)
      Call .MainMenu.Click(c_AllActions)
      Call .PopupMenu.Click(c_SendToVer)
      Builtin.Delay(1000)
    End With  
    Call ClickCmdButton(5, "²Ûá")
    Builtin.Delay(2000)
    Call Close_Pttel("frmPttel")
  End Function
  
  'Հաստատում է պայմանագիրը
  Public Function Verify(FolderPath) 
    Call wTreeView.DblClickItem(FolderPath)
    Call Rekvizit_Fill("Dialog", 1, "General", "NUM", DocNum) 
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    BuiltIn.Delay(4000)
    
    If wMDIClient.VBObject("frmPttel").VBObject("tdbgView").ApproxCount <> 0 Then
    Builtin.Delay(2000)
      Call wMainForm.MainMenu.Click(c_AllActions)
      Call wMainForm.PopupMenu.Click(c_ToConfirm)
      Builtin.Delay(1000)
      Call ClickCmdButton(1, "Ð³ëï³ï»É")
    Else 
      Log.Error(DocNum & " համարի պայմանագիրը չի գտնվել Հաստատվող փաստաթղթեր 1-ում")  
    End If   
      
    Builtin.Delay(2000)
    Call Close_Pttel("frmPttel")
  End Function 
  
  Public Sub OpenInFolder(FolderName)
    Call LetterOfCredit_Filter_Fill(FolderName, DocLevel, DocNum)
  End Sub
  
  Public Sub CloseAgr()
      Builtin.Delay(2000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_AgrClose)
    Builtin.Delay(1000)
    
    Call Rekvizit_Fill("Dialog", 1, "General", "DATECLOSE", CloseDate)
    Call ClickCmdButton(2, "Î³ï³ñ»É")
  End Sub
  
  Public Sub OpenAgr()
	Builtin.Delay(1000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_AgrOpen)
    Builtin.Delay(1000)
    
    Call ClickCmdButton(5, "²Ûá")
  End Sub
  
End Class 

Public Function New_LoanDocument()
  Set New_LoanDocument = New LoanDocument
End Function

'------------------------------------------------------------------------------
'Տոկոսների նշանակում:
'Տոկոսադրույքներ գործողության կատարում
'opDate - Ամսաթիվ
'Prc - Տոկոասադրույք
'NonUsedPrc - Չօգտ.մասի տոկոասադրույք
'------------------------------------------------------------------------------
Function ChangeRete(opDate, Prc, NonUsedPrc)
  Dim Rekv, calcPRBase
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_TermsStates & "|" & c_Percentages & "|" & c_Percentages)
  Builtin.Delay(2000)
		'ISN-ի վերագրում փոփոխականին
  calcPRBase = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
  'Ամսաթիվ դաշտի լրացում
  Call Rekvizit_Fill("Document", 1, "General", "DATE", opDate)
  'Տոկոասադրույք դաշտի լրացում
  Call Rekvizit_Fill("Document", 1, "General", "PCAGR", Prc & "[Tab]")
  
  Set Rekv = wMDIClient.VBObject("frmASDocForm").WaitVBObject("AS_LABELPCNOCHOOSE", delay_small)
  If Rekv.Exists Then
   'Չօգտ.մասի տոկոասադրույք դաշտի լրացում
   Call Rekvizit_Fill("Document", 1, "General", "PCNOCHOOSE", NonUsedPrc & "[Tab]" & "365")
  End If 
  Call ClickCmdButton(1, "Î³ï³ñ»É")
		ChangeRete = calcPRBase
End Function

'------------------------------------------------------------------------------
'Տոկոսների նշանակում:
'Արդյունավետ տոկոսադրույք գործողության կատարում
'opDate - Ամսաթիվ
'EffRete - Արդյունավետ տոկոասադրույք
'ActRete - Փաստացի տոկոասադրույք
'------------------------------------------------------------------------------
Function ChangeEffRete(opDate, EffRete, ActRete)
		Dim calcPRBase
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_TermsStates & "|" & c_Percentages & "|" & c_EffRate)
  Builtin.Delay(2000)
		'ISN-ի վերագրում փոփոխականին
  calcPRBase = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
  'Ամսաթիվ դաշտի լրացում
  Call Rekvizit_Fill("Document", 1, "General", "DATE", opDate)
  'Արդյունավետ տոկոասադրույք դաշտի լրացում
  Call Rekvizit_Fill("Document", 1, "General", "PCNDER", EffRete)
  'Փաստացի տոկոասադրույք դաշտի լրացում
  Call Rekvizit_Fill("Document", 1, "General", "PCNDERALL", ActRete & "[Tab]" & "365")
  
  Call ClickCmdButton(1, "Î³ï³ñ»É")
		ChangeEffRete = calcPRBase
End Function

'------------------------------------------------------------------------------
'Օբյեկտիվ ռիսկի դասիչ գործողության կատարում
'------------------------------------------------------------------------------
Function ObjectiveRisk(Date, Level)
  Dim calcPRBase
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_TermsStates & "|" & c_Risking & "|" & c_ObjRiskCat)
  Builtin.Delay(2000)
		'ISN-ի վերագրում փոփոխականին
  calcPRBase = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
  
   'Ամսաթիվ դաշտի լրացում
    Call Rekvizit_Fill("Document", 1, "General", "DATE", Date)
    'Ռիսկի դասիչ դաշտի լրացում
    Call Rekvizit_Fill("Document", 1, "General", "RISK", Level)
  
    Call ClickCmdButton(1, "Î³ï³ñ»É")
		ObjectiveRisk = calcPRBase
End Function

'------------------------------------------------------------------------------
' Կանխավ վճարված տոկոսների վերադարձ
'------------------------------------------------------------------------------
Function ReturnPrepaidRates(Date, Sum)
  Builtin.Delay(2000)
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_Opers & "|" & c_GiveAndBack & "|" & c_ReturnPrepaidInt)
  Builtin.Delay(2000)
  
  ReturnPrepaidRates = wMDIClient.VBObject("frmASDocForm").DocFormCommon.Doc.isn
  
    'Ամսաթիվ դաշտի լրացում
    Call Rekvizit_Fill("Document", 1, "General", "DATE", Date)
    'Ռիսկի դասիչ դաշտի լրացում
    Call Rekvizit_Fill("Document", 1, "General", "SUMMA", Sum)
    
    Call ClickCmdButton(1, "Î³ï³ñ»É")
    Call ClickCmdButton(5, "²Ûá")

End Function

'_____________________________________________________________________________________
' Գործողության ջնջում"Դիտում և խմբաոգրում" -ից
'_____________________________________________________________________________________
Sub Delete_ViewEdit(sDate, fDate, opType)
  Dim frmAsMsgBox
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_ViewEdit & "|" & opType)
  
  Call Rekvizit_Fill("Dialog", 1, "General", "START", sDate) 
  Call Rekvizit_Fill("Dialog", 1, "General", "END", fDate) 
  If Asbank.VBObject("frmAsUstPar").WndCaption <> "Բանկի արդյունավետ տոկոսադրույք" And Asbank.VBObject("frmAsUstPar").WndCaption <> "Հաշվարկման ամսաթվեր" Then
    Call Rekvizit_Fill("Dialog", 1, "CheckBox", "ONLYCH", 1) 
  End If 
  Asbank.VBObject("frmAsUstPar").VBObject("CmdOK").ClickButton
  
  wMDIClient.VBObject("frmPttel_2").VBObject("tdbgView").MoveLast
  Do While wMDIClient.VBObject("frmPttel_2").VBObject("tdbgView").ApproxCount <> 0
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_Delete)
    Builtin.Delay(2000)
    Set frmAsMsgBox = Asbank.WaitVbObject("frmAsMsgBox", 1000)
    If frmAsMsgBox.Exists Then 
      frmAsMsgBox.VBObject("cmdButton").ClickButton
      Log.Message(opType & "-ը ջնջելիս բացվել է սխալի պատուհան")
      Exit Do
    End If
    Asbank.VBObject("frmDeleteDoc").VBObject("YesButton").ClickButton
    Set frmAsMsgBox = Asbank.WaitVbObject("frmAsMsgBox", 1000)
    If frmAsMsgBox.Exists Then 
      frmAsMsgBox.VBObject("cmdButton").ClickButton
      Log.Message(opType & "-ը ջնջելիս բացվել է սխալի պատուհան")
      Exit Do
    End If
  Loop
  Builtin.Delay(2000)
  wMDIClient.VBObject("frmPttel_2").Close
End Sub

Sub Loan_Attraction(fBASE,Date,Sum,CashOrNo,acc)
  Builtin.Delay(2000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_Opers & "|" & c_GiveAndBack & "|" & c_LoanAttraction)
  Builtin.Delay(2000)
    
    'ISN-ի վերագրում փոփոխականին
    fBASE = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
  
    'Ամսաթիվ դաշտի լրացում
    Call Rekvizit_Fill("Document",1,"General","DATE",Date)
    'Գումար դաշտի լրացում
    Call Rekvizit_Fill("Document",1,"General","SUMMA",Sum)
    'Կանխիկ/Անկանխիկ դաշտի լրացում
    Call Rekvizit_Fill("Document",1,"General","CASHORNO",CashOrNo)
    'Հաշիվ դաշտի լրացում 
    Call Rekvizit_Fill("Document",1,"General","ACCCORR",acc)
  
    Call ClickCmdButton(1, "Î³ï³ñ»É")
    Builtin.Delay(2000)
    Call ClickCmdButton(5, "²Ûá")
  
End Sub

'--------------------------------------------------------------------------------------
' Î³ÝË³í í×³ñí³Í Վարձավճարի í»ñ³¹³ñÓ ÷³ëïÃÕÃÇ Éñ³óáõÙ :
'--------------------------------------------------------------------------------------
'dateStart - ³Ùë³ÃÇí ¹³ßïÇ ³ñÅ»ù
'summ - ¶áõÙ³ñ ¹³ßïÇ ³ñÅ»ù
'rpBase - ö³ëï³ÃÕÃÇ ISN

Sub Return_Payed_Rent(Date, Summa, CashOrNo, Account, rpBase)
    
    
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_Opers & "|" & c_GiveAndBack & "|" & c_ReturnPrepaidRent)
    
    rpBase = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
    '²Ùë³ÃÇí ¹³ßïÇ Éñ³óáõÙ
    Call Rekvizit_Fill("Document",1,"General","DATE",Date)
    '¶áõÙ³ñ ¹³ßïÇ Éñ³óáõÙ
    Call Rekvizit_Fill("Document",1,"General","SUMMA",Summa)
    'Î³ÝËÇÏ/²ÝÏ³ÝËÇÏ ¹³ßïÇ Éñ³óáõÙ
    Call Rekvizit_Fill("Document",1,"General","CASHORNO",CashOrNo)
    BuiltIn.Delay(1000)    
    'Ð³ßÇí ¹³ßïÇ Éñ³óáõÙ
    Call Rekvizit_Fill("Document",1,"General","ACCCORR",Account)
    
    If p1.WaitVBObject("frmAsMsgBox",1000).Exists Then
        Call ClickCmdButton(5, "Î³ï³ñ»É")
    End If
    
    'Կատարել կոճակի սեղմում
    Call ClickCmdButton(1, "Î³ï³ñ»É")
    Call ClickCmdButton(5, "²Ûá")
'    Call ClickCmdButton(5, "Î³ï³ñ»É")
    BuiltIn.Delay(3000)
    
End Sub