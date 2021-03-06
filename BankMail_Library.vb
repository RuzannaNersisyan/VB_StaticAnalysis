Option Explicit
'USEUNIT Library_Common
'USEUNIT Library_Contracts
'USEUNIT Online_PaySys_Library
'USEUNIT Constants
'USEUNIT Payment_Order_ConfirmPhases_Library
'USEUNIT Subsystems_SQL_Library
'USEUNIT Online_PaySys_Library
'USEUNIT Library_CheckDB

'----------------------------------------------------------------------------------------
' àõÕ³ñÏí³Õ ³ñÅ»ÃÕÃ³ÛÇÝ Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñ ÃÕÃ³å³Ý³ÏáõÙ ÷³ëï³ÃÕÃÇ ³éÏ³ÛáõÃÛ³Ý ëïáõ·áõÙ
'üáõÝÏóÇ³Ý í»ñ³¹³ñÓÝáõÙ ¿ true ÷³ëï³ÃÕÃÇ ³éÏ³ÛáõÃÛ³Ý ¹»åùáõÙ, Ñ³Ï³é³Ï ¹»åùáõÙ` false :
'----------------------------------------------------------------------------------------
Function BankMail_Check_Doc_In_Sending_SecrOrd_Folder(fISN)
  Dim is_exists : is_exists = False
  Dim colN
    
  Call wTreeView.DblClickItem("|BankMail ²Þî|àõÕ³ñÏíáÕ Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñ|àõÕ³ñÏíáÕ ³ñÅ»ÃÕÃ³ÛÇÝ Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñ")
  If p1.WaitVBObject("frmAsUstPar", 3000).Exists Then
    Call ClickCmdButton(2, "Î³ï³ñ»É")
  Else 
    Log.Error "Can't find frmAsUstPar window", "", pmNormal, ErrorColor
  End If    
  If wMDIClient.WaitVBObject("frmPttel", 3000).Exists Then
    colN = wMDIClient.vbObject("frmPttel").GetColumnIndex("DOCNUM")
    If SearchInPttel("frmPttel", colN, fISN) Then
      is_exists = true
    End If
  Else
    Log.Message "The sending documnet frmPttel doesn't exist", "", pmNormal, ErrorColor
  End If
    
  BankMail_Check_Doc_In_Sending_SecrOrd_Folder = is_exists
End Function

'--------------------------------------------------------------------------
'Արժեթղթերի ազատ առաքման / գրավից հանման հանձնարարական փաստաթղթերի լրացում
'--------------------------------------------------------------------------
Sub BankMail_FreeDeliver_Doc_Fill(docNumber, stockID, volume, senderAcc,senderName ,_
                                recstockAcc,recstockName ,fISN )
    
  Dim rekvName, docN
    
  If wMDIClient.WaitVBObject("frmASDocForm", 3000).Exists Then
    'Ստեղծվող ISN - ի փաստատթղթի  վերագրում փոփոխականին
    fISN = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
    'Փաստաթղթի N դաշտի արժեքի վերագրում փոփոխականին
    rekvName = GetVBObject("BMDOCNUM", wMDIClient.vbObject("frmASDocForm"))
    docN = wMDIClient.vbObject("frmASDocForm").vbObject("TabFrame").vbObject(rekvName).Text
    docNumber = Left(docN,6)
    'Արժեթղթերի իդենտիֆիկատոր դաշտի լրացում
    Call Rekvizit_Fill("Document", 1, "General", "STOCKID", stockID)
    'Արժեթղթերի ծավալ դաշտի լրացում
    Call Rekvizit_Fill("Document", 1, "General", "VOLUME", volume)
    'Արժեթղթեր առաքողի հաշիվ դաշտի լրացում
    Call Rekvizit_Fill("Document", 1, "General", "SSENDER", "^A[Del]" & senderAcc)
    'Արժեթղթեր առաքող դաշտի լրացում
    Call Rekvizit_Fill("Document", 1, "General", "SSNAME", "^A[Del]" & senderName)
    'Արժեթղթեր ստացողի հաշիվ դաշտի լրացում
    Call Rekvizit_Fill("Document", 1, "General", "SRECEIVER", "^A[Del]" & recstockAcc)
    'Արժեթղթեր ստացող դաշտի լրացում
    Call Rekvizit_Fill("Document", 1, "General", "SRNAME", "^A[Del]" & recstockName)
    Call ClickCmdButton(1, "Î³ï³ñ»É")
  Else 
    Log.Error "Can't open frmASDocForm window", "", pmNormal, ErrorColor
  End If
End Sub

'----------------------------------------------------------------------------------
' Վճարման հանձնարագրերի լրացում
' ordType - Վճարման հանձնարագրերի տեսակ
' fISN - փաստաթղթի ISN
' wAcsBranch - Գրասենյակ
' wAcsDepart - Բաժին դաշտ
' payDate - Ամսաթիվ
' docNum - Փաստաթղի համար
' cliCode - Վճարող հաճախորդի կոդ
' accDB - Հաշիվ դեբետ
' payer - Վճարող
' ePayer - Վճարող անգլերեն
' taxCods - ՀՎՀՀ (Վճարող)
' jurState - Իրավաբանական կարգավիճակ(վճարող)
' dbDropDown - բուլյան փոփոխական
' coaNum - Հաշվարկային պլանի համար
' balAcc - Հ/Պ հաշվեկշռային հաշիվ
' accMask - Հաշվի շաբլոն
' accCur - Արժույթ
' accType - Հաշվի տիպ
' cliName - Հաճախորդի անվանում
' cCode - Հաճախորդ
' accNote - Նշում
' accNote2 - Նշում 2
' accNote3 - Նշում 3
' acsBranch - Գրասենյակ
' acsDepart - Բաժին
' acsType - Հասանելության տիպ
' pCardNum - Քարտի համար
' socCard - Սոցիալական քարտ
' accCR - Հաշիվ կրեդիտ
' receiver - Ստացող
' eReceiver - Ստացող անգլերեն
' summa - Գումար
' wCur - Արժույթ
' wAim - Նպատակ
' jurStatR - Իրավաբանական կարգավիճակ (Ստացողի)
' bankCr - Ստացող բանկ
' authorPerson - Լիազորված անձ
' addInfo - Լրացուցիչ ինֆորմացիա
' wAddress - Հասցե
' authPerson - Լիազորված անձ
' rInfo - Լրացուցիչ ինֆորմացիա
Sub PaymOrdToBeSentFill(ordType, fISN, wAcsBranch, wAcsDepart, payDate, docNum, cliCode, accDB, payer, ePayer, taxCods,_
                                                jurState, dbDropDown, coaNum, balAcc, accMask, accCur, accType, cliName, cCode, accNote, accNote2,_
                                                accNote3, acsBranch, acsDepart, acsType, pCardNum, socCard, accCR, receiver, eReceiver, summa, wCur,_
                                                wAim, jurStatR, bankCr, authorPerson, addInfo, wAddress, authPerson, rInfo)
  Dim  param
      
  BuiltIn.Delay(3000)
  Call wMainForm.MainMenu.Click(c_AllActions)      
  Select Case ordType
      Case "PAY"
          Call wMainForm.PopupMenu.Click(c_PayOrds & "|" & c_PayOrdToBeSent)
      Case "BDG"
          Call wMainForm.PopupMenu.Click(c_PayOrds & "|" & c_PayOrdBdg)
      Case "WPD"
          Call wMainForm.PopupMenu.Click("Վճարման հանձնարարագրեր|Վճարման հանձնարարագիր (ուղ.) լրացուցիչ վճ. տվյալներով")
      Case "WOA"
          Call wMainForm.PopupMenu.Click("Վճարման հանձնարարագրեր|Անհաշիվ փոխանցում")
      Case "INB"
          Call wMainForm.PopupMenu.Click("Վճարման հանձնարարագրեր|Ներբանկային վճարման հանձնարարագիր (անձն. տվ.)")
  End Select
  BuiltIn.Delay(1000)
      
  fISN = wMDIClient.VBObject("frmASDocForm").DocFormCommon.Doc.ISN
      
  param = GetVBObject("DOCNUM", wMDIClient.VBObject("frmASDocForm"))
  docNum = wMDIClient.VBObject("frmASDocForm").vbObject("TabFrame").VBObject(param).Text
     
 ' Գրասենյակ դաշտի լրացում
 Call Rekvizit_Fill("Document", 1, "General", "ACSBRANCH", wAcsBranch)
 ' Բաժին դաշտի լրացում
 Call Rekvizit_Fill("Document", 1, "General", "ACSDEPART", wAcsDepart & "[Tab]")
 BuiltIn.Delay(1000)
 ' Ամսաթիվ դաշտի լրացում
 Call Rekvizit_Fill("Document", 1, "General", "DATE", "![End][Del]" & payDate)
 ' Հաշիվ դեբետ դաշտի լրացում
 Call Rekvizit_Fill("Document", 1, "General", "ACCDB", "[BS][BS][BS][BS][BS][BS]" & accDB)
 ' Վճարող հաճախորդի կոդ դաշտի լրացում
 Call Rekvizit_Fill("Document", 1, "General", "CLICODE", cliCode)
     
 If dbDropDown Then
    'Հաշիվ դեբետ դաշտի կողքի կոճակի սեղմում
    Call ClickDropDown(1, "ACCDB")
    ' Հաշվարկային պլանի համար դաշտի լրացում
    Call Rekvizit_Fill("Dialog", 1, "General", "COANUM", coaNum)
    ' Հ/Պ հաշվեկշռային հաշիվ դաշտի լրացում
    Call Rekvizit_Fill("Dialog", 1, "General", "BALACC", balAcc)
    ' Հաշվի շաբլոն դաշտի լրացում
    Call Rekvizit_Fill("Dialog", 1, "General", "ACCMASK", accMask)
    ' Արժույթ դաշտի լրացում
    Call Rekvizit_Fill("Dialog", 1, "General", "ACCCUR", accCur)
    ' Հաշվի տիպ դաշտի լրացում
    Call Rekvizit_Fill("Dialog", 1, "General", "ACCTYPE", accType)
    ' Հաճախորդի անվանում դաշտի լրացում
    Call Rekvizit_Fill("Dialog", 1, "General", "CLNAME", cliName)
    ' Հաճախորդ դաշտի լրացում
    Call Rekvizit_Fill("Dialog", 1, "General", "CLICOD", cCode)
    ' Նշում դաշտի լրացում
    Call Rekvizit_Fill("Dialog", 1, "General", "ACCNOTE", accNote)
    ' Նշում 2 դաշտի լրացում
    Call Rekvizit_Fill("Dialog", 1, "General", "ACCNOTE2", accNote2)
    ' Նշում 3 դաշտի լրացում
    Call Rekvizit_Fill("Dialog", 1, "General", "ACCNOTE3", accNote3)
    ' Գրասենյակ դաշտի լրացում
    Call Rekvizit_Fill("Dialog", 1, "General", "ACSBRANCH", acsBranch)
    ' Բաժին դաշտի լրացում
    Call Rekvizit_Fill("Dialog", 1, "General", "ACSDEPART", acsDepart)
    ' Հասանելության տիպ դաշտի լրացում
    Call Rekvizit_Fill("Dialog", 1, "General", "ACSTYPE", acsType)
    ' Քարտի համար դաշտի լրացում
    Call Rekvizit_Fill("Dialog", 1, "General", "PCARD", pCardNum)

     ' Կատարել կոճակի սեղմում 
     Call ClickCmdButton(2, "Î³ï³ñ»É") 
     Sys.Process("Asbank").vbObject("frmModalBrowser").vbObject("tdbgView").Keys("[Enter]") 
     wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame").VBObject("TextC").Click
  End If
       
  ' Վճարող դաշտի լրացում
  Call Rekvizit_Fill("Document", 1, "General", "PAYER", "![End][Del]" & payer)
  ' Վճարող(անգլ) դաշտի լրացում
  Call Rekvizit_Fill("Document", 1, "General", "EPAYER", ePayer)    
  If  ordType <> "INB" Then
       ' ՀՎՀՀ (Վճարող) դաշտի լրացում
       Call Rekvizit_Fill("Document", 1, "General", "TAXCODSD", "![End][Del]" & taxCods)
       ' Իրավաբանական կարգավիճակ(վճարող) դաշտի լրացում
       Call Rekvizit_Fill("Document", 1, "General", "JURSTAT", jurState)
  End If
  If ordType = "BDG" Then
      ' Սոցիալական քարտ դաշտի լրացում
      Call Rekvizit_Fill("Document", 1, "General", "REGNUM", socCard)
  End If
  If ordType <> "WOA" and ordType <> "INB" Then    
     ' Հաշիվ կրեդիտ դաշտի լրացում
     Call Rekvizit_Fill("Document", 1, "General", "ACCCR", "[BS][BS][BS][BS][BS][BS]" & accCR)
  End If
  ' Ստացող դաշտի լրացում
  Call Rekvizit_Fill("Document", 1, "General", "RECEIVER", receiver)
  ' Ստացող(անգլ.) դաշտի լրացում
  Call Rekvizit_Fill("Document", 1, "General", "ERECEIVER", eReceiver)
  ' Գումար դաշտի լրացում
  Call Rekvizit_Fill("Document", 1, "General", "SUMMA", summa)
  ' Արժույթ դաշտի լրացում
  Call Rekvizit_Fill("Document", 1, "General", "CUR", wCur)
  ' Նպատակ դաշտի լրացում
  Call Rekvizit_Fill("Document", 1, "General", "AIM", wAim)
  If ordType = "WPD" Or ordType = "WOA" Then
     ' Իրավաբանական կարգավիճակ (Ստացողի) դաշտի լրացում
     Call Rekvizit_Fill("Document", 1, "General", "JURSTATR", jurStatR)
     If ordType = "WOA" Then
        ' Ստացող բանկ դաշտի լրացում
        Call Rekvizit_Fill("Document", 1, "General", "BANKCR", bankCr)
     End If
     ' Անցում Վճարողի/ատացողի լրացուցիչ տվյալներ  
     ' Լիազորված անձ դաշտի լրացում
     Call Rekvizit_Fill("Document", 2, "General", "SMNG", authorPerson)
     ' Լրացուցիչ ինֆորմացիա դաշտի լրացում
     Call Rekvizit_Fill("Document", 2, "General", "SINFO", addInfo)
     ' Հասցե դաշտի լրացում
     Call Rekvizit_Fill("Document", 2, "General", "RADDRESS", wAddress)
     ' Լիազորված անձ դաշտի լրացում   
     Call Rekvizit_Fill("Document", 2, "General", "RMNG", authPerson)
     ' Լրացուցիչ ինֆորմացիա դաշտի լրացում
     Call Rekvizit_Fill("Document", 2, "General", "RINFO", rInfo)
  End If
      
  ' կատարել կոճակի սեղմում
  Call ClickCmdButton(1, "Î³ï³ñ»É")
      
  BuiltIn.Delay(1000)
  wMDIClient.VBObject("FrmSpr").Close
End Sub

' Պայամանագրի վավերացում
' colN - Թղթապանակում սյան համարը
' docNum - Փասատթղթի համարը
' action - Գործողության տիպ
' doNum - բացված պատուհանի տեսակ
' doActio - Սեղմվող կոճակի անվանում
Function ConfirmContractDoc(colN, docNum, action, doNum, doAction)
  Dim status : status = False
    
  BuiltIn.Delay(3000)
  Do Until wMDIClient.VBObject("frmPttel").VBObject("tdbgView").EOF
    If  Trim(wMDIClient.VBObject("frmPttel").VBObject("tdbgView").Columns.Item(colN).Value) = Trim(docNum) Then
      BuiltIn.Delay(3000)
      'Կատարել բոլոր գործողությունները
      Call wMainForm.MainMenu.Click(c_AllActions)
      'Գործողության տիպը պայամանագրի նկատմամբ
      Call wMainForm.PopupMenu.Click(action)
      BuiltIn.Delay(1000)
      'կոճակի սեղմում
      Call ClickCmdButton(doNum, doAction)
      If p1.WaitVBObject("frmAsMsgBox", 10000).Exists Then
        Call ClickCmdButton(5, "OK")
      End if
      BuiltIn.Delay(2000)
      status = True
      Exit Do                   
    Else
      wMDIClient.VBObject("frmPttel").VBObject("tdbgView").MoveNext
    End If
  Loop 
      
  BuiltIn.Delay(2000)
  ConfirmContractDoc = status   
End Function 

' ՊայամանագÇñÁ áõÕ³ñÏ»É BankMail
' colN - Թղթապանակում սյան համարը
' docNum - Փասատթղթի համարը
' action - Գործողության տիպ
' doNum - բացված պատուհանի տեսակ
' doActio - Սեղմվող կոճակի անվանում
Function Contract_To_Bank_Mail(colN, docNum)
  Dim status : status = False
    
  BuiltIn.Delay(3000)
  Do Until wMDIClient.VBObject("frmPttel").VBObject("tdbgView").EOF
    If  Trim(wMDIClient.VBObject("frmPttel").VBObject("tdbgView").Columns.Item(colN).Value) = Trim(docNum) Then
      BuiltIn.Delay(3000)
      ' Կատարել բոլոր գործողությունները
      Call wMainForm.MainMenu.Click(c_AllActions)
      ' Գործողության տիպը պայամանագրի նկատմամբ
      Call wMainForm.PopupMenu.Click(c_SendBM)
      BuiltIn.Delay(1000)
      '  կոճակի սեղմում
      Call ClickCmdButton(5, "²Ûá")
      if p1.WaitVBObject("frmAsMsgBox", 2000).Exists Then
        Call MessageExists(2, "Ð³Õáñ¹³·ñáõÃÛáõÝÝ»ñÇ áõÕ³ñÏáõÙÝ ³í³ñïí»ó")
        Call ClickCmdButton(5, "OK")
      End if
      BuiltIn.Delay(2000)
      status = True
      Exit Do                   
    Else
      wMDIClient.VBObject("frmPttel").VBObject("tdbgView").MoveNext
    End If
  Loop 
      
  BuiltIn.Delay(2000)
  Contract_To_Bank_Mail = status   
End Function 

' Պայմանագրի առկա լինելը ստուգող ֆունկցիա
' colN - Թղթապանակում սյան համարը
' docTypeName - Փաստաթղթի տեսակ 
Function CheckContractDoc(colN, docTypeName)
  Dim  status : status = False

  BuiltIn.Delay(3000)
  Do Until wMDIClient.VBObject("frmPttel").VBObject("tdbgView").EOF
    If Trim(wMDIClient.VBObject("frmPttel").VBObject("tdbgView").Columns.Item(colN).Value) = Trim(docTypeName) Then
      Log.Message("Փաստաթուղթն առկա է ")
      BuiltIn.Delay(2000)
      status = True
      Exit Do             
    Else
       wMDIClient.VBObject("frmPttel").VBObject("tdbgView").MoveNext
    End If
  Loop 
      
  CheckContractDoc =  status
End Function
 
' Խմբային հիշարար օրդերի  հաստատում
' grRemOrdISN - Խմբային հիշարար օրդերի ISN
' grRemOrdNum - Խմբային հիշարար օրդերի Փաստաթղթի N
' wDate - Խմբային հիշարար օրդերի Ամսաթիվ
Sub GroupReminderOrdersVer(grRemOrdISN, grRemOrdNum, wDate)
  Dim param
      
 ' Խմբային հիշարար օրդերի ISN ի ստացում
  grRemOrdISN = wMDIClient.VBObject("frmASDocForm").DocFormCommon.Doc.ISN
  ' Խմբային հիշարար օրդերի Փաստաթղթի N արժեքի ստացում
  param = GetVBObject("DOCNUM", wMDIClient.VBObject("frmASDocForm"))
  grRemOrdNum = wMDIClient.VBObject("frmASDocForm").vbObject("TabFrame").VBObject(param).Text
  ' Խմբային հիշարար օրդերի Ամսաթիվ դաշտի լրացում
  Call Rekvizit_Fill("Document", 1, "General", "DATE", "![End][Del]" & wDate)
  ' Կատարել կոճակի սեղմում
  Call ClickCmdButton(1, "Î³ï³ñ»É")
  ' Փակել տպելու ձևը
  BuiltIn.Delay(1000)
  wMDIClient.VBObject("FrmSpr").Close
End Sub
 
' Պայմանագիրն առկայության ստուգում Հաճախորդներ թղթապանակում
' docTypeName - Փաստաթղթի տեսակ 
' commentName - Մեկնաբանություն
Function CheckPayOrderAvailableOrNot(docTypeName, commentName)
  Dim status : status = False
      
  BuiltIn.Delay(3000)
  ' Կատարել բոլոր գործողությունները
  Call wMainForm.MainMenu.Click(c_AllActions)
  ' Մուտք հաճախորդի թղթապանակ
  Call wMainForm.PopupMenu.Click(c_ClFolder)
      
  If Not wMDIClient.WaitVBObject("frmPttel_2", 10000).Exists Then
        Log.Error("Հաճախորդի թղթապանակը չի բացվել")
        Exit Function
  End If

  Do Until wMDIClient.VBObject("frmPttel_2").VBObject("tdbgView").EOF
   If Trim(wMDIClient.VBObject("frmPttel_2").VBObject("tdbgView").Columns.Item(0).Value) = Trim(docTypeName) and  Trim( wMDIClient.VBObject("frmPttel_2").VBObject("tdbgView").Columns.Item(1).Value) = Trim(commentName) Then
      Log.Message("Փաստաթուղթն առկա է")
      status = True
      Exit Do     
   Else
      wMDIClient.VBObject("frmPttel_2").VBObject("tdbgView").MoveNext
   End If
  Loop 
      
  CheckPayOrderAvailableOrNot = status       
End Function

' Դիտել Վճարման հանձնարարագրի պայմանագիրն 
'wDate - Ժամանակահատված
'fISN - ծնող փաստաթղթի ISN
'childISN - զավակ փաստաթղթի ISN
'status - բուլեան փոփոխական
'wDateTime - ժամանակ դաշտ
Sub WiewPayOrderFromTransferSent(wDate, fISN, childISN, status, wDateTime)
  BuiltIn.Delay(2000)
  Call wTreeView.DblClickItem("|BankMail ²Þî|àõÕ³ñÏíáÕ Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñ|àõÕ³ñÏíáÕ ÷áË³ÝóáõÙÝ»ñ")
  
  If Not p1.WaitVBObject("frmAsUstPar", 2000).Exists Then
    Log.Error("Ուղարկվող դիալոգը չի բացվել")
    Exit Sub
  End If    
      
  ' Ժամանակահատվածի սկիզբ դաշտի լրացում
  Call Rekvizit_Fill("Dialog", 1, "General", "PERN", "![End][Del]" & wDate)
  ' Ժամանակահատվածի ավարտ դաշտի լրացում
  Call Rekvizit_Fill("Dialog", 1, "General", "PERK", "![End][Del]" & wDate)
  Call ClickCmdButton(2, "Î³ï³ñ»É")

  If Not wMDIClient.WaitVBObject("frmPttel", 3000).Exists Then
    Log.Error("Ուղարկվող թղթապանակը չի բացվել")
    Exit Sub
  End If      
      
  Do Until wMDIClient.VBObject("frmPttel").VBObject("tdbgView").EOF
    If  Trim(wMDIClient.VBObject("frmPttel").VBObject("tdbgView").Columns.Item(2).Value) = Trim(fISN) Then
      BuiltIn.Delay(2000)
      ' Կատարել բոլոր գործողությունները
      Call wMainForm.MainMenu.Click(c_AllActions)
      ' Դիտել Վճարման հանձնարարագրի պայմանագիրն
      Call wMainForm.PopupMenu.Click(c_View)
                        
      childISN = wMDIClient.VBObject("frmASDocForm").DocFormCommon.Doc.ISN
                        
      If status Then
        ' Ժամանակ դաշտի արժեքի ստացում
        wDateTime = wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame").VBObject("TextC_4").Text
      End If
                        
      ' Հաստատել կոճակի սեղմում
      Call ClickCmdButton(1, "OK")
      Exit Do                   
    Else
        wMDIClient.VBObject("frmPttel").VBObject("tdbgView").MoveNext
    End If
  Loop  
End Sub

' wDate - Ժամանակահատված, ամսաթիվ 
' vioid - ՃՈ որոշման համար դաշտի արժեք
' fISN - Վճարման հանձնարարագիր փաստաթուղթի ISN
' docNum - Վճարման հանձնարարագիր փաստաթուղթի համար
' payer - Վճարող դաշտի արժեք
' accCR - Հաշիվ կրեդիտ դաշտի արժեք
' wAim - Նպատակ դաշտի արժեք
' accDB - Հաշիվ դեբետ դաշտի արժեք
' receiver - Ստացող դաշտի արժեք
' sMes1 - Հաղորդագրություն 1 դաշտի արժեք
' sMes2 - Հաղորդագրություն 2 դաշտի արժեք
' aim - դրսից ստացվող նպատակ դաշտի արժեք, որ համեմատվում է ծրագրում ստացված արժեքի հետ
' ՃՈ տուգանքի վճարում
Sub PaymentOfTrafficPenalty(wDate, vioid, fISN, docNum, payer, accCR, wAim, accDB, receiver, sMes1, sMes2, aim)

      ' Մուտք Աշխատանքային փաստաթղթեր թղթապանակ
      Call wTreeView.DblClickItem("|Ð³×³Ëáñ¹Ç ëå³ë³ñÏáõÙ ¨ ¹ñ³Ù³ñÏÕ |²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ")

      If Not p1.WaitVBObject("frmAsUstPar", 3000).Exists Then
             Log.Error("Աշխատանքային փաստաթղթեր դիալոգը չի բացվել")
             Exit Sub
      End If
      
      ' Ժամանակահատվածի սկիզբ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "PERN",  "![End][Del]" & wDate)
      ' Ժամանակահատվածի ավարտ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "PERK", "![End][Del]" & wDate)
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
      
      If Not wMDIClient.WaitVBObject("frmPttel", 3000).Exists Then
             Log.Error("Աշխատանքային փաստաթղթեր թղթապանակը չի բացվել ")
             Exit Sub
      End If
      
      ' Գործողություններ /  Բոլոր գործողություններ 
      Call wMainForm.MainMenu.Click(c_AllActions)
      ' բացել ՃՈ տուգանքի վճարում դիալոգը
      Call wMainForm.PopupMenu.Click(c_PayOrds & "|" & c_PayForTraficPenalty)
      
      BuiltIn.Delay(1000)
      ' ՃՈ որոշման համար դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "VIOID", vioid)
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
      
      If Not wMDIClient.WaitVBObject("frmASDocForm", 3000).Exists Then
            Log.Error("Վճարման հանձնարարագիր փաստաթուղթը չի բացվել")
            Exit Sub
      End If
      
      ' Վճարման հանձնարարագիր փաստաթուղթի ISN -ի ստացում
      fISN = wMDIClient.VBObject("frmASDocForm").DocFormCommon.Doc.ISN
      
      ' Ամսաթիվ դաշտի լրացում
      Call Rekvizit_Fill("Document", 1, "General", "DATE", "![End][Del]" & wDate)
      ' Հաշիվ դեբետ դաշտի լրացում
      Call Rekvizit_Fill("Document", 1, "General", "ACCDB", "[BS][BS][BS][BS][BS][BS]" & accDB)
      ' Ստացող դաշտի լրացում
      Call Rekvizit_Fill("Document", 1, "General", "RECEIVER", receiver)
      
      ' Վճարման հանձնարարագիր փաստաթուղթի համարի ստացում
      docNum = Get_Rekvizit_Value("Document", 1, "General", "DOCNUM")
      ' Վճարող դաշտի արժեքի ստացում
      payer = Get_Rekvizit_Value("Document", 1, "Comment", "PAYER")
      ' Հաշիվ կրեդիտ դաշտի արժեքի ստացում
      accCR = Get_Rekvizit_Value("Document", 1, "Bank", "ACCCR")
      ' Նպատակ դաշտի արժեքի ստացում
      wAim = Get_Rekvizit_Value("Document", 1, "Comment", "AIM")
      
      Call GoTo_ChoosedTab(5)
      ' Հաղորդագրություն 1 դաշտի արժեքի ստացում
      sMes1 = Get_Rekvizit_Value("Document", 5, "Comment", "MESSAGE1")
      ' Հաղորդագրություն 2 դաշտի արժեքի ստացում
      sMes2 = Get_Rekvizit_Value("Document", 5, "Comment", "MESSAGE2")
      
      If Trim(sMes1) <> "CPO/" & fISN Then
          Log.Error("Հաղորդագրություն 1 դաշտի սխալ արժեք")
      Else
          Log.Message("Հաղորդագրություն 1 դաշտի ճիշտ արժեք")
      End If
      
      If Trim(sMes2) <> "PDD/OT/" & vioid & "/PAT/190315/0" Then
          Log.Error("Հաղորդագրություն 2 դաշտի սխալ արժեք")
      Else
          Log.Message("Հաղորդագրություն 2 դաշտի ճիշտ արժեք")
      End If  
      
      If Trim(wAim) <> Trim(aim) Then
          Log.Error("Նպատակ դաշտի սխալ արժեք")   
      Else
          Log.Message("Նպատակ դաշտի ճիշտ արժեք")
      End If 
      Call ClickCmdButton(1, "Î³ï³ñ»É")
      
      BuiltIn.Delay(1000)
      wMDIClient.VBObject("FrmSpr").Close
End Sub

' Թղթապանակ մուտք գործելուց դիալոգում տվյալներ լրացում և թղթապանակում պայմանագրի առկայության ստուգում
' workEnvName = թղթապանակի անվանումը, որտեղ պետք է մուտք կատարվի
' workEnv = թղթապանակի անվանումը
' stRekName = Ժամանակահատվածի սկիզբ ռեկվիզիտի անվանումը
' wDateS =  Ժամանակահատվածի սկիզբ դաշտ
' endRekName = Ժամանակահատվածի ավարտ ռեկվիզիտի անվանումը
' wDateE = Ժամանակահատվածի ավարտ դաշտ
' status = բուլեան փոփոխական
' isnRekName = ISN -ի ռեկվիզիտի անվանումը
' fISN = ISN -ի արժեք
Function AccessFolder(workEnvName, workEnv, stRekName, wDateS, endRekName, wDateE, wStatus, isnRekName, fISN)
      Dim state : state = False
      
      BuiltIn.Delay(2000)
      Call wTreeView.DblClickItem(workEnvName)

      If Not p1.WaitVBObject("frmAsUstPar", 2000).Exists Then
             Log.Error( workEnv & "դիալոգը չի բացվել")
             Exit Function
      End If
      
      ' Ժամանակահատվածի սկիզբ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", stRekName, "![End][Del]" & wDateS)
      ' Ժամանակահատվածի ավարտ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", endRekName, "![End][Del]" & wDateE)
      
	    If wStatus Then
            Call Rekvizit_Fill("Dialog", 1, "General", isnRekName, fISN)
      End If
      
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
      
      If Not wMDIClient.WaitVBObject("frmPttel", 3000).Exists Then
           Log.Error(workEnv & "թղթապանակը չի բացվել ")
           Exit Function
      End If

      If Not wStatus Then
            state = True
            AccessFolder = state
            Exit Function
      End If
      
      ' Ստուգել պայամանագրի առկայությունը թղթապանակում
      If wMDIClient.VBObject("frmPttel").VBObject("tdbgView").ApproxCount <> 1 Then
            Log.Error("Փաստաթուղթն առկա ãէ " & workEnv & " թղթապանակում")
            Exit Function
      End If
      
      state = True
      AccessFolder = state    
End Function

Function EditPeymentOrder(jurState, volort, payScale, chrgSum) 
    Dim state : state = False
      
    BuiltIn.Delay(3000)
    ' Կատարել բոլոր գործողությունները
    Call wMainForm.MainMenu.Click(c_AllActions)
    ' Խմբագրել Վճարման հանձնարարագրի պայմանագիրը
    Call wMainForm.PopupMenu.Click(c_ToEdit)
              
    If wMDIClient.WaitVBObject("frmASDocForm", 3000).Exists Then
        ' Իրավաբանական կարգավիճակ(վճարող) դաշտի լրացում
        Call Rekvizit_Fill("Document", 1, "General", "JURSTAT", jurState)
        ' Գրոծունեության ոլորտ դաշտի լրացում
        Call Rekvizit_Fill("Document", 3, "General", "VOLORT", volort)
        ' Գանձման տեսակ դաշտի լրացում
        Call Rekvizit_Fill("Document", 3, "General", "PAYSCALE", payScale)
        ' Գանձման գումար դաշտի լրացում
        Call Rekvizit_Fill("Document", 3, "General", "CHRGSUM", chrgSum)
        ' Կատարել կոճակի սեղմում
        Call ClickCmdButton(1, "Î³ï³ñ»É")  
        state = True
    Else
        Log.Error("Վճարման հանձնարարագրի պայմանագիրը չի բացվել")
        Exit Function
    End If

    EditPeymentOrder = state
End Function

' Վավերացնել վճարման հանձնարարագիրը
Sub VerifyPaymentOrder(todayD, fISN, wDocDate)   
      BuiltIn.Delay(3000)
      Call wMainForm.MainMenu.Click(c_AllActions)
      ' Վավերացնել վճարման հանձնարարագիրը
      Call wMainForm.PopupMenu.Click(c_ToProcess)
      
      If Not p1.WaitVBObject("frmASDocFormModal", 3000).Exists Then
            Log.Error("Վճարման հանձնարարագիրը չի բացվել")
            Exit Sub
      End If
      
      wDocDate = p1.VbObject("frmASDocFormModal").VbObject("TabFrame").VbObject("TDBDate").Text
      
     ' Ամսաթիվ դաշտի լրացում
     Call Rekvizit_Fill("DocumentModal", 1, "General", "DATE", todayD)
      
      ' ISN ի ստացում
      fISN = p1.VBObject("frmASDocFormModal").DocFormCommon.Doc.ISN
      Log.Message("Փաստաթղթի ISN` " & fISN)
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(4, "Î³ï³ñ»É")
      
      BuiltIn.Delay(1000)
      wMDIClient.VBObject("frmPttel").Close
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''  |BankMail ²Þî|öáË³ÝóáõÙÝ»ñ|àõÕ³ñÏí³Í ÷áË³ÝóáõÙÝ»ñ ÃÕÃ³å³Ý³ÏÇó
''  endDocISN-áí»Éù³ÛÇÝ ý³ÛÉÇ ³ÝáõÝÁ:
''  * ²ÝÑñ³Å»ßï ¿ ·ïÝí»É ïíÛ³É å³Ý³ÏáõÙ:
Public Function Get_File_Name(endDocISN)
  Dim fileOut : fileOut = NULL 

  Do Until Sys.Process("Asbank").vbObject("MainForm").Window("MDIClient", "", 1).vbObject("frmPttel").vbObject("tdbgView").EOF 
    If Trim(Sys.Process("Asbank").vbObject("MainForm").Window("MDIClient", "", 1).vbObject("frmPttel").vbObject("tdbgView").Columns.Item(3).Text) = Trim(endDocISN) Then
        fileOut = Trim(Sys.Process("Asbank").vbObject("MainForm").Window("MDIClient", "", 1).vbObject("frmPttel").vbObject("tdbgView").Columns.Item(11).Text)
        Exit Do
    Else
        Call wMDIClient.vbObject("frmPttel").vbObject("tdbgView").MoveNext
    End If
  Loop
  
  Get_File_Name = fileOut
  
End Function

'---------------------------------------------------------------------------------------------
'Èñ³óÝ»É ¹³ßïÇ ³ñÅ»ùÁ
'---------------------------------------------------------------------------------------------
'formType  - Éñ³óíáÕ ýáñÙ³ÛÇ ï»ë³ÏÁ
'    1 - Document 
'    2 - Dialog
'tabNumber - ¾çÇ Ñ³Ù³ñÁ
'rekvType  - ¹³ßïÇ ï»ë³ÏÁ
'    G - General (ÁÝ¹Ñ³Ýáõñ)
'    M - Masc   (? Ýß³ÝÝ»ñáí ¿ Éñ³óí³Í) 
'    Ch - Check Box (ÜßÇã) 
'rekvName  - ¹³ßïÇ ³ÝáõÝÁ
'rekvValue - Éñ³óíáÕ ³ñÅ»ùÁ
'    Null-Ç ¹»åùáõÙ áãÇÝã ãÇ Éñ³óíÇ, ÙÛáõë ¹»åù»ñáõÙ ¹³ßïÇ ³ñÅ»ùÁ Ï÷áË³ñÇÝíÇ ÷áË³Ýó³Í ³ñÅ»ùáí
'    Ch ï»ë³ÏÇ ¹»åùáõÙ ÷áË³Ýó»É 1(Üßí³Í ¿) Ï³Ù 0(Üßí³Í ã¿) 
Public Sub FillRekv(byval formType, byval tabNumber, byval rekvType, _
                    byval rekvName, byval rekvValue)

  Dim rekvObj, sTab, wTabStrip
    
  
    Select Case formType
      Case 1 ' Document
      
          If Not isnull(rekvValue) Then
              rekvValue = "^A" & "[Del]" & rekvValue
          End If
          
          sTab = "TabFrame"
          If tabNumber <> 1 Then
            sTab = sTab & "_" & tabNumber
            Set wTabStrip = wMDIClient.vbObject("frmASDocForm").vbObject("TabStrip")
            wTabStrip.SelectedItem = wTabStrip.Tabs(tabNumber)
          End If
          
          Select Case rekvType
            Case "G"
              rekvObj = GetVBObject(rekvName, wMDIClient.vbObject("frmASDocForm"))
              wMDIClient.vbObject("frmASDocForm").vbObject(sTab).vbObject(rekvObj).Keys(rekvValue & "[Tab]")
              
            Case "Ch"
              rekvObj = GetVBObject(rekvName, wMDIClient.vbObject("frmASDocForm"))
              wMDIClient.vbObject("frmASDocForm").vbObject(sTab).vbObject(rekvObj).Value = rekvValue
              wMDIClient.vbObject("frmASDocForm").vbObject(sTab).vbObject(rekvObj).Keys("[Tab]")
            Case Else
              Log.Error("Unknown rekvizit type of document.")
          End Select
          
      Case 2  ' Dialog 
          rekvObj = GetVBObject_Dialog(rekvName, Sys.Process("Asbank").vbObject("frmAsUstPar"))
          
          sTab = "TabFrame"
          If tabNumber <> 1 Then
            sTab = sTab & "_" & tabNumber
            
            Set wTabStrip = wMDIClient.vbObject("frmAsUstPar").vbObject("TabStrip")
            wTabStrip.SelectedItem = wTabStrip.Tabs(tabNumber)
          End If
          
          Select Case rekvType
            Case "G"
          
              If not isnull(rekvValue) Then 
                rekvValue = "^A" & "[Del]" & rekvValue
              End If
             
              Sys.Process("Asbank").vbObject("frmAsUstPar").vbObject(sTab).vbObject(rekvObj).Keys(rekvValue & "[Tab]")
              
            Case "M"  ' Masc  
              Sys.Process("Asbank").vbObject("frmAsUstPar").vbObject(sTab).vbObject(rekvObj).Keys(rekvValue & "[Tab]")
              
            Case "Ch" ' Check Box  
              Sys.Process("Asbank").vbObject("frmAsUstPar").vbObject(sTab).vbObject(rekvObj).value = rekvValue
              Sys.Process("Asbank").vbObject("frmAsUstPar").vbObject(sTab).vbObject(rekvObj).Keys("[Tab]")
            
            Case Else 
              Log.Error("Unknown rekvizit type of dialog.")
          End select
    End Select 
End Sub
                                                                                   
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Տվյալ ֆունկցիան կատարում է հիմնական ուղորդիչ ծառում, որոշակի ճանապարհով(sPath)գտնվող պանակում որոշակի
' դաշտի(nColumNumber) արժեքով(sValue) փաստաթղտի առկայության ստուգումը(օրինակ որաշակի ISN ունեցող փաստաթղթի):
Public Function Folder_Data_Check(sWorkSpace, sPath, arrDlgRekvs, nColumnNumber, sValue)
  Dim i, bExists 
  bExists = False
  
  Call ChangeWorkspace(sWorkSpace)
  
  ' Հիմնական ուղորդիչ ծառով անցում թղթապանակ
  wMDIClient.VBObject("frmExplorer").VBObject("tvTreeView").DblClickItem(sPath)
  
  If Not p1.WaitVBObject("frmAsUstPar", 3000).Exists Then
    Log.Error("Սպասվող ֆիլտրի դիալոգը չհայտնվեց")
    Exit Function
  End If
   
  For i = 0 To UBound(arrDlgRekvs) 
     Call FillRekv(2, arrDlgRekvs(i)(0), arrDlgRekvs(i)(1), arrDlgRekvs(i)(2), arrDlgRekvs(i)(3))
  Next   

  Call ClickCmdButton(2, "Î³ï³ñ»É") 
  
  ' Թղթապանակում փաստաթղթի առկայության ստուգում 
  If wMDIClient.WaitVBObject("frmPttel", 6000).Exists Then
      Do Until wMDIClient.VBObject("frmPttel").vbObject("tdbgView").EOF Or bExists = True
        If Trim(wMDIClient.VBObject("frmPttel").vbObject("tdbgView").Columns.Item(nColumnNumber).Text) = Trim(sValue) Then
            bExists = True
        Else
            Call wMDIClient.VBObject("frmPttel").vbObject("tdbgView").MoveNext
        End If
      Loop
  Else
      Log.Error("Թղթապանակը հնարավոր չեղավ բացել")
  End If
  
  Folder_Data_Check = bExists
End Function

Public Function Folder_Data_Check_MultipleValue(sWorkSpace, sPath, arrDlgRekvs, arrSearchCol, arrSearchValue)

  Dim  i, bExists
  bExists = False

  Call ChangeWorkspace(sWorkSpace)
  
  ' Հիմնական ուղորդիչ ծառով անցում թղթապանակ
  BuiltIn.Delay(1000)
  wMDIClient.VBObject("frmExplorer").VBObject("tvTreeView").DblClickItem(sPath)
  
  If Not p1.WaitVBObject("frmAsUstPar", 3000).Exists Then
    Log.Error("Սպասվող ֆիլտրի դիալոգը չհայտնվեց")
    Exit Function
  End If
   
  For i = 0 To UBound(arrDlgRekvs) 
     Call FillRekv(2, arrDlgRekvs(i)(0), arrDlgRekvs(i)(1), arrDlgRekvs(i)(2), arrDlgRekvs(i)(3))
  Next   
  
  Call ClickCmdButton(2, "Î³ï³ñ»É") 
  
  ' Թղթապանակում փաստաթղթի առկայության ստուգում
  If wMDIClient.WaitVBObject("frmPttel", 3000).Exists Then
      Do Until wMDIClient.VBObject("frmPttel").vbObject("tdbgView").EOF Or bExists = True
        For i=0 To UBound(arrSearchCol) 
            if Trim(wMDIClient.VBObject("frmPttel").vbObject("tdbgView").Columns.Item(arrSearchCol(i)).Text) <> Trim(arrSearchValue(i)) Then
              exit For
            end if
        next
        if i = UBound(arrSearchCol) + 1 then
            bExists = True
        Else
            Call wMDIClient.VBObject("frmPttel").vbObject("tdbgView").MoveNext
        End If
      Loop
  Else
      Log.Error("Թղթապանակը հնարավոր չեղավ բացել")
  End If
  
  Folder_Data_Check_MultipleValue = bExists
End Function


'BankMail-ի պահոցի հաղորդագրություններ Ֆիլտրի կլասս
Class BMStorageMsgFilter
    Public sDate
    Public eDate
    Public sendRec
    Public mt
    Public number
    Public state
    Public uniqueN
    Public view
    Public fill
    Private Sub Class_Initialize()
        sDate = "  /  /  "
        eDate = "  /  /  "
        sendRec = ""
        mt = ""
        number = ""
        state= ""
        uniqueN = ""
        view = "BMCBTINT"
        fill = "0"
    End Sub
End Class

Function New_BankMailStorageMsgFilter ()
    Set New_BankMailStorageMsgFilter = New BMStorageMsgFilter  
End Function
'Մուտք BankMail-ի պահոցի հաղորդագրություններ հաշվեվություն
Sub GoTo_BankMail_StorageMessages(bmStorageMsg, folderDirect)
    Call wTreeView.DblClickItem(folderDirect)
    BuiltIn.Delay(2000)
    If p1.WaitVBObject("frmAsUstPar",1000).Exists Then
        'Ժամանակահատվածի սկիզբ
        Call Rekvizit_Fill("Dialog",1,"General", "SDATE","[Home]![End][Del]" & bmStorageMsg.sDate)
        'Ժամանակահատվածի վերջ
        Call Rekvizit_Fill("Dialog",1,"General", "EDATE","[Home]![End][Del]" & bmStorageMsg.eDate)
        'Ուղարկաված/Ստացված
        Call Rekvizit_Fill("Dialog",1,"General", "SR","[Home]![End][Del]" & bmStorageMsg.sendRec)
        'Հաղորդագրության տիպ
        Call Rekvizit_Fill("Dialog",1,"General", "MT","[Home]![End][Del]" & bmStorageMsg.mt)
        'Համար
        Call Rekvizit_Fill("Dialog",1,"General", "NUMBER","[Home]![End][Del]" & bmStorageMsg.number) 
        'Կարգավիճակ
        Call Rekvizit_Fill("Dialog",1,"General", "STATUS","[Home]![End][Del]" & bmStorageMsg.state) 
        'Ունիկալ համար
        Call Rekvizit_Fill("Dialog",1,"General", "UNIQUE","[Home]![End][Del]" & bmStorageMsg.uniqueN)
        'Դիտելու ձև
        Call Rekvizit_Fill("Dialog",1,"General", "SELECTED_VIEW","[Home]![End][Del]" & bmStorageMsg.view) 
        'Լրացնել
        Call Rekvizit_Fill("Dialog",1,"General", "EXPORT_EXCEL","[Home]![End][Del]" & bmStorageMsg.fill)         
        Call ClickCmdButton(2, "Î³ï³ñ»É")
        Call WaitForExecutionProgress() 
    Else 
        Log.Error "Filter Window not Found"    
    End If    
End Sub

'Ընդունել BankMail Համակարգից գործողությունը կատարող ֆունկցիա
'messageCount- Ընդունվող հաղորդագրությունների քանակը
Sub Recieve_From_BankMail (messageCount)
    Call wTreeView.DblClickItem("|BankMail ²Þî|Ð³Õáñ¹³·ñáõÃÛáõÝÝ»ñÇ ÁÝ¹áõÝáõÙ|ÀÝ¹áõÝ»É BankMail Ñ³Ù³Ï³ñ·Çó")
    BuiltIn.Delay(5000) 
    If messageCount = 1 Then
        Call MessageExists(2,"Ð³Õáñ¹³·ñáõÃÛ³Ý ÁÝ¹áõÝáõÙÝ ³í³ñïí»ó") 
    ElseIF messageCount > 1 Then
        Call MessageExists(2,messageCount & " Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñÇ ÁÝ¹áõÝáõÙÝ ³í³ñïí»ó")
    Else
        Call MessageExists(2,"ÀÝ¹áõÝÙ³Ý »ÝÃ³Ï³ Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñ ãÏ³Ý")
    End If     
    If p1.WaitVBObject("frmAsMsgBox",2000).Exists Then
        Call ClickCmdButton(5, "OK")
    Else
        Log.Error "No Message Found",,,ErrorColor
    End If
End Sub



'Ստացված խառը հաղորդագրություններ Ֆիլտրի կլասս (BankMail)
Class RecievedMsgFilterBankMail
    Public sDate 
    Public eDate
    Public mt
    Public state
    Public reciever
    Public view
    Public fillInto
    Private Sub Class_Initialize
        sDate = ""
        eDate = ""
        mt = ""
        state = ""
        reciever = "1"
        view = "SWinp"
        fillInto = "0"
    End Sub
End Class
    
Function New_Recieved_Msg_BankMail()
    Set New_Recieved_Msg_BankMail = new RecievedMsgFilterBankMail
End Function    
    
    
'Ընդունված հաղորդագրություններ Ֆիլտրի լրացում (BankMail)
Sub Fill_Recieved_Messages_Filter_BankMail (recMsg)     
        'Սկզբի ամսաթիվ
        Call Rekvizit_Fill("Dialog",1,"General", "PERN","[Home]![End][Del]" & recMsg.sDate)
        'Վերջի ամսաթիվ
        Call Rekvizit_Fill("Dialog",1,"General", "PERK","[Home]![End][Del]" & recMsg.eDate)
        'Հաղ. Տիպ
        Call Rekvizit_Fill("Dialog",1,"General", "MT","[Home]![End][Del]" & recMsg.mt)
        'Կարգավիճակ
        Call Rekvizit_Fill("Dialog",1,"General", "RCCTL","[Home]![End][Del]" & recMsg.state)
        'Հասցեատեր
        Call Rekvizit_Fill("Dialog",1,"General", "TO","[Home]![End][Del]" & recMsg.reciever)
        'Դիտելու ձև
        Call Rekvizit_Fill("Dialog",1,"General", "SELECTED_VIEW","[Home]![End][Del]" & recMsg.view)
        'Լրացնել
        Call Rekvizit_Fill("Dialog",1,"General", "EXPORT_EXCEL","[Home]![End][Del]" & recMsg.fillInto)
        Call ClickCmdButton(2, "Î³ï³ñ»É")      
End Sub

'Մուտք Ընդունված հաղորդագրություններ/Ստացված Խառը հաղորդագրություններ թղթապանակ(BankMail)
Sub GoTo_Recieved_Messages_BankMail (RecievedFilter, folderDirect)
    Call wTreeView.DblClickItem(folderDirect)
    If p1.WaitVBObject("frmAsUstPar",3000).Exists Then
        Call Fill_Recieved_Messages_Filter (RecievedFilter)
        Call WaitForExecutionProgress()
    Else
        Log.Error "Can Not Open Recieved Filter",,,ErrorColor      
    End If 
End Sub


'Ստացված Փոխանցումներ Ֆիլտրի կլասս (BankMail)
Class RecievedTransFilterBankMail
    Public sDate 
    Public eDate
    Public mt
    Public bankCode
    Public state
    Public reciever
    Public view
    Public fillInto
    Private Sub Class_Initialize
        sDate = ""
        eDate = ""
        mt = ""
        bankCode = ""
        state = ""
        reciever = "1"
        view = "SWinp"
        fillInto = "0"
    End Sub
End Class
    
Function New_Recieved_Trans_BankMail()
    Set New_Recieved_Trans_BankMail = new RecievedTransFilterBankMail
End Function    
    
    
'Ընդունված Փոխանցումներ Ֆիլտրի լրացում (BankMail)
Sub Fill_Recieved_Transfer_Filter_BankMail (recTrans)     
        'Սկզբի ամսաթիվ
        Call Rekvizit_Fill("Dialog",1,"General", "PERN","[Home]![End][Del]" & recTrans.sDate)
        'Վերջի ամսաթիվ
        Call Rekvizit_Fill("Dialog",1,"General", "PERK","[Home]![End][Del]" & recTrans.eDate)
        'Հաղ. Տիպ
        Call Rekvizit_Fill("Dialog",1,"General", "MT","[Home]![End][Del]" & recTrans.mt)
        'Բանկի կոդ
        Call Rekvizit_Fill("Dialog",1,"General", "BANKCTL","[Home]![End][Del]" & recTrans.bankCode)        
        'Կարգավիճակ
        Call Rekvizit_Fill("Dialog",1,"General", "RCCTL","[Home]![End][Del]" & recTrans.state)
        'Հասցեատեր
        Call Rekvizit_Fill("Dialog",1,"General", "TO","[Home]![End][Del]" & recTrans.reciever)
        'Դիտելու ձև
        Call Rekvizit_Fill("Dialog",1,"General", "SELECTED_VIEW","[Home]![End][Del]" & recTrans.view)
        'Լրացնել
        Call Rekvizit_Fill("Dialog",1,"General", "EXPORT_EXCEL","[Home]![End][Del]" & recTrans.fillInto)
        Call ClickCmdButton(2, "Î³ï³ñ»É")      
End Sub

'Մուտք Ընդունված հաղորդագրություններ/Ստացված Փոխանցումներ թղթապանակ(BankMail)
Sub GoTo_Recieved_Transfer_BankMail (RecievedFilter, folderDirect)
    Call wTreeView.DblClickItem(folderDirect)
    If p1.WaitVBObject("frmAsUstPar",3000).Exists Then
        Call Fill_Recieved_Transfer_Filter_BankMail (RecievedFilter)
        Call WaitForExecutionProgress()
    Else
        Log.Error "Can Not Open Recieved Filter",,,ErrorColor      
    End If 
End Sub

'BankMail -ի MT100 Վճարման հանձնարարագիր փաստաթղթի Ընդհանուր էջի կլասս
Class BankMailMT100Common
    Public isn
    Public div
    Public dep
    Public note
    Public docN
    Public fDate
    Public payerAcc
    Public regNum
    Public payer
    Public recieverAcc
    Public reciever    
    Public sum
    Public cur
    Public aim
    Public tabN
    Public check
    Sub Class_Initialize
        isn = ""
        div = ""
        dep = ""
        note = ""
        docN = ""
        fDate = "  /  /  "
        payerAcc = ""
        regNum = ""
        payer = ""
        recieverAcc = ""
        reciever = ""
        sum = "0.00"
        cur = ""
        aim = ""
        tabN = 1
        check = False   
    End Sub
End Class

Function New_BankMailMT100Common()
    Set New_BankMailMT100Common = new BankMailMT100Common
End Function

'BankMail -ի MT100 Վճարման հանձնարարագիր փաստաթղթի Ընդհանուր էջի ստուգում
Sub BankMail_MT100_CommonTab_Check(bmMT100Com)
    Call GoTo_ChoosedTab(bmMT100Com.tabN)
    'Գրասենյակ դաշտի ստուգում
    Call Compare_Two_Values("Գրասենյակ",Get_Rekvizit_Value("Document",bmMT100Com.tabN,"Mask","ACSBRANCH"),bmMT100Com.div)
    'Բաժին դաշտի ստուգում
    Call Compare_Two_Values("Բաժին",Get_Rekvizit_Value("Document",bmMT100Com.tabN,"Mask","ACSDEPART"),bmMT100Com.dep)
    'Նշում դաշտի ստուգում
    Call Compare_Two_Values("Նշում",Get_Rekvizit_Value("Document",bmMT100Com.tabN,"Mask","PAYNOTE"),bmMT100Com.note)
    'Փաստաթղթի N դաշտի ստուգում
    Call Compare_Two_Values("Փաստաթղթի N",Get_Rekvizit_Value("Document",bmMT100Com.tabN,"General","BMDOCNUM"),bmMT100Com.docN)
    'Ամսաթիվ դաշտի ստուգում
    Call Compare_Two_Values("Ամսաթիվ",Get_Rekvizit_Value("Document",bmMT100Com.tabN,"General","DATE"),bmMT100Com.fDate)
    'Վճարողի Հաշիվ դաշտի ստուգում
    Call Compare_Two_Values("Վճարողի Հաշիվ",Get_Rekvizit_Value("Document",bmMT100Com.tabN,"Bank","ACCDB"),bmMT100Com.payerAcc)  
    'Սոցիալական քարտ դաշտի ստուգում
    Call Compare_Two_Values("Սոցիալական քարտ",Get_Rekvizit_Value("Document",bmMT100Com.tabN,"General","SOCCARD"),bmMT100Com.regNum)    
    'Վճարող դաշտի ստուգում
    Call Compare_Two_Values("Վճարող",Get_Rekvizit_Value("Document",bmMT100Com.tabN,"Comment","PAYER"),bmMT100Com.payer)
    'Ստացողի Հաշիվ դաշտի ստուգում
    Call Compare_Two_Values("Ստացողի Հաշիվ",Get_Rekvizit_Value("Document",bmMT100Com.tabN,"Bank","ACCCR"),bmMT100Com.recieverAcc)
    'Ստացող դաշտի ստուգում
    Call Compare_Two_Values("Ստացող",Get_Rekvizit_Value("Document",bmMT100Com.tabN,"Comment","RECEIVER"),bmMT100Com.reciever)
    'Գումար դաշտի ստուգում
    Call Compare_Two_Values("Գումար",Get_Rekvizit_Value("Document",bmMT100Com.tabN,"General","SUMMA"),bmMT100Com.sum)
    'Արժույթ դաշտի ստուգում
    Call Compare_Two_Values("Արժույթ",Get_Rekvizit_Value("Document",bmMT100Com.tabN,"Mask","CUR"),bmMT100Com.cur)
    'Նպատակ դաշտի ստուգում
    Call Compare_Two_Values("Նպատակ",Get_Rekvizit_Value("Document",bmMT100Com.tabN,"Comment","AIM"),bmMT100Com.aim)
End Sub

'BankMail -ի MT100 Վճարման հանձնարարագիր փաստաթղթի Լրացուցիչ էջի կլասս
Class BankMailMT100Add
    Public pack
    Public payDate
    Public fileName
    Public dirName
    Public srDate
    Public repayDate
    Public tabN
    Public check
    
    Private Sub Class_Initialize
        pack = ""
        payDate = "  /  /  "
        fileName = ""
        dirName = ""
        srDate = "  /  /  "
        repayDate = "  /  /  "
        tabN = 2
        check = False         
    End Sub
End Class

Function New_BankMailMT100Add()
    Set New_BankMailMT100Add = new BankMailMT100Add
End Function

'BankMail -ի MT100 Վճարման հանձնարարագիր փաստաթղթի Լրացուցիչ էջի ստուգում
Sub BankMail_MT100_AddTab_Check(bmMT100Add)
    Call GoTo_ChoosedTab(bmMT100Add.tabN)
    'Փաթեթի համար դաշտի ստուգում
    Call Compare_Two_Values("Փաթեթի համար",Get_Rekvizit_Value("Document",bmMT100Add.tabN,"General","ACSBRANCH"),bmMT100Add.pack)
    'Վճարման օր դաշտի ստուգում
    Call Compare_Two_Values("Վճարման օր",Get_Rekvizit_Value("Document",bmMT100Add.tabN,"General","ACSBRANCH"),bmMT100Add.payDate)
    'Ֆայլի անուն/Հաղ. N դաշտի ստուգում
    Call Compare_Two_Values("Ֆայլի անուն/Հաղ. N",Get_Rekvizit_Value("Document",bmMT100Add.tabN,"General","ACSBRANCH"),bmMT100Add.fileName)
    'Դիրեկտորիայի անուն դաշտի ստուգում
    Call Compare_Two_Values("Դիրեկտորիայի անուն",Get_Rekvizit_Value("Document",bmMT100Add.tabN,"General","ACSBRANCH"),bmMT100Add.dirName)
    'Ամսաթիվ(Ուղարկման/Ստացման) դաշտի ստուգում
    Call Compare_Two_Values("Ամսաթիվ(Ուղարկման/Ստացման)",Get_Rekvizit_Value("Document",bmMT100Add.tabN,"General","ACSBRANCH"),bmMT100Add.srDate)
    'Մարման ամսաթիվ դաշտի ստուգում
    Call Compare_Two_Values("Մարման ամսաթիվ",Get_Rekvizit_Value("Document",bmMT100Add.tabN,"General","ACSBRANCH"),bmMT100Add.repayDate)
End Sub   

Class BankMailMT100 
    Public commonTab
    Public addTab
    Private Sub Class_Initialize
        Set commonTab = New_BankMailMT100Common()
        Set addTab = New_BankMailMT100Add()
    End Sub
End Class

Function New_PaymentOrderRecieved()
    Set New_PaymentOrderRecieved = New BankMailMT100
End Function

'BankMail -ի MT100 Վճարման հանձնարարագիր փաստաթղթի ստուգում
Sub BankMail_MT100_Check(MT100)
    'Փաստաթղթի isn-ի ստացում
    MT100.commonTab.isn = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
    'Ընդհանուր
    If MT100.commonTab.check Then
        Call BankMail_MT100_CommonTab_Check(MT100.commonTab)
    End If
    'Լրացուցիչ
    If MT100.addTab.check Then
        Call BankMail_MT100_AddTab_Check(MT100.addTab)
    End If
End Sub 