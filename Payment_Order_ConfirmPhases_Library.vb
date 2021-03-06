Option Explicit
'USEUNIT Library_Common
'USEUNIT Online_PaySys_Library
'USEUNIT Constants
'USEUNIT Library_Contracts
'USEUNIT Library_Colour
Dim fCount, lCount, dCount
'----------------------------------------------------------------------
'Վճարման հանձնարարագիր (ուղ.) տեսակի վճարային փաստաթղթի լրացում
'----------------------------------------------------------------------

'office - Գրասենյակ/Բաժին դաշտի արժեքը
'department - Հաշվապահություն դաշտի արժեքը
'docNumber - Փաստաթղթի համարը
'data - Ամսաթիվ դաշտի արժեքը
'accDeb - true արժեքի դեպքում սեղմվում է Հաշիվ դեբետ կոճակը
'chart - Հաշվային պլանի համար դաշտի արժեքը
'balAcc - Հ/Պ հաշվեկշռային հաշիվ դաշտի արժեքը
'accMask - Հաշվի շաբլոն դաշտի արժեքը
'accCur - Արժույթ դաշտի արժեքը
'accType - Հաշվի տիպ դաշտի արժեքը
'clientName - Հաճախորդի անվանում դաշտի արժեքը
'client - Հաճախորդ դաշտի արժեքը
'newAcc - Նոր հաշիվ դաշտի արժեքը
'note1 - Նշում դաշտի արժեքը
'note2 - Նշում 2 դաշտի արժեքը
'note3 - Նշում 3 դաշտի արժեքը
'branch - Գրասենյակ դաշտի արժեքը
'depart - Բաժին դաշտի արժեքը
'acsType - Հասան-ն տիպ դաշտի արժեքը
'cardNum - Քարտի համար դաշտի արժեքը
'payer - Վճարող դաշտի արժեքը
'epayer - Վճարող(անգլ) դաշտի արժեքը
'taxCod - ՀՎՀՀ (Վճարող) դաշտի արժեքը
'socCard - Սոցիալական քարտ դաշտի արժեքը
'accCredit - Հաշիվ կրեդիտ դաշտի արժեքը
'receiver - Ստացող դաշտի արժեքը
'eReceiver - Ստացող (անգլ.) դաշտի արժեքը
'summa - Գումար դաշտի արժեքը
'currency - Արժույթ դաշտի արժեքը
'aim - Նպատակ դաշտի արժեքը
'fISN - Փաստաթղթի ISN
Sub PayOrder_Send_Fill(office, department, docNumber, data, accDeb, acDBValue, chart, balAcc, accMask, accCur, accType, clientName, client, _
                       note1, note2, note3, branch, depart, acsType, cardNum, payer, epayer, taxCod , socCard, accCredit, _
                       receiver, eReceiver, summa, curr, aim , fISN)
    
  Dim rekvObj
  
  BuiltIn.Delay(3000)
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_PayOrds & "|" & c_PayOrdToBeSent)
  BuiltIn.Delay(1000)
        
  'Ստեղծվող ISN - ի փաստատթղթի  վերագրում փոփոխականին
  fISN = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
  'Գրասենյակ դաշտի լրացում
  Call Rekvizit_Fill("Document", 1, "General", "ACSBRANCH", office)
  'Բաժին դաշտի լրացում
  Call Rekvizit_Fill("Document", 1, "General", "ACSDEPART", department)
    
  'Փաստաթղթի N դադաշտիշտի արժեքի վերագրում փոփոխականին
  rekvObj = GetVBObject("DOCNUM", wMDIClient.vbObject("frmASDocForm"))
  docNumber = wMDIClient.vbObject("frmASDocForm").vbObject("TabFrame").vbObject(rekvObj).Text
    
  ' Ամսաթիվ դաշտի լրացում
  Call Rekvizit_Fill("Document", 1, "General", "DATE", data)
    
  If accDeb Then
    'Հաշիվ դեբետ կոճակի սեղմում
    wMDIClient.vbObject("frmASDocForm").vbObject("TabFrame").vbObject("ASAmACC").vbObject("CmdViewUser").Click()
    'Հաշվային պլանի համար դաշտի լրացում
    Call Rekvizit_Fill("Dialog", 1, "General", "COANUM", "1")
    'Հ/Պ հաշվեկշռային հաշիվ դաշտի լրացում
    Call Rekvizit_Fill("Dialog", 1, "General", "BALACC", balAcc)
    'Հաշվի շաբլոն դաշտի լրացում
    Call Rekvizit_Fill("Dialog", 1, "General", "AccMask", accMask)
    'Արժույթ դաշտի լրացում
    Call Rekvizit_Fill("Dialog", 1, "General", "AccCur", accCur)
    'Հաշվի տիպ դաշտի լրացում
    Call Rekvizit_Fill("Dialog", 1, "General", "AccType", accType)
    'Հաճախորդի անվանում դաշտի լրացում
    Call Rekvizit_Fill("Dialog", 1, "General", "ClName", clientName)
    'Հաճախորդ դաշտի լրացում
    Call Rekvizit_Fill("Dialog", 1, "General", "CliCod", client)
    'Նշում դաշտի լրացում
    Call Rekvizit_Fill("Dialog", 1, "General", "AccNote", note1)
    'Նշում2 դաշտի լրացում
    Call Rekvizit_Fill("Dialog", 1, "General", "AccNote2", note2)
    'Նշում3 դաշտի լրացում
    Call Rekvizit_Fill("Document", 1, "General", "AccNote3", note3)
    'Գրասենյակ դաշտի լրացում
    Call Rekvizit_Fill("Dialog", 1, "General", "AcsBranch", branch)
    'Բաժին դաշտի լրացում
    Call Rekvizit_Fill("Dialog", 1, "General", "AcsDepart", depart)
    'Հասանելության տիպ դաշտի լրացում
    Call Rekvizit_Fill("Dialog", 1, "General", "AcsType", acsType)
    'Կատարել կոճակի սեղմում
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    p1.vbObject("frmModalBrowser").vbObject("tdbgView").Keys("[Enter]")
  Else
    Call Rekvizit_Fill("Document", 1, "General", "ACCDB", "^A[Del]" & acDBValue)
  End If
       
  'Վճարող դաշտի լրացում
  Call Rekvizit_Fill("Document", 1, "General", "PAYER", payer)
  'Վճարող անգլերեն դաշտի լրացում
  '    rekvName = GetVBObject("EPAYER", wMDIClient.vbObject("frmASDoacForm"))
  '    wMDIClient.vbObject("frmASDocForm").vbObject("TabFrame").vbObject(rekvName).Keys(epayer & "[Tab]")
  'ՀՎՀՀ(վճարող) դաշտի լրացում
  Call Rekvizit_Fill("Document", 1, "General", "TAXCODSD", taxCod)           
  'Սոցիալական քարտ դաշտի լրացում
  '    rekvName = GetVBObject("REGNUM", wMDIClient.vbObject("frmASDocForm"))
  '    wMDIClient.vbObject("frmASDocForm").vbObject("TabFrame").vbObject(rekvName).Keys(socCard & "[Tab]")
  'Հաշիվ կրեդիտ դաշտի լրացում
  Call Rekvizit_Fill("Document", 1, "General", "ACCCR", "^A[Del]" & accCredit)
  'Ստացող դաշտի լրացում
  Call Rekvizit_Fill("Document", 1, "General", "RECEIVER", receiver)
  'Ստացող(անգլ.) դաշտի լրացում
  Call Rekvizit_Fill("Document", 1, "General", "ERECEIVER", eReceiver)
  'Գումար դաշտի լրացում
  Call Rekvizit_Fill("Document", 1, "General", "SUMMA", summa)
  'Արժույթ դաշտի լրացում
  Call Rekvizit_Fill("Document", 1, "General", "CUR", curr)
  'Նպատակ դաշտի լրացում
  Call Rekvizit_Fill("Document", 1, "General", "AIM", aim)
  '   'Դրամարկղ դաշտի լրացում
  '    Call Rekvizit_Fill("Document",4,"General","KASSA","01")
  '    Call Rekvizit_Fill("Document",4,"General","KASSIMV","021")                
  'Կատարել կոճակի սեղմում
  Call ClickCmdButton(1, "Î³ï³ñ»É")
    
  'Տպելու ձև պատուհանի փակում
  BuiltIn.Delay(delay_small)
  wMDIClient.vbObject("FrmSpr").Close
End Sub

'----------------------------------------------------------------------
'ì×³ñÙ³Ý Ñ³ÝÓÝ³ñ³ñ³·Çñ (áõÕ.) -Ç "ì³í»ñ³óÝ»É" ·áñÍáÕáõÃÛ³Ý Ï³ï³ñáõÙ:
'üáõÝÏóÇ³Ý »ÝÃ³¹ñáõÙ ¿ ÷³ëï³ÃÕÃÇ ³éÏ³ÛáõÃÛáõÝÁ :
'----------------------------------------------------------------------
'verify - true ³ñÅ»ùÇ ¹»åùáõÙ Ñ³ëï³ïíáõÙ ¿ ,  fasle-Ç ¹»åùáõÙ ` áã
Sub PaySys_Verify(verify)
  BuiltIn.Delay(3000)
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_ToConfirm)
  If wMDIClient.WaitVBObject("frmASDocForm", 3000).Exists Then 
    If verify Then
      Call ClickCmdButton(1, "Ð³ëï³ï»É")
    Else
      Call ClickCmdButton(1, "Ø»ñÅ»É")
      Call ClickCmdButton(11, "Î³ï³ñ»É") 
    End If    
  Else 
    Log.Error "Can't open frmASDocForm widow.", "", pmNormal, ErrorColor
  End If
End Sub

'------------------------------------------------------------------------------
'ì»ñëïáõ·íáÕ ÷³ëï³ÃÕÃ»ñ ÃÕÃ³å³Ý³ÏáõÙ í×³ñÙ³Ý Ñ³ÝÓÝ³ñ³ñ·ñÇ ³éÏ³ÛáõÃÛ³Ý ëïáõ·áõÙ:
'ýáõÝÏó³Ç³Ý í»ñ³¹³ñÓÝáõÙ ¿ true , »Ã» ÷³ëï³ÃáõÕÃÁ ³éÏ³ ¿ , false` »Ã» µ³ó³Ï³ÛáõÙ:
'------------------------------------------------------------------------------
'docNum - ÷³ëï³ÃÕÃÇ Ñ³Ù³ñ
Function PaySys_Check_Doc_In_InspecdetDoc_Folder(docNum)
  Dim is_exists : is_exists = False
  Dim colN
    
  BuiltIn.Delay(3000)
  Call wTreeView.DblClickItem("|ÎñÏÝ³ÏÇ Ùáõïù³·ñáÕÇ ²Þî                 |ÎñÏÝ³ÏÇ Ùáõïù³·ñíáÕ ÷³ëï³ÃÕÃ»ñ")
    
  If wMDIClient.WaitVBObject("frmPttel", 2000).Exists Then
    colN = wMDIClient.vbObject("frmPttel").GetColumnIndex("DOCNUM")
    If SearchInPttel("frmPttel", colN, docNum) Then
      is_exists = true
    Else 
      Log.Error "Can't find serached row where document N = " & docNum, "", pmNormal, ErrorColor
    End If
  Else
    Log.Error "The double input frmPttel does't exist", "", pmNormal, ErrorColor
  End If
    
  PaySys_Check_Doc_In_InspecdetDoc_Folder = is_exists
End Function

'------------------------------------------------------------------------------
'Վերստուգվող փաստաթղթեր թղթապանակում վճարման հանձնարարգրի Հաստատում :Եթե նոր
'ներմուծված արժեքները սխալ են, ապա ֆունկցիան վերադարձնում է fasle, եթե հաստատվում է ` true :
'Ֆունկցիան ենթադրում է փաստաթղթի առկայությունը :
'------------------------------------------------------------------------------
'accCred - Հաշիվ կրեդիտ դաշտի արժեքը
'summa - Գումար դաշտի արժեքը
Function PaySys_Verify_Doc_In_InspecdetDoc_Folder(accCred, summa)
    Dim isverify : isverify = False
    
    BuiltIn.Delay(3000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    wMainForm.PopupMenu.click(c_DoubleInput)
    If wMDIClient.WaitVBObject("frmASDocForm", 3000).Exists Then
        'Հաշիվ կրեդիտ դաշտի լրացում
        Call Rekvizit_Fill("Document", 1, "General", "AccCr", "^A[Del]" & accCred)
        'Գումար դաշտի լրացում
        Call Rekvizit_Fill("Document", 1, "General", "Summa", summa)
        Call ClickCmdButton(1, "Î³ï³ñ»É")
    Else
        Log.Message("Can't open the document for double input")
    End If
    If p1.WaitVBObject("frmAsMsgBox", 2000).Exists Then
        Call ClickCmdButton(3, "²Ûá")
    Else
        isverify = True
    End If
    
    PaySys_Verify_Doc_In_InspecdetDoc_Folder = isverify
End Function

'------------------------------------------------------------------------------
'²ñï³ùÇÝ ÷áË³ÝóáõÙÝ»ñÇ áõÕ³ñÏíáÕ Ñ³ÝÓÝ³ñ³·ñ»ñ ÃÕÃ³å³Ý³ÏáõÙ í×³ñÙ³Ý Ñ³ÝÓÝ³ñ³ñ³·ñÇ
'³éÏ³ÛáõÃÛ³Ý ëïáõ·áõÙ :üáõÝÏóÇ³Ý í»ñ³¹³ñÓÝáõÙ ¿ true, »Ã» ÷³ëï³ÃáõÕÃÁ ³éÏ³ ¿ ¨
'false` »Ã» ³ÛÝ µ³ó³Ï³ÛáõÙ ¿:
'------------------------------------------------------------------------------
'startDate - àÕ³ñÏíáÕ Ñ³ÝÓÝ³ñ³ñ·ñ»ñ ýÇÉïÇ Å³Ù³Ý³Ï³Ñ³ïí³Í ¹³ßïÇ ëÏ½µÝ³Ï³Ý ³Ùë³ÃÇí
'endDate - àÕ³ñÏíáÕ Ñ³ÝÓÝ³ñ³ñ·ñ»ñ ýÇÉïÇ Å³Ù³Ý³Ï³Ñ³ïí³Í ¹³ßïÇ í»ñçÝ³Ï³Ý ³Ùë³ÃÇí
'docNum - ö³ëï³ÃÕÃÇ Ñ³Ù³ñÁ
Function PaySys_Check_Doc_In_ExternalTransfer_Folder(startDate, endDate , docNum)
  Dim is_exists : is_exists = false
  Dim colN
  
  BuiltIn.Delay(delay_middle)
  Call wTreeView.DblClickItem("|²ñï³ùÇÝ ÷áË³ÝóáõÙÝ»ñÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|àõÕ³ñÏíáÕ Ñ³ÝÓÝ³ñ³ñ³·ñ»ñ|àõÕ³ñÏíáÕ Ñ³ÝÓÝ³ñ³ñ³·ñ»ñ")
  If p1.WaitVBObject("frmAsUstPar", 3000).Exists Then
    Call Rekvizit_Fill("Dialog", 1, "General", "PERN", startDate)
    Call Rekvizit_Fill("Dialog", 1, "General", "PERK", endDate)
    Call Rekvizit_Fill("Dialog", 1, "General", "USER", "^A" & "[Del]")
    Call ClickCmdButton(2, "Î³ï³ñ»É")
  Else 
    Log.Error "Can't find frmAsUstPar window", "", pmNormal, ErrorColor
  End If

  If wMDIClient.WaitVBObject("frmPttel", 3000).Exists Then
    colN = wMDIClient.vbObject("frmPttel").GetColumnIndex("DOCNUM")
    If SearchInPttel("frmPttel", colN, docNum) Then
      is_exists = true
    End If
  Else
    Log.Message "The sending documnet frmPttel doesn't exist", "", pmNormal, ErrorColor
  End If
  
  PaySys_Check_Doc_In_ExternalTransfer_Folder = is_exists
End Function

'----------------------------------------------------------------------------------
'Ուղարկվող հանձնարագրեր թղթապանակից վճարման հանձնարարագրի ուղարկում BankMail բաժին :
'---------------------------------------------------------------------------------
Sub PaySys_Sendto_BankMail()
  BuiltIn.Delay(delay_middle)
  Call wMainForm.MainMenu.Click(c_AllActions)
  wMainForm.PopupMenu.click(c_SendToBM)
  BuiltIn.Delay(delay_middle)

  If p1.WaitVBObject("frmAsMsgBox", delay_middle).Exists Then
    Call ClickCmdButton(5, "²Ûá")
    if p1.WaitVBObject("frmAsMsgBox", 1000).Exists Then
      Call ClickCmdButton(5, "OK")
    End if
  Else
    Log.Message "Message box must be exist", "", pmNormal, ErrorColor
  End If
End Sub

'----------------------------------------------------------------------------------------
'BankMAil-Ç àõÕ³ñÏíáÕ ÷áË³ÝóáõÙÝ»ñ ÃÕÃ³å³Ý³ÏáõÙ í×³ñÙ³Ý Ñ³ÝÓÝ³ñ³ñ·ñÇ ³éÏ³ÛáõÃÛ³Ý ëïáõ·áõÙ:
'ºÃ» Ñ³ÝÓÝ³ñ³ñ·ÇñÁ ³éÏ³ ¿ , ³å³ ýáõÝÏóÇ³Ý í»ñ³¹³ñÓÝáõÙ ¿ true, fasle ` »Ã» ³ÛÝ µ³ó³Ï³ÛáõÙ ¿ :
'----------------------------------------------------------------------------------------
'startDate - àõÕ³ñÏíáÕ ýÇÉïñÇ êÏ½µÇ ³Ùë³ÃÇí ¹³ßïÇ ³ñÅ»ù
'endDate - àõÕ³ñÏíáÕ ýÇÉïñÇ ì»ñçÇ ³Ùë³ÃÇí ¹³ßïÇ ³ñÅ»ù
'docISN - ö³ëï³ÃÕÃÇ ISN-Á
Function PaySys_Check_Doc_In_BankMail_Folder(startDate, endDate , docISN)
  Dim is_exists : is_exists = false
  Dim colN
    
  Call wTreeView.DblClickItem("|BankMail ²Þî|àõÕ³ñÏíáÕ Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñ|àõÕ³ñÏíáÕ ÷áË³ÝóáõÙÝ»ñ")
  If p1.WaitVBObject("frmAsUstPar", 3000).Exists Then
    Call Rekvizit_Fill("Dialog", 1, "General", "PERN", startDate)
    Call Rekvizit_Fill("Dialog", 1, "General", "PERK", endDate)
    Call ClickCmdButton(2, "Î³ï³ñ»É")
  Else 
    Log.Error "Can't find frmAsUstPar window", "", pmNormal, ErrorColor
  End If
  
  wMDIClient.Refresh
  BuiltIn.Delay(delay_middle)
  If wMDIClient.WaitVBObject("frmPttel", delay_middle).Exists Then
    colN = wMDIClient.vbObject("frmPttel").GetColumnIndex("DOCNUM")
    If SearchInPttel("frmPttel", colN, docISN) Then
      is_exists = true
    Else 
      Log.Error "Can't find serached row where document isn = " & docISN, "", pmNormal, ErrorColor
    End If
  Else
    Log.Message "The sending documnet frmPttel doesn't exist", "", pmNormal, ErrorColor
  End If
  PaySys_Check_Doc_In_BankMail_Folder = is_exists
End Function

'----------------------------------------------------------------------------------------
'ö³ëï³ÃÕÃÇ áõÕ³ñÏáõÙ Ù³ëÝ³ÏÇ ËÙµ³·ñÙ³Ý
'----------------------------------------------------------------------------------------
Function PaySys_SendTo_Partial_Edit()
    BuiltIn.Delay(3000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    wMainForm.PopupMenu.click(c_SendToPartEd)
    If p1.WaitVBObject("frmAsUstPar", 3000).Exists Then
      Call ClickCmdButton(2, "Î³ï³ñ»É")
    Else 
      Log.Error "Can't find frmAsUstPar window", "", pmNormal, ErrorColor
    End If
End Function

'----------------------------------------------------------------------------------------
'ö³ëï³ÃÕÃÇ çÝçáõÙ Ñ³ßíÇ ³éÝ»Éáí Ñ³ßíÇíÁ Ï³ÝËÇÏ³ÛÇÝ ¿ Ã» áã :
'----------------------------------------------------------------------------------------
Sub Paysys_Delete_Doc(cash)
  BuiltIn.Delay(3000)
  Call wMainForm.MainMenu.Click(c_AllActions)
  wMainForm.PopupMenu.Click(c_Delete)
  BuiltIn.Delay(delay_middle)
  If cash Then  
    Call ClickCmdButton(5, "Î³ï³ñ»É")
  End If
  Call ClickCmdButton(3, "²Ûá")
End Sub                 

'----------------------------------------------------------------------
' ö³ëï³ÃÕÃÇ ³éÏ³ÛáõÃÛ³Ý ëïáõ·áõÙ 2-ñ¹ Ð³ëï³ïáÕÇ Ùáï
'----------------------------------------------------------------------
'docNum - ö³ëï³ÃÕÃÇ
'startDate - üÇÉïñáõÙ Éñ³óíáÕ ³Ùë³ÃÇí(êÏÇ½µ)
'endDate - üÇÉïñáõÙ Éñ³óíáÕ ³Ùë³ÃÇí(ì»ñç)
Function PaySys_Check_Doc_In_Verifier(docNum, startDate, endDate, workpaper)
  Dim exist, verifyDocuments, colN
  exist = False
    
  BuiltIn.Delay(3000)
  wMDIClient.Refresh
    
  Set verifyDocuments = New_VerificationDocument()
  verifyDocuments.User = "^A[Del]"
  Call GoToVerificationDocument(workpaper, verifyDocuments)

  If wMDIClient.WaitVBObject("frmPttel", 3000).Exists Then
    colN = wMDIClient.vbObject("frmPttel").GetColumnIndex("DOCNUM")
    If SearchInPttel("frmPttel", colN, docNum) Then
      exist = True
    Else 
      Log.Error "Can't find row with " & docNum & "document number", "", pmNormal, ErrorColor
    End If
  Else
      Log.Error "Verifiers folder view doesn't exists", "", pmNormal, ErrorColor
  End If
    
  PaySys_Check_Doc_In_Verifier = exist
End Function

'----------------------------------------------------------------------
'ì×³ñÙ³Ý Ñ³ÝÓÝ³ñ³ñ³·Çñ (áõÕ.) -Ç "ì³í»ñ³óÝ»É" ·áñÍáÕáõÃÛ³Ý Ï³ï³ñáõÙ:
'üáõÝÏóÇ³Ý »ÝÃ³¹ñáõÙ ¿ ÷³ëï³ÃÕÃÇ ³éÏ³ÛáõÃÛáõÝÁ :
'----------------------------------------------------------------------
'verify - true ³ñÅ»ùÇ ¹»åùáõÙ Ñ³ëï³ïíáõÙ ¿ ,  fasle-Ç ¹»åùáõÙ ` áã
Sub PaySys_Send_To_CheckUp()
  BuiltIn.Delay(3000)
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_SendToDoubleInput)
  If p1.WaitVBObject("frmAsMsgBox", 3000).Exists Then
    Call ClickCmdButton(5, "²Ûá")
  Else
    Log.Error "Can't find frmAsMsgBox window", "", pmNormal, ErrorColor
  End If
End Sub

'----------------------------------------------------------------------
'ì×³ñÙ³Ý Ñ³ÝÓÝ³ñ³ñ³·Çñ (áõÕ.) -Ç "àõÕ³ñÏ»É Ñ³ëï³ïÙ³Ý" ·áñÍáÕáõÃÛ³Ý Ï³ï³ñáõÙ:
'üáõÝÏóÇ³Ý »ÝÃ³¹ñáõÙ ¿ ÷³ëï³ÃÕÃÇ ³éÏ³ÛáõÃÛáõÝÁ :
'----------------------------------------------------------------------
Sub PaySys_Send_To_Verify()
    BuiltIn.Delay(3000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_SendToVer)
    If p1.WaitVBObject("frmAsMsgBox", 3000).Exists Then 
      Call ClickCmdButton(5, "²Ûá")
    Else 
      Log.Error "Can't open frmAsMsgBox window", "", pmNormal, ErrorColor
    End If
End Sub

'----------------------------------------------------------------------
'ì×³ñÙ³Ý Ñ³ÝÓÝ³ñ³ñ³·Çñ (áõÕ.) -Ç "àõÕ³ñÏ»É ³ñï³ùÇÝ µ³ÅÇÝ" ·áñÍáÕáõÃÛ³Ý Ï³ï³ñáõÙ:
'üáõÝÏóÇ³Ý »ÝÃ³¹ñáõÙ ¿ ÷³ëï³ÃÕÃÇ ³éÏ³ÛáõÃÛáõÝÁ :
'----------------------------------------------------------------------
Sub PaySys_Send_To_External()
    BuiltIn.Delay(3000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_SendToExternalSec)
    BuiltIn.Delay(1000)
    Call ClickCmdButton(5, "²Ûá")
End Sub

'----------------------------------------------------------------------
'²¹ÙÇÝÇëïñ³ïáñÇ ²Þî-áõÙ ³í»É³óÝáõÙ ¿ "ÆÙ ÷³ëï³ÃÕÃ»ñ" ÃÕÃ³å³Ý³ÏÁ 
'"Ü»ñÙáõÍ»ÉÏ³ñ·³íáñáõÙÝ»ñÁ" ÑÝ³ñ³íáñáõÃÛ³Ùµ :
'----------------------------------------------------------------------
Sub Insert_MyDocs()
  Dim colN
  
  BuiltIn.Delay(delay_middle)
  Call wMDIClient.VBObject("frmExplorer").VBObject("tvTreeView").DblClickItem("|²¹ÙÇÝÇëïñ³ïáñÇ ²Þî|ú·ï³·áñÍáÕÝ»ñ ¨ ²Þî|²ßË³ï³Ýù³ÛÇÝ ï»Õ»ñ (²Þî)")
  If wMDIClient.WaitVBObject("frmPttel", delay_middle).Exists Then
    colN = wMDIClient.vbObject("frmPttel").GetColumnIndex("CODE")
    If Not SearchInPttel("frmPttel", colN, "GlavBux") Then
      Log.Message "Can't find searched row where Code = GlavBux", "", pmNormal, ErrorColor
    End If
  Else 
    Log.Message "Can't find frmPttel window", "", pmNormal, ErrorColor
  End If
  
  BuiltIn.Delay(delay_middle)
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click("Դիզայներ")
  BuiltIn.Delay(delay_middle)
  wMDIClient.VBObject("frmEditNavTreeNew").VBObject("TreeView").Keys("[End]")
    
  wMainForm.VBObject("tbToolBar").Window("ToolbarWindow32", "", 1).ClickItem(6)
  wMDIClient.VBObject("frmEditNavTreeNew").Close
  wMDIClient.vbObject("frmPttel").vbObject("tdbgView").MoveFirst
  colN = wMDIClient.vbObject("frmPttel").GetColumnIndex("CODE")
  Call SearchInPttel("frmPttel", colN, "ADMIN")
  
  BuiltIn.Delay(delay_middle)
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click("Դիզայներ")
  BuiltIn.Delay(delay_middle)
  wMainForm.VBObject("tbToolBar").Window("ToolbarWindow32", "", 1).ClickItem(7)
  BuiltIn.Delay(delay_middle) 
  wMDIClient.VBObject("frmEditNavTreeNew").Close 
  Call ClickCmdButton(5, "²Ûá")
End Sub

'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'êïáõ·áõÙ ¿ áñ, ·³ÝÓÙ³Ý ï»ë³ÏÁ ×Çßï Éñ³óí³Í ÉÇÝÇ
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'docN - å³ÛÙ³Ý³·ñÇ Ñ³Ù³ñ
'tabN - Ã³µ-Ç Ñ³Ù³ñ
'Num - ¹³ßïÇ Ñ³Ù³ñ
'chargeType - ·³ÝÓÙ³Ý ï»ë³Ï
'chargePercent - ·³ÝÓÙ³Ý ïáÏáë
'chargeSum - ·³ÝÓÙ³Ý ·áõÙ³ñ
Sub Check_Charges(docN, tabN, Num, chargeType, chargePercent, chargeSum)
  Dim is_exists : is_exists = False
   
  If Trim(wMDIClient.vbObject("frmPttel").vbObject("tdbgView").Columns.Item(2).text) = docN Then
    BuiltIn.Delay(3000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_ToEdit)                   

    If Not Trim(wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame_" & tabN).VBObject("ASTypeTree_" & Num).VBObject("TDBMask").Text) = chargeType Then
      Log.Error("Type doesn't match" & wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame_" & tabN).VBObject("ASTypeTree_" & Num).VBObject("TDBMask").Text)
    End If
    If Not Trim(wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame_" & tabN).VBObject("TDBNumber_2").Text) = chargePercent Then 
      Log.Error("Type doesn't match" & wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame_" & tabN).VBObject("TDBNumber_2").Text)
    End If          
    If Not Trim(wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame_" & tabN).VBObject("TDBNumber_3").Text) = chargeSum Then 
      Log.Error("Type doesn't match" & wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame_" & tabN).VBObject("TDBNumber_3").Text)
    End If
  Else
    Log.Error("Document doesn't exists")
  End If
  Call ClickCmdButton(1, "Î³ï³ñ»É")
End Sub


'Վճարման հանձնարարագիր (ուղ.) փաստաթղթի Ընդհանուր էջի կլասս
Class PayOrderSentCommon
    Public isn
    Public div
    Public dep
    Public fDate
    Public note
    Public docN
    Public payClientCode
    Public rezident
    Public accD
    Public payer
    Public legalPos
    Public payerEng
    Public taxCode
    Public regNum
    Public accC
    Public areaCode
    Public reciever
    Public recieverEng
    Public sum
    Public cur
    Public aim
    Public tabN
    Public check
    
    Sub Class_Initialize
        isn = ""
        div = ""
        dep = ""
        fDate = "  /  /    "
        note = ""
        docN = ""
        payClientCode = ""
        rezident = ""
        accD = ""
        payer = ""
        legalPos = ""
        payerEng = ""
        taxCode = ""
        regNum = ""
        accC = ""
        areaCode = ""
        reciever = ""
        recieverEng = ""
        sum = "0.00"
        cur = ""
        aim = ""
        tabN = 1
        check = False   
    End Sub
End Class
 
Function New_PaymentOrderSentCommon()
    Set New_PaymentOrderSentCommon = new PayOrderSentCommon
End Function

'Վճարման հանձնարարագիր (ուղ.) փաստաթղթի Ընդհանուր էջի ստուգում
Sub Payment_Order_Sent_CommonTab_Check(PayOrdSentCom)
    Call GoTo_ChoosedTab(PayOrdSentCom.tabN)
    'Գրասենյակ դաշտի ստուգում
    Call Compare_Two_Values("Գրասենյակ",Get_Rekvizit_Value("Document",PayOrdSentCom.tabN,"Mask","ACSBRANCH"),PayOrdSentCom.div)
    'Բաժին դաշտի ստուգում
    Call Compare_Two_Values("Բաժին",Get_Rekvizit_Value("Document",PayOrdSentCom.tabN,"Mask","ACSDEPART"),PayOrdSentCom.dep)
    'Ամսաթիվ դաշտի ստուգում
    Call Compare_Two_Values("Ամսաթիվ",Get_Rekvizit_Value("Document",PayOrdSentCom.tabN,"General","DATE"),PayOrdSentCom.fDate)
    'Նշում դաշտի ստուգում
    Call Compare_Two_Values("Նշում",Get_Rekvizit_Value("Document",PayOrdSentCom.tabN,"Mask","PAYNOTE"),PayOrdSentCom.note)
    'Փաստաթղթի N դաշտի ստուգում
    Call Compare_Two_Values("Փաստաթղթի N",Get_Rekvizit_Value("Document",PayOrdSentCom.tabN,"General","DOCNUM"),PayOrdSentCom.docN)
    'Վճարող հաճախորդի կոդ դաշտի ստուգում
    Call Compare_Two_Values("Վճարող հաճախորդի կոդ",Get_Rekvizit_Value("Document",PayOrdSentCom.tabN,"Mask","CLICODE"),PayOrdSentCom.payClientCode)
    'Ռեզիդենտություն դաշտի ստուգում
    Call Compare_Two_Values("Ռեզիդենտություն",Get_Rekvizit_Value("Document",PayOrdSentCom.tabN,"Mask","RES"),PayOrdSentCom.rezident)
    'Հաշիվ Դեբետ դաշտի ստուգում
    Call Compare_Two_Values("Հաշիվ Դեբետ",Get_Rekvizit_Value("Document",PayOrdSentCom.tabN,"Bank","ACCDB"),PayOrdSentCom.accD)
    'Վճարող դաշտի ստուգում
    Call Compare_Two_Values("Վճարող",Get_Rekvizit_Value("Document",PayOrdSentCom.tabN,"Comment","PAYER"),PayOrdSentCom.payer)
    'Իրավաբանական կարգավիճակ դաշտի ստուգում
    Call Compare_Two_Values("Իրավաբանական կարգավիճակ",Get_Rekvizit_Value("Document",PayOrdSentCom.tabN,"Mask","JURSTAT"),PayOrdSentCom.legalPos)
    'Վճարող (անգ.) դաշտի ստուգում
    Call Compare_Two_Values("Վճարող (անգ.)",Get_Rekvizit_Value("Document",PayOrdSentCom.tabN,"Comment","EPAYER"),PayOrdSentCom.payerEng)
    'ՀՎՀՀ(վճարող) դաշտի ստուգում
    Call Compare_Two_Values("ՀՎՀՀ(վճարող)",Get_Rekvizit_Value("Document",PayOrdSentCom.tabN,"Comment","TAXCODSD"),PayOrdSentCom.taxCode)
    'Սոցիալական քարտ դաշտի ստուգում
    Call Compare_Two_Values("Սոցիալական քարտ",Get_Rekvizit_Value("Document",PayOrdSentCom.tabN,"General","REGNUM"),PayOrdSentCom.regNum)
    'Հաշիվ Կրեդիտ դաշտի ստուգում
    Call Compare_Two_Values("Հաշիվ կրեդիտ",Get_Rekvizit_Value("Document",PayOrdSentCom.tabN,"Bank","ACCCR"),PayOrdSentCom.accC)  
    'Ստացող դաշտի ստուգում
    Call Compare_Two_Values("Ստացող",Get_Rekvizit_Value("Document",PayOrdSentCom.tabN,"Comment","RECEIVER"),PayOrdSentCom.reciever)
    'Ստացող (անգ.) դաշտի ստուգում
    Call Compare_Two_Values("Ստացող (անգ.)",Get_Rekvizit_Value("Document",PayOrdSentCom.tabN,"Comment","ERECEIVER"),PayOrdSentCom.recieverEng)    
    'Գումար դաշտի ստուգում
    Call Compare_Two_Values("Գումար",Get_Rekvizit_Value("Document",PayOrdSentCom.tabN,"General","SUMMA"),PayOrdSentCom.sum)
    'Արժույթ դաշտի ստուգում
    Call Compare_Two_Values("Արժույթ",Get_Rekvizit_Value("Document",PayOrdSentCom.tabN,"Mask","CUR"),PayOrdSentCom.cur)
    'Նպատակ դաշտի ստուգում
    Call Compare_Two_Values("Նպատակ",Get_Rekvizit_Value("Document",PayOrdSentCom.tabN,"Comment","AIM"),PayOrdSentCom.aim)
End Sub

'Վճարման հանձնարարագիր (ուղ.) փաստաթղթի Լրացուցիչ էջի կլասս
Class PayOrderSentAdd
    Public payDate
    Public docN
    Public pack
    Public transitAcc
    Public correspondentAcc
    Public correspondentAccCB
    Public transitAcc2
    Public recPaySys
    Public sentPaySys
    Public onOrder
    Public transferAim
    Public refuse
    Public accType
    Public tabN
    Public check
    
    Private Sub Class_Initialize
        payDate = "  /  /  "
        docN = ""
        pack = ""
        transitAcc = ""
        correspondentAcc = ""
        correspondentAccCB = ""
        transitAcc2 = ""
        recPaySys = ""
        sentPaySys = ""
        onOrder = 0
        transferAim = ""
        refuse = ""
        accType = ""
        tabN = 2
        check = False         
    End Sub
End Class

Function New_PaymentOrderSentAdd()
    Set New_PaymentOrderSentAdd = new PayOrderSentAdd
End Function

'Վճարման հանձնարարագիր (ուղ.) փաստաթղթի Լրացուցիչ էջի ստուգում
Sub Payment_Order_Sent_AddTab_Check(PayOrdSentAdd)
    Call GoTo_ChoosedTab(PayOrdSentAdd.tabN)
    'Վճարման օր դաշտի ստուգում
    Call Compare_Two_Values("Վճարման օր",Get_Rekvizit_Value("Document",PayOrdSentAdd.tabN,"General","PAYDATE"),PayOrdSentAdd.payDate)
    'Փաստաթղթի N(20) դաշտի ստուգում
    Call Compare_Two_Values("Փաստաթղթի N(20) ",Get_Rekvizit_Value("Document",PayOrdSentAdd.tabN,"General","BMDOCNUM"),PayOrdSentAdd.docN)
    'Փաթեթի համարը դաշտի ստուգում
    Call Compare_Two_Values("Փաթեթի համարը",Get_Rekvizit_Value("Document",PayOrdSentAdd.tabN,"General","PACK"),PayOrdSentAdd.pack)
    'Տարանցիկ հաշիվ դաշտի ստուգում
    Call Compare_Two_Values("Տարանցիկ հաշիվ",Get_Rekvizit_Value("Document",PayOrdSentAdd.tabN,"Mask","TCORRACC"),PayOrdSentAdd.transitAcc)
    'Թղթակցային հաշիվ դաշտի ստուգում
    Call Compare_Two_Values("Թղթակցային հաշիվ",Get_Rekvizit_Value("Document",PayOrdSentAdd.tabN,"Mask","CORRACC"),PayOrdSentAdd.correspondentAcc)
    'Թղթակցային հաշիվ ԿԲ-ում դաշտի ստուգում
    Call Compare_Two_Values("Թղթակցային հաշիվ ԿԲ-ում",Get_Rekvizit_Value("Document",PayOrdSentAdd.tabN,"Bank","CORRACCCB"),PayOrdSentAdd.correspondentAccCB)
    'Տարանցիկ հաշիվ 2 դաշտի ստուգում
    Call Compare_Two_Values("Տարանցիկ հաշիվ 2",Get_Rekvizit_Value("Document",PayOrdSentAdd.tabN,"Mask","TCORRACCT"),PayOrdSentAdd.transitAcc2)
    'Ընդ. վճ. համակարգ դաշտի ստուգում
    Call Compare_Two_Values("Ընդ. վճ. համակարգ",Get_Rekvizit_Value("Document",PayOrdSentAdd.tabN,"Mask","PAYSYSIN"),PayOrdSentAdd.recPaySys)
    'Ուղ. վճ համակարգ դաշտի ստուգում
    Call Compare_Two_Values("Ուղ. վճ համակարգ",Get_Rekvizit_Value("Document",PayOrdSentAdd.tabN,"Mask","PAYSYSOUT"),PayOrdSentAdd.sentPaySys)
    'Համաձայն հրահանգի դաշտի ստուգում
    Call Compare_Two_Values("Համաձայն հրահանգի",Get_Rekvizit_Value("Document",PayOrdSentAdd.tabN,"CheckBox","ONORDER"),PayOrdSentAdd.onOrder)
    'Փոխանցման նպատակ դաշտի ստուգում
    Call Compare_Two_Values("Փոխանցման նպատակ",Get_Rekvizit_Value("Document",PayOrdSentAdd.tabN,"Mask","PAYAIM"),PayOrdSentAdd.transferAim)
    'Մերժում դաշտի ստուգում
    Call Compare_Two_Values("Մերժում",Get_Rekvizit_Value("Document",PayOrdSentAdd.tabN,"General","REFUSE"),PayOrdSentAdd.refuse)
    'Հաշվի տիպ դաշտի ստուգում
    Call Compare_Two_Values("Հաշվի տիպ",Get_Rekvizit_Value("Document",PayOrdSentAdd.tabN,"Mask","ACCTYPE"),PayOrdSentAdd.accType)    
End Sub

'Վճարման հանձնարարագիր (ուղ.) փաստաթղթի Գանձում փոխանցումից էջի կլասս
Class PaymentOrderSentTransCharge
    Public chargesAcc
    Public cur
    Public CBCourse
    Public chargeType1
    Public interest1
    Public sum1
    Public incomeAcc1
    Public chargeType2
    Public interest2
    Public sum2
    Public incomeAcc2 
    Public purSale
    Public optype
    Public opPlace
    Public time
    Public busField
    Public comment
    Public tabN
    Public check
    Private Sub Class_Initialize 
        chargesAcc = ""
        cur = ""
        CBCourse = "0/0"
        chargeType1 = ""
        interest1 = "0.0000"
        sum1 = "0.00"
        incomeAcc1 = ""
        chargeType2 = ""
        interest2 = "0.0000"
        sum2 = "0.00"
        incomeAcc2 = "" 
        purSale = ""
        optype = ""
        opPlace = ""
        If aqDateTime.Compare(aqConvert.DateTimeToFormatStr(aqDateTime.Time, "%H:%M"), "16:00") < 0 Then
            time = "1"
        Else
            time = "2"
        End If 
        busField = ""
        comment = ""
        tabN = 3
        check = False
    End Sub 
End Class

Function New_PaymentOrderSentTransCharge()
    Set New_PaymentOrderSentTransCharge = new PaymentOrderSentTransCharge
End Function
'Վճարման հանձնարարագիր (ուղ.) փաստաթղթի Գանձում փոխանցումից էջի ստուգում
Sub Payment_Order_Sent_TChargeTab_Check(PayOrdSentTCharge)
    Call GoTo_ChoosedTab(PayOrdSentTCharge.tabN)
    'Գանձման հաշիվ դաշտի ստուգում
    Call Compare_Two_Values("Գանձման հաշիվ",Get_Rekvizit_Value("Document",PayOrdSentTCharge.tabN,"Mask","CHRGACC"),PayOrdSentTCharge.chargesAcc)
    'Արժույթ դաշտի ստուգում
    Call Compare_Two_Values("Արժույթ",Get_Rekvizit_Value("Document",PayOrdSentTCharge.tabN,"Mask","CHRGCUR"),PayOrdSentTCharge.cur)
    'ԿԲ փոխարժեք դաշտի ստուգում
    Call Compare_Two_Values("ԿԲ փոխարժեք",Get_Rekvizit_Value("Document",PayOrdSentTCharge.tabN,"Course","CHRGCBCRS"),PayOrdSentTCharge.CBCourse)
    'Գանձման տեսակ դաշտի ստուգում
    Call Compare_Two_Values("Գանձման տեսակ",Get_Rekvizit_Value("Document",PayOrdSentTCharge.tabN,"Mask","PAYSCALE"),PayOrdSentTCharge.chargeType1)
    'Տոկոս դաշտի ստուգում
    Call Compare_Two_Values("Տոկոս",Get_Rekvizit_Value("Document",PayOrdSentTCharge.tabN,"General","PRSNT"),PayOrdSentTCharge.interest1)
    'Գումար դաշտի ստուգում
    Call Compare_Two_Values("Գումար",Get_Rekvizit_Value("Document",PayOrdSentTCharge.tabN,"General","CHRGSUM"),PayOrdSentTCharge.sum1)
    'Եկամտի հաշիվ դաշտի ստուգում
    Call Compare_Two_Values("Եկամտի հաշիվ",Get_Rekvizit_Value("Document",PayOrdSentTCharge.tabN,"Mask","CHRGINC"),PayOrdSentTCharge.incomeAcc1)
    'Գանձման տեսակ 2 դաշտի ստուգում
    Call Compare_Two_Values("Գանձման տեսակ 2",Get_Rekvizit_Value("Document",PayOrdSentTCharge.tabN,"Mask","PAYSCALE2"),PayOrdSentTCharge.chargeType2)
    'Տոկոս 2 դաշտի ստուգում
    Call Compare_Two_Values("Տոկոս 2",Get_Rekvizit_Value("Document",PayOrdSentTCharge.tabN,"General","PRSNT2"),PayOrdSentTCharge.interest2)
    'Գումար 2 դաշտի ստուգում
    Call Compare_Two_Values("Գումար 2",Get_Rekvizit_Value("Document",PayOrdSentTCharge.tabN,"General","CHRGSUM2"),PayOrdSentTCharge.sum2)
    'Եկամտի հաշիվ 2 դաշտի ստուգում
    Call Compare_Two_Values("Եկամտի հաշիվ 2",Get_Rekvizit_Value("Document",PayOrdSentTCharge.tabN,"Mask","CHRGINC2"),PayOrdSentTCharge.incomeAcc2)
    'Առք/Վաճառք դաշտի ստուգում
    Call Compare_Two_Values("Առք/Վաճառք",Get_Rekvizit_Value("Document",PayOrdSentTCharge.tabN,"Mask","CUPUSA"),PayOrdSentTCharge.purSale)
    'Գործողության տեսակ դաշտի ստուգում
    Call Compare_Two_Values("Գործողության տեսակ",Get_Rekvizit_Value("Document",PayOrdSentTCharge.tabN,"Mask","CURTES"),PayOrdSentTCharge.optype)
    'Ժամանակ դաշտի ստուգում
    Call Compare_Two_Values("Ժամանակ",Get_Rekvizit_Value("Document",PayOrdSentTCharge.tabN,"Mask","TIME"),PayOrdSentTCharge.time)
    'Գոծողության վայր դաշտի ստուգում
    Call Compare_Two_Values("Գոծողության վայր",Get_Rekvizit_Value("Document",PayOrdSentTCharge.tabN,"Mask","CURVAIR"),PayOrdSentTCharge.opPlace)
    'Գործունեության ոլորտ դաշտի ստուգում
    Call Compare_Two_Values("Գործունեության ոլորտ",Get_Rekvizit_Value("Document",PayOrdSentTCharge.tabN,"Mask","VOLORT"),PayOrdSentTCharge.busField)
    'Մեկնաբանություն դաշտի ստուգում
    Call Compare_Two_Values("Մեկնաբանություն",Get_Rekvizit_Value("Document",PayOrdSentTCharge.tabN,"General","COMM"),PayOrdSentTCharge.comment)
End Sub

'Վճարման հանձնարարագիր (ուղ.) փաստաթղթի Դրամարկղ էջի կլասս
Class PaymentOrderSentCashDesk
    Public cashDesk
    Public cashLabel
    Public depositor
    Public idNum
    Public passType
    Public address
    Public eMail
    Public base
    Public transInput
    Public transCurr
    Public chargeInput
    Public chargeCurr
    Public dateOfBirth
    Public birthPlace
    Public stateRegCertNum
    Public tabN
    Public check 
    Private Sub Class_Initialize
        cashDesk = ""
        cashLabel = ""
        depositor = ""
        idNum = ""
        passType = ""
        address = ""
        eMail = ""
        base = ""
        transInput = "0.00"
        transCurr = ""
        chargeInput = "0.00"
        chargeCurr = ""
        dateOfBirth = "  /  /  "
        birthPlace = ""
        stateRegCertNum = ""
        tabN = 4
        check = False     
    End Sub
End Class

Function New_PaymentOrderSentCashDesk()
    Set New_PaymentOrderSentCashDesk = new PaymentOrderSentCashDesk
End Function
'Վճարման հանձնարարագիր (ուղ.) փաստաթղթի Դրամարկղ էջի ստուգում
Sub Foreign_Payment_Order_Sent_CDesk_Check(PayOrdSentCashDesk)
    Call GoTo_ChoosedTab(PayOrdSentCashDesk.tabN)
    'Դրամարկղ դաշտի ստուգում
    Call Compare_Two_Values("Դրամարկղ",Get_Rekvizit_Value("Document",PayOrdSentCashDesk.tabN,"Mask","KASSA"),PayOrdSentCashDesk.cashDesk)
    'Դրամարկղի նիշ դաշտի ստուգում
    Call Compare_Two_Values("Դրամարկղի նիշ",Get_Rekvizit_Value("Document",PayOrdSentCashDesk.tabN,"Mask","KASSIMV"),PayOrdSentCashDesk.cashLabel)
    'Մուծող դաշտի ստուգում
    Call Compare_Two_Values("Մուծող",Get_Rekvizit_Value("Document",PayOrdSentCashDesk.tabN,"Comment","PAYER1"),PayOrdSentCashDesk.depositor)
    'Անձը հաստատող փաստ. դաշտի ստուգում
    Call Compare_Two_Values("Անձը հաստատող փաստ.",Get_Rekvizit_Value("Document",PayOrdSentCashDesk.tabN,"General","PASSNUM"),PayOrdSentCashDesk.idNum)
    'Անձը հաստատող փաստ. տեսակ դաշտի ստուգում
    Call Compare_Two_Values("Անձը հաստատող փաստ. տեսակ",Get_Rekvizit_Value("Document",PayOrdSentCashDesk.tabN,"Mask","PASSTYPE"),PayOrdSentCashDesk.passType)
    'Հասցե դաշտի ստուգում
    Call Compare_Two_Values("Հասցե",Get_Rekvizit_Value("Document",PayOrdSentCashDesk.tabN,"General","ADDRESS"),PayOrdSentCashDesk.address)
    'Էլ. հասցե դաշտի ստուգում
    Call Compare_Two_Values("Էլ. հասցե",Get_Rekvizit_Value("Document",PayOrdSentCashDesk.tabN,"General","EMAIL"),PayOrdSentCashDesk.eMail)
    'Հիմք դաշտի ստուգում
    Call Compare_Two_Values("Հիմք",Get_Rekvizit_Value("Document",PayOrdSentCashDesk.tabN,"Comment","BASE"),PayOrdSentCashDesk.base)
    'Մուտք փոխանցման համար դաշտի ստուգում
    Call Compare_Two_Values("Մուտք փոխանցման համար",Get_Rekvizit_Value("Document",PayOrdSentCashDesk.tabN,"General","INTRANSSUM"),PayOrdSentCashDesk.transInput)
    'Արժույթ դաշտի ստուգում
    Call Compare_Two_Values("Արժույթ",Get_Rekvizit_Value("Document",PayOrdSentCashDesk.tabN,"Mask","INTRANSCUR"),PayOrdSentCashDesk.transCurr)
    'Մուտք գանձման համար դաշտի ստուգում
    Call Compare_Two_Values("Մուտք գանձման համար",Get_Rekvizit_Value("Document",PayOrdSentCashDesk.tabN,"General","INCHRGSUM"),PayOrdSentCashDesk.chargeInput)
    'Արժույթ 2 դաշտի ստուգում
    Call Compare_Two_Values("Արժույթ 2",Get_Rekvizit_Value("Document",PayOrdSentCashDesk.tabN,"Mask","INCHRGCUR"),PayOrdSentCashDesk.chargeCurr)
    'Ծննդյան ամսաթիվ դաշտի ստուգում
    Call Compare_Two_Values("Ծննդյան ամսաթիվ",Get_Rekvizit_Value("Document",PayOrdSentCashDesk.tabN,"General","DATEBIRTH"),PayOrdSentCashDesk.dateOfBirth)
    'Ծննդավայր դաշտի ստուգում
    Call Compare_Two_Values("Ծննդավայր",Get_Rekvizit_Value("Document",PayOrdSentCashDesk.tabN,"Comment","BIRTHPLACE"),PayOrdSentCashDesk.birthPlace)
    'Պետ. գրանցման վկայականի համար դաշտի ստուգում
    Call Compare_Two_Values("Պետ. գրանցման վկայականի համար",Get_Rekvizit_Value("Document",PayOrdSentCashDesk.tabN,"General","REGCERT"),PayOrdSentCashDesk.stateRegCertNum)
End Sub

'Վճարման հանձնարարագիր (ուղ.) փաստաթղթի Վճարման տվյալներ էջի կլասս
Class PaymentOrderSentPayData
    Public payCode
    Public reportCode
    Public report
    Public msgCode(6)
    Public msg(6)
    Public rezident
    Public legalPos
    Public areaCode
    Public taxCode
    Public regNum
    Public name
    Public idNum
    Public passType
    Public address
    Public authPerson
    Public addInfo
    Public tabN
    Public check 
    Private Sub Class_Initialize
        Dim i
        payCode = ""
        reportCode = ""
        report = ""
        For i = 0 to 6
            msgCode(i) = ""
            msg(i) = ""
        Next    
        rezident = ""
        legalPos = ""
        areaCode = ""
        taxCode = ""
        regNum = ""
        name = ""
        idNum = ""
        passType = ""
        address = ""
        authPerson = ""
        addInfo = ""
        tabN = 5
        check = False     
    End Sub
End Class

Function New_PaymentOrderSentPayData()
    Set New_PaymentOrderSentPayData = New PaymentOrderSentPayData
End Function

'Վճարման հանձնարարագիր (ուղ.) փաստաթղթի Վճարման տվյալներ էջի ստուգում
Sub Payment_Order_Sent_PayData_Check(PayOrdSentPayData)
    Dim i
    Call GoTo_ChoosedTab(PayOrdSentPayData.tabN)
    'Հանձնարարագրի կոդ դաշտի ստուգում
    Call Compare_Two_Values("Հանձնարարագրի կոդ",Get_Rekvizit_Value("Document",PayOrdSentPayData.tabN,"Mask","PAYMENTCODE"),PayOrdSentPayData.payCode)
    'Հաշվետվողականության կոդ դաշտի ստուգում
    Call Compare_Two_Values("Հաշվետվողականության կոդ",Get_Rekvizit_Value("Document",PayOrdSentPayData.tabN,"Mask","REPORTCODE"),PayOrdSentPayData.reportCode)
    'Հաշվետվողականության կոդ 2 դաշտի ստուգում
    Call Compare_Two_Values("Հաշվետվողականության կոդ 2",Get_Rekvizit_Value("Document",PayOrdSentPayData.tabN,"Comment","REPORT"),PayOrdSentPayData.report)
    'Հաղորդագրություններ դաշտերի ստուգում
    For i = 1 to 6
        Call Compare_Two_Values("Հաղորդագրության կոդ "&i ,Get_Rekvizit_Value("Document",PayOrdSentPayData.tabN,"Mask","MSGCODE"&i),PayOrdSentPayData.msgCode(i))
        Call Compare_Two_Values("Հաղորդագրություն "&i ,Get_Rekvizit_Value("Document",PayOrdSentPayData.tabN,"Comment","MESSAGE"&i),PayOrdSentPayData.msg(i))
    Next
    'Ռեզիդենտություն դաշտի ստուգում
    Call Compare_Two_Values("Ռեզիդենտություն",Get_Rekvizit_Value("Document",PayOrdSentPayData.tabN,"Mask","DBRES"),PayOrdSentPayData.rezident)
    'Իրավաբանական կարգավիճակ դաշտի ստուգում
    Call Compare_Two_Values("Իրավաբանական կարգավիճակ",Get_Rekvizit_Value("Document",PayOrdSentPayData.tabN,"Mask","DBJURSTAT"),PayOrdSentPayData.legalPos)
    'Տարած. կոդ դաշտի ստուգում
    Call Compare_Two_Values("Տարած. կոդ",Get_Rekvizit_Value("Document",PayOrdSentPayData.tabN,"Mask","DBAREA"),PayOrdSentPayData.areaCode)
    'ՀՎՀՀ դաշտի ստուգում
    Call Compare_Two_Values("ՀՎՀՀ",Get_Rekvizit_Value("Document",PayOrdSentPayData.tabN,"Comment","DBTAXCOD"),PayOrdSentPayData.taxCode)
    'Սոց. քարտ դաշտի ստուգում
    Call Compare_Two_Values("Սոց. քարտ",Get_Rekvizit_Value("Document",PayOrdSentPayData.tabN,"General","DBREGNUM"),PayOrdSentPayData.regNum)
    'Անվանում դաշտի ստուգում
    Call Compare_Two_Values("Անվանում",Get_Rekvizit_Value("Document",PayOrdSentPayData.tabN,"Comment","DEBTOR"),PayOrdSentPayData.name)
    'Անձնագիր դաշտի ստուգում
    Call Compare_Two_Values("Անձնագիր",Get_Rekvizit_Value("Document",PayOrdSentPayData.tabN,"General","DBPASS"),PayOrdSentPayData.idNum)
    'Անձնագրի տիպ դաշտի ստուգում
    Call Compare_Two_Values("Անձնագրի տիպ",Get_Rekvizit_Value("Document",PayOrdSentPayData.tabN,"Mask","DBPASSTYPE"),PayOrdSentPayData.passType)
    'Հասցե դաշտի ստուգում
    Call Compare_Two_Values("Հասցե",Get_Rekvizit_Value("Document",PayOrdSentPayData.tabN,"General","DBADDRESS"),PayOrdSentPayData.address)
    'Լիազորված անձ դաշտի ստուգում
    Call Compare_Two_Values("Լիազորված անձ",Get_Rekvizit_Value("Document",PayOrdSentPayData.tabN,"General","DBMNG"),PayOrdSentPayData.authPerson)
    'Լրացուցիչ ինֆորմացիա դաշտի ստուգում
    Call Compare_Two_Values("Լրացուցիչ ինֆորմացիա",Get_Rekvizit_Value("Document",PayOrdSentPayData.tabN,"General","DBINFO"),PayOrdSentPayData.addInfo)
End Sub   


Class PaymentOrderSent 
    Public commonTab
    Public addTab
    Public tChargeTab
    Public cDeskTab
    Public payDataTab
    Public attachTab
    Private Sub Class_Initialize
        Set commonTab = New_PaymentOrderSentCommon()
        Set addTab = New_PaymentOrderSentAdd()
        Set tChargeTab = New_PaymentOrderSentTransCharge()
        Set cDeskTab = New_PaymentOrderSentCashDesk()
        Set payDataTab = New_PaymentOrderSentPayData()
        Set attachTab = New_Attached_Tab(fCount, lCount, dCount)
        attachTab.tabN = 6
    End Sub
End Class

Function New_PaymentOrderSent(fCount ,lCount , dCount)
    Set New_PaymentOrderSent = New PaymentOrderSent
End Function

'Վճարման հանձնարարագիր (ուղ.) փաստաթղթի ստուգում
Sub Payment_Order_Sent_Check(PayOrdSent)
    'Փաստաթղթի isn-ի ստացում
    PayOrdSent.commonTab.isn = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
    'Ընդհանուր
    If PayOrdSent.commonTab.check Then
        Call Payment_Order_Sent_CommonTab_Check(PayOrdSent.commonTab)
    End If
    'Լրացուցիչ
    If PayOrdSent.addTab.check Then
        Call Payment_Order_Sent_AddTab_Check(PayOrdSent.addTab)
    End If
    'Գանձում փոխանցումից
    If PayOrdSent.tChargeTab.check Then
        Call Payment_Order_Sent_TChargeTab_Check(PayOrdSent.tChargeTab)
    End If    
    'Դրամարկղ
    If PayOrdSent.cDeskTab.check Then
        Call Foreign_Payment_Order_Sent_CDesk_Check(PayOrdSent.cDeskTab)
    End If 
    'Վճարման տվյալներ
    If PayOrdSent.payDataTab.check Then
        Call Payment_Order_Sent_PayData_Check(PayOrdSent.payDataTab)
    End If
    'Կցված
    Call Attach_Tab_Check(PayOrdSent.attachTab)     
End Sub