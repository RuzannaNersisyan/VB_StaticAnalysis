'USEUNIT Library_Common       
'USEUNIT Payment_Order_ConfirmPhases_Library  
'USEUNIT BankMail_Library  
'USEUNIT Constants
 
'Ñ³ßÇíÁ փակփլ/բացել
public sub AccCloseOrOpen(acc, action) 
  Call ChangeWorkspace(c_ChiefAcc)
  ' ÐÇÙÝ³Ï³Ý áõÕáñ¹Çã Í³éáí ³ÝóáõÙ ÃÕÃ³å³Ý³Ï
  wMDIClient.VBObject("frmExplorer").VBObject("tvTreeView").DblClickItem("|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|Ð³ßÇíÝ»ñ")
  
  ' Հաշվի շաբլոն դաշտի լրացում
  Call Rekvizit_Fill("Dialog", 1, "General", "ACCMASK", acc)
  Call ClickCmdButton(2, "Î³ï³ñ»É")
    
  if action = 0 Then
      if wMDIClient.VBObject("frmPttel").VBObject("tdbgView").Columns.Item(7).value = "" then
        BuiltIn.Delay(3000)
        'Ð³Ùå³ï³ëË³Ý ïáÕÇ íñ³ ë»ÕÙ»É §ö³Ï»É¦
        Call wMainForm.MainMenu.Click(c_AllActions)
        Call wMainForm.PopupMenu.Click(c_Close)
        BuiltIn.Delay(1000)
        Call ClickCmdButton(2, "Î³ï³ñ»É")
      Else
        Log.Message("Հաշիվն արդեն փակ է")
      end if
  else
    if wMDIClient.VBObject("frmPttel").VBObject("tdbgView").Columns.Item(7).value <> "" then
        BuiltIn.Delay(3000)
        'Ð³Ùå³ï³ëË³Ý ïáÕÇ íñ³ ë»ÕÙ»É §Բացել¦
        Call wMainForm.MainMenu.Click(c_AllActions)
        Call wMainForm.PopupMenu.Click(c_Open)
        BuiltIn.Delay(1000)
        Call ClickCmdButton(5, "²Ûá")
    Else
      Log.Message("Հաշիվն արդեն բաց է")
    end if
  end if  
end sub
 
 ' Հաղորդագրությունների ավտոմատ մշակում
Sub AutoMessageProcessing(clCount, delayTime)
 ' Մուտք Հաղորդագրություննների ավտոմատ մշակում      
      Call wTreeView.DblClickItem("|Ð»é³Ñ³ñ Ñ³Ù³Ï³ñ·»ñ|²íïáÙ³ï Ï³ï³ñíáÕ ·áñÍáÕáõÃÛáõÝÝ»ñ|Ð³Õáñ¹³·ñáõÃÛáõÝÝ»ñÇ ³íïáÙ³ï Ùß³ÏáõÙ")
      
      If  Not p1.WaitVBObject("frmCallBackOnTimer", 3000).Exists Then
          Log.Error("Հազորդագրությունների ավտոմատ մշակում պատուհանը չի բացվել")
          Exit Sub
      End If
      
      ' Կատարել կոճակի սեղմում
      p1.VBObject("frmCallBackOnTimer"& clCount).VBObject("cmdStop_Start").Click
      BuiltIn.Delay(delayTime)
      ' Դադարեցնել կոճակի սեղմում
      p1.VBObject("frmCallBackOnTimer"& clCount).VBObject("cmdCancel").Click
End Sub
 
' Պայմանագրի առկայության ստուգումը մշակման ենթակա մուտքային հաղորդագրություններ (Ընդհանուր) թղթապանակում
' todayDMY = Ամսաթիվ դաշտ
' system = Համակարգ դաշտ
' cliMask = Հաճախորդի կոդ
' amount = Պայմանագրի գումար չափ
' wState = Պայմանագրիը վիճակ
' msgType = Հաղորդագրության տեսակ
Function CheckContractRemoteSystems(direction, todayDMY, system, cliMask, msgType, amount, dirName, wState)
      Dim  wStatus : wStatus = False
      
      Call wTreeView.DblClickItem(direction)
      
      If Not p1.WaitVBObject("frmAsUstPar", 3000).Exists Then
          Log.Message(dirName &"դիալոգը չի բացվել")
          Exit Function
      End If
      
      ' Ժամանակահատվածի սկիզբ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "SDATE", todayDMY)
      ' Ժամանակահատվածի ավարտ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "EDATE", todayDMY)
      ' Համակարգ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "SYSTEM", system)
      ' Հաճախորդի կոդ  դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "CLIMASK", cliMask)
      ' Հաղորդագրության տեսակ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "MSGTYPE", msgType)
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
      
      If Not wMDIClient.WaitVBObject("frmPttel", 3000).Exists Then
            Log.Error(dirName &"թղթապանակը չի բացվել")
            CheckContractRemoteSystems = wStatus
            Exit Function
      ElseIf wMDIClient.VBObject("frmPttel").VBObject("tdbgView").ApproxCount <> 1 Then
            Log.Error("Պայմանագրիը առկա չէ " & dirName & "թղթապանակում")
            CheckContractRemoteSystems = wStatus
            Exit Function
      ElseIf Not CompareFieldValue("frmPttel", "SUMMA", amount) Then
            Log.Error("Պայմանագրի գումար դաշտի արժեքը չի համապատասխանում նախնական տրված Գումար դաշտի արժեքի հետ")
      ElseIf Not CompareFieldValue("frmPttel", "STATENAME", wState) Then
            Log.Error("Պայմանագրի վիճակ դաշտը պետք է լինի " & wState )
      End If
      
      wStatus = True
      
      CheckContractRemoteSystems = wStatus
End Function

' Պայամանագրի հաստատում/ մերժում
' colN - Հերթական սյան համարը
' docNum - Փաստաթղթի համարը
' action - Գործողության տեսակը
' basis - Հիմք դաշտի ծրագրային անվանումը
' refuse - Դաշտի լրացման արժեքը
' doNum - Բերված հաղորդագրության պատուհանի տեսակը
' doActio - Կոճակի անվանումը
Function ExcludeContractDoc(colN, docNum, action, basis, refuse, doNum, doActio)
  Dim status : status = False

  Do Until wMDIClient.VBObject("frmPttel").VBObject("tdbgView").EOF
    If  Trim(wMDIClient.VBObject("frmPttel").VBObject("tdbgView").Columns.Item(colN).Value) = Trim(docNum) Then
      ' Կատարել բոլոր գործողությունները
      Call wMainForm.MainMenu.Click(c_AllActions)
      ' Վավերացնել Վճարման հանձնարարագրի պայմանագիրն
      Call wMainForm.PopupMenu.Click(action)
      ' Հիմք դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", basis, refuse)
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(doNum, doActio)
      status = True
      Exit Do                   
    Else
      wMDIClient.VBObject("frmPttel").VBObject("tdbgView").MoveNext
    End If
  Loop 
      
  ExcludeContractDoc = status
End Function 

' Մերժել վճարման հանձնարարագիրը
' colN - Հերթական սյան համարը
' docNum - Փաստաթղթի համարը
' action - Գործողության տեսակը
' basis - Հիմք դաշտի ծրագրային անվանում
' refuse - Նպատակ դաշտի լրացման արժեքը
' doNum - Բերված հաղորդագրության պատուհանի տեսակը
' doActio - Կոճակի անվանումը
' fISN - Պայմանագրի ISN
Sub RejectPaymentOrder(colN, docNum, action, fISN, ordDocNum, basis, refuse, doNum, doActio)
  Do Until wMDIClient.VBObject("frmPttel").VBObject("tdbgView").EOF
     If  Trim(wMDIClient.VBObject("frmPttel").VBObject("tdbgView").Columns.Item(colN).Value) = Trim(docNum) Then
        BuiltIn.Delay(1000)
        ' Կատարել բոլոր գործողությունները
        Call wMainForm.MainMenu.Click(c_AllActions)
        ' Վավերացնել Վճարման հանձնարարագրի պայմանագիրն
        Call wMainForm.PopupMenu.Click(action)
                      
        If p1.WaitVBObject("frmAsUstPar", 2000).Exists Then
          ' Կատարել կոճակի սեղմում
          Call ClickCmdButton(2, "Î³ï³ñ»É")
        End If
                        
        If Not wMDIClient.WaitVBObject("frmASDocForm", 2000).Exists Then
              Log.Error("Հիշարար օրդեր փաստաթուղթը չի բացվել")
              Exit Sub
        End If
                        
        ' Փաստաթղթի ISN -ի ստացում
        fISN = wMDIClient.VBObject("frmASDocForm").DocFormCommon.Doc.ISN             
        ' Փաստաթղթի համարի ստացում
        ordDocNum = Get_Rekvizit_Value("Document", 1, "General", "DOCNUM")
        ' Նպատակ դաշտի լրացում
        Call Rekvizit_Fill("Document", 1, "General", basis, refuse)
        ' Կատարել կոճակի սեղմում
        Call ClickCmdButton(doNum, doActio)
        BuiltIn.Delay(1000)
        wMDIClient.VBObject("FrmSpr").Close
        Exit Do                   
     Else
        wMDIClient.VBObject("frmPttel").VBObject("tdbgView").MoveNext
     End If
  Loop 
End Sub 

' Հաշիվը սառեցնել/ ապասառեցնել
' accMask - Հաշվի շաբլոն
' frozen - Հաշվի վիճակը դարձնել դաշտ
' wAim - Նպատակ
Sub FreezeAccOrNo(accMask, frozen, wAim)

      ' Մուտք աշխատանքային փաստաթղթեր
      Call wTreeView.DblClickItem("|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|Ð³ßÇíÝ»ñ")
      
      If Not p1.WaitVBObject("frmAsUstPar", 3000).Exists Then
            Log.Error("Հաշիվներ դիալոգը չի բացվել")
            Exit Sub
      End If
      
      ' Հաշվի շաբլոն դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "ACCMASK", accMask)
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
          
      If Not wMDIClient.WaitVBObject("frmPttel", 3000).Exists Then
            Log.Error("Հաշիվներ թղթապանակը չի բացվել")
            Exit Sub
      ElseIf wMDIClient.VBObject("frmPttel").VBObject("tdbgView").ApproxCount <> 1 Then
            Log.Error(accMask &"Հաշվով միայն մեկ պայմանագիր պետք է լինի")
            Exit Sub
      End If
      
      BuiltIn.Delay(3000)
      ' Կատարել բոլոր գործողությունները
      Call wMainForm.MainMenu.Click(c_AllActions)
      ' Սառեցնել/ապասառեցնել գործողության կատարում
      Call wMainForm.PopupMenu.Click(c_FreezeUnfreeze)
      
      If Not p1.WaitVBObject("frmAsUstPar", 3000).Exists Then
            Log.Error("Սառեցնել/ապասառեցնել դիալոգը չի բացվել")
            Exit Sub
      End If
      
      ' "Հաշվի վիճակը դարձնել" դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "FROZEN", frozen)
      ' Նպատակ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "AIM", wAim)
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
      
      BuiltIn.Delay(2000)
      wMDIClient.VBObject("frmPttel").Close
End Sub

' Մուտք թղթապանակ
' directFolder  - թղթապանակ մուտք գործելու ճանապարհը
' folderName - թղթապանակի անվանումը
' wDayS - Ժամանակահատվածի սկիզբ
' wDayE - Ժամանակահատվածի ավարտ
' wCur - Արժույթ
' wAdmin - Կատարող 
' wDocType - Փաստաթղթի տեսակ
Function EnterFolder(directFolder, folderName, wDayS, wDayE, wCur, wUser, wDocType)
      Dim wState : wState = False
      
      Call wTreeView.DblClickItem(directFolder)
      BuiltIn.Delay(1000)
      
      If Not p1.WaitVBObject("frmAsUstPar", 3000).Exists Then
            Log.Error(folderName &" դիալոգը չի բացվել")
            Exit Function
      End If
      ' Ժամանակահատվածի սկիզբ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "PERN", wDayS)
      ' Ժամանակահատվածի ավարտ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "PERK", wDayE)
      ' Արժույթ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "CUR", wCur)
      ' Կատարողներ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "USER", wUser)
      ' Փաստաթղթի տեսակ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "DOCTYPE", wDocType)
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")

      If wMDIClient.WaitVBObject("frmPttel", 3000).Exists Then
            Log.Message(folderName &" թղթապանակը բացվել է")
            wState = True
      End If
      
      EnterFolder = wState
End Function


' Կանխիկացման հայտ փաստաթղթի փնտրում ISN-ով
Function SearchCashWithdrawalReqByISN(fISN)
  Dim status, i
  status = False
      
  For i = 0 To wMDIClient.VBObject("frmPttel").VBObject("tdbgView").ApproxCount - 1
    BuiltIn.Delay(1000) 
    ' Դիտել գործողության կատարում
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_View)          
    ' Ստուգում որ կանխիկացման հայտ փաստաթուղթը բացվել է
    If  Not wMDIClient.WaitVBObject("frmASDocForm", 3000).Exists Then
          Log.Error("Կանխիկացման հայտ փաստաթուղթը չի բացվել")
          SearchCashWithdrawalReqByISN = status
          Exit Function
    End If
    ' Ստուգում արդյոք հավասար են կանխիկացման հայտ փաստաթղթի ISN-ը SQL-ով սարքված հայտի ISN -ի հետ
    If Trim(wMDIClient.VBObject("frmASDocForm").DocFormCommon.Doc.ISN) = Trim(fISN) Then
          wMDIClient.VBObject("frmASDocForm").Close
          Exit For
    Else
          wMDIClient.VBObject("frmASDocForm").Close
    End If
    ' Անցնել հաջորդ տող
    wMDIClient.VBObject("frmPttel").VBObject("tdbgView").MoveNext  
  Next
  
  ' Ստուգում որ փաստաթուղթը գտնվել է
  If i = wMDIClient.VBObject("frmPttel").VBObject("tdbgView").ApproxCount then
      Log.Error("Փաստաթուղթը բացակայում է Աշխատանքային փաստաթղթեր թղթապանակից")
      SearchCashWithdrawalReqByISN = status
      Exit Function
  Else
      Log.Message("Փաստաթուղթը գտնվում է Աշխատանքային փաստաթղթեր թղթապանակում")                                                           
  End If
  status = True
      
  SearchCashWithdrawalReqByISN = status  
End Function

' Գործողություն Փաստաթղթի հետ
Sub ActionWithDoc(action, doNum, doActio)
  BuiltIn.Delay(3000)
  ' Բոլոր գործողություններ
  Call wMainForm.MainMenu.Click(c_AllActions)
  ' Գործողության անվանում
  Call wMainForm.PopupMenu.Click(action)
  ' Այո կոճակի սեղմում
  Call ClickCmdButton(doNum, doActio)
End Sub