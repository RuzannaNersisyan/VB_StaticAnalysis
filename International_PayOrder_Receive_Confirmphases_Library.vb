Option Explicit
'USEUNIT Library_Common
'USEUNIT Online_PaySys_Library
'USEUNIT Constants
'USEUNIT Library_Contracts

'----------------------------------------------------------------------
'Միջազգային վճարման հանձնարարագիր (ուղ.) տեսակի վճարային փաստաթղթի լրացում
'----------------------------------------------------------------------
'department - Բաժին դաշտի արժեք
'docNumber - Փաստաթղթի համար դաշտի դաշտի
'data - Ամսաթիվ դաշտի արժեք
'payerAcc - վճարողի հաշիվ դաշտի արժեք
'payer - Վճարող դաշտի արժեք
'IBAN - true արժեքի դեպքում սեղմվում է IBAN կոճակը, fasle - ի դեպքում` ոչ
'country - Հաշիվ ֆիլտրի(IBAN կոճակը սեղմելիս) Երկիր դաշտի արժեք
'acc - Հաշիվ ֆիլտրի(IBAN կոճակը սեղմելիս) Հաշիվ դաշտի արժեք
'recAcc - Ստացողի հաշիվ դաշտի արժեք
'receiver - Ստացող դաշտի արժեք
'summa - Գումար դաշտի արժեք
'curr - Արժույթ դաշտի արժեք
'recCorrBank - Վճարող բանկի թղթակցի հաշիվ
'transAcc - Տարանցիկ հաշիվ
'fISN - Փաստաթղթի ISN
Sub  International_PayOrder_Recipient_Fill(fISN, office, department, docNumber, data, recAcc, receiver, recInfo, payerAcc, payer, payerAddr, country, acc,_
                                                                                   summa, curr, aim, recCorrBank, transAcc, IBAN )
    
  Dim rekvName
    
  If wMDIClient.WaitVBObject("frmASDocForm", 6000).Exists Then 
    Call GoTo_ChoosedTab(1)
    'Ստեղծվող ISN - ի փաստաթղթի վերագրում փոփոխականին 
    fISN = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
    'Գրասենյակ դաշտի լրացում
    Call Rekvizit_Fill("Document", 1, "General", "ACSBRANCH", office)
    'Բաժին դաշտի լրացում
    Call Rekvizit_Fill("Document", 1, "General", "ACSDEPART", department)
    
    'Փաստաթղթի N դաշտի արժեքի վերագրում փոփոխականին
    rekvName = GetVBObject("DOCNUM", wMDIClient.vbObject("frmASDocForm"))
    docNumber = wMDIClient.vbObject("frmASDocForm").vbObject("TabFrame").vbObject(rekvName).Text

    ' Ամսաթիվ դաշտի լրացում
    Call Rekvizit_Fill("Document", 1, "General", "DATE", data)
    'Ստացողի հաշիվ (59) դաշտի լրացում
    Call Rekvizit_Fill("Document", 1, "General", "ACCCR", "^A[Del]" & recAcc)
    'Ստացող դաշտի լրացում
    Call Rekvizit_Fill("Document", 1, "General", "RECEIVER", receiver)
    'Ստացողի հասցե դաշտի լրացում
    Call Rekvizit_Fill("Document", 1, "General", "RECADDR", recInfo)
    
    'Վճարողի հաշիվ դաշտի լրացում
    '    Call Rekvizit_Fill("Document", 1, "General", "ACCDB", "^A[Del]" & payerAcc)
    'Վճարող դաշտի լրացում
    Call Rekvizit_Fill("Document", 1, "General", "PAYER", payer)
    'Վճարողի հասցե դաշտի լրացում
    Call Rekvizit_Fill("Document", 1, "General", "PAYADDR", payerAddr)
    
    If IBAN Then
      'IBAN կոճակի սեղմում
      wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame").VBObject("AsTpComment_3").VBObject("CmdAdditional_2").Click()
      ' Երկիր դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "COUNTRY", country)
      'Հաշիվ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "ACC", acc)
      'Կարատել կոճակի սեղմում
      Sys.Process("Asbank").VBObject("frmAsUstPar").VBObject("CmdOK").Click()
    Else
      'Վճարողի հաշիվ դաշտի լրացում
      Call Rekvizit_Fill("Document", 1, "General", "ACCDB", "^A[Del]" & payerAcc)
    End If
    
    'Գումար դաշտի լրացում
    Call Rekvizit_Fill("Document", 1, "General", "SUMMA", summa)
    'Արժույթ դաշտի լրացում
    Call Rekvizit_Fill("Document", 1, "General", "CUR", curr)
    'Նպատակ դաշտի լրացում
    Call Rekvizit_Fill("Document", 1, "General", "AIM", aim) 

    ' Անցում 3.Ֆին. կազմակերպություն էջին
    Call GoTo_ChoosedTab(3)
    
    'Ստացող բանկի թղթակից դաշտի լրացում
    Call Rekvizit_Fill("Document", 3, "General", "RCORBANK", recCorrBank)

    'Անցում 2.Լրացուցիչ էջին
    Call GoTo_ChoosedTab(2)
    
    'Տարանցիկ հաշիվ դաշտի լրացում
    Call Rekvizit_Fill("Document", 2, "General", "TCORRACC", transAcc)
    
    'Կատարել կոճակի սեղմում
    Call ClickCmdButton(1, "Î³ï³ñ»É")
  Else 
    Log.Error "Can't find frmASDocForm window", "", pmNormal, ErrorColor
  End If
End Sub

'----------------------------------------------------------------------
'ØÇç³½·³ÛÇÝ í×³ñÙ³Ý Ñ³ÝÓÝ³ñ³ñ·ñÇ áõÕáñÏáõÙ Ñ³ëï³ïÙ³Ý ë¨ óáõó³Ï
'----------------------------------------------------------------------
Sub International_PayOrder_Send_To_BlackList()
    BuiltIn.Delay(3000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_SendtoVerBL)
    If p1.WaitvbObject("frmAsMsgBox", 2000).Exists Then
    Call ClickCmdButton(5, "²Ûá")
    Call ClickCmdButton(5, "²Ûá")
    End if
    BuiltIn.Delay(1000)
    wMDIClient.vbObject("frmPttel").Close()
End Sub