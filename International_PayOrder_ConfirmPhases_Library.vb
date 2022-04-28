Option Explicit
'USEUNIT Library_Common
'USEUNIT Library_Contracts
'USEUNIT Online_PaySys_Library
'USEUNIT Constants
'USEUNIT Library_Colour
Dim chargesCount, fCount, lCount, dCount
'----------------------------------------------------------------------
'ØÇç³½·³ÛÇÝ í×³ñÙ³Ý Ñ³ÝÓÝ³ñ³ñ³·Çñ (áõÕ.) ï»ë³ÏÇ í×³ñ³ÛÇÝ ÷³ëï³ÃÕÃÇ Éñ³óáõÙ
'----------------------------------------------------------------------
'office - ¶ñ³ë»ÝÛ³Ï ¹³ßïÇ ³ñÅ»ù
'department - ´³ÅÇÝ ¹³ßïÇ ³ñÅ»ù
'docNumber - ö³ëï³ÃÕÃÇ Ñ³Ù³ñ ¹³ßïÇ ¹³ßïÇ
'data - ²Ùë³ÃÇí ¹³ßïÇ ³ñÅ»ù
'clTrans - Ð³×³Ëáñ¹Ç ÷áË³ÝóáõÙ ¹³ßïÇ ³ñÅ»ù
'res - è»½Ç¹»ÝïáõÃÛáõÝ ¹³ßïÇ ³ñÅ»ù
'payerInfo - ì×³ñáÕÇ ïí. ïÇå ¹³ßïÇ ³ñÅ»ù
'payerAcc - í×³ñáÕÇ Ñ³ßÇí ¹³ßïÇ ³ñÅ»ù
'payer - ì×³ñáÕ ¹³ßïÇ ³ñÅ»ù
'payerAddr - ì×³ñáÕÇ Ñ³ëó» ¹³ßïÇ ³ñÅ»ù
'recdataType - êï³óáÕÇ ïí. ïÇå ¹³ßïÇ ³ñÅ»ù
'IBAN - true ³ñÅ»ùÇ ¹»åùáõÙ ë»ÕÙíáõÙ ¿ IBAN Ïá×³ÏÁ, fasle - Ç ¹»åùáõÙ` áã
'country - Ð³ßÇí ýÇÉïñÇ(IBAN Ïá×³ÏÁ ë»ÕÙ»ÉÇë) ºñÏÇñ ¹³ßïÇ ³ñÅ»ù
'acc - Ð³ßÇí ýÇÉïñÇ(IBAN Ïá×³ÏÁ ë»ÕÙ»ÉÇë) Ð³ßÇí ¹³ßïÇ ³ñÅ»ù
'recAcc - êï³óáÕÇ Ñ³ßÇí ¹³ßïÇ ³ñÅ»ù
'receiver - êï³óáÕ ¹³ßïÇ ³ñÅ»ù
'recCountry - êï³óáÕÇ »ñÏÇñ ¹³ßïÇ ³ñÅ»ùÁ 
'recAddr - êï³óáÕÇ Ñ³ëó» ¹³ßïÇ ³ñÅ»ùÁ
'summa - ¶áõÙ³ñ ¹³ßïÇ ³ñÅ»ù
'curr - ²ñÅáõÛÃ ¹³ßïÇ ³ñÅ»ù
'paycorrBank - ì×³ñáÕÇ ÃÕÃ³ÏÇó µ³ÝÏ ¹³ßïÇ ³ñÅ»ù(3-ñ¹ ¿ç) 
'paycorrAcc - ì×³ñáÕ µ³ÝÏÇ ÃÕÃ³ÏóÇ Ñ³ßÇí
'medBankDataType - ØÇçÝáñ¹ µ³ÝÏÇ ïí, ïÇå ¹³ßïÇ ³ñÅ»ù
'medBank - ØÇçÝáñ¹ µ³ÝÏ ¹³ßïÇ ³ñÅ»ùÁ 
'medBankAcc - ØÇçÝáñ¹ µ³ÝÏÇ Ñ³ßÇí ¹³ßïÇ ³ñÅ»ù
'recOrgDataType - êï³óáÕ Ï³½Ù³Ï»ñåáõÃÛ³Ý ïí. ïÇå ¹³ßïÇ ³ñÅ»ùÁ 
'recOrg - êï³óáÕ Ï³½Ù³Ï»ñåáõÃÛáõÝ
'recOrgAcc - êï³óáÕÏ³½Ù³Ï»ñå. Ñ³ßÇí
'fISN - ö³ëï³ÃÕÃÇ ISN 
Sub International_PayOrder_Send_Fill(office, department, docNumber, data, clTrans, res, payerInfo, payerAcc, payer, payerAddr,_
                       recdataType, IBAN,country, acc, recAcc, receiver, recCountry, recAddr, summa, curr, paycorrBank , paycorrAcc, medBankDataType, _
                       medBank,medBankAcc, recOrgDataType, recOrg, recOrgAcc,aim, fISN)

  Dim rekvName
    
  BuiltIn.Delay(3000)
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click("Վճարման հանձնարարագրեր|Միջազգ. վճարման հանձնարարագիր (ուղ.)")
  
  If wMDIClient.vbObject("frmASDocForm").Exists Then 
    Call GoTo_ChoosedTab(1)
    'êï»ÕÍíáÕ ISN - Ç ÷³ëï³ïÃÕÃÇ  í»ñ³·ñáõÙ ÷á÷áË³Ï³ÝÇÝ
    fISN = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
    '¶ñ³ë»ÝÛ³Ï ¹³ßïÇ Éñ³óáõÙ
    Call Rekvizit_Fill("Document", 1, "General", "ACSBRANCH", office)
    '´³ÅÇÝ ¹³ßïÇ Éñ³óáõÙ
    Call Rekvizit_Fill("Document", 1, "General", "ACSDEPART", department)
    
    'ö³ëï³ÃÕÃÇ N ¹³ßïÇ ³ñÅ»ùÇ í»ñ³·ñáõÙ ÷á÷áË³Ï³ÝÇÝ
    docNumber = Get_Rekvizit_Value("Document" ,1, "General", "DOCNUM")
    '    rekvName = GetVBObject("DOCNUM", wMDIClient.vbObject("frmASDocForm"))
    '    docNumber = wMDIClient.vbObject("frmASDocForm").vbObject("TabFrame").vbObject(rekvName).Text
    
    ' ²Ùë³ÃÇí ¹³ßïÇ Éñ³óáõÙ
    Call Rekvizit_Fill("Document", 1, "General", "DATE", data)
    'Ð³×³Ëáñ¹Ç ÷áË³ÝóáõÙ ¹³ßïÇ Éñ³óáõÙ
    Call Rekvizit_Fill("Document", 1, "General", "CLITRANS", clTrans)
    'è»½Ç¹»ÝïáõÃÛáõÝ ¹³ßïÇ Éñ³óáõÙ
    Call Rekvizit_Fill("Document", 1, "General", "RES", res)
    'Ð³×³Ëáñ¹Ç ïí. ïÇå ¹³ßïÇ Éñ³óáõÙ
    Call Rekvizit_Fill("Document", 1, "General", "PAYOP", payerInfo)
    'ì×³ñáÕÇ Ñ³ßÇí ¹³ßïÇ Éñ³óáõÙ
    Call Rekvizit_Fill("Document", 1, "General", "ACCDB", "^A[Del]" & payerAcc)
    'ì×³ñáÕ ¹³ßïÇ Éñ³óáõÙ
    Call Rekvizit_Fill("Document", 1, "General", "PAYER", payer)
    'ì×³ñáÕÇ Ñ³ëó» ¹³ßïÇ Éñ³óáõÙ
    Call Rekvizit_Fill("Document", 1, "General", "PAYADDR", payerAddr)
    'êï³óáÕÇ ïí. ïÇå ¹³ßïÇ Éñ³óáõÙ
    Call Rekvizit_Fill("Document", 1, "General", "RECOP", recdataType)
    
    If IBAN Then
      'IBAN Ïá×³ÏÇ ë»ÕÙáõÙ
      wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame").VBObject("AsTpComment_3").VBObject("CmdAdditional_2").Click()
      ' ºñÏÇñ ¹³ßïÇ Éñ³óáõÙ
      Call Rekvizit_Fill("Dialog", 1, "General", "Country", country)
      'Ð³ßÇí ¹³ßïÇ Éñ³óáõÙ
      Call Rekvizit_Fill("Dialog", 1, "General", "Acc", acc)
      'Î³ï³ñ»É Ïá×³ÏÇ ë»ÕÙáõÙ 
      Call ClickCmdButton(2, "Î³ï³ñ»É")
    Else
      'êï³óáÕÇ Ñ³ßÇí ¹³ßïÇ Éñ³óáõÙ
      Call Rekvizit_Fill("Document", 1, "General", "ACCCR", recAcc)
    End If
    
    'êï³óáÕ ¹³ßïÇ Éñ³óáõÙ
    Call Rekvizit_Fill("Document", 1, "General", "RECEIVER", receiver)
    'êï³óáÕÇ »ñÏÇñ ¹³ßïÇ Éñ³óáõÙ
    Call Rekvizit_Fill("Document", 1, "General", "Country", recCountry)
    'êï³óáÕÇ Ñ³ëó» ¹³ßïÇ Éñ³óáõÙ
    Call Rekvizit_Fill("Document", 1, "General", "RECADDR", recAddr)
    '¶áõÙ³ñ ¹³ßïÇ Éñ³óáõÙ
    Call Rekvizit_Fill("Document", 1, "General", "SUMMA", summa)
    '²ñÅáõÛÃ ¹³ßïÇ Éñ³óáõÙ
    Call Rekvizit_Fill("Document", 1, "General", "CUR", curr)
    
    Call GoTo_ChoosedTab(2)
    'Üå³ï³Ï ¹³ßïÇ Éñ³óáõÙ
    Call Rekvizit_Fill("Document", 2, "General", "AIM", aim)
    
    '²ÝóáõÙ 3.üÇÝ. Ï³½Ù³Ï»ñåáõÃÛáõÝ ¿çÇÝ
    Call GoTo_ChoosedTab(3)
    'ì×³ñáÕ µ³ÝÏÇ ÃÕÃ³ÏÇó ¹³ßïÇ Éñ³óáõÙ
    Call Rekvizit_Fill("Document", 3, "General", "PCORBANK", paycorrBank)
    'ì×³ñáÕ µ³ÝÏÇ ÃÕÃ³ÏÇó  Ñ³ßÇí ¹³ßïÇ Éñ³óáõÙ
    Call Rekvizit_Fill("Document", 3, "General", "PCORID", paycorrAcc)
    'ØÇçÝáñ¹ µ³ÝÏÇ ïí. ïÇå ¹³ßïÇ Éñ³óáõÙ
    Call Rekvizit_Fill("Document", 3, "General", "MEDOP", medBankDataType)
    'ØÇçÝáñ¹ µ³ÝÏ ¹³ßïÇ Éñ³óáõÙ
    Call Rekvizit_Fill("Document", 3, "General", "MEDBANK", medBank)
    'ØÇçÝáñ¹ µ³ÝÏÇ Ñ³ßÇí ¹³ßïÇ Éñ³óáõÙ
    Call Rekvizit_Fill("Document", 3, "General", "MEDID", medBankAcc)
    'êï³óáÕ Ï³½Ù³Ï»ñå. ïí. ïÇå ¹³ßïÇ Éñ³óáõÙ
    Call Rekvizit_Fill("Document", 3, "General", "RINSTOP", recOrgDataType)
    'êï³óáÕ Ï³½Ù³Ï»ñå.  ¹³ßïÇ Éñ³óáõÙ
    Call Rekvizit_Fill("Document", 3, "General", "RECINST", recOrg)
    'êï³óáÕ Ï³½Ù³Ï»ñå. Ñ³ßÇí ¹³ßïÇ Éñ³óáõÙ
    Call Rekvizit_Fill("Document", 3, "General", "RINSTID", recOrgAcc)
    
    'Î³ï³ñ»É Ïá×³ÏÇ ë»ÕÙáõÙ
    Call ClickCmdButton(1, "Î³ï³ñ»É")
  Else 
    Log.Error "Can't find frmASDocForm window", "", pmNormal, ErrorColor
  End If
End Sub


'------------------------------------------------------------------------------------------
'ØÇç³½·³ÛÇÝ í×³ñÙ³Ý Ñ³ÝÓÝ³ñ³ñ³·Çñ (áõÕ.) ï»ë³ÏÇ í×³ñ³ÛÇÝ ÷³ëï³ÃÕÃÇ áõÕ³ñÏáõÙ SWIFT µ³ÅÇÝ
'------------------------------------------------------------------------------------------
Sub PaySys_Sento_SWIFT()
    BuiltIn.Delay(3000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    wMainForm.PopupMenu.click(c_SendToSW)
    If p1.WaitVBObject("frmAsMsgBox", 2000).Exists Then
        Call ClickCmdButton(5, "²Ûá")
        If p1.WaitVBObject("frmAsMsgBox", 2000).Exists Then
            If MessageExists(2,"¶áñÍáÕáõÃÛ³Ý µ³ñ»Ñ³çáÕ ³í³ñï") Then
              Call ClickCmdButton(5, "OK")
            End If
        End If
    Else
        Log.Message("Message box must be exist")
    End If
End Sub
'----------------------------------------------------------------------------------------
'SWIFT-Ç àõÕ³ñÏíáÕ ÷áË³ÝóáõÙÝ»ñ ÃÕÃ³å³Ý³ÏáõÙ í×³ñÙ³Ý Ñ³ÝÓÝ³ñ³ñ·ñÇ ³éÏ³ÛáõÃÛ³Ý ëïáõ·áõÙ:
'ºÃ» Ñ³ÝÓÝ³ñ³ñ·ÇñÁ ³éÏ³ ¿ , ³å³ ýáõÝÏóÇ³Ý í»ñ³¹³ñÓÝáõÙ ¿ true, fasle ` »Ã» ³ÛÝ µ³ó³Ï³ÛáõÙ ¿ :
'----------------------------------------------------------------------------------------
'startDate - àõÕ³ñÏíáÕ ýÇÉïñÇ êÏ½µÇ ³Ùë³ÃÇí ¹³ßïÇ ³ñÅ»ù
'endDate - àõÕ³ñÏíáÕ ýÇÉïñÇ ì»ñçÇ ³Ùë³ÃÇí ¹³ßïÇ ³ñÅ»ù
'docISN - ö³ëï³ÃÕÃÇ ISN-Á
Function PaySys_Check_Doc_In_SWIFT_Folder(startDate, endDate , docISN)
  Dim is_exists : is_exists = False
  Dim colN
    
  BuiltIn.Delay(2000)
  Call wTreeView.DblClickItem("|S.W.I.F.T. ²Þî                  |àõÕ³ñÏíáÕ Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñ|àõÕ³ñÏíáÕ ÷áË³ÝóáõÙÝ»ñ")
  If p1.WaitVBObject("frmAsUstPar", delay_middle).Exists Then
    Call Rekvizit_Fill("Dialog", 1, "General", "PERN", startDate)
    Call Rekvizit_Fill("Dialog", 1, "General", "PERK", endDate)
    Call ClickCmdButton(2, "Î³ï³ñ»É")
  Else 
    Log.Error "Can't find frmAsUstPar window", "", pmNormal, ErrorColor
  End If
    
  If wMDIClient.WaitVBObject("frmPttel", 3000).Exists Then
    colN = wMDIClient.vbObject("frmPttel").GetColumnIndex("DOCNUM")
    If SearchInPttel("frmPttel", colN, docISN) Then
      is_exists = true
    Else 
      Log.Error "Can't find serached row where document N = " & docISN, "", pmNormal, ErrorColor
    End If
  Else
    Log.Message "The sending documnet frmPttel doesn't exist", "", pmNormal, ErrorColor
  End If

  PaySys_Check_Doc_In_SWIFT_Folder = is_exists    
End Function


'Միջազգային Վճարման հանձնարարագիր փաստաթղթի Ընդհանուր էջի կլասս
Class ForeignPaymentOrderSentCommon
    Public isn
    Public div
    Public dep
    Public fDate
    Public docN
    Public clientTransfer
    Public payerCode
    Public residence
    Public ordClientType
    Public payerAcc
    Public payer
    Public payerAddress
    Public benefClientType
    Public benefClientAcc
    Public receiver
    Public recieverIsFinInstRA
    Public receiptCountry
    Public recieverAdress
    Public amount
    Public cur
    Public tabN
    Public check
    
    Private Sub Class_Initialize
        isn = ""
        div = ""
        dep = ""
        fDate = "  /  /  "
        docN = ""
        clientTransfer = ""
        payerCode = ""
        residence = ""
        ordClientType = ""
        payerAcc = "77700/"
        payer = ""
        payerAddress = ""
        benefClientType = ""
        benefClientAcc = ""
        receiver = ""
        recieverIsFinInstRA = 0
        receiptCountry = ""
        recieverAdress = ""
        amount = "0.00"
        cur = ""
        tabN = 1
        check = True
    End Sub
End Class

Function New_ForeignPaymentOrderSentCommon()
    Set New_ForeignPaymentOrderSentCommon = new ForeignPaymentOrderSentCommon
End Function

'Միջազգային Վճարման հանձնարարագիր փաստաթղթի Ընդհանուր էջի ստուգում
Sub Foreign_Payment_Order_Sent_CommonTab_Check(ForeignPayOrdSentCom)
    Call GoTo_ChoosedTab(ForeignPayOrdSentCom.tabN)
    'Փաստաթղթի isn-ի ստացում
    ForeignPayOrdSentCom.isn = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
    'Գրասենյակ դաշտի ստուգում
    Call Compare_Two_Values("Գրասենյակ",Get_Rekvizit_Value("Document",ForeignPayOrdSentCom.tabN,"Mask","ACSBRANCH"),ForeignPayOrdSentCom.div)
    'Բաժին դաշտի ստուգում
    Call Compare_Two_Values("Բաժին",Get_Rekvizit_Value("Document",ForeignPayOrdSentCom.tabN,"Mask","ACSDEPART"),ForeignPayOrdSentCom.dep)
    'Ամսաթիվ դաշտի ստուգում
    Call Compare_Two_Values("Ամսաթիվ",Get_Rekvizit_Value("Document",ForeignPayOrdSentCom.tabN,"General","DATE"),ForeignPayOrdSentCom.fDate)
    'Փաստաթղթի N դաշտի ստուգում
    Call Compare_Two_Values("Փաստաթղթի N",Get_Rekvizit_Value("Document",ForeignPayOrdSentCom.tabN,"General","DOCNUM"),ForeignPayOrdSentCom.docN)
    'Հաճախորդի փոխանցում դաշտի ստուգում
    Call Compare_Two_Values("Հաճախորդի փոխանցում",Get_Rekvizit_Value("Document",ForeignPayOrdSentCom.tabN,"Mask","CLITRANS"),ForeignPayOrdSentCom.clientTransfer)
    'Վճարող հաճախորդի կոդ դաշտի ստուգում
    Call Compare_Two_Values("Վճարող հաճախորդի կոդ",Get_Rekvizit_Value("Document",ForeignPayOrdSentCom.tabN,"Mask","CLICODE"),ForeignPayOrdSentCom.payerCode)
    'Ռեզիդենտություն դաշտի ստուգում
    Call Compare_Two_Values("Ռեզիդենտություն",Get_Rekvizit_Value("Document",ForeignPayOrdSentCom.tabN,"Mask","RES"),ForeignPayOrdSentCom.residence)
    'Վճարողի տվ. տիպ դաշտի ստուգում
    Call Compare_Two_Values("Վճարողի տվ. տիպ",Get_Rekvizit_Value("Document",ForeignPayOrdSentCom.tabN,"Mask","PAYOP"),ForeignPayOrdSentCom.ordClientType)
    'Վճարողի հաշիվ դաշտի ստուգում
    Call Compare_Two_Values("Վճարողի հաշիվ",Get_Rekvizit_Value("Document",ForeignPayOrdSentCom.tabN,"Bank","ACCDB"),ForeignPayOrdSentCom.payerAcc)
    'Վճարող դաշտի ստուգում
    Call Compare_Two_Values("Վճարող դաշտի",Get_Rekvizit_Value("Document",ForeignPayOrdSentCom.tabN,"Comment","PAYER"),ForeignPayOrdSentCom.payer)
    'Վճարողի հասցե դաշտի ստուգում
    Call Compare_Two_Values("Վճարողի հասցե",Get_Rekvizit_Value("Document",ForeignPayOrdSentCom.tabN,"Comment","PAYADDR"),ForeignPayOrdSentCom.payerAddress)
    'Ստացողի տվ. տիպ դաշտի ստուգում
    Call Compare_Two_Values("Ստացողի տվ. տիպ",Get_Rekvizit_Value("Document",ForeignPayOrdSentCom.tabN,"Mask","RECOP"),ForeignPayOrdSentCom.benefClientType)
    'Ստացողի հաշիվ դաշտի ստուգում
    Call Compare_Two_Values("Ստացողի հաշիվ",Get_Rekvizit_Value("Document",ForeignPayOrdSentCom.tabN,"Comment","ACCCR"),ForeignPayOrdSentCom.benefClientAcc)
    'Ստացող դաշտի ստուգում
    Call Compare_Two_Values("Ստացող",Get_Rekvizit_Value("Document",ForeignPayOrdSentCom.tabN,"Comment","RECEIVER"),ForeignPayOrdSentCom.receiver)
    'Ստացողը ՀՀ ֆին. կազմ. է դաշտի ստուգում
    Call Compare_Two_Values("Ստացողը ՀՀ ֆին. կազմ. է",Get_Rekvizit_Value("Document",ForeignPayOrdSentCom.tabN,"CheckBox","ISFINCOMPR"),ForeignPayOrdSentCom.recieverIsFinInstRA)
    'Ստացման երկիր դաշտի ստուգում
    Call Compare_Two_Values("Ստացման երկիր",Get_Rekvizit_Value("Document",ForeignPayOrdSentCom.tabN,"Mask","COUNTRY"),ForeignPayOrdSentCom.receiptCountry)
    'Ստացողի հասցե դաշտի ստուգում
    Call Compare_Two_Values("Ստացողի հասցե",Get_Rekvizit_Value("Document",ForeignPayOrdSentCom.tabN,"Comment","RECADDR"),ForeignPayOrdSentCom.recieverAdress)
    'Գումար դաշտի ստուգում
    Call Compare_Two_Values("Գումար",Get_Rekvizit_Value("Document",ForeignPayOrdSentCom.tabN,"General","SUMMA"),ForeignPayOrdSentCom.amount)
    'Արժույթ դաշտի ստուգում
    Call Compare_Two_Values("Արժույթ",Get_Rekvizit_Value("Document",ForeignPayOrdSentCom.tabN,"Mask","CUR"),ForeignPayOrdSentCom.cur)
End Sub

'Միջազգային Վճարման հանձնարարագիր փաստաթղթի Լրացուցիչ էջի կլասս
Class ForeignPaymentOrderSentAdd
    Public valDate
    Public docN
    Public pack
    Public idNum
    Public socCard
    Public aim
    Public addInfo
    Public regulatoryReporting
    Public transactionCode
    Public transitAcc
    Public correspondentAcc
    Public transitAcc2
    Public recPaySys
    Public sentPaySys
    Public benefCountry
    Public onOrder
    Public transferAim
    Public tabN
    Public check
    
    Private Sub Class_Initialize
        valDate = "  /  /  "
        docN = ""
        pack = ""
        idNum = ""
        socCard = ""
        aim = ""
        addInfo = ""
        regulatoryReporting = ""
        transactionCode = ""
        transitAcc = ""
        correspondentAcc = ""
        transitAcc2 = ""
        recPaySys = ""
        sentPaySys = ""
        benefCountry = ""
        onOrder = 0
        transferAim = ""
        tabN = 2
        check = True            
    End Sub
End Class

Function New_ForeignPaymentOrderSentAdd()
    Set New_ForeignPaymentOrderSentAdd = new ForeignPaymentOrderSentAdd
End Function

'Միջազգային Վճարման հանձնարարագիր փաստաթղթի Լրացուցիչ էջի ստուգում
Sub Foreign_Payment_Order_Sent_AddTab_Check(ForeignPayOrdSentAdd)
    Call GoTo_ChoosedTab(ForeignPayOrdSentAdd.tabN)
    'Վճարման օր դաշտի ստուգում
    Call Compare_Two_Values("Վճարման օր",Get_Rekvizit_Value("Document",ForeignPayOrdSentAdd.tabN,"General","PAYDATE"),ForeignPayOrdSentAdd.valDate)
    'Փաստաթղթի N(20) դաշտի ստուգում
    Call Compare_Two_Values("Փաստաթղթի N(20) ",Get_Rekvizit_Value("Document",ForeignPayOrdSentAdd.tabN,"General","BMDOCNUM"),ForeignPayOrdSentAdd.docN)
    'Փաթեթի համարը դաշտի ստուգում
    Call Compare_Two_Values("Փաթեթի համարը",Get_Rekvizit_Value("Document",ForeignPayOrdSentAdd.tabN,"General","PACK"),ForeignPayOrdSentAdd.pack)
    'Անձը հաստատող փաստ. դաշտի ստուգում
    Call Compare_Two_Values("Անձը հաստատող փաստ.",Get_Rekvizit_Value("Document",ForeignPayOrdSentAdd.tabN,"General","PASSNUM"),ForeignPayOrdSentAdd.idNum)
    'Սոցիալական քարտ դաշտի ստուգում
    Call Compare_Two_Values("Սոցիալական քարտ",Get_Rekvizit_Value("Document",ForeignPayOrdSentAdd.tabN,"General","REGNUM"),ForeignPayOrdSentAdd.socCard)
    'Նպատակ դաշտի ստուգում
    Call Compare_Two_Values("Նպատակ",Get_Rekvizit_Value("Document",ForeignPayOrdSentAdd.tabN,"Comment","AIM"),ForeignPayOrdSentAdd.aim)
    'Լրացուցիչ ինֆորմացիա դաշտի ստուգում
    Call Compare_Two_Values("Լրացուցիչ ինֆորմացիա",Get_Rekvizit_Value("Document",ForeignPayOrdSentAdd.tabN,"Comment","ADDINFO"),ForeignPayOrdSentAdd.addInfo)
    'Ղեկավարող հաշվետվություն դաշտի ստուգում
    Call Compare_Two_Values("Ղեկավարող հաշվետվություն",Get_Rekvizit_Value("Document",ForeignPayOrdSentAdd.tabN,"General","REGREP"),ForeignPayOrdSentAdd.regulatoryReporting)
    'Գործառնության կոդ դաշտի ստուգում
    Call Compare_Two_Values("Գործառնության կոդ",Get_Rekvizit_Value("Document",ForeignPayOrdSentAdd.tabN,"General","TRANSCODE"),ForeignPayOrdSentAdd.transactionCode)
    'Տարանցիկ հաշիվ դաշտի ստուգում
    Call Compare_Two_Values("Տարանցիկ հաշիվ",Get_Rekvizit_Value("Document",ForeignPayOrdSentAdd.tabN,"Mask","TCORRACC"),ForeignPayOrdSentAdd.transitAcc)
    'Թղթակցային հաշիվ դաշտի ստուգում
    Call Compare_Two_Values("Թղթակցային հաշիվ",Get_Rekvizit_Value("Document",ForeignPayOrdSentAdd.tabN,"Mask","CORRACC"),ForeignPayOrdSentAdd.correspondentAcc)
    'Տարանցիկ հաշիվ 2 դաշտի ստուգում
    Call Compare_Two_Values("Տարանցիկ հաշիվ 2",Get_Rekvizit_Value("Document",ForeignPayOrdSentAdd.tabN,"Mask","TCORRACCT"),ForeignPayOrdSentAdd.transitAcc2)
    'Ընդ. վճ. համակարգ դաշտի ստուգում
    Call Compare_Two_Values("Ընդ. վճ. համակարգ",Get_Rekvizit_Value("Document",ForeignPayOrdSentAdd.tabN,"Mask","PAYSYSIN"),ForeignPayOrdSentAdd.recPaySys)
    'Ուղ. վճ համակարգ դաշտի ստուգում
    Call Compare_Two_Values("Ուղ. վճ համակարգ",Get_Rekvizit_Value("Document",ForeignPayOrdSentAdd.tabN,"Mask","PAYSYSOUT"),ForeignPayOrdSentAdd.sentPaySys)
    'Ստացողի ռեզ. երկիր դաշտի ստուգում
    Call Compare_Two_Values("Ստացողի ռեզ. երկիր",Get_Rekvizit_Value("Document",ForeignPayOrdSentAdd.tabN,"Mask","BENCLICOUNTRY"),ForeignPayOrdSentAdd.benefCountry)
    'Համաձայն հրահանգի դաշտի ստուգում
    Call Compare_Two_Values("Համաձայն հրահանգի",Get_Rekvizit_Value("Document",ForeignPayOrdSentAdd.tabN,"CheckBox","ONORDER"),ForeignPayOrdSentAdd.onOrder)
    'Փոխանցման նպատակ դաշտի ստուգում
    Call Compare_Two_Values("Փոխանցման նպատակ",Get_Rekvizit_Value("Document",ForeignPayOrdSentAdd.tabN,"Mask","PAYAIM"),ForeignPayOrdSentAdd.transferAim)
End Sub

'Միջազգային Վճարման հանձնարարագիր փաստաթղթի Ֆին.կազմ. էջի կլասս
Class ForeignPaymentOrderSentFinOrg
    Public reference
    Public ordInstType
    Public ordInst
    Public ordInstPID
    Public payBankCorr
    Public payBankCorrAcc
    Public paybankCorrType
    Public intInstType
    Public intInst
    Public intInstPID
    Public accWithInstType
    Public accWithInst
    Public accWithInstPID
    Public uniqueID
    Public tabN
    Public check
    Private Sub Class_Initialize
        reference = ""
        ordInstType = ""
        ordInst = ""
        ordInstPID = ""
        payBankCorr = ""
        payBankCorrAcc = ""
        paybankCorrType = ""
        intInstType = ""
        intInst = ""
        intInstPID = ""
        accWithInstType = ""
        accWithInst = ""
        accWithInstPID = ""
        uniqueID = ""
        tabN = 3
        check = True
    End Sub    
End Class

Function New_ForeignPaymentOrderSentFinOrg()
    Set New_ForeignPaymentOrderSentFinOrg = new ForeignPaymentOrderSentFinOrg
End Function

'Միջազգային Վճարման հանձնարարագիր փաստաթղթի Ֆին.կազմ. էջի ստուգում
Sub Foreign_Payment_Order_Sent_FinOrg_Check(ForeignPayOrdSentFinOrg)
    Call GoTo_ChoosedTab(ForeignPayOrdSentFinOrg.tabN)
    'Հղում դաշտի ստուգում
    Call Compare_Two_Values("Հղում",Get_Rekvizit_Value("Document",ForeignPayOrdSentFinOrg.tabN,"General","REF"),ForeignPayOrdSentFinOrg.reference)
    'Վճարող կազմակերպ. տվ. տիպ դաշտի ստուգում
    Call Compare_Two_Values("Վճարող կազմակերպ. տվ. տիպ",Get_Rekvizit_Value("Document",ForeignPayOrdSentFinOrg.tabN,"Mask","PINSTOP"),ForeignPayOrdSentFinOrg.ordInstType)
    'Վճարող կազմակերպ. դաշտի ստուգում
    Call Compare_Two_Values("Վճարող կազմակերպ.",Get_Rekvizit_Value("Document",ForeignPayOrdSentFinOrg.tabN,"Comment","PAYINST"),ForeignPayOrdSentFinOrg.ordInst)
    'Վճարող կազմակերպ. հաշիվ դաշտի ստուգում
    Call Compare_Two_Values("Վճարող կազմակերպ. հաշիվ",Get_Rekvizit_Value("Document",ForeignPayOrdSentFinOrg.tabN,"Comment","PINSTID"),ForeignPayOrdSentFinOrg.ordInstPID)
    'Վճարող բանկի թղթակից դաշտի ստուգում
    Call Compare_Two_Values("Վճարող բանկի թղթակից",Get_Rekvizit_Value("Document",ForeignPayOrdSentFinOrg.tabN,"Comment","PCORBANK"),ForeignPayOrdSentFinOrg.payBankCorr)
    'Վճարող բանկի թղթակցի հաշիվ դաշտի ստուգում
    Call Compare_Two_Values("Վճարող բանկի թղթակցի հաշիվ",Get_Rekvizit_Value("Document",ForeignPayOrdSentFinOrg.tabN,"Comment","PCORID"),ForeignPayOrdSentFinOrg.payBankCorrAcc)
    'Միջնորդ բանկի տվ. տիպ դաշտի ստուգում
    Call Compare_Two_Values("Միջնորդ բանկի տվ. տիպ",Get_Rekvizit_Value("Document",ForeignPayOrdSentFinOrg.tabN,"Mask","MEDOP"),ForeignPayOrdSentFinOrg.intInstType)
    'Միջնորդ բանկ դաշտի ստուգում
    Call Compare_Two_Values("Միջնորդ բանկ",Get_Rekvizit_Value("Document",ForeignPayOrdSentFinOrg.tabN,"Comment","MEDBANK"),ForeignPayOrdSentFinOrg.intInst)
    'Միջնորդ բանկի հաշիվ դաշտի ստուգում
    Call Compare_Two_Values("Միջնորդ բանկի հաշիվ",Get_Rekvizit_Value("Document",ForeignPayOrdSentFinOrg.tabN,"Comment","MEDID"),ForeignPayOrdSentFinOrg.intInstPID)
    'Ստացող կազմակերպ. տվ. տիպ դաշտի ստուգում
    Call Compare_Two_Values("Ստացող կազմակերպ. տվ. տիպ",Get_Rekvizit_Value("Document",ForeignPayOrdSentFinOrg.tabN,"Mask","RINSTOP"),ForeignPayOrdSentFinOrg.accWithInstType)
    'Ստացող կազմակերպ. դաշտի ստուգում
    Call Compare_Two_Values("Ստացող կազմակերպ.",Get_Rekvizit_Value("Document",ForeignPayOrdSentFinOrg.tabN,"Comment","RECINST"),ForeignPayOrdSentFinOrg.accWithInst)
    'Ստացող կազմակերպ. հաշիվ դաշտի ստուգում
    Call Compare_Two_Values("Ստացող կազմակերպ. հաշիվ",Get_Rekvizit_Value("Document",ForeignPayOrdSentFinOrg.tabN,"Comment","RINSTID"),ForeignPayOrdSentFinOrg.accWithInstPID)
    'Ունիկալ համար դաշտի ստուգում
    Call Compare_Two_Values("Ունիկալ համար",Get_Rekvizit_Value("Document",ForeignPayOrdSentFinOrg.tabN,"General","UNIQUENUM"),ForeignPayOrdSentFinOrg.uniqueID)    
End Sub

'Միջազգային Վճարման հանձնարարագիր փաստաթղթի Գանձում փոխանցումից էջի կլասս
Class ForeignPaymentOrderSentTransCharge
    Public detailsOfCharge
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
    Public legPosition
    Public busField
    Public comment
    Public tabN
    Public check
    Private Sub Class_Initialize 
        detailsOfCharge = ""
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
        time = ""
        legPosition = ""
        busField = ""
        comment = ""
        tabN = 4
        check = True
    End Sub 
End Class

Function New_ForeignPaymentOrderSentTransCharge()
    Set New_ForeignPaymentOrderSentTransCharge = new ForeignPaymentOrderSentTransCharge
End Function
'Միջազգային Վճարման հանձնարարագիր փաստաթղթի Գանձում փոխանցումից էջի ստուգում
Sub Foreign_Payment_Order_Sent_TChargeTab_Check(ForeignPayOrdSentTCharge)
    Call GoTo_ChoosedTab(ForeignPayOrdSentTCharge.tabN)
    'Ծախսերի մանրամասնություն դաշտի ստուգում
    Call Compare_Two_Values("Ծախսերի մանրամասնություն",Get_Rekvizit_Value("Document",ForeignPayOrdSentTCharge.tabN,"Mask","EXPTYPE"),ForeignPayOrdSentTCharge.detailsOfCharge)
    'Գանձման հաշիվ դաշտի ստուգում
    Call Compare_Two_Values("Գանձման հաշիվ",Get_Rekvizit_Value("Document",ForeignPayOrdSentTCharge.tabN,"Mask","CHRGACC"),ForeignPayOrdSentTCharge.chargesAcc)
    'Արժույթ դաշտի ստուգում
    Call Compare_Two_Values("Արժույթ",Get_Rekvizit_Value("Document",ForeignPayOrdSentTCharge.tabN,"Mask","CHRGCUR"),ForeignPayOrdSentTCharge.cur)
    'ԿԲ փոխարժեք դաշտի ստուգում
    Call Compare_Two_Values("ԿԲ փոխարժեք",Get_Rekvizit_Value("Document",ForeignPayOrdSentTCharge.tabN,"Course","CHRGCBCRS"),ForeignPayOrdSentTCharge.CBCourse)
    'Գանձման տեսակ դաշտի ստուգում
    Call Compare_Two_Values("Գանձման տեսակ",Get_Rekvizit_Value("Document",ForeignPayOrdSentTCharge.tabN,"Mask","PAYSCALE"),ForeignPayOrdSentTCharge.chargeType1)
    'Տոկոս դաշտի ստուգում
    Call Compare_Two_Values("Տոկոս",Get_Rekvizit_Value("Document",ForeignPayOrdSentTCharge.tabN,"General","PRSNT"),ForeignPayOrdSentTCharge.interest1)
    'Գումար դաշտի ստուգում
    Call Compare_Two_Values("Գումար",Get_Rekvizit_Value("Document",ForeignPayOrdSentTCharge.tabN,"General","CHRGSUM"),ForeignPayOrdSentTCharge.sum1)
    'Եկամտի հաշիվ դաշտի ստուգում
    Call Compare_Two_Values("Եկամտի հաշիվ",Get_Rekvizit_Value("Document",ForeignPayOrdSentTCharge.tabN,"Mask","CHRGINC"),ForeignPayOrdSentTCharge.incomeAcc1)
    'Գանձման տեսակ 2 դաշտի ստուգում
    Call Compare_Two_Values("Գանձման տեսակ 2",Get_Rekvizit_Value("Document",ForeignPayOrdSentTCharge.tabN,"Mask","PAYSCALE2"),ForeignPayOrdSentTCharge.chargeType2)
    'Տոկոս 2 դաշտի ստուգում
    Call Compare_Two_Values("Տոկոս 2",Get_Rekvizit_Value("Document",ForeignPayOrdSentTCharge.tabN,"General","PRSNT2"),ForeignPayOrdSentTCharge.interest2)
    'Գումար 2 դաշտի ստուգում
    Call Compare_Two_Values("Գումար 2",Get_Rekvizit_Value("Document",ForeignPayOrdSentTCharge.tabN,"General","CHRGSUM2"),ForeignPayOrdSentTCharge.sum2)
    'Եկամտի հաշիվ 2 դաշտի ստուգում
    Call Compare_Two_Values("Եկամտի հաշիվ 2",Get_Rekvizit_Value("Document",ForeignPayOrdSentTCharge.tabN,"Mask","CHRGINC2"),ForeignPayOrdSentTCharge.incomeAcc2)
    'Առք/Վաճառք դաշտի ստուգում
    Call Compare_Two_Values("Առք/Վաճառք",Get_Rekvizit_Value("Document",ForeignPayOrdSentTCharge.tabN,"Mask","CUPUSA"),ForeignPayOrdSentTCharge.purSale)
    'Գործողության տեսակ դաշտի ստուգում
    Call Compare_Two_Values("Գործողության տեսակ",Get_Rekvizit_Value("Document",ForeignPayOrdSentTCharge.tabN,"Mask","CURTES"),ForeignPayOrdSentTCharge.optype)
    'Գոծողության վայր դաշտի ստուգում
    Call Compare_Two_Values("Գոծողության վայր",Get_Rekvizit_Value("Document",ForeignPayOrdSentTCharge.tabN,"Mask","CURVAIR"),ForeignPayOrdSentTCharge.opPlace)
    'Ժամանակ դաշտի ստուգում
    Call Compare_Two_Values("Ժամանակ",Get_Rekvizit_Value("Document",ForeignPayOrdSentTCharge.tabN,"Mask","TIME"),ForeignPayOrdSentTCharge.time)
    'Իրավաբանական կարգավիճակ դաշտի ստուգում
    Call Compare_Two_Values("Իրավաբանական կարգավիճակ",Get_Rekvizit_Value("Document",ForeignPayOrdSentTCharge.tabN,"Mask","JURSTAT"),ForeignPayOrdSentTCharge.legPosition)
    'Գործունեության ոլորտ դաշտի ստուգում
    Call Compare_Two_Values("Գործունեության ոլորտ",Get_Rekvizit_Value("Document",ForeignPayOrdSentTCharge.tabN,"Mask","VOLORT"),ForeignPayOrdSentTCharge.busField)
    'Մեկնաբանություն դաշտի ստուգում
    Call Compare_Two_Values("Մեկնաբանություն",Get_Rekvizit_Value("Document",ForeignPayOrdSentTCharge.tabN,"General","COMM"),ForeignPayOrdSentTCharge.comment)
End Sub
'Միջազգային Վճարման հանձնարարագիր փաստաթղթի Դրամարկղ էջի կլասս
Class ForeignPaymentOrderSentCashDesk
    Public cashDesk
    Public cashLabel
    Public depositor
    Public idNum
    Public eMail
    Public base
    Public transInput
    Public transCurr
    Public chargeInput
    Public chargeCurr
    Public dateOfBirth
    Public birthPlace
    Public stateRegCertNum
    Public taxCode
    Public tabN
    Public check 
    Private Sub Class_Initialize
        cashDesk = ""
        cashLabel = ""
        depositor = ""
        idNum = ""
        eMail = ""
        base = ""
        transInput = "0.00"
        transCurr = ""
        chargeInput = "0.00"
        chargeCurr = ""
        dateOfBirth = "  /  /  "
        birthPlace = ""
        stateRegCertNum = ""
        taxCode = ""
        tabN = 5
        check = True     
    End Sub
End Class

Function New_ForeignPaymentOrderSentCashDesk()
    Set New_ForeignPaymentOrderSentCashDesk = new ForeignPaymentOrderSentCashDesk
End Function
'Միջազգային Վճարման հանձնարարագիր փաստաթղթի Դրամարկղ էջի ստուգում
Sub Foreign_Payment_Order_Sent_CDesk_Check(ForeignPayOrdSentCashDesk)
    Call GoTo_ChoosedTab(ForeignPayOrdSentCashDesk.tabN)
    'Դրամարկղ դաշտի ստուգում
    Call Compare_Two_Values("Դրամարկղ",Get_Rekvizit_Value("Document",ForeignPayOrdSentCashDesk.tabN,"Mask","KASSA"),ForeignPayOrdSentCashDesk.cashDesk)
    'Դրամարկղի նիշ դաշտի ստուգում
    Call Compare_Two_Values("Դրամարկղի նիշ",Get_Rekvizit_Value("Document",ForeignPayOrdSentCashDesk.tabN,"Mask","KASSIMV"),ForeignPayOrdSentCashDesk.cashLabel)
    'Մուծող դաշտի ստուգում
    Call Compare_Two_Values("Մուծող",Get_Rekvizit_Value("Document",ForeignPayOrdSentCashDesk.tabN,"Comment","PAYER1"),ForeignPayOrdSentCashDesk.depositor)
    'Անձը հաստատող փաստ. դաշտի ստուգում
    Call Compare_Two_Values("Անձը հաստատող փաստ.",Get_Rekvizit_Value("Document",ForeignPayOrdSentCashDesk.tabN,"General","PASSNUM1"),ForeignPayOrdSentCashDesk.idNum)
    'Էլ. հասցե դաշտի ստուգում
    Call Compare_Two_Values("Էլ. հասցե",Get_Rekvizit_Value("Document",ForeignPayOrdSentCashDesk.tabN,"General","EMAIL"),ForeignPayOrdSentCashDesk.eMail)
    'Հիմք դաշտի ստուգում
    Call Compare_Two_Values("Հիմք",Get_Rekvizit_Value("Document",ForeignPayOrdSentCashDesk.tabN,"Comment","BASE"),ForeignPayOrdSentCashDesk.base)
    'Մուտք փոխանցման համար դաշտի ստուգում
    Call Compare_Two_Values("Մուտք փոխանցման համար",Get_Rekvizit_Value("Document",ForeignPayOrdSentCashDesk.tabN,"General","INTRANSSUM"),ForeignPayOrdSentCashDesk.transInput)
    'Արժույթ դաշտի ստուգում
    Call Compare_Two_Values("Արժույթ",Get_Rekvizit_Value("Document",ForeignPayOrdSentCashDesk.tabN,"Mask","INTRANSCUR"),ForeignPayOrdSentCashDesk.transCurr)
    'Մուտք գանձման համար դաշտի ստուգում
    Call Compare_Two_Values("Մուտք գանձման համար",Get_Rekvizit_Value("Document",ForeignPayOrdSentCashDesk.tabN,"General","INCHRGSUM"),ForeignPayOrdSentCashDesk.chargeInput)
    'Արժույթ 2 դաշտի ստուգում
    Call Compare_Two_Values("Արժույթ 2",Get_Rekvizit_Value("Document",ForeignPayOrdSentCashDesk.tabN,"Mask","INCHRGCUR"),ForeignPayOrdSentCashDesk.chargeCurr)
    'Ծննդյան ամսաթիվ դաշտի ստուգում
    Call Compare_Two_Values("Ծննդյան ամսաթիվ",Get_Rekvizit_Value("Document",ForeignPayOrdSentCashDesk.tabN,"General","DATEBIRTH"),ForeignPayOrdSentCashDesk.dateOfBirth)
    'Ծննդավայր դաշտի ստուգում
    Call Compare_Two_Values("Ծննդավայր",Get_Rekvizit_Value("Document",ForeignPayOrdSentCashDesk.tabN,"Comment","BIRTHPLACE"),ForeignPayOrdSentCashDesk.birthPlace)
    'Պետ. գրանցման վկայականի համար դաշտի ստուգում
    Call Compare_Two_Values("Պետ. գրանցման վկայականի համար",Get_Rekvizit_Value("Document",ForeignPayOrdSentCashDesk.tabN,"General","REGCERT"),ForeignPayOrdSentCashDesk.stateRegCertNum)
    'ՀՎՀՀ (Վճարող) դաշտի ստուգում
    Call Compare_Two_Values("ՀՎՀՀ (Վճարող)",Get_Rekvizit_Value("Document",ForeignPayOrdSentCashDesk.tabN,"Comment","TAXCODSD"),ForeignPayOrdSentCashDesk.taxCode)
End Sub
Class ForeignPaymentOrderSentOther
    Public comission
    Public accType
    Public fType
    Public sendRec
    Public msgType
    Public phone
    Public refusal
    Public orderCode()
    Public curr()
    Public sum()
    Public corrBankChargesCount
    Public tabN
    Public check
    Private Sub Class_Initialize
        comission = "0.00"
        accType = ""
        fType = ""
        sendRec = ""
        msgType = ""
        phone = ""
        refusal = ""
        corrBankChargesCount = chargesCount
        Redim orderCode(corrBankChargesCount)
        Redim curr(corrBankChargesCount)
        Redim sum(corrBankChargesCount)
        For corrBankChargesCount = 0 To chargesCount - 1
            orderCode(corrBankChargesCount) = ""
            curr(corrBankChargesCount) = ""
            sum(corrBankChargesCount) = "0.00"
        Next
        tabN = 6
        check = True
    End Sub
End Class

Function New_ForeignPaymentOrderSentOther(chargeCount)
    chargesCount = chargeCount
    Set New_ForeignPaymentOrderSentOther = new ForeignPaymentOrderSentOther
End Function

Sub Foreign_Payment_Order_Sent_OtherTab_Check(ForeignPayOrdSentOther)
    Dim grid, rowN , i
    Call GoTo_ChoosedTab(ForeignPayOrdSentOther.tabN)
    'Միջնորդավճար դաշտի ստուգում
    Call Compare_Two_Values("Միջնորդավճար",Get_Rekvizit_Value("Document",ForeignPayOrdSentOther.tabN,"General","COMMISSION"),ForeignPayOrdSentOther.comission)
    'Հաշվի տիպ դաշտի ստուգում
    Call Compare_Two_Values("Հաշվի տիպ",Get_Rekvizit_Value("Document",ForeignPayOrdSentOther.tabN,"Mask","ACCTYPE"),ForeignPayOrdSentOther.accType) 
    'Տիպ դաշտի ստուգում
    Call Compare_Two_Values("Տիպ",Get_Rekvizit_Value("Document",ForeignPayOrdSentOther.tabN,"Mask","CORTYPE"),ForeignPayOrdSentOther.fType) 
    'Ուղարկող/Ստացող դաշտի ստուգում
    Call Compare_Two_Values("Ուղարկող/Ստացող",Get_Rekvizit_Value("Document",ForeignPayOrdSentOther.tabN,"Comment","SNDREC"),ForeignPayOrdSentOther.sendRec) 
    'Հաղ. տիպ դաշտի ստուգում
    Call Compare_Two_Values("Հաղ. տիպ",Get_Rekvizit_Value("Document",ForeignPayOrdSentOther.tabN,"General","MT"),ForeignPayOrdSentOther.msgType) 
    'Հեռախոս դաշտի ստուգում
    Call Compare_Two_Values("Հեռախոս",Get_Rekvizit_Value("Document",ForeignPayOrdSentOther.tabN,"General","PHONE"),ForeignPayOrdSentOther.phone) 
    'Մերժում դաշտի ստուգում
    Call Compare_Two_Values("Մերժում",Get_Rekvizit_Value("Document",ForeignPayOrdSentOther.tabN,"General","REFUSE"),ForeignPayOrdSentOther.refusal)
    'Թղթակից բանկի գանձումներ աղյուսակի ստուգում
    Set grid = wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame_" & ForeignPayOrdSentOther.tabN).VBObject("DocGrid")
    For i = 0 to ForeignPayOrdSentOther.corrBankChargesCount - 1
        rowN = Get_Cell_Row_Grid (1, "Document", 6, ForeignPayOrdSentOther.orderCode(i))
        'Կարգադրության կոդ սյան ստուգում
        Call Check_Value_Grid (0 ,rowN, "Document", 6, ForeignPayOrdSentOther.orderCode(i))
        'Արժ. սյան ստուգում
        Call Check_Value_Grid (1 ,rowN, "Document", 6, ForeignPayOrdSentOther.curr(i))
        'Գումար սյան ստուգում
        Call Check_Value_Grid (2 ,rowN, "Document", 6, ForeignPayOrdSentOther.sum(i))
    Next
    If grid.ApproxCount <> ForeignPayOrdSentOther.corrBankChargesCount Then
        Log.Error "Charges count is not equal to " & ForeignPayOrdSentOther.corrBankChargesCount & " It is " & grid.ApproxCount
    End If 
End Sub

Class ForeignPaymentOrderSent
    Public commonTab
    Public addTab
    Public finOrgTab
    Public tChargeTab
    Public cDeskTab
    Public otherTab
    Public attachTab
    Private Sub Class_Initialize
        Set commonTab = New_ForeignPaymentOrderSentCommon()
        Set addTab = New_ForeignPaymentOrderSentAdd()
        Set finOrgTab = New_ForeignPaymentOrderSentFinOrg()
        Set tChargeTab = New_ForeignPaymentOrderSentTransCharge()
        Set cDeskTab = New_ForeignPaymentOrderSentCashDesk()
        Set otherTab = New_ForeignPaymentOrderSentOther(chargesCount)
        Set attachTab = New_Attached_Tab(fCount, lCount, dCount)
        attachTab.tabN = 7
    End Sub
End Class

Function New_ForeignPaymentOrderSent(chargeC, fileC ,linkC , deleteC)
    chargesCount = chargeC
    fCount = fileC
    lCount = linkC
    dCount = deleteC
    Set New_ForeignPaymentOrderSent = new ForeignPaymentOrderSent
End Function

Sub Foreign_Payment_Order_Sent_Check(FPO)
    'Ընդհանուր էջի ստուգում
    If FPO.commonTab.check Then
        Call Foreign_Payment_Order_Sent_CommonTab_Check(FPO.commonTab)        
    End If
    'Լրացուցիչ էջի ստուգում
    If FPO.addTab.check Then
        Call Foreign_Payment_Order_Sent_AddTab_Check(FPO.addTab)   
    End If
    'Ֆին. կազմակերպ. էջի ստուգում
    If FPO.finOrgTab.check Then
        Call Foreign_Payment_Order_Sent_FinOrg_Check(FPO.finOrgTab)    
    End If
    'Գանձում փոխանցումից էջի ստուգում
    If FPO.tChargeTab.check Then
        Call Foreign_Payment_Order_Sent_TChargeTab_Check(FPO.tChargeTab)
    End If 
    'Դրամարկղ էջի ստուգում
    If FPO.cDeskTab.check Then
        Call Foreign_Payment_Order_Sent_CDesk_Check(FPO.cDeskTab)
    End If
    'Այլ էջի ստուգում
    If FPO.otherTab.check Then
        Call Foreign_Payment_Order_Sent_OtherTab_Check(FPO.otherTab)
    End If
    'Կցված էջի ստուգում 
    Call Attach_Tab_Check(FPO.attachTab)    
End Sub