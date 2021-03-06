Option Explicit
'USEUNIT Library_Common
'USEUNIT Library_Contracts
'USEUNIT Online_PaySys_Library
'USEUNIT Library_Colour
'USEUNIT Constants

Dim i : i = 0

'-------------------------------------------------
'Արժեթղթերի փոխանակման հաստատում փաստաթղթի լրացում
'-------------------------------------------------
'docNumber - Փաստաթղթի համարը
'summa - Գումար դաշտի արժեք
'nbAcc - Հաշիվ դաշտի արժեք
'aim - Նպատակ դաշտի արժեք
'fISN - Փաստատթղթի ISN-ը
Sub SWIFT_Doc_Fill(docNumber, ref, opType, orgType1 , firstOrg, opType2 , secOrg, date1, date2, _
                   curr1, curr2, sendRec, summ, opType3, thirdOrg, opType4, fourthOrg,fISN )
    
  Dim rekvName, docN
    
  If wMDIClient.WaitVBObject("frmASDocForm", 3000).Exists Then
    'Ստեղծվող ISN - ի փաստատթղթի  վերագրում փոփոխականին
    fISN = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
    'Փաստաթղթի N դաշտի արժեքի վերագրում փոփոխականին
    docN = Get_Rekvizit_Value("Document",1,"General","BMDOCNUM")
    docNumber = Left(docN, 6)
    'Հղում դաշտի լրացում
    Call Rekvizit_Fill("Document", 1, "General", "REF", ref)
    'Գործողության տեսակ դաշտի լրացում
    Call Rekvizit_Fill("Document", 1, "General", "OPERTYPE", opType)
    'Առաջին կազմակերպ. տվ. տիպ դաշտի լրացում
    Call Rekvizit_Fill("Document", 1, "General", "PARTYAOP", orgType1)
    'Առաջին կազմակերպություն դաշտի լրացում
    Call Rekvizit_Fill("Document", 1, "General", "PARTYA", firstOrg)
    'Երկրորդ կազմակերպ. տվ. տիպ դաշտի լրացում
    Call Rekvizit_Fill("Document", 1, "General", "PARTYBOP", opType2)
    'Երկրորդ կազմակերպություն դաշտի լրացում
    Call Rekvizit_Fill("Document", 1, "General", "PARTYB", secOrg)
    
    'Անցում 2.Գործառնության տվյալներ էջին
    Call GoTo_ChoosedTab(2)
    'Ամսաթիվ դաշտի լրացում
    Call Rekvizit_Fill("Document", 2, "General", "DATE", date1)
    'Վճարման օր  դաշտի լրացում
    Call Rekvizit_Fill("Document", 2, "General", "PAYDATE", date2)
    'Արժույթ(առք) դաշտի լրացում
    Call Rekvizit_Fill("Document", 2, "General", "CURB", curr1)
    'Արժույթ(վաճառք) դաշտի լրացում
    Call Rekvizit_Fill("Document", 2, "General", "CURS", curr2)
    'Գումար(առք) դաշտի լրացում
    Call Rekvizit_Fill("Document", 2, "General", "SUMMAB", summ)
    'Ուղարկող/Ստացող դաշտի լրացում
    Call Rekvizit_Fill("Document", 2, "General", "SNDREC", sendRec)
    
    'Անցում 3.տվյալներ (առք) էջին
    Call GoTo_ChoosedTab(3)
    'Ստացող գործակալի տվ. տիպ դաշտի լրացում
    Call Rekvizit_Fill("Document", 3, "General", "RAGENTBOP", opType3)
    'Ստացող գործակալ դաշտի լրացում
    Call Rekvizit_Fill("Document", 3, "General", "RAGENTB", thirdOrg)
    
    'Անցում 4.տվյալներ (Վաճառք) էջին
    Call GoTo_ChoosedTab(4)
    'Ստացող գործակալի տվ. տիպ դաշտի լրոցում
    Call Rekvizit_Fill("Document", 4, "General", "RAGENTSOP", opType4)
    'Ստացող գործակալ դաշտի լրացում
    Call Rekvizit_Fill("Document", 4, "General", "RAGENTS", fourthOrg)
    
    Call ClickCmdButton(1, "Î³ï³ñ»É")
  Else
    Log.Error "Can't open frmASDocForm window", "", pmNormal, ErrorColor
  End If    
End Sub

'----------------------------------------------------------------------------------------
' àõÕ³ñÏíáÕ ÷áË³ÝóáõÙÝ»ñ ÃÕÃ³å³Ý³ÏáõÙ ÷³ëï³ÃÕÃÇ ³éÏ³ÛáõÃÛ³Ý ëïáõ·áõÙ
'üáõÝÏóÇ³Ý í»ñ³¹³ñÓÝáõÙ ¿ true ÷³ëï³ÃÕÃÇ ³éÏ³ÛáõÃÛ³Ý ¹»åùáõÙ, Ñ³Ï³é³Ï ¹»åùáõÙ` false :
'----------------------------------------------------------------------------------------
Function SWIFT_Check_Doc_In_Sending_SecrOrd_Folder(fISN)
  Dim is_exist : is_exist = False
  Dim colN
    
  BuiltIn.Delay(2000)
  Call wTreeView.DblClickItem("|S.W.I.F.T. ²Þî                  |àõÕ³ñÏíáÕ Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñ|àõÕ³ñÏíáÕ ÷áË³ÝóáõÙÝ»ñ")
  If p1.WaitVBObject("frmAsUstPar", 1000).Exists Then
    Call Rekvizit_Fill("Dialog", 1, "General", "PERN", "!" & "[End]" & "[Del]")
    Call Rekvizit_Fill("Dialog", 1, "General", "PERK", "!" & "[End]" & "[Del]")
    Call ClickCmdButton(2, "Î³ï³ñ»É")
  Else 
    Log.Error "Can't find frmAsUstPar window", "", pmNormal, ErrorColor
  End if
  If wMDIClient.WaitVBObject("frmPttel", 3000).Exists Then
    colN = wMDIClient.vbObject("frmPttel").GetColumnIndex("DOCNUM")
    If SearchInPttel("frmPttel", colN, fISN) Then
    is_exist = True
  End If
  Else
    Log.Message "The sending documents frmPttel does't exist", "", pmNormal, ErrorColor
  End If
    
  SWIFT_Check_Doc_In_Sending_SecrOrd_Folder = is_exist
End Function



' SWIFT ԱՇՏ/ Նոր հաղորդագրություններ/ Փոխանցում իր հաշիվներով ՀՏ 200
Class TransferToHisAccounts
  
    Public fISN
    Public acsBranch
    Public acsDepart
    Public docNum
    Public wDate
    Public rinStop
    Public recOrgAcc
    Public recOrg
    Public wSumma
    Public wCur
    Public txKey
    Public wPackN
    Public addInfo
    Public fileName
    Public directoryName
    Public bmioDate
    Public bmioTime
    Public repaymentDate
    Public wDeadline
    Public wPrior
    Public BankingCham
    Public sendRec
    Public CorBankAcc
    Public CorBank
    Public IntBankDataType
    Public IntBankAcc
    Public IntBank
    Public clcikBOrNo
    Public clcikBOrNo2
    Public clcikBOrNo3
    Public finOrginization(2)

    Private Sub Class_Initialize
        fISN = ""
        acsBranch = ""
        acsDepart = ""
        docNum = ""
        wDate = ""
        rinStop = ""
        recOrgAcc = ""
        recOrg = ""
        wSumma = ""
        wCur = ""
        txKey = ""
        wPackN = ""
        addInfo = ""
        fileName = ""
        directoryName = ""
        bmioDate = ""
        bmioTime = ""
        repaymentDate = ""
        wDeadline = ""
        wPrior = ""
        BankingCham = ""
        sendRec = ""
        CorBankAcc = ""
        CorBank = ""
        IntBankDataType = ""
        IntBankAcc = ""
        IntBank = ""
        clcikBOrNo = False
        clcikBOrNo2 = False
        clcikBOrNo3 = False
    
        For i = 0 to 2
            Set finOrginization(i) = New_FinancialOrganizations()
        Next
        
    End Sub  
End Class

Function New_TransferToHisAccounts()
    Set New_TransferToHisAccounts= NEW TransferToHisAccounts      
End Function

' Լրացնել "Փոխանցում իր հաշիվներով" փաստաթղթի դաշտերը
Sub Fill_TransferToHisAccounts(TransferToHisAccounts)
      
      Dim frmASDocForm, wStatus
      ' Վերցնել փաստաթղթի ISN-ը
      Set frmASDocForm = wMDIClient.VBObject("frmASDocForm")
      TransferToHisAccounts.fISN = frmASDocForm.DocFormCommon.Doc.ISN
      ' Գրասենյակ դաշտի լրացում
      Call Rekvizit_Fill("Document", 1, "General", "ACSBRANCH", TransferToHisAccounts.acsBranch)
      ' Բաժին դաշտի լրացում
      Call Rekvizit_Fill("Document", 1, "General", "ACSDEPART", TransferToHisAccounts.acsDepart)
      ' Ստանալ Փաստաթղթի N-ը
      TransferToHisAccounts.docNum = Get_Rekvizit_Value("Document",1,"General","BMDOCNUM")
      ' Ամսաթիվ դաշտի լրացում
      Call Rekvizit_Fill("Document", 1, "General", "DATE", TransferToHisAccounts.wDate)
      ' Ստացող կազմակերպ տվ. տիպ դաշտի լրացում
      Call Rekvizit_Fill("Document", 1, "General", "RINSTOP", TransferToHisAccounts.rinStop)
      ' Ստացող կազմակերպ. հաշիվ դաշտի լրացում
      Call Rekvizit_Fill("Document", 1, "General", "RINSTID", TransferToHisAccounts.recOrgAcc)

     wStatus = True
     BuiltIn.Delay(1500)
     
     If TransferToHisAccounts.clcikBOrNo Then
     
           wStatus = False
           frmASDocForm.VBObject("TabFrame").VBObject("AsTpComment_2").VBObject("CmdAdditional").Click
           If p1.WaitVBObject("frmAsUstPar",2000).Exists Then
                   ' Ստացող կազմակերպ. դաշտի լրացում
                   wStatus = FinancialOrganizationsFilter(TransferToHisAccounts.finOrginization(0))
            Else  
                  Log.Error("Ֆինանսական կազմակերպություններ դիալոգը չի բացվել")
            End If
      
     Else
        Call Rekvizit_Fill("Document", 1, "General", "RECINST", TransferToHisAccounts.recOrg)
     End If     
      
      If Not wStatus Then
            Call Rekvizit_Fill("Document", 1, "General", "RECINST", TransferToHisAccounts.recOrg)
      End If
              
      ' Գումար դաշտի լրացում
      Call Rekvizit_Fill("Document", 1, "General", "SUMMA", TransferToHisAccounts.wSumma)
      ' Արժույթ դաշտի լրացում
      Call Rekvizit_Fill("Document", 1, "General", "CUR", TransferToHisAccounts.wCur)
      ' Բանալի դաշտի լրացում
      Call Rekvizit_Fill("Document", 1, "General", "TXKEY", TransferToHisAccounts.txKey)
              
      ' Փաթեթի համարը դաշտի լրացում
      Call Rekvizit_Fill("Document", 2, "General", "PACK", TransferToHisAccounts.wPackN)
      ' Լրացուցիչ ինֆորմացիա դաշտի լրացում
      Call Rekvizit_Fill("Document", 2, "General", "ADDINFO", TransferToHisAccounts.addInfo)
      ' Ֆայլի անուն դաշտի լրացում
      Call Rekvizit_Fill("Document", 2, "General", "BMNAME", TransferToHisAccounts.fileName)
      ' Դիրեկտորայի անուն դաշտի լրացում
      Call Rekvizit_Fill("Document", 2, "General", "BMDIRECT", TransferToHisAccounts.directoryName)
      ' Ամսաթիվ (Ուղարկման/Ստացման) դաշտի լրացում
      Call Rekvizit_Fill("Document", 2, "General", "BMIODATE", TransferToHisAccounts.bmioDate)
      ' Ժամանակ (Ուղարկման/Ստացման) դաշտի լրացում
      Call Rekvizit_Fill("Document", 2, "General", "BMIOTIME", TransferToHisAccounts.bmioTime)
      ' Մարման ամսաթիվ դաշտի լրացում
      Call Rekvizit_Fill("Document", 2, "General", "QDATE", TransferToHisAccounts.repaymentDate)
      ' Վերջնաժամկետ դաշտի լրացում
      Call Rekvizit_Fill("Document", 2, "General", "TRAILER", TransferToHisAccounts.wDeadline)
      ' Կարգ դաշտի լրացում
      Call Rekvizit_Fill("Document", 2, "General", "PRIOR", TransferToHisAccounts.wPrior)
      ' Բանկային առաջնություն դաշտի լրացում
      Call Rekvizit_Fill("Document", 2, "General", "BANKPRIOR", TransferToHisAccounts.BankingCham)
       
      ' Անցնել Ֆիզ. կազմակերպ. բաժին
       frmASDocForm.vbObject("TabStrip").SelectedItem =  frmASDocForm.vbObject("TabStrip").Tabs(3)
      
      wStatus = True
      BuiltIn.Delay(1500)
            
      If TransferToHisAccounts.clcikBOrNo Then
      
            wStatus = False
            frmASDocForm.VBObject("TabFrame_3").VBObject("AsTpComment_5").VBObject("CmdAdditional").Click
            If p1.WaitVBObject("frmAsUstPar",2000).Exists Then
                  ' Ուղարկող/Ստացող դաշտի լրացում
                  wStatus = FinancialOrganizationsFilter(TransferToHisAccounts.finOrginization(1))
            Else  
                  Log.Error("Ֆինանսական կազմակերպություններ դիալոգը չի բացվել")
            End If
            
       Else
           Call Rekvizit_Fill("Document", 1, "General", "RECINST", TransferToHisAccounts.sendRec)
       End If    
      
      If Not wStatus Then
            Call Rekvizit_Fill("Document", 3, "General", "SNDREC", TransferToHisAccounts.sendRec)
      End If
      
      ' Վճարող բանկի թղթակից հաշիվ դաշտի լրացում
      Call Rekvizit_Fill("Document", 3, "General", "PCORID", TransferToHisAccounts.CorBankAcc)
      ' Վճարող բանկի թղթակից դաշտի լրացում
      Call Rekvizit_Fill("Document", 3, "General", "PCORBANK", TransferToHisAccounts.CorBank)
      ' Միջնորդ բանկ տվ. տիպ դաշտի լրացում
      Call Rekvizit_Fill("Document", 3, "General", "MEDOP", TransferToHisAccounts.IntBankDataType)
      ' Միջնորդ բանկի հաշիվ դաշտի լրացում
      Call Rekvizit_Fill("Document", 3, "General", "MEDID", TransferToHisAccounts.IntBankAcc)
      
      wStatus = True
      BuiltIn.Delay(1500)
      
      If TransferToHisAccounts.clcikBOrNo Then
      
          wStatus = False
          frmASDocForm.VBObject("TabFrame_3").VBObject("AsTpComment_8").VBObject("CmdAdditional").Click
          If p1.WaitVBObject("frmAsUstPar",2000).Exists Then
                ' Միջնորդ բանկ դաշտի լրացում
                wStatus = FinancialOrganizationsFilter(TransferToHisAccounts.finOrginization( 2))
          Else  
                Log.Error("Ֆինանսական կազմակերպություններ դիալոգը չի բացվել")
          End If
      
      Else
           Call Rekvizit_Fill("Document", 1, "General", "RECINST", TransferToHisAccounts.IntBank)
      End If 
       
      If Not wStatus Then
            Call Rekvizit_Fill("Document", 3, "General", "", TransferToHisAccounts.IntBank)
      End If
      
      Call ClickCmdButton(1, "Î³ï³ñ»É")
End Sub


Function FinancialOrganizationsFilter(finOrg)
      
      Dim wStatus
      wStatus = False
      
      Log.Message("Ֆինանսական կազմակերպություններ դիալոգը բացվել է")
      ' Կոդ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "CODE", finOrg.wCode)
      ' Անվանում դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "NAME", finOrg.wName)
      ' Հասցե դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "ADDRESS", finOrg.wAddress)
      ' Երկիր դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "COUNTRY", finOrg.wCountry)
      ' Քաղաք դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "CITY", finOrg.wCity)
            
      Call ClickCmdButton(2, "Î³ï³ñ»É")
      
      BuiltIn.Delay(1500)
            
      If  Sys.Process("Asbank").VBObject("frmModalBrowser").VBObject("tdbgView").ApproxCount >= 1 Then
             Sys.Process("Asbank").VBObject("frmModalBrowser").VBObject("tdbgView").Keys("[Enter]")
             wStatus = True
      Else 
            Sys.Process("Asbank").VBObject("frmModalBrowser").Close
            Log.Error("Այդպիսի տվյալներով ֆինանսական կազմակերպություն չի գտնվել")
      End If

      FinancialOrganizationsFilter = wStatus
End Function

Class FinancialOrganizations
        Public wCode 
        Public wName
        Public wAddress 
        Public wCountry 
        Public wCity
    
        Private Sub Class_Initialize 
              wCode = ""
              wName= ""
              wAddress = ""
              wCountry = ""
              wCity = ""
        End Sub
End Class

Function New_FinancialOrganizations()
        Set New_FinancialOrganizations = NEW  FinancialOrganizations  
End Function 

' Մուտք SWIFT ԱՇՏ/ Փոխանցումներ/Ուղարկված փոխանցումներ թղթապանակ
Class OpenSentTransfersFolder
  
      Public folderDirect
      Public stDate
      Public endDate
      Public messType
      Public wState
      Public wUser
      Public wAddressee
      Public eRecipient
      Public messN
      Public shoePaySys

    Private Sub Class_Initialize
    
        folderDirect = ""
        stDate = ""
        endDate = ""
        messType = ""
        wState = ""
        wUser = ""
        wAddressee = ""
        eRecipient = ""
        messN = ""
        shoePaySys = 0
        
    End Sub  
End Class

Function New_OpenSentTransfersFolder()
    Set New_OpenSentTransfersFolder= NEW OpenSentTransfersFolder      
End Function

' Լրացնել "Ուղարկված փոխանցումներ" փաստաթղթի դաշտերը
Sub Fill_OpenSentTransfersFolder(OpenSentTransfersFolder)
      
      ' Մուտք թղթապանակ
      Call wTreeView.DblClickItem(OpenSentTransfersFolder.folderDirect)
      ' Ժամանակահատվածի սկիզբ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "PERN",  "^A[Del]" & OpenSentTransfersFolder.stDate)
      ' Ժամանակահատվածի ավարտ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "PERK",  "^A[Del]" & OpenSentTransfersFolder.endDate)
      ' Հաղ. տիպ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "MT", OpenSentTransfersFolder.messType)
      ' Կարգավիճակ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "STATCTL", OpenSentTransfersFolder.wState)
      ' Կատարող  դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "USER",  "^A[Del]" & OpenSentTransfersFolder.wUser)
      ' Հասցեատեր դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "TO",  "^A[Del]" & OpenSentTransfersFolder.wAddressee)
      ' Ստացող դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "SNDREC", OpenSentTransfersFolder.eRecipient)
      ' Հաղորդագրության միարժեք համար դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "UETR", OpenSentTransfersFolder.messN)
      ' Ցույց տալ ընդ. վճ. համակարգը դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "CheckBox", "SHOWPAYSYSIN", OpenSentTransfersFolder.shoePaySys)

      Call ClickCmdButton(2, "Î³ï³ñ»É")
End Sub



' "Միջբանկային փոխանցում ՀՏ202" փաստաթղթի ստեղծում
Class InterbankTransfer

      Public fISN
      Public docNum
      Public InterbankTransferGeneral
      Public InterbankTransferAdditional
      Public InterbankTransferFinancialOrg
      Public clcikBOrNo
      Public clcikBOrNo2
      Public clcikBOrNo3
      Public clcikBOrNo4
      Public clcikBOrNo5
      Public clcikBOrNo6
      Public clcikBOrNo7
      Public  FinOrganization(6) 
      
      Private Sub Class_Initialize
      
          fISN = ""
          docNum = ""
          Set InterbankTransferGeneral = New_InterbankTransferGeneral()
          Set InterbankTransferAdditional = New_InterbankTransferAdditional()
          Set InterbankTransferFinancialOrg = New_InterbankTransferFinancialOrg()
          
          clcikBOrNo = False
          clcikBOrNo2 = False
          clcikBOrNo3 = False
          clcikBOrNo4 = False
          clcikBOrNo5 = False
          clcikBOrNo6 = False
          clcikBOrNo7 = False

        For i = 0 to 6
            Set FinOrganization(i) = New_FinancialOrganizations()
        Next
        
      End Sub

End Class


Function New_InterbankTransfer()
    Set New_InterbankTransfer = NEW InterbankTransfer      
End Function



' "Միջբանկային Փոխանցում ՀՏ202 "փաստաթղթում /General - "/Ընդանուր" tab-ի Class
Class InterbankTransferGeneral
      Public FillTab
      Public acsBranch
      Public acsDepart
      Public messSingleVal
      Public serviceType
      Public MessType
      Public wReference
      Public wDate
      Public recOrgDataType
      Public expectedMessage
      Public expectedMessage2
      Public recOrgAcc
      Public wRecOrg
      Public recDataType
      Public recAcc
      Public wReceiver
      Public  wSumma
      Public wCur
      Public wTxKey
      
      Private Sub Class_Initialize
              FillTab = False
              acsBranch = ""
              acsDepart = ""
              messSingleVal = ""
              serviceType = ""
              MessType = ""
              wReference = ""
              wDate = ""
              recOrgDataType = ""
              expectedMessage = ""
              expectedMessage2 = ""
              recOrgAcc = ""
              wRecOrg = ""
              recDataType = ""
              recAcc = ""
              wReceiver = ""
              wSumma = ""
              wCur = ""
              wTxKey = ""
    End Sub  
End Class

Function New_InterbankTransferGeneral()
    Set New_InterbankTransferGeneral = NEW InterbankTransferGeneral      
End Function


' Լրացնել "Մջբանկային Փոխանցում ՀՏ202" փաստաթղթի "Ընդհանուր" բաժնի դաշտերը
Sub Fill_InterbankTransferGeneral(InterbankTransfer)
  
      Dim wStatus
      If InterbankTransfer.InterbankTransferGeneral.FillTab Then
      
            ' Գրասենյակ դաշտի լրացում
            Call Rekvizit_Fill("Document", 1, "General", "ACSBRANCH",  InterbankTransfer.InterbankTransferGeneral.acsBranch)
            ' Բաժին դաշտի լրացում
            Call Rekvizit_Fill("Document", 1, "General", "ACSDEPART",  InterbankTransfer.InterbankTransferGeneral.acsDepart)
            ' Հաղորդագրության միարժեք համար դաշտի լրացում
            Call Rekvizit_Fill("Document", 1, "General", "UETR",  InterbankTransfer.InterbankTransferGeneral.messSingleVal)
            ' Ծառայության տեսակ դաշտի լրացում
            Call Rekvizit_Fill("Document", 1, "General", "STID",  InterbankTransfer.InterbankTransferGeneral.serviceType)
            ' Հաղ տիպ դաշտի լրացում
            Call Rekvizit_Fill("Document", 1, "General", "MT",  InterbankTransfer.InterbankTransferGeneral.MessType)
            ' Հղում դաշտի լրացում
            Call Rekvizit_Fill("Document", 1, "General", "REF",  InterbankTransfer.InterbankTransferGeneral.wReference)
            ' Ամսաթիվ դաշտի լրացում
            Call Rekvizit_Fill("Document", 1, "General", "DATE",  InterbankTransfer.InterbankTransferGeneral.wDate)
            ' Ստացող կազմակերպ տվ. տիպ դաշտի լրացում
            Call Rekvizit_Fill("Document", 1, "General", "RINSTOP",  InterbankTransfer.InterbankTransferGeneral.recOrgDataType)
            ' Ստացող կազմակերպ. հաշիվ դաշտի լրացում
            Call Rekvizit_Fill("Document", 1, "General", "RINSTID",  InterbankTransfer.InterbankTransferGeneral.recOrgAcc)

            wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame").VBObject("AsTpComment").VBObject("CmdAdditional_2").Click
           
            wStatus =  MessageExists(2, InterbankTransfer.InterbankTransferGeneral.expectedMessage)
          
            If wStatus Then
                Log.Message("IBAN հաղորդագրությունը ճիշտ է")
                Call ClickCmdButton(5, "OK")
            End If
         
            ' Ստացող կազմակերպ. դաշտի լրացում
            wStatus = True
            If  InterbankTransfer.clcikBOrNo Then
     
                wStatus = False
                wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame").VBObject("AsTpComment_2").VBObject("CmdAdditional").Click
           
                If p1.WaitVBObject("frmAsUstPar",2000).Exists Then
                    ' Ֆինանսական կազմակերպություններ դիալոգի լրացում
                    wStatus = FinancialOrganizationsFilter(InterbankTransfer.FinOrganization(0))
                Else  
                    Log.Error("Ֆինանսական կազմակերպություններ դիալոգը չի բացվել")
                End If
            Else
                Call Rekvizit_Fill("Document", 1, "General", "RECINST",  InterbankTransfer.InterbankTransferGeneral.wRecOrg)
            End If     
      
            If Not wStatus Then
                Log.Message("Ստացող կազմակերպ. դաշտի արժեքը լրացվել է ինփութից")
                Call Rekvizit_Fill("Document", 1, "General", "RECINST",  InterbankTransfer.InterbankTransferGeneral.wRecOrg)
            End If
          
            ' Ստացողի տվ. տիպ դաշտի լրացում
            Call Rekvizit_Fill("Document", 1, "General", "RECOP",   InterbankTransfer.InterbankTransferGeneral.recDataType)
            ' Ստացողի հաշիվ դաշտի լրացում
            Call Rekvizit_Fill("Document", 1, "General", "ACCCR",   InterbankTransfer.InterbankTransferGeneral.recAcc)
            ' Սեղմել IBAN կոճակը
            wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame").VBObject("AsTpComment_3").VBObject("CmdAdditional_2").Click
           
            wStatus =  MessageExists(2, InterbankTransfer.InterbankTransferGeneral.expectedMessage2)
          
            If wStatus Then
                Log.Message("IBAN հաղորդագրությունը ճիշտ է")
                Call ClickCmdButton(5, "OK")
            End If
          
            wStatus = True
            If  InterbankTransfer.clcikBOrNo2 Then
            
                wStatus = False
                wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame").VBObject("AsTpComment_4").VBObject("CmdAdditional").Click
           
                If p1.WaitVBObject("frmAsUstPar",2000).Exists Then
                    ' Ֆինանսական կազմակերպություններ դիալոգի լրացում
                    wStatus = FinancialOrganizationsFilter(InterbankTransfer.FinOrganization(1))
                Else  
                    Log.Error("Ֆինանսական կազմակերպություններ դիալոգը չի բացվել")
                End If
                
            Else
                ' Ստացող դաշտի լրացում
                Call Rekvizit_Fill("Document", 1, "General", "RECEIVER",   InterbankTransfer.InterbankTransferGeneral.wReceiver)
            End If     
                
            If Not wStatus Then
                ' Ստացող դաշտի լրացում
                Log.Message("Ստացող դաշտի արժեքը լրացվել է ինփութից")
                Call Rekvizit_Fill("Document", 1, "General", "RECEIVER",   InterbankTransfer.InterbankTransferGeneral.wReceiver)
            End If
           
            ' Գումար դաշտի լրացում
            Call Rekvizit_Fill("Document", 1, "General", "SUMMA",   InterbankTransfer.InterbankTransferGeneral.wSumma)
            ' Արժույթ դաշտի լրացում
            Call Rekvizit_Fill("Document", 1, "General", "CUR",   InterbankTransfer.InterbankTransferGeneral.wCur)
            ' Բանալի դաշտի լրացում
            Call Rekvizit_Fill("Document", 1, "General", "TXKEY",   InterbankTransfer.InterbankTransferGeneral.wTxKey)
            
      End If
      
End Sub

' "Միջբանկային Փոխանցում ՀՏ202" փաստաթղթի "Լրացուցիչ" բաժնի դաշտերի կլասը
Class InterbankTransferAdditional
      Public FillTab
      Public packN
      Public addInfo
      Public wCode
      Public wValue
      Public wDeviation
      Public fileName
      Public directName
      Public dateSentRec
      Public timeSentRec
      Public repayRate
      Public wTrailer
      Public wPriority
      Public accServBankRef
      Public bankPriority
      
      Private Sub Class_Initialize
             FillTab = True
             packN = ""
             addInfo = ""
             wCode = ""
             wValue = ""
             wDeviation = ""
             fileName = ""
             directName = ""
             dateSentRec = ""
             timeSentRec = ""
             repayRate = ""
             wTrailer = ""
             wPriority = ""
             accServBankRef = ""
             bankPriority = ""
    End Sub  
End Class

Function New_InterbankTransferAdditional()
    Set New_InterbankTransferAdditional = NEW InterbankTransferAdditional      
End Function

' Լրացնել "Միջբանկային Փոխանցում ՀՏ202" փաստաթղթի "Լրացուցիչ" բաժնի դաշտերը
Sub Fill_InterbankTransferAdditional(AdditionalTab)
      
      If AdditionalTab.FillTab Then
            ' Փաթեթի համարը դաշտի լրացում
            Call Rekvizit_Fill("Document", 2, "General", "PACK",  AdditionalTab.packN)
            ' Լրացուցիչ ինֆորմացիա դաշտի լրացում
            Call Rekvizit_Fill("Document", 2, "General", "ADDINFO",  AdditionalTab.addInfo)
      
            With wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame_2").VBObject("DocGrid")
                  ' Կոդ դաշտի լրացում
                  .Row = 0
                  .Col = 0
                  .Keys(AdditionalTab.wCode & "[Enter]")
                  ' Արժեք դաշտի լրացում
                  .Col = 1
                  .Keys(AdditionalTab.wValue & "[Enter]" )
                  ' Շեղում դաշտի լրացում
                  .Col = 2
                  .Keys(AdditionalTab.wDeviation & "[Enter]" )
            End With 
      
            ' Ֆայլի անուն դաշտի լրացում
            Call Rekvizit_Fill("Document", 2, "General", "BMNAME",  AdditionalTab.fileName)
            ' Դիրեկտորայի անուն դաշտի լրացում
            Call Rekvizit_Fill("Document", 2, "General", "BMDIRECT",  AdditionalTab.directName)
            ' Ամսաթիվ (Ուղարկման/Ստացման) դաշտի լրացում
            Call Rekvizit_Fill("Document", 2, "General", "BMIODATE",  AdditionalTab.dateSentRec)
            ' Ժամանակ (Ուղարկման/Ստացման) դաշտի լրացում
            Call Rekvizit_Fill("Document", 2, "General", "BMIOTIME",  AdditionalTab.timeSentRec)
            ' Մարման ամսաթիվ դաշտի լրացում
            Call Rekvizit_Fill("Document", 2, "General", "QDATE",  AdditionalTab.repayRate)
            ' Վերջնահատված դաշտի լրացում
            Call Rekvizit_Fill("Document", 2, "General", "TRAILER",  AdditionalTab.wTrailer)
            ' Կարգ դաշտի լրացում
            Call Rekvizit_Fill("Document", 2, "General", "PRIOR",  AdditionalTab.wPriority)
            ' Հաշիվը ստացող բանկի հղում դաշտի լրացում
            Call Rekvizit_Fill("Document", 2, "General", "ABANKREF",  AdditionalTab.accServBankRef)
            ' Բանկային առաջնություն դաշտի լրացում
            Call Rekvizit_Fill("Document", 2, "General", "BANKPRIOR",  AdditionalTab.bankPriority)
      End If 
          
End Sub


' "Միջբանկային Փոխանցում ՀՏ202" փաստաթղթի "Ֆին. կազմակերպ." բաժնի դաշտերի կլասը
Class InterbankTransferFinancialOrg

      Public FillTab
      Public sendRec
      Public payOrgDataType
      Public payOrgAcc
      Public expectedMessage3
      Public expectedMessage4
      Public expectedMessage5
      Public expectedMessage6
      Public payingOrg
      Public payBankDataType
      Public payCorrAcc
      Public payBankCorr
      Public recBankDataType
      Public recCorrAcc
      Public recCorr
      Public medBankDataType
      Public medBankAcc
      Public medBank
      
      Private Sub Class_Initialize
            FillTab = True
            sendRec = ""
            payOrgDataType = ""
            payOrgAcc = ""
            expectedMessage3 = ""
            expectedMessage4 = ""
            expectedMessage5 = ""
            expectedMessage6 = ""
            payingOrg = ""
            payBankDataType = ""
            payCorrAcc = ""
            payBankCorr = ""
            recBankDataType = ""
            recCorrAcc = ""
            recCorr = ""
            medBankDataType = ""
            medBankAcc = ""
            medBank = ""
    End Sub  
    
End Class

Function New_InterbankTransferFinancialOrg()
    Set New_InterbankTransferFinancialOrg = NEW InterbankTransferFinancialOrg      
End Function

' Լրացնել "Միջբանկային Փոխանցում ՀՏ202" փաստաթղթի "Ֆին. կազմակերպ." բաժնի դաշտերը
Sub Fill_InterbankTransferFinancialOrg(InterbankTransfer)
      
      Dim wStatus, wValue
      If InterbankTransfer.InterbankTransferFinancialOrg.FillTab Then

             ' Անցնել Ֆին. կազմակերպ. բաժին
             wMDIClient.VBObject("frmASDocForm").VBObject("TabStrip").SelectedItem = wMDIClient.VBObject("frmASDocForm").VBObject("TabStrip").Tabs(3)
             
            ' Ուղարկող/Ստացող դաշտի լրացում
            wStatus = True
            If  InterbankTransfer.clcikBOrNo3 Then
     
                wStatus = False
                wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame_3").VBObject("AsTpComment_7").VBObject("CmdAdditional").Click
           
                If p1.WaitVBObject("frmAsUstPar",2000).Exists Then
                    ' Ֆինանսական կազմակերպություններ դիալոգի լրացում
                    wStatus = FinancialOrganizationsFilter(InterbankTransfer.FinOrganization(2))
                Else  
                    Log.Error("Ֆինանսական կազմակերպություններ դիալոգը չի բացվել")
                End If
            Else
                Call Rekvizit_Fill("Document", 3, "General", "SNDREC",  InterbankTransfer.InterbankTransferFinancialOrg.sendRec)
            End If     
      
            If Not wStatus Then
            Log.Message("Ուղարկող/Ստացող դաշտի արժեքը լրացվել է ինփութից")
                Call Rekvizit_Fill("Document", 3, "General", "SNDREC",  InterbankTransfer.InterbankTransferFinancialOrg.sendRec)
            End If
            
            ' Վճարող կազմակերպ. տվ. տիպ դաշտի լրացում
            Call Rekvizit_Fill("Document", 3, "General", "PINSTOP",  InterbankTransfer.InterbankTransferFinancialOrg.payOrgDataType & "[Tab]")
            ' Վճարող կազմակերպ. հաշիվ դաշտի լրացում
            Call Rekvizit_Fill("Document", 3, "General", "PINSTID",  InterbankTransfer.InterbankTransferFinancialOrg.payOrgAcc)
            
            ' IBAN կոճակի սեղմում և հաղորդագրության ստուգում
            wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame_3").VBObject("AsTpComment_8").VBObject("CmdAdditional_2").Click
           
            wStatus =  MessageExists(2, InterbankTransfer.InterbankTransferFinancialOrg.expectedMessage3)
          
            If wStatus Then
                Log.Message("IBAN հաղորդագրությունը ճիշտ է")
                Call ClickCmdButton(5, "OK")
            End If
            
            ' Վճարող կազմակերպ. դաշտի լրացում
            wStatus = True
            If  InterbankTransfer.clcikBOrNo4 Then
     
                wStatus = False
                wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame_3").VBObject("AsTpComment_9").VBObject("CmdAdditional").Click
           
                If p1.WaitVBObject("frmAsUstPar",2000).Exists Then
                    ' Ֆինանսական կազմակերպություններ դիալոգի լրացում
                    wStatus = FinancialOrganizationsFilter(InterbankTransfer.FinOrganization(3))
                Else  
                    Log.Error("Ֆինանսական կազմակերպություններ դիալոգը չի բացվել")
                End If
            Else
                Call Rekvizit_Fill("Document", 3, "General", "PAYINST",  InterbankTransfer.InterbankTransferFinancialOrg.payingOrg)
            End If     
      
            If Not wStatus Then
                Log.Message("Վճարող կազմակերպ. դաշտի արժեքը լրացվել է ինփութից")
                Call Rekvizit_Fill("Document", 3, "General", "PAYINST",  InterbankTransfer.InterbankTransferFinancialOrg.payingOrg)
            End If
            
            ' Վճարող բանկի թղթակից տվ. տիպ դաշտի լրացում
            Call Rekvizit_Fill("Document", 3, "General", "PCOROP",  InterbankTransfer.InterbankTransferFinancialOrg.payBankDataType)
            ' Վճարող բանկի թղթակից հաշիվ դաշտի լրացում
            Call Rekvizit_Fill("Document", 3, "General", "PCORID",  InterbankTransfer.InterbankTransferFinancialOrg.payCorrAcc)
            
            ' IBAN կոճակի սեղմում և հաղորդագրության ստուգում
            wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame_3").VBObject("AsTpComment_10").VBObject("CmdAdditional_2").Click
           
            wStatus =  MessageExists(2, InterbankTransfer.InterbankTransferFinancialOrg.expectedMessage4)
          
            If wStatus Then
                Log.Message("IBAN հաղորդագրությունը ճիշտ է")
                Call ClickCmdButton(5, "OK")
            End If
            
            ' Վճարող բանկի թղթակից դաշտի լրացում
            wStatus = True
            If  InterbankTransfer.clcikBOrNo5 Then
     
                wStatus = False
                wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame_3").VBObject("AsTpComment_11").VBObject("CmdAdditional").Click
           
                If p1.WaitVBObject("frmAsUstPar",2000).Exists Then
                    ' Ֆինանսական կազմակերպություններ դիալոգի լրացում
                    wStatus = FinancialOrganizationsFilter(InterbankTransfer.FinOrganization(4))
                Else  
                    Log.Error("Ֆինանսական կազմակերպություններ դիալոգը չի բացվել")
                End If
            Else
                  Call Rekvizit_Fill("Document", 3, "General", "PCORBANK",  InterbankTransfer.InterbankTransferFinancialOrg.payBankCorr)
            End If     
      
            If Not wStatus Then
                  Log.Message("Վճարող բանկի թղթակից դաշտի արժեքը լրացվել է ինփութից")
                  Call Rekvizit_Fill("Document", 3, "General", "PCORBANK",  InterbankTransfer.InterbankTransferFinancialOrg.payBankCorr)
            End If
            
            ' Ստացող բանկի թղթակից տվ. տիպ դաշտի լրացում
            Call Rekvizit_Fill("Document", 3, "General", "RCOROP",  InterbankTransfer.InterbankTransferFinancialOrg.recBankDataType)
            ' Ստացող բանկի թղթակից հաշիվ դաշտի լրացում
            Call Rekvizit_Fill("Document", 3, "General", "RCORID",  InterbankTransfer.InterbankTransferFinancialOrg.recCorrAcc)
            
            ' Ստացող բանկի թղթակից հաշիվ
            wValue = Get_Rekvizit_Value("Document",3,"Comment","RCORID")
            
            If Not wValue = "" Then
                  ' IBAN կոճակի սեղմում և հաղորդագրության ստուգում
                  wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame_3").VBObject("AsTpComment_12").VBObject("CmdAdditional_2").Click
           
                  wStatus =  MessageExists(2, InterbankTransfer.InterbankTransferFinancialOrg.expectedMessage5)
          
                  If wStatus Then
                      Log.Message("IBAN հաղորդագրությունը ճիշտ է")
                      Call ClickCmdButton(5, "OK")
                  End If
            End If
            
            ' Ստացող բանկի թղթակից դաշտի լրացում
            wStatus = True
            If  InterbankTransfer.clcikBOrNo6 Then
     
                wStatus = False
                wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame_3").VBObject("AsTpComment_13").VBObject("CmdAdditional").Click
           
                If p1.WaitVBObject("frmAsUstPar",2000).Exists Then
                      ' Ֆինանսական կազմակերպություններ դիալոգի լրացում
                      wStatus = FinancialOrganizationsFilter(InterbankTransfer.FinOrganization(5))
                Else  
                      Log.Error("Ֆինանսական կազմակերպություններ դիալոգը չի բացվել")
                End If
            Else
                  Call Rekvizit_Fill("Document", 3, "General", "RCORBANK",  InterbankTransfer.InterbankTransferFinancialOrg.recCorr)
            End If     
      
            If Not wStatus Then
            Log.Message("Ստացող բանկի թղթակից դաշտի արժեքը լրացվել է ինփութից")
                  Call Rekvizit_Fill("Document", 3, "General", "RCORBANK",  InterbankTransfer.InterbankTransferFinancialOrg.recCorr)
            End If
            
            ' Միջնորդ բանկի տվ. տիպ դաշտի լրացում
            Call Rekvizit_Fill("Document", 3, "General", "MEDOP",  InterbankTransfer.InterbankTransferFinancialOrg.medBankDataType)
            ' Միջնորդ բանկի հաշիվ դաշտի լրացում
            Call Rekvizit_Fill("Document", 3, "General", "MEDID",  InterbankTransfer.InterbankTransferFinancialOrg.medBankAcc)
            
            ' IBAN կոճակի սեղմում և հաղորդագրության ստուգում
            wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame_3").VBObject("AsTpComment_14").VBObject("CmdAdditional_2").Click
           
            wStatus =  MessageExists(2, InterbankTransfer.InterbankTransferFinancialOrg.expectedMessage6)
          
            If wStatus Then
                Log.Message("IBAN հաղորդագրությունը ճիշտ է")
                Call ClickCmdButton(5, "OK")
            End If
            
            ' Միջնորդ բանկ դաշտի լրացում
            wStatus = True
            If  InterbankTransfer.clcikBOrNo7 Then
     
                wStatus = False
                wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame_3").VBObject("AsTpComment_15").VBObject("CmdAdditional").Click
           
                If p1.WaitVBObject("frmAsUstPar",2000).Exists Then
                    ' Ֆինանսական կազմակերպություններ դիալոգի լրացում
                    wStatus = FinancialOrganizationsFilter(InterbankTransfer.FinOrganization(6))
                Else  
                    Log.Error("Ֆինանսական կազմակերպություններ դիալոգը չի բացվել")
                End If
            Else
                Call Rekvizit_Fill("Document", 3, "General", "MEDBANK",  InterbankTransfer.InterbankTransferFinancialOrg.medBank)
            End If     
      
            If Not wStatus Then
                 Log.Message("Միջնորդ բանկ դաշտի արժեքը լրացվել է ինփութից")
                 Call Rekvizit_Fill("Document", 3, "General", "MEDBANK",  InterbankTransfer.InterbankTransferFinancialOrg.medBank)
            End If
            
      End If 
      
End Sub


' "Միջբանկային փոխանցում հտ 202" փաստաթղթի ստեղծման ֆունկցիա
Function CreateInterbankTransfer(IntTransfer)

    ' ISN-ի վերագրում փոփոխականին
    IntTransfer.fISN = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
    
    ' Ստանալ Փաստաթղթի N-ը
    IntTransfer.docNum = Get_Rekvizit_Value("Document",1,"General","BMDOCNUM")
    
    ' Լրացնել "Ընդանուր" Tab-ի ռեկվիզիտները
    Call Fill_InterbankTransferGeneral(IntTransfer)
    
    ' Լրացնել "Լրացուցիչ" Tab-ի ռեկվիզիտները
    Call Fill_InterbankTransferAdditional(IntTransfer.InterbankTransferAdditional)
    
    ' Լրացնել "Ֆին. կազմակերպ" Tab-ի ռեկվիզիտները
    Call Fill_InterbankTransferFinancialOrg(IntTransfer)
    
    ' Սեղմել "Կատարել" կոճակը
    Call ClickCmdButton(1, "Î³ï³ñ»É")
End Function    