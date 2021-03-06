Option Explicit

'USEUNIT Library_Common  
'USEUNIT Akreditiv_Library 
'USEUNIT Pledge_Library
'USEUNIT Group_Operations_Library
'USEUNIT Constants
'USEUNIT Mortgage_Library

'Test Case Id 165783
'Test Case Id 165786
'Test Case Id 165787
'Test Case Id 165790
'Test Case Id 165791
'Test Case Id 165792

Sub Pledge_Actions_Test(DocumentType)
  Dim fDATE, sDATE, attr, frmAsMsgBox, FrmSpr
  Dim Pledge, CollectFromProvision_ISN, GiveCredit_ISN, PercentCapISN, Repay_ISN, Store_ISN,_
      WriteOut_ISN
  Dim Overdraft, FolderName, opDate, Sum, opPerSum, calcDate, exTerm, MainSum, PerSum, Prc,_
      NonUsedPrc, EffRete, ActRete, Typ
  
'--------------------------------------
  Set attr = Log.CreateNewAttributes
  attr.BackColor = RGB(0, 255, 255)
  attr.Bold = True
  attr.Italic = True
'--------------------------------------  

  ''1.Համակարգ մուտք գործել ARMSOFT օգտագործողով
  fDATE = "20260101"
  sDATE = "20140101"
  Call Initialize_AsBank("bank", sDATE, fDATE)
  Login("ARMSOFT")
  
  'Գրավի ստեղծում
  Set Pledge = New_PledgeDoc()
  With Pledge
    .Date = "221018" 
    .GiveDate = "221018"
    .Client = "00000001"
    .Value = 100000
    .Count = 1
    
    Select Case DocumentType
        Case 1
          Call ChangeWorkspace(c_GivenPledge)
          FolderName = "|îñ³Ù³¹ñí³Í ·ñ³í|"
          .PledgeKind = "|îñ³Ù³¹ñí³Í ·ñ³í|"
        Case 2   
          Call ChangeWorkspace(c_RecPledge)
          FolderName = "|êï³óí³Í ·ñ³í|"
          .PledgeKind = "|êï³óí³Í ·ñ³í|"
        Case 3
          Call ChangeWorkspace(c_GivenDepPledge)
          FolderName = "|îñ³Ù³¹ñí³Í ³í³Ý¹³ÛÇÝ ·ñ³í|"
          .PledgeKind = "|îñ³Ù³¹ñí³Í ³í³Ý¹³ÛÇÝ ·ñ³í|"  
          .DocNum1 = "V-000281"
        Case 4
          Call ChangeWorkspace(c_RecDepPledge)
          FolderName = "|êï³óí³Í ³í³Ý¹³ÛÇÝ ·ñ³í|"
          .PledgeKind = "|êï³óí³Í ³í³Ý¹³ÛÇÝ ·ñ³í|"  
          .DocNum1 = "11"  
        Case 5
          Call ChangeWorkspace(c_GivenGuarantee)
          FolderName = "|îñ³Ù³¹ñí³Í »ñ³ßË³íáñáõÃÛáõÝ|"
          .PledgeKind = "|îñ³Ù³¹ñí³Í »ñ³ßË³íáñáõÃÛáõÝ|"  
          .PledgeType = 2
        Case 6
          Call ChangeWorkspace(c_RecGuarantee)
          FolderName = "|êï³óí³Í »ñ³ßË³íáñáõÃÛáõÝ|"
          .PledgeKind = "|êï³óí³Í »ñ³ßË³íáñáõÃÛáõÝ|"  
          .PledgeType = 2  
    End Select
    
    If Right(.PledgeKind, 17) = "»ñ³ßË³íáñáõÃÛáõÝ|" Then
      Call .CreateGuarantee(FolderName & "Üáñ å³ÛÙ³Ý³·ñÇ ëï»ÕÍáõÙ")
    Else 
      Call .CreatePledge(FolderName & "Üáñ å³ÛÙ³Ý³·ñÇ ëï»ÕÍáõÙ")
    End If  
    Log.Message(.DocNum)
    
    'Պայմանագրին ուղղարկել հաստատման
    .SendToVerify(Null)
    'Հաստատել
    .Verify(FolderName & "Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I")
  
    .OpenInFolder(FolderName)
    
    If .PledgeKind = "|îñ³Ù³¹ñí³Í »ñ³ßË³íáñáõÃÛáõÝ|" Then
      Typ = 2
    Else
      Typ = 1
    End If
    
    Call Log.Message("Գրավի տրամադրում",,,attr)
    Call GiveReturnPledge(.Date, c_Give, Typ)
    
    Call Log.Message("Ճշգրտում/վերագնահատում",,,attr)
    If Right(.PledgeKind, 15) = "³í³Ý¹³ÛÇÝ ·ñ³í|" Then
      Call RevaluationDepPledge(.Date, 100000, "2")
    Else
      Call Revaluation(.Date, 200000, 2, Typ)  
    End If
    
    Call Log.Message("Գրավի վերադարձ",,,attr)
    Call GiveReturnPledge(.Date, c_Return, Typ)
    
    Call Log.Message("Պայմանագրի փակում",,,attr)
    .CloseDate = .Date
    .CloseAgr()
    
    Call Log.Message("Պայմանագրի բացում",,,attr)
    .OpenAgr()
  
    Call Log.Message("Բոլոր փաստաթղթերի ջնջում",,,attr)

    BuiltIn.Delay(2000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_OpersView)
    BuiltIn.Delay(2000)
  
    Call Rekvizit_Fill("Dialog", 1, "General", "START", "^A[Del]" & "[Tab]")
    Call Rekvizit_Fill("Dialog", 1, "General", "END", "^A[Del]" & "[Tab]")
    Call Rekvizit_Fill("Dialog", 1, "General", "DEALTYPE", "^A[Del]" &"[Tab]")
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    BuiltIn.Delay(2000)
    
    wMDIClient.VBObject("frmPttel_2").VBObject("tdbgView").MoveLast
    While wMDIClient.VBObject("frmPttel_2").VBObject("tdbgView").ApproxCount <> 0
      Call wMainForm.MainMenu.Click(c_Opers & "|" & c_Delete)
      BuiltIn.Delay(1000)
      Call ClickCmdButton(3, "²Ûá")
    Wend
    Call Close_Pttel("frmPttel_2")
    
    Call wMainForm.MainMenu.Click(c_Opers & "|" & c_Delete)
    BuiltIn.Delay(1000)
    Call ClickCmdButton(3, "²Ûá")
    
    Call Close_AsBank()
  End With  
End Sub