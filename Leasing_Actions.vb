Option Explicit

'USEUNIT Library_Common 
'USEUNIT Financial_Leasing_Library 
'USEUNIT Akreditiv_Library
'USEUNIT Loan_Agreements_Library 
'USEUNIT Loan_Agreemnts_With_Schedule_Library
'USEUNIT Overdraft_NewCases_Library
'USEUNIT Group_Operations_Library
'USEUNIT Constants
'USEUNIT Mortgage_Library

'Test case ID 165764

Sub Leasing_Actions_Test(DocumentType)
  Dim fDATE, sDATE, attr
  Dim Loan, FolderName, opDate, calcDate, debtDate, exTerm, MainSum
  Dim Leasing,PerSum, Prc, NonUsedPrc, EffRete, ActRete, summ
  
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
  
  ''2.Մուտք գործել "Ենթահամակրգեր(ՀԾ)"
  Call ChangeWorkspace(c_Subsystems)  

  ''3.Գրաֆիկով լիզինգի պայմանագրի ստեղծում
  FolderName = "|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|ä³ÛÙ³Ý³·ñ»ñ|î»Õ³µ³ßËí³Í ÙÇçáóÝ»ñ|üÇÝ³Ýë³Ï³Ý ÉÇ½ÇÝ· (ï»Õ³µ³ßËí³Í)|"
  
  Set Leasing = New_LeasingDoc()
  With Leasing
    .CalcAcc = "00000113032"
    .Date = "221018"
    .GiveDate = "221018"
    .StartDate = "221018"
    .Summa = 50000
    .BuyPrice = 10000
    .PaperCode = 111
    .Term = "221019"
    .DatesFillType = 1
    .office = "00"
    .department = "1"

    Select Case DocumentType
        Case 1 
          .DocType = "¶ñ³ýÇÏáí ÉÇ½ÇÝ·Ç å³ÛÙ³Ý³·Çñ"
        Case 2  
          .DocType = "ÈÇ½ÇÝ·"
          .LastDate = .Term
    End Select
  
    Call .CreateLeasing(FolderName & "Üáñ å³ÛÙ³Ý³·ñÇ ëï»ÕÍáõÙ")
    
    Log.Message(.DocNum)
  
    If .DocType = "¶ñ³ýÇÏáí ÉÇ½ÇÝ·Ç å³ÛÙ³Ý³·Çñ" Then
      'Մարման գրաֆիկի նշանակում
      BuiltIn.Delay(2000)
      Call wMainForm.MainMenu.Click(c_AllActions)
      Call wMainForm.PopupMenu.Click(c_RepaySchedule) 
      BuiltIn.Delay(2000)  
    End If  
  
    Call Close_Pttel("frmPttel")
  
    'Պայմանագիրը ուղարկել հաստատման
    .SendToVerify(FolderName & "²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ")
    'Վավերացնել պայմանագիրը
    .Verify(FolderName & "Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I")
  
    Call LetterOfCredit_Filter_Fill(FolderName, .DocLevel, .DocNum)
  
    opDate = "221018"
    Call Log.Message("Գանձում տրամադրումից",,,attr)
    Call Collect_From_Provision(opDate, summ, 2, .CalcAcc, Null)
  
    Call Log.Message("Լիզինգի տրամադրում",,,attr)
    Call Give_Leasing(opDate)
  
    Call Log.Message("Տոկոսների հաշվարկ",,,attr)
    opDate = "141118"
    Call Calculate_Percents(opDate, opDate, False)

    exTerm = "221020"
    opDate = "151118"
    If .DocType = "¶ñ³ýÇÏáí ÉÇ½ÇÝ·Ç å³ÛÙ³Ý³·Çñ" Then
      Call Log.Message("Գրաֆիկի վերանայում",,,attr)
      Call Fading_Schedule_Fill(opDate, exTerm, .Summa)
      
      Call Log.Message("Պարտքերի մարում",,,attr)
      MainSum = 10000
      PerSum = Null
      Call Fade_Debt(opDate, Null, "", MainSum, PerSum, False)
    Else 
      Call Log.Message("Պարտքերի մարում",,,attr)
      debtDate = "221118"
      Call Leasing_Fade_Debt(Null, opDate, debtDate, perSum, 2, "", .DocNum)  
    End If
      
    Call Log.Message("Տոկոսադրույքներ",,,attr)
    Prc = 15
    NonUsedPrc = 10
    Call ChangeRete(opDate, Prc, NonUsedPrc)
    
    Call Log.Message("Արդյունավետ տոկոսադրույք",,,attr)
    Call ChangeEffRete(opDate, EffRete, ActRete)
    
    If .DocType = "¶ñ³ýÇÏáí ÉÇ½ÇÝ·Ç å³ÛÙ³Ý³·Çñ" Then
      Call Log.Message("Սուբսիդավորման տոկոսադրույք",,,attr)
      Call SubsidyRate_DocFill(opDate, "^A[Del]" & 5)
     'Ջնջել Սուբսիդավորման տոկոսադրույքը
      Call Delete_ViewEdit("^A[Del]", "^A[Del]", c_Percentages & "|" & c_SubsidyRate)
    End If 
    
    Call Log.Message("Տոկոսների հաշվարկ",,,attr)
    Call Calculate_Percents(opDate, opDate, False)
      
    Call Log.Message("Օբյեկտիվ ռիսկի դասիչ",,,attr)
    Call ObjectiveRisk(opDate, "04")
    
    Call Log.Message("Ռիսկի դասիչ և պահուստավորման տոկոս",,,attr)
    Call FillDoc_Risk_Classifier(opDate, "05", 100)
    
    Call Log.Message("Պահուստավորում",,,attr)
    If .DocType = "¶ñ³ýÇÏáí ÉÇ½ÇÝ·Ç å³ÛÙ³Ý³·Çñ" Then
      Call FillDoc_Store(opDate, Null)
    Else  
      Call wMainForm.MainMenu.Click(c_AllActions)
      Call wMainForm.PopupMenu.Click(c_Opers & "|" & c_Store & "|" & c_Store)
      Call Rekvizit_Fill("Document", 1, "General", "DATE", "^A[Del]" & opDate)
      With wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame").VBObject("DocGrid")
        .Row = 0
        .Col = 1 
        .Keys("10")  
      End With 
      Call ClickCmdButton(1, "Î³ï³ñ»É") 
    End If
    
    Call Log.Message("Դուրս գրում",,,attr)
    Call FillDoc_WriteOut(opDate, Null)
    
    opDate = "161118"
    Call Log.Message("Դուրս գրածի վերականգնում",,,attr)
    Call WriteOffReconstruction(opDate, "", "")
    
    Call Log.Message("Պարտքերի մարում",,,attr)
    If .DocType = "¶ñ³ýÇÏáí ÉÇ½ÇÝ·Ç å³ÛÙ³Ý³·Çñ" Then
      Call Fade_Debt(opDate, Null, "221020", 40000, 394.5, False)
    Else 
      debtDate = "221019"
      Call Leasing_Fade_Debt(Null, opDate, debtDate, perSum, 2, "", .DocNum)  
    End If
    
    Call Log.Message("Պայմանագրի փակում",,,attr)
    .CloseDate = opDate
    .CloseAgr()
  
    Call Log.Message("Պայմանագրի բացում",,,attr)
    'Պայմանագրի բացում
    .OpenAgr()
  
    Call Log.Message("Բոլոր փաստաթղթերի ջնջում",,,attr)
    
    'Ջնջում "Գործողությունների դիտումից"
    Call Close_Pttel("frmPttel")
    Call GroupDelete(FolderName, 1, .DocNum, "^A[Del]", "^A[Del]", "")
    Call LetterOfCredit_Filter_Fill(FolderName, 1, .DocNum)
  
    'Ռիսկի դասիչների ջնջում
    Call Delete_ViewEdit("^A[Del]", "^A[Del]", c_Risking & "|" & c_RisksPersRes)
    Call Delete_ViewEdit("^A[Del]", "^A[Del]", c_Risking & "|" & c_ObjRiskCat)
    
    'Ջնջել տոկոսների նշանակումները
    Call Delete_ViewEdit("^A[Del]", "^A[Del]", c_Percentages & "|" & c_Percentages)
    Call Delete_ViewEdit("^A[Del]", "^A[Del]", c_Percentages & "|" & c_EffRate)

    'Ջնջում "Գործողությունների դիտումից"
    Call Close_Pttel("frmPttel")
    Call GroupDelete(FolderName, 1, .DocNum, "^A[Del]", "^A[Del]", "^A[Del]")
    Call LetterOfCredit_Filter_Fill(FolderName, 1, .DocNum)
  End With
  
  'Ջնջել Լիզինգի պայմանագիրը  
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_Delete)
  BuiltIn.Delay(2000)
  Call ClickCmdButton(3, "²Ûá")

  Call Close_AsBank()    
End Sub