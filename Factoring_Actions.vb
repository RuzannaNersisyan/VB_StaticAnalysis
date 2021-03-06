Option Explicit

'USEUNIT Library_Common  
'USEUNIT Factoring_Library
'USEUNIT Payment_Order_ConfirmPhases_Library
'USEUNIT Akreditiv_Library
'USEUNIT Loan_Agreements_Library 
'USEUNIT Loan_Agreemnts_With_Schedule_Library
'USEUNIT Deposit_Contract_Library
'USEUNIT Overdraft_NewCases_Library
'USEUNIT Group_Operations_Library
'USEUNIT Overdraft_NewCases_Library
'USEUNIT Constants
'USEUNIT Mortgage_Library

'Test Case Id 165831
'Test Case Id 165832
'Test Case Id 165834

Sub Factoring_Actions_Test(DocumentType)
  Dim fDATE, sDATE, attr
  Dim Factoring, calcDate, FolderName, PerAcc, opDate, exTerm
  Dim MainSum, PerSum, Prc, NonUsedPrc, summa
    
    
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
  
  FolderName = "|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|ä³ÛÙ³Ý³·ñ»ñ|î»Õ³µ³ßËí³Í ÙÇçáóÝ»ñ|ØØÄä ü³ÏïáñÇÝ·|"
  
  Call Log.Message("Ֆակտորինգի պայմանագրի ստեղծում",,,attr)
  Set Factoring = New_FactoringDoc()
  With Factoring
    .PayerAcc = "03485010100"
    .LenderAcc = "00000113032"
    .Amount = 100000
    .Date = "220419" 
    .GiveDate = "220419"
    .Term = "220420"
    .DocLevel = 1
    .PaidAmount = 100000
    .PaperCode = 333
    
    Select Case DocumentType
        Case 1
          .DocType = "îáÏáë.»Ï.µ»ñáÕ ý³ÏïáñÇÝ·"
          .DocTypeNum = "2"
        Case 2
          .DocType = "ü³ÏïáñÇÝ·"
          .DocTypeNum = "5"
        Case 3
          .DocType = "¶ñ³ýÇÏáí ý³ÏïáñÇÝ·Ç å³ÛÙ³Ý³·Çñ"
          .DocTypeNum = "8"
    End Select
    
    Call .CreateFactoring(FolderName & "Üáñ å³ÛÙ³Ý³·ñÇ ëï»ÕÍáõÙ")
    Log.Message(.DocNum)
    
    If .DocTypeNum = "8" Then
      'Մարման գրաֆիկի նշանակում
      Call wMainForm.MainMenu.Click(c_AllActions)
      Call wMainForm.PopupMenu.Click(c_RepaySchedule) 
      BuiltIn.Delay(2000)         
      wMDIClient.VBObject("frmPttel").VBObject("tdbgView").MoveNext
    End If
    
    'Պայմանագրը ուղարկել հաստատման                               
    Call PaySys_Send_To_Verify()
    
      'Վերցնել "Տոկոսային եկամուտների հաշիվ" ռեկվիզիտի արժեքը`
      wMDIClient.VBObject("frmPttel").Refresh
      Call Find_Doc_By("ü³ÏïáñÇÝ·Ç Ñ³ßí³å³Ñ³Ï³Ý Ñ³í»Éí³Í", 32, 0, "")
      BuiltIn.Delay(1000)
      Call wMainForm.MainMenu.Click(c_AllActions)
      Call wMainForm.PopupMenu.Click(c_View)
      BuiltIn.Delay(1000)
  
      PerAcc = Get_Rekvizit_Value("Document",2,"Mask","ACCPERINC")
      BuiltIn.Delay(1000)
      wMDIClient.VBObject("frmASDocForm").Close
      BuiltIn.Delay(1000)
      Call Close_Pttel("frmPttel")
    
      'Մուտք գործել "Հաճախորդի սպասարկում և դրամարկղ" 
      Call ChangeWorkspace(c_CustomerService)
      Call wTreeView.DblClickItem("|Ð³×³Ëáñ¹Ç ëå³ë³ñÏáõÙ ¨ ¹ñ³Ù³ñÏÕ |Ð³ßÇíÝ»ñ")
      BuiltIn.Delay(2000)
      Call Rekvizit_Fill("Dialog", 1, "General", "ACCMASK", PerAcc)
      Call ClickCmdButton(2, "Î³ï³ñ»É") 
      BuiltIn.Delay(2000)
      'Փոխել Տոկոսային եկամուտների հաշիվի սոտորին սահմանը
      Call wMainForm.MainMenu.Click(c_AllActions)
      Call wMainForm.PopupMenu.Click(c_ChangeLowerBound)
      BuiltIn.Delay(2000)
      Call Rekvizit_Fill("Document", 1, "General", "CHGDATE", "220419")
      Call Rekvizit_Fill("Document", 1, "General", "LLIMIT", "-999999999")
      Call ClickCmdButton(1, "Î³ï³ñ»É") 
      BuiltIn.Delay(2000)
      Call Close_Pttel("frmPttel")
    
    Call ChangeWorkspace(c_Subsystems)
    Call LetterOfCredit_Filter_Fill(FolderName & "ä³ÛÙ³Ý³·ñ»ñ|", .DocLevel, .DocNum)
      
    Call Log.Message("Գանձում տրամադրումից",,,attr)
    Call Collect_From_Provision(.Date, summa, 2, .PayerAcc, Null)
    
    Call Log.Message("ՄՄԺՊ ֆակտորինգի տրամադրում",,,attr)
    Call GiveFactoring(.Date, 2, Null)
    
    Call Log.Message("Տոկոսների հաշվարկ",,,attr)
    calcDate = "210519"
    Call Calculate_Percents(calcDate, calcDate, False)
    
    opDate = "220519"
    exTerm = "220421"
    If .DocTypeNum = "8" Then
      Call Log.Message("Գրաֆիկի վերանայում",,,attr)
      Call Fading_Schedule_Fill(opDate, exTerm, .Amount)
    Else
      Call Log.Message("Ժամկետների վերանայում",,,attr)
      Call Deposit_Extension(opDate, exTerm, "", .Paragraph, .Direction, c_TermsStates & "|" & c_Dates & "|" & c_ReviewTerms)
    End If
    
    Call Log.Message("Պարտքերի մարում",,,attr)
    MainSum = 10000
    PerSum = Null
    Call Fade_Debt(opDate, Null, "", MainSum, PerSum, False)
  
    Call Log.Message("Արդյունավետ տոկոսադրույք",,,attr)
    Call ChangeEffRete(opDate, "", "")
    
    Call Log.Message("Բանկի արդյունավետ տոկոսադրույք",,,attr)
    Call BankEffective_InterestRate_DocFill(opDate, "")
    
    calcDate = "220519" 
    Call Log.Message("Տոկոսների հաշվարկ",,,attr)
    Call Calculate_Percents(calcDate, calcDate, False)
    
    Call Log.Message("Օբյեկտիվ ռիսկի դասիչ",,,attr)
    Call ObjectiveRisk(opDate, "04")
    
    Call Log.Message("Ռիսկի դասիչ և պահուստավորման տոկոս",,,attr)
    Call FillDoc_Risk_Classifier(opDate, "05", 100)
    
    Call Log.Message("Պահուստավորում",,,attr)
    Call FillDoc_Store(opDate, Null)
    
    Call Log.Message("Դուրս գրում",,,attr)
    Call FillDoc_WriteOut(opDate, Null)
    
    opDate = "230519"
    Call Log.Message("Դուրս գրածի վերականգնում",,,attr)
    Call WriteOffReconstruction(opDate, "", "")
    
    Call Log.Message("Պարտքերի մարում",,,attr)
    Call Fade_Debt(opDate, Null, exTerm, "", "", False)

    Call Log.Message("Պայմանագրի փակում",,,attr)
    .CloseDate = opDate
    BuiltIn.Delay(2000)
    .CloseAgr()
    
    Call Log.Message("Պայմանագրի բացում",,,attr)
    'Պայմանագրի բացում
    .OpenAgr()

    Call Log.Message("Բոլոր փաստաթղթերի ջնջում",,,attr)
    
    BuiltIn.Delay(2000)
    Call Close_Pttel("frmPttel")
    Call GroupDelete(FolderName & "ä³ÛÙ³Ý³·ñ»ñ|", 1, .DocNum, "^A[Del]", "^A[Del]", "^A[Del]")
    
    Call LetterOfCredit_Filter_Fill(FolderName & "ä³ÛÙ³Ý³·ñ»ñ|", 1, .DocNum)
  
    'Ռիսկի դասիչների ջնջում
    Call Delete_ViewEdit("^A[Del]", "^A[Del]", c_Risking & "|" & c_RisksPersRes)
    Call Delete_ViewEdit("^A[Del]", "^A[Del]", c_Risking & "|" & c_ObjRiskCat)
    
    'Ջնջել տոկոսների նշանակումները
    Call Delete_ViewEdit("^A[Del]", "^A[Del]", c_Percentages & "|" & c_Percentages)
    Call Delete_ViewEdit("^A[Del]", "^A[Del]", c_Percentages & "|" & c_EffRate)
    Call Delete_ViewEdit("^A[Del]", "^A[Del]", c_Percentages & "|" & c_BankEffRate)

    'Ջնջել Ֆակտորինգի պայմանագիրը  
    BuiltIn.Delay(1000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_Delete)
    Call ClickCmdButton(3, "²Ûá")
  
  End With  
  
  Call Close_AsBank()  
End Sub