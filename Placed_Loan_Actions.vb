Option Explicit

'USEUNIT Library_Common  
'USEUNIT Akreditiv_Library
'USEUNIT Loan_Agreements_Library 
'USEUNIT Credit_Line_Library
'USEUNIT Deposit_Contract_Library
'USEUNIT Loan_Agreements_With_Schedule_Linear_Library
'USEUNIT Loan_Agreemnts_With_Schedule_Library
'USEUNIT Overdraft_NewCases_Library
'USEUNIT Group_Operations_Library
'USEUNIT Constants
'USEUNIT Library_Colour
'USEUNIT Mortgage_Library

'Test case ID 165690
'Test case ID 165693
'Test case ID 165697
'Test case ID 165700

Sub Placed_Loan_Actions_Test(DocumentType)
  Dim fDATE, sDATE, attr
  Dim Loan, FolderName, opDate, calcDate, exTerm, MainSum, PerSum, Prc, NonUsedPrc,_
      EffRete, ActRete, sum
  
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
  
  ''3.Վարկային գիծ պայմանագրի ստեղծում
  FolderName = "|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|ä³ÛÙ³Ý³·ñ»ñ|î»Õ³µ³ßËí³Í ÙÇçáóÝ»ñ|ì³ñÏ»ñ (ï»Õ³µ³ßËí³Í)|"
  Set Loan = New_LoanDocument()
  With Loan
    .CalcAcc = "00000113032"                                    
    .Limit = 100000
    .Date = "221018" 
    .GiveDate = "221018"
    .Term = "221019"
    .FirstDate = "221018"
    .PaperCode = 555
    
    Select Case DocumentType
        Case 1
          .DocType = "ì³ñÏ³ÛÇÝ ·ÇÍ"
        Case 2
          .DocType = "ØÇ³Ý·³ÙÛ³ í³ñÏ"
        Case 3
          .DocType = "¶ñ³ýÇÏáí í³ñÏ³ÛÇÝ å³ÛÙ³Ý³·Çñ" 
        Case 4
          .DocType = "¶ñ³ýÇÏáí í³ñÏ³ÛÇÝ å³ÛÙ³Ý³·Çñ (·Í³ÛÇÝ)"  
          .CheckPayDates = 1
          .FillRoundPr = ""
          .NonUsedPercent = 0
          .Paragraph = ""
          .PayDates = 22
    End Select
    
    Call .CreatePlLoan(FolderName & "Üáñ å³ÛÙ³Ý³·ñÇ ëï»ÕÍáõÙ")
  
    Log.Message(.DocNum)

    Call Close_Pttel("frmPttel")
  
    'Պայմանագրին ուղղարկել հաստատման
    .SendToVerify(FolderName & "²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ")
    'Հաստատել
    .Verify(FolderName & "Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I")
  
    Call LetterOfCredit_Filter_Fill(FolderName & "ä³ÛÙ³Ý³·ñ»ñ|", .DocLevel, .DocNum)
  
    Call Log.Message("Գանձում տրամադրումից",,,attr)
    Call Collect_From_Provision(.Date, sum, 2, .CalcAcc, Null)
  
    Call Log.Message("Վարկի տրամադրում",,,attr)
    If Left(.DocType, 28) = "¶ñ³ýÇÏáí í³ñÏ³ÛÇÝ å³ÛÙ³Ý³·Çñ" Then
      Call Give_Credit(.Date, 100000, 2, .CalcAcc, Null)
    Else
      Call Give_Credit(.Date, 80000, 2, .CalcAcc, Null)
    End If  
  
    Call Log.Message("Տոկոսների հաշվարկ",,,attr)
    calcDate = "211118"
    Call Calculate_Percents(calcDate, calcDate, False)
    
    opDate = "221118"
    If .DocType = "ì³ñÏ³ÛÇÝ ·ÇÍ" Then
      Call Log.Message("Սահմանաչափի փոփոխում",,,attr)
      Call Change_Limit(opDate , 200000)
    
      Call Log.Message("Տոկոսների կապիտալացում",,,attr)
      Call Percent_Capitalization(Null , opDate, "")
    End If
    
    exTerm = "221020"
    If .DocType = "¶ñ³ýÇÏáí í³ñÏ³ÛÇÝ å³ÛÙ³Ý³·Çñ" Then
      Call Log.Message("Գրաֆիկի վերանայում",,,attr)
      Call Fading_Schedule_Fill(opDate, exTerm, .Limit)
    ElseIf .DocType <> "¶ñ³ýÇÏáí í³ñÏ³ÛÇÝ å³ÛÙ³Ý³·Çñ (·Í³ÛÇÝ)" Then
      Call Log.Message("Ժամկետների վերանայում",,,attr)
      Call Deposit_Extension(opDate, exTerm, "", .Paragraph, .Direction, c_TermsStates & "|" & c_Dates & "|" & c_ReviewTerms)
    End If
    Call Close_Pttel("frmPttel")
    
    If .DocType <> "¶ñ³ýÇÏáí í³ñÏ³ÛÇÝ å³ÛÙ³Ý³·Çñ (·Í³ÛÇÝ)" Then
      ''Հաստատել ժամկետների վերանայումը
      .Verify(FolderName & "Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I")
    End If
    Call LetterOfCredit_Filter_Fill(FolderName & "ä³ÛÙ³Ý³·ñ»ñ|", .DocLevel, .DocNum)
    
    Call Log.Message("Պարտքերի մարում",,,attr)
    MainSum = 10000
    PerSum = Null
    Call Fade_Debt(opDate, Null, "", MainSum, PerSum, False)
  
    Call Log.Message("Տոկոսադրույքներ",,,attr)
    Prc = 15
    NonUsedPrc = 10
    Call ChangeRete(opDate, Prc, NonUsedPrc)
    
    Call Log.Message("Արդյունավետ տոկոսադրույք",,,attr)
    Call ChangeEffRete(opDate, EffRete, ActRete)
    
    If .DocType = "ØÇ³Ý·³ÙÛ³ í³ñÏ" or .DocType = "¶ñ³ýÇÏáí í³ñÏ³ÛÇÝ å³ÛÙ³Ý³·Çñ" Then
      Call Log.Message("Բանկի արդյունավետ տոկոսադրույք",,,attr)
      Call BankEffective_InterestRate_DocFill(opDate, "")
      If Left(.DocType, 28) = "¶ñ³ýÇÏáí í³ñÏ³ÛÇÝ å³ÛÙ³Ý³·Çñ" Then
        Call Log.Message("Սուբսիդավորման տոկոսադրույք",,,attr)
        Call SubsidyRate_DocFill(opDate, "^A[Del]" & 5)
       'Ջնջել Սուբսիդավորման տոկոսադրույքը
        Call Delete_ViewEdit("^A[Del]", "^A[Del]", c_Percentages & "|" & c_SubsidyRate)
      End If
    ElseIf .DocType = "ì³ñÏ³ÛÇÝ ·ÇÍ" Then
      Call Log.Message("Գծայնության դադարեցում",,,attr)
      Call Credit_Line_Stop_Recovery_DocFill(opDate, 1)
    End If  
    
    calcDate = "221118" 
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
    
    opDate = "231118"
    Call Log.Message("Դուրս գրածի վերականգնում",,,attr)
    Call WriteOffReconstruction(opDate, "", "")
    
    Call Log.Message("Պարտքերի մարում",,,attr)
    Call Fade_Debt(opDate, Null, "221020", "", "", False)

    Call Log.Message("Պայմանագրի փակում",,,attr)
    .CloseDate = opDate
    .CloseAgr()
  
    'Պայմանագրի բացում
    .OpenAgr()
    Call Close_Pttel("frmPttel")

    Call Log.Message("Բոլոր փաստաթղթերի ջնջում",,,attr)
    
    Call DeleteAllActions("|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|ä³ÛÙ³Ý³·ñ»ñ|î»Õ³µ³ßËí³Í ÙÇçáóÝ»ñ|ì³ñÏ»ñ (ï»Õ³µ³ßËí³Í)",.DocNum,"010118","010121")

  End With
  
  Call Close_AsBank()     
End Sub
