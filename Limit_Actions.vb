Option Explicit

'USEUNIT Library_Common  
'USEUNIT Limit_Library
'USEUNIT Loan_Agreements_With_Schedule_Linear_Library
'USEUNIT Akreditiv_Library
'USEUNIT Derivative_Tools_Library
'USEUNIT Loan_Agreements_Library 
'USEUNIT Loan_Agreemnts_With_Schedule_Library
'USEUNIT Group_Operations_Library
'USEUNIT Constants
'USEUNIT Mortgage_Library

'Test Case ID 165830

Sub Limit_Actions()
  Dim fDATE, sDATE, attr
  Dim LimitDoc, FolderName, opDate, calcDate, exTerm, MainSum, PerSum, Prc, NonUsedPrc,_
      EffRete, ActRete, Sum
  Dim Loan    
  
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

  ''3.Սահմանաչափի պայմանագրի ստեղծում
  FolderName = "|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|ä³ÛÙ³Ý³·ñ»ñ|î»Õ³µ³ßËí³Í ÙÇçáóÝ»ñ|ê³ÑÙ³Ý³ã³÷»ñ|"
  Set LimitDoc = New_LimitDocument()
  With LimitDoc
    .Client = "00000001"
    .Limit = 1000000
    .Date = "221019" 
    .GiveDate = "221019"
    .Term = "221020"
        
    Call .CreateLimit(FolderName & "Üáñ å³ÛÙ³Ý³·ñÇ ëï»ÕÍáõÙ")
  
    Log.Message(.DocNum)

    'Պայմանագրին ուղղարկել հաստատման
    .SendToVerify(Null)
    'Հաստատել
    .Verify(FolderName & "Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I")
  
    .OpenInFolder(FolderName)

    Call Log.Message("Գծայնության վերականգնում",,,attr)
    Call Credit_Line_Stop_Recovery_DocFill(.Date, 2)
    
    Call Log.Message("Նոր պայմանագրի ստեղծում` Գրաֆիով վարկային պայմանագիր",,,attr)

    BuiltIn.Delay(2000)
    Call Close_Pttel("frmPttel")
    Set Loan = New_LoanDocument()
    With Loan
      .CalcAcc = "77786271031"                                    
      .Limit = 1000000
      .Date = "221019" 
      .GiveDate = "221019"
      .Term = "221020"
      .FirstDate = "221019"
      .PaperCode = 555
      .DocType = "¶ñ³ýÇÏáí í³ñÏ³ÛÇÝ å³ÛÙ³Ý³·Çñ" 
        
      Call .CreatePlLoan("|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|ä³ÛÙ³Ý³·ñ»ñ|î»Õ³µ³ßËí³Í ÙÇçáóÝ»ñ|ì³ñÏ»ñ (ï»Õ³µ³ßËí³Í)|Üáñ å³ÛÙ³Ý³·ñÇ ëï»ÕÍáõÙ")
  
      Log.Message(.DocNum)

      'Պայմանագրին ուղղարկել հաստատման
      .SendToVerify(Null)
      'Հաստատել
      .Verify("|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|ä³ÛÙ³Ý³·ñ»ñ|î»Õ³µ³ßËí³Í ÙÇçáóÝ»ñ|ì³ñÏ»ñ (ï»Õ³µ³ßËí³Í)|Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I")
  
      Call LetterOfCredit_Filter_Fill("|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|ä³ÛÙ³Ý³·ñ»ñ|î»Õ³µ³ßËí³Í ÙÇçáóÝ»ñ|ì³ñÏ»ñ (ï»Õ³µ³ßËí³Í)|ä³ÛÙ³Ý³·ñ»ñ|", .DocLevel, .DocNum)
  
      Call Log.Message("Գանձում տրամադրումից",,,attr)
      Call Collect_From_Provision(.Date, Sum, 2, .CalcAcc, Null)
  
      Call Log.Message("Վարկի տրամադրում",,,attr)
      Call Give_Credit(.Date, .Limit, 2, .CalcAcc, Null)
      
      Call Log.Message("Տոկոսների հաշվարկ",,,attr)
      Call Calculate_Percents(.Date, .Date, False)
      
      opDate = "231019"
      Call Log.Message("Վարկի պարտքերի մարում",,,attr)
      Call Fade_Debt(opDate, Null, "221021", "", "", False)
      
      Call Log.Message("Վարկի պայմանագրի փակում",,,attr)
      BuiltIn.Delay(2000)
      .CloseDate = opDate
      .CloseAgr()
    End With

    BuiltIn.Delay(2000)
    Call Close_Pttel("frmPttel")
    BuiltIn.Delay(2000)
    
    .OpenInFolder(FolderName)
    
    exTerm = "221021"
    opDate = "231019"
    Call Log.Message("Ժամկետների վերանայում",,,attr)
    Call ReviewTerms(opDate, exTerm, 1)
    
    Call Log.Message("Արդյունավետ տոկոսադրույք",,,attr)
    Call ChangeEffRete(opDate, EffRete, ActRete)
    
    Call Log.Message("Օբյեկտիվ ռիսկի դասիչ",,,attr)
    Call ObjectiveRisk(opDate, "04")
    
    Call Log.Message("Ռիսկի դասիչ և պահուստավորման տոկոս",,,attr)
    Call FillDoc_Risk_Classifier(opDate, "05", 100)

    Call Log.Message("Չօգտ.մասի պահուստավորում",,,attr) 
    Call Compl_Actions_Reservation(1, opDate, 100000)
    .OpenInFolder(FolderName)
    
    'Լրացնել "Մարման աղբյուր" դաշտը
    BuiltIn.Delay(2000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_ToEdit)
    Call Rekvizit_Fill("Document", 9, "General", "REPSOURCE", 1) 
    Call ClickCmdButton(1, "Î³ï³ñ»É")
    
    Call Log.Message("Պայմանագրի փակում",,,attr)
    .CloseDate = opDate
    .CloseAgr()
  
    ''Ջնջել բոլոր պայմանագրերը
  
    'Պայմանագրի բացում
    .OpenAgr()

    Call Log.Message("Բոլոր փաստաթղթերի ջնջում",,,attr)
  
    BuiltIn.Delay(2000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_OpersView)
  
    Call Rekvizit_Fill("Dialog", 1, "General", "START", "^A[Del]" )
    Call Rekvizit_Fill("Dialog", 1, "General", "END", "^A[Del]" )
    Call Rekvizit_Fill("Dialog", 1, "General", "DEALTYPE", "^A[Del]")
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    
    BuiltIn.Delay(2000)
    Call wMainForm.MainMenu.Click(c_Opers & "|" & c_Delete)
    Call ClickCmdButton(3, "²Ûá")
    wMDIClient.VBObject("frmPttel_2").Close
    
    'Ռիսկի դասիչների ջնջում
    Call Delete_ViewEdit("^A[Del]", "^A[Del]", c_Risking & "|" & c_RisksPersRes)
    Call Delete_ViewEdit("^A[Del]", "^A[Del]", c_Risking & "|" & c_ObjRiskCat)
    
    'Ջնջել տոկոսների նշանակումները
    Call Delete_ViewEdit("^A[Del]", "^A[Del]", c_Percentages & "|" & c_EffRate)

   'Ջնջում "Սահմանաչափերից"
    Call Delete_ViewEdit("^A[Del]", "^A[Del]", c_Other & "|" & c_Limits)
    
    BuiltIn.Delay(2000)
    Call Close_Pttel("frmPttel")
    
    'Պայմանագրի բացում
    Call LetterOfCredit_Filter_Fill("|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|ä³ÛÙ³Ý³·ñ»ñ|î»Õ³µ³ßËí³Í ÙÇçáóÝ»ñ|ì³ñÏ»ñ (ï»Õ³µ³ßËí³Í)|ä³ÛÙ³Ý³·ñ»ñ|", Loan.DocLevel, Loan.DocNum)
    Loan.OpenAgr()
    Call Close_Pttel("frmPttel")
    Call GroupDelete("|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|ä³ÛÙ³Ý³·ñ»ñ|î»Õ³µ³ßËí³Í ÙÇçáóÝ»ñ|ì³ñÏ»ñ (ï»Õ³µ³ßËí³Í)|ä³ÛÙ³Ý³·ñ»ñ|", 1, Loan.DocNum, "^A[Del]", "^A[Del]", "^A[Del]")
    
    'Ջնջել Վարկի պայմանագիրը  
    Call LetterOfCredit_Filter_Fill("|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|ä³ÛÙ³Ý³·ñ»ñ|î»Õ³µ³ßËí³Í ÙÇçáóÝ»ñ|ì³ñÏ»ñ (ï»Õ³µ³ßËí³Í)|ä³ÛÙ³Ý³·ñ»ñ|", Loan.DocLevel, Loan.DocNum)
    
    BuiltIn.Delay(2000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_Delete)
    Call ClickCmdButton(3, "²Ûá")
    
    BuiltIn.Delay(2000)
    Call Close_Pttel("frmPttel")
    .OpenInFolder(FolderName)
    
    'Ջնջել Սահմանաչափի պայմանագիրը  
    BuiltIn.Delay(2000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_Delete)
    Call ClickCmdButton(3, "²Ûá")
    
    Call Close_AsBank()   
  End With  
End Sub