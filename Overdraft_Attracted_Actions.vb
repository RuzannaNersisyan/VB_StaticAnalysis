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
'USEUNIT Mortgage_Library

'Test Case Id 165765
'Test Case Id 165767

Sub Overdraft_Attracted_Actions_Test(DocumentType)
    Dim fDATE, sDATE, FolderName, my_vbObj
    Dim Overdraft, opDate, exTerm, MainSum, PerSum, Prc, NonUsedPrc
    Dim attr
      
    ''Համակարգ մուտք գործել ARMSOFT օգտագործողով
    fDATE = "20260101"
    sDATE = "20140101"
    Call Initialize_AsBank("bank", sDATE, fDATE)
    Login("ARMSOFT")
  
  '--------------------------------------
    Set attr = Log.CreateNewAttributes
    attr.BackColor = RGB(0, 255, 255)
    attr.Bold = True
    attr.Italic = True
  '--------------------------------------  
    Call ChangeWorkspace(c_Subsystems)
  
    ''1.Ներգրավված օվերդրաֆտ պայմանագրի ստեղծում
    Set Overdraft = New_OverdraftAttrDoc()
    With Overdraft
      .CalcAcc = "00000113032"                                    
      .Limit = 100000
      .Date = "221018" 
      .GiveDate = "221018"
      .Term = "221019"
      .PaperCode = 555
    
    Select Case DocumentType
      Case 1 
       .DocType = "úí»ñ¹ñ³ýï"
      Case 2
       .DocType = "ØÇ³Ý·³ÙÛ³ ûí»ñ¹ñ³ýï" 
    End Select
    
    FolderName = "|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|ä³ÛÙ³Ý³·ñ»ñ|Ü»ñ·ñ³íí³Í ÙÇçáóÝ»ñ|Ü»ñ·ñ³íí³Í ûí»ñ¹ñ³ýï|"
    
    Call .CreateAttrOverdraft(FolderName & "Üáñ å³ÛÙ³Ý³·ñÇ ëï»ÕÍáõÙ")
  
    Log.Message(.DocNum)

    Call Close_Pttel("frmPttel")
  
    'Պայմանագրին ուղղարկել հաստատման
    .SendToVerify(FolderName & "²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ")
    'Հաստատել
    .Verify(FolderName & "Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I")
  
    Call LetterOfCredit_Filter_Fill(FolderName, .DocLevel, .DocNum)
    
    Call Log.Message("Օվերդրաֆտի ներգրավում",,,attr)
    Call Attraction(c_OverdraftAttraction, .Date, 80000, "", "")
    
    opDate = "211118"
    Call Log.Message("Տոկոսների հաշվարկ",,,attr)
    Call Calculate_Percents(opDate, opDate, False)

    opDate = "221118"
    Call Log.Message("Տոկոսների կապիտալացում",,,attr)
    Call Percent_Capitalization(Null , opDate, "")
    
    exTerm = "221020"
    Call Log.Message("Ժամկետների վերանայում",,,attr)
    Call Deposit_Extension(opDate, exTerm, "", .Paragraph, .Direction, c_TermsStates & "|" & c_Dates & "|" & c_ReviewTerms)

    Call Log.Message("Պարտքերի մարում",,,attr)
    MainSum = 10000
    Call Fade_Debt(opDate, Null, "", MainSum, PerSum, False)

    Call Log.Message("Տոկոսադրույքներ",,,attr)
    Prc = 15
    NonUsedPrc = 10
    Call ChangeRete(opDate, Prc, NonUsedPrc)
    
    Call Log.Message("Տոկոսների հաշվարկ",,,attr)
    Call Calculate_Percents(opDate, opDate, False)

    opDate = "231118"
    Call Log.Message("Պարտքերի մարում",,,attr)
    Call Fade_Debt(opDate, Null, Null,  70733.8, 29.1, False)
      
    Call Log.Message("Պայմանագրի փակում",,,attr)
    .CloseDate = opDate
    .CloseAgr()
  
    'Պայմանագրի բացում
    .OpenAgr()

    Call Log.Message("Բոլոր փաստաթղթերի ջնջում",,,attr)
    'Ջնջում "Գործողությունների դիտումից"
    Call Close_Pttel("frmPttel")
    Call GroupDelete(FolderName, 1, .DocNum, "^A[Del]", "^A[Del]", "^A[Del]")
    
    Call LetterOfCredit_Filter_Fill(FolderName, 1, .DocNum)
  End With
  
  'Ջնջել "Պայմ.մարման ժամկետներ"-ը
  Call Delete_ViewEdit("^A[Del]", "^A[Del]", c_Dates & "|" & c_AgrDates)

  BuiltIn.Delay(2000)
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_Folders & "|" & c_AgrFolder)
  BuiltIn.Delay(2000)
  
  Set my_vbObj = wMDIClient.VBObject("frmPttel_2").VBObject("tdbgView")
  With my_vbObj  
     .MoveFirst
     Do While (Not .EOF)
      Log.Message(Left(.Columns.Item(0).Text, 38))
      If Left(.Columns.Item(0).Text, 15) = "îáÏáë³¹ñáõÛùÝ»ñ" then 
        Call wMainForm.MainMenu.Click(c_AllActions)
        Call wMainForm.PopupMenu.Click(c_Delete)
        BuiltIn.Delay(1000)
        Call ClickCmdButton(3, "²Ûá")
        Exit Do   
      Else
        Call .MoveNext
      End If
     Loop 
  End With
  Call Close_Pttel("frmPttel_2")
  
  'Ջնջել օվերդրաֆտի պայմանագիրը
  BuiltIn.Delay(1000)
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_Delete)
  BuiltIn.Delay(1000)
  Call ClickCmdButton(3, "²Ûá")
  
  Call Close_AsBank() 
End Sub