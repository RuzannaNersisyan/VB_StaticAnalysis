Option Explicit

'USEUNIT Library_Common  
'USEUNIT Repo_Library
'USEUNIT Payment_Order_ConfirmPhases_Library
'USEUNIT Akreditiv_Library
'USEUNIT Loan_Agreements_Library 
'USEUNIT Deposit_Contract_Library
'USEUNIT Loan_Agreemnts_With_Schedule_Library
'USEUNIT Group_Operations_Library
'USEUNIT Constants
'USEUNIT Mortgage_Library

'Test Case Id 165780

Sub Inverse_Repo_Actions_Test()
  Dim fDATE, sDATE, attr, opDate, exTerm, Prc, NonUsedPrc, EffRete, ActRete, FolderName
  Dim client, curr, acc, summa, date, kindscale,per, baj, dateGive, dateAgr, DateFill,_
      startDate,CheckPayDates, PayDates, Paragraph, Direction ,secState, secClass,_
      security, Price, fBASE, DocNum
'--------------------------------------
  Set attr = Log.CreateNewAttributes
  attr.BackColor = RGB(0, 255, 255)
  attr.Bold = True
  attr.Italic = True
'-------------------------------------- 
  
  ''Համակարգ մուտք գործել ARMSOFT օգտագործողով
  fDATE = "20260101"
  sDATE = "20140101"
  Call Initialize_AsBank("bank", sDATE, fDATE)
  Login("ARMSOFT")
  
  FolderName = "|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|ä³ÛÙ³Ý³·ñ»ñ|Ü»ñ·ñ³íí³Í ÙÇçáóÝ»ñ|Ð³Ï³¹³ñÓ é»åá Ñ³Ù³Ó³ÛÝ³·ñ»ñ|"
  
  Call ChangeWorkspace(c_Subsystems)
  Call wTreeView.DblClickItem(FolderName & "Üáñ å³ÛÙ³Ý³·ñÇ ëï»ÕÍáõÙ")
  
  acc = "00000113032"
  date = "221018"
  kindscale = ""
  per = 12
  baj = 365
  dateGive = "221018"
  dateAgr = "221019"
  DateFill = 1
  startDate = "221018"
  CheckPayDates = 0
  Paragraph = 1 
  Direction = 2
  secState = 1
  secClass = 6
  security = "R-0007"
  Call Inverse_Repo_Create(client, curr, acc, summa, date, kindscale,per, baj, dateGive,_
                            dateAgr, DateFill,startDate,CheckPayDates, PayDates, Paragraph,_
                            Direction ,secState, secClass, security,Price,fBASE, DocNum)

  Log.Message(DocNum)                            
  ''4.Պայմանագրը ուղարկել հաստատման                               
  Call PaySys_Send_To_Verify()
                                 
  ''5.Հաստատել պայմանագիրը
  BuiltIn.Delay(2000)
  Call Close_Pttel("frmPttel")
  Call wTreeView.DblClickItem(FolderName & "Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I")
  'Լրացնել "Պայմանագարի համար"   
  Call Rekvizit_Fill("Dialog",1,"General","NUM",DocNum)
  'Սեղմել "Կատարել" կոճակը
  Call ClickCmdButton(2, "Î³ï³ñ»É")

  'Հաստատել Հաստատող փաստաթղթեր 1- ում
  Call PaySys_Verify(True)
  BuiltIn.Delay(2000)
  Call Close_Pttel("frmPttel")
  
  Call LetterOfCredit_Filter_Fill(FolderName, 1, DocNum)
  
  Call Log.Message("Ռեպոյի ներգրավում",,,attr)
  Call InverseRepoAttraction(date)
  
  opDate = "211118"
  Call Log.Message("Տոկոսների հաշվարկ",,,attr)
  Call Calculate_Percents(opDate, opDate, False)

  opDate = "221118"
  exTerm = "221020"
  Call Log.Message("Ժամկետների վերանայում",,,attr)
  Call Deposit_Extension(opDate, exTerm, "", Paragraph, Direction, c_TermsStates & "|" & c_Dates & "|" & c_ReviewTerms)
  
  Call Log.Message("Տոկոսադրույքներ",,,attr)
  Prc = 15
  NonUsedPrc = 10
  Call ChangeRete(opDate, Prc, NonUsedPrc)
    
  Call Log.Message("Արդյունավետ տոկոսադրույք",,,attr)
  Call ChangeEffRete(opDate, EffRete, ActRete)
    
  Call Log.Message("Բանկի արդյունավետ տոկոսադրույք",,,attr)
  Call BankEffective_InterestRate_DocFill(opDate, "")
  
  Call Log.Message("Տոկոսների հաշվարկ",,,attr)
  Call Calculate_Percents(opDate, opDate, False)
  
  opDate = "231118"
  Call Log.Message("Պարտքերի մարում",,,attr)
  Call InverseRepoRepay(opDate, 43294821.60, 459043.70, 2)
                                                          
  Call Log.Message("Պայմանագրի փակում",,,attr)
  BuiltIn.Delay(2000)
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_AgrClose)
  Call Rekvizit_Fill("Dialog", 1, "General", "DATECLOSE", opDate)
  Call ClickCmdButton(2, "Î³ï³ñ»É")

  BuiltIn.Delay(2000)
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_AgrOpen)
  BuiltIn.Delay(2000)
  Call ClickCmdButton(5, "²Ûá")
  
  'Ջնջել բոլոր պայմանագրերը
  Call Log.Message("Բոլոր փաստաթղթերի ջնջում",,,attr)
  
 'Ջնջում հաշվարկման ամսաթվերից
  Call Delete_ViewEdit(opDate, opDate, c_Other & "|" & c_CalcDates)
  
 'Ջնջել բոլոր գործողությունները
  BuiltIn.Delay(2000)
  Call Close_Pttel("frmPttel")
  Call GroupDelete(FolderName, 1, DocNum, "^A[Del]", "^A[Del]", "^A[Del]")
    
  Call LetterOfCredit_Filter_Fill(FolderName, 1, DocNum)

  'Ջնջել տոկոսների նշանակումները
  Call Delete_ViewEdit("^A[Del]", "^A[Del]", c_Percentages & "|" & c_Percentages)
  Call Delete_ViewEdit("^A[Del]", "^A[Del]", c_Percentages & "|" & c_EffRate)
  Call Delete_ViewEdit("^A[Del]", "^A[Del]", c_Percentages & "|" & c_BankEffRate)
  
  'Ջնջել Հակադարձ ռեպոյի պայմանագիրը
  BuiltIn.Delay(2000)
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_Delete)
  BuiltIn.Delay(2000)
  Call ClickCmdButton(3, "²Ûá")
  
  Call Close_AsBank()  
End Sub