Option Explicit

'USEUNIT Library_Common  
'USEUNIT Deposit_Contract_Library
'USEUNIT Payment_Order_ConfirmPhases_Library
'USEUNIT Akreditiv_Library
'USEUNIT Loan_Agreements_Library 
'USEUNIT Credit_Line_Library
'USEUNIT Loan_Agreemnts_With_Schedule_Library
'USEUNIT Group_Operations_Library
'USEUNIT Constants
'USEUNIT Mortgage_Library

'Test case Id 165743
'Test case Id 165746
'Test case ID 165752

Sub Deposit_Actions_Test(DocumentType)
  Dim fDATE, sDATE, attr, frmAsMsgBox, FrmSpr, FolderName, fBASE_depInv
  Dim opDate, Sum, exTerm, MainSum, Prc, NonUsedPrc, EffRete, ActRete
  Dim fBASE,DocNum,template,depositContractType,colItem, perSum, _
      ClientCode,thirdPerson,curr,CalcAcc,thirdAcc,perAcc, chbKap,_
      chbAuto,chbEx, Date,kindScale,scale,withScale,depositPer,part,per,GiveDate,_
      Term,startDate,period,direction,payDates, sumsDateFillType, sumsFillType    
  Dim InvDocNum, DocLevel  
  Dim FillType, FirstDate, EndDate
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
  
  FolderName = "|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|ä³ÛÙ³Ý³·ñ»ñ|Ü»ñ·ñ³íí³Í ÙÇçáóÝ»ñ|²í³Ý¹Ý»ñ (Ý»ñ·ñ³íí³Í)|"
  
  ''2.Մուտք գործել "Ենթահամակրգեր(ՀԾ)"
  Call ChangeWorkspace(c_Subsystems) 
  
  wTreeView.DblClickItem(FolderName & "¸³ï³ñÏ å³ÛÙ³Ý³·Çñ")
  
  ''3.Ներգրավված ավանդի պայմանագրի ստեղծում
  Select Case  DocumentType
  Case 1 
    depositContractType = "²í³Ý¹³ÛÇÝ å³ÛÙ³Ý³·Çñ"
  Case 2
    depositContractType = "ØÇ³Ý·³ÙÛ³ ³í³Ý¹"  
  Case 3 
    depositContractType = "¶ñ³ýÇÏáí ³í³Ý¹³ÛÇÝ å³ÛÙ³Ý³·Çñ"    
  End Select 
  
  colItem = "0"
  CalcAcc = "03485010100"
  Sum = 100000
  chbKap = 0
  chbAuto = 1
  chbEx = 1
  Date = "121118"
  kindScale = "1"
  depositPer = 10
  part = 365
  per = 0.5
  GiveDate = "121118"
  Term = "121119"
  period = 1
  direction = 2
  scale = False
  startDate = Date
  
  If depositContractType = "¶ñ³ýÇÏáí ³í³Ý¹³ÛÇÝ å³ÛÙ³Ý³·Çñ" Then
    FillType = 1
    payDates = 12
    sumsDateFillType = 1
    sumsFillType = "01"
    Call Deposit_Contract_With_Schedule_Fill(fBASE,DocNum,depositContractType,colItem,template, _
                              ClientCode,thirdPerson,curr,CalcAcc,thirdAcc,perAcc, Sum,chbKap,_
                              chbAuto,Date,GiveDate,Term, FillType, FirstDate, EndDate, payDates,sumsDateFillType,sumsFillType,direction,kindScale,depositPer,part,per)
                              
    'Մարման գրաֆիկի նշանակում   
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_RepaySchedule)
    Call Find_Data("¶ñ³ýÇÏáí ³í³Ý¹³ÛÇÝ å³ÛÙ³Ý³·Çñ- "& DocNum &" {êäÀ111}",0)
  Else 
    Call Deposit_Contract_Fill(fBASE,DocNum,template,depositContractType,colItem, _
                               ClientCode,thirdPerson,curr,CalcAcc,thirdAcc,perAcc,Sum,chbKap,_
                               chbAuto,chbEx,Date,kindScale,scale,withScale,depositPer,part,per, GiveDate,_
                               Term,startDate,period,direction)
  End If
  
  Log.Message(DocNum)                               

  ''4.Պայմանագրը ուղարկել հաստատման                               
  Call PaySys_Send_To_Verify()
  BuiltIn.Delay(3000)
  Call Close_Pttel("frmPttel")
                                 
  ''5.Հաստատել պայմանագիրը
  Call wTreeView.DblClickItem(FolderName & "Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I")
  'Լրացնել "Պայմանագարի համար"   
  Call Rekvizit_Fill("Dialog",1,"General","NUM",DocNum)
  'Սեղմել "Կատարել" կոճակը
  Call ClickCmdButton(2, "Î³ï³ñ»É")
  'Հաստատել Հաստատող փաստաթղթեր 1- ում
  Call PaySys_Verify(True)
  Call Close_Pttel("frmPttel")
  
'  wTreeView.DblClickItem(FolderName & "ä³ÛÙ³Ý³·ñ»ñ")
'  'Լրացնել "Պայմանագրի Մակարդակ" դաշտը
'  Call Rekvizit_Fill("Dialog",1,"General","LEVEL",DocLevel)
'    'Լրացնել "Պայմանագրի համար" դաշտը
'  Call Rekvizit_Fill("Dialog",1,"General","NUM",DocNum)
'    'Սեղմեձլ "Կատարել" կոճակը
'  Call ClickCmdButton(2, "Î³ï³ñ»É")
  DocLevel = 1
  Call LetterOfCredit_Filter_Fill(FolderName, DocLevel, DocNum)
  
  Call Log.Message("Ավանդի ներգրավում",,,attr)
  Call Deposit_Involvment(fBASE_depInv, InvDocNum, Date, Sum, 2, CalcAcc)

  opDate = "111218"    
  Call Log.Message("Տոկոսների հաշվարկ",,,attr)
  Call Calculate_Percent(Null, opDate , opDate)
  
  opDate = "121218" 
  Call Log.Message("Տոկոսների կապիտալացում",,,attr)
  Call Percent_Capitalization(Null, opDate, "")
  
  exTerm = "121120"   
  If depositContractType = "¶ñ³ýÇÏáí ³í³Ý¹³ÛÇÝ å³ÛÙ³Ý³·Çñ" Then
    Call Log.Message("Գրաֆիկի վերանայում",,,attr)
    Call Fading_Schedule_Fill(opDate, exTerm, 100715)
  Else
    Call Log.Message("Ժամկետների վերանայում",,,attr)
    Call Deposit_Extension(opDate, exTerm, "", period, Direction, c_TermsStates & "|" & c_Dates & "|" & c_ReviewTerms)
  End If
  
  Call Log.Message("Պարտքերի մարում",,,attr)
  MainSum = 10000
  perSum = ""
  Call Debt_Repayment(Null, opDate, MainSum,perSum, 2, CalcAcc,docNum, 2)

  Call Log.Message("Տոկոսադրույքներ",,,attr)
  Prc = 15
  NonUsedPrc = 10
  Call ChangeRete(opDate, Prc, NonUsedPrc)
    
  Call Log.Message("Արդյունավետ տոկոսադրույք",,,attr)
  Call ChangeEffRete(opDate, EffRete, ActRete)

  Call Log.Message("Պարտքերի մարում",,,attr)
  MainSum =  90715
  perSum = ""
  Call Debt_Repayment(Null, opDate, MainSum,perSum, 2, CalcAcc,docNum, 2)
  
  Call Log.Message("Պայմանագրի փակում",,,attr)
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_AgrClose)
    
  Call Rekvizit_Fill("Dialog", 1, "General", "DATECLOSE", opDate)
  Call ClickCmdButton(2, "Î³ï³ñ»É")

  Call Log.Message("Պայմանագրի բացում",,,attr)
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_AgrOpen)
  Call ClickCmdButton(5, "²Ûá")
  
  Call Log.Message("Բոլոր փաստաթղթերի ջնջում",,,attr)
  'Ջնջել բոլոր գործողությունները
  Call Close_Pttel("frmPttel")
  Call GroupDelete(FolderName, 1, docNum, "^A[Del]", "^A[Del]", "")
  Call LetterOfCredit_Filter_Fill(FolderName, 1, docNum)
  
  'Ջնջել տոկոսների նշանակումները
  Call Delete_ViewEdit("^A[Del]", "^A[Del]", c_Percentages & "|" & c_Percentages)
  Call Delete_ViewEdit("^A[Del]", "^A[Del]", c_Percentages & "|" & c_EffRate)

  If depositContractType <> "¶ñ³ýÇÏáí ³í³Ý¹³ÛÇÝ å³ÛÙ³Ý³·Çñ" Then
    'Ջնջել ժամկետների վերանայումը
    Call Delete_ViewEdit("^A[Del]", "^A[Del]", c_Dates & "|" & c_AgrDates) 
  Else
    Call Close_Pttel("frmPttel")
    Call GroupDelete(FolderName, 1, docNum, "^A[Del]", "^A[Del]", "")
    Call LetterOfCredit_Filter_Fill(FolderName, 1, docNum)
  End If
  
  'Ջնջել ավանդի պայմանագիրը
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_Delete)
  Call ClickCmdButton(3, "²Ûá")
  
  Call Close_AsBank()  
End Sub