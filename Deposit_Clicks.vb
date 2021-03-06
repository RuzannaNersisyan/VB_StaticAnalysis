Option Explicit

'USEUNIT Library_Common  
'USEUNIT Akreditiv_Library 
'USEUNIT Constants

'Test Case Id - 160288

Sub Deposit_Clicks_Test()
  Dim fDATE, sDATE, AsUstPar, Count, i, DocNum, DocType, DocLevel, FolderName
  Dim arrayWaitForDoc, arrayWaitForView, arrayWaitForModalBrowser, arrayFrmSpr
  Dim attr
        
  ''Համակարգ մուտք գործել ARMSOFT օգտագործողով
  fDATE = "20220101"
  sDATE = "20140101"
  Call Initialize_AsBank("bank", sDATE, fDATE)
  Login("ARMSOFT")

'--------------------------------------
  Set attr = Log.CreateNewAttributes
  attr.BackColor = RGB(255, 255, 0)
  attr.Bold = True
  attr.Italic = True
'-------------------------------------- 
  ReDim arrayWaitForDoc(15)          'պետք է բացվի Doc
  arrayWaitForDoc = Array(c_ToEdit, c_View,_
                           c_Safety & "|" & c_AgrBindNew & "|" & c_AgrBind,_
                           c_Opers & "|" & c_GiveAndBack & "|" & c_PayOffDebt,_
                           c_Opers & "|" & c_GiveAndBack & "|" & c_ReturnPrepaidInt,_
                           c_Opers & "|" & c_GiveAndBack & "|" & c_FadeDepFromPercent,_ 
                           c_Opers & "|" & c_Interests & "|" & c_PrcAccruing,_
                           c_Opers & "|" & c_Interests & "|" & c_AccAdjust,_
                           c_Opers & "|" & c_AgrBreak,_
                           c_TermsStates & "|" & c_Dates & "|" & c_ReviewTerms,_
                           c_TermsStates & "|" & c_Dates & "|" & c_OtherPaySchedule,_
                           c_TermsStates & "|" & c_Percentages & "|" & c_Percentages,_
                           c_TermsStates & "|" & c_Percentages & "|" & c_EffRate,_
                           c_TermsStates & "|" & c_Other & "|" & c_RecalculateRate)
  ReDim arrayWaitForView(15)          'պետք է բացվի View
  arrayWaitForView = Array(c_DocumentLog, c_Folders & "|" & c_ClFolder, c_Folders & "|" & c_AgrFolder,_
                            c_References & "|" & c_CheckPastdueSums,_
                            c_Safety & "|" & c_AgrBindNew & "|" & c_LinksOfAgreement, c_OpersView,_
                            c_ViewEdit & "|" & c_Dates & "|" & c_AgrDates,_
                            c_ViewEdit & "|" & c_Dates & "|" & c_PerDates,_
                            c_ViewEdit & "|" & c_Percentages & "|" & c_Percentages,_
                            c_ViewEdit & "|" & c_Percentages & "|" & c_EffRate,_
                            c_ViewEdit & "|" & c_Other & "|" & c_AccAdjust,_
                            c_ViewEdit & "|" & c_Other & "|" & c_TaxRates,_
                            c_ViewEdit & "|" & c_Other & "|" & c_RecalcRate,_
                            c_ViewEdit & "|" & c_Other & "|" & c_CalcDates,_
                            c_AccEntries & "|" & c_ForBal)
  ReDim arrayWaitForModalBrowser(2)   'պետք է բացվի ModalBrowser
  arrayWaitForModalBrowser = Array(c_Safety & "|" & c_AgrOpen & "|" & c_Mortgage,_
                                    c_Safety & "|" & c_AgrBind & "|" & c_Mortgage)
  ReDim arrayFrmSpr(3)        
  arrayFrmSpr = Array(c_References & "|" & c_CommView, c_References & "|" & c_CliRepaySchedule, c_References & "|" &  c_Statement)

  Call ChangeWorkspace(c_Subsystems)
  FolderName = "|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|ä³ÛÙ³Ý³·ñ»ñ|î»Õ³µ³ßËí³Í ÙÇçáóÝ»ñ|î»Õ³µ³ßËí³Í ³í³Ý¹Ý»ñ|"
  
   ''1.Տեղաբաշխված ավանդներ/Ավանդային պայմանագիր
  Call Log.Message("Տեղաբաշխված ավանդներ/Ավանդային պայմանագիր",,,attr)
  DocLevel = 1
  DocNum = "^A[Del]" & "0000000009"
  Call LetterOfCredit_Filter_Fill(FolderName, DocLevel, DocNum)
  
  Count = 14
  For i = 0 To Count-1
  If i = 4 Then
    Call OnClick(c_Opers & "|" & c_GiveAndBack & "|" & c_GiveDeposit, "frmASDocForm")
  ElseIf i <> 10 and i <> 11 And i <> 12 And i <> 18 And i <> 20 Then
    Call OnClick(arrayWaitForDoc(i), "frmASDocForm")
  End If  
  Next
  Call OnClick(c_Opers  & "|" & c_GiveAndBack & "|" & c_WriteOffRepay, "frmASDocForm")
  Call OnClick(c_Opers  & "|" & c_Store & "|" & c_Store, "frmASDocForm")
  Call OnClick(c_InputPrimaryContract, "frmASDocForm")
  
  Count = 15
  For i = 0 To Count-1
    If i <> 3 and i <> 11 Then
      BuiltIn.Delay(100)
      Call OnClick(arrayWaitForView(i), "AsView")
    End If  
  Next 
  Call OnClick(c_AccEntries & "|" & c_ForOffBal, "AsView")
  
  Count = 2
'  For i = 0 To Count-1
'    Call OnClick(arrayWaitForModalBrowser(i), "frmModalBrowser")
'  Next 
'  Call OnClick(c_Safety & "|" & c_AgrOpen & "|" & c_Mortgage, "frmModalBrowser")
  Call OnClick(c_Safety & "|" & c_AgrBind & "|" & c_Guarantee, "frmModalBrowser")
  
  Count = 3
  For i = 0 To Count-1
    Call OnClick(arrayFrmSpr(i), "FrmSpr")
  Next 
  
  'Ջնջել   
  Call OnClick(c_Delete, "frmDeleteDoc")
  
  'Պայմանագրի փակում
  Call OnClick(c_AgrClose, "frmAsUstPar")

  wMDIClient.VBObject("frmPttel").Close
  
  FolderName = "|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|ä³ÛÙ³Ý³·ñ»ñ|Ü»ñ·ñ³íí³Í ÙÇçáóÝ»ñ|²í³Ý¹Ý»ñ (Ý»ñ·ñ³íí³Í)|"
  ''1.Ավանդային պայմանագիր
  Call Log.Message("Ավանդներ (ներգրավված)/Ավանդային պայմանագիր",,,attr)
  DocLevel = 1
  DocNum = "^A[Del]" & "A-000306"
  Call LetterOfCredit_Filter_Fill(FolderName, DocLevel, DocNum)
                                                            
  Count = 14
  For i = 0 To Count-1
    Call OnClick(arrayWaitForDoc(i), "frmASDocForm")
  Next
  Call OnClick(c_Opers & "|" & c_PassSums, "frmASDocForm")
  
  Count = 15
  For i = 0 To Count-1
    Call OnClick(arrayWaitForView(i), "AsView")
  Next 
  
'  Count = 2
'  For i = 0 To Count-1
'    Call OnClick(arrayWaitForModalBrowser(i), "frmModalBrowser")
'  Next 
  
  Count = 3
  For i = 0 To Count-1
    Call OnClick(arrayFrmSpr(i), "FrmSpr")
  Next 
  
  'Ջնջել   
  Call OnClick(c_Delete, "frmDeleteDoc")
  
  'Պայմանագրի փակում
  Call OnClick(c_AgrClose, "frmAsUstPar")

  wMDIClient.VBObject("frmPttel").Close

  ''2.Միանգամյա ավանդ
  Call Log.Message("Ավանդներ (ներգրավված)/Միանգամյա ավանդ",,,attr)
  DocLevel = 1
  DocNum = "^A[Del]" & "A-000304"
  Call LetterOfCredit_Filter_Fill(FolderName, DocLevel, DocNum)
  
  Count = 14
  For i = 0 To Count-1
  If i <> 4 and i <> 5 Then
    Call OnClick(arrayWaitForDoc(i), "frmASDocForm")
  End If  
  Next
  Call OnClick(c_TermsStates & "|" & c_Percentages & "|" & c_BankEffRate, "frmASDocForm")
  Call OnClick(c_Opers & "|" & c_PassSums, "frmASDocForm")
  
  Count = 15
  For i = 0 To Count-1
    Call OnClick(arrayWaitForView(i), "AsView")
  Next 
  Call OnClick(c_ViewEdit & "|" & c_Percentages & "|" & c_BankEffRate, "AsView")
    
'  Count = 2
'  For i = 0 To Count-1
'    Call OnClick(arrayWaitForModalBrowser(i), "frmModalBrowser")
'  Next 
  
  Count = 3
  For i = 0 To Count-1
    Call OnClick(arrayFrmSpr(i), "FrmSpr")
  Next 
  
  'Ջնջել   
  Call OnClick(c_Delete, "frmDeleteDoc")
  
  'Պայմանագրի փակում
  Call OnClick(c_AgrClose, "frmAsUstPar")

  wMDIClient.VBObject("frmPttel").Close
  
  Call Close_AsBank()
End Sub      