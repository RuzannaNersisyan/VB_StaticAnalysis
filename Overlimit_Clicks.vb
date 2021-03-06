Option Explicit

'USEUNIT Library_Common  
'USEUNIT Akreditiv_Library 
'USEUNIT Constants

'Test Case Id - 160490

Sub Overlimit_Clicks_Test()
  Dim fDATE, sDATE, Count, i, DocNum, DocType, DocLevel, FolderName
  Dim arrayWaitForDoc, arrayWaitForView, arrayWaitForModalBrowser, arrayFrmSpr
  Dim attr   
  
  ''1, Համակարգ մուտք գործել ARMSOFT օգտագործողով
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

  FolderName = "|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|ä³ÛÙ³Ý³·ñ»ñ|î»Õ³µ³ßËí³Í ÙÇçáóÝ»ñ|¶»ñ³Í³Ëë|"
  Call ChangeWorkspace(c_Subsystems)
  
  ReDim arrayWaitForDoc(15)          'պետք է բացվի Doc
  arrayWaitForDoc = Array(c_ToEdit, c_View, c_Folders & "|" & c_CurrentSchedule,_
                           c_InputPrimaryContract,_
                           c_Opers & "|" & c_GiveAndBack & "|" & c_Overlimit,_
                           c_Opers & "|" & c_GiveAndBack & "|" & c_PayOffDebt,_
                           c_Opers & "|" & c_GiveAndBack & "|" & c_WriteOffRepay,_
                           c_Opers & "|" & c_Store & "|" & c_Store,_
                           c_Opers & "|" & c_Interests & "|" & c_PrcAccruing,_
                           c_Opers & "|" & c_WriteOff & "|" & c_WriteOff,_
                           c_Opers & "|" & c_WriteOff & "|" & c_WriteOffBack,_
                           c_Opers & "|" & c_DebtLet,_
                           c_TermsStates & "|" & c_Dates & "|" & c_ReviewSchedule,_
                           c_TermsStates & "|" & c_Risking & "|" & c_RiskCatPerRes,_
                           c_TermsStates & "|" & c_Percentages & "|" & c_Percentages)
  ReDim arrayWaitForView(12)          'պետք է բացվի View
  arrayWaitForView = Array(c_DocumentLog, c_Folders & "|" & c_ClFolder,_
                            c_Folders & "|" & c_AgrFolder, c_Folders & "|" & c_SchFolder,_
                            c_References & "|" & c_CheckPastdueSums, c_OpersView,_
                            c_ViewEdit & "|" & c_Risking & "|" &  c_RisksPersRes,_
                            c_ViewEdit & "|" & c_Percentages & "|" & c_Percentages,_
                            c_ViewEdit & "|" & c_Other & "|" & c_AccAdjust,_
                            c_ViewEdit & "|" & c_Other & "|" & c_CalcDates,_
                            c_AccEntries & "|" & c_ForBal, c_AccEntries & "|" & c_ForOffBal)
  ReDim arrayFrmSpr(2)        
  arrayFrmSpr = Array(c_References & "|" & c_CommView, c_References & "|" &  c_Statement)

  ''1.Գերածախս(գրաֆիկով պայմ.)
  Call Log.Message("Գերածախս(գրաֆիկով պայմ.)",,,attr)
  DocLevel = 1  
  DocNum = "^A[Del]" & "000387400"
  Call LetterOfCredit_Filter_Fill(FolderName, DocLevel, DocNum)
  
  Count = 15
  For i = 0 To Count-1
    Call OnClick(arrayWaitForDoc(i), "frmASDocForm")
  Next
  
  Count = 12
  For i = 0 To Count-1
    Call OnClick(arrayWaitForView(i), "AsView")
  Next 
  
  Count = 2
  For i = 0 To Count-1
    Call OnClick(arrayFrmSpr(i), "FrmSpr")
  Next 
  
  'Ջնջել   
  Call OnClick(c_Delete, "frmDeleteDoc")
  
  'Պայմանագրի փակում
  Call OnClick(c_AgrClose, "frmAsUstPar")

  wMDIClient.VBObject("frmPttel").Close

  ''2.Գերածախս անժամկետ
  Call Log.Message("Գերածախս անժամկետ",,,attr)
  DocLevel = 1  
  DocNum = "^A[Del]" & "77808541849"
  Call LetterOfCredit_Filter_Fill(FolderName, DocLevel, DocNum)
  
  Count = 15
  For i = 0 To Count-1
    If i <> 2 and i <> 12 Then
      Call OnClick(arrayWaitForDoc(i), "frmASDocForm")
    End If
  Next
  Call OnClick(c_Opers & "|" & c_Interests & "|" & c_AccAdjust, "frmASDocForm")
  
  Count = 12
  For i = 0 To Count-1
    If i <> 3 Then
      Call OnClick(arrayWaitForView(i), "AsView")
    End If    
  Next 
  
  Count = 2
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