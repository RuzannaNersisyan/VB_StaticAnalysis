Option Explicit

'USEUNIT Library_Common  
'USEUNIT Akreditiv_Library 
'USEUNIT Constants

'Test Case Id - 160484

Sub Overdraft_Clicks_Test()
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

  FolderName = "|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|ä³ÛÙ³Ý³·ñ»ñ|î»Õ³µ³ßËí³Í ÙÇçáóÝ»ñ|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|"
  Call ChangeWorkspace(c_Subsystems)
 
  Call Log.Message("Գրաֆիկով օվերդրաֆտ",,,attr)
  
  'Գրաֆիկով օվերդրաֆտի համար
  DocLevel = 2
  DocNum = 3756
	Call LetterOfCredit_Filter_Fill(FolderName, DocLevel, DocNum)
   
  ReDim arrayWaitForDoc(21)          'պետք է բացվի Doc
  arrayWaitForDoc = Array(c_ToEdit, c_View, c_Folders & "|" & c_CurrentSchedules,_
                           c_Safety & "|" & c_AgrOpen & "|" & c_Guarantee,_
                           c_Safety & "|" & c_AgrBindNew & "|" & c_AgrBind,_
                           c_InputPrimaryContract,_
                           c_Opers & "|" & c_GiveAndBack & "|" & c_GiveOverdraft,_
                           c_Opers & "|" & c_GiveAndBack & "|" & c_PayOffDebt,_
                           c_Opers & "|" & c_GiveAndBack & "|" & c_ReturnPrepaidInt,_
                           c_Opers & "|" & c_GiveAndBack & "|" & c_WriteOffRepay,_
                           c_Opers & "|" & c_Store & "|" & c_Store,_
                           c_Opers & "|" & c_Store & "|" & c_UnusedPartStore,_
                           c_Opers & "|" & c_WriteOff & "|" & c_WriteOffBack,_
                           c_Opers & "|" & c_DebtLet,_
                           c_TermsStates & "|" & c_Dates & "|" & c_ReviewSchedule,_
                           c_TermsStates & "|" & c_Dates & "|" & c_OtherPaySchedule,_
                           c_TermsStates & "|" & c_Risking & "|" & c_RiskCatPerRes,_
                           c_TermsStates & "|" & c_Percentages & "|" & c_Percentages,_
                           c_TermsStates & "|" & c_Percentages & "|" & c_EffRate,_
                           c_TermsStates & "|" & c_Other & "|" & c_Limit,_
                           c_TermsStates & "|" & c_StopLine)
  ReDim arrayWaitForView(16)          'պետք է բացվի View
  arrayWaitForView = Array(c_DocumentLog, c_Folders & "|" & c_ClFolder,_
                            c_Folders & "|" & c_AgrFolder, c_Folders & "|" & c_SchFolder,_
                            c_References & "|" & c_CheckInterest,_
                            c_References & "|" & c_CheckPastdueSums,_
                            c_Safety & "|" & c_AgrBindNew & "|" & c_LinksOfAgreement, c_OpersView,_
                            c_ViewEdit & "|" &  c_Risking & "|" &  c_RisksPersRes,_
                            c_ViewEdit & "|" & c_Percentages & "|" & c_Percentages,_
                            c_ViewEdit & "|" & c_Percentages & "|" & c_EffRate,_
                            c_ViewEdit & "|" & c_Other & "|" & c_AccAdjust,_
                            c_ViewEdit & "|" & c_Other & "|" & c_Limits,_
                            c_ViewEdit & "|" & c_Other & "|" & c_CalcDates,_
                            c_ViewEdit & "|" & c_LineBrRec, c_AccEntries & "|" & c_ForBal,_
                            c_AccEntries & "|" & c_ForOffBal)                            
  ReDim arrayWaitForModalBrowser(3)   'պետք է բացվի ModalBrowser
  arrayWaitForModalBrowser = Array(c_Safety & "|" & c_AgrOpen & "|" & c_Mortgage,_
                                    c_Safety & "|" & c_AgrBind & "|" & c_Mortgage,_
                                    c_Safety & "|" & c_AgrBind & "|" & c_Guarantee)
  ReDim arrayFrmSpr(3)        
  arrayFrmSpr = Array(c_References & "|" & c_CommView, c_References & "|" & c_CliRepaySchedule, c_References & "|" &  c_Statement)
                                                              
  Count = 21
  For i = 0 To Count-1
    Call OnClick(arrayWaitForDoc(i), "frmASDocForm")
  Next
  
  Count = 17
  For i = 0 To Count-1
    Call OnClick(arrayWaitForView(i), "AsView")
  Next 
  
  Count = 3
  For i = 0 To Count-1
    Call OnClick(arrayWaitForModalBrowser(i), "frmModalBrowser")
  Next 
  
  Count = 3
  For i = 0 To Count-1
    Call OnClick(arrayFrmSpr(i), "FrmSpr")
  Next 
  
  'Ջնջել   
  Call OnClick(c_Delete, "frmDeleteDoc")
  
  'Պայմանագրի փակում
  Call OnClick(c_AgrClose, "frmAsUstPar")

  wMDIClient.VBObject("frmPttel").Close
  
  ''Գրաֆիկով օվերդրաֆտի(Արտոնյալ ժամկետով) համար DocType = "8S"
  Call Log.Message("Գրաֆիկով օվերդրաֆտի(Արտոնյալ ժամկետով)",,,attr)
  DocLevel = 1  
  DocNum = "ST-006"
  Call LetterOfCredit_Filter_Fill(FolderName, DocLevel, DocNum)
  
  Count = 21
  For i = 0 To Count-1
    If i <> 2 Then
      Call OnClick(arrayWaitForDoc(i), "frmASDocForm")
    End If  
  Next
  Call OnClick(c_Folders & "|" & c_CurrentSchedule, "frmASDocForm") 

  Count = 17
  For i = 0 To Count-1
    Call OnClick(arrayWaitForView(i), "AsView")
  Next 
  Call OnClick(c_ViewEdit & "|" & c_Other & "|" & c_OvGrPerCalcDates, "AsView")
  
  Count = 3
  For i = 0 To Count-1
    Call OnClick(arrayWaitForModalBrowser(i), "frmModalBrowser")
  Next 
  
  Count = 3
  For i = 0 To Count-1
    Call OnClick(arrayFrmSpr(i), "FrmSpr")
  Next 

  'Ջնջել   
  Call OnClick(c_Delete, "frmDeleteDoc")
  
  'Պայմանագրի փակում
  Call OnClick(c_AgrClose, "frmAsUstPar")

  wMDIClient.VBObject("frmPttel").Close
  
  ''Բարդ օվերդրաֆտի համար 
  Call Log.Message("Բարդ օվերդրաֆտ",,,attr)
  DocLevel = 2
  DocNum = "0092"
	Call LetterOfCredit_Filter_Fill(FolderName, DocLevel, DocNum)
    
  ReDim arrayWaitForDoc(20)          'պետք է բացվի Doc
  arrayWaitForDoc = Array(c_ToEdit, c_View,_
                           c_Safety & "|" & c_AgrOpen & "|" & c_Guarantee,_
                           c_Safety & "|" & c_AgrBindNew & "|" & c_AgrBind,_
                           c_OpenSubAgr, c_InputPrimaryContract,_
                           c_Opers & "|" & c_IntRepayment,_
                           c_Opers & "|" & c_Interests & "|" & c_PrcAccruing,_
                           c_Opers & "|" & c_Interests & "|" & c_AccAdjust,_
                           c_Opers & "|" & c_Store & "|" & c_Store,_
                           c_Opers & "|" & c_Store & "|" & c_UnusedPartStore,_
                           c_Opers & "|" & c_WriteOff & "|" & c_WriteOff,_
                           c_Opers & "|" & c_WriteOff & "|" & c_WriteOffBack,_
                           c_Opers & "|" & c_WriteOff & "|" & c_DebtLet,_
                           c_TermsStates & "|" & c_Dates & "|" & c_ReviewTerms,_
                           c_TermsStates & "|" & c_Dates & "|" & c_OtherPaySchedule,_
                           c_TermsStates & "|" & c_Risking & "|" & c_RiskCatPerRes,_
                           c_TermsStates & "|" & c_Percentages & "|" & c_Percentages,_
                           c_TermsStates & "|" & c_Percentages & "|" & c_EffRate,_
                           c_TermsStates & "|" & c_Other & "|" & c_Limit,_
                           c_TermsStates & "|" & c_StopLine)
  ReDim arrayWaitForView(16)          'պետք է բացվի View
  arrayWaitForView = Array(c_DocumentLog, c_Folders & "|" & c_ClFolder,_
                            c_Folders & "|" & c_AgrFolder, c_Folders & "|" & c_AgrChildren,_
                            c_Safety & "|" & c_AgrBindNew & "|" & c_LinksOfAgreement, c_OpersView,_
                            c_ViewEdit & "|" & c_Dates & "|" & c_AgrDates,_
                            c_ViewEdit & "|" & c_Dates & "|" & c_PerDates,_
                            c_ViewEdit & "|" &  c_Risking & "|" &  c_RisksPersRes,_
                            c_ViewEdit & "|" & c_Percentages & "|" & c_Percentages,_
                            c_ViewEdit & "|" & c_Percentages & "|" & c_EffRate,_
                            c_ViewEdit & "|" & c_Other & "|" & c_AccAdjust,_
                            c_ViewEdit & "|" & c_Other & "|" & c_Limits,_
                            c_ViewEdit & "|" & c_Other & "|" & c_CalcDates,_
                            c_ViewEdit & "|" & c_LineBrRec,_
                            c_AccEntries & "|" & c_ForBal, c_AccEntries & "|" & c_ForOffBal)                           
                                                        
  Count = 20
  For i = 0 To Count-1
    Call OnClick(arrayWaitForDoc(i), "frmASDocForm")
  Next 

  Count = 17
  For i = 0 To Count-1
    Call OnClick(arrayWaitForView(i), "AsView")
  Next 
  
  Count = 3
  For i = 0 To Count-1
    Call OnClick(arrayWaitForModalBrowser(i), "frmModalBrowser")
  Next 
  
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