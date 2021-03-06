Option Explicit

'USEUNIT Library_Common  
'USEUNIT Akreditiv_Library 
'USEUNIT Constants

'Test Case Id - 160496

Sub Securities_Clicks_Test()
  Dim fDATE, sDATE, Count, i, DocNum, DocType, DocLevel, FolderName
  Dim arrayWaitForDoc, arrayWaitForDoc1, arrayWaitForView, arrayWaitForView1
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
  
  ReDim arrayWaitForDoc(9)          'պետք է բացվի Doc
  arrayWaitForDoc = Array(c_ToEdit, c_View,_
                           c_Opers & "|" & c_GiveAndBack & "|" & c_CliSecTrade,_
                           c_Opers & "|" & c_GiveAndBack & "|" & c_PayOffDebt,_
                           c_Opers & "|" & c_GiveAndBack & "|" & c_SecSell,_
                           c_Opers & "|" & c_Interests & "|" & c_PrcAccruing,_
                           c_Opers & "|" & c_Pledging & "|" & c_SecPledging,_
                           c_Opers & "|" & c_Pledging & "|" & c_SecPledgeOut,_
                           c_Opers & "|" & c_Reclassification)
  ReDim arrayWaitForView(7)          'պետք է բացվի View
  arrayWaitForView = Array(c_DocumentLog, c_Folders & "|" & c_AgrFolder,_
                            c_References & "|" & c_CheckPastdueSums, c_OpersView,_
                            c_ViewEdit & "|" & c_Other & "|" & c_CalcDates,_
                            c_AccEntries & "|" & c_ForBal, c_AccEntries & "|" & c_ForOffBal)                            

  FolderName = "|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|ä³ÛÙ³Ý³·ñ»ñ|î»Õ³µ³ßËí³Í ÙÇçáóÝ»ñ|²ñÅ»ÃÕÃ»ñ ØØÄä|"
  Call ChangeWorkspace(c_Subsystems)
  
  ''1.Արժեթղթեր ՄՄԺՊ/Տոկոսային եկ.բերող արժեթուղթ
  Call Log.Message("Արժեթղթեր ՄՄԺՊ/Տոկոսային եկ.բերող արժեթուղթ",,,attr)
  DocLevel = 1
  DocNum = "S-0045"
  Call LetterOfCredit_Filter_Fill(FolderName, DocLevel, DocNum)
                            
  Count = 9
  For i = 0 To Count-1
    Call OnClick(arrayWaitForDoc(i), "frmASDocForm")
  Next
  
  Count = 7
  For i = 0 To Count-1
    Call OnClick(arrayWaitForView(i), "AsView")
  Next 
  
  Call OnClick(c_References & "|" & c_CommView, "FrmSpr")
  
  'Ջնջել   
  Call OnClick(c_Delete, "frmDeleteDoc")
  
  'Պայմանագրի փակում
  Call OnClick(c_AgrClose, "frmAsUstPar")

  wMDIClient.VBObject("frmPttel").Close

  ''2.Արժեթղթեր ՄՄԺՊ/Տոկոսային եկ.չբերող արժեթուղթ
  Call Log.Message("Արժեթղթեր ՄՄԺՊ/Տոկոսային եկ.չբերող արժեթուղթ",,,attr)
  DocLevel = 1
  DocNum = "S-0031"
  Call LetterOfCredit_Filter_Fill(FolderName, DocLevel, DocNum)

  Count = 9
  For i = 0 To Count-2
    If i <> 4 Then
      Call OnClick(arrayWaitForDoc(i), "frmASDocForm")
    End If   
  Next
  
  Count = 7
  For i = 0 To Count-1
    Call OnClick(arrayWaitForView(i), "AsView")
  Next 
  
  Call OnClick(c_References & "|" & c_CommView, "FrmSpr")
  
  'Ջնջել   
  Call OnClick(c_Delete, "frmDeleteDoc")
  
  'Պայմանագրի փակում
  Call OnClick(c_AgrClose, "frmAsUstPar")

  wMDIClient.VBObject("frmPttel").Close
  
  FolderName = "|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|ä³ÛÙ³Ý³·ñ»ñ|î»Õ³µ³ßËí³Í ÙÇçáóÝ»ñ|²ñÅ»ÃÕÃ»ñ í³×³éùÇ|"
  
  ''3.Արժեթղթեր վաճառքի/Տոկոսային եկ.բերող արժեթուղթ
  Call Log.Message("Արժեթղթեր վաճառքի/Տոկոսային եկ.բերող արժեթուղթ",,,attr)
  DocLevel = 1
  DocNum = "SS-0000004"
  Call LetterOfCredit_Filter_Fill(FolderName, DocLevel, DocNum)
                            
  Count = 9
  For i = 0 To Count-1
    Call OnClick(arrayWaitForDoc(i), "frmASDocForm")
  Next
  Call OnClick(c_Opers & "|" & c_SecPrCorr, "frmASDocForm")
  
  Count = 7
  For i = 0 To Count-1
    Call OnClick(arrayWaitForView(i), "AsView")
  Next 
  Call OnClick(c_ViewEdit & "|" & c_Other & "|" & c_PriceCorrDates, "AsView")
  Call OnClick(c_RevRepoSells, "AsView")
  
  Call OnClick(c_References & "|" & c_CommView, "FrmSpr")
  
  'Ջնջել   
  Call OnClick(c_Delete, "frmDeleteDoc")
  
  'Պայմանագրի փակում
  Call OnClick(c_AgrClose, "frmAsUstPar")

  wMDIClient.VBObject("frmPttel").Close
  
  FolderName = "|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|ä³ÛÙ³Ý³·ñ»ñ|î»Õ³µ³ßËí³Í ÙÇçáóÝ»ñ|²ñÅ»ÃÕÃ»ñ í»ñ³í³×³éùÇ|"
  
  ''4.Արժեթղթեր վերավաճառքի/Տոկոսային եկ.բերող արժեթուղթ
  Call Log.Message("Արժեթղթեր վերավաճառքի/Տոկոսային եկ.բերող արժեթուղթ",,,attr)
  DocLevel = 1
  DocNum = "RS-0000009"
  Call LetterOfCredit_Filter_Fill(FolderName, DocLevel, DocNum)
  
  Count = 9
  For i = 0 To Count-2
    Call OnClick(arrayWaitForDoc(i), "frmASDocForm")
  Next
  Call OnClick(c_Opers & "|" & c_SecPrCorr, "frmASDocForm")
  
  Count = 7
  For i = 0 To Count-1
    Call OnClick(arrayWaitForView(i), "AsView")
  Next 
  Call OnClick(c_ViewEdit & "|" & c_Other & "|" & c_PriceCorrDates, "AsView")
  Call OnClick(c_RevRepoSells, "AsView")
  
  Call OnClick(c_References & "|" & c_CommView, "FrmSpr")
  
  'Ջնջել   
  Call OnClick(c_Delete, "frmDeleteDoc")
  
  'Պայմանագրի փակում
  Call OnClick(c_AgrClose, "frmAsUstPar")

  wMDIClient.VBObject("frmPttel").Close
  
  ''4.Արժեթղթեր վերավաճառքի/Տոկոսային եկ.չբերող արժեթուղթ
  Call Log.Message("Արժեթղթեր վերավաճառքի/Տոկոսային եկ.չբերող արժեթուղթ",,,attr)
  DocLevel = 1
  DocNum = "RS-0000011"
  Call LetterOfCredit_Filter_Fill(FolderName, DocLevel, DocNum)
  
  Count = 9
  For i = 0 To Count-1
    If i <> 4 and i <> 5 Then 
      Call OnClick(arrayWaitForDoc(i), "frmASDocForm")
    End If  
  Next
  Call OnClick(c_Opers & "|" & c_SecPrCorr, "frmASDocForm")
  
  Count = 7
  For i = 0 To Count-1
    If i <> 2 Then
      Call OnClick(arrayWaitForView(i), "AsView")
    End If  
  Next 
  Call OnClick(c_ViewEdit & "|" & c_Other & "|" & c_PriceCorrDates, "AsView")
  Call OnClick(c_RevRepoSells, "AsView")
  
  Call OnClick(c_References & "|" & c_CommView, "FrmSpr")
  
  'Ջնջել   
  Call OnClick(c_Delete, "frmDeleteDoc")
  
  'Պայմանագրի փակում
  Call OnClick(c_AgrClose, "frmAsUstPar")

  wMDIClient.VBObject("frmPttel").Close

    
  FolderName = "|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|ä³ÛÙ³Ý³·ñ»ñ|î»Õ³µ³ßËí³Í ÙÇçáóÝ»ñ|àã å»ï³Ï³Ý ³ñÅ»ÃÕÃ»ñ|"
  
  ''4.Ոչ պետական արժեթղթեր/Տոկոսային եկ.բերող արժեթուղթ
  Call Log.Message("Ոչ պետական արժեթղթեր/Տոկոսային եկ.բերող արժեթուղթ",,,attr)
  DocLevel = 1
  DocNum = "000053"
  Call LetterOfCredit_Filter_Fill(FolderName, DocLevel, DocNum)
  
  ReDim arrayWaitForDoc1(6)          'պետք է բացվի Doc
  arrayWaitForDoc1 = Array(c_Folders & "|" & c_CurrentSchedule,_ 
                            c_Opers & "|" & c_Store & "|" & c_Store, c_Opers & "|" & c_SecPrCorr,_
                            c_TermsStates & "|" & c_Risking & "|" & c_ObjRiskCat,_
                            c_TermsStates & "|" & c_Risking & "|" & c_RiskCatPerRes,_
                            c_TermsStates & "|" & c_Percentages & "|" & c_Percentages)
  ReDim arrayWaitForView1(5) 
  arrayWaitForView1 = Array(c_Folders & "|" & c_ClFolder, c_Folders & "|" & c_SchFolder,_
                            c_References & "|" & c_CheckInterest,_
                            c_ViewEdit & "|" & c_Other & "|" & c_PriceCorrDates,_
                            c_ViewEdit & "|" & c_Risking & "|" & c_RisksPersRes,_
                            c_ViewEdit & "|" & c_Risking & "|" & c_ObjRiskCat,_
                            c_RevRepoSells)                           
                            
  Count = 9
  For i = 0 To Count-1
    If arrayWaitForDoc(i) <> c_Opers & "|" & c_Reclassification Then
      Call OnClick(arrayWaitForDoc(i), "frmASDocForm")
    End If  
  Next
  
  Count = 6
  For i = 0 To Count-1
    Call OnClick(arrayWaitForDoc1(i), "frmASDocForm")
  Next

  Count = 7
  For i = 0 To Count-1
    Call OnClick(arrayWaitForView(i), "AsView")
  Next
  Call OnClick(c_ViewEdit & "|" & c_Percentages & "|" & c_Percentages, "AsView")
  
  Count = 5
  For i = 0 To Count-1
    Call OnClick(arrayWaitForView1(i), "AsView")
  Next
  
  wMDIClient.VBObject("frmPttel").Close
  
  ''5.Ոչ պետական արժեթղթեր/Տոկոսային եկ.չբերող արժեթուղթ
  Call Log.Message("Ոչ պետական արժեթղթեր/Տոկոսային եկ.բերող արժեթուղթ",,,attr)
  DocLevel = 1
  DocNum = "000052"
  Call LetterOfCredit_Filter_Fill(FolderName, DocLevel, DocNum)

  Count = 9
  For i = 0 To Count-1
    If arrayWaitForDoc(i) <> c_Opers & "|" & c_Reclassification Then
      Call OnClick(arrayWaitForDoc(i), "frmASDocForm")
    End If  
  Next

  Count = 6
  For i = 1 To Count-2
    Call OnClick(arrayWaitForDoc1(i), "frmASDocForm")
  Next
  
  Count = 7
  For i = 0 To Count-1
    Call OnClick(arrayWaitForView(i), "AsView")
  Next
  Call OnClick(c_ViewEdit & "|" & c_Dates & "|" & c_AgrDates, "AsView")
  Call OnClick(c_ViewEdit & "|" & c_Dates & "|" & c_PerDates, "AsView")
  
  Count = 5
  For i = 0 To Count-1
    If i <> 1 and i <> 2 Then
      Call OnClick(arrayWaitForView1(i), "AsView")
    End If  
  Next
  
  Call Close_AsBank()
End Sub 