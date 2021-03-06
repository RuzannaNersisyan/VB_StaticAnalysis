Option Explicit

'USEUNIT Library_Common  
'USEUNIT Akreditiv_Library 
'USEUNIT Constants

'Test Case Id - 160422

Sub GivenPledge_Clicks_Test()
  Dim fDATE, sDATE, Count, i, DocLevel, DocNum, FolderName, CalcAcc
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
  Call ChangeWorkspace(c_GivenPledge)
  
  ReDim arrayWaitForDoc(5)          'պետք է բացվի Doc
  arrayWaitForDoc = Array(c_ToEdit, c_View, c_Opers & "|" & c_AdjRev,_
                           c_Opers & "|" & c_Return, c_Addition)
  ReDim arrayWaitForView(6)          'պետք է բացվի View
  arrayWaitForView = Array(c_DocumentLog, c_Folders & "|" & c_ClFolder,_
                            c_Folders & "|" & c_AgrFolder,_
                            c_Folders & "|" & c_CollAgrFolder,_
                            c_OpersView,_
                            c_AccEntries & "|" & c_ForOffBal)

  ''1.Գրավի պայմանագիր` Արժեթղթեր                           
  Call Log.Message("Գրավի պայմանագիր` Արժեթղթեր",,,attr)                                                      
  DocNum = "M10171"                          
  Call wTreeView.DblClickItem("|îñ³Ù³¹ñí³Í ·ñ³í|ä³ÛÙ³Ý³·ñ»ñ")
  With AsBank.VBObject("frmAsUstPar")
    .VBObject("TabFrame").VBObject("AsTypeFolder").VBObject("TDBMask").Keys(DocNum & "[Tab]")
    .VBObject("CmdOK").ClickButton
  End With
  
  Count = 5
  For i = 0 To Count-1
    Call OnClick(arrayWaitForDoc(i), "frmASDocForm")
  Next
  
  Count = 6
  For i = 0 To Count-1
    Call OnClick(arrayWaitForView(i), "AsView")
  Next 
  
  Call OnClick(c_References & "|" & c_CommView, "FrmSpr")
  
  'Ջնջել   
  Call OnClick(c_Delete, "frmDeleteDoc")
  
  'Պայմանագրի փակում
  Call OnClick(c_AgrClose, "frmAsUstPar")

  wMDIClient.VBObject("frmPttel").Close
  
  Call ChangeWorkspace(c_GivenDepPledge)
  
  ''2.Ավանադային գրավի պայմանագիր                           
  Call Log.Message("Ավանդային գրավի պայմանագիր",,,attr)                                                      
  DocNum = "M30000"                          
  Call wTreeView.DblClickItem("|îñ³Ù³¹ñí³Í ³í³Ý¹³ÛÇÝ ·ñ³í|ä³ÛÙ³Ý³·ñ»ñ")
  With AsBank.VBObject("frmAsUstPar")
    .VBObject("TabFrame").VBObject("AsTypeFolder").VBObject("TDBMask").Keys(DocNum & "[Tab]")
    .VBObject("CmdOK").ClickButton
  End With
  
  Count = 5
  For i = 0 To Count-1
    Call OnClick(arrayWaitForDoc(i), "frmASDocForm")
  Next
  Call OnClick(c_Opers & "|" & c_Give, "frmASDocForm")
  
  Count = 6
  For i = 0 To Count-1
    If i <> 3 Then
      Call OnClick(arrayWaitForView(i), "AsView")
    End If  
  Next 
  
  Call OnClick(c_References & "|" & c_CommView, "FrmSpr")
  
  'Ջնջել   
  Call OnClick(c_Delete, "frmDeleteDoc")
  
  'Պայմանագրի փակում
  Call OnClick(c_AgrClose, "frmAsUstPar")

  wMDIClient.VBObject("frmPttel").Close  
  
  Call ChangeWorkspace(c_GivenGuarantee)
  
  ReDim arrayWaitForDoc(10)          'պետք է բացվի Doc
  arrayWaitForDoc = Array(c_ToEdit, c_View,_
                           c_Safety & "|" & c_AgrOpen & "|" & c_Guarantee,_
                           c_Safety & "|" & c_AgrOpen & "|" & c_DepMort,_
                           c_InputPrimaryContract,_
                           c_Opers & "|" & c_Outstanding & "|" & c_AdjRev,_
                           c_Opers & "|" & c_Outstanding & "|" & c_Return,_
                           c_Opers & "|" & c_Store,_
                           c_TermsStates & "|" & c_Risking & "|" & c_RiskCatPerRes,_
                           c_TermsStates & "|" & c_Risking & "|" & c_ObjRiskCat)
  ReDim arrayWaitForView(7)          'պետք է բացվի View
  arrayWaitForView = Array(c_DocumentLog, c_Folders & "|" & c_ClFolder,_
                            c_Folders & "|" & c_AgrFolder,_
                            c_OpersView,_
                            c_ViewEdit & "|" & c_Risking & "|" & c_RisksPersRes,_
                            c_ViewEdit & "|" & c_Risking & "|" & c_ObjRiskCat,_
                            c_AccEntries & "|" & c_ForOffBal)
  ReDim arrayWaitForModalBrowser(4)   'պետք է բացվի ModalBrowser
  arrayWaitForModalBrowser = Array(c_Safety & "|" & c_AgrOpen & "|" & c_Mortgage, c_Safety & "|" & c_AgrBind & "|" & c_Mortgage,_
                                    c_Safety & "|" & c_AgrBind & "|" & c_Guarantee, c_Safety & "|" & c_AgrBind & "|" & c_DepMort)
                            
  ''2.Երաշխավաորություն                           
  Call Log.Message("Երաշխավաորություն",,,attr)                                                      
  DocNum = "M20171"                          
  Call wTreeView.DblClickItem("|îñ³Ù³¹ñí³Í »ñ³ßË³íáñáõÃÛáõÝ|ä³ÛÙ³Ý³·ñ»ñ")
  With AsBank.VBObject("frmAsUstPar")
    .VBObject("TabFrame").VBObject("AsTypeFolder").VBObject("TDBMask").Keys(DocNum & "[Tab]")
    .VBObject("CmdOK").ClickButton
  End With
  
  Count = 10
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
  
  Call Close_AsBank()
End Sub 