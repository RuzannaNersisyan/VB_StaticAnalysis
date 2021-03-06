Option Explicit

'USEUNIT Library_Common  
'USEUNIT Akreditiv_Library 
'USEUNIT Constants

'Test Case Id - 160492

Sub ReceivedPledge_Clicks_Test()
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
  Call ChangeWorkspace(c_RecPledge)
  
  ReDim arrayWaitForDoc(4)          'պետք է բացվի Doc
  arrayWaitForDoc = Array(c_ToEdit, c_View, c_Opers & "|" & c_AdjRev,_
                           c_Opers & "|" & c_Return)
  ReDim arrayWaitForView(6)          'պետք է բացվի View
  arrayWaitForView = Array(c_DocumentLog, c_Folders & "|" & c_ClFolder,_
                            c_Folders & "|" & c_AgrFolder,_
                            c_Folders & "|" & c_CollAgrFolder,_
                            c_OpersView,_
                            c_AccEntries & "|" & c_ForOffBal)
  
  ''1.Այլ գրավ                            
  Call Log.Message("Այլ գրավ",,,attr)                                                      
  DocNum = "N10199"                          
  Call wTreeView.DblClickItem("|êï³óí³Í ·ñ³í|ä³ÛÙ³Ý³·ñ»ñ")
  With AsBank.VBObject("frmAsUstPar")
    .VBObject("TabFrame").VBObject("AsTypeFolder").VBObject("TDBMask").Keys(DocNum & "[Tab]")
    .VBObject("CmdOK").ClickButton
  End With
  
  Count = 4
  For i = 0 To Count-1
    Call OnClick(arrayWaitForDoc(i), "frmASDocForm")
  Next
  
  Count = 6
  For i = 0 To Count-1
    BuiltIn.Delay(100)
    Call OnClick(arrayWaitForView(i), "AsView")
  Next 
  
  Call OnClick(c_References & "|" & c_CommView, "FrmSpr")
  
  'Ջնջել   
  Call OnClick(c_Delete, "frmDeleteDoc")
  
  'Պայմանագրի փակում
  Call OnClick(c_AgrClose, "frmAsUstPar")

  wMDIClient.VBObject("frmPttel").Close
  
  ''2.Ոսկի, թանկարժեք իրեր                            
  Call Log.Message("Ոսկի, թանկարժեք իրեր",,,attr)                                                      
  DocNum = "N10191"                          
  Call wTreeView.DblClickItem("|êï³óí³Í ·ñ³í|ä³ÛÙ³Ý³·ñ»ñ")
  With AsBank.VBObject("frmAsUstPar")
    .VBObject("TabFrame").VBObject("AsTypeFolder").VBObject("TDBMask").Keys(DocNum & "[Tab]")
    .VBObject("CmdOK").ClickButton
  End With
  
  Count = 4
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
  
  ''3.Արժեթղթեր                            
  Call Log.Message("Արժեթղթեր",,,attr)                                                      
  DocNum = "N10182"                          
  Call wTreeView.DblClickItem("|êï³óí³Í ·ñ³í|ä³ÛÙ³Ý³·ñ»ñ")
  With AsBank.VBObject("frmAsUstPar")
    .VBObject("TabFrame").VBObject("AsTypeFolder").VBObject("TDBMask").Keys(DocNum & "[Tab]")
    .VBObject("CmdOK").ClickButton
  End With
  
  Count = 4
  For i = 0 To Count-1
    Call OnClick(arrayWaitForDoc(i), "frmASDocForm")  
  Next
  Call OnClick(c_Addition, "frmASDocForm")
  
  Count = 6
  For i = 0 To Count-1
    If i <> 1 Then
      Call OnClick(arrayWaitForView(i), "AsView")
    End If  
  Next 
  
  Call OnClick(c_References & "|" & c_CommView, "FrmSpr")
  
  'Ջնջել   
  Call OnClick(c_Delete, "frmDeleteDoc")
  
  'Պայմանագրի փակում
  Call OnClick(c_AgrClose, "frmAsUstPar")

  wMDIClient.VBObject("frmPttel").Close
  
  ''4.Գրավի պայմանագիր                            
  Call Log.Message("Գրավի պայմանագիր",,,attr)                                                      
  DocNum = "N10224"                          
  Call wTreeView.DblClickItem("|êï³óí³Í ·ñ³í|ä³ÛÙ³Ý³·ñ»ñ")
  With AsBank.VBObject("frmAsUstPar")
    .VBObject("TabFrame").VBObject("AsTypeFolder").VBObject("TDBMask").Keys(DocNum & "[Tab]")
    .VBObject("CmdOK").ClickButton
  End With
  
  Count = 4
  For i = 0 To Count-1
    If i <> 2 Then
      Call OnClick(arrayWaitForDoc(i), "frmASDocForm")
    Else 
      Call OnClick(c_Opers & "|" & c_Give, "frmASDocForm")  
    End If  
  Next
  Call OnClick(c_AgrBind & "|" & c_AgrBind, "frmASDocForm") 
  Call OnClick(c_AgrBind & "|" & c_CutAgrBind, "frmASDocForm") 
  
  Count = 6
  For i = 0 To Count-1
    Call OnClick(arrayWaitForView(i), "AsView")
  Next 
  Call OnClick(c_AgrBind & "|" & c_LinksOfAgreement, "AsView")
  
  Call OnClick(c_References & "|" & c_CommView, "FrmSpr")
  
  'Ջնջել   
  Call OnClick(c_Delete, "frmDeleteDoc")
  
  'Պայմանագրի փակում
  Call OnClick(c_AgrClose, "frmAsUstPar")

  wMDIClient.VBObject("frmPttel").Close
  
  Call ChangeWorkspace(c_RecDepPledge)
  ''5.Ավանդային գրավի պայմանագիր                            
  Call Log.Message("Ավանդային գրավի պայմանագիր",,,attr)                                                      
  DocNum = "000184"                          
  Call wTreeView.DblClickItem("|êï³óí³Í ³í³Ý¹³ÛÇÝ ·ñ³í|ä³ÛÙ³Ý³·ñ»ñ")
  With AsBank.VBObject("frmAsUstPar")
    .VBObject("TabFrame").VBObject("AsTypeFolder").VBObject("TDBMask").Keys(DocNum & "[Tab]")
    .VBObject("CmdOK").ClickButton
  End With
  
  Count = 4
  For i = 0 To Count-1
    Call OnClick(arrayWaitForDoc(i), "frmASDocForm")  
  Next
  Call OnClick(c_Addition, "frmASDocForm")
  Call OnClick(c_Opers & "|" & c_Give, "frmASDocForm")  
  
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
  
  Call ChangeWorkspace(c_RecGuarantee)
  ''5.Ավանդային գրավի պայմանագիր                            
  Call Log.Message("Երաշխավորություն",,,attr)                                                      
  DocNum = "N20175"                          
  Call wTreeView.DblClickItem("|êï³óí³Í »ñ³ßË³íáñáõÃÛáõÝ|ä³ÛÙ³Ý³·ñ»ñ")
  With AsBank.VBObject("frmAsUstPar")
    .VBObject("TabFrame").VBObject("AsTypeFolder").VBObject("TDBMask").Keys(DocNum & "[Tab]")
    .VBObject("CmdOK").ClickButton
  End With
  
  Count = 4
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

  Call Close_AsBank()
End Sub