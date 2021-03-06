Option Explicit
'USEUNIT Library_Common
'USEUNIT Online_PaySys_Library
'USEUNIT Online_PaySys_Send_Library
'USEUNIT Payment_Library

Sub Online_PaySys_Send_Reject_Establish_Test()
    
    Dim startDATE, fDATE , fDOCNUM , fBASE
    Utilities.ShortDateFormat = "yyyymmdd"
    startDATE = "20101016"
    fDATE = "20111221"
    'Test StartUp start
    Call Initialize_AsBank("bank", startDATE, fDATE)
    Login("OPERATOR")
    'Test StartUp end
    
    'ä³ÛÙ³Ý³·ñÇ ëï»ÕÍáõÙ ë¨³·Çñ ·áñÍáÕáõÃÛ³Ùµ
    Call Online_PaySys_Go_To_Agr_WorkPapers("|Ð³×³Ëáñ¹Ç ëå³ë³ñÏáõÙ ¨ ¹ñ³Ù³ñÏÕ |²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ", Null, Null)
    Call Online_PaySys_Prepare_To_Create(2)
    Call Online_PaySys_Send_Fill("T", Null, Null, "Poxos Hastatvoxyan", "AH987654", Null, Null, Null, _
                                 "Petros Hastatvoxyan", "AM", "200000", Null , Null , Null , "030211", _
                                 "001" , Null, Null, Null , Null , Null, Null, Null , 1 , fDOCNUM, fBASE)
    
    'ê¨³·ñ»ñ ÃÕÃ³å³Ý³ÏÇó ÷³ëï³ÃÕÃÇ áõÕ³ñÏáõÙ ³ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ ÃÕÃ³å³Ý³Ï
    Call wTreeView.DblClickItem("|Ð³×³Ëáñ¹Ç ëå³ë³ñÏáõÙ ¨ ¹ñ³Ù³ñÏÕ |ÂÕÃ³å³Ý³ÏÝ»ñ|ú·ï³·áñÍáÕÇ ë¨³·ñ»ñ")
    Sys.Process("Asbank").vbObject("frmAsUstPar").vbObject("CmdOK").Click()
    BuiltIn.Delay(delay_middle)
    
    Do Until wMDIClient.vbObject("frmPttel").vbObject("tdbgView").EOF
        
        If Trim(wMDIClient.vbObject("frmPttel").vbObject("tdbgView").Columns.Item(2).Text) = Trim(fBASE) Then
            BuiltIn.Delay(delay_middle)
            Call wMainForm.MainMenu.Click("¶áñÍáÕáõÃÛáõÝÝ»ñ|´áÉáñ ·áñÍáÕáõÃÛáõÝÝ»ñÁ . . .")
            BuiltIn.Delay(delay_middle)
            Call wMainForm.PopupMenu.Click("ÊÙµ³·ñ»É")
            BuiltIn.Delay(delay_middle)
            Call Online_PaySys_Send_Fill("T", Null, Null, null, Null, Null, Null, Null, _
                                         Null, Null, Null, Null , Null , Null , Null, _
                                         "001" , Null, Null, Null , Null , Null, Null, Null , 0 , fDOCNUM, fBASE)
            Exit Do
        Else
            Call wMDIClient.vbObject("frmPttel").vbObject("tdbgView").MoveNext
        End If
    Loop
    
    'ä³ÛÙ³Ý³·ñÇ áõÕ³ñÏáõÙ ¹ñ³Ù³ñÏÕ
    Call Online_PaySys_Go_To_Agr_WorkPapers("|Ð³×³Ëáñ¹Ç ëå³ë³ñÏáõÙ ¨ ¹ñ³Ù³ñÏÕ |²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ", Null, Null)
    Call Online_PaySys_Send_To_Verify(1, fDOCNUM)
    
    'ä³ÛÙ³Ý³·ñÇ í³í»ñ³óáõÙ ¹ñ³Ù³ñÏÕáõÙ
    Login("CASHIER")
    Call Online_PaySys_Go_To_Agr_WorkPapers("|¸ñ³Ù³ñÏÕ|ö³ëï³ÃÕÃ»ñ ¹ñ³Ù³ñÏÕáõÙ", Null, Null)
    Call Online_PaySys_Send_Back_From_Cash(fDOCNUM, True, True)
    
    'ä³ÛÙ³Ý³·ñÇ Ù»ñÅáõÙ 1 Ñ³ëï³ïáÕÇ ÏáÕÙÇó
    Login("VERIFIER")
    Call Online_PaySys_Verify(fDOCNUM , False)
    
    'ä³ÛÙ³Ý³·ñÇ áõÕ³ñÏáõÙ ¹ñ³Ù³ñÏÕ
    Login("OPERATOR")
    Call Online_PaySys_Go_To_Agr_WorkPapers("|Ð³×³Ëáñ¹Ç ëå³ë³ñÏáõÙ ¨ ¹ñ³Ù³ñÏÕ |²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ", Null, Null)
    Call Online_PaySys_Send_To_Verify(1, fDOCNUM)
    
    'ä³ÛÙ³Ý³·ñÇ í³í»ñ³óáõÙ ¹ñ³Ù³ñÏÕáõÙ
    Login("CASHIER")
    Call Online_PaySys_Go_To_Agr_WorkPapers("|¸ñ³Ù³ñÏÕ|ö³ëï³ÃÕÃ»ñ ¹ñ³Ù³ñÏÕáõÙ", Null, Null)
    Call Online_PaySys_Send_Back_From_Cash(fDOCNUM, True, False)
    
    'ä³ÛÙ³Ý³·ñÇ Ñ³ëï³ïáõÙ 1 Ñ³ëï³ïáÕÇ ÏáÕÙÇó
    Login("VERIFIER")
    Call Online_PaySys_Verify(fDOCNUM , True)
    
    'Test CleanUp start
    Login("OPERATOR")
    'ö³ëï³ÃÕÃÇ Ñ»é³óáõÙ
    Call Online_PaySys_Go_To_Agr_WorkPapers("|Ð³×³Ëáñ¹Ç ëå³ë³ñÏáõÙ ¨ ¹ñ³Ù³ñÏÕ |Ð³ßí³éí³Í í×³ñ³ÛÇÝ ÷³ëï³ÃÕÃ»ñ", Null, Null)
    Online_PaySys_Delete_Agr(fDOCNUM)
    Call Close_AsBank()
    'Test CleanUp end
    
End Sub