Option Explicit
'USEUNIT Library_Common
'USEUNIT Payment_Order_ConfirmPhases_Library
'USEUNIT Online_PaySys_Library
'USEUNIT CashInput_Confirmphases_Library
'USEUNIT Currency_Exchange_Confirmphases_Library
'USEUNIT CashOutput_Confirmpases_Library
'USEUNIT Constants

'Test case ID 165605

Sub CashOutput_Allconditions_Test()
    BuiltIn.Delay(20000)
    
    Dim fDATE, startDATE , docNumber, summa, fISN, draft, accTemp, data
    Dim confInput, confPath, docExist, isDel, rCount
    
    data = null
    Utilities.ShortDateFormat = "yyyymmdd"
    startDATE = "20030101"
    fDATE = "20250101"
    confPath = "X:\Testing\CashOutput confirm phases\CashOutput_Allconditions.txt"
    accTemp = "33170160500"                
    summa = "220000"
    draft = False
    
    'Test StartUp 
    Call Initialize_AsBank("bank", startDATE, fDATE)
    
    'Î³ñ·³íáñáõÙÝ»ñÇ Ý»ñÙáõÍáõÙ
    confInput = Input_Config(confPath)
    If Not confInput Then
        Log.Error("The configuration doesn't input")
    End If
    
    Call ChangeWorkspace(c_ChiefAcc)
    'ì×³ñÙ³Ý Ñ³ÝÓÝ³ñ³ñ·ñÇ ëï»ÕÍáõÙ
    Call CashOutput_Doc_Fill(docNumber, accTemp, summa, fISN, draft)
    Call ClickCmdButton(5, "Î³ï³ñ»É")
    
    'îå»Éáõ Ó¨ å³ïáõÑ³ÝÇ ÷³ÏáõÙ
    BuiltIn.Delay(1000)
    wMDIClient.vbObject("FrmSpr").Close
    BuiltIn.Delay(1000)
    wMDIClient.vbObject("frmPttel").Close
    
    'ö³ëï³ÃÕÃÇ áõÕ³ñÏáõÙ Ñ³ëï³ïÙ³Ý ë¨ óáõó³Ï
    Call ChangeWorkspace(c_CustomerService)
    'ö³ëï³ÃÕÃÇ ³éÏ³ÛáõÃÛ³Ý ²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ ÃÕÃ³å³Ý³ÏáõÙ
    docExist = Online_PaySys_Check_Doc_In_Workpapers(docNumber, null, Null)
    If Not docExist Then
        Log.Error("The document with number " & docNumber & " doesn't exist in workpaper documents")
        Exit Sub
    End If
    
    Call Online_PaySys_Send_To_Verify(3)
    
    'ê¨ óáõó³ÏÇó ÷³ëï³ÃÕÃÇ Ñ³ëï³ïáõÙ Ñ³ÙÁÝÏÝáõÙÝ»ñÇ Ù³ëÇÝ ÇÝýáñÙ³óÇ³Ý ëïáõ·»Éáõó Ñ»ïá
    Call ChangeWorkspace(c_BLVerifyer)
    docExist = Online_PaySys_Check_Doc_In_Black_List(docNumber)
    If docExist = False Then
        Log.Error("Document with specified ID " & docNumber & "doesn't exists in Black list folder")
        Exit Sub
    End If
    
    rCount = Online_PaySys_Check_Assertion_In_Black_List()
    If rCount <> 2 Then
        Log.Error("There must be 2 row")
        Exit Sub
    End If
    
    Call PaySys_Verify( True)
    
    'ö³ëï³ÃÕÃÇ í³í»ñ³óáõÙ
    Login("VERIFIER")
    
    'ö³ëï³ÃÕÃÇ ³éÏ³ÛáõÃÛ³Ý ëïáõ·áõÙ 1-ÇÝ Ñ³ëï³ïáÕÇ Ùáï
    docExist = Online_PaySys_Check_Doc_In_Verifier(docNumber, null, Null)
    If Not docExist Then
        Log.Error("The document with number " & docNumber & " doesn't exist in 1st verify documents")
        Exit Sub
    End If
    
    'ö³ëï³ÃÕÃÇ í³í»ñ³óáõÙ 1-ÇÝ Ñ³ëï³áïÕÇ ÏáÕÙÇó
    Call PaySys_Verify(True)
    
    Login("ARMSOFT")
    
    '¶ÉË³íáñ Ñ³ßí³å³ÑÇ ÁÝ¹Ñ³Ýáõñ ¹ÇïáõÙ ÃÕÃ³å³Ý³ÏáõÙ ÷³ëï³ÃÕÃÇ ³éÏ³ÛáõÃÛ³Ý ëïáõ·áõÙ
    Call ChangeWorkspace(c_ChiefAcc)
    Log.Message(fISN)
    docExist = Check_Doc_In_GeneralView_Folder(fISN)
    If Not docExist Then
        Log.Error("The document with number " & fISN & " must exist in general view folder")
        Exit Sub
    End If
    
    'ö³ëï³ÃÕÃÇ Ñ»é³óáõÙ
    Call Online_PaySys_Delete_Agr()
    
    'Test CleanUp
    Call Close_AsBank()
End Sub