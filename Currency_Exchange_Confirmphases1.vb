Option Explicit
'USEUNIT Library_Common
'USEUNIT Payment_Order_ConfirmPhases_Library
'USEUNIT Online_PaySys_Library
'USEUNIT Currency_Exchange_Confirmphases_Library
'USEUNIT Constants

'Test Case ID 165600

Sub Currency_Exchange_Allverify_Test()
    BuiltIn.Delay(20000)
    
    Dim fDATE, startDATE , data , office, department, docNumber, accDeb, accCred, cur1, cur2, summa1
    Dim aim , ptype, clientCode, clientName, fISN, draft
    Dim confInput, confPath, docExist, isDel, rCount
    
    data = "211211"
    Utilities.ShortDateFormat = "yyyymmdd"
    startDATE = "20030101"
    fDATE = "20250101"
    confPath = "X:\Testing\Currency exchange confirmphases\Currency_exchange_Allverify_New.txt"
    data = Null                 
    department = Null
    accDeb = null         
    accCred = Null
    cur1 = "001"
    cur2 = "003"
    summa1 = "250"
    aim = "To verify all exchanges"
    ptype = "01"
    clientCode = "00034851"
    clientName = Null
    draft = True
    
    'Test StartUp 
    Call Initialize_AsBank("bank", startDATE, fDATE)
    Call ChangeWorkspace(c_Admin)
    
    'Î³ñ·³íáñáõÙÝ»ñÇ Ý»ñÙáõÍáõÙ
    confInput = Input_Config(confPath)
    If Not confInput Then
        Log.Error("The configuration doesn't input")
    End If
    
    Call ChangeWorkspace(c_CustomerService)
    'ì×³ñÙ³Ý Ñ³ÝÓÝ³ñ³ñ·ñÇ ëï»ÕÍáõÙ
    Call Online_PaySys_Go_To_Agr_WorkPapers("|Ð³×³Ëáñ¹Ç ëå³ë³ñÏáõÙ ¨ ¹ñ³Ù³ñÏÕ |²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ", data, data)
    
    Call Currency_Exchange_Doc_Fill(department, docNumber, accDeb, accCred, cur1, cur2, summa1, aim , ptype, clientCode, clientName, fISN, draft)
    
    'îå»Éáõ Ó¨ å³ïáõÑ³ÝÇ ÷³ÏáõÙ
    BuiltIn.Delay(1000)
    wMDIClient.vbObject("frmPttel").Close
    
    'ö³ëï³ÃÕÃÇ ³éÏ³ÛáõÃÛ³Ý ëïáõ·áõÙ ë¨³·ñ»ñ ÃÕÃ³å³Ý³ÏáõÙ
    docExist = Online_PaySys_Check_Doc_In_Drafts(fISN)
    If Not docExist Then
        Log.Error("The document with number " & docNumber & " doesn't exist in drafts folder")
        Exit Sub
    End If
    
    Call Currency_Exchange_Verify_From_Draft(summa1)
    
    BuiltIn.Delay(1000)
    wMDIClient.vbObject("FrmSpr").Close
    BuiltIn.Delay(1000)
    wMDIClient.vbObject("frmPttel").Close
    
    docExist = Online_PaySys_Check_Doc_In_Workpapers(docNumber, data, data)
    If Not docExist Then
        Log.Error("The document with number " & docNumber & " doesn't exist in workpapers folder")
        Exit Sub
    End If
    
    'ö³ëï³ÃÕÃÇ í³í»ñ³óáõÙ
    Call PaySys_Verify(True)
    
    Login("VERIFIER")
    'ö³ëï³ÃÕÃÇ ³éÏ³ÛáõÃÛ³Ý ëïáõ·áõÙ 1-ÇÝ Ñ³ëï³ïáÕÇ Ùáï
    docExist = Online_PaySys_Check_Doc_In_Verifier(docNumber, data, data)
    If Not docExist Then
        Log.Error("The document with number " & docNumber & " doesn't exist in 1st verify documents")
        Exit Sub
    End If
    
    'ö³ëï³ÃÕÃÇ í³í»ñ³óáõÙ 1-ÇÝ Ñ³ëï³áïÕÇ ÏáÕÙÇó
    Call PaySys_Verify(True)
    
    Login("ARMSOFT")
    '¶ÉË³íáñ Ñ³ßí³å³ÑÇ ÁÝ¹Ñ³Ýáõñ ¹ÇïáõÙ ÃÕÃ³å³Ý³ÏáõÙ ÷³ëï³ÃÕÃÇ ³éÏ³ÛáõÃÛ³Ý ëïáõ·áõÙ
    Call ChangeWorkspace(c_ChiefAcc)
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