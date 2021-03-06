Option Explicit
'USEUNIT Library_Common
'USEUNIT Payment_Order_ConfirmPhases_Library
'USEUNIT Online_PaySys_Library
'USEUNIT PayOrder_Receive_ConfirmPhases_Library
'USEUNIT International_PayOrder_Receive_Confirmphases_Library
'USEUNIT Constants

'Test Case ID 165626

Sub International_PayOrder_Receive_Allconditions_Test()
    BuiltIn.Delay(20000)
    
    Dim fDATE, startDATE , data, payer, office, department, docNumber
    Dim receiver, summa, fISN, confInput, confPath, docExist , inspDocVerify
    Dim payerAcc, IBAN, country, acc, transAcc, recAcc, curr, recCorrBank, rCount
    Dim     recInfo, payerAddr, aim 
    
    data = "211217"
    Utilities.ShortDateFormat = "yyyymmdd"
    startDATE = "20030101"
    fDATE = "20250101"
    confPath = "X:\Testing\International (Receive) ConfPhases\International_PayOrder_Receive_Allconditions.txt"                  
    data = aqConvert.DateTimeToFormatStr(aqDateTime.Today(), "%d/%m/%y")    
    office = Null
    department = Null
    payerAcc = Null
    payer = "MOUSTAFA ABBES"
    IBAN = True
    country = "CZ"
    acc = "11111111111111111111"
    recAcc = "77700/03485190101"
    receiver = Null
    summa = "250000"
    curr = "001"
    recCorrBank = "CITIATWXXXX"
    transAcc = Null    
    recInfo = Null
    payerAddr = Null
    aim = "npatak"
    
    'Test StartUp 
    Call Initialize_AsBank("bank", startDATE, fDATE)
    Call ChangeWorkspace(c_Admin)
    
    'Î³ñ·³íáñáõÙÝ»ñÇ Ý»ñÙáõÍáõÙ
    confInput = Input_Config(confPath)
    If Not confInput Then
        Log.Error("The configuration doesn't input")
    End If
    
    Call ChangeWorkspace(c_ExternalTransfers)
    'ØÇç. í×³ñÙ³Ý Ñ³ÝÓÝ³ñ³ñ³·Çñ (ëï.)-Ç ëï»ÕÍáõÙ
    Call wTreeView.DblClickItem("|²ñï³ùÇÝ ÷áË³ÝóáõÙÝ»ñÇ ²Þî|Üáñ ÷³ëï³ÃÕÃ»ñ|ØÇç. í×³ñÙ³Ý Ñ³ÝÓÝ³ñ³ñ³·Çñ (ëï.)")
    Call International_PayOrder_Recipient_Fill( fISN, office, department, docNumber, data, recAcc, receiver, recInfo, payerAcc, payer, payerAddr, country, acc,_
                                                                                        summa, curr, aim, recCorrBank, transAcc, IBAN )
    
    Call ClickCmdButton(5, "Î³ï³ñ»É")
    
    'îå»Éáõ Ó¨ å³ïáõÑ³ÝÇ ÷³ÏáõÙ
    BuiltIn.Delay(1000)
    wMDIClient.vbObject("FrmSpr").Close
    
    'ö³ëï³ÃÕÃÇ ³éÏ³ÛáõÃÛ³Ý ëïáõ·áõÙ Ñ³ßí³éÙ³Ý »ÝÃ³Ï³ ëï³óí³Í Ñ³ÝÓÝ³ñ³ñ·ñ»ñÇ ÃÕÃ³å³Ý³ÏáõÙ
    docExist = Check_Doc_In_UnderRegistration_Folder (docNumber, data, data)
    If docExist = False Then
        Log.Error("Document with specified ID " & docNumber & "doesn't exists in under registration folder")
        Exit Sub
    End If
    
    'ö³ëï³ÃÕÃÇ áõÕ³ñÏáõÙ Ñ³ëï³ïÙ³Ý ë¨ óáõó³Ï
    Call International_PayOrder_Send_To_BlackList()
    
    'ê¨ óáõó³ÏÇó ÷³ëï³ÃÕÃÇ Ñ³ëï³ïáõÙ Ñ³ÙÁÝÏÝáõÙÝ»ñÇ Ù³ëÇÝ ÇÝýáñÙ³óÇ³Ý ëïáõ·»Éáõó Ñ»ïá
    Call ChangeWorkspace(c_BLVerifyer)
    docExist = Online_PaySys_Check_Doc_In_Black_List(docNumber)
    If docExist = False Then
        Log.Error("Document with specified ID " & docNumber & "doesn't exists in Black list folder")
        Exit Sub
    End If
    
    rCount = Online_PaySys_Check_Assertion_In_Black_List()
    If rCount <> 1 Then
        Log.Error("There must be 1 row")
        Exit Sub
    End If
    
    Call PaySys_Verify( True)
    
    'ö³ëï³ÃÕÃÇ í³í»ñ³óáõÙ
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
    '²ñï³ùÇÝ ÷áË³ÝóáõÙÝ»ñÇ Ñ³ßí³éí³Í ëï³óí³Í ÷áË³ÝóáõÙÝ»ñ ÃÕÃ³å³Ý³ÏáõÙ Ñ³ÝÓÝ³ñ³ñ³·ñÇ ³éÏ³ÛáõÃÛ³Ý ëïáõ·áõÙ
    Call ChangeWorkspace(c_ExternalTransfers)
    log.Message(docNumber)
    docExist = Check_Doc_In_Registered_Folder(docNumber , data, data)
    If Not docExist Then
        Log.Error("The document with number " & docNumber & " must exist in registered transfers folder")
        Exit Sub
    End If
    
    'Ößï»É Ù³ñáõÙÁ ·áñÍáÕáõÃÛ³Ý Ï³ï³ñáõÙ
    Call Clarify_Fading()
    
    '²ñï³ùÇÝ ÷áË³ÝóáõÙÝ»ñÇ Ð³ßí³éÙ³Ý »ÝÃ³Ï³ ÃÕÃ³å³ÝÏáõÙ Ñ³ÝÓÝ³ñ³ñ·ñÇ ³éÏ³ÛáõÃÛ³Ý ëïáõ·áõÙ
    docExist = Check_Doc_In_UnderRegistration_Folder(docNumber, data, data)
    If Not docExist Then
        Log.Error("After fadeing order from registered payment orders folder with number " & docNumber & " must exist in under registration folder " )
    Else
        'ö³ëï³ÃÕÃÇ áõÕ³ñÏáõÙ Ù³ëÝ³ÏÇ ËÙµ³·ñÙ³Ý
        Call Paysys_Delete_Doc(False)
    End If
    
    'Test CleanUp
    Call Close_AsBank()
End Sub