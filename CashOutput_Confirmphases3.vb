Option Explicit
'USEUNIT Library_Common
'USEUNIT Payment_Order_ConfirmPhases_Library
'USEUNIT Online_PaySys_Library
'USEUNIT Currency_Exchange_Confirmphases_Library
'USEUNIT CashInput_Confirmphases_Library
'USEUNIT CashOutput_Confirmpases_Library
'USEUNIT Constants

'Test case number - 165054

Sub CashOutput_Pass_Test()
  Dim fDATE, startDATE , docNumber, summa, fISN, draft, accTemp, data
  Dim confInput, confPath, docExist, isDel, rCount , inspDocVerify
    
  data = Null
  Utilities.ShortDateFormat = "yyyymmdd"
  startDATE = "20030101"
  fDATE = "20250101"
  confPath = "X:\Testing\CashOutput confirm phases\CashOutput_Pass.txt"
  accTemp = "03485190101"
  summa = "250000"
  draft = False   
  
  BuiltIn.Delay(20000)                
       
  'Test StartUp 
  Call Initialize_AsBank("bank", startDATE, fDATE)
    
  'Î³ñ·³íáñáõÙÝ»ñÇ Ý»ñÙáõÍáõÙ
  confInput = Input_Config(confPath)
  If Not confInput Then
  Log.Error("The configuration doesn't input")
  End If
    
  'ì×³ñÙ³Ý Ñ³ÝÓÝ³ñ³ñ·ñÇ ëï»ÕÍáõÙ
  Call ChangeWorkspace(c_ChiefAcc)
  Call CashOutput_Doc_Fill(docNumber, accTemp, summa, fISN, draft)
    
  'îå»Éáõ Ó¨ å³ïáõÑ³ÝÇ ÷³ÏáõÙ
  BuiltIn.Delay(1000)
  wMDIClient.vbObject("FrmSpr").Close
    
  BuiltIn.Delay(1000)
  wMDIClient.vbObject("frmPttel").Close
    
  '²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ ÃÕÃ³å³Ý³ÏáõÙ ÷³ëï³ÃÕÃÇ ³éÏ³ÛáõÃÛ³Ý ëïáõ·áõÙ
  Call ChangeWorkspace(c_CustomerService)
  docExist = Online_PaySys_Check_Doc_In_Workpapers(docNumber, data, data)
  If Not docExist Then
    Log.Error("The document with number " & docNumber & " doesn't exist in workpapers folder")
    Exit Sub
  End If
    
  'ö³ëï³ÃÕÃÇ áõÕ³ñÏáõÙ ÏñÏÝ³ÏÇ Ùáõïù³·ñÙ³Ý
  Call CashInput_Send_To_CheckUp()
    
  'ö³ëï³ÃÕÃÇ í³í»ñ³óáõÙ í»ñëïáõ·áÕÇ ÷³ëï³ÃÕÃ»ñÇó
  Login("DOUBLEINPUTOPERATOR")
  docExist = PaySys_Check_Doc_In_InspecdetDoc_Folder(docNumber)
  If Not docExist Then
    Log.Error("The document with number " & docNumber & "doesn't exist in inspected documents folder")
    Exit Sub
  End If
    
  'ö³ëï³ïÃÕÃÇ ÏñÏÝ³ÏÇ Ùáõïù³·ñáõÙ
  inspDocVerify = CashOutput_Verify_Doc_In_InspecdetDoc_Folder(accTemp, summa)
  If Not inspDocVerify Then
    Log.Error("Wrong double input values ")
    Exit Sub
  End If
    
  'ö³ëï³ÃÕÃÇ ³éÏ³ÛáõÃÛ³Ý ëïáõ·áõÙ 1-ÇÝ Ñ³ëï³ïáÕÇ Ùáï
  Login("VERIFIER")
  docExist = Online_PaySys_Check_Doc_In_Verifier(docNumber, data, data)
  If docExist Then
    Log.Error("The document with number " & docNumber & " mustn't exist in 1st verify documents")
    Exit Sub
  End If
    
  'ö³ëï³ÃÕÃÇ ³éÏ³ÛáõÃÛ³Ý ëïáõ·áõÙ 2-ñ¹ Ñ³ëï³ïáÕÇ Ùáï
  Login("VERIFIER2")
  docExist = PaySys_Check_Doc_In_Verifier(docNumber, data, data, "|Ð³ëï³ïáÕ II ²Þî|Ð³ëï³ïíáÕ í×³ñ³ÛÇÝ ÷³ëï³ÃÕÃ»ñ")
  If Not docExist Then
    Log.Error("The document with number " & docNumber & "doesn't exist in 2nd verify documents")
    Exit Sub
  End If
    
  'ö³ëï³ÃÕÃÇ í³í»ñ³óáõÙ 2-ñ¹ Ñ³ëï³ïáÕÇ ÏáÕÙÇó
  Call PaySys_Verify(True)
    
  '¶ÉË³íáñ Ñ³ßí³å³ÑÇ ÁÝ¹Ñ³Ýáõñ ¹ÇïáõÙ ÃÕÃ³å³Ý³ÏáõÙ ÷³ëï³ÃÕÃÇ ³éÏ³ÛáõÃÛ³Ý ëïáõ·áõÙ
  Login("ARMSOFT")
  Call ChangeWorkspace(c_ChiefAcc)
  BuiltIn.Delay(6000)
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