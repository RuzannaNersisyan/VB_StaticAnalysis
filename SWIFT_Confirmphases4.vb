Option Explicit
'USEUNIT Library_Common
'USEUNIT Payment_Order_ConfirmPhases_Library
'USEUNIT Online_PaySys_Library
'USEUNIT Currency_Exchange_Confirmphases_Library
'USEUNIT CashInput_Confirmphases_Library
'USEUNIT SWIFT_Confirmphases_Library  
'USEUNIT Constants

'Test case ID 165598

Sub SWIFT_DifferConditions_Test()
   
    Dim fDATE, startDATE , docNumber, stockID, ref, opType, orgType1 , firstOrg, opType2 , secOrg, date1, date2
    Dim curr1, curr2, sendRec, summ, opType3, thirdOrg, recOrg , opType4, fourthOrg, fISN
    Dim confInput, confPath, docExist, isDel, rCount, data , curr21, curr22
   
    data = Null
    Utilities.ShortDateFormat = "yyyymmdd"
    startDATE = "20030101"
    fDATE = "20250101"
    confPath = "X:\Testing\SWIFT confirm phases\SWIFT_DifferentConditions.txt"
    ref = "28101198"
    opType = "AMND"
    orgType1 = "A"
    firstOrg = "CITICZPXXXX"
    opType2 = "A"
    secOrg = "ABCBBSNSXXX"
    date1 = aqConvert.DateTimeToStr(aqDateTime.Today)
    date2 = aqConvert.DateTimeToStr(aqDateTime.Today + 12)
    curr1 = "001"
    curr2 = "003"
    curr21 = "000"
    curr22 = "003"
    sendRec = "CITIATWXXXX"
    summ = "250000"
    opType3 = "A"
    thirdOrg = "CITIATWXXXX"
    opType4 = "A"
    fourthOrg = "CITIATWXXXX"
    
    'Test StartUp 
    Call Initialize_AsBank("bank", startDATE, fDATE)
    
    'Î³ñ·³íáñáõÙÝ»ñÇ Ý»ñÙáõÍáõÙ
    confInput = Input_Config(confPath)
    If Not confInput Then
        Log.Error("The configuration doesn't input")
    End If
    
    Call SetParameter("SWIN", "X:\Testing\SWIFT\IN\")
'     Login("SWIFT")
'    
'    Call wTreeView.DblClickItem("|S.W.I.F.T. ²Þî                  |Üáñ Ñ³Õáñ¹³·ñáõÃÛáõÝ|²ñï³ñÅáõÛÃÇ ÷áË³Ý³ÏÙ³Ý Ñ³ëï³ïáõÙ (Ðî 300)")
'     Call SWIFT_Doc_Fill(docNumber, ref, opType, orgType1 , firstOrg, opType2 , secOrg, date1, date2, _
'                        curr1, curr2, sendRec, summ, opType3, thirdOrg, opType4, fourthOrg, fISN )
'    
'    Login("ARMSOFT")
'    
'    Call ChangeWorkspace("S.W.I.F.T. ²Þî")
'    docExist = SWIFT_Check_Doc_In_Sending_SecrOrd_Folder(fISN)
'    If Not docExist Then
'        Log.Error("The document with number " & docNumber & " doesn't exist in workpapers folder")
'        Exit Sub
'    End If
'    'ö³ëï³ÃÕÃÇ áõÕ³ñÏáõÙ Ñ³ëï³ïÙ³Ý
'    Call Online_PaySys_Send_To_Verify(2)
'    
'    
'    Login("VERIFIER")
'    '1-ÇÝ ö³ëï³ÃÕÃÇ ³éÏ³ÛáõÃÛ³Ý ëïáõ·áõÙ 1-ÇÝ Ñ³ëï³ïáÕÇ Ùáï
'    docExist = Online_PaySys_Check_Doc_In_Verifier(docNumber, data, data)
'    If Not docExist Then
'        Log.Error("The document with number " & docNumber & " doesn't exist in 1st verify documents")
'        Exit Sub
'    End If
'    '1-ÇÝ ö³ëï³ÃÕÃÇ í³í»ñ³óáõÙ 1-ÇÝ Ñ³ëï³ïáÕÇ ÏáÕÙÇó
'    Call PaySys_Verify(True)
'    
'
'    Login("ARMSOFT")
'    Call ChangeWorkspace("S.W.I.F.T. ²Þî")
'    docExist = SWIFT_Check_Doc_In_Sending_SecrOrd_Folder(fISN)
'    If Not docExist Then
'        Log.Error("The document with number " & docNumber & " doesn't exist in workpapers folder")
'        Exit Sub
'    End If
'    '1-ÇÝ ö³ëï³ÃÕÃÇ Ñ»é³óáõÙ
'    Call Online_PaySys_Delete_Agr()
'    
'    Sys.Process("Asbank").vbObject("MainForm").Window("MDIClient", "", 1).vbObject("frmPttel").Close()
     
     Login("SWIFT")  
     Call wTreeView.DblClickItem("|S.W.I.F.T. ²Þî                  |Üáñ Ñ³Õáñ¹³·ñáõÃÛáõÝ|²ñï³ñÅáõÛÃÇ ÷áË³Ý³ÏÙ³Ý Ñ³ëï³ïáõÙ (Ðî 300)")
    Call SWIFT_Doc_Fill(docNumber, ref, opType, orgType1 , firstOrg, opType2 , secOrg, date1, date2, _
                        curr21, curr22, sendRec, summ, opType3, thirdOrg, opType4, fourthOrg, fISN )
    
    Login("ARMSOFT")
    
    Call ChangeWorkspace(c_SWIFT)
    Log.Message(fISN)
    docExist = SWIFT_Check_Doc_In_Sending_SecrOrd_Folder(fISN)
    If Not docExist Then
        Log.Error("The document with number " & docNumber & " doesn't exist in workpapers folder")
        Exit Sub
    End If
    
    'ö³ëï³ÃÕÃÇ áõÕ³ñÏáõÙ Ñ³ëï³ïÙ³Ý
    Call Online_PaySys_Send_To_Verify(2)
    
    Login("VERIFIER")
    '2-ñ¹ ö³ëï³ÃÕÃÇ ³éÏ³ÛáõÃÛ³Ý ëïáõ·áõÙ 1-ÇÝ Ñ³ëï³ïáÕÇ Ùáï
    docExist = Online_PaySys_Check_Doc_In_Verifier(docNumber, data, data)
    If Not docExist Then
        Log.Error("The document with number " & docNumber & " doesn't exist in 1st verify documents")
        Exit Sub
    End If
    
    '2-ñ¹  ö³ëï³ÃÕÃÇ í³í»ñ³óáõÙ 1-ÇÝ Ñ³ëï³ïáÕÇ ÏáÕÙÇó
    Call PaySys_Verify(True)
    
    Login("ARMSOFT")
    Call ChangeWorkspace(c_SWIFT)
    docExist = SWIFT_Check_Doc_In_Sending_SecrOrd_Folder(fISN)
    If Not docExist Then
        Log.Error("The document with number " & docNumber & " doesn't exist in workpapers folder")
        Exit Sub
    End If
    
    '2-ñ¹ ÷³ëï³ÃÕÃÇ Ñ»é³óáõÙ
    Call Online_PaySys_Delete_Agr()
    
    'Test CleanUp
    Call Close_AsBank()
End Sub