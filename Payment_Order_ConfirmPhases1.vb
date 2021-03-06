Option Explicit
'USEUNIT Library_Common
'USEUNIT Library_Colour
'USEUNIT Payment_Order_ConfirmPhases_Library
'USEUNIT Online_PaySys_Library
'USEUNIT Constants

'Test Case N 165047

Sub Payment_Order_AllFerify_Test()
    BuiltIn.Delay(20000) 
    
    Dim fDATE, startDATE , data , office, department, docNumber, accDeb, acDBValue, chart, balAcc, accMask, accCur, accType, clientName, client
    Dim note1, note2, note3, branch, depart, acsType, cardNum, payer, epayer, taxCod , socCard, accCredit
    Dim receiver, eReceiver, summa, curr, aim , fISN, confInput, confPath, docExist, isDel, colN
    data = "211211"
    Utilities.ShortDateFormat = "yyyymmdd"
    startDATE = "20030101"
    fDATE = "20250101"
    confPath = "X:\Testing\Order confirm phases\AllVerify_New.txt"
    
    data = "220612"
    office = Null             
    department = Null
    accDeb = False
    acDBValue = "77700/000001100"
    chart = Null
    balAcc = "10"
    accMask = Null
    accCur = Null
    accType = Null
    clientName = Null
    client = Null
    note1 = Null
    note2 = Null
    note3 = Null
    branch = Null
    depart = Null
    acsType = Null
    cardNum = Null
    payer = "Petrosyan Petr"
    epayer = Null
    taxCod = Null
    socCard = Null
    accCredit = "10300/4200012    "
    receiver = "Mozart"
    eReceiver = Null
    summa = "1000"
    curr = "000"
    aim = "Kap chuni"
    
    'Test StartUp 
    Call Initialize_AsBank("bank", startDATE, fDATE)
    Call ChangeWorkspace(c_Admin)
    Call Insert_MyDocs()
    Login("ARMSOFT")
    
    'Î³ñ·³íáñáõÙÝ»ñÇ Ý»ñÙáõÍáõÙ
    Log.Message "Կարգավորումների ներմուծում", "", pmNormal, DivideColor
    confInput = Input_Config(confPath)
    If Not confInput Then
        Log.Error("The configuration doesn't input")
    End If
    
    'ì×³ñÙ³Ý Ñ³ÝÓÝ³ñ³ñ·ñÇ ëï»ÕÍáõÙ
    Log.Message "Վճարման հանձնարարագրի ստեղծում", "", pmNormal, DivideColor
    Call ChangeWorkspace(c_CustomerService)
    Call Online_PaySys_Go_To_Agr_WorkPapers("|Ð³×³Ëáñ¹Ç ëå³ë³ñÏáõÙ ¨ ¹ñ³Ù³ñÏÕ |²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ", data, data)
    Call PayOrder_Send_Fill(office, department, docNumber, data, accDeb, acDBValue, chart, balAcc, accMask, accCur, accType, clientName, client, _
                            note1, note2, note3, branch, depart, acsType, cardNum, payer, epayer, taxCod , socCard, accCredit, _
                            receiver, eReceiver, summa, curr, aim , fISN)
    
    'ö³ëï³ÃÕÃÇ í³í»ñ³óáõÙ
    Log.Message "Փաստաթղթի վավերացում", "", pmNormal, DivideColor
    BuiltIn.Delay(3000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_SendToVer)
    BuiltIn.Delay(3000)
    Call ClickCmdButton(5, "²Ûá")
    
    BuiltIn.Delay(3000)
    wMDIClient.vbObject("frmPttel").Close() 
    
    'ö³ëï³ÃÕÃÇ ³éÏ³ÛáõÃÛ³Ý ëïáõ·áõÙ 1-ÇÝ Ñ³ëï³ïáÕում
    Log.Message "Փաստաթղթի առկայության ստուգում 1-ին հաստատողում", "", pmNormal, DivideColor
    Login("VERIFIER")
    docExist = Online_PaySys_Check_Doc_In_Verifier(docNumber, data, data)
    If Not docExist Then
      Log.Error("The document with number " & docNumber & " doesn't exist in 1st verify documents")
      Exit Sub
    End If
    
    'ö³ëï³ÃÕÃÇ í³í»ñ³óáõÙ 1-ÇÝ Ñ³ëï³áïÕÇ ÏáÕÙÇó
    Log.Message "Փաստաթղթի վավերացում 1-ին հատատողի կողմից", "", pmNormal, DivideColor
    Call PaySys_Verify(True)
    BuiltIn.Delay(3000)
    wMDIClient.vbObject("frmPttel").Close() 
    
    '---------------------
    Login("ARMSOFT")
    Call ChangeWorkspace(c_CustomerService)
    Call Online_PaySys_Go_To_Agr_WorkPapers("|Ð³×³Ëáñ¹Ç ëå³ë³ñÏáõÙ ¨ ¹ñ³Ù³ñÏÕ |²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ", data, data)
    If wMDIClient.WaitVBObject("frmPttel", 3000).Exists Then 
      colN = wMDIClient.vbObject("frmPttel").GetColumnIndex("DOCNUM")
      If SearchInPttel("frmPttel", colN, docNumber) Then
        Call PaySys_Verify(True)
      End If
    Else
      Log.Error "Pttel doesn't opened", "", pmNormal, ErrorColor
    End If 
      
    BuiltIn.Delay(3000)
    wMDIClient.vbObject("frmPttel").Close() 
    '---------------------
    
    '²ñï³ùÇÝ ÷áË³ÝóáõÙÝ»ñáõÙ Ñ³ÝÓÝ³ñ³ñ·ñÇ ³éÏ³ÛáõÃÛ³Ý ëïáõ·áõÙ
    Log.Message "Արտաքին փոխանցումներում հանձնարարագրի առկայության ստուգում", "", pmNormal, DivideColor
    Call ChangeWorkspace(c_ExternalTransfers)
    docExist = PaySys_Check_Doc_In_ExternalTransfer_Folder(data, data , docNumber)
    If Not docExist Then
      Log.Error("The document with number " & docNumber & " must exist in external transfers folder")
      Exit Sub
    End If
    
    'Ð³ÝÓÝ³ñ³ñ·ñÇ áõÕ³ñÏáõÙ BankMail µ³ÅÇÝ
    Log.Message "Հանձնարարագրի ուղարկում BankMail բաժին", "", pmNormal, DivideColor
    Call PaySys_Sendto_BankMail()
    
    'ö³ëï³ïÃÕÃÇ ³éÏ³ÛáõÃÛ³Ý ëïáõ·áõÙ BankMail-Ç áõÕ³ñÏíáÕ ÷áË³ÝóáõÙÝ»ñ ÃÕÃ³å³Ý³ÏáõÙ
    Log.Message "Փաստաթղթի առկայության ստուգում BankMail-ի ուղարկվող փոխանցումներ թղթապանակում", "", pmNormal, DivideColor
    Login("BANKMAIL")
    docExist = PaySys_Check_Doc_In_BankMail_Folder(data, data , fISN)
    If Not docExist Then
      Log.Error("The document with ISN " & fISN & " must exsits in sending BankMail folder")
      Exit Sub
    End If
    
    'ö³ëï³ÃÕÃÇ Ñ»é³óáõÙ
    Log.Message "Փաստաթղթի հեռացում", "", pmNormal, DivideColor
    Call Paysys_Delete_Doc(False)
    Login("ARMSOFT")
    Call ChangeWorkspace(c_ExternalTransfers)
    
    docExist = PaySys_Check_Doc_In_ExternalTransfer_Folder(data, data , docNumber)
    If Not docExist Then
      Log.Error("After deleteing in BankMail the document with number " & docNumber & " must exist in external transfers folder " )
    Else
      Call PaySys_SendTo_Partial_Edit()
    End If
    
    Login("OPERATOR")
    docExist = Online_PaySys_Check_Doc_In_Workpapers(docNumber, data, data)
    If Not docExist Then
      Log.Error("After deleteing in external transfers folder the document with number " & docNumber & " must exist in workpapers " )
    Else
      Call Paysys_Delete_Doc(True)
    End If
    
    'Test CleanUp 
    Call Close_AsBank()
End Sub