Option Explicit
'USEUNIT Library_Common
'USEUNIT Payment_Order_ConfirmPhases_Library
'USEUNIT Online_PaySys_Library
'USEUNIT PayOrder_Receive_ConfirmPhases_Library
'USEUNIT Constants 

'Test Case ID 165485

Sub Payment_Order_Receive_Reject_Test()
    
    Dim fDATE, startDATE , data, aim ,payer, accCredit, accDeb, office, department, department1,docNumber
    Dim trAcc, receiver, summa, fISN, confInput, confPath, docExist , inspDocVerify 
    
    data = "211216"
    Utilities.ShortDateFormat = "yyyymmdd"
    startDATE = "20030101"       
    fDATE = "20250101"
    confPath = "X:\Testing\PayOrder confirm phases(Receive)\PayOrder_Receive_Reject.txt"
    data = Null
    office = Null
    department = Null
    accCredit = "77700/03485190101"
    payer = "Petrosyan Petros"
    accDeb = "12400/123450789  "
    receiver = Null      
    summa = "25000"
    aim = "For reject"
    trAcc = Null
    
    'Test StartUp 
    Call Initialize_AsBank("bank", startDATE, fDATE)
    
    'Կարգավորումների ներմուծում
    confInput = Input_Config(confPath)
    If Not confInput Then
        Log.Error("The configuration doesn't input")
    End If
    
    Call ChangeWorkspace(c_ExternalTransfers)
    'Վճարման հանձնարարգիր (ստ.)-ի ստեղծում
    Call wTreeView.DblClickItem("|²ñï³ùÇÝ ÷áË³ÝóáõÙÝ»ñÇ ²Þî|Üáñ ÷³ëï³ÃÕÃ»ñ|ì×³ñÙ³Ý Ñ³ÝÓÝ³ñ³ñ³·Çñ (ëï.)")
    Call PayOrder_Receive_Fill(office, department, docNumber, data, accDeb, payer, accCredit, receiver, summa, aim , trAcc, fISN)
    
    'Տպելու ձև պատուհանի փակում
    BuiltIn.Delay(1000)
    wMDIClient.vbObject("FrmSpr").Close
    
    'Փաստաթղթի առկայության ստուգում հաշվառման ենթակա ստացված հանձնարարգրերի թղթապանակում
    docExist = Check_Doc_In_UnderRegistration_Folder (docNumber, data, data)
    If docExist = False Then
        Log.Error("Document with specified ID " & docNumber & "doesn't exists in under registration folder")
        Exit Sub
    End If
    
    'Փաստաթղթի ուղարկում վերստուգման
    Call PaySys_Send_To_CheckUp()
    
    Login("DOUBLEINPUTOPERATOR")
    docExist = PaySys_Check_Doc_In_InspecdetDoc_Folder(docNumber)
    If Not docExist Then
        Log.Error("The document with number " & docNumber & "doesn't exist in inspected documents folder")
        Exit Sub
    End If
    
    'Փաստատթղթի կրկնակի մուտքագրում
    inspDocVerify = PaySys_Verify_Doc_In_InspecdetDoc_Folder(accCredit, Null)
    If Not inspDocVerify Then
        Log.Error("Wrong double input values ")
    End If
    
    'Փաստաթղթի վավերացում
    Login("VERIFIER")
    'Փաստաթղթի առկայության ստուգում 1-ին հաստատողի մոտ
    docExist = Online_PaySys_Check_Doc_In_Verifier(docNumber, data, data)
    If Not docExist Then
        Log.Error("The document with number " & docNumber & " doesn't exist in 1st verify documents")
        Exit Sub
    End If
    
    'Փաստաթղթի վավերացում 1-ին  հաստատողի կողմից
    Call PaySys_Verify(True)
    
    'Փաստաթղթի առկայության ստուգում 2-րդ հաստատողի մոտ
    Login("VERIFIER2")
    docExist = PaySys_Check_Doc_In_Verifier(docNumber, data, data, "|Ð³ëï³ïáÕ II ²Þî|Ð³ëï³ïíáÕ í×³ñ³ÛÇÝ ÷³ëï³ÃÕÃ»ñ")
    If Not docExist Then
        Log.Error("The document with number " & docNumber & "doesn't exist in 2nd verify documents")
        Exit Sub
    End If
    
    'Փաստաթղթի մերժում 2-րդ հաստատողի կողմից
    Call PaySys_Verify(False)
    
    Login("ARMSOFT")
    Call ChangeWorkspace(c_ExternalTransfers)
    
    'Փաստաթղթի առկայության ստուգում հաշվառման ենթակա ստացված հանձնարարգրերի թղթապանակում
    docExist = Check_Doc_In_UnderRegistration_Folder (docNumber, data, data)
    If docExist = False Then
        Log.Error("Document with specified ID " & docNumber & "doesn't exists in under registration folder")
        Exit Sub
    End If
    
    'Փաստաթղթի ուղարկում վերստուգման
    Call PaySys_Send_To_CheckUp()
    
    Login("DOUBLEINPUTOPERATOR")
    docExist = PaySys_Check_Doc_In_InspecdetDoc_Folder(docNumber)
    If Not docExist Then
        Log.Error("The document with number " & docNumber & "doesn't exist in inspected documents folder")
        Exit Sub
    End If
    
    'Փաստատթղթի կրկնակի մուտքագրում
    inspDocVerify = PaySys_Verify_Doc_In_InspecdetDoc_Folder(accCredit, Null)
    If Not inspDocVerify Then
        Log.Error("Wrong double input values ")
    End If
    
    'Փաստաթղթի առկայության ստուգում 1-ին հաստատողի մոտ
    Login("VERIFIER")
    docExist = Online_PaySys_Check_Doc_In_Verifier(docNumber, data, data)
    If Not docExist Then
        Log.Error("The document with number " & docNumber & " must exist in 1st verify documents")
        Exit Sub
    End If
    
    'Փաստաթղթի վավերացում 1-ին  հաստատողի կողմից
    Call PaySys_Verify(True)
    
    'Փաստաթղթի առկայության ստուգում 2-րդ հաստատողի մոտ
    Login("VERIFIER2")
    docExist = PaySys_Check_Doc_In_Verifier(docNumber, data, data, "|Ð³ëï³ïáÕ II ²Þî|Ð³ëï³ïíáÕ í×³ñ³ÛÇÝ ÷³ëï³ÃÕÃ»ñ")
    If Not docExist Then
        Log.Error("The document with number " & docNumber & "doesn't exist in 2nd verify documents")
        Exit Sub
    End If
    
    'Փաստաթղթի մերժում 2-րդ հաստաոտղի կողմից
    Call PaySys_Verify(True)
    
    'Փաստաթղթի առկայության ստուգում 3-րդ հաստատողի մոտ
    Login("VERIFIER3")
    docExist = PaySys_Check_Doc_In_Verifier(docNumber, data, data, "|Ð³ëï³ïáÕ III ²Þî|Ð³ëï³ïíáÕ í×³ñ³ÛÇÝ ÷³ëï³ÃÕÃ»ñ")
    If Not docExist Then
        Log.Error("The document with number " & docNumber & "doesn't exist in 2nd verify documents")
        Exit Sub
    End If
    
    'Փաստաթղթի հաստատում 3-րդ հաստաոտղի կողմից
    Call PaySys_Verify(True)
    
    Login("ARMSOFT")
    'Արտաքին փոխանցումների հաշվառված ստացված փոխանցումներ թղթապանակում հանձնարարագրի առկայության ստուգում
    Call ChangeWorkspace(c_ExternalTransfers)
    docExist = Check_Doc_In_Registered_Folder(docNumber , data, data)
    If Not docExist Then
        Log.Error("The document with number " & docNumber & " must exist in registered transfers folder")
        Exit Sub
    End If
    
    'Ճշտել մարումը գործողության կատարում
    Call Clarify_Fading()
    
    'Test CleanUp
    'Արտաքին փոխանցումների Տարանցիկ թղթապանկում հանձնարարգրի առկայության ստուգում
    docExist = Check_Doc_In_Transit_Folder(docNumber, data, data)
    If Not docExist Then
        Log.Error("After fadeing documnent from registered payment orders folder  " & docNumber & " must exist in transit folder " )
    Else
        'Փաստաթղթի ուղարկում մասնակի խմբագրման
        Call PaySys_SendTo_Partial_Edit()
    End If
    
    BuiltIn.Delay(1000)
    wMDIClient.VBObject("frmPttel").close
    
    'Փաստաթղթի առկայության ստուգում Մասնակի խմբագրվող հանձնարարագրեր թղթապանակում 
    docExist = Check_Doc_In_Partial_Edit_Folder(docNumber, data, data)
    If Not docExist Then
        Log.Error("After sending to partial edit from transit folder the documnent with number " & docNumber & " must exist in partial editing folder " )
    Else
        'Փաստաթղթի հեռացում 
        Call Paysys_Delete_Doc(False)
    End If
    
    'Test CleanUp
    Call Close_AsBank()
End Sub