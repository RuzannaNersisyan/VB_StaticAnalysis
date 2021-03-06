Option Explicit
'USEUNIT Library_Common
'USEUNIT Payment_Order_ConfirmPhases_Library
'USEUNIT Online_PaySys_Library
'USEUNIT PayOrder_Receive_ConfirmPhases_Library
'USEUNIT Constants

'Test Case ID 165482

Sub Payment_Order_Receive_Allconditions_Test()
    BuiltIn.Delay(20000)
    
    Dim fDATE, startDATE , data, aim ,payer, accCredit, accDeb, office, department, department1,docNumber
    Dim trAcc, receiver, summa, fISN, confInput, confPath, docExist, rCount 
    
    data = "211218"
    Utilities.ShortDateFormat = "yyyymmdd"
    startDATE = "20030101"
    fDATE = "20250101"
    confPath = "X:\Testing\PayOrder confirm phases(Receive)\PayOrder_Receive_AllConditions.txt"
    data = Null
    office = Null
    department = Null
    accCredit = "77700/03485190101"
    payer = "MOUSTAFA ABBES"
    accDeb = "12400/123450789  "
    receiver = Null
    summa = "15000"
    aim = "Sev cucak"
    trAcc = Null                       
    
    'Test StartUp 
    Call Initialize_AsBank("bank", startDATE, fDATE)
    
    'Կարգավորումների ներմուծում
    confInput = Input_Config(confPath)
    If Not confInput Then
        Log.Error("The configuration doesn't input")
        Exit Sub
    End If
    
    Call ChangeWorkspace(c_ExternalTransfers)
    'Վճարման հանձնարարգիր (ստ.)-ի ստեղծում
    Call wTreeView.DblClickItem("|²ñï³ùÇÝ ÷áË³ÝóáõÙÝ»ñÇ ²Þî|Üáñ ÷³ëï³ÃÕÃ»ñ|ì×³ñÙ³Ý Ñ³ÝÓÝ³ñ³ñ³·Çñ (ëï.)")
    Call PayOrder_Receive_Fill(office, department, docNumber, data, accDeb, payer, accCredit, receiver, summa, aim , trAcc, fISN)
    
    Call ClickCmdButton(5, "Î³ï³ñ»É")
    
    'Տպելու ձև պատուհանի փակում
    BuiltIn.Delay(2000)
    wMDIClient.vbObject("FrmSpr").Close
    
    'Փաստաթղթի առկայության ստուգում հաշվառման ենթակա ստացված հանձնարարգրերի թղթապանակում
    docExist = Check_Doc_In_UnderRegistration_Folder (docNumber, data, data)
    If docExist = False Then
        Log.Error("Document with specified ID " & docNumber & "doesn't exists in under registration folder")
        Exit Sub
    End If
    
    'Փաստաթղթի ուղարկում հաստատման սև ցուցակ
    Call Online_PaySys_Send_To_Verify(3)
    
    'Սև ցուցակից փաստաթղթի հաստատում համընկնումների մասին ինֆորմացիան ստուգելուց հետո
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
    
    'Փաստաթղթի վավերացում
    Login("VERIFIER")
    'Փաստաթղթի առկայության ստուգում 1-ին հաստատողի մոտ
    docExist = Online_PaySys_Check_Doc_In_Verifier(docNumber, data, data)
    If Not docExist Then
        Log.Error("The document with number " & docNumber & " doesn't exist in 1st verify documents")
        Exit Sub
    End If
    
    'Փաստաթղթի վավերացում 1-ին հաստաոտղի կողմից
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
    
    'Արտաքին փոխանցումների Հաշվառման ենթակա թղթապանկում հանձնարարգրի առկայության ստուգում
    docExist = Check_Doc_In_UnderRegistration_Folder(docNumber, data, data)
    If Not docExist Then
        Log.Error("After fadeing order from registered payment orders folder with number " & docNumber & " must exist in under registration folder " )
    Else
        'Փաստաթղթի ուղարկում մասնակի խմբագրման
        Call Paysys_Delete_Doc(False)
    End If
    
    'Test CleanUp
    Call Close_AsBank()
End Sub