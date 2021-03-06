'USEUNIT SWIFT_International_Payorder_Library
'USEUNIT Library_Common
'USEUNIT Acc_Statements_Library
'USEUNIT Payment_Except_Library   
'USEUNIT Constants

'Test case Id 165862

Sub Payment_Except_Cards_Test()

    Dim fDATE, startDATE ,cardNumber , sDate, eDate
    Dim docExist,fileName1 , fileName2,template,toFile,param
    
    fDATE = "20260101"
    startDATE = "20070101"
    cardNumber = "9051190200005849"
    sDate = "01/01/07"
    eDate = "01/01/08"
    template = "CardStateCB_AS\7"
    
    'Î
    isExists = aqFile.Exists(Project.Path& "Stores\Payment_Excerpt_htm_Templates\cards.txt")
    If isExists Then
      aqFileSystem.DeleteFile(Project.Path& "Stores\Payment_Excerpt_htm_Templates\cards.txt")
    End If
    
    'Test StartUp start
    Call Initialize_AsBank("bank", startDATE, fDATE)
    'Test StartUp end  
    
    Call ChangeWorkspace(c_CardsSV)
    'Պլաստիկ քարտի առկայության ստուգում "Պլաստիկ քարտեր " թղթապանակում 
    docExist = Check_CardExist_In_Carsds_Folder(cardNumber)
   
    If Not docExist Then
        Log.Error("Card with number " & cardNumber & " isn't exist in cards folder")
        Exit Sub
    End If  
    
    isExist = View_Card_Statment(sDate, eDate,template)
    If Not isExist Then
        Log.Error("Document statement doesn't exist")
        Exit Sub
    End If
    BuiltIn.Delay(10000)
    fileName = ListFiles("C:\Users\"& Sys.UserName & "\AppData\Local\Temp\AS-BANK")
    fileName1 = "C:\Users\" & Sys.UserName & "\AppData\Local\Temp\AS-BANK\" & Trim(fileName)
    fileName2 = Project.Path & "Stores\Payment_Excerpt_htm_Templates\cardsReal.txt"
    Log.Message(fileName1)
    toFile = Project.Path & "Stores\Payment_Excerpt_htm_Templates\cards.txt"
    Call Read_Write_File(fileName1, toFile)
    
    param = "<[/]span><span>(0[1-9]|[1-2][0-9]|3[0-1]).(0[1-9]|1[0-2]).[0-9]{4}<[/]span>, <span>(2[0-3]|[01][0-9]):[0-5][0-9]<[/]span>"
    Call Compare_Files(fileName2, toFile,param)
    
    Call Close_AsBank()
    
End Sub