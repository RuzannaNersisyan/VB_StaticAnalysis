'USEUNIT SWIFT_International_Payorder_Library
'USEUNIT Library_Common
'USEUNIT Payment_Except_Library

'Test case Id 165863

Sub Payment_Except_Credit_Test()

  Dim startDATE,fDATE,docType, docNum, cpath,isExist
  Dim sDate, eDate,template,fileName,toFile,isExists,param
   
   startDATE  = "20010101"
   fDATE = "20260101"
   docType = 1
   docNum = "V-002536"
   cpath = "|ì³ñÏ»ñ (ï»Õ³µ³ßËí³Í)|ä³ÛÙ³Ý³·ñ»ñ"
   sDate = "010314"
   eDate = "010414"   
   template = "AgrStateCredit_AS\7"
    
    'Կատարում է ստուգում,եթե նման անունով ֆայլ կա տրված թղթապանակում ,ջնջում է
    isExists = aqFile.Exists(Project.Path& "Stores\Payment_Excerpt_htm_Templates\credit.txt")
    If isExists Then
      aqFileSystem.DeleteFile(Project.Path& "Stores\Payment_Excerpt_htm_Templates\credit.txt")
    End If
    
  'Test StartUp start
    Call Initialize_AsBank("bank", startDATE, fDATE)
    Call Login("CREDITOPERATOR")
    isExist = Contracts_Filter_Fill(docType, docNum, cpath)
    If Not isExist Then 
        Log.Error("Document with number " & docNum & "doesn't exist")
        Exit Sub
    End If    
    isExist = View_Statment (sDate, eDate,False,template)
    If Not isExist Then
        Log.Error("Document statement doesn't exist")
        Exit Sub
    End If
    BuiltIn.Delay(15000)
    fileName = ListFiles("C:\Users\"& Sys.UserName & "\AppData\Local\Temp\AS-BANK")
    fileName1 = "C:\Users\"& Sys.UserName & "\AppData\Local\Temp\AS-BANK\" & Trim(fileName)
    fileName2 = Project.Path & "Stores\Payment_Excerpt_htm_Templates\creditReal.txt"
    Log.Message(fileName1)
    toFile = Project.Path & "Stores\Payment_Excerpt_htm_Templates\credit.txt"
    Call Read_Write_File(fileName1, toFile)
    
    param = "([0-3][0-9].[0-1][0-9].[2][0-3][0-9][0-9]..[0-2][0-9]:[0-5][0-9])"
    Call Compare_Files(fileName2, toFile,param)
    
    Call Close_AsBank()
    
End Sub