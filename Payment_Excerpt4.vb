'USEUNIT Library_Common
'USEUNIT Payment_Except_Library

'Test case ID 165664

Sub LoanAgr_With_Schedule_Statment_Test3()
    
    Dim fDATE, startDATE , cpath, docType, docNum , sDate, eDate, savePath, fName
    Dim docExist, fIdent , fileName1 , fileName2,template
    
    fDATE = "20250101"
    startDATE = "20030101"
    cpath = "|ì³ñÏ»ñ (ï»Õ³µ³ßËí³Í)|ä³ÛÙ³Ý³·ñ»ñ"
    docType = "1"
    docNum = "ST-004"
    sDate = "17/03/11"
    eDate = "17/04/11"
    savePath = Project.Path & "Stores\Statement_Check_140414\Statement_Actual\"
    fName = "4.txt"
    fileName1 = Project.Path & "Stores\Statement_Check_140414\Statement_Actual\4.txt"
    fileName2 = Project.Path & "Stores\Statement_Check_140414\Statement_Expected\ST-004_Expected.txt"
    template = ""
    
    'Test StartUp 
    Call Initialize_AsBank("bank", startDATE, fDATE)
    Call Login("CREDITOPERATOR")

    docExist = Contracts_Filter_Fill(docType, docNum, cpath)
    If Not docExist Then
        Log.Error("Document with number" & docNum & "doesn't exist")
        Exit Sub
    End If
    
    docExist = View_Statment (sDate, eDate, True,template)
    If Not docExist Then
        Log.Error("Document statement doesn't exist")
        Exit Sub
    End If
    
    Call SaveDoc(savePath, fName)
    fIdent = Compare_Files(fileName1, fileName2, "")
    If Not fIdent Then
        Log.Error(fileName1 & "and" & fileName2 &" :Files are not identical" )
    End If
    
    'Test CleanUp 
    Call Close_AsBank()
End Sub