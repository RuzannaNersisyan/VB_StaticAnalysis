'USEUNIT Library_Templates
'USEUNIT Template_Checker
'USEUNIT Template_NecessaryDocuments
'USEUNIT Library_Common

'Test Case N 159882

Private fCode

Sub DoubleImport
    
    Call TemplateStartUp()
    
    fCode = "AAA"
    
    ' Ý³Ëáñáù ÷áñÓ»É Ñ»é³óÝ»É
    Call TemplateFilter (" ", " ", " ")
    DeleteTemplate(Array(fCode))
    Call CloseTemplateGridWindow
        
    'êï»ÕÍ»É Ã»Ù÷É»ÛÃ
    Call TemplateFilter (" ", " ", " ")
    
    fName = "TestTemplateAAA"
    fEname = "TemplateAAA"
    fType = "1"
    Connectivity = True
    Updateable = False
    Utilities.ShortDateFormat = "dd/mm/yy"
    Call CreateTemplate(fCode, fName, fEname, fType, Connectivity, Updateable)
    Call CheckTemplate (fCode, fName, fEname, fType, Connectivity, Updateable, "", Utilities.DateToStr(Utilities.Date()))
    
    'Ü»ñÙáõÍ»É ý³ÛÉ
    ImportType = "ImportWithNoClick"
    Call ImportFile(ImportType, templatePath & "\Loan_Distributed_scheduled_1.xlsx")
    Call CloseTemplateGridWindow
    Call CheckImportFile(fCode, fType, templatePath & "\Loan_Distributed_scheduled_1.xlsx")
    
    
    DocType = "PkCash"
    ActiveBanadzev = ""
    Call DocTemplateFilter (DocType)
    Call AddTemplateToDoc(fCode, ActiveBanadzev)
    Call CloseTemplateGridWindow
    Call CheckTemplateMapping(fCode, fType, DocType, ActiveBanadzev)
    
    'îå»É ¹áÏáõÙ»ÝïÁ
    GroupCashInputISN = "76963468"
    Call PrintDocument(GroupCashInputISN, "excel", fName, templatePath & "\Loan_Distributed_scheduled_1.xlsx", Array(), False)
    
    ' Ü»ñÙáõÍ»É »ñÏñáñ¹ ³Ý·³Ù
    Call TemplateFilter (fCode, " ", " ")
    ImportType = "ImportWithNoClick"
    fType2 = "0"
    Call EditTemplate(fCode, fName, fEname, fType2, Connectivity, Updateable)
    Call ImportFile(ImportType, templatePath & "\Group_Cash_In_1.doc")
    Call CloseTemplateGridWindow
    Call CheckImportFile(fCode, fType2, templatePath & "\Group_Cash_In_1.doc")
    
    Call PrintDocument(GroupCashInputISN, "word", fName, templatePath & "\Group_Cash_In_1.doc", Array(), False)
    
    Call DocTemplateFilter (DocType)
    Call DeleteTemplateForDoc(fCode)
    Call CloseTemplateGridWindow
    
    Call TestCleanUp()
End Sub

'-------------------------------------------------------------------------------------------------------

Private Sub TestCleanUp()
    Call TemplateFilter (" ", " ", " ")
    
    bAnswer = DeleteTemplate(Array(fCode))
    If bAnswer Then
        Log.Message("The Template was deleted successfully!!")
    Else
        Log.Error("Wrong something in deletion template!!")
    End If
    
    Call CheckDeleteTemplate(fCode, "0")
    
    Call CloseTemplateGridWindow
    TemplateCleanUp()
End Sub