'USEUNIT Library_Templates
'USEUNIT Template_Checker
'USEUNIT Library_Common

'Test Case N 159963

Private fCode

Sub WrongScenar1_WithoutImportFileOpenFolderAndExport
    
    Call TemplateStartUp()
    
'    fCode = "TemplateêË³Éêó»Ý³ñ1"
     fCode="TemplateWrong1"
    Call TemplateFilter (" ", " ", " ")
     fName = "TestTemplateWrong1"
'    fName = "TestTemplateêË³Éêó»Ý³ñ1"
    fEname = "TestTemplate"
    fType = "2"
    Connectivity = True
    Updateable = True
    Utilities.ShortDateFormat = "dd/mm/yy"
    Call CreateTemplate(fCode, fName, fEname, fType, Connectivity, Updateable)
    
    Call Check_UnableOpenFolder
    Call CloseTemplateGridWindow
    
    Call TestCleanUp()
End Sub

'-------------------------------------------------------------------------------------------------------

Private Sub TestCleanUp()
    Call TemplateFilter (" ", " ", " ")
    
    bAnswer = DeleteTemplate(Array(fCode))
    If bAnswer Then
        Log.Message("Template with code " & fCode & " is deleted!!!")
    Else
        Log.Error("Something wrong with deletion template with code " & fCode & " ... ")
    End If
    
    Call CloseTemplateGridWindow
    
    TemplateCleanUp()
End Sub