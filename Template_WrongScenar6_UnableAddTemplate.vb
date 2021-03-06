'USEUNIT Library_Templates
'USEUNIT Template_Checker
'USEUNIT Library_Common

'Test Case N 160007

Private fCode

Sub WrongScenar6_UnableAddTemplate
    
    Call TemplateStartUp()
    
    Call TemplateFilter (" ", " ", " ")
    fCode = "WrongTemplate6"
    fName = "WrongTestTemplate6"
    fEname = "WrongTest3Template6"
    fType = "0"
    Connectivity = False
    Updateable = True
    Utilities.ShortDateFormat = "dd/mm/yy"
    Call CreateTemplate(fCode, fName, fEname, fType, Connectivity, Updateable)
    Call CloseTemplateGridWindow
    Call CheckTemplate(fCode, fName, fEname, fType, Connectivity, Updateable, "", Utilities.DateToStr(Utilities.Date()))
    Call DocTemplateFilter ("PkCash")
    Call Check_UnableAddTamplateToDoc(fCode)
    Call CloseTemplateGridWindow
    
    Call TestCleanUp()
End Sub

'-------------------------------------------------------------------------------------------------------

Private Sub TestCleanUp()
    Call TemplateFilter (" ", " ", " ")
    
    bAnswer = DeleteTemplate(Array(fCode))
    If bAnswer Then
        Log.Message("Deletion of Template ended successfully!!")
    Else
        Log.Error("In deletion something wrong!")
    End If
    Call CheckDeleteTemplate(fCode, "0")
    
    Call CloseTemplateGridWindow
    
    TemplateCleanUp()
End Sub