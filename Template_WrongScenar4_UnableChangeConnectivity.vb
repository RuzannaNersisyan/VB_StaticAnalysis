'USEUNIT Library_Templates
'USEUNIT Template_Checker
'USEUNIT Library_Common

'Test Case N 159997

Private fCode

Sub WrongScenar4_UnableChangeConnectivity
    
    Call TemplateStartUp()
    
    fCode = "WrongTemplate4"
    
    Call TemplateFilter (" ", " ", " ")
    
    fName = "WrongTestTemplate4"
    fEname = "WRTemplate4"
    fType = "1"
    Connectivity = True
    Updateable = False
    Utilities.ShortDateFormat = "dd/mm/yy"
    Call CreateTemplate(fCode, fName, fEname, fType, Connectivity, Updateable)
    Call CheckTemplate(fCode, fName, fEname, fType, Connectivity, Updateable, "", Utilities.DateToStr(Utilities.Date()))
    
    DocType = "PkCash"
    ActiveBanadzev = ""
    Call SeeDocList_AddDoc(DocType, ActiveBanadzev )
    Call Check_UnableChangeConnectivity(fCode)
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