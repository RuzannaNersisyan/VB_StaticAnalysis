'USEUNIT Library_Templates
'USEUNIT Template_Checker
'USEUNIT Library_Common

'Test Case N 159965

Private fCode

Sub WrongScenar2_WithoutImportPrint
    
    Call TemplateStartUp()
    
    fCode = "WrongTemplate2"
    
    Call TemplateFilter (" ", " ", " ")
    
    fName = "WrongTestTemplate2"
    fEname = ""
    fType = "0"
    Connectivity = True
    Updateable = True
    Utilities.ShortDateFormat = "dd/mm/yy"
    Call CreateTemplate(fCode, fName, fEname, fType, Connectivity, Updateable)
    Call CheckTemplate (fCode, fName, fEname, fType, Connectivity, Updateable, "", Utilities.DateToStr(Utilities.Date()))
    Call CloseTemplateGridWindow
    
    GroupCashInputISN = "76963468"
    ActiveBanadzev = "Doc.Grid(""SubSums"").Value(0, ""SUMMA"") + Doc.Grid(""SubSums"").Value(1, ""SUMMA"") "
    
    Call DocTemplateFilter ("PkCash")
    Call AddTemplateToDoc(fCode, ActiveBanadzev)
    Call CloseTemplateGridWindow
    Call PrintDocument(GroupCashInputISN, "word", fName, templatePath & "\EmptyDoc.doc", Array(), True)
    
    call TestCleanUp()    
End Sub

'-------------------------------------------------------------------------------------------------------
Private Sub  TestCleanUp()
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