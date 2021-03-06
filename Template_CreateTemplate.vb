'USEUNIT Library_Templates
'USEUNIT Template_Checker
'USEUNIT Library_Common
'USEUNIT Template_Checker

'Test Case N 159833

Sub CreateTemplatesTest
    BuiltIn.Delay(20000)
    
    Call TemplateStartUp()
    
    Call TemplateFilter (" ", " ", " ")
    fCode = "X"
    fName = "TestTemplate1"
    fEname = ""
    fType = "0"
    Connectivity = True
    Updateable = True
    Utilities.ShortDateFormat = "dd/mm/yy"
    Call CreateTemplate(fCode, fName, fEname, fType, Connectivity, Updateable)
    Call CheckTemplate(fCode, fName, fEname, fType, Connectivity, Updateable, "", Utilities.DateToStr(Utilities.Date()))
    
    fCode = "T2"
    fName = "TestTemplate2"
    fEname = "Template2"
    fType = "1"
    Connectivity = True
    Updateable = False
    Utilities.ShortDateFormat = "dd/mm/yy"
    Call CreateTemplate(fCode, fName, fEname, fType, Connectivity, Updateable)
    Call CheckTemplate(fCode, fName, fEname, fType, Connectivity, Updateable, "", Utilities.DateToStr(Utilities.Date()))
    
    fCode = "Template3"
    fName = "TestTemplate3"
    fEname = "Test3Template3Template3Template3Template3Template3"
    fType = "0"
    Connectivity = False
    Updateable = True
    Utilities.ShortDateFormat = "dd/mm/yy"
    Call CreateTemplate(fCode, fName, fEname, fType, Connectivity, Updateable)
    Call CheckTemplate(fCode, fName, fEname, fType, Connectivity, Updateable, "", Utilities.DateToStr(Utilities.Date()))
    
    fCode = "Template4Template4TM"
    fName = "TestTemplate4"
    fEname = "TestTemplate4"
    fType = "1"
    Connectivity = False
    Updateable = False
    Utilities.ShortDateFormat = "dd/mm/yy"
    Call CreateTemplate(fCode, fName, fEname, fType, Connectivity, Updateable)
    Call CheckTemplate(fCode, fName, fEname, fType, Connectivity, Updateable, "", Utilities.DateToStr(Utilities.Date()))
    
    fCode = "Template5"
    fName = "TestTemplate5"
    fEname = "TestTemplate5"
    fType = "2"
    Connectivity = True
    Updateable = True
    Utilities.ShortDateFormat = "dd/mm/yy"
    Call CreateTemplate(fCode, fName, fEname, fType, Connectivity, Updateable)
    Call CheckTemplate(fCode, fName, fEname, fType, Connectivity, Updateable, "", Utilities.DateToStr(Utilities.Date()))
    Call CloseTemplateGridWindow
    
    Call TemplateCleanUp()
    
End Sub