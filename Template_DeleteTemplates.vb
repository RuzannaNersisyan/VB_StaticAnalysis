'USEUNIT Library_Templates
'USEUNIT Template_Checker
'USEUNIT Library_Common

'Test Case N 159918

Sub DeleteTemplatesTest
    
    Call TemplateStartUp()
    
    Template1 = "X"
    Template2 = "T2"
    Template3 = "Template3"
    Template4 = "Template4Template4TM"
    Template5 = "Template5"
    
    Template1Edit = "XEdit"
    Template4Edit = "Template4Template"
    
    Dim TemplateCode(4)
    TemplateCode(0) = Template1Edit
    TemplateCode(1) = Template2
    TemplateCode(2) = Template3
    TemplateCode(3) = Template4Edit
    TemplateCode(4) = Template5
    Dim bAnswer
    
    Call TemplateFilter (" ", " ", " ")
    bAnswer = DeleteTemplate(TemplateCode)
    If bAnswer = True Then
        Log.Message("Templates deleted successfully!")
    Else
        Log.Error("Something wrong with deleting all templates!")
    End If
    
    Call CheckDeleteTemplate(Template1Edit, "1")
    Call CheckDeleteTemplate(Template2, "1")
    Call CheckDeleteTemplate(Template3, "0")
    Call CheckDeleteTemplate(Template4Edit, "0")
    Call CheckDeleteTemplate(Template5, "2")
    
    Call CloseTemplateGridWindow
    
    Call TemplateCleanUp()
    
End Sub