'USEUNIT Library_Common
'USEUNIT Constants

'-------------------------------------------------------------------------------------------------------
'Â»Ù÷É»ÛÃÇ üÇÉïñ
Sub TemplateFilter (fCode, fType, Connectivity)
    Log.Message("Template Filter opening started...")
    
    BuiltIn.Delay(2000)
    Call ChangeWorkspace(c_Admin)
    Call wTreeView.DblClickItem("|²¹ÙÇÝÇëïñ³ïáñÇ ²Þî|Ð³Ù³Ï³ñ·³ÛÇÝ ³ßË³ï³ÝùÝ»ñ|îå»Éáõ Ó¨³ÝÙáõßÝ»ñ|îå»Éáõ Ó¨³ÝÙáõßÝ»ñÇ Õ»Ï³í³ñáõÙ")
    BuiltIn.Delay(2000)
    Call Rekvizit_Fill("Dialog", 1, "General", "CODE",  fCode)
    Call Rekvizit_Fill("Dialog", 1, "General", "TYPE",  fType)
    Call Rekvizit_Fill("Dialog", 1, "General", "DOCCONNECTED",  Connectivity)
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    
    Log.Message("Template Filter opening ended!")
End Sub

'------------------------------------------------------------------------------------------------
Sub CheckFilterResult(TemplateCode)
    Set grid = wMDIClient.vbObject("frmPttel").vbObject("tdbgView")
    wMDIClient.vbObject("frmPttel").vbObject("tdbgView").MoveFirst
    
    For k = LBound(TemplateCode) To UBound(TemplateCode)
        Do Until wMDIClient.vbObject("frmPttel").vbObject("tdbgView").EOF
            
            If Trim(grid.Columns.Item(0).Text) = Trim(TemplateCode(k)) Then
                Log.Message("The " & TemplateCode(k) & " Code is present!! ")
                Exit Do
            Else
                wMDIClient.vbObject("frmPttel").vbObject("tdbgView").MoveNext
                
            End If
        Loop
        
        If wMDIClient.vbObject("frmPttel").vbObject("tdbgView").EOF Then
            Log.Error("The " & TemplateCode(k) & " Code IS NOT Present!! ")
        End If
        wMDIClient.vbObject("frmPttel").vbObject("tdbgView").MoveFirst
    Next
End Sub

'-----------------------------------------------------------------------
Sub CloseTemplateGridWindow
    BuiltIn.Delay(3000)
    wMDIClient.vbObject("frmPttel").Close()
End Sub

'-------------------------------------------------------------------------------------------------
'Â»Ù÷É»ÛÃÇ ëï»ÕÍáõÙ
Sub CreateTemplate(fCode, fName, fEname, fType, Connectivity, Updateable)
    Log.Message("Template creation with code " & fCode & "  is started....")
    
    BuiltIn.Delay(3000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_Add)
    BuiltIn.Delay(2000)
    Call Rekvizit_Fill("Dialog", 1, "General", "CODE", fCode)
    Call Rekvizit_Fill("Dialog", 1, "General", "NAME", fName)
    Call Rekvizit_Fill("Dialog", 1, "General", "ENAME", fEname)
    Call Rekvizit_Fill("Dialog", 1, "General", "TYPE", fType)
    If Connectivity = True Then
      Call Rekvizit_Fill("Dialog", 1, "CheckBox", "DOCCONNECTED", cbChecked)
    Else
      Call Rekvizit_Fill("Dialog", 1, "CheckBox", "DOCCONNECTED", cbUnChecked)
    End If
    If Updateable = True Then
      Call Rekvizit_Fill("Dialog", 1, "CheckBox", "UPDATEABLE", cbChecked)
    Else
      Call Rekvizit_Fill("Dialog", 1, "CheckBox", "UPDATEABLE", cbUnchecked)
    End If
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    
    Log.Message("Template creation with code " & fCode & " ended!")
End Sub

'-------------------------------------------------------------------------------------------------------
'Â»Ù÷É»ÛÃÇ ËÙµ³·ñáõÙ
Sub EditTemplate(fCode, fName, Ename, fType, Connectivity, Updateable)
    Log.Message("Editing of " & fCode & " is started... ")
    
    BuiltIn.Delay(3000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_ToEdit)
    BuiltIn.Delay(2000)
    
    fCodeInitial = Get_Rekvizit_Value("Dialog", 1, "General", "CODE")
    Call Rekvizit_Fill("Dialog", 1, "General", "CODE", "![End][Del]" & fCode)
    Call Rekvizit_Fill("Dialog", 1, "General", "NAME", "![End][Del]" & fName)
    Call Rekvizit_Fill("Dialog", 1, "General", "ENAME", "![End][Del]" & Ename)
    Call Rekvizit_Fill("Dialog", 1, "General", "TYPE", "![End][Del]" & fType)
     If Connectivity = True Then
      Call Rekvizit_Fill("Dialog", 1, "CheckBox", "DOCCONNECTED", 1)
    Else
      Call Rekvizit_Fill("Dialog", 1, "CheckBox", "DOCCONNECTED", 0)
    End If
    If Updateable = True Then
      Call Rekvizit_Fill("Dialog", 1, "CheckBox", "UPDATEABLE", 1)
    Else
      Call Rekvizit_Fill("Dialog", 1, "CheckBox", "UPDATEABLE", 0)
    End If
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    
    Log.Message("Editing of " & fCodeInitial & "-> " & fCode &" is ended... ")
End Sub

'----------------------------------------------------------------------------------------------------
Function DeleteTemplate(TemplateCode)
    Log.Message("Starting Deleting template...")
    
    bResult = False
    
    Set grid = wMDIClient.vbObject("frmPttel").vbObject("tdbgView")
    For k = LBound(TemplateCode) To UBound(TemplateCode)
        wMDIClient.vbObject("frmPttel").vbObject("tdbgView").MoveFirst
        Do Until wMDIClient.vbObject("frmPttel").vbObject("tdbgView").EOF
            If Trim(grid.Columns.Item(0).Text) = Trim(TemplateCode(k)) Then
                BuiltIn.Delay(3000)
                Call wMainForm.MainMenu.Click(c_AllActions)
                Call wMainForm.PopupMenu.Click(c_Delete)
                BuiltIn.Delay(1000)
                Call ClickCmdButton(5, "²Ûá")
                bResult = True
                Exit Do
            End If
            wMDIClient.vbObject("frmPttel").vbObject("tdbgView").MoveNext
        Loop
    Next
    
    DeleteTemplate = bResult
    
    Log.Message("Ending Deleting template...")
End Function

'-----------------------------------------------------------------------------------------------
'ü³ÛÉÇ Ý»ñÙáõÍáõÙ
Sub ImportFile(ImportType, Path)    
    Log.Message("Importing file is started....")
    
    BuiltIn.Delay(3000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click("Ներմուծել ֆայլը")
    
    Select Case ImportType
        Case "ImportWithClick"
            p1.vbObject("frmAsUstPar").vbObject("TabFrame").vbObject("AsTypePath").vbObject("CmdViewPath").ClickButton
            p1.Window("#32770", "Select File", 1).Window("ComboBoxEx32", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1).DblClick
            p1.Window("#32770", "Select File", 1).Window("ComboBoxEx32", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1).Keys(Path & "[Tab]")
            p1.Window("#32770", "Select File", 1).Window("Button", "&Open", 1).ClickButton
        Case "ImportWithNoClick"
            p1.vbObject("frmAsUstPar").vbObject("TabFrame").vbObject("AsTypePath").vbObject("TxtPath").Text = ""
            p1.vbObject("frmAsUstPar").vbObject("TabFrame").vbObject("AsTypePath").vbObject("TxtPath").Keys(Path & "[Tab]")
    End Select
    
    p1.vbObject("frmAsUstPar").vbObject("CmdOK").ClickButton
    Log.Message("Importing file ended!")
    
    BuiltIn.Delay(delay_middle)
End Sub

'--------------------------------------------------------------------------------------------------------
Sub Check_UnableOpenFolder
    Log.Message("Starting checking Opening  folder ...")
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click("Բացել ֆայլի թղթապանակը")
    
    Set my_vbobj = p1.WaitVBObject("frmAsMsgBox", 1000)
    If my_vbobj.Exists Then
        p1.vbObject("frmAsMsgBox").vbObject("cmdButton").ClickButton
        Log.Message("Rigth! You cant open folder before importing file!!!")
    Else
        Log.Error("Wrong! You cant open folder before importing file!!!")
    End If
    
    Log.Message("Ending checking Opening  folder ...")
End Sub

'------------------------------------------------------------------------------------------------------
'¸Çï»É ÷³ëÃ³ÃÕÃ»ñÇ óáõó³ÏÁ ¨ ³í»É³óÝ»É
Sub SeeDocList_AddDoc(DocType, ActiveBanadzev)
    Log.Message("Seeing the document list and adding document started.... ")
    
    BuiltIn.Delay(3000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click("Դիտել փաստաթղթերի ցուցակը")
    BuiltIn.Delay(2000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_Add)
    BuiltIn.Delay(2000)
    
    Call Rekvizit_Fill("Dialog", 1, "General", "DocType", DocType)
    Call Rekvizit_Fill("Dialog", 1, "General", "Access", ActiveBanadzev)
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    
    BuiltIn.Delay(1000)
    wMDIClient.vbObject("frmPttel_2").Close()
    
    Log.Message("Seeing the document list and adding document ended! ")
End Sub

'--------------------------------------------------------------------------------------------------------
'¸Çï»É ÷³ëÃ³ÃÕÃ»ñÇ óáõó³ÏÁ ¨  ËÙµ³·ñ»É
Sub SeeDocList_EditDoc(DocType, ActiveBanadzev)
    Log.Message("Seeing the document list and editing document started.... ")
    
    BuiltIn.Delay(3000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click("Դիտել փաստաթղթերի ցուցակը")
    BuiltIn.Delay(3000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_ToEdit)
    BuiltIn.Delay(2000)
    
    Call Rekvizit_Fill("Dialog", 1, "General", "DocType", "![End][Del]" & DocType)
    Call Rekvizit_Fill("Dialog", 1, "General", "Access", "![End][Del]" & ActiveBanadzev)
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    
    BuiltIn.Delay(2000)
    wMDIClient.vbObject("frmPttel_2").Close()
    
    Log.Message("Seeing the document list and editing document ended! ")
End Sub

'--------------------------------------------------------------------------------------------------------
'ö³ëï³ÃÕÃÇ ïå»Éáõ Ó¨³ÝÙáõßÝ»ñ
Sub DocTemplateFilter (DocType)
    BuiltIn.Delay(2000)
    Call wTreeView.DblClickItem("|²¹ÙÇÝÇëïñ³ïáñÇ ²Þî|Ð³Ù³Ï³ñ·³ÛÇÝ ³ßË³ï³ÝùÝ»ñ|îå»Éáõ Ó¨³ÝÙáõßÝ»ñ|ö³ëï³ÃÕÃÇ ïå»Éáõ Ó¨³ÝÙáõßÝ»ñ")
    BuiltIn.Delay(1000)
    Call Rekvizit_Fill("Dialog", 1, "General", "DocType", DocType)
    Call ClickCmdButton(2, "Î³ï³ñ»É")
End Sub

'--------------------------------------------------------------------------------------------------------
'²í»É³óÝ»É Ó¨³ÝÙáõß ÷³ëï³ÃÕÃÇ íñ³
Sub AddTemplateToDoc(TeplateCode, ActiveBanadzev)
    Log.Message("Adding Template to Document  started....")
    
    BuiltIn.Delay(3000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_Add)
    BuiltIn.Delay(2000)
    
    Call Rekvizit_Fill("Dialog", 1, "General", "CODE", TeplateCode)
    Call Rekvizit_Fill("Dialog", 1, "General", "ACCESS", ActiveBanadzev)
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    
    Log.Message("Adding Template to Document ended!....")
End Sub

'--------------------------------------------------------------------------------------------------------
Sub Check_UnableAddTamplateToDoc(TemplateCode)
    Log.Message("Starting  Checking unable add template to Doc......")
    
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_Add)
    
    Str = GetVBObject_Dialog ("CODE", p1.vbObject("frmAsUstPar"))
    p1.vbObject("frmAsUstPar").vbObject("TabFrame").vbObject(Str).Keys("^[Down]")
    
    Set grid = p1.vbObject("frmModalBrowser_2").vbObject("tdbgView")
     'Set grid = Aliases.Asbank.VBObject("frmModalBrowser").VBObject("tdbgView")
    grid.MoveFirst
    Do Until grid.EOF
        
        If Trim(grid.Columns.Item(1).Text) = Trim(TemplateCode) Then
            Log.Error ("WRONG!!! Template Code was found and can add template to Doc")
            Exit Do
        End If
        grid.MoveNext
    Loop
    
    If grid.EOF Then
        Log.Message ("All is OK!!!  Template Code was not found and can NOT add template to Doc")
    End If
    
    Sys.Process("Asbank").VBObject("frmModalBrowser_2").Close
    Sys.Process("Asbank").vbObject("frmAsUstPar").vbObject("CmdCancel").ClickButton
    
    Log.Message("Ending checking unable add template to Doc......")
End Sub

'--------------------------------------------------------------------------------------------------------
Sub Check_UnableAddDoc
    Log.Message("Satrting checking add doc to Template.......")
    
    Call wMainForm.MainMenu.Click(c_AllActions)
    
    menuCount = wMainForm.PopupMenu.Count
    k = wMainForm.PopupMenu.Items(0).Caption
    For j = 1 To menuCount
        If k = "Դիտել փաստաթղթերի ցուցակը" Then
            Log.Error("WRONG!!!Template can not connect with doc....")
            Exit For
        ElseIf j = menuCount Then
            Log.Message("RIGTH!!!Template can not connect with doc....")
            Exit For
        End If
        k = wMainForm.PopupMenu.Items(j).Caption
    Next
    
    Log.Message("Ending  checking add doc to Template....")
End Sub

'--------------------------------------------------------------------------------------------------------
Sub Check_UnableChangeConnectivity(fCode)
    Log.Message("Starting checking editing template connectivity...")
    
    BuiltIn.Delay(3000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_ToEdit)
    BuiltIn.Delay(1000)
    Str = GetVBObject_Dialog ("DOCCONNECTED", p1.vbObject("frmAsUstPar"))
    If p1.vbObject("frmAsUstPar").vbObject("TabFrame").vbObject(Str).Enabled = False Then
        Log.Message("RIGTH! Doc had been added to Template, And Can not edit template connectivity!")
    ElseIf p1.vbObject("frmAsUstPar").vbObject("TabFrame").vbObject(Str).Enabled = True Then
        Log.Error("WRONG! Doc had been added to Template, but it is possible to edit template connectivity !")
    End If
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    
    Log.Message("Ending checking editing template connectivity...")
End Sub

'--------------------------------------------------------------------------------------------------------
'ÊÙµ³·ñ»É ³í»É³óñ³Í Ó¨³ÝÙáõßÁ Ó³ëï³ÃÕÃÇ íñ³
Sub EditTemplateForDoc(TemplateCode, ActiveBanadzev)
    Log.Message("Editing Template For Document started....")
    
    Set grid = wMDIClient.vbObject("frmPttel").vbObject("tdbgView")
    grid.MoveFirst()
    Do While Not grid.EOF
        If Trim(grid.Columns.Item(2).Text) = TemplateCode Then
            Exit Do
        End If
        grid.MoveNext()
    Loop
    If grid.EOF Then
        Log.Error("There  is NO Templaate With " & TemplateCode & " Code!!!")
    Else
        BuiltIn.Delay(3000)
        Call wMainForm.MainMenu.Click(c_AllActions)
        Call wMainForm.PopupMenu.Click(c_ToEdit)
        BuiltIn.Delay(2000)
        Call Rekvizit_Fill("Dialog", 1, "General", "Access", "![End][Del]" & ActiveBanadzev)
        Call ClickCmdButton(2, "Î³ï³ñ»É")
    End If
    
    Log.Message("Editing Template for Document ended!")
End Sub

'--------------------------------------------------------------------------------------------------------
'æÝç»É ³í»É³óñ³Í Ó¨³ÝÙáõßÁ Ó³ëï³ÃÕÃÇ íñ³ÛÇó
Sub DeleteTemplateForDoc(TemplateCode)
    
    Log.Message("Deleting Template for document started...")
    
    Set grid = wMDIClient.vbObject("frmPttel").vbObject("tdbgView")
    grid.MoveFirst()
    Do While Not grid.EOF
        If Trim(grid.columns.Item(2).Text) = TemplateCode Then
            Exit Do
        End If
        grid.MoveNext()
    Loop
    
    If grid.EOF Then
        Log.Error("There  is NO Templaate With " & TemplateCode & " Code!!!")
    Else
        Call wMainForm.MainMenu.Click(c_AllActions)
        Call wMainForm.PopupMenu.Click(c_Delete)
        
        p1.vbObject("frmAsMsgBox").vbObject("cmdButton").ClickButton
        Log.Message("The " & TemplateCode & " Template was DELETED!! ")
    End If
    
    Log.Message("Deleting Template for document ended!")
End Sub

'-------------------------------------------------------------------------------------------------------
Public Sub TemplateStartUp()
    Utilities.ShortDateFormat = "yyyymmdd"
    endDATE = Utilities.DateToStr(Utilities.Date())
    startDATE = Utilities.DateToStr(Utilities.IncMonth(Utilities.Now, -24))
    
    Call Initialize_AsBank("bank", startDATE, endDATE)
    Call login("Armsoft")
    call ChangeWorkspace(c_Admin)
End Sub

'-------------------------------------------------------------------------------------------------------
Public Sub TemplateCleanUp()
    Call Close_AsBank
    BuiltIn.Delay(3000)
End Sub