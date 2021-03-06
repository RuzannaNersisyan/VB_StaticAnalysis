'USEUNIT Constants
'USEUNIT Library_Colour
'USEUNIT Library_Contracts
'USEUNIT Mortgage_Library
'USEUNIT Payment_Except_Library
'USEUNIT SWIFT_International_Payorder_Library

Dim files_count, links_count, del_count, fil_Count
Dim fCount, lCount, dCount

Public Const col_item = 1

' Delay - Ý»ñÇ Ñ³Ù³ñ ·Éáµ³É ÏáÝëï³ïÝ»ñÇ Ñ³Ûï³ñ³ñáõÙ ¨ ³ñÅ»ù³íáñáõÙ
Public Const delay_small = 500
Public Const delay_middle = 1000
Public Const delay_best = 3000
Public Const delay_big = 5000
Public Const delay_agr_send_huge = 50000
Public Const delay_LoanRegister = 500000

Dim DocNum
Dim DebitAccount
Dim CreditAccount
Dim TransitAccount
Dim CorrespondingAccount
Dim BranchCorrespondingAccount
Dim IncomeAccount
Dim Count_Messages
Dim PaymentSystemType

Dim p1
Dim AsBank
Dim wMainForm
Dim wMDIClient
Dim wTreeView
Dim w4
Dim w5
Dim w6

Public cConnectionString
Public OLAPConnectionString

Const stepTimeOut = 500
Const printTimeOut = 10000
Const WaitWordTimeOut = 2000
Const FillDocumentDelay = 120000
Const WaitDocStartFilling = 5000

Dim g_currentDate
Dim g_currentDate_For_Check
Dim g_currentDate_SQL
Dim g_currentDate_Line_SQL

docspath = ProjectSuite.Path & "AsBank\Stores\Files"
templatePath = docspath + "\Templates"

Function GetVBObject (Rekv, obj)
  If obj.DocFormCommon.Doc.Control(Rekv).index = 0 Then
     GetVBObject = obj.DocFormCommon.Doc.Control(Rekv).Name
  Else
     GetVBObject = obj.DocFormCommon.Doc.Control(Rekv).Name & "_" & CStr(obj.DocFormCommon.Doc.Control(Rekv).index + 1)
  End If
End Function

Function GetVBObject_Dialog (Rekv, obj)
  If obj.indcontrols(Rekv).index = 0 Then
     GetVBObject_Dialog = obj.indcontrols(Rekv).Name
  Else
     GetVBObject_Dialog = obj.indcontrols(Rekv).Name & "_" & CStr(obj.indcontrols(Rekv).index + 1)
  End If  
End Function

Sub GetConfigInformation(TestArea, ServerName, DatabaseName)  
    ' Create COM object
    Set Doc = Sys.OleObject("Msxml2.DOMDocument.6.0")
    Doc.async = False
    
    TestConfigPath = ProjectSuite.Path & "Project_AsBank\TestConfig.xml"
    
    Call Doc.load(TestConfigPath)
    Set Nodes = Doc.selectNodes("//DatabaseConfiguration")
    For i = 0 To Nodes.Length - 1
        Set Node = Nodes.Item(i)
        If Nodes(i).getElementsByTagName("TestArea").Item(0).text = testArea Then
            ServerName = Nodes(i).getElementsByTagName("ServerName").Item(0).text
            DatabaseName = Nodes(i).getElementsByTagName("DatabaseName").Item(0).text
        End If
    Next
End Sub

'----------------------------------------------------------------
Sub Login (Name)
    BuiltIn.Delay(2000)
    Call wMainForm.MainMenu.Click("Աշխատանք|Աշխատանքի սկիզբ")
    BuiltIn.Delay(2000)
    If p1.vbObject("frmLogin").Exists Then
      Call p1.vbObject("frmLogin").vbObject("txtPassword").Keys("![Tab]")
      Call p1.vbObject("frmLogin").vbObject("txtUserNameCombo").Keys(Name) 
      Call ClickCmdButton(0, "Î³ï³ñ»É")
    Else
      Log.Error "Can't find login window", "", pmNormal, ErrorColor
    End If 
    
    If p1.WaitVBObject("frmAsMsgBox", 3000).Exists Then
      Call p1.vbObject("frmAsMsgBox").vbObject("cmdButton_2").ClickButton
    End If
    BuiltIn.Delay(1000)
End Sub

'---------------------------------------------------------------------------------------------
'workspace
Public Sub ChangeWorkspace(workspace)
  Dim frmMainForm,frmMDIClient , frmExplorer, tvTreeView
  Sys.Process("Asbank").Refresh
  Set frmMainForm = Sys.Process("Asbank").VBObject("MainForm") 
  Set frmMDIClient = frmMainForm.Window("MDIClient", "", 1)
  Set frmExplorer = frmMDIClient.VBObject("frmExplorer")
  set tvTreeView= Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient").VBObject("frmExplorer").VBObject("tvTreeView")
  frmMDIClient.Refresh
  
    Call tvTreeView.ClickR()
  Call tvTreeView.PopupMenu.Click(workspace)
End Sub

'--------------------------------------------------------------
Sub Initialize_AsBankQA(DateStart, DateEnd)    
    TestedApps.Items("Asbank_QA").Terminate
    BuiltIn.Delay(3000) 
    
'    Call GetConfigInformation(TestArea, ServerName, DatabaseName)
    
  'TestedApps.Asbank.Parameters = """qasql\bank"" ""bankTesting_Sat"" ""armsoft"""

    Set p1 = TestedApps.Items("Asbank_QA").Run
    If (Not p1.Exists) Then
        Call Log.Error("ASBank application not found")
        Exit Sub
    End If
        
    k1 = Sys.Process("Asbank").vbObject("MainForm").WaitProperty("VisibleOnScreen", True, 20000)
    Set wMainForm = Sys.Process("Asbank").vbObject("MainForm")
    
    k2 = wMainForm.Window("MDIClient").WaitProperty("VisibleOnScreen", True, 20000)
    Set wMDIClient = wMainForm.Window("MDIClient")
    
    Set wMainForm = Sys.Process("Asbank").vbObject("MainForm")
    Set wMDIClient = wMainForm.Window("MDIClient")
    
    k3 = wMDIClient.vbObject("frmExplorer").WaitProperty("VisibleOnScreen", True, 20000)
    Set wfrmExplorer = wMDIClient.vbObject("frmExplorer")
    
    k4 = wfrmExplorer.vbObject("tvTreeView").WaitProperty("VisibleOnScreen", True, 20000)
    Set wTreeView = wfrmExplorer.vbObject("tvTreeView")
    
    Dim DatabaseName, ServerName
    DatabaseName = "bankTesting_QA"
    ServerName = "QASQLBANK"
    BuiltIn.delay(1000)
    cConnectionString = "Provider=SQLOLEDB.1;Password=sasa111;Persist Security Info=True;User ID=sa;Initial Catalog="& DatabaseName &";Data Source=" & ServerName
    
    Set aCon = ADO.CreateConnection
    aCon.ConnectionString = cConnectionString
    ' Opens the connection
    aCon.Open
    ' Creates a command and specifies its parameters
    Set aCmd = ADO.CreateCommand
    aCmd.ActiveConnection = aCon ' Connection
    aCmd.CommandType = adCmdStoredProc ' Command type
    aCmd.CommandText = "asTest_ChangeDateTime"
    
    aCmd.Parameters.Append aCmd.CreateParameter("@OperStart", 200, 1, 255, DateStart)
    aCmd.Parameters.Append aCmd.CreateParameter("@OperEnd", 200, 1, 255, DateEnd)
    
    aCmd.Execute
    aCon.Close
       
    Set AsBank = p1
    g_currentDate = aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%d%m%y")
    g_currentDate_For_Check = aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%d/%m/%y")
    g_currentDate_SQL = aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%Y%m%d")
    g_currentDate_Line_SQL = aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%Y-%m-%d")
End Sub

Sub Initialize_AsBank(TestArea, DateStart, DateEnd)
    if Sys.WaitProcess("Asbank", 1000).Exists Then
       Sys.Process("Asbank").Terminate
    End If            
    BuiltIn.Delay(3000)
    
    Call GetConfigInformation(TestArea, ServerName, DatabaseName)
    
   TestedApps.Asbank.Parameters = """"& ServerName &""" """& DatabaseName &""" ""armsoft"""
    
    Set p1 = TestedApps.Items("Asbank").Run
    If (Not Sys.Process("Asbank").Exists) Then
        Call Log.Error("ASBank application not found")
        Exit Sub
    End If
    
    ''''''''''''''''''''''''''''
    '  Set Proc = Sys.Process("AsBank")
    '  R = SetAffinity(Proc.ProcessID, 1)
    ''''''''''''''''''''''''''''
    '  Call p1.VBObject("frmLogin").VBObject("txtUserName").Keys("armsoft")
    '  Select Case dataBase
    '  Case "bankShowTesting"
    '    Call p1.VBObject("frmLogin").VBObject("cmbConfig").Keys("b")
    '  Case "AregakT"
    '    Call p1.VBObject("frmLogin").VBObject("cmbConfig").Keys("a")
    '  End Select
    '  Call p1.VBObject("frmLogin").VBObject("cmdOK").ClickButton
    
    k1 = Sys.Process("Asbank").vbObject("MainForm").WaitProperty("VisibleOnScreen", True, 20000)
    Set wMainForm = Sys.Process("Asbank").vbObject("MainForm")
    
    k2 = wMainForm.Window("MDIClient").WaitProperty("VisibleOnScreen", True, 20000)
    Set wMDIClient = wMainForm.Window("MDIClient")
    
    Set wMainForm = Sys.Process("Asbank").vbObject("MainForm")
    Set wMDIClient = wMainForm.Window("MDIClient")
    
    k3 = wMDIClient.vbObject("frmExplorer").WaitProperty("VisibleOnScreen", True, 20000)
    Set wfrmExplorer = wMDIClient.vbObject("frmExplorer")
    
    k4 = wfrmExplorer.vbObject("tvTreeView").WaitProperty("VisibleOnScreen", True, 20000)
    Set wTreeView = wfrmExplorer.vbObject("tvTreeView")
    
    '
    '  ' Creates ADO connection
    '  Set aCon = ADO.CreateConnection
    '  ' Sets up the connection parameters
    '   'aCon.ConnectionString = "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=bankShowTesting;Data Source=BANK-SERVER"
    '  Select Case dataBase
    '  Case "bankShowTesting"
    '    '  aCon.ConnectionString = cConnectionString_BankShowTesting
    '    cConnectionString = GetConnectionString(1)
    '  Case "AregakT"
    '      'aCon.ConnectionString = cConnectionString_AregakT
    '    cConnectionString = GetConnectionString(2)
    '  End Select

    BuiltIn.delay(1000)
    cConnectionString = "Provider=SQLOLEDB.1;Password=sasa111;Persist Security Info=True;User ID=sa;Initial Catalog="& DatabaseName &";Data Source=" & ServerName
    
    Set aCon = ADO.CreateConnection
    aCon.ConnectionString = cConnectionString
    ' Opens the connection
    aCon.Open
    ' Creates a command and specifies its parameters
    Set aCmd = ADO.CreateCommand
    aCmd.ActiveConnection = aCon ' Connection
    aCmd.CommandType = adCmdStoredProc ' Command type
    aCmd.CommandText = "asTest_ChangeDateTime"
    
    aCmd.Parameters.Append aCmd.CreateParameter("@OperStart", 200, 1, 255, DateStart)
    aCmd.Parameters.Append aCmd.CreateParameter("@OperEnd", 200, 1, 255, DateEnd)
    
    aCmd.Execute
    aCon.Close    
    
    Set AsBank = p1
    g_currentDate = aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%d%m%y")
    g_currentDate_For_Check = aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%d/%m/%y")
    g_currentDate_SQL = aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%Y%m%d")
    g_currentDate_Line_SQL = aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%Y-%m-%d")
End Sub

'---------------------------------------------------------------------------------------------
Sub Close_AsBank
    Dim wMainForm
    Set wMainForm = Sys.Process("Asbank").VBObject("MainForm") 
    Set wMDIClient = wMainForm.Window("MDIClient", "", 1) 
    Set wMainForm = Sys.Process("Asbank").vbObject("MainForm")
    Call wMainForm.MainMenu.Click("Աշխատանք|Աշխատանքի ավարտ")
    Sys.Process("Asbank").Close
    Delay(delay_middle)
End Sub

'-------------------------------db params---------------------------------
'--------------------------------------------------------
' ì»ñ³¹³ñÓÝáõÙ ¿ å³ñ³Ù»ïñÇ ³ñÅ»ùÁ
' »Ã» Ï³ Ïïñí³Íù Áëï Áëï û·ï³·áñÍáÕÇ, ³å³ ïíÛ³É û·ï³·áñÍáÕÇ Ñ³Ù³ñ ³ñÅ»ùÁ
'--------------------------------------------------------
' paramName - å³ñ³Ù»ïñÇ ³ÝáõÝÁ
' userId    - û·ï³·áñÍáÕÇ Ïá¹Á
Public Function Param(paramName, userId)
    Dim dbConnection, dbCommand 
 
    Set dbConnection = ADO.CreateConnection
    
    dbConnection.ConnectionString = cConnectionString
    dbConnection.Open()
    
    Set dbCommand = ADO.CreateCommand
    dbCommand.ActiveConnection = dbConnection
    dbCommand.CommandType = adCmdStoredProc
    dbCommand.CommandText = "asp_ParamValue"
    
    dbCommand.Parameters.Append dbCommand.CreateParameter("@ParamId", DB.adChar, DB.adParamInput, 20, paramName)
    dbCommand.Parameters.Append dbCommand.CreateParameter("@UserID", DB.adSmallInt, DB.adParamInput,, userId)
    dbCommand.Parameters.Append dbCommand.CreateParameter("@VType", DB.adChar, DB.adParamOutput, 32)
    dbCommand.Parameters.Append dbCommand.CreateParameter("@Value", DB.adVarChar, DB.adParamOutput, 255)
    dbCommand.Parameters.Append dbCommand.CreateParameter("@SignPermanent", DB.adChar, DB.adParamOutput, 1)
    dbCommand.Parameters.Append dbCommand.CreateParameter("@UserDimension", DB.adChar, DB.adParamOutput, 1)
        
    dbCommand.Execute
          
    If IsNull(dbCommand.Parameters("@Value").Value) Then
        dbConnection.Close() 
        Param = ""
        Log.Error(paramName & " անունով պարամետր չի գտնվել")
        Exit Function 
    Else
        Param = Trim(dbCommand.Parameters("@Value").Value)    
    End If
    
    dbConnection.Close()
End Function

Public Sub SetParameter(paramName, paramValue)
    Set dbConnection = ADO.CreateConnection
    
    dbConnection.ConnectionString = cConnectionString
    dbConnection.Open()
    
    Set dbCommand = ADO.CreateCommand
    dbCommand.ActiveConnection = dbConnection
    dbCommand.CommandType = adCmdText
    dbCommand.CommandText = " Update PARAMS set fVALUE = '" & paramValue &"' where fPARID = '" & paramName & "'"
    
    Set commandResult = dbCommand.Execute
    dbConnection.Close()
End Sub

'------------------- printing functionality------------------------------------------------------
' ---------------------------------------------------------------------------------------------------
Public Function PrintDocument(fISN, printFormat, templateName, originalDoc, NotInListTemplates, bSave)   
    Call LocateDocument(fISN)
    Call Print(printFormat, templateName, NotInListTemplates)
    
    If bSave Then
        resultFileName = docspath & "\Actual Documents\" & templateName &"_"& _
                         fISN & "." &GetTemplateExtension(templateName)
    Else
        resultFileName = ""
    End If
    
    Call SavePrintForm(fISN, printFormat, resultFileName)
    wMDIClient.frmPttel.Close()
    
    PrintDocument = resultFileName
End Function

'-----------------------------------------------------------------------------------------------
Private Sub Print(printFormat, templateName, NotInListTemplates)
    Set asbank = p1
    wMDIClient.frmPttel.SetFocus()
    
    If printFormat = "word" Then
        Call Sys.Process("Asbank").VBObject("MainForm").VBObject("tbToolBar").Window("ToolbarWindow32", "", 1).ClickItem(21)
    ElseIf printFormat = "excel" Then
        Call Sys.Process("Asbank").VBObject("MainForm").VBObject("tbToolBar").Window("ToolbarWindow32", "", 1).ClickItem(22)
    End If
    
    Set frmModalBrowser = asbank.WaitVBObject("frmModalBrowser", 1000)
    If frmModalBrowser.Exists = True Then
        
        For Each tmp in NotInListTemplates
            frmModalBrowser.tdbgView.MoveFirst()
            
            Do Until frmModalBrowser.tdbgView.EOF
                If frmModalBrowser.tdbgView.Columns.Item(0).Text = tmp Then
                    Log.Error("Item from Not in list array found " & tmp)
                    Exit Do
                End If
                frmModalBrowser.tdbgView.MoveNext()
            Loop
        Next
        
        frmModalBrowser.tdbgView.MoveFirst()
        Do Until frmModalBrowser.tdbgView.EOF
            If frmModalBrowser.tdbgView.Columns.Item(0).Text = templateName Then
                frmModalBrowser.Keys("[Enter]")
                Exit Do
            End If
            frmModalBrowser.tdbgView.MoveNext()
        Loop
    End If
    
    If printFormat = "word" Then
        WaitForWordToFill(templateName)
    ElseIf printFormat = "excel" Then
        WaitForExcelToFill(templateName)
    End If
End Sub

'-----------------------------------------------------------------------------------------------
Private Function GetTemplateExtension(templateName)
    GetTemplateExtension = ""
    Set dbConnection = ADO.CreateConnection
    
    dbConnection.ConnectionString = cConnectionString
    dbConnection.Open
    
    Set dbCommand = ADO.CreateCommand
    dbCommand.ActiveConnection = dbConnection
    dbCommand.CommandType = adCmdText
    dbCommand.CommandText = " select fFILE from TEMPLATES where fCAPTION='" & templateName & "'"
    
    Set commandResult = dbCommand.Execute
    
    If commandResult.RecordCount = 0 Then
        errorText = "Record with Name=" & templateName & " could not be found."
        Log.Error(errorText)
    Else
        Log.Message("Record was found!!!")
        If commandResult.RecordCount = 1 Then
            filePath = Trim(commandResult("fFILE").Value)
            If filePath <> "" Then
                pos = InStrRev(filePath, ".")
                GetTemplateExtension = Mid(filePath, pos + 1, Len(filePath) - pos)
            End If '
        Else
            Log.Error("More then one record found.")
        End If
        
    End If
    dbConnection.Close
End Function

'-----------------------------------------------------------------------------------------------
Public Sub SavePrintForm(fISN, printFormat, resultFileName)   
    If printFormat = "word" Then
        Set wordApp = WaitForObject("Word.Application", printTimeOut)
        If wordApp Is Nothing Then
            Log.Error("word document was not found.")
        Else
            If resultFileName <> "" Then
                wordApp.ActiveDocument.saveas(resultFileName)
            End If
            
            wordApp.ActiveDocument.Close()
            wordApp.Quit()
            Set wordApp = Nothing
            sleepTime = 0
            Do While sleepTime < printTimeOut And Sys.WaitProcess("WINWORD", stepTimeOut).Exists
                Delay(stepTimeOut)
                sleepTime = sleepTime + stepTimeOut
            Loop
            
            If Sys.WaitProcess("WINWORD", stepTimeOut).Exists Then
                Call Sys.Process("WINWORD").Terminate
                Log.Warning("WINWORD process was terminatied by Test Complete")
            End If
        End If
        
    ElseIf printFormat = "excel" Then
        Set xlApp = WaitForObject("Excel.Application", printTimeOut)
        If xlapp Is Nothing Then
            Log.Error("excel was not found.")
        Else
            If resultFileName <> "" Then
                Xlapp.activeworkbook.saveas(resultFileName)
            End If
            Xlapp.ActiveWorkbook.Close SaveChanges = False
            Xlapp.Quit()
            Set xlapp = Nothing     
            If Sys.WaitProcess("EXCEL", stepTimeOut).Exists Then
                Call Sys.Process("EXCEL").Terminate
                Log.Warning("EXCEL process was terminatied by Test Complete")
            End If
        End If
    End If
End Sub

'-----------------------------------------------------------------------------------------------
Public Function CheckDocReadOnly(docPath)
    
    Set objWord = CreateObject("Word.Application")
    Set objDoc = objWord.Documents.Open(docPath)
    
    If objDoc.ProtectionType = 2 Then
        CheckDocReadOnly = True
    Else
        CheckDocReadOnly = False
    End If
    
    objDoc.Close
    objWord.Quit
End Function

'-----------------------------------------------------------------------------------------------
Public Function WaitForObject(objType, timeout)
    sleepTime = 0
    Set obj = Nothing
    While sleepTime < timeout And obj Is Nothing
        delay(stepTimeOut)
        sleepTime = sleepTime + stepTimeOut
        Set obj = GetObject(, objType)
    Wend
    Set WaitForObject = obj
End Function

'-------------------------------------------------------------------------------------------------
Public Function WaitForWordToFill(Templ)
    WaitForWordToFill = WaitForDocToFill("WINWORD",Templ)
End Function

'-------------------------------------------------------------------------------------------------
Public Function WaitForExcelToFill(Templ)
    WaitForExcelToFill = WaitForDocToFill("EXCEL",Templ)
End Function

''''''''''''''''''''''''''''''''''''''''''''''''' 
'''''msgbox ýáñÙ³ÛÇ ·áÛáõÃÛ³Ý ëïáõ·áõÙ''''''''''' 
'''''''''''''''''''''''''''''''''''''''''''''''''
Function Is_MSG_Exist(tYPE_MSG)
  Dim  my_vbObj,rESULT
  Sys.Process("Asbank").Refresh
  rESULT=false
  Set my_vbObj = Sys.Process("Asbank").WaitVBObject(tYPE_MSG, delay_middle)
   If my_vbObj.Exists Then
     rESULT=True
   end If
  Is_MSG_Exist=rESULT
End Function

'-------------------------------------------------------------------------------------------------
Private Function WaitForDocToFill(docType,Templ)
    Dim IsVisible , CheckModalWin, tYPE_MSG, CountRow, TemplExist, TemplCheck , C_count
    Log.Message("start waiting for " & docType & " doc" & aqConvert.DateTimeToStr(aqDateTime.Time))
    bSuccesful = True
    
    Set doc = Sys.WaitProcess(docType, WaitWordTimeOut)
    If doc.Exists Then
        '  ëå³ë»É ÙÇÝã¨  WINWORD Ï³Ù EXCEL ÑÇÙ³Ý³Ï³Ý å³ïáõÑ³ÝÁ »ñ¨³ ¿Ïñ³ÝÇÝ
        Select Case doctype
            Case "WINWORD"
                Set window = doc.WaitWindow("WINWORD","Microsoft Word Document" , 1, WaitDocStartFilling)
            Case "EXCEL"
                Set window = doc.WaitWindow("XLMAIN", , , WaitDocStartFilling)
            Case Else
                window = Null
        End Select
        
        If Not (doc.WaitProperty("CPUUsage", 0, FillDocumentDelay)) Then
            Log.Error("Probably the report hanged")
            bSuccesful = False
        End If
    Log.Message("exixts")
    Else
        CheckModalWin = False
        tYPE_MSG = "frmModalBrowser"
        CheckModalWin = Is_MSG_Exist(tYPE_MSG)
        If CheckModalWin Then
            CountRow = Sys.Process("Asbank").vbObject("frmModalBrowser").vbObject("tdbgView").ApproxCount
            Sys.Process("Asbank").vbObject("frmModalBrowser").vbObject("tdbgView").Click()
            CheckModalWin = False
            tYPE_MSG = "frmModalBrowser"
            CheckModalWin = Is_MSG_Exist(tYPE_MSG)
            If CheckModalWin Then
                Do While CountRow>0
                    TemplCheck = Trim(Sys.Process("Asbank").vbObject("frmModalBrowser").vbObject("tdbgView").NativeVBObject)
                    If TemplCheck = Templ Then
                        Log.Message(Sys.Process("Asbank").vbObject("frmModalBrowser").vbObject("tdbgView").Bookmark)
                        Sys.Process("Asbank").vbObject("frmModalBrowser").vbObject("tdbgView").Keys("[Enter]")
                        Set doc = Sys.WaitProcess(docType, WaitWordTimeOut)
                        If doc.Exists Then
                            '  ëå³ë»É ÙÇÝã¨  WINWORD Ï³Ù EXCEL ÑÇÙ³Ý³Ï³Ý å³ïáõÑ³ÝÁ »ñ¨³ ¿Ïñ³ÝÇÝ
                            Select Case doctype
                                Case "WINWORD"
                                    Set window = doc.WaitWindow("OpusApp", "*- Microsoft Word", 1, WaitDocStartFilling)
                                Case "EXCEL"
                                    Set window = doc.WaitWindow("XLMAIN", , , WaitDocStartFilling)
                                Case Else
                                    window = Null
                            End Select
                            If Not (doc.WaitProperty("CPUUsage", 0, FillDocumentDelay) )Then
                                Log.Error("Probably the report hanged")
                                bSuccesful = False
                            End If
                        End If
                        CountRow = 0
                    Else
                        'Ñ³çáñ¹ ïáÕÇ ³ÝóáõÙ
                        Sys.Process("Asbank").vbObject("frmModalBrowser").vbObject("tdbgView").MoveNext
                        CountRow = CountRow -1
                        Log.Message(Sys.Process("Asbank").vbObject("frmModalBrowser").vbObject("tdbgView").Bookmark)
                    End If
                Loop
            End If
        Else
            Select Case errorType
                Case 0
                    Log.Message("No active " & docType & " process found")
                Case 1
                    Log.Warning("No active " & docType & " process found")
                Case 2
                    Log.Error("No active " & docType & " process found")
            End Select
            bSuccesful = False
        End If
    End If
    ' »Ã» Ñ³çáÕ ¿ ³í³ñïí»É ÙÇ ùÇã ¿É ëå³ë»É
    ' »Ã» áã ³å³ ¿É ÑáõÛë ãÏ³
    If bSuccesful Then
        Log.Message("Window opened succesful")
    End If
    WaitForDocToFill = bSuccesful
    Log.Message("end waiting for " & docType & " doc" & aqConvert.DateTimeToStr(aqDateTime.Time))
End Function

'-------------------------------------------------------------------------------------------------
Public Function LocateDocument(fISN)
    wMDIClient.frmExplorer.SetFocus()
    Call wMDIClient.frmExplorer.vbObject("tvTreeView").DblClickItem("|²¹ÙÇÝÇëïñ³ïáñÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|êï»ÕÍí³Í ÷³ëï³ÃÕÃ»ñ")
    
    Set control = asbank.frmAsUstPar
    Set control2 = control.vbObject("TabFrame")
    
    Call control2.TDBNumber.Keys(fISN & "[Enter]")
    'Call control2.TDBNumber_2.Keys(fISN & "[Enter]")
    
    control2.AsTypeView.TDBMask.Text = ""
    Call control2.AsTypeView.TDBMask.Keys("[Enter]")
    Call control.vbObject("CmdOK").ClickButton()
    
    Set frmPttel = wMDIClient.WaitvbObject("frmPttel", delay_small)
    If Not frmPttel.Exists Then
        Log.Error("frmPttel was expected.")
    Else
        Set grid = wMDIClient.vbObject("frmPttel").vbObject("tdbgView")
        If grid.EOF Then
            LocateDocument = False
        Else
            LocateDocument = True
        End If
    End If
End Function

'---------------------------------------------------------------------------------------------
'¸³ï³ñÏ»É ÃÕÃ³å³Ý³ÏÇ å³ñáõÝ³ÏáõÃÛáõÝÁ
'---------------------------------------------------------------------------------------------
'folderpath - ÃÕÃ³å³Ý³ÏÇ ×³Ý³å³ñÑÁ
Function Empty_Folder(folderpath)
  Dim fso,fFolder, fFiles, file

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fFolder = fso.GetFolder(folderpath)
    Set fFiles = fFolder.Files
  
    For Each file in fFiles
        Utilities.DeleteFile(file)  
        BuiltIn.Delay(3000) 
    Next   
End Function

'---------------------------------------------------------------------------------------------
'ê»ÕÙ»É Ññ³Ù³Ý³ÛÇÝ Ïá×³ÏÁ
'---------------------------------------------------------------------------------------------
'formType  - Éñ³óíáÕ ýáñÙ³ÛÇ ï»ë³ÏÁ
'    1 - Document 
'    2 - Dialog
'    3 - DeleteDoc
'    4 - Modal Document
'    5 - Message Box 
'    6 - Options
'caption - ë»ÕÙíáÕ Ïá×³ÏÇ ³ÝáõÝÁ
'formType - µ³óí³Í å³ïáõÑ³ÝÇ ï»ë³ÏÁ ûñ.` 2 - "frmASDocForm"
'³ñÅ»ùÁ Ï³ñ»ÉÇ ¿ ·ñ»É Ý³¨ Ñ»ï¨Û³É Ï»ñå` "2_2"
'³é³çÇÝ 2-Á å³ïáõÑ³ÝÇ ï»ë³ÏÇ Ñ³Ù³ñ ¿
'»ñÏñáñ¹Á` ÝáõÛÝ ïÇåÇ µ³óí³Í å³ïáõÑ³ÝÝ»ñÇó` 2-ñ¹Á
'¹»åùÁ ·áñÍáõÙ ¿, »ñµ å³ïáõÑ³ÝÇ ³ÝáõÝÁ áõÝÇ ÝÙ³Ý³ïÇå Ó¨` "frmASDocForm_2"
Public Sub ClickCmdButton(formType, caption)
Dim arrayProp, arrValues, objButton, winN, form_Type
  
    'winN - ÷á÷áË³Ï³ÝÁ ëï³ÝáõÙ ¿ formType-Ç í»ñçÇÝ »ñÏáõ ÝÇß»ñÁ
    winN = Right(formType, 2)
    '»Ã» winN-Ç ³é³çÇÝ ÝÇßÁ "_" ¿, ³å³ form_Type ÷á÷áË³Ï³ÝÇÝ í»ñ³·ñáõÙ »Ýù ¿ formType-Ç ³ñÅ»ùÁ Ñ³Ý³Í í»ñçÇÝ »ñÏáõ ÝÇßÁ
    'Ñ³Ï³é³Ï ¹»åùáõÙ form_Type-ÇÝ í»ñ³·ñáõÙ formType ³ñÅ»ùÁ
    If Left(winN, 1) = "_" Then 
        form_Type = Left(formType, Len(formType) - 2)
    Else 
        form_Type = formType
        winN = ""
    End If 
		 
    BuiltIn.Delay(1500)
    arrayProp = Array("Caption", "WndClass")
    arrValues = Array(caption, "ThunderRT6CommandButton")
    wMDIClient.Refresh
		
    Select Case form_Type
        Case 1 ' Document
            Set objButton = wMDIClient.VBObject("frmASDocForm" & winN).FindChild(arrayProp, arrValues, 50)
 
        Case 2 ' Dialog
            Set objButton = p1.vbObject("frmAsUstPar" & winN).FindChild(arrayProp, arrValues)
      
        Case 3 ' DeleteDoc
            Set objButton = p1.VBObject("frmDeleteDoc" & winN).FindChild(arrayProp, arrValues)
          
        Case 4 ' Modal View of Document
            Set objButton = p1.vbObject("frmASDocFormModal" & winN).FindChild(arrayProp, arrValues)  
     
        Case 5 ' Message Box 
            Set objButton = p1.VBObject("frmAsMsgBox" & winN).FindChild(arrayProp, arrValues)  
										
        Case 6 ' Options
            Set objButton = p1.vbObject("frmOptions" & winN).FindChild(arrayProp, arrValues)
            
        Case 7 ' frmPttelFilter
            Set objButton = p1.vbObject("frmPttelFilter" & winN).FindChild(arrayProp, arrValues)	
						
        Case 8 ' frmTreeNode
            Set objButton = p1.vbObject("frmTreeNode" & winN).FindChild(arrayProp, arrValues)		
						
        Case 9 ' frmEditJob
            Set objButton = p1.vbObject("frmEditJob" & winN).FindChild(arrayProp, arrValues)		
						
        Case 10 ' frmRolePropN
            Set objButton = p1.vbObject("frmRolePropN" & winN).FindChild(arrayProp, arrValues)
			
        Case 11 ' frmRefuseComment
            Set objButton = p1.vbObject("frmRefuseComment" & winN).FindChild(arrayProp, arrValues)
      
        Case 12 ' frmSprFind
            Set objButton = p1.vbObject("frmSprFind" & winN).FindChild(arrayProp, arrValues)
                
        Case 13 'frmImport
            Set objButton = wMDIClient.vbObject("frmImport" & winN).FindChild(arrayProp, arrValues) 
            
        Case 14 'frmOLAPExp
            Set objButton = wMDIClient.vbObject("frmOLAPExp" & winN).FindChild(arrayProp, arrValues) 
                
        Case 0 ' frmLogin
            Set objButton = p1.vbObject("frmLogin" & winN).FindChild(arrayProp, arrValues)
    End Select
		
    objButton.Click()  
    BuiltIn.Delay(1000)
End Sub


'---------------------------------------------------------------------------------------------
'Î³ÝãáõÙ ¿ ¹³ßïÇ ClickDropDown-Á Ctrl+Ü»ñù¨ Ïá×³ÏÝ»ñÇ ÙÇçáóáí
'---------------------------------------------------------------------------------------------
'formType  - Éñ³óíáÕ ýáñÙ³ÛÇ ï»ë³ÏÁ
'    1 - Document 
'    2 - Dialog
'rekvName  - ¹³ßïÇ ³ÝáõÝÁ
Public Sub ClickDropDown(formType, rekvName)
    Dim rekvObj

    Select Case formType
    Case 1 ' Document      
        rekvObj = GetVBObject(rekvName, wMDIClient.vbObject("frmASDocForm"))
        wMDIClient.vbObject("frmASDocForm").vbObject("TabFrame").vbObject(rekvObj).Keys("^[Down]")      
    Case 2 ' Dialog          
        rekvObj = GetVBObject_Dialog("rekvName", Sys.Process("Asbank").vbObject("frmAsUstPar"))
        Sys.Process("Asbank").vbObject("frmAsUstPar").vbObject("TabFrame").vbObject(rekvObj).Keys("^[Down]")              
    End Select
End Sub

'---------------------------------------------------------------------------------------------
'Èñ³óÝ»É ¹³ßïÇ ³ñÅ»ùÁ
'---------------------------------------------------------------------------------------------
'    1 - Document 
'    2 - Dialog
'tabN - ¾çÇ Ñ³Ù³ñÁ
'rekvType  - ¹³ßïÇ ï»ë³ÏÁ
'    General (ÁÝ¹Ñ³Ýáõñ)
'    Masc   (? Ýß³ÝÝ»ñáí ¿ Éñ³óí³Í) 
'    CheckBox (ÜßÇã)
'rekvName  - ¹³ßïÇ ³ÝáõÝÁ
'rekvValue - Éñ³óíáÕ ³ñÅ»ùÁ
'    Null-Ç ¹»åùáõÙ áãÇÝã ãÇ Éñ³óíÇ, ÙÛáõë ¹»åù»ñáõÙ ¹³ßïÇ ³ñÅ»ùÁ Ï÷áË³ñÇÝíÇ ÷áË³Ýó³Í ³ñÅ»ùáí
'    Ch ï»ë³ÏÇ ¹»åùáõÙ ÷áË³Ýó»É 0(Üßí³Í ¿) Ï³Ù 1(Üßí³Í ã¿) 
'formType - µ³óí³Í å³ïáõÑ³ÝÇ ï»ë³ÏÁ ûñ.` 2 - "frmASDocForm"
'³ñÅ»ùÁ Ï³ñ»ÉÇ ¿ ·ñ»É Ý³¨ Ñ»ï¨Û³É Ï»ñå` "2_2"
'³é³çÇÝ 2-Á å³ïáõÑ³ÝÇ ï»ë³ÏÇ Ñ³Ù³ñ ¿
'»ñÏñáñ¹Á` ÝáõÛÝ ïÇåÇ µ³óí³Í å³ïáõÑ³ÝÝ»ñÇó` 2-ñ¹Á
'¹»åùÁ ·áñÍáõÙ ¿, »ñµ å³ïáõÑ³ÝÇ ³ÝáõÝÁ áõÝÇ ÝÙ³Ý³ïÇå Ó¨` "frmASDocForm_2"
Public Sub Rekvizit_Fill(formType, tabN, rekvType, rekvName, rekvValue)
  
  Dim rekvObj, sTab, wTabStrip, winN, form_Type
  
		'winN - ÷á÷áË³Ï³ÝÁ ëï³ÝáõÙ ¿ formType-Ç í»ñçÇÝ »ñÏáõ ÝÇß»ñÁ
		winN = Right(formType, 2)
		'»Ã» winN-Ç ³é³çÇÝ ÝÇßÁ "_" ¿, ³å³ form_Type ÷á÷áË³Ï³ÝÇÝ í»ñ³·ñáõÙ »Ýù ¿ formType-Ç ³ñÅ»ùÁ Ñ³Ý³Í í»ñçÇÝ »ñÏáõ ÝÇßÁ
		'Ñ³Ï³é³Ï ¹»åùáõÙ form_Type-ÇÝ í»ñ³·ñáõÙ formType ³ñÅ»ùÁ
		if Left(winN, 1) = "_" then 
				form_Type = Left(formType, Len(formType) - 2)
		else 
			form_Type = formType
			winN = ""
		end if 
	
  Select case form_Type
    Case "Document"
						sTab = "TabFrame"
						If tabN <> 1 Then
								sTab = sTab & "_" & tabN
								Set wTabStrip = wMDIClient.VBObject("frmASDocForm" & winN).VBObject("TabStrip")
								wTabStrip.SelectedItem = wTabStrip.Tabs(tabN)
						End If
						'      Set wMDIClient = Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1)
						Select Case rekvType
								Case "General"
										wMDIClient.Refresh
										rekvObj = GetVBObject(rekvName, wMDIClient.vbObject("frmASDocForm" & winN))
										If Not wMDIClient.vbObject("frmASDocForm" & winN).vbObject(sTab).vbObject(rekvObj).ReadOnly Then
												' wMDIClient.vbObject("frmASDocForm").vbObject(sTab).vbObject(rekvObj).Keys("![End]" & "[Del]")
												wMDIClient.vbObject("frmASDocForm").vbObject(sTab).vbObject(rekvObj).Keys(rekvValue & "[Tab]")
										End If
								Case "CheckBox"
										rekvObj = GetVBObject(rekvName, wMDIClient.vbObject("frmASDocForm" & winN))
										If wMDIClient.vbObject("frmASDocForm" & winN).vbObject(sTab).vbObject(rekvObj).Enabled Then
												wMDIClient.vbObject("frmASDocForm" & winN).vbObject(sTab).vbObject(rekvObj).Value = rekvValue
												wMDIClient.vbObject("frmASDocForm" & winN).vbObject(sTab).vbObject(rekvObj).Keys("[Tab]")
										End If
								'        Case "Mask" 
								'            Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmASDocForm").VBObject("TabFrame").VBObject("ASTypeTree").VBObject("CmdViewTree").Click()
								'            Set Sys.Process("Asbank").VBObject("frmDynamicTreeBrow").VBObject("TreeBrow").wSelection = rekvVAlue
								Case Else
										Log.Error("Unknown rekvizit type of document.")
						End select
						
				Case "DocumentModal"
						sTab = "TabFrame"
						If tabN <> 1 Then
								sTab = sTab & "_" & tabN
								Set wTabStrip = p1.VBObject("frmASDocFormModal" & winN).VBObject("TabStrip")
								wTabStrip.SelectedItem = wTabStrip.Tabs(tabN)
						End If
						Select Case rekvType
								Case "General"
										p1.Refresh
										rekvObj = GetVBObject(rekvName, p1.vbObject("frmASDocFormModal" & winN))
										If Not p1.vbObject("frmASDocFormModal" & winN).vbObject(sTab).vbObject(rekvObj).ReadOnly Then
												p1.vbObject("frmASDocFormModal").vbObject(sTab).vbObject(rekvObj).Keys(rekvValue & "[Tab]")
										End If
								Case "CheckBox"
										rekvObj = GetVBObject(rekvName, p1.vbObject("frmASDocFormModal" & winN))
										If p1.vbObject("frmASDocFormModal" & winN).vbObject(sTab).vbObject(rekvObj).Enabled Then
												p1.vbObject("frmASDocFormModal" & winN).vbObject(sTab).vbObject(rekvObj).Value = rekvValue
												p1.vbObject("frmASDocFormModal" & winN).vbObject(sTab).vbObject(rekvObj).Keys("[Tab]")
										End If
								Case Else
										Log.Error("Unknown rekvizit type of document.")
						End select
				
				Case "TreeNode"
'						sTab = "TabFrame"
						If tabN <> 1 Then
								sTab = sTab & "_" & tabN
								Set wTabStrip = p1.VBObject("frmTreeNode" & winN)
								wTabStrip.SelectedItem = wTabStrip.Tabs(tabN)
						End If
						Select Case rekvType
								Case "General"
										p1.Refresh
'										rekvObj = GetVBObject(rekvName, p1.vbObject("frmTreeNode" & winN))
										p1.vbObject("frmTreeNode").vbObject(rekvName).Keys(rekvValue & "[Tab]")
										
								Case "CheckBox"
'										rekvObj = GetVBObject(rekvName, p1.vbObject("frmTreeNode" & winN))
										If p1.vbObject("frmTreeNode" & winN).vbObject(rekvName).Enabled Then
												p1.vbObject("frmTreeNode" & winN).vbObject(rekvName).Value = rekvValue
												p1.vbObject("frmTreeNode" & winN).vbObject(rekvName).Keys("[Tab]")
										End If
								Case Else
										Log.Error("Unknown rekvizit type of document.")
						End select
						
						Case "EditJob"
'						sTab = "TabFrame"
						If tabN <> 1 Then
								sTab = sTab & "_" & tabN
								Set wTabStrip = p1.VBObject("frmEditJob" & winN)
								wTabStrip.SelectedItem = wTabStrip.Tabs(tabN)
						End If
						Select Case rekvType
								Case "General"
										p1.Refresh
'										rekvObj = GetVBObject(rekvName, p1.vbObject("frmTreeNode" & winN))
										p1.vbObject("frmEditJob").vbObject(rekvName).Keys(rekvValue & "[Tab]")
										
								Case "CheckBox"
'										rekvObj = GetVBObject(rekvName, p1.vbObject("frmTreeNode" & winN))
										If p1.vbObject("frmEditJob" & winN).vbObject(rekvName).Enabled Then
'												p1.vbObject("frmEditJob" & winN).vbObject(rekvName).Value = rekvValue
												p1.vbObject("frmEditJob" & winN).vbObject(rekvName).Keys("[Tab]")
										End If
								Case Else
										Log.Error("Unknown rekvizit type of document.")
						End select
      
				Case "Dialog"
						rekvObj = GetVBObject_Dialog(rekvName, p1.VBObject("frmAsUstPar" & winN))
						sTab = "TabFrame"
						If tabN <> 1 Then
								sTab = sTab & "_" & tabN
								Set wTabStrip = p1.vbObject("frmAsUstPar" & winN).vbObject("TabStrip")
								wTabStrip.SelectedItem = wTabStrip.Tabs(tabN)
						End If
						Select Case rekvType
								Case "General"
										rekvObj = GetVBObject_Dialog(rekvName, p1.vbObject("frmAsUstPar" & winN))
										'              Sys.Process("Asbank").vbObject("frmAsUstPar").vbObject(sTab).vbObject(rekvObj).Keys("![End]" & "[Del]")
										p1.vbObject("frmAsUstPar" & winN).vbObject(sTab).vbObject(rekvObj).Keys(rekvValue & "[Tab]")
								Case "CheckBox"
										rekvObject = GetVBObject_Dialog(rekvName, p1.vbObject("frmAsUstPar" & winN))
										p1.vbObject("frmAsUstPar" & winN).vbObject(sTab).vbObject(rekvObj).Value = rekvValue
										'Sys.Process("Asbank").vbObject("frmAsUstPar").vbObject(sTab).vbObject(rekvObj).Keys("[Tab]")
								Case Else
										Log.Error("Unknown rekvizit type of document.")
						End Select  
      
    Case "OLAPNavigator"
						Select Case rekvType
								Case "General"
										wMDIClient.Refresh
										If Not wMDIClient.vbObject("frmOLAPNav").vbObject(rekvName).ReadOnly Then
												wMDIClient.vbObject("frmOLAPNav").vbObject(rekvName).Keys(rekvValue & "[Tab]")
										End If
								Case "CheckBox"
										rekvObj = GetVBObject(rekvName, wMDIClient.vbObject("frmOLAPNav"))
										If wMDIClient.vbObject("frmASDocForm").vbObject(rekvObj).Enabled Then
												wMDIClient.vbObject("frmASDocForm").vbObject(rekvObj).Value = rekvValue
												wMDIClient.vbObject("frmASDocForm").vbObject(rekvObj).Keys("[Tab]")
										End If
								'        Case "Mask" 
								'            Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmASDocForm").VBObject("TabFrame").VBObject("ASTypeTree").VBObject("CmdViewTree").Click()
								'            Set Sys.Process("Asbank").VBObject("frmDynamicTreeBrow").VBObject("TreeBrow").wSelection = rekvVAlue
								Case Else
										Log.Error("Unknown rekvizit type of document.")
						End select
      
      Case "OLAPExport"
						Select Case rekvType
								Case "General"
										wMDIClient.Refresh
										If Not wMDIClient.vbObject("frmOLAPExp").vbObject(rekvName).ReadOnly Then
												wMDIClient.vbObject("frmOLAPExp").vbObject(rekvName).Keys(rekvValue & "[Tab]")
										End If
								Case "CheckBox"
										If wMDIClient.vbObject("frmOLAPExp").vbObject(rekvName).Enabled Then
												wMDIClient.vbObject("frmOLAPExp").vbObject(rekvName).Value = rekvValue
												wMDIClient.vbObject("frmOLAPExp").vbObject(rekvName).Keys("[Tab]")
										End If
								'        Case "Mask" 
								'            Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmASDocForm").VBObject("TabFrame").VBObject("ASTypeTree").VBObject("CmdViewTree").Click()
								'            Set Sys.Process("Asbank").VBObject("frmDynamicTreeBrow").VBObject("TreeBrow").wSelection = rekvVAlue
								Case Else
										Log.Error("Unknown rekvizit type of document.")
						End select
  End select
End Sub

'---------------------------------------------------------------------------------------------
'Վերցնել ¹³ßïÇ ³ñÅ»ùÁ
'---------------------------------------------------------------------------------------------
'formType  - ýáñÙ³ÛÇ ï»ë³ÏÁ
'    1.Document 
'    2.Dialog
'tabN - ¾çÇ Ñ³Ù³ñÁ
'rekvType  - ¹³ßïÇ ï»ë³ÏÁ
'    1.General - առանց մեկնաբանության դաշտեր(ÁÝ¹Ñ³Ýáõñ և ամսաթիվ)
'    2.Mask - "TDBMask" վերջավորությամբ
'    3.CheckBox (ÜßÇã)
'    4.Label - վերցնում է դաշտին կից մեկնաբանությունը
'    5.Comment - "TDBComment" վերջավորությամբ
'    6.Course - "AsCourse" վերջավորությամբ
'    7.Number - "TDBNumber" վերջավորությամբ
'rekvName  - ¹³ßïÇ ³ÝáõÝÁ
Function Get_Rekvizit_Value(formType,tabN,rekvType,rekvName)
  
    Dim wMDIClient,rekvObj,sTab,wTabStrip
 
    Select case formType
      Case "Document"
             
        Set wMDIClient = Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1)
        sTab = "TabFrame"
        If tabN <> 1 Then
          sTab = sTab & "_" & tabN
          Set wTabStrip = wMDIClient.VBObject("frmASDocForm").VBObject("TabStrip")
          wTabStrip.SelectedItem = wTabStrip.Tabs(tabN)
        End If
              
        rekvObj = GetVBObject(rekvName, wMDIClient.vbObject("frmASDocForm"))
         
        Select Case rekvType
          Case "General"
            Get_Rekvizit_Value = wMDIClient.vbObject("frmASDocForm").vbObject(sTab).vbObject(rekvObj).Text
          Case "Mask"
            Get_Rekvizit_Value = wMDIClient.vbObject("frmASDocForm").vbObject(sTab).vbObject(rekvObj).vbObject("TDBMask").Text
										Case "Bank"
            Get_Rekvizit_Value = wMDIClient.vbObject("frmASDocForm").vbObject(sTab).vbObject(rekvObj).vbObject("TDBBank").Text
          Case "CheckBox"
            Get_Rekvizit_Value = wMDIClient.vbObject("frmASDocForm").vbObject(sTab).vbObject(rekvObj).Value
          Case "Label"
            Get_Rekvizit_Value = Trim(wMDIClient.vbObject("frmASDocForm").vbObject(sTab).vbObject(rekvObj).VBObject("TxTLabel").Text) 
          Case "Comment"
            Get_Rekvizit_Value = wMDIClient.vbObject("frmASDocForm").vbObject(sTab).vbObject(rekvObj).VBObject("TDBComment").Text
          Case "Course"
            Get_Rekvizit_Value = wMDIClient.vbObject("frmASDocForm").vbObject(sTab).vbObject(rekvObj).VBObject("AsCourse").NativeVBObject
          Case Else
            Log.Error("Unknown rekvizit type of document."),,,ErrorColor
        End select
    
      Case "Dialog"
          rekvObj = GetVBObject_Dialog(rekvName,p1.VBObject("frmAsUstPar"))
          sTab = "TabFrame"
          If tabN <> 1 Then
            sTab = sTab & "_" & tabN
            Set wTabStrip = p1.vbObject("frmAsUstPar").vbObject("TabStrip")
            wTabStrip.SelectedItem = wTabStrip.Tabs(tabN)
          End If
        Select Case rekvType
          Case "General"
            Get_Rekvizit_Value = Sys.Process("Asbank").vbObject("frmAsUstPar").vbObject(sTab).vbObject(rekvObj).Text
          Case "Mask"
            Get_Rekvizit_Value = Sys.Process("Asbank").vbObject("frmAsUstPar").vbObject(rekvObj).vbObject("TDBMask").Text
          Case "CheckBox"
            Get_Rekvizit_Value = Sys.Process("Asbank").vbObject("frmAsUstPar").vbObject(sTab).vbObject(rekvObj).Value
          Case "Label"
            Get_Rekvizit_Value = Trim(Sys.Process("Asbank").vbObject("frmAsUstPar").vbObject(sTab).vbObject(rekvObj).VBObject("TxTLabel").Text) 
          Case "Comment"
            Get_Rekvizit_Value = Sys.Process("Asbank").vbObject("frmAsUstPar").vbObject(sTab).vbObject(rekvObj).VBObject("TDBComment").Text
          Case Else
            Log.Error("Unknown rekvizit type of document."),,,ErrorColor
        End select
  End select
End Function

'---------------------------------------------------------------------------------------------
'Ստուգել ¹³ßïÇ ³ñÅ»ùÁ խմբագրվող է թէ ոչ
'---------------------------------------------------------------------------------------------
'formType  - ýáñÙ³ÛÇ ï»ë³ÏÁ
'    1.Document 
'    2.Dialog
'tabN - ¾çÇ Ñ³Ù³ñÁ
'rekvType  - ¹³ßïÇ ï»ë³ÏÁ
'    1.General - առանց մեկնաբանության դաշտեր(ÁÝ¹Ñ³Ýáõñ և ամսաթիվ)
'    2.Mask - "TDBMask" վերջավորությամբ
'    3.Comment - "TDBComment" վերջավորությամբ
'    4.CheckBox - CheckBox համար
'    5.Course - "TDBNumber" վերջավորությամբ
'    6.Course1 - "TDBNumber1" վերջավորությամբ
'    7.Course2 - "TDBNumber2" վերջավորությամբ
'rekvName  - ¹³ßïÇ ³ÝáõÝÁ
'ExpectedType սպասվող տեսակ
'    1.True - խմբագրվող է
'    2.False - խմբագրվող չէ
Function Check_ReadOnly(formType,tabN,rekvType,rekvName,ExpectedType)

    Dim CurrentStatus,wMDIClient,rekvObj,sTab,wTabStrip
 
    Select case formType
      Case "Document"
        
        Set wMDIClient = Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1)
        sTab = "TabFrame"
        If tabN <> 1 Then
          sTab = sTab & "_" & tabN
          Set wTabStrip = wMDIClient.VBObject("frmASDocForm").VBObject("TabStrip")
          wTabStrip.SelectedItem = wTabStrip.Tabs(tabN)
        End If
      
        rekvObj = GetVBObject(rekvName, wMDIClient.vbObject("frmASDocForm"))
      
        Select Case rekvType
          Case "General"
            CurrentStatus = wMDIClient.vbObject("frmASDocForm").vbObject(sTab).vbObject(rekvObj).ReadOnly
          Case "Mask"
            CurrentStatus = wMDIClient.vbObject("frmASDocForm").vbObject(sTab).vbObject(rekvObj).vbObject("TDBMask").ReadOnly
          Case "Comment"
            CurrentStatus = wMDIClient.vbObject("frmASDocForm").vbObject(sTab).vbObject(rekvObj).VBObject("TDBComment").ReadOnly
          Case "CheckBox"
            CurrentStatus = Not(wMDIClient.vbObject("frmASDocForm").vbObject(sTab).vbObject(rekvObj).Enabled)
'          Case "Course"
'            CurrentStatus = wMDIClient.vbObject("frmASDocForm").vbObject(sTab).VBObject("TDBNumber").ReadOnly
          Case "Course1"
            CurrentStatus = wMDIClient.vbObject("frmASDocForm").vbObject(sTab).vbObject(rekvObj).VBObject("TDBNumber1").ReadOnly
          Case "Course2"
            CurrentStatus = wMDIClient.vbObject("frmASDocForm").vbObject(sTab).vbObject(rekvObj).VBObject("TDBNumber2").ReadOnly
          Case Else
            Log.Error("Unknown rekvizit type of document."),,,ErrorColor
        End select
    
      Case "Dialog"
          rekvObj = GetVBObject_Dialog(rekvName,p1.VBObject("frmAsUstPar"))
          sTab = "TabFrame"
          If tabN <> 1 Then
            sTab = sTab & "_" & tabN
            Set wTabStrip = p1.vbObject("frmAsUstPar").vbObject("TabStrip")
            wTabStrip.SelectedItem = wTabStrip.Tabs(tabN)
          End If
        Select Case rekvType
          Case "General"
            CurrentStatus = Sys.Process("Asbank").vbObject("frmAsUstPar").vbObject(sTab).vbObject(rekvObj).ReadOnly
          Case "Mask"
            CurrentStatus = Sys.Process("Asbank").vbObject("frmAsUstPar").vbObject(sTab).vbObject(rekvObj).vbObject("TDBMask").ReadOnly
          Case "Comment"
            CurrentStatus = Sys.Process("Asbank").vbObject("frmAsUstPar").vbObject(sTab).vbObject(rekvObj).VBObject("TDBComment").ReadOnly
          Case "CheckBox"
            CurrentStatus = Not(Sys.Process("Asbank").vbObject("frmAsUstPar").vbObject(sTab).vbObject(rekvObj).Enabled)   
          Case Else
            Log.Error("Unknown rekvizit type of document."),,,ErrorColor
        End select
  End select
  
    If CurrentStatus = ExpectedType Then
       Log.Message "ReadOnly is correct For - "& rekvName & " - Rekvizit",,,DivideColor2
       Check_ReadOnly = True
    Else
       Log.Error "ReadOnly For -"& rekvName &"- Must be = " & ExpectedType &" ,It Was = "& CurrentStatus ,,,ErrorColor
       Check_ReadOnly = False
    End If 
End Function

'-----------------------------------------------------------------------------
' Կատարում է Caption գործողությունը և սպասում Object տիպի օբյեկտի
'-----------------------------------------------------------------------------  
Function OnClick(Caption, Object)
  Dim frmAsUstPar, frmModalBrowser, frmASDocForm, AsView, FrmSpr, Rekv
  Dim i, j, attr1
'-------------------------------------  
  Set attr1 = Log.CreateNewAttributes
  attr1.BackColor = RGB(250, 30, 100)
  attr1.Bold = True
'-------------------------------------  
  BuiltIn.Delay(1000) 
  wMainForm.MainMenu.Click(c_AllActions)
  wMainForm.PopupMenu.Click(Caption)
  BuiltIn.Delay(1000)
  
  Select Case Object
  Case "frmASDocForm"    
    Set frmASDocForm = wMDIClient.WaitVBObject("frmASDocForm", delay_small)
    If frmASDocForm.Exists Then
      frmASDocForm.Close
    Else 
      Call Log.Error("'" & Caption & "'" & "գործողությունը կատարելիս փաստաթուղը չի բացվել:",,,attr1)  
    End If 
  Case "frmDeleteDoc"  
    'Պայմանագրի վրա 'Ջնջել' գործողությունը կատարելիս կամ բացվում է "frmDeleteDoc" կամ "frmAsMsgBox"
     Set frmDeleteDoc = AsBank.WaitVBObject("frmDeleteDoc", delay_small)
     If frmDeleteDoc.Exists Then
        frmDeleteDoc.Close
     Else
        Set frmAsMsgBox = AsBank.VBObject("frmAsMsgBox")
        If frmAsMsgBox.Exists Then
          frmAsMsgBox.Close
        Else
          Call Log.Error("'" & Caption & "'" & "գործողության վրա click անելիս սպասված հաղորդագրությունը չի բացվել:",,,attr1)    
        End If 
     End If
   Case "AsView"
    'Կարող է բացվել դիալոգ
     Set frmAsUstPar = AsBank.WaitVBObject("frmAsUstPar", delay_small)
     If frmAsUstPar.Exists Then
        Asbank.VBObject("frmAsUstPar").VBObject("CmdOK").ClickButton
     End If
     
     Set AsView = wMDIClient.WaitVBObject("frmPttel_2", delay_middle)
     If AsView.Exists Then
       AsView.Close
     Else 
       Call Log.Error("'" & Caption & "'" & " թղթապանակը չի բացվել:",,,attr1)
     End If
   Case "frmAsUstPar"
     Set frmAsUstPar = AsBank.WaitVBObject("frmAsUstPar", delay_small) 
     If frmAsUstPar.Exists Then
       frmAsUstPar.Close
     Else 
       Call Log.Error("'" & Caption & "'" & "գործողության վրա click անելիս դիալոգը չի բացվել:",,,attr1)  
     End If
   Case "frmModalBrowser"
      Set frmModalBrowser = AsBank.WaitVBObject("frmModalBrowser", delay_small)
      i = 0
      While  i <> frmModalBrowser.VBObject("tdbgView").ApproxCount
        j = 0
        For j = 0 To i-1
          frmModalBrowser.VBObject("tdbgView").MoveNext
        Next
      
        Call frmModalBrowser.Keys("[Enter]")
        'Պետք է բացվի Doc
        If wMDIClient.VBObject("frmASDocForm").Exists Then
          wMDIClient.VBObject("frmASDocForm").Close
        Else 
          Call Log.Error("'" & Caption & "'" & "գործողության վրա click անելիս փաստաթուղը չի բացվել:",,,attr1)  
        End If 
        i = i+1
		BuiltIn.Delay(1000) 
        wMainForm.MainMenu.Click(c_AllActions)
        wMainForm.PopupMenu.Click(Caption)
		BuiltIn.Delay(1000) 
      Wend
      frmModalBrowser.Close 
   Case "FrmSpr"
    'Կարող է բացվել դիալոգ
     Set frmAsUstPar = AsBank.WaitVBObject("frmAsUstPar", delay_small)
     If frmAsUstPar.Exists Then
        If Caption = c_References & "|" & c_CommView Then
          frmAsUstPar.VBObject("TabFrame").VBObject("TDBDate").Keys("^A[Del]" & "[Tab]") 
        End If
        frmAsUstPar.VBObject("CmdOK").ClickButton
     End If
	 BuiltIn.Delay(2000) 
     'Կարող է հայտնվել հաղորդագրություն
     Set frmAsMsgBox = AsBank.WaitVBObject("frmAsMsgBox", delay_small)
     If frmAsMsgBox.Exists Then
        frmAsMsgBox.VBObject("cmdButton").ClickButton
     End If
	 BuiltIn.Delay(3000) 
    'Պետք է բացվի քաղավածք
     Set FrmSpr = wMDIClient.WaitVBObject("FrmSpr", delay_big)
     If FrmSpr.Exists Then
       FrmSpr.Close
     Else 
       Call Log.Error("'" & Caption & "'" & " քաղավածքը չի բացվել:",,,attr1)        
     End If

  End Select 
End Function

'---------------------------------------------------------------------------
'Կարդում է և վերարտագրում է մի ֆայլից մյուսը
'---------------------------------------------------------------------------
Sub Read_Write_File(fromFile, toFile)
  const  ForReading = 1
  Const  ForWriting = 2, ForAppending = 8, TristateFalse = 0

    ' Creates a new file object
    Set FS = Sys.OleObject("Scripting.FileSystemObject") 
    Set fFile = FS.OpenTextFile(fromFile, ForReading)
    While Not fFile.AtEndOfStream
         s = fFile.ReadLine
             ' Creates a new file object
    Set  fsk = CreateObject("Scripting.FileSystemObject")
    If Not fsk.FileExists(toFile) Then
          Set fk = fsk.CreateTextFile(toFile)
    Else
          Set fk = fsk.OpenTextFile(toFile, ForAppending, TristateFalse)
    End If
    fk.Write s
    fk.Close
    WEnd
    fFile.Close
    
    ' Creates a new file object
    Set  fs = CreateObject("Scripting.FileSystemObject")
    If Not fs.FileExists(toFile) Then
          Set fFile = fs.CreateTextFile(toFile)
    Else
          Set fFile = fs.OpenTextFile(toFile, ForAppending, TristateFalse)
    End If
    fFile.Write s
    fFile.Close
End Sub

'---------------------------------------------------------------------------
' Օրինակ՝ Left_Align("123", 6) = "123   "
'---------------------------------------------------------------------------
Function Left_Align(string, size)
  Dim length
  
  string = Trim(string)
  length = Len(string)
  If length > size Then 
    Left_Align = string
    Log.Message "The size of the string is greater than the size to be left aligned"
  Else
    Left_Align = string & SPACE(size - length)
  End If
End Function

'---------------------------------------------------------------------------
' Օրինակ՝ Right_Align("123", 6) = "   123"
'---------------------------------------------------------------------------
Function Right_Align(string, size)
  Dim length
  string = Trim(string)
  length = Len(string)
  If length > size Then 
    Right_Align = string
    Log.Message "The size of the string is greater than the size to be left aligned"
  Else
    Right_Align = SPACE(size - length) & string
  End If
End Function

'---------------------------------------------------------------------------
' տողի առկա լինելը ստուգող ֆունկցիա
'---------------------------------------------------------------------------
' PttelName - համապատասխան Pttel-ի անունը (օր.՝"frmPttel", "frmPttel_2")
' colN - Թղթապանակում սյան համարը
' SearchValue - փնտրվող արժեքը
Function SearchInPttel(PttelName, colN, SearchValue)
      Dim i, tdbgView, status

      status = False
      Set tdbgView = wMDIClient.VBObject(PttelName).VBObject("tdbgView")

      tdbgView.MoveFirst
      For i = 0 to tdbgView.ApproxCount - 1
          If Trim(tdbgView.Columns.Item(colN).Value) = Trim(SearchValue)  Then
              status = True   
														Exit For        
          Else
              If i < tdbgView.ApproxCount - 1 Then
                  tdbgView.MoveNext
              End If    
          End If
      Next 

     SearchInPttel =  status
End Function

'---------------------------------------------------------------------------
' Ստուգում է բացված պատուհանի հաղորդագրությունը
'---------------------------------------------------------------------------
' FormType - պատուհանի տեսակ
' ExpectedMessage - սպասվող հաղորդագրություն
Function MessageExists( FormType, ExpectedMessage )
    Dim MessageWin,ActualMessage

    Select Case FormType
      Case 1 ' DeleteDoc
          Set MessagesWin = p1.WaitVBObject("frmDeleteDoc",2000)
          If MessagesWin.Exists Then 
              ActualMessage = p1.VBObject("frmDeleteDoc").VBObject("LblMsg").NativeVBObject
          Else
              Log.Warning "Message window doesn't exist!!!",,, WarningColor
              Exit Function
          End If    
      Case 2 ' Message Box 
          Set MessageWin = p1.WaitVBObject("frmAsMsgBox",2000) 
          If MessageWin.Exists Then
              ActualMessage = p1.VBObject("frmAsMsgBox").VBObject("lblMessage").NativeVBObject
          Else
              Log.Warning "Message window doesn't exist!!!",,, WarningColor
              Exit Function
          End If  
      Case 3 ' Message Box
          Set MessageWin = p1.WaitVBObject("frmAsUstPar",2000) 
          If MessageWin.Exists Then
              ActualMessage = p1.VBObject("frmAsUstPar").VBObject("LabelCntrl").NativeVBObject
          Else
              Log.Warning "Message window doesn't exist!!!",,, WarningColor
              Exit Function
          End If  
     End Select
       
      If Trim(ActualMessage) = Trim(ExpectedMessage) Then
          MessageExists = True
          Log.Message "Message is correct" ,,, MessageColor   
      Else
          MessageExists = False
          Log.Error "Message must be = " & ExpectedMessage & ",It was = " & ActualMessage ,,, ErrorColor
      End If 
End Function

'---------------------------------------------------------------------------
' Փնտրում և հեռացնում է տողը pttel-ից
'---------------------------------------------------------------------------
' PttelName - համապատասխան Pttel-ի անունը (օր.՝"frmPttel", "frmPttel_2")
' ColumnN - Թղթապանակում սյան համարը
' SearchValue - փնտրվող արժեքը
' ExpectedMessage - սպասվող հաղորդագրություն
Sub SearchAndDelete( PttelName, ColumnN, SearchValue, ExpectedMessage )
    If SearchInPttel(PttelName,ColumnN, SearchValue) Then
        Call wMainForm.MainMenu.Click(c_Opers & "|" & c_Delete)
        BuiltIn.Delay(delay_small) 
        Call MessageExists(1,ExpectedMessage)
        Call ClickCmdButton(3, "²Ûá") 
    Else
        Log.Error "Can Not find this row!",,,ErrorColor
    End If 
End Sub 

'------------------------------------------------------------------------------------
'Համեմատում է տրված երկու արժեքները 
'------------------------------------------------------------------------------------
Function Compare_Two_Values(Name,ActualValue,ExpectedValue)
				Dim isEqual : isEqual = True
    If Not Trim(ActualValue) = Trim(ExpectedValue) Then
							 isEqual = False
        Log.Error Name & " field - is NOT equal to = "& Trim(ExpectedValue) &" ,It was = "& Trim(ActualValue),,,ErrorColor
    End If
				Compare_Two_Values = isEqual
End Function

'---------------------------------------------------------------------------
' Ֆունկցիան դիտել գործողությունից հետո վերցնում է փաստաթղթի ISN-ը`
'---------------------------------------------------------------------------
Function GetIsn()
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_View)    
    If wMDIClient.WaitvbObject("frmASDocForm", 3000).Exists Then
        GetIsn = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
        Call Close_Window(wMDIClient, "frmASDocForm")
    Else
        Log.Error "Can Not Open Window to get ISN " ,,,ErrorColor    
    End If
End Function

'------------------------------------------------------------------------------------
' Ֆունկցիան Pttel - ի բացվելու դեպքում վերադաձնում է True
'------------------------------------------------------------------------------------
Function WaitForPttel(PttelName)
    Dim timeout,Exists
    timeout = 0
    Exists = False
    Do Until Exists Or timeout = 20
      If wMDIClient.WaitVBObject(PttelName, delay_middle).Exists Then
          Exists = True
          BuiltIn.Delay(2000)
      End If
      timeout = timeout + 1
    Loop
    WaitForPttel = Exists
End Function   

'---------------------------------------------------------------------------
' Pttel-ում Համեմատում է տողի colN -րդ սյան արժեքը սպասվող արժեքի հետ
'---------------------------------------------------------------------------
' PttelName - համապատասխան Pttel-ի անունը (օր.՝"frmPttel", "frmPttel_2")
' columnName - Թղթապանակում սյան անունը (օր.՝"fPenRem", "fAccRem")
' SearchValue - փնտրվող արժեքը
Function CompareFieldValue(PttelName, columnName, ExpectedValue)
    Dim column_number, actual_value
    
    BuiltIn.Delay(delay_small)
    column_number = wMDIClient.VBObject(PttelName).GetColumnIndex(columnName)
    actual_value = wMDIClient.VBObject(PttelName).VBObject("tdbgView").Columns.Item(column_number).Text
    
    If Trim(actual_value) = Trim(ExpectedValue)  Then
      CompareFieldValue = true
    else
      CompareFieldValue = false
      Log.Error column_number & " Rd field Value is NOT equal to = "& Trim(ExpectedValue) &" ,It was = "& Trim(actual_value),,,ErrorColor
    End If
End Function

'---------------------------------------------------------------------------
' Pttel-ում Համեմատում է տողի colN -րդ սյան արժեքը սպասվող արժեքի հետ
'---------------------------------------------------------------------------
' PttelName - համապատասխան Pttel-ի անունը (օր.՝"frmPttel", "frmPttel_2")
' columnName - Թղթապանակում սյան անունը (օր.՝"fPenRem", "fAccRem")
' SearchValue - փնտրվող արժեքը
Function Compare_ColumnFooterVlaue(PttelName, columnName, ExpectedValue)
    Dim column_number, actual_value
    
    BuiltIn.Delay(delay_small)
    column_number = wMDIClient.VBObject(PttelName).GetColumnIndex(columnName) 
    actual_value = wMDIClient.VBObject(PttelName).VBObject("tdbgView").Columns.Item(column_number).FooterText
    
    If Trim(actual_value) = Trim(ExpectedValue)  Then
      Compare_ColumnFooterVlaue = true
    else
      Compare_ColumnFooterVlaue = false
      Log.Error column_number & " Footer field Value is NOT equal to = "& Trim(ExpectedValue) &" ,It was = "& Trim(actual_value), "", pmNormal, ErrorColor
    End If
End Function

'---------------------------------------------------------------------------
' CheckPttel_RowCount - Համեմատում է pttel-ի տողերի քանակը
'---------------------------------------------------------------------------
Sub CheckPttel_RowCount(PttelName, ExpectedRowCount)
    Dim ActualRowCount
    
    If WaitForPttel(PttelName) Then
        ActualRowCount = wMDIClient.VBObject(PttelName).VBObject("tdbgView").ApproxCount
        
        If ActualRowCount = ExpectedRowCount Then
            Log.Message "Row Count is correct" ,,, MessageColor  
        Else 
            Log.Error "Row Count is Not equal to = "& ExpectedRowCount &" ,It Was = "& ActualRowCount  ,,,ErrorColor
        End If  
    Else 
        Log.Error PttelName & "  does not Exist!" ,,,ErrorColor
    End If  
End Sub

'---------------------------------------------------------------------------
' ExportToExcel - Արտահանում է թղթապանակի տողերը
'---------------------------------------------------------------------------
' PttelName - համապատասխան Pttel-ի անունը (օր.՝"frmPttel", "frmPttel_2")
' Path - արտահանման ճանապահը (օր.՝"Stores\ExpectedReports\PlasticCards\Actual\Actual.xlsx")
Sub ExportToExcel(PttelName,Path)
    Dim Exists,MessageWin,RowCount
    
    'Î³ï³ñáõÙ ¿ ëïáõ·áõÙ,»Ã» ÝÙ³Ý ³ÝáõÝáí ý³ÛÉ Ï³ ïñí³Í ÃÕÃ³å³Ý³ÏáõÙ ,çÝçáõÙ ¿   
    Exists = aqFile.Exists(Path)
    If Exists Then
        aqFileSystem.DeleteFile(Path)
    End If
    BuiltIn.Delay(delay_middle)

    If WaitForPttel(PttelName) Then
        
        RowCount = wMDIClient.VBObject(PttelName).VBObject("tdbgView").ApproxCount
        wMDIClient.VBObject(PttelName).Keys("^[F2]")
        
        BuiltIn.Delay(4000)
        
        If RowCount >= 20000 Then 
            Set MessageWin = p1.WaitVBObject("frmAsMsgBox", delay_small)
            If MessageWin.Exists Then
                Call MessageExists(2,"îáÕ»ñÇ ù³Ý³ÏÁ 20000-Çó ³í»É ¿, ³ñï³Ñ³ÝáõÙÁ Ïï¨Ç »ñÏ³ñ, Ð³ëï³ï»ù:")
                Call ClickCmdButton(5, "²Ûá") 
                BuiltIn.Delay(RowCount * 1.3)
            Else 
                Log.Error "frmAsMsgBox does not Exists!" ,,,ErrorColor
            End If
        End If 
        If Sys.Process("EXCEL").Exists Then
				    BuiltIn.Delay(10000)
            Sys.Process("EXCEL").Window("XLMAIN", "* - Excel", 1).Keys("[F12]")
            Sys.Process("EXCEL").Window("#32770", "Save As", 1).Keys(Path & "[Enter]")
        Else 
            Log.Error "Excel does not Open!" ,,,ErrorColor
        End If 
    Else 
        Log.Error PttelName & " does not Exist!" ,,,ErrorColor
    End If   
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''columnSorting''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''
'Պրոցեդուրան սորտավորում է պատուհանի նշված սյունը
'colName - սորտավորվող սյան անունը (անունների զանգված)
'sortColCount - սորտավորվող սյուների քանակը
'frmWin - պատուհանի տեսակը
Sub columnSorting(colName, sortColCount, frmWin)
		Dim i, colNum, RowCount
    
    RowCount = wMDIClient.VBObject("frmPttel").VBObject("tdbgView").ApproxCount
		For i = 0 to sortColCount - 1
				colNum =	wMDIClient.VBObject(frmWin).GetColumnIndex(colName(i))
				wMDIClient.VBObject(frmWin).Keys("[Hold]" & "^!" & (colNum + 1))
        BuiltIn.Delay(1000 + (RowCount * 0.2))
		Next
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''FastColumnSorting''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''
'Պրոցեդուրան սորտավորում է պատուհանի նշված սյունը
'colName - սորտավորվող սյան անունը (անունների զանգված)
'sortColCount - սորտավորվող սյուների քանակը
'frmWin - պատուհանի տեսակը
Sub FastColumnSorting(colName, sortColCount, frmWin)
		Dim i, colNum
		for i = 0 to sortColCount - 1
				colNum =	wMDIClient.VBObject(frmWin).GetColumnIndex(colName(i))
				wMDIClient.VBObject(frmWin).Keys("[Hold]" & "^!" & (colNum + 1))
    BuiltIn.Delay(700)
		next
End Sub

'--------------------------------------------------------------------------------------------
' Ֆունկցիան սպասում է այնքան մինչև "կատարման ընթացքը" վերջանա Pttel- ի բացվելու դեպքում վերադաձնում է True
'--------------------------------------------------------------------------------------------
Function WaitForExecutionProgress()
    Dim timeout,Exists
    timeout = 0
    Exists = False
    BuiltIn.Delay(1000)
    Do Until Exists Or timeout = 250
      If p1.WaitVBObject("frmPttelProgress", 2000).Exists Then
          Exists = False
          BuiltIn.Delay(2000)
      Else
          Exists = True
      End If
      timeout = timeout + 1
    Loop
    WaitForExecutionProgress = Exists
End Function 

'-------------------------------------------------------------------
' ExportToTXTFromPttel - Արտահանում է թղթապանակի տողերը բացված Pttel-ից
'-------------------------------------------------------------------
' PttelName - համապատասխան Pttel-ի անունը (օր.՝"frmPttel", "frmPttel_2")
' Path - արտահանման ճանապահը (օր.՝"Stores\ExpectedReports\PlasticCards\Actual\Actual.txt")
Sub ExportToTXTFromPttel(PttelName,Path)
    Dim exists,ViewWin,SaveAsWin
    
    'Î³ï³ñáõÙ ¿ ëïáõ·áõÙ,»Ã» ÝÙ³Ý ³ÝáõÝáí ý³ÛÉ Ï³ ïñí³Í ÃÕÃ³å³Ý³ÏáõÙ ,çÝçáõÙ ¿   
    exists = aqFile.Exists(Path)
    If exists Then
        aqFileSystem.DeleteFile(Path)
    End If
    BuiltIn.Delay(delay_middle)

    If WaitForPttel(PttelName) Then
        wMDIClient.VBObject(PttelName).Keys("^[F5]")
        BuiltIn.Delay(3000)
        
        Set ViewWin = wMDIClient.WaitVBObject("FrmSpr",500)
        If ViewWin.Exists Then
            
            'Սեղմել "Հիշել որպես"
            Call wMainForm.MainMenu.Click(c_SaveAs)
            BuiltIn.Delay(1000)
            
            Set SaveAsWin = p1.WaitWindow("#32770", "ÐÇß»É áñå»ë", 1, 500)
            If SaveAsWin.Exists Then
                p1.Window("#32770", "ÐÇß»É áñå»ë", 1).Window("DUIViewWndClassName", "", 1).Window("DirectUIHWND", "", 1).Window("FloatNotifySink", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1).Keys(Path)
                p1.Window("#32770", "ÐÇß»É áñå»ë", 1).Window("Button", "&Save", 1).Click()
            Else
                Log.Error "Save As Window does not Open!" ,,,ErrorColor
            End If
            BuiltIn.Delay(1000)
            ViewWin.Close
        Else
            Log.Error "View Window does not Open!" ,,,ErrorColor
        End If
    Else 
        Log.Error "Pttel window does not Exist!" ,,,ErrorColor
    End If   
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''SaveRAM_RowsLimit''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պրոցեդուրան դրույթների միջից փոփոխում է Տեսնել հիշողությունը սկսած(տողերի քանակ) դաշի արժեքը
'rekvValue - դաշտի նոր լրացվող արժեքը
Sub SaveRAM_RowsLimit(rekvValue)
		wMainForm.Keys("^o")
		if p1.WaitVBObject("frmOptions", 5000).Exists then
				p1.vbObject("frmOptions").VBObject("OptionsGeneral").VBObject("FrameGen").vbObject("TDBNRowsRAMLimit").Value = rekvValue
				p1.vbObject("frmOptions").VBObject("OptionsGeneral").VBObject("FrameGen").vbObject("TDBNRowsRAMLimit").Keys("[Tab]")   
				Call ClickCmdButton(6, "Î³ï³ñ»É")      
				BuiltIn.Delay(2000)
		else
				Log.Error "Can't open frmOptions window.", "", pmNormal, ErrorColor
		end if
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''Check_AgreementExisting''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պայմանագրեր թղթապանակում փաստատթղթի առկայության ստուգում 
'Ֆունկցիան վերադարձնում է true, եթե պայմանագիրը առկա է և false, եթե այն բացակայում է 
'agreement - Պայմանագրեր մուտք գործելու կլասի օբյեկտ
'FolderName - Պայմանագրի ճանապարհը
Function Check_AgreementExisting(FolderName, agreement)
  Dim isExist : isExist = true
  Call GoTo_Contracts(FolderName, agreement)
  wMDIClient.Refresh
  If wMDIClient.vbObject("frmPttel").vbObject("tdbgView").ApproxCount <> 1 Then
						Log.Error "There are no document with specified ID or there are more than one. There are " &_
						wMDIClient.vbObject("frmPttel").vbObject("tdbgView").ApproxCount & "rows.", "", pmNormal, ErrorColor
      isExist = false
  End If
  Check_AgreementExisting = isExist
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''-- DeleteAllActions --''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Ֆունկցիան հեռացնում է պայմանագրի բոլոր գործողությունները
Sub DeleteAllActions(FolderPath,DocNum,StartDate,EndDate)
    Dim Pttel, i, tdbgView, MessageWin
    
    Call wTreeView.DblClickItem(FolderPath & "|¶áñÍáÕáõÃÛáõÝÝ»ñ, ÷á÷áËáõÃÛáõÝÝ»ñ|ä³ÛÙ³Ý³·ñÇ µáÉáñ ·áñÍáÕáõÃÛáõÝÝ»ñ")
    wMDIClient.Refresh
    
    BuiltIn.Delay(1000)
    Call Rekvizit_Fill("Dialog", 1, "General", "START", "^A[Del]" & StartDate)
    Call Rekvizit_Fill("Dialog", 1, "General", "END", "^A[Del]" & EndDate)
    Call Rekvizit_Fill("Dialog", 1, "General", "NUM", DocNum)
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    BuiltIn.Delay(5000)
    
    Set Pttel = wMDIClient.VBObject("frmPttel")
    If WaitForPttel("frmPttel") Then
      Set tdbgView = wMDIClient.VBObject("frmPttel").VBObject("tdbgView")
      BuiltIn.Delay(3000)
      tdbgView.MoveLast
      
      For i = 1 to tdbgView.ApproxCount
          If tdbgView.ApproxCount = 0 Then
              Exit For
          Else
              Call wMainForm.MainMenu.Click(c_Opers & "|" & c_Delete)
              BuiltIn.Delay(2000) 
              Set MessageWin = p1.WaitVBObject("frmAsMsgBox", delay_small)
              If MessageWin.Exists Then
                  Call MessageExists(2,"àõß³¹ñáõÃÛáõÝ:"&vbCrLf&"¸áõù ÷áñÓáõÙ »ù Ñ»é³óÝ»É §ËÙµ³ÛÇÝ ·áñÍáÕáõÃÛáõÝÝ»ñÇ¦ ÷³ëï³ÃáõÕÃÁ")
                  Call ClickCmdButton(5, "²Ûá") 
              End If 
              
              BuiltIn.Delay(1000) 
              Call ClickCmdButton(3, "²Ûá") 
              BuiltIn.Delay(1000) 
              If tdbgView.ApproxCount = 0 Then
                  Exit For
              End If
              wMainForm.Keys("^r")
              BuiltIn.Delay(3000) 
              tdbgView.MoveLast
          End If   
      Next 
        BuiltIn.Delay(1000) 
        Call Close_Pttel("frmPttel") 
    Else
        Log.Error "Can Not Open Պայմանագրի բոլոր գործողությունները Window",,,ErrorColor         
    End If 
    If Pttel.Exists Then
        Log.Error "Can Not Close Պայմանագրի բոլոր գործողությունները Window",,,ErrorColor
    End If 
End Sub

'Excel ý³ÛÉÇ å³Ñå³ÝÙ³Ý ýáõÝÏóÇ³
'fileName - ³ñï³Ñ³ÝíáÕ ý³ÛÉÇ å³Ñå³ÝÙ³Ý ×³Ý³å³ñÑ
Sub SaveExcelFile(fileName)
		Dim confSaveAs
		if Sys.WaitProcess("EXCEL", 15000).Exists then
    Sys.Process("EXCEL").Window("XLMAIN", "* - Excel", 1).Keys("[F12]")
    Sys.Process("EXCEL").Window("#32770", "Save As", 1).Keys(fileName & "[Enter]")
				if Sys.WaitProcess("EXCEL", 3000).Window("#32770", "Confirm Save As", 1).Exists then
				  Set confSaveAs = Sys.Process("EXCEL").Window("#32770", "Confirm Save As", 1)
						confSaveAs.UIAObject("Confirm_Save_As").Window("CtrlNotifySink", "", 7).Window("Button", "&Yes", 1).Click
				end if
  else 
      Log.Error "Excel does not Open!" , "",pmNormal, ErrorColor
  end if 
End	Sub

'--------------------------------------------------------------------------------
'Մուտք է գործում Ադմինիստրատորի ԱՇՏ 4.0|Պարամետրեր|Համակարգի պարամետրերի ուղղորդիչ թղթապանակ
'--------------------------------------------------------------------------------
Sub GoTo_SystemParameters(SystemParams) 
    Dim FilterWin
    
    Call wTreeView.DblClickItem("|²¹ÙÇÝÇëïñ³ïáñÇ ²Þî 4.0|ä³ñ³Ù»ïñ»ñ|Ð³Ù³Ï³ñ·Ç å³ñ³Ù»ïñ»ñÇ áõÕÕáñ¹Çã")
    Set FilterWin = p1.WaitVBObject("frmAsUstPar",delay_middle)
    BuiltIn.Delay(delay_middle) 
    
    If FilterWin.Exists Then
        'Լրացնում է "Շաբլոն" դաշտը
        p1.VBObject("frmAsUstPar").VBObject("TabFrame").VBObject("TDBMask").keys(SystemParams)
        Call ClickCmdButton(2, "Î³ï³ñ»É")
        
        Call WaitForExecutionProgress()
    Else
        Log.Error "Can Not Open System Parameters Filter",,,ErrorColor      
    End If 
End Sub

'-------------------------------------------------
' Կատարել "Խմբագրել ռեեստրի կարգավիճակը"  Գործողություն
'------------------------------------------------
Sub EditRegisterStatus(Status,Comment)
    Dim DocForm
    
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_EditRegisterStatus)
    BuiltIn.Delay(1500)
    Set DocForm = AsBank.WaitVBObject("frmAsUstPar", 2000)
    
    If DocForm.Exists Then
        'Լրացնել "Ռեեստրի Կարգավիճակ" դաշտը
        Call Rekvizit_Fill("Dialog", 1, "General", "REPSTATUS", Status)
        'Լրացնել "Մեկնաբանություն" դաշտը
        Call Rekvizit_Fill("Dialog", 1, "General", "COMM", Comment)
        
        Call ClickCmdButton(2, "Î³ï³ñ»É")
    Else
        Log.Error "Can Not Open Edit Register Status Window",,,ErrorColor         
    End If    
    BuiltIn.Delay(1500)
    If DocForm.Exists Then
        Log.Error "Can Not Close Edit Register Status Window",,,ErrorColor
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''Փակել պատուհանը'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''
' parentPath - պատուհանին հասնելու ճանապարհը
' windowName - պատուհանի անունը
Sub Close_Window(parentPath, windowName)
				If parentPath.WaitVBObject(windowName, 3000).Exists Then
								parentPath.VBObject(windowName).Close
				Else
								Log.Error "Can't close window", "", pmNormal, ErrorColor    
				End If  
End Sub

'----------------------------------------------------------------------------
'--------------------------Ֆայլը փաստաթղթին կցելու ֆունկցիա------------------------
'----------------------------------------------------------------------------
'Filepath-ֆայլի ուղղությունը օր. Filepath = Project.Path & "Stores\MemorialOrder\ForTest.xlsx"
'TabN- "Կցված" Tab-ի համարը
'Filename - ֆայլի անունը օր՝ ForTest.xlsx
Sub Attach_File_ToDoc (filePath, tabN, fileName)
    Dim sTab, wTabStrip, count, path, listViewAttachments, item, expMessage
    
    sTab = "TabFrame"
    If wMDIClient.WaitVBObject("frmASDocForm",1000).Exists Then
        'Անցում համպատասխան Tab
        If tabN <> 1 Then
    								sTab = sTab & "_" & tabN
    								Set wTabStrip = wMDIClient.VBObject("frmASDocForm").VBObject("TabStrip")
    								wTabStrip.SelectedItem = wTabStrip.Tabs(tabN)
        End If
        Set listViewAttachments = wMDIClient.VBObject("frmASDocForm").VBObject(sTab).VBObject("AsAttachments1").VBObject("ListViewAttachments")
        count = listViewAttachments.wItemCount  'Մինչև կցելը առկա ֆայլերի քանակը      
        wMDIClient.VBObject("frmASDocForm").VBObject(sTab).VBObject("AsAttachments1").VBObject("CmdAdd").Click
        BuiltIn.Delay (3000)   
        If asbank.Window("#32770", "Open", 1).exists Then
           asbank.Window("#32770", "Open", 1).Window("ComboBoxEx32", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1).Keys(filePath)
           asbank.Window("#32770", "Open", 1).Window("Button", "&Open", 1).click       
        End If
        BuiltIn.Delay (4000)
        
        Set messageWin = asbank.WaitVBObject("frmAsMsgBox",1000) 
        'Ստուգում է արդյոք ֆայլի կրկնվելու վերաբերյալ հաղորդագրություն կա թե ոչ
        If messageWin.Exists Then
            expMessage = fileName & " - ²Û¹åÇëÇ ³ÝáõÝáí ý³ÛÉ ³ñ¹»Ý Ï³, Ã³ñÙ³óÝ»±É:"
            actMessage = asbank.VBObject("frmAsMsgBox").VBObject("lblMessage").NativeVBObject
            'Հաղորդագրությունում հնարավոր տողադարձերի փոփոխում բացատի
            Set reg = New RegExp      
            reg.Global = True
            reg.Pattern = "\s"
            actMessageNew = reg.Replace (actMessage , " ")
            actMessageNew = Replace(actMessageNew,"   ", " ")
            'Ակտուալ և սպասվող հաղորդագրությունների համեմատում
            If Trim(actMessageNew) = Trim(expMessage) Then
               Call ClickCmdButton(5, "²Ûá")
               Log.Message "File " & filePath & " has been updated",,,MessageColor 
            Else
                Log. Error "Expected message is " & expMessage & " Actual message is " & actMessageNew ,,, ErrorColor 
            End If
        Else
            'Ստուգում է արդյոք ֆայլը ավելացավ ցուցակում թե ոչ
            If listViewAttachments.wItemCount=count+1 Then
               Log.Message "File " & filePath & " has been attached",,,MessageColor
            Else
                Log.Error "File" & filePath & " has not been attached",,,ErrorColor
            End If  
        End If
    End If
End Sub

'-------------------------------------------------------------
'listViewAttachments տիպի աղյուսակում Ֆայլը կամ հղումը գտնելու ֆունկցիա
'-------------------------------------------------------------
'filename-ֆայլի անունը կամ հղումը օրինակ՝ ForTest.xslx կամ D:\Testing\TestsAsBank\AsBank\Stores\MemorialOrder\ForTest.xslx
'tabN- "Կցված" Tab-ի համարը 
'Վերադարձնում է true եթե ֆայլը գտնվել է
Function SearchInAttachList (fileName,tabN)
    Dim i, listViewAttachments, status, sTab, wTabStrip
    sTab = "TabFrame"
    If wMDIClient.WaitVBObject("frmASDocForm",3000).Exists Then
        'Անցում համպատասխան Tab
        If tabN <> 1 Then
        sTab = sTab & "_" & tabN
        Set wTabStrip = wMDIClient.VBObject("frmASDocForm").VBObject("TabStrip")
        wTabStrip.SelectedItem = wTabStrip.Tabs(tabN)
        End If
    End If 
    status = False
    Set listViewAttachments = wMDIClient.VBObject("frmASDocForm").VBObject(sTab).VBObject("AsAttachments1").VBObject("ListViewAttachments")
    listViewAttachments.Click
    listViewAttachments.Keys("[Home]")
    For i = 0 to listViewAttachments.wItemCount - 1
        If Trim(listViewAttachments.wItem(i,1)) = Trim(fileName)  Then
           status = True   
           Exit For
        Else
            If i < listViewAttachments.wItemCount - 1 Then
               listViewAttachments.Keys("[Down]")
            End If
        End If
    Next
    If status = False Then
       Log.Error "File " & fileName & " not found",,,ErrorColor 
    End If
    SearchInAttachList = status
End Function

'-------------------------------------------------------------
'listViewAttachments տիպի աղյուսակում Ֆայլը կամ հղումը ջնջելու ֆունկցիա
'-------------------------------------------------------------
'filename - ֆայլը/հղումը օրինակ՝ ForTest.xslx կամ D:\Testing\TestsAsBank\AsBank\Stores\MemorialOrder\ForTest.xslx
'tabN- "Կցված" Tab-ի համարը 
Sub Delete_Attached_File_Doc (fileName,tabN) 
    If SearchInAttachList(fileName,tabN) Then
        Dim sTab, wTabStrip, count, reg, reg1, expMessage, actMessageNew , actMessage, messageWin
        sTab = "TabFrame"
        If tabN <> 1 Then
           sTab = sTab & "_" & tabN
        End If
        count = wMDIClient.VBObject("frmASDocForm").VBObject(sTab).VBObject("AsAttachments1").VBObject("ListViewAttachments").wItemCount
        If wMDIClient.VBObject("frmASDocForm").VBObject(sTab).VBObject("AsAttachments1").VBObject("CmdDelete").Enabled Then
            wMDIClient.VBObject("frmASDocForm").VBObject(sTab).VBObject("AsAttachments1").VBObject("CmdDelete").Click
            BuiltIn.delay(1000)
            Set messageWin = p1.WaitVBObject("frmAsMsgBox",2000) 
            If messageWin.Exists Then
                expMessage = "æÝç»±É " & fileName & " Ïóí³Í ý³ÛÉÁ/ÑÕáõÙÁ:"
                actMessage = asbank.VBObject("frmAsMsgBox").VBObject("lblMessage").NativeVBObject
                'Հաղորդագրությունում հնարավոր տողադարձերի փոփոխում բացատի
                Set reg = New RegExp      
                reg.Global = True
                reg.Pattern = "\s"
                actMessageNew = reg.Replace (actMessage , " ")
                actMessageNew = Replace(actMessageNew,"   ", " ")
                If Trim(actMessageNew) = Trim(expMessage) Then
                   Call ClickCmdButton(5, "²Ûá")
                Else
                    Log. Error "Expected message is " & expMessage & " Actual message is " & actMessageNew ,,, ErrorColor 
                End If
                'Ստուգում է արդյոք ֆայլը ջնջվեց ցուցակից թե ոչ
                If wMDIClient.VBObject("frmASDocForm").VBObject(sTab).VBObject("AsAttachments1").VBObject("ListViewAttachments").wItemCount=count-1 Then
                   Log.Message "File " & fileName & " has been deleted",,,MessageColor
                Else
                    Log.Message "File " & filename & " has not been deleted",,,ErrorColor
                End If 
            Else 
                Log. Error "Message doesn't exists",,,ErrorColor
            End If    
        Else 
            Log.Message "Delete Button is disabled",,,MessageColor
        End If    
    Else 
        Log.Error "File " & filename & " not found",,,ErrorColor
    End If
End Sub

'-----------------------------------
'Ֆայլի հղումը փաստաթղթին կցելու ֆունկցիա
'----------------------------------
'Filepath- Ֆայլի հղումը օր. Filepath = Project.Path & "Stores\MemorialOrder\ForTest.xlsx"
'TabN- "Կցված" Tab-ի համարը
'description - Նկարագրություն
Sub Attach_Link_ToDoc (filePath, TabN, description)
    Dim sTab, wTabStrip, count, expectedMessage, strLength, messageWin, actMessage
    sTab = "TabFrame"
    If wMDIClient.WaitVBObject("frmASDocForm",1000).Exists Then
        'Անցում համպատասխան Tab
        If tabN <> 1 Then
    								sTab = sTab & "_" & tabN
    								Set wTabStrip = wMDIClient.VBObject("frmASDocForm").VBObject("TabStrip")
    								wTabStrip.SelectedItem = wTabStrip.Tabs(tabN)
        End If
        count = wMDIClient.VBObject("frmASDocForm").VBObject(sTab).VBObject("AsAttachments1").VBObject("ListViewAttachments").wItemCount
        wMDIClient.VBObject("frmASDocForm").VBObject(sTab).VBObject("AsAttachments1").VBObject("CmdAddLink").Click
        BuiltIn.Delay (3000)   
        If asbank.waitVBObject("frmAsUstPar",1000).exists Then
           'Լրացնում է հղում և նկարագրություն դաշտերը
           Call Rekvizit_Fill ("Dialog",1,"General","LINK",filePath) 
           Call Rekvizit_Fill ("Dialog",1,"General","COMM",description)
           Call ClickCmdButton (2,"Î³ï³ñ»É")    
        End If
        BuiltIn.Delay (5000)
        strLength = Len(filePath)
        'Ստուգում է ֆայլի հղման սիմվոլների քանակը և եթե այն 78 -ից մեծ է կատարում է տողադարձ 78-րդ սիմվոլից հետո հղումը կրկնվելու վերաբերյալ հաղորդագրությունը ստուգելու համար 
        If strLength <= 78 Then
           expectedMessage = filePath & vbNewLine & "²Û¹åÇëÇ ý³ÛÉ/ÑÕáõÙ ³ñ¹»Ý Ï³"
        Else 
            expectedMessage = Left(filePath , 78) & vbNewLine & Mid (filePath , 79) & vbNewLine & "²Û¹åÇëÇ ý³ÛÉ/ÑÕáõÙ ³ñ¹»Ý Ï³"
        End If    
        Set messageWin = asbank.WaitVBObject("frmAsMsgBox",1000) 
        'Ստուգում է արդյոք հղումը կրկնվելու վերաբերյալ հաղորդագրություն կա թե ոչ
        If messageWin.Exists Then
           actMessage = messageWin.VBObject("lblMessage").NativeVBObject
           If Trim(actMessage) = Trim(expectedMessage) Then
              Call ClickCmdButton(5 , "OK")
              Log.Message "Link " & filePath & " already exists",,,DivideColor2
           Else
               Log.Error "Expected message is " & expectedMessage & " Actual message is " & actMessage ,,, ErrorColor
           End If
        Else
            'Ստուգում է արդյոք հղումը ավելացավ ցուցակում թե ոչ
            If wMDIClient.VBObject("frmASDocForm").VBObject(sTab).VBObject("AsAttachments1").VBObject("ListViewAttachments").wItemCount=count+1 Then
               Log.Message "Link " & filepath & " has been attached",,,MessageColor
            Else
                Log.Message "Link " & filepath & " has not been attached",,,ErrorColor
            End If
        End If  
    End If
End Sub

'Կցել գործողության կլասս
Class Attached_Tab
    Public tabN
    Public filesCount
    Public linksCount
    Public delRowsCount
    Public addFiles()
    Public fileName()
    Public addLinks()
    Public delFiles()
    Public linkName()
    Private Sub Class_Initialize()
        Dim i
        tabN = 4
        filesCount = files_count
        linksCount = links_count
        delRowsCount = del_count
        Redim addFiles(filesCount)
        Redim fileName(filesCount)
        For i = 0 To filesCount - 1
            addFiles(i) = ""
            fileName(i) = ""
        Next
        Redim addLinks(linksCount)
        Redim linkName(linksCount)
        For i = 0 To linksCount - 1
            addLinks(i) = ""
            linkName(i) = ""
        Next
        Redim delFiles(delRowsCount)
        For i = 0 To delRowsCount - 1
            delFiles(i) = ""
        Next
    End Sub
End Class

Function New_Attached_Tab(fCount, lCount, dCount)
    files_count = fCount
    links_count = lCount
    del_count = dCount
    Set New_Attached_Tab = New Attached_Tab
End Function

'Ֆայլեր և հղումներ կցող և ջնջող ֆունկցիա
Sub Fill_Attached_Tab(Attached)
    Dim i
    For i = 0 To Attached.filesCount - 1
        Call Attach_File_ToDoc (Attached.addFiles(i), Attached.tabN, Attached.fileName(i))
    Next
    For i = 0 To Attached.linksCount - 1
        Call Attach_Link_ToDoc (Attached.addLinks(i), Attached.TabN, Attached.linkName(i))
    Next
    For i = 0 To Attached.delRowsCount - 1
        Call Delete_Attached_File_Doc (Attached.delFiles(i), Attached.tabN) 
    Next
End Sub

'"Կցված" էջը ստուգող ֆունկցիա
Sub Attach_Tab_Check(tabAttach)
    Dim i, expCount, count
    'Անցում Կցված էջ 
    Call GoTo_ChoosedTab(tabAttach.tabN)   
    'Ստուգել, որ ֆայլերը առկա են 
    For i = 0 To tabAttach.filesCount - 1
        Call SearchInAttachList (tabAttach.fileName(i), tabAttach.tabN) 
    Next
    'Ստուգել, որ հղումները առկա են
    For i = 0 To tabAttach.linksCount - 1
        Call SearchInAttachList (tabAttach.addLinks(i), tabAttach.tabN)
    Next
    'Համեմատում է Աղյուսակում առկա և ակնկալվող ֆայլերի և հղումների քանակը,առկա տողերի սխալ քանակի դեպքում լոգավորում է Error
    count = wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame_"&tabAttach.tabN).VBObject("AsAttachments1").VBObject("ListViewAttachments").wItemCount
    expCount = tabAttach.filesCount + tabAttach.linksCount
    If expCount <> count Then
       Log.Error "Attached files and links count is " & count & ". Expected value is " & expCount ,,, ErrorColor
    End If 
End Sub

'------------------------------------------------------------------------------------
'-------------Grid - ի սյուների խմբագրվող լինելը ստուգող ֆունկցիա-----------------------------
'------------------------------------------------------------------------------------
'columnCount-Սյուների քանակը
'tabN- էջը որի վրա գտնվում է Grid-ը
'colReadOnlyArray- Grid-ի սյուների չխմբագրվող լինելու զանգված օր.՝ Array(True, False, True)
'Ֆունկցիան վերադարձնում է True կամ False արժեք
Function Check_ReadOnly_Grid (columnCount, tabN,  colReadOnlyArray)
    Dim grid, sTab, i, j, status, statusEnd
    statusEnd = True
    sTab = "TabFrame"
    If tabN <> 1 Then
        sTab = sTab & "_" & tabN
    End If
    Set grid = wMDIClient.VBObject("frmASDocForm").VBObject(sTab).VBObject("DocGrid")
    For i = 0 to columnCount - 1
        For j = 0 to grid.ApproxCount 
            With grid
                 .row = j
                 status = .Columns.Item(i).Locked
            End With
            If not status = colReadOnlyArray (i) Then
               statusEnd = False
               Log.Error "( " & j & ", " & i & ") ReadOnly value is not " & colReadOnlyArray(i) & ". It is " & status ,,,ErrorColor
            End If   
        Next    
    Next
    Check_ReadOnly_Grid = statusEnd
End Function

'------------------------------------------------------------------------------------
'-------------Grid -ի բջիջների արժեքները ստուգող ֆունկցիա-----------------------------------
'------------------------------------------------------------------------------------
'columnN - սյան համարը
'rowN - տողի համարը
'docType - պատուհանի տեսակը, որտեղ գտնվում է գրիդը
'tabN - էջը որի վրա գտնվում է Grid-ը
'expectedValue - սպասվող արժեք
'Ֆունկցիան ստանում է True կամ False արժեք
Function Check_Value_Grid (columnN, rowN, docType, tabN, expectedValue)
    Dim grid, sTab, actValue, status
    status = True
    
    Select Case docType
        Case "Document"
            sTab = "TabFrame"
            If tabN <> 1 Then
                sTab = sTab & "_" & tabN
            End If
            Set grid = wMDIClient.VBObject("frmASDocForm").VBObject(sTab).VBObject("DocGrid")
        Case "OLAP"
            Set grid = wMDIClient.VBObject("frmOLAPExp").VBObject("TDBGrid1")
    End Select
    
    With grid
         .row = rowN
         actValue = .Columns.Item(columnN).Text
    End With
    If not Trim(actValue) = Trim(expectedValue) Then
       status = False
       Log.Error "(" & rowN & "," & columnN & ")  Value is not " & expectedValue & ". It is " & actValue ,,,ErrorColor
    End If
    Check_Value_Grid = status           
End Function

'------------------------------------------------------------------------------------
'-------------Grid -ի բջիջի տողը դուրս բերող ֆունկցիա--------------------------------------
'------------------------------------------------------------------------------------
'columnN - սյան համարը
'docType - պատուհանի տեսակը, որտեղ գտնվում է գրիդը
'tabN - էջը որի վրա գտնվում է Grid-ը
'cellValue - սպասվող արժեք
'Ֆունկցիան ստանում է տողի համարը
Function Get_Cell_Row_Grid (columnN, docType, tabN, cellValue)
    Dim docGrid, sTab
    Get_Cell_Row_Grid = Null
    
    Select Case docType
        Case "Document"
            sTab = "TabFrame"
            If tabN <> 1 Then
                sTab = sTab & "_" & tabN
            End If
            Set docGrid = wMDIClient.VBObject("frmASDocForm").VBObject(sTab).VBObject("DocGrid")
        Case "OLAP"
            Set docGrid = wMDIClient.VBObject("frmOLAPExp").VBObject("TDBGrid1")
    End Select
    BuiltIn.Delay (1000)
    docGrid.MoveFirst
    Do Until docGrid.EOF
       If Trim (docGrid.Columns.Item(columnN).text) = Trim(cellValue) Then
           Get_Cell_Row_Grid = docGrid.row
           Exit Function
        Else
            docGrid.MoveNext
        End If 
    Loop
    If Get_Cell_Row_Grid= Null  Then
       Log.Error "Cell with value " & cellValue & " in column " & columnN & " doesn't exists",,,ErrorColor                  
    End If
End Function


' Փնտրել հանգույց/էլեմենտ ծառում գործողություն
Function Find_Tree_Element(count, arr)
        Dim i, status
        status = True
        For  i = 0 To count - 1
            While Trim(wMDIClient.VBObject("frmEditTree").VBObject("TreeView").SelectedItem) <> Trim(arr(i))
                   wMDIClient.VBObject("frmEditTree").VBObject("TreeView").Keys("[Down]")
            Wend
             If Trim(wMDIClient.VBObject("frmEditTree").VBObject("TreeView").SelectedItem) = Trim(arr(i)) Then
                wMDIClient.VBObject("frmEditTree").VBObject("TreeView").Keys("[Enter]")
                If i <>  count - 1 Then
                   wMDIClient.VBObject("frmEditTree").VBObject("TreeView").Keys("[Down]")
                End If
            Else                                                      
                status = False
            End If
        Next
        Find_Tree_Element = status
End Function

'Կարգավորումները փաստաթղթով ներմուծող ֆունկցիա
Sub Settings_Import(filePath,folderDirect) 
    Call wTreeView.DblClickItem(folderDirect)
    If wMDIClient.WaitVBObject("frmImport",2000).Exists Then
        wMDIClient.VBObject("frmImport").VBObject("PathOlap").Keys("^A[Del]" & filePath & "[Enter]")
        Call ClickCmdButton(13, "Üß»É µáÉáñÁ")
        Call ClickCmdButton(13, "Î³ï³ñ»É")
        If MessageExists(2, "îíÛ³ÉÝ»ñÇ Ý»ñÙáõÍáõÙÁ ³Ýó³í µ³ñ»Ñ³çáÕ:") Then
            Call ClickCmdButton(5, "OK")
        End If    
        Call Close_Window(wMDIClient, "frmImport")
    Else
        Log.Error "Տվյալների ներմուծման պատուհանը չի բացվել"
    End If              
End Sub

'Աղյուսակների ֆիլտրման կլասս
'Ֆիլտրերի քանակը 
Class filter_Pttel
    Public andOr 
    Public colName 
    Public cond 
    Public val 
    Public valEnd 
    Public condCount
      
    Private Sub Class_Initialize
        fCount = fil_Count
        ReDim andOr (fCount)
        ReDim colName (fCount)
        ReDim cond (fCount)
        ReDim val (fCount)
        ReDim valEnd (fCount)
        For condCount = 0 to fCount
            andOr(condCount) = 0
            colName(condCount) = 0
            cond(condCount) = 0
            val(condCount) = ""
            valEnd (condCount) = ""
        Next
        condCount = fCount
    End Sub
End Class

'filterCount-Ֆիլտրի պայմանների քանակը 
Function New_Filter_Pttel (filterCount)
    fil_Count = filterCount
    Set New_Filter_Pttel = new filter_Pttel
End Function

'--------------------------------------------------
'------------Աղյուսակները Ֆիլտրող ֆունկցիա--------------
'---------------Ctrl+H գործողությամբ------------------
Sub Pttel_Filtering (filter_Pttel, pttelName)    
    Dim i, j, TDBGridFilter, k
    If wMDIClient.WaitVBObject(pttelName, 30000). exists Then
        wMDIClient.VBObject(pttelName).Keys ("^h")
        If asbank.waitVBObject("frmPttelFilter",1000).exists Then
           Do 
             asbank.VBObject("frmPttelFilter").VBObject("FilterControl").VBObject("ToolbarFilterActions").Window("msvb_lib_toolbar", "", 1).ClickItem(2)     
           Loop Until asbank.VBObject("frmPttelFilter").VBObject("FilterControl").VBObject("TDBGridFilter").ApproxCount = 0
           For i = 0 to filter_Pttel.condCount - 1
             asbank.VBObject("frmPttelFilter").VBObject("FilterControl").VBObject("ToolbarFilterActions").Window("msvb_lib_toolbar", "", 1).ClickItem(1)
           Next 
           Set TDBGridFilter = asbank.VBObject("frmPttelFilter").VBObject("FilterControl").VBObject("TDBGridFilter")
           For j = 0 to filter_Pttel.condCount - 1  
              TDBGridFilter.Keys("[Home]")
              'Եվ/ կամ սյան լրացում
              If j <> 0 Then
                 With TDBGridFilter 
                      .row = j
                      .col = 1
                      .Keys ("~[Down]")
                      k = 0
                      Do Until k = filter_Pttel.andOr(j)
                         .Keys ("[Down]")
                         k = k+1
                      Loop   
                      .Keys ("[Enter]")
                 End With  
                 TDBGridFilter.Keys ("[Right]")
              End If
             'Սյան անվանում սյան լրացում
              With TDBGridFilter
                      .row = j
                      .col = 2
                      .Keys ("~[Down]")
                      k = 0
                      Do Until k = filter_Pttel.colName(j)
                         .Keys ("[Down]")
                         k = k+1
                      Loop   
                      .Keys ("[Enter]")
              End With
              'Պայման սյան լրացում
              With TDBGridFilter
                      .row = j
                      .col = 3
                      .Keys ("~[Down]")
                      k = 0
                      Do Until k = filter_Pttel.cond(j) 
                         .Keys ("[Down]")
                         k = k+1
                      Loop   
                      .Keys ("[Enter]")
              End With
              'Արժեք սյան լրացում 
              With TDBGridFilter
                      .row = j
                      .col = 4
                      .Keys (filter_Pttel.val(j) & "[Right]")
              End With
              'Միջև պայմանի դեպքում երկրորդ Արժեք սյան լրացում
              If filter_Pttel.cond(j) = 8 Then
                  With TDBGridFilter
                          .row = j
                          .col = 5
                          .Keys (filter_Pttel.valEnd(j) & "[Right]")
                  End With
              End If   
           Next
           With TDBGridFilter
               .MoveLast
             If Trim(.Columns.Item(4).Value) = "" or .Columns.Item(4).Value = "  /  /  "  Then
                asbank.VBObject("frmPttelFilter").VBObject("FilterControl").VBObject("ToolbarFilterActions").Window("msvb_lib_toolbar", "", 1).ClickItem(2)
             End If  
           End With
           Call ClickCmdButton (7, "Î³ï³ñ»É")
        Else 
            Log.Error "Filter window doesn't exists",,,ErrorColor
        End If
    Else
        Log.Error "Pttel not found",,,ErrorColor
    End If           
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''Find_Word''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''
' objWhereSearch - օբյեկտի ճանապարհը, որում փնտրելու ենք
' searchedValue - փնտրվող արժեքը
Function Find_Word(objWhereSearch, searchedValue)
    objWhereSearch.Keys("[Home]")
    objWhereSearch.Keys("^f")
    If p1.WaitVBObject("frmSprFind", 1000).Exists Then
        ' Փնտրվող արժեքի ներմուծում
        p1.VBObject("frmSprFind").VBObject("Frame1").VBObject("TDBMask1").Keys("[F2]" & searchedValue)
        ' Սեղմել "Հաջորդը" կոճակը
        Call ClickCmdButton(12, "Ð³çáñ¹Á")
        If p1.WaitVBObject("frmAsMsgBox", 2000).Exists Then
            Call MessageExists(2, "²Û¹åÇëÇ ·ñ³éáõÙ ãÏ³")
            Call ClickCmdButton(5, "OK")
            Find_Word = False
        Else
            Find_Word = True
        End If
        Call Close_Window(p1, "frmSprFind")
    Else
        Log.Error "Can't open frmSprFind window.", "", pmNormal, ErrorColor
    End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''Fill_Grid_Field''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Գրիդի դաշտի լրացում
' colN - սյան համարը
' rowN - տողի համարը
' docType - պատուհանի տեսակը, որտեղ գտնվում է գրիդը
' fieldType - դաշտի տեսակը
' tabN - էջը որտեղ գտնվում է Grid-ը
' value - լրացվող արժեք
Sub Fill_Grid_Field(colN, rowN, docType, fieldType, tabN, value)
    Dim grid, sTab
    
    Select Case docType
        Case "Document"
            sTab = "TabFrame"
            If tabN <> 1 Then
                sTab = sTab & "_" & tabN
            End If
            Set grid = wMDIClient.VBObject("frmASDocForm").VBObject(sTab).VBObject("DocGrid")
        Case "OLAP"
            Set grid = wMDIClient.VBObject("frmOLAPExp").VBObject("TDBGrid1")
    End Select
    With grid
        .Col = colN
        .Row = rowN
        Select Case fieldType
        Case "General"
            If docType = "OLAP" Then
                .Keys(value)
                .Keys("[Right]" & "[Left]")
            End If
            .Keys(value & "[Right]")
        Case "CheckBox"
            If docType = "OLAP" Then
                If Abs(.Text) <> Abs(value) Then
                    .Keys(" ")
                End If
            End If
        End Select
    End With
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''Get_Grid_Value'''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Գրիդի դաշտի արժեքի ստացման ֆունկցիա
' colN - սյան համարը
' rowN - տողի համարը
' docType - պատուհանի տեսակը, որտեղ գտնվում է գրիդը
Function Get_Grid_Value(colN, rowN, docType)
    Dim grid, sTab
    
    Select Case docType
        Case "Document"
            sTab = "TabFrame"
            If tabN <> 1 Then
                sTab = sTab & "_" & tabN
            End If
            Set grid = wMDIClient.VBObject("frmASDocForm").VBObject(sTab).VBObject("DocGrid")
        Case "OLAP"
            Set grid = wMDIClient.VBObject("frmOLAPExp").VBObject("TDBGrid1")
    End Select
    With grid
        .Col = colN
        .Row = rowN
        Get_Grid_Value = grid.Text
    End With
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''Search_In_Grid''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Գրիդի արժեքի փնտրման ֆունկցիա
' colN - սյան համարը 
' docType - պատուհանի տեսակը, որտեղ գտնվում է գրիդը
' tabN - էջը որտեղ գտնվում է Grid-ը
' searchedValue - փնտրվող արժեք
Function Search_In_Grid(colN, docType, tabN, searchedValue)
    Dim grid, sTab, i, isExist
    
    isExist = False
    Select Case docType
        Case "Document"
            sTab = "TabFrame"
            If tabN <> 1 Then
                sTab = sTab & "_" & tabN
            End If
            Set grid = wMDIClient.VBObject("frmASDocForm").VBObject(sTab).VBObject("DocGrid")
        Case "OLAP"
            Set grid = wMDIClient.VBObject("frmOLAPExp").VBObject("TDBGrid1")
        Case "OLAPNavigator"
            Set grid = wMDIClient.VBObject("frmOLAPNav").VBObject("TDBGSections")
    End Select
    With grid
        .Col = colN
        For i = 1 To .ApproxCount 
            .Row = i 
            If .Columns.Item(colN).Text = searchedValue Then
                isExist = True
                Exit For
            End If
        Next
    End With
    
    Search_In_Grid = isExist
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''Delete_Grid_Row'''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Գրիդի տողի ջնջում
' colN - սյան համարը 
' docType - պատուհանի տեսակը, որտեղ գտնվում է գրիդը
' tabN - էջը որտեղ գտնվում է Grid-ը
' searchedValue - ջնջվող տողի դաշտի արժեքը ըստ սյան 
Sub Delete_Grid_Row(colN, docType, tabN, searchedValue)
    Dim grid, sTab
    
    sTab = "TabFrame"
    If tabN <> 1 Then
        sTab = sTab & "_" & tabN
    End If
    If Search_In_Grid(colN, docType, tabN, searchedValue) Then
        Set grid = wMDIClient.VBObject("frmASDocForm").VBObject(sTab).VBObject("DocGrid")
        grid.Keys("^d")
    Else
        Log.Error "Can't find " & searchedValue & "value in " & colN & "column", "", pmNormal, ErrorColor
    End If
    
End Sub

'Կատարում է Դիտել փաստաթուղթը գործողությունը և համեմատում փաստաթուղթը օրինակի հետ
'savePath - Փաստաթղթի պահպանման թղթապանակը
'fileName- Պահպանվող ֆայլի անվանումը
'pathExp- Օրինակի ճանապարհը 
Sub View_Doc_Action (savePath, fileName, pathExp, regex)
    
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_ViewDoc)
    BuiltIn.Delay(2000)
    If wMDIClient.WaitVBObject("FrmSpr",2000).Exists Then
        Call SaveDoc(savePath, fileName)
        Call Compare_Files(savePath & fileName, pathExp, regex)
        Call Close_Window(wMDIClient, "FrmSpr" )
    Else
        Log.Error "Can't find document print view",,,ErrorColor
    End If    
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''Excel_Find_Word'''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''
' objWhereSearch - օբյեկտի ճանապարհը, որում փնտրելու ենք
' searchedValue - փնտրվող արժեքը
Function Excel_Find_Word(objWhereSearch, searchedValue)
    Dim findWin
    
    objWhereSearch.Keys("[Home]")
    objWhereSearch.Keys("^f")
    If Sys.Process("EXCEL").WaitWindow("bosa_sdm_XL9", "Find and Replace", 1, 3000).Exists Then
        Set findWin = Sys.Process("EXCEL").Window("bosa_sdm_XL9", "Find and Replace", 1)
        ' Փնտրվող արժեքի ներմուծում
        findWin.Window("EDTBX", "", 1).Keys(searchedValue)
        ' Սեղմել "Find All" կոճակը
        findWin.Window("EDTBX", "", 1).Keys("[Tab][Tab]" &  "[Enter]")
        If findWin.Window("XLTFRCLASS", "", 1).Window("SysListView32", "", 1).wItemCount > 0 Then
            Excel_Find_Word = True
        Else
            Excel_Find_Word = False
        End If
        findWin.Close
    Else
        Log.Error "Can't open ""Find and Replace"" window.", "", pmNormal, ErrorColor
    End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''Delete_Conjunction'''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Պրոցեդուրան ջնջում է Կապակցված հաճախորդին ըստ մեկնաբանության
' delRowCount - Ջնջվող հաճախորդների քանակը
' pttelName - Պտտելի անունը
' searchedComment - փնտրվող մեկնաբանությունը
' expectedMessage - սպասվող հաղորդագրություն
Sub Delete_Conjunction(delRowCount, pttelName, searchedComment, expectedMessage)
    Dim i
    
    BuiltIn.Delay(3000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_ConjuctPersons & "|" & c_ViewAllConjucts) 
    If wMDIClient.WaitvbObject(pttelName, 3000).Exists Then
        For i = 0 To delRowCount - 1
            Call SearchAndDelete(pttelName, 0, searchedComment, expectedMessage)
            Call SearchAndDelete(pttelName, 0, searchedComment, expectedMessage)
        Next
        Call Close_Window(wMDIClient, pttelName)
    Else 
        Log.Error "Can't find "& pttelName & " window.", "", pmNormal, ErrorColor    
    End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''Call_Function'''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''
' moduleName - մոդուլի անունը 
' functionName - ֆունկցիայի անունը
' date - հաշվետվության սկզբի ամսաթիվ
Sub Call_Function(moduleName, functionName, date)
    BuiltIn.Delay(2000)
    Call wTreeView.DblClickItem("|²¹ÙÇÝÇëïñ³ïáñÇ ²Þî 4.0|Ð³Ù³Ï³ñ·³ÛÇÝ ³ßË³ï³ÝùÝ»ñ|Ð³Ù³Ï³ñ·³ÛÇÝ ·áñÍÇùÝ»ñ|üáõÝÏóÇ³ÛÇ Ï³Ýã")
    If p1.WaitVBObject("frmAsUstPar", 2000).Exists Then
        Call Rekvizit_Fill("Dialog", 1, "General", "MODULENAME", moduleName)
        Call Rekvizit_Fill("Dialog", 1, "General", "SUBNAME", functionName)
        Call ClickCmdButton(2, "Î³ï³ñ»É")
        If p1.WaitVBObject("frmAsUstPar", 2000).Exists Then
            Call Rekvizit_Fill("Dialog", 1, "General", "SDATE", date)
            Call ClickCmdButton(2, "Î³ï³ñ»É")
            If p1.WaitVBObject("frmAsMsgBox", 6000).Exists Then
                Call MessageExists(2, "¶áñÍáÕáõÃÛ³Ý µ³ñ»Ñ³çáÕ ³í³ñï")
                Call ClickCmdButton(5, "OK")
            Else
                Log.Error "Can't open frmAsMsgBox window.", "", pmNormal, ErrorColor
            End If
        Else
            Log.Error "Can't open frmAsUstPar window.", "", pmNormal, ErrorColor
        End If
    Else
        Log.Error "Can't open frmAsUstPar window.", "", pmNormal, ErrorColor
    End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''Export_From_OLTP''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Խմբերի արտահանում OLAP բազա
' toOlap - ExportToOlap կլասի օբյեկտ
' expDelay - արտահանման կատարման ժամանակահատված
Sub Export_From_OLTP(toOlap, expDelay)
    BuiltIn.Delay(2000)
    Call wTreeView.DblClickItem("|OLAP ·áñÍ³éÝ³í³ñÇ ²Þî|²ñï³Ñ³ÝáõÙ OLTP Ñ³Ù³Ï³ñ·Çó")
    If wMDIClient.WaitVBObject("frmOLAPExp", 3000).Exists Then 
        ' Լրացնել Հաշվետվությունների արտահանում (OLAP) պատուհանը
        Call Fill_Report_Export(toOlap)
        ' Սեղմել Կատարել կոճակը
        Call ClickCmdButton(14, "Î³ï³ñ»É")
        ' Ստուգել բաղված պատուհանի հաղորդագրությունը և սեղմել OK կոճակը
        If p1.WaitVBObject("frmAsMsgBox", expDelay).Exists Then
            Call MessageExists(2, "²ñï³Ñ³ÝáõÙÁ ³í³ñïí³Í ¿")
            Call ClickCmdButton(5, "OK")
        Else
            Log.Error "Can't open frmAsMsgBox window.", "", pmNormal, ErrorColor
        End If
        ' Փակել Արտահանում OLTP համակարգից պատուհանը
        Call Close_Window(wMDIClient, "frmOLAPExp")
    Else
        Log.Error "Can't open frmOLAPExp window.", "", pmNormal, ErrorColor
    End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''Check_Exported_Groups'''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Ֆունկցիան ստուգում է արտահանվել է խումբը թե ոչ 
' startDate - Սկիզբ դաշտի արժեք
' endDate - Վերջ դաշտի արժեք
' groupCode - Խմբի կոդ դաշտի արժեք
' expectedRowCount - սպասվող արտահանված տողերի քանակ 
Function Check_Exported_Groups(startDate, endDate, groupCode, expectedRowCount)
    BuiltIn.Delay(2000)
    Call wTreeView.DblClickItem("|OLAP ·áñÍ³éÝ³í³ñÇ ²Þî|ÀÝ¹áõÝí³Í ËÙµ»ñÇ áõÕÕáñ¹Çã")
    If wMDIClient.WaitVBObject("frmOLAPNav", 3000).Exists Then 
        wMDIClient.VBObject("frmOLAPNav").Maximize
        ' Լրացնել Սկիզբ դաշտը 
        Call Rekvizit_Fill("OLAPNavigator", 1, "General", "TDBDpern", startDate)
        ' Լրացնել վերջ դաշտը 
        Call Rekvizit_Fill("OLAPNavigator", 1, "General", "TDBDperk", endDate)
        ' Ստուգել, որ առկա է սպասվող քանակի 24 հաշվետվություն արտահանած տողեր 
        If Search_In_Grid(0, "OLAPNavigator", 1, groupCode) Then
            wMDIClient.VBObject("frmOLAPNav").VBObject("TDBGView").SetFocus
            If wMDIClient.VBObject("frmOLAPNav").VBObject("TDBGView").ApproxCount = expectedRowCount Then
                Check_Exported_Groups = True
            Else 
                Log.Error "Row count must be " & expectedRowCount & " , but it is " & _ 
                wMDIClient.VBObject("frmOLAPNav").VBObject("TDBGView").ApproxCount, "", pmNormal, ErrorColor
                Check_Exported_Groups = False
            End If
        Else
            Log.Error "Can't find row with " & groupCode & " group code.", "", pmNormal, ErrorColor
        End If
        wMDIClient.VBObject("frmOLAPNav").Minimize
    Else
        Log.Error "Can't open frmOLAPNav window.", "", pmNormal, ErrorColor
    End If
End Function

'Փաստաթղթի Վավերացում
Sub Confirm_Document()
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_ToConfirm)
    If wMDIClient.WaitVBObject("frmASDocForm",4000).Exists Then 
        Call ClickCmdButton(1, "Ð³ëï³ï»É")
    Else    
        Log.Error "Document window not found",,,ErrorColor
    End If
End Sub

'__________________________________________________________
'Համեմատում է ֆայլերը 
'__________________________________________________________
Sub Compare_Files(fromFile, toFile, param)

    Dim fso, file1, file2, regEx, attr
    Dim fileText1, fileText2, newText1, newText2
    Const ForReading = 1
    
'--------------------------------------
    Set attr = Log.CreateNewAttributes
    attr.BackColor = RGB(0, 255, 255)
    attr.Bold = True
    attr.Italic = True
'-------------------------------------- 
 
    ' Creates the FileSystemObject object
    Set fso = CreateObject("Scripting.FileSystemObject")
    BuiltIn.Delay(2000)
  
    ' Reads the first text file
    Set file1 = fso.OpenTextFile(fromFile, ForReading)
    fileText1 = file1.ReadAll
    file1.Close

    ' Reads the second text file
    Set file2 = fso.OpenTextFile(toFile, ForReading)
    fileText2 = file2.ReadAll
    file2.Close
    BuiltIn.Delay(2000)
    ' Creates the regular expression object
    Set regEx = New RegExp 
    regEx.Pattern = param
    If param <> "" Then
      regEx.IgnoreCase = True
    Else 
      regEx.IgnoreCase = False
    End If
    regEx.Global = True

    BuiltIn.Delay(2000)
    ' Replaces the text matching the specified date/time format with <ignore>
    newText1 = regEx.Replace(fileText1, "<ignore>")
    BuiltIn.Delay(2000)
    newText2 = regEx.Replace(fileText2, "<ignore>")
    BuiltIn.Delay(2000)
  
    ' Compares the text
    If newText1 = newText2 Then
        Call Log.Message("The files are identical.",,, attr)
    Else
        Log.Error("The files are NOT identical.")
    End If 
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''  üáõÝÏóÇ³Ý Ï³ï³ñáõÙ ¿ fileName1, fileName2 ×³Ý³å³ÑÝ»ñáí ·ïÝíáÕ ý³ÛÉ»ñÇ Ñ³Ù»Ù³ïáõÙ:
''  Ð³Ù»Ù³ïÙ³Ý Å³Ù³Ý³Ï dictExcludedPatterns-áõÙ å³ñáõÝ³ÏíáÕ ïáÕ³ÛÇÝ ß³µÉáÝÝ»ñÇ Ñ»ï Ñ³ÙÝÏÝáÕ
''  ïáÕ³ÛÇÝ ³ñï³Ñ³ÛïáõÃÛáõÝÝ»ñÁ ÷áË³ñÇÝíáõÙ »Ý "<ignore>"-áí:
Function Compare_Files_With_Patterns_Array(fileName1, fileName2, dictExcludedPatterns)

  Dim fso, file1, file2, regEx, iCount
  Dim fileText1, fileText2
  Const ForReading = 1

  ' Creates the FileSystemObject object
  Set fso = CreateObject("Scripting.FileSystemObject")

  ' Reads the first text file
  If Not fso.FileExists(fileName1) Then
    Log.Error("First file`" & fileName1 & " don't exits")
    Compare_Files_With_Patterns_Array = False
    Exit Function
  End If
  Set file1 = fso.OpenTextFile(fileName1, ForReading)
  fileText1 = file1.ReadAll
  file1.Close

  ' Reads the second text file
  If Not fso.FileExists(fileName2) Then
    Log.Error("Second file`" & fileName2 & " don't exits")
    Compare_Files_With_Patterns_Array = False
    Exit Function
  End If
  Set file2 = fso.OpenTextFile(fileName2, ForReading)
  fileText2 = file2.ReadAll
  file2.Close

  ' Creates the regular expression object
  Set regEx = New RegExp

  ' Specifies the pattern for the date/time mask
  ' MM.DD.YYYY HH:MM:SSLL (for example: 4/25/2006 10:51:35AM)
  ' pattern is` "\d{1,2}.\d{1,2}.\d{2,4}\s\d{1,2}:\d{2}:\d{2}\w{2}"
  For Each iCount in dictExcludedPatterns.Items
    regEx.Pattern = iCount
    regEx.IgnoreCase = True
    regEx.Global = True

    ' Replaces the text matching the specified date/time format with <ignore>
    fileText1 = regEx.Replace(fileText1, "<ignore>")
    fileText2 = regEx.Replace(fileText2, "<ignore>")
  Next

  ' Compares the text
  If fileText1 = fileText2 Then
    Compare_Files_With_Patterns_Array = True
    Log.Message "Files are identical.", "", pmNormal, MessageColor
  Else
    Compare_Files_With_Patterns_Array = False
    Log.Error "Files are NOT identical.", "", pmNormal, ErrorColor
  End If 
End Function

'------------------------------------------------------------------------------
' ÐÇß»É áñå»ë ·áñÍáÕáõÃÛ³Ý Ï³ï³ñáõÙ : üáõÝÏóÇ³Ý ÑÇßáõÙ ¿ ÷³ëï³ÃáõÕÃÁ ïñí³Í
'×³Ý³ñ³å³ñÑáí :
'------------------------------------------------------------------------------
'savePath - ü³ÛÉÇ å³ÑÙ³Ý ×³Ý³å³ñÑÁ
Sub SaveDoc( savePath, fName)
    BuiltIn.Delay(5000)
    Call wMainForm.MainMenu.Click(c_SaveAs)
    Sys.Process("Asbank").Window("#32770", "ÐÇß»É áñå»ë", 1).SaveFile savePath & fName
End Sub

'----------------------------------------------------------------------------------------
'Î³ñ·³íáñáõÙÝ»ñÇ Ý»ñÙáõÍáõÙ Ñ³Ù³Ï³ñ· : üáõÝÏóÇ³Ý í»ñ³¹³ñÓÝáõÙ ¿ true, »Ã» Ï³ñ·³íáñáõÙÝ»ñÁ
'µ³ñ»Ñ³çáÕ Ý»ñÙáõÍí»É »Ý , false` »Ã» áã:
'----------------------------------------------------------------------------------------
'confPath - Î³ñ·³íáñÙ³Ý ×³Ý³å³ñÑ
Function Input_Config(confPath)
  Dim cInput : cInput = False  
    
  BuiltIn.Delay(2000)
  Call ChangeWorkspace(c_BM)
  Call wTreeView.DblClickItem("|BankMail ²Þî|ÆÙ ÷³ëï³ÃÕÃ»ñ (ARMSOFT)|Ü»ñÙáõÍ»É Ï³ñ·³íáñáõÙÝ»ñÁ")
    
  If p1.WaitVBObject("frmAsUstPar", 3000).Exists Then
    Call Rekvizit_Fill("Dialog", 1, "General", "PATH", confPath)
    Call ClickCmdButton(2, "Î³ï³ñ»É")
  Else
    Log.Error "Can't find frmAsUstPar window", "", pmNormal, ErrorColor
  End If
    
  If p1.WaitVBObject("frmAsMsgBox", 10000).Exists Then
    If MessageExists(2, "¶áñÍáÕáõÃÛ³Ý µ³ñ»Ñ³çáÕ ³í³ñï") Then
      BuiltIn.Delay(4000)
      Call ClickCmdButton(5, "OK") 
      cInput = True
    Else
      cInput = False
    End If
  Else
    Log.Error "Can't find frmAsMsgBox window", "", pmNormal, ErrorColor
  End If
    
  Input_Config = cInput
  Call ChangeWorkspace(c_Admin)
End Function

'ֆունկցիան Պարամետրերում տալիս է արժեք
Sub SetParameter_InPttel(ParameterName, Value)
    If SearchInPttel("frmPttel", 1, ParameterName) Then
        Call wMainForm.MainMenu.Click(c_AllActions)
        Call wMainForm.PopupMenu.Click(c_ToEdit)
        If p1.WaitVbObject("frmAsUstPar", 3000).Exists Then
            'Լրացնում է "Արժեք" դաշտը
            Call Rekvizit_Fill("Dialog", 1, "General", "VALUE", "^A[Del]" & Value)
            'Սեղմել "Կատարել"
            Call ClickCmdButton(2, "Î³ï³ñ»É")
        Else
            Log.Error "Can't find frmAsUstPar window",,, ErrorColor
        End If
    Else
        Log.Error "Can Not find ("& ParameterName &") Parameter row!",,,ErrorColor
    End If
End Sub

'--------------------------------------------------------------------------------------
'Ð³ßí³é»É ·áñÍáÕáõÃÛ³Ý Ï³ï³ñáõÙ :
'--------------------------------------------------------------------------------------
Sub Register_Payment()
    BuiltIn.Delay(3000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_DoTrans)
    BuiltIn.Delay(2000)
    If MessageExists(2, "Ð³ßí³é»É") Then
        Call ClickCmdButton(5, "²Ûá")
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''--- GoToFolder_ByDocNum ---'''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Լրացնելով տրված "Փաստաթղթի N"-ը մուրք է գործում թղթապանակ
'FolderName - գտնբելու ճանապարհը
'FieldName - "Փաստաթղթի N" դաշտի անվանումը Օր.՝"NUM"
'DocNum - "Փաստաթղթի N"
Sub GoToFolder_ByDocNum(FolderName,FieldName,DocNum)
    
    Dim DocForm
    Call wTreeView.DblClickItem(FolderName)
    
    Set DocForm = p1.WaitVBObject("frmAsUstPar",2000)
    If DocForm.Exists Then
        'Լրացնում է "ä³ÛÙ³Ý³·ñÇ N" դաշտը
        Call Rekvizit_Fill("Dialog", 1, "General", FieldName, DocNum)
        Call ClickCmdButton(2, "Î³ï³ñ»É")
        BuiltIn.Delay(2000)
		Else 
				Log.Error "Can't open Filter widow!", "", pmNormal, ErrorColor
		End If
    Call WaitForPttel("frmPttel")
End Sub