Option Explicit

'USEUNIT Library_Common
'USEUNIT OLAP_Library
'USEUNIT Constants

'Test Case ID 166356

Sub Accountanty_24_Test()     

    Dim exists, SPath, userName, Passord, StartDate, EndDate, AccNumber, TreeLevel, CBranch, Language 
    Dim EPath1, EPath2, resultWorksheet(3), Thousand, RequesQuery, i, j, DB1, param, CurrDate, windExists, Cont
    Dim groupName, DateS, DateE, expOlap, expTXT, DateStart, DateEnd

    SPath = Project.Path & "Stores\Actual_OLAP"
    EPath1 = Project.Path & "Stores\Actual_OLAP\16600_24.xls"
    EPath2 = Project.Path & "Stores\Expected_OLAP\16600_24_190328.xls"
    
    For i = 1 To 2
      resultWorksheet(i) = Project.Path & "Stores\Result_Olap\Result_16600_24_sheet_" & i  & ".xls"
    Next
     'Î³ï³ñáõÙ ¿ ëïáõ·áõÙ,»Ã» ÝÙ³Ý ³ÝáõÝáí ý³ÛÉ Ï³ ïñí³Í ÃÕÃ³å³Ý³ÏáõÙ ,çÝçáõÙ ¿   
    exists = aqFile.Exists(EPath1)
    If exists Then
        aqFileSystem.DeleteFile(EPath1)
    End If
  
    DateStart = "20120101"
    DateEnd = "20240101"
    Call Initialize_AsBankQA(DateStart, DateEnd) 
    Call ChangeWorkspace(c_Admin40)
    
    Call wTreeView.DblClickItem("|²¹ÙÇÝÇëïñ³ïáñÇ ²Þî 4.0|Ð³Ù³Ï³ñ·³ÛÇÝ ³ßË³ï³ÝùÝ»ñ|Ð³Ù³Ï³ñ·³ÛÇÝ ·áñÍÇùÝ»ñ|üáõÝÏóÇ³ÛÇ Ï³Ýã")
    Call Rekvizit_Fill("Dialog",1,"General","MODULENAME","Util")
    Call Rekvizit_Fill("Dialog",1,"General","SUBNAME","UpdateDataForRep24")
    Call ClickCmdButton(2,"Î³ï³ñ»É")
    Call Rekvizit_Fill("Dialog",1,"General","SDATE","010114")
    Call ClickCmdButton(2,"Î³ï³ñ»É")
    BuiltIn.Delay(5000)
    Call ClickCmdButton(5,"OK")
    
    groupName = "FORM24"
    DateS = "280214"
    DateE = "280214"   
    expOlap = 1
    expTXT = 0
    Call ChangeWorkspace(c_OLAPAdmin)
    Call wTreeView.DblClickItem("|OLAP ³¹ÙÇÝÇëïñ³ïáñÇ ²Þî|OLAP ËÙµ»ñÇ ï»Õ»Ï³ïáõ")
    aqPerformance.Start()
    Call Find_Group_and_Export(groupName,DateS,DateE,expOlap,expTXT)
    Log.Message(aqPerformance.Value)    
    'ê»ÕÙ»É Ok Ïá×³ÏÁ 
    If p1.WaitVBObject("frmAsMsgBox", 10000).Exists Then 
      If Not Trim(p1.VBObject("frmAsMsgBox").VBObject("lblMessage").Caption) = "²ñï³Ñ³ÝáõÙÁ ³í³ñïí³Í ¿" Then 
        Log.Error "The actual message is " & p1.VBObject("frmAsMsgBox").VBObject("lblMessage").Caption,,,ErrorColor
      End If
      Call ClickCmdButton(5,"OK")
    End If
    
    Call Close_AsBank()
'    
    TestedApps.killproc.Run()
    
    Call Initialize_Excel ()
      
    Call Sys.Process("EXCEL").Window("XLMAIN", "Excel", 1).Window("FullpageUIHost").Window("NetUIHWND").Click(505, 236)
    
    Call AddOLAPAddIn ()
 
     userName = "ADMIN" 
     Passord= ""
     DB1 = "bankTesting_QA"
   
    'Î³ï³ñ»É ³ßË³ï³ÝùÇ ëÏÇ½µ
    Call Start_Work(userName ,Passord,DB1 )
     i = 0
     j = 20
    
    '´³ó»É Ñ³ßí»ïíáõÃÛ³Ü Ó¨³ÝÙáõß ïíÛ³ÉÝ»ñÇ å³ÑáóÇó
    Call Open_Accountanty(i,j)

    windExists = True
    CurrDate = Null
    StartDate = "28022014"
    EndDate = "28022014"
    AccNumber = 1
    TreeLevel = Null
    CBranch = "99997"
    Language  = "Հայերեն"
    Thousand = cbChecked
    RequesQuery = "60"
    param = "16600_24.xls"
    Cont = True    
    'Ð³ßí³ñÏ»É Ñ³ßí»ïíáõÃÛáõÝÁ 
    Call Calculate_Report_Range(windExists,CurrDate,StartDate,EndDate,AccNumber,TreeLevel,CBranch,Language ,Thousand,RequesQuery,param)
     'ä³Ñ»É ý³ÛÉÁ ACTUAL_OLAP ÃÕÃ³å³Ý³ÏáõÙ
   Call Save_To_Folder(SPath,param,Cont)

   'Ð³Ù»Ù³ï»É »ñÏáõ EXCEL ý³ÛÉ»ñ
    Call CompareTwoExcelFiles(EPath1, EPath2, resultWorksheet)
 
  
    'Î³ï³ñ»É ²ßË³ï³ÝùÇ ³í³ñï
    Sys.Process("EXCEL").Window("XLMAIN", "" & param & "  [Compatibility Mode] - Excel", 1).Window("EXCEL2", "", 2).ToolBar("Ribbon").Window("MsoWorkPane", "Ribbon", 1).Window("NUIPane", "", 1).Window("NetUIHWND", "", 1).Keys("~X")
    Sys.Process("EXCEL").Window("XLMAIN", "" & param & "  [Compatibility Mode] - Excel", 1).Window("EXCEL2", "", 2).ToolBar("Ribbon").Window("MsoWorkPane", "Ribbon", 1).Window("NUIPane", "", 1).Window("NetUIHWND", "", 1).Keys("Y7")
 
    'ö³Ï»É EXCEL- Á
    Call CloseAllExcelFiles()
    
'    Call Close_AsBank()
  '  TestedApps.killproc.Run()
End Sub