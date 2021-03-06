Option Explicit
'USEUNIT International_PayOrder_Receive_Confirmphases_Library
'USEUNIT International_PayOrder_ConfirmPhases_Library
'USEUNIT PayOrder_Receive_ConfirmPhases_Library
'USEUNIT SWIFT_International_Payorder_Library
'USEUNIT Payment_Order_ConfirmPhases_Library
'USEUNIT Payment_Except_Library
'USEUNIT Subsystems_SQL_Library
'USEUNIT Online_PaySys_Library
'USEUNIT Akreditiv_Library
'USEUNIT Library_Common
'USEUNIT Constants

'Test case Id 166764

Sub SWIFT_Acc_Extract_FileAct_Test(SysType)

    Dim max,min,rand, sDATE, fDATE,DocNum,cashOutN
    Dim fromFile,toFile,what,fWith,isExists,fBASE,param
    Dim queryString,sql_Value, colNum,sql_isEqual,result,fOBJECT
    Dim curr_date,category,receipt, bankName,fileName
    Dim startDate , endDate , stype , bank , comm , showAcc , countPeriod
    
    sDATE = "20010101"
    fDATE = "20250101"
    param = "(((:[2][8][A-Z])|([1][3][D])|([2][0]:)|[2][1])([0-9:]+)|(:[2][8]:[0-9:]+))|([[6][2][F]:[C]......)|[$]................................"
     
    Call Initialize_AsBank("bank", sDATE, fDATE)
    
    aqFileSystem.DeleteFile(Project.Path & "Stores\SWIFTtest\Actual\FileAct\*.RJE")
    
    Select Case SysType
              Case 1
                  Call SetParameter("SWIN", Project.Path& "Stores\SWIFTtest\Actual\")
                  Call SetParameter("SWFAIN", Project.Path& "Stores\SWIFTtest\Actual\FileAct\")
                  Call SetParameter("SWFAOUT", Project.Path& "Stores\SWIFTtest\Import\FileAct\")
                  Call SetParameter("SWOUT", Project.Path& "Stores\SWIFTtest\Import\")
                  Call SetParameter("SWTMPDIR", "\\host2\Sys\Testing\SWIFT\tmp\")
                  Call SetParameter("SWSPFSIN", "")
                  Call SetParameter("SWSPFSCLIENTS", "")
                  
              Case 2
                  Call SetParameter("SWIN", "")
                  Call SetParameter("SWFAIN", "")
                  Call SetParameter("SWFAOUT", "")
                  Call SetParameter("SWOUT", "")
                  Call SetParameter("SWTMPDIR", "\\host2\Sys\Testing\SWIFT\tmp\")
                  Call SetParameter("SWSPFSIN", Project.Path& "Stores\SWIFTtest\Actual\FileAct\")
                  Call SetParameter("SWSPFSCLIENTS", "UBSWCHZHXXX")
      End Select
    
    Call Create_Connection()
    
    Call ChangeWorkspace(c_SWIFT)
    Call wTreeView.DblClickItem("|S.W.I.F.T. ²Þî                  |î»Õ»Ï³ïáõÝ»ñ|Ð³ßÇíÝ»ñ")
    Call Rekvizit_Fill("Dialog",1,"General","ACCMASK" ,"00100770100"  )
    'Sys.Process("Asbank").VBObject("frmAsUstPar").VBObject("CmdOK").Click()
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    
    'Քաղվածքի տրամադրում SWIFT
    startDate = "060611"
    endDate = aqConvert.DateTimeToFormatStr(aqDateTime.Today(), "%d/%m/%y")    
    stype = "950"
    bank =  "UBSWCHZHXXX" 
    comm = 1
    showAcc = 1
    countPeriod = 1
    Call Acc_State_SWIFT(startDate , endDate , stype , bank , comm , showAcc , countPeriod)
    Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").Close()
    Call wTreeView.DblClickItem("|S.W.I.F.T. ²Þî                  |àõÕ³ñÏíáÕ Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñ|àõÕ³ñÏíáÕ Ë³éÁ Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñ")
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    
    'Ֆիլտրել օրերի քանակը
    Call wMainForm.MainMenu.Click("Դիտում |bankFiltr")    
    'Քաղվածքի խմբագրում
    bankName = NULL
    Call Edit_Acc_State(bankName,stype)
    'Ուղարկել SWIFT կամ հաստատման
    Call Send_SWIFT_or_Confirm(category,receipt)
    Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").Close()
    
    BuiltIn.Delay(1000)
    fileName = ListFiles(Project.Path & "Stores\SWIFTtest\Actual\FileAct")
    toFile = Project.Path & "Stores\SWIFTtest\Actual\FileAct\" & Trim(fileName)
    fromFile = Project.Path &"Stores\SWIFTtest\Expected\MT950_410.RJE"
    
    Call Compare_Files(fromFile, toFile,param)
      
'--------------------------------------940----------------------------------------------------------------------------------
   
    aqFileSystem.DeleteFile(Project.Path & "Stores\SWIFTtest\Actual\FileAct\*.RJE")

    Call wTreeView.DblClickItem("|S.W.I.F.T. ²Þî                  |î»Õ»Ï³ïáõÝ»ñ|Ð³ßÇíÝ»ñ")
    Call Rekvizit_Fill("Dialog",1,"General","ACCMASK" ,"00100770100"  )
    'Sys.Process("Asbank").VBObject("frmAsUstPar").VBObject("CmdOK").Click()
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    
    'Քաղվածքի տրամադրում SWIFT
    startDate = "060611"
    endDate = aqConvert.DateTimeToFormatStr(aqDateTime.Today(), "%d/%m/%y")    
    stype = "940"
    bank =  "UBSWCHZHXXX" 
    comm = 1
    showAcc = 1
    countPeriod = 1
    Call Acc_State_SWIFT(startDate , endDate , stype , bank , comm , showAcc , countPeriod)
    Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").Close()
    Call wTreeView.DblClickItem("|S.W.I.F.T. ²Þî                  |àõÕ³ñÏíáÕ Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñ|àõÕ³ñÏíáÕ Ë³éÁ Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñ")
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    
    'Ֆիլտրել օրերի քանակը
    Call wMainForm.MainMenu.Click("Դիտում |bankFiltr")
    bankName = NULL
    'Քաղվածքի խմբագրում
    Call Edit_Acc_State(bankName,stype)
    'Ուղարկել SWIFT կամ հաստատման
    Call Send_SWIFT_or_Confirm(category,receipt)
    Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").Close()
    
    BuiltIn.Delay(1000)
    fileName = ListFiles(Project.Path & "Stores\SWIFTtest\Actual\FileAct")
    toFile = Project.Path & "Stores\SWIFTtest\Actual\FileAct\" & Trim(fileName)
    fromFile = Project.Path &"Stores\SWIFTtest\Expected\MT940_410.RJE"
    
    Call Compare_Files(fromFile, toFile,param)
    
    
'--------------------------------------941----------------------------------------------------------------------------------
   
    aqFileSystem.DeleteFile(Project.Path & "Stores\SWIFTtest\Actual\FileAct\*.RJE")
    
    Call wTreeView.DblClickItem("|S.W.I.F.T. ²Þî                  |î»Õ»Ï³ïáõÝ»ñ|Ð³ßÇíÝ»ñ")
    Call Rekvizit_Fill("Dialog",1,"General","ACCMASK" ,"00100770100"  )
    'Sys.Process("Asbank").VBObject("frmAsUstPar").VBObject("CmdOK").Click()
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    
    'Քաղվածքի տրամադրում SWIFT
    startDate = "060611"
    endDate = aqConvert.DateTimeToFormatStr(aqDateTime.Today(), "%d/%m/%y")    
    stype = "941"
    bank =  "UBSWCHZHXXX" 
    comm = 1
    showAcc = 1
    countPeriod = 1
    Call Acc_State_SWIFT(startDate , endDate , stype , bank , comm , showAcc , countPeriod)
    Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").Close()
    Call wTreeView.DblClickItem("|S.W.I.F.T. ²Þî                  |àõÕ³ñÏíáÕ Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñ|àõÕ³ñÏíáÕ Ë³éÁ Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñ")
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    
    'Ֆիլտրել օրերի քանակը
    Call wMainForm.MainMenu.Click("Դիտում |bankFiltr")
    bankName = NULL
    'Քաղվածքի խմբագրում
    Call Edit_Acc_State(bankName,stype)
    'Ուղարկել SWIFT կամ հաստատման
    Call Send_SWIFT_or_Confirm(category,receipt)
    Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").Close()
    
    BuiltIn.Delay(1000)
    fileName = ListFiles(Project.Path & "Stores\SWIFTtest\Actual\FileAct")
    toFile = Project.Path & "Stores\SWIFTtest\Actual\FileAct\" & Trim(fileName)
    fromFile = Project.Path &"Stores\SWIFTtest\Expected\MT941_410.RJE"
    
    Call Compare_Files(fromFile, toFile,param)
    
'--------------------------------------942----------------------------------------------------------------------------------
   
    aqFileSystem.DeleteFile(Project.Path & "Stores\SWIFTtest\Actual\FileAct\*.RJE")
    
    Call wTreeView.DblClickItem("|S.W.I.F.T. ²Þî                  |î»Õ»Ï³ïáõÝ»ñ|Ð³ßÇíÝ»ñ")
    Call Rekvizit_Fill("Dialog",1,"General","ACCMASK" ,"00100770100"  )
    'Sys.Process("Asbank").VBObject("frmAsUstPar").VBObject("CmdOK").Click()
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    
    'Քաղվածքի տրամադրում SWIFT
    startDate = "060611"
    endDate = aqConvert.DateTimeToFormatStr(aqDateTime.Today(), "%d/%m/%y")    
    stype = "942"
    bank =  "UBSWCHZHXXX" 
    comm = 1
    showAcc = 1
    countPeriod = 1
    Call Acc_State_SWIFT(startDate , endDate , stype , bank , comm , showAcc , countPeriod)
    Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").Close()
    Call wTreeView.DblClickItem("|S.W.I.F.T. ²Þî                  |àõÕ³ñÏíáÕ Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñ|àõÕ³ñÏíáÕ Ë³éÁ Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñ")
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    
    'Ֆիլտրել օրերի քանակը
    Call wMainForm.MainMenu.Click("Դիտում |bankFiltr")
    bankName = NULL
    'Քաղվածքի խմբագրում
    Call Edit_Acc_State(bankName,stype)
    'Ուղարկել SWIFT կամ հաստատման
    Call Send_SWIFT_or_Confirm(category,receipt)
    Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").Close()
    
    BuiltIn.Delay(1000)
    fileName = ListFiles(Project.Path & "Stores\SWIFTtest\Actual\FileAct")
    toFile = Project.Path & "Stores\SWIFTtest\Actual\FileAct\" & Trim(fileName)
    fromFile = Project.Path &"Stores\SWIFTtest\Expected\MT942_410.RJE"
    
    Call Compare_Files(fromFile, toFile,param)
    
    Call Close_AsBank()    
    
End Sub