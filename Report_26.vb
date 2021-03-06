'USEUNIT SWIFT_International_Payorder_Library
'USEUNIT Library_Common
'USEUNIT OLAP_Library
'USEUNIT Constants

Option Explicit
'Test Case ID 123846

Sub Report_26_Test()  
  
    Dim dateS,dateE,repay,trans,DateStart,DateEnd,branch,newForm
    Dim file1, file2, param,exists
    Dim EPath1, EPath2, resultWorksheet
    DateStart = "20120101"
    DateEnd = "20240101" 
     
    EPath1 = Project.Path & "Stores\CB\Actual\26_2.xlsx"
    EPath2 = Project.Path & "Stores\CB\Expected\Expected 26_2.xlsx"
    resultWorksheet = Project.Path & "Stores\Result_Olap\Result_CB_26.xlsx"
    
     'Î³ï³ñáõÙ ¿ ëïáõ·áõÙ,»Ã» ÝÙ³Ý ³ÝáõÝáí ý³ÛÉ Ï³ ïñí³Í ÃÕÃ³å³Ý³ÏáõÙ ,çÝçáõÙ ¿   
    exists = aqFile.Exists(EPath1)
    If exists Then
        aqFileSystem.DeleteFile(EPath1)
    End If
    
    
    'Մուտք գործել ՀԾ- Բանկ համակարգ ARMSOFT օգտագործողով
    Call Initialize_AsBankQA(DateStart, DateEnd) 
 
    'Անցում կատարել "Ենթահամակարգեր" ԱՇՏ
    Call ChangeWorkspace(c_Subsystems)
    Call wTreeView.DblClickItem("|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|Ð³ßí»ïíáõÃÛáõÝÝ»ñ,  Ù³ïÛ³ÝÝ»ñ|Î´ Ñ³ßí»ïíáõÃÛáõÝÝ»ñ|26 ²í³Ý¹Ý»ñÇ ¹ÇÙ³ó »ñ³ßË. í×³ñáõÙÝ»ñÇ  í»ñ³µ»ñÛ³É")
    
    dateS = "280214"
    repay = "4000000"
    trans = "2000000"
    newForm = 0
    
    'Լրացնում է  "Ամսաթիվ" դաշտը
    Call Rekvizit_Fill("Dialog",1 ,"General" ,"DATEREP" ,dateS)
    'Լրացնում է "Դրամայաին ավանդների սահման" դաշտը
    Call Rekvizit_Fill("Dialog",1 ,"General" ,"CRANGEAMD" ,repay)
    'Լրացնում է "Արժույթային ավանդների սահման" դաշտը
    Call Rekvizit_Fill("Dialog",1 ,"General" ,"CRANGECUR" ,trans)
    'Լրացնում է "Ցույց տալ բոլոր բաղադրիչ գումարները" նշիչը
    Call Rekvizit_Fill("Dialog",1 ,"CheckBox" ,"SHOWCOMMENT" ,newForm)
    'Սեղմել "Կատարել" կոճակը
    Call ClickCmdButton(2 ,"Î³ï³ñ»É")
    
    BuiltIn.Delay(250000)
    
    'Սեղմել "Հիշել որպես"
    Call wMainForm.MainMenu.Click(c_SaveAs)
    Sys.Process("Asbank").Window("#32770", "ÐÇß»É áñå»ë", 1).Window("DUIViewWndClassName", "", 1).Window("DirectUIHWND", "", 1).Window("FloatNotifySink", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1).Keys(Project.Path & "Stores\CB\Actual\26.txt")
    Sys.Process("Asbank").Window("#32770", "ÐÇß»É áñå»ë", 1).Window("Button", "&Save", 1).Click()
    
    'Համեմատել ֆայլերը
    file1 = Project.Path & "Stores\CB\Actual\26.txt"
    file2 = Project.Path & "Stores\CB\Expected\Expected 26_1.txt"
    Call Compare_Files(file1, file2,param)
    
    Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("FrmSpr").Close()
    
    Call wTreeView.DblClickItem("|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|Ð³ßí»ïíáõÃÛáõÝÝ»ñ,  Ù³ïÛ³ÝÝ»ñ|Î´ Ñ³ßí»ïíáõÃÛáõÝÝ»ñ|26_4 ²í³Ý¹Ý»ñÇ ³é³í»É³·áõÛÝ ïáÏáë³¹ñáõÛùÝ»ñÁ")
    
    dateS = "010214"
    dateE = "280214"
    
    'Լրացնում է  "Ժամանակահատված(սկիզբ)" դաշտը
    Call Rekvizit_Fill("Dialog",1 ,"General" ,"DATEBEG" ,dateS)
    'Լրացնում է "Ժամանակահատված(վերջ)" դաշտը
    Call Rekvizit_Fill("Dialog",1 ,"General" ,"DATEEND" ,dateE)
    'Սեղմել "Կատարել" կոճակը
    Call ClickCmdButton(2 ,"Î³ï³ñ»É")    
    BuiltIn.Delay(6000)
    
    'Արտահանել EXCEL ֆայլ
    Call Sys.Process("Asbank").VBObject("MainForm").VBObject("tbToolBar").Window("ToolbarWindow32", "", 1).ClickItem(27)        
    BuiltIn.Delay(2000)
    Sys.Process("EXCEL").Window("XLMAIN", "* - Excel", 1).Window("XLDESK", "", 1).Window("EXCEL7", "*", 1).Keys("[F12]")

    Builtin.Delay(2000)
    Sys.Process("EXCEL").Window("#32770", "Save As", 1).Window("DUIViewWndClassName", "", 1).Window("DirectUIHWND", "", 1).Window("FloatNotifySink", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1).Keys(Project.Path & "Stores\CB\Actual\26_2.xlsx")
    Sys.Process("EXCEL").Window("#32770", "Save As", 1).Window("Button", "&Save", 1).Click()
    
    'Համեմատել ֆայլերը
     Call CompareTwoExcelFiles(EPath1, EPath2,resultWorksheet)
     
    'Փակել ՀԾ - Բանկ համակարգը
    Call Close_AsBank()
    Call CloseAllExcelFiles()  
    
End Sub