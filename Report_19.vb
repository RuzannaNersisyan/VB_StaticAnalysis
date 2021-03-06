'USEUNIT SWIFT_International_Payorder_Library
'USEUNIT Library_Common
'USEUNIT OLAP_Library
'USEUNIT Constants

Option Explicit
'Test Case ID 123836

Sub Report_19_Test()
  
      Dim coaNum,dateE,trans,cbCode,DateStart,DateEnd,branch,nbTurn
      Dim file1, file2, param,exists
      Dim EPath1, EPath2, resultWorksheet
    
      DateStart = "20120101"
      DateEnd = "20240101" 

      EPath1 = Project.Path & "Stores\CB\Actual\19.xlsx"
      EPath2 = Project.Path & "Stores\CB\Expected\Expected 19.xlsx"
      resultWorksheet = Project.Path & "Stores\Result_Olap\Result_CB_19.xlsx"
    
     'Î³ï³ñáõÙ ¿ ëïáõ·áõÙ,»Ã» ÝÙ³Ý ³ÝáõÝáí ý³ÛÉ Ï³ ïñí³Í ÃÕÃ³å³Ý³ÏáõÙ ,çÝçáõÙ ¿   
      exists = aqFile.Exists(EPath1)
      If exists Then
          aqFileSystem.DeleteFile(EPath1)
      End If

      'Մուտք գործել ՀԾ- Բանկ համակարգ ARMSOFT օգտագործողով
      Call Initialize_AsBankQA(DateStart, DateEnd) 
 
      'Անցում կատարել "Ենթահամակարգեր" ԱՇՏ
      Call ChangeWorkspace(c_Subsystems)
      Call wTreeView.DblClickItem("|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|Ð³ßí»ïíáõÃÛáõÝÝ»ñ,  Ù³ïÛ³ÝÝ»ñ|Î´ Ñ³ßí»ïíáõÃÛáõÝÝ»ñ|19 ²ñï³ñÅáõÛÃÇ ³éù/í³×³éù")
    
      dateE = "280214"  
      trans = 0
      nbTurn = 1
    
      'Լրացնում է  "Ժամանակահատված(սկիզբ)" դաշտը
      Call Rekvizit_Fill("Dialog",1 ,"General" ,"SDATE" ,dateE)
      'Լրացնում է  "Ժամանակահատված(վերջ)" դաշտը
      Call Rekvizit_Fill("Dialog",1 ,"General" ,"EDATE" ,dateE)
      'Լրացնում է "Ժամանակ" դաշտը
      Call Rekvizit_Fill("Dialog",1 ,"General" ,"TIME" ,coaNum)
      'Լրացնում է "Ներառել հաշվեկշռային գործողությունները" նշիչը
      Call Rekvizit_Fill("Dialog",1 ,"CheckBox" ,"BALOP" ,nbTurn)
      'Լրացնում է "Ներառել դիլինգային պայմանագրերը" նշիչը
      Call Rekvizit_Fill("Dialog",1 ,"CheckBox" ,"OFFBOP" ,nbTurn)
      'Լրացնում է "Միայն առքի/վաճառքի գործողությունները" նշիչը
      Call Rekvizit_Fill("Dialog",1 ,"CheckBox" ,"ONLYPUSA" ,trans)
      'Լրացնում է "Խմբավորել նաև ըստ գործ. ոլորտի " նշիչը
      Call Rekvizit_Fill("Dialog",1 ,"CheckBox" ,"GRPBYVOL" ,trans)
      'Լրացնում է "Ցույց տալ բացված տեսքով" նշիչը
      Call Rekvizit_Fill("Dialog",1 ,"CheckBox" ,"UNGROUP" ,trans)
      'Լրացնում է "Գրասենյակ" դաշտը
      Call Rekvizit_Fill("Dialog",1 ,"General" ,"ACSBRANCH" ,branch)
      'Սեղմել "Կատարել" կոճակը
      Call ClickCmdButton(2 ,"Î³ï³ñ»É")
    
      BuiltIn.Delay(8000)
        
      'Դասավորել ըստ ²éùÇ ÙÇçÇÝ ÷áË³ñÅ»ù ëÛ³Ý 
      Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").VBObject("tdbgView").Keys("^8")
      BuiltIn.Delay(2000)     
      Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").VBObject("TDBGView").MoveFirst
      If Not Trim(Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").VBObject("tdbgView").Columns.Item(7).Text) = "0.0000" Then 
        Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").VBObject("tdbgView").Keys("^8")
      End If
      'êáñï³íáñ»É Áëï ì³×³éùÇ Í³í³É ëÛ³Ý
      Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").VBObject("tdbgView").Keys("[Hold]" & "^!"& "9")
    
      'Արտահանել EXCEL ֆայլ
      Call Sys.Process("Asbank").VBObject("MainForm").VBObject("tbToolBar").Window("ToolbarWindow32", "", 1).ClickItem(27)    
      BuiltIn.Delay(2000)    
      Sys.Process("EXCEL").Window("XLMAIN", "* - Excel", 1).Window("XLDESK", "", 1).Window("EXCEL7", "*", 1).Keys("[F12]")

      Builtin.Delay(2000)
      Sys.Process("EXCEL").Window("#32770", "Save As", 1).Window("DUIViewWndClassName", "", 1).Window("DirectUIHWND", "", 1).Window("FloatNotifySink", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1).Keys(Project.Path & "Stores\CB\Actual\19.xlsx")
      Sys.Process("EXCEL").Window("#32770", "Save As", 1).Window("Button", "&Save", 1).Click()    
      BuiltIn.Delay(3000)

      'Ð³Ù»Ù³ï»É »ñÏáõ EXCEL ý³ÛÉ»ñ
      Call CompareTwoExcelFiles(EPath1, EPath2,resultWorksheet)
    
      'Փակել ՀԾ - Բանկ համակարգը
      Call Close_AsBank()
      Call CloseAllExcelFiles()
    
End Sub
