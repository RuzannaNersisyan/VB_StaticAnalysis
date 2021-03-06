'USEUNIT SWIFT_International_Payorder_Library
'USEUNIT Library_Common
'USEUNIT OLAP_Library
'USEUNIT Constants

Option Explicit
'Test Case ID 122692

Sub Report_09_Test()
  
      Dim dateS,dateE,avarage,DateStart,DateEnd,branch,nbTurn
      Dim file1, file2, param
    
      DateStart = "20120101"
      DateEnd = "20240101" 

      'Մուտք գործել ՀԾ- Բանկ համակարգ ARMSOFT օգտագործողով
      Call Initialize_AsBankQA(DateStart, DateEnd) 
 
      'Անցում կատարել "Ենթահամակարգեր" ԱՇՏ
      Call ChangeWorkspace(c_Subsystems)
      Call wTreeView.DblClickItem("|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|Ð³ßí»ïíáõÃÛáõÝÝ»ñ,  Ù³ïÛ³ÝÝ»ñ|Î´ Ñ³ßí»ïíáõÃÛáõÝÝ»ñ|09 Ü»ñ·ñ³íí³Í ¨ ï»Õ³µ³ß. ÙÇç. Å³ÙÏ»ï³ÛÝáõÃÛáõÝ")
    
      dateE = "280214" 
      nbTurn = 0
    
      'Լրացնում է  "Ժամանակահատված" դաշտը
      Call Rekvizit_Fill("Dialog",1 ,"General" ,"DATEBEG" ,dateE)
      'Լրացնում է "ժամանակահատված" դաշտը
      Call Rekvizit_Fill("Dialog",1 ,"General" ,"DATEEND" ,dateE)
      'Լրացնում է "Գրասենյակ" դաշտը
      Call Rekvizit_Fill("Dialog",1 ,"General" ,"BRANCH" ,branch)
      'Լրացնում է "Հաշվարկել բանկի հետ կախված փոփոխությունները"  նշիչը
      Call Rekvizit_Fill("Dialog",1 ,"CheckBox" ,"CALCATMCOD" ,nbTurn)
      'Լրացնում է Խմբավորել ըստ հաճախորդների նշիչը
      Call Rekvizit_Fill("Dialog",1 ,"CheckBox" ,"CALCLIABIL" ,nbTurn)
      'Լրացնում է Խմբավորել ըստ հաճախորդների դաշտը
      Call Rekvizit_Fill("Dialog",1 ,"General" ,"LIABILITIES" ,avarage)
      'Լրացնում է Խմբավորել ըստ հաճախորդների նշիչը
      Call Rekvizit_Fill("Dialog",1 ,"CheckBox" ,"SHOWCOMMENT" ,nbTurn)
      'Սեղմել "Կատարել" կոճակը
      Call ClickCmdButton(2 ,"Î³ï³ñ»É")
    
      BuiltIn.Delay(2400000)
    
      'Սեղմել "Հիշել որպես"
      Call wMainForm.MainMenu.Click(c_SaveAs)
      Sys.Process("Asbank").Window("#32770", "ÐÇß»É áñå»ë", 1).Window("DUIViewWndClassName", "", 1).Window("DirectUIHWND", "", 1).Window("FloatNotifySink", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1).Keys(Project.Path & "Stores\CB\Actual\09_1.txt")
      Sys.Process("Asbank").Window("#32770", "ÐÇß»É áñå»ë", 1).Window("Button", "&Save", 1).Click()
    
      'Համեմատել ֆայլերը
      file1 = Project.Path & "Stores\CB\Actual\09_1.txt"
      file2 = Project.Path & "Stores\CB\Expected\Expected 09_1.txt"
      Call Compare_Files(file1, file2,param)
    
      Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("FrmSpr").Close()
    
      Call wTreeView.DblClickItem("|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|Ð³ßí»ïíáõÃÛáõÝÝ»ñ,  Ù³ïÛ³ÝÝ»ñ|Î´ Ñ³ßí»ïíáõÃÛáõÝÝ»ñ|09_2 /9-ñ¹ Ñ³ßí»ïíáõÃÛ³Ý 2-ñ¹ ï³ñµ»ñ³Ï/")
    
      nbTurn = 1
    
      'Լրացնում է  "Ժամանակահատված(սկիզբ)" դաշտը
      Call Rekvizit_Fill("Dialog",1 ,"General" ,"DATEBEG" ,dateE)
      'Լրացնում է "ժամանակահատված(վերջ)" դաշտը
      Call Rekvizit_Fill("Dialog",1 ,"General" ,"DATEEND" ,dateE)
      'Լրացնում է "Գրասենյակ" դաշտը
      Call Rekvizit_Fill("Dialog",1 ,"General" ,"BRANCH" ,branch)
      'Լրացնում է "Հաշվել "  դաշտը
      Call Rekvizit_Fill("Dialog",1 ,"CheckBox" ,"CALCATMCOD" ,nbTurn)
      'Լրացնում է Խմբավորել ըստ հաճախորդների դաշտը
      Call Rekvizit_Fill("Dialog",1 ,"CheckBox" ,"CALCLIABIL" ,nbTurn)
      'Լրացնում է Խմբավորել ըստ հաճախորդների դաշտը
      Call Rekvizit_Fill("Dialog",1 ,"General" ,"LIABILITIES" ,avarage)
      'Լրացնում է Խմբավորել ըստ հաճախորդների դաշտը
      Call Rekvizit_Fill("Dialog",1 ,"CheckBox" ,"SHOWCOMMENT" ,nbTurn)
      'Սեղմել "Կատարել" կոճակը
      Call ClickCmdButton(2 ,"Î³ï³ñ»É")
    
      BuiltIn.Delay(5000000)
    
      'Սեղմել "Հիշել որպես"
      Call wMainForm.MainMenu.Click(c_SaveAs)
      Sys.Process("Asbank").Window("#32770", "ÐÇß»É áñå»ë", 1).Window("DUIViewWndClassName", "", 1).Window("DirectUIHWND", "", 1).Window("FloatNotifySink", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1).Keys(Project.Path & "Stores\CB\Actual\09_2.txt")
      Sys.Process("Asbank").Window("#32770", "ÐÇß»É áñå»ë", 1).Window("Button", "&Save", 1).Click()
    
      'Համեմատել ֆայլերը
      file1 = Project.Path & "Stores\CB\Actual\09_2.txt"
      file2 = Project.Path & "Stores\CB\Expected\Expected 09_2.txt"
      Call Compare_Files(file1, file2,param)
    
      'Փակել ՀԾ - Բանկ համակարգը    
      Call Close_AsBank()
    
End Sub