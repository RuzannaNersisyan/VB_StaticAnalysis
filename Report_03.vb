'USEUNIT SWIFT_International_Payorder_Library
'USEUNIT Library_Common
'USEUNIT OLAP_Library
'USEUNIT Constants

Option Explicit
'Test Case ID 121738

Sub Report_03_Test()
  
      Dim dateS,dateE,repay,trans,cbCode,DateStart,DateEnd,branch,nbTurn
      Dim file1, file2, param
    
      DateStart = "20120101"
      DateEnd = "20240101" 
      
      'Մուտք գործել ՀԾ- Բանկ համակարգ ARMSOFT օգտագործողով
      Call Initialize_AsBankQA(DateStart, DateEnd) 
 
      'Անցում կատարել "Ենթահամակարգեր" ԱՇՏ
      Call ChangeWorkspace(c_Subsystems)
      Call wTreeView.DblClickItem("|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|Ð³ßí»ïíáõÃÛáõÝÝ»ñ,  Ù³ïÛ³ÝÝ»ñ|Î´ Ñ³ßí»ïíáõÃÛáõÝÝ»ñ|03_3,4 ´³ÝÏÇ Ñ»ï Ï³åí³Í ³ÝÓ³Ýó ÝÏ³ïÙ³Ùµ å³Ñ³ÝçÝ»ñ")
    
      dateS = "010214"
      dateE = "280214" 
      trans = 1
      nbTurn = 1
    
      'Լրացնում է  "Ժամանակահատված(սկիզբ)" դաշտը
      Call Rekvizit_Fill("Dialog",1 ,"General" ,"DATEBEG" ,dateS)
      'Լրացնում է "ժամանակահատված(վերջ)" դաշտը
      Call Rekvizit_Fill("Dialog",1 ,"General" ,"DATEEND" ,dateE)
      'Լրացնում է "Գրասենյակ" դաշտը
      Call Rekvizit_Fill("Dialog",1 ,"General" ,"BRANCH" ,branch)
      'Լրացնում է "Հաշվարկել բանկի հետ կախված փոփոխությունները"  նշիչը
      Call Rekvizit_Fill("Dialog",1 ,"CheckBox" ,"SHEET4" ,trans)
      'Լրացնում է Խմբավորել ըստ հաճախորդների նշիչը
      Call Rekvizit_Fill("Dialog",1 ,"CheckBox" ,"GROUP" ,nbTurn)
      'Սեղմել "Կատարել" կոճակը
      Call ClickCmdButton(2 ,"Î³ï³ñ»É")
    
      BuiltIn.Delay(250000)
    
      'Սեղմել "հիշել որպես"
      Call wMainForm.MainMenu.Click(c_SaveAs)
      Sys.Process("Asbank").Window("#32770", "ÐÇß»É áñå»ë", 1).Window("DUIViewWndClassName", "", 1).Window("DirectUIHWND", "", 1).Window("FloatNotifySink", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1).Keys(Project.Path & "Stores\CB\Actual\03.txt")
      Sys.Process("Asbank").Window("#32770", "ÐÇß»É áñå»ë", 1).Window("Button", "&Save", 1).Click()
    
      'Համեմատել ֆայլերը
      file1 = Project.Path & "Stores\CB\Actual\03.txt"
      file2 = Project.Path & "Stores\CB\Expected\Expected 03.txt"
      Call Compare_Files(file1, file2,param)
    
      'Փակել ՀԾ - Բանկ համակարգը
      Call Close_AsBank()
    
End Sub