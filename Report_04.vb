'USEUNIT SWIFT_International_Payorder_Library
'USEUNIT Library_Common
'USEUNIT OLAP_Library
'USEUNIT Constants

Option Explicit
'Test Case ID 121745

Sub Report_04_Test()
  
      Dim dateS,DateStart,DateEnd
      Dim file1, file2, param
    
      DateStart = "20120101"
      DateEnd = "20240101" 

      'Մուտք գործել ՀԾ- Բանկ համակարգ ARMSOFT օգտագործողով
      Call Initialize_AsBankQA(DateStart, DateEnd) 
 
      'Անցում կատարել "Ենթահամակարգեր" ԱՇՏ
      Call ChangeWorkspace(c_Subsystems)
      Call wTreeView.DblClickItem("|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|Ð³ßí»ïíáõÃÛáõÝÝ»ñ,  Ù³ïÛ³ÝÝ»ñ|Î´ Ñ³ßí»ïíáõÃÛáõÝÝ»ñ|04_2 /4 Ñ³ßí»ïíáõÃÛ³Ý 2 ¿çÁ")
    
      dateS = "310116"
    
      'Լրացնում է  "Ամսաթիվ" դաշտը
      Call Rekvizit_Fill("Dialog",1 ,"General" ,"DATEREP" ,"^A[Del]" & dateS)
      'Սեղմել "Կատարել" կոճակը
      Call ClickCmdButton(2 ,"Î³ï³ñ»É")
    
      BuiltIn.Delay(8000)
    
      'Սեղմել "Հիշել որպես"
      Call wMainForm.MainMenu.Click(c_SaveAs)
      Sys.Process("Asbank").Window("#32770", "ÐÇß»É áñå»ë", 1).Window("DUIViewWndClassName", "", 1).Window("DirectUIHWND", "", 1).Window("FloatNotifySink", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1).Keys(Project.Path & "Stores\CB\Actual\04.txt")
      Sys.Process("Asbank").Window("#32770", "ÐÇß»É áñå»ë", 1).Window("Button", "&Save", 1).Click()
    
      'Համեմատել ֆայլերը
      file1 = Project.Path & "Stores\CB\Actual\04.txt"
      file2 = Project.Path & "Stores\CB\Expected\Expected 04.txt"
      Call Compare_Files(file1, file2,param)
    
      'Փակել ՀԾ - Բանկ համակարգը
      Call Close_AsBank()
    
End Sub