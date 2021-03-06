'USEUNIT SWIFT_International_Payorder_Library
'USEUNIT Library_Common
'USEUNIT OLAP_Library
'USEUNIT Constants

Option Explicit
'Test Case ID 121735

Sub Report_01_Test()
  
      Dim coaNum,dateE,repay,trans,cbCode,DateStart,DateEnd,branch,nbTurn
      Dim file1, file2, param
    
      DateStart = "20120101"
      DateEnd = "20240101" 

      'Մուտք գործել ՀԾ- Բանկ համակարգ ARMSOFT օգտագործողով
      Call Initialize_AsBankQA(DateStart, DateEnd) 
 
      'Անցում կատարել "Ենթահամակարգեր" ԱՇՏ
      Call ChangeWorkspace(c_Subsystems)
      Call wTreeView.DblClickItem("|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|Ð³ßí»ïíáõÃÛáõÝÝ»ñ,  Ù³ïÛ³ÝÝ»ñ|Î´ Ñ³ßí»ïíáõÃÛáõÝÝ»ñ|01 Ð³ßí»ÏßÇé")
    
      coaNum = "1"
      dateE = "280214"  
      cbCode = "99997"
      repay = 0
      trans = 0
      nbTurn = 1
    
      'Լրացնում է  "Ամսաթիվ" դաշտը
      Call Rekvizit_Fill("Dialog",1 ,"General" ,"DATE" ,dateE)
      'Լրացնում է "Հաշվային պլանի համար" դաշտը
      Call Rekvizit_Fill("Dialog",1 ,"General" ,"COANUM" ,coaNum)
      'Լրացնում է "Գրասենյակ" դաշտը
      Call Rekvizit_Fill("Dialog",1 ,"General" ,"BRANCH" ,branch)
      'Լրացնում է "ՀՀ ԿԲ կոդ" դաշտը
      Call Rekvizit_Fill("Dialog",1 ,"General" ,"CBCODE" ,cbCode)
      'Լրացնում է "Մարել մարման ենթակա հաշիվերը" նշիչը
      Call Rekvizit_Fill("Dialog",1 ,"CheckBox" ,"REPAY" ,repay)
      'Լրացնում է "Տեղափոխել տեղափոխման ենթակա հաշիվները" նշիչը
      Call Rekvizit_Fill("Dialog",1 ,"CheckBox" ,"TRANS" ,trans)
      'Լրացնում է "Ետհաշվեկշիռ" նշիչը
      Call Rekvizit_Fill("Dialog",1 ,"CheckBox" ,"NBBAL" ,nbTurn)   'NBBAL
      'Սեղմել "Կատարել" կոճակը
      Call ClickCmdButton(2 ,"Î³ï³ñ»É")
    
      BuiltIn.Delay(400000)
    
      'Սեղմել "Հիշել որպես"
      Call wMainForm.MainMenu.Click(c_SaveAs)
      Sys.Process("Asbank").Window("#32770", "ÐÇß»É áñå»ë", 1).Window("DUIViewWndClassName", "", 1).Window("DirectUIHWND", "", 1).Window("FloatNotifySink", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1).Keys(Project.Path & "Stores\CB\Actual\01.txt")
      Sys.Process("Asbank").Window("#32770", "ÐÇß»É áñå»ë", 1).Window("Button", "&Save", 1).Click()
    
      'Համեմատել ֆայլերը 
      file1 = Project.Path & "Stores\CB\Actual\01.txt"
      file2 = Project.Path & "Stores\CB\Expected\Expected 01.txt"
      Call Compare_Files(file1, file2, param)
    
      'Փակել ՀԾ- Բանկ համակարգը
      Call Close_AsBank()
    
End Sub