'USEUNIT SWIFT_International_Payorder_Library
'USEUNIT Library_Common
'USEUNIT OLAP_Library
'USEUNIT Constants
'USEUNIT Library_Colour
'USEUNIT Payment_Except_Library

Option Explicit
'Test Case ID 123849

Sub Report_24_Test()
  
      Dim dateS, dateE, DateStart, DateEnd, groupName, expOlap, expTXT
      Dim startDate, endDate, foreignTransacts, internalTransacts
      Dim actualFilePath, actualFile, expectedFile
   
      DateStart = "20120101"
      DateEnd = "20240101" 

      'Մուտք գործել ՀԾ- Բանկ համակարգ ARMSOFT օգտագործողով
      Call Initialize_AsBankQA(DateStart, DateEnd) 
      Call ChangeWorkspace(c_Admin40)
    
      'Աշխատեցնել UpdateDataForRep24 ֆունկցիան
      Call wTreeView.DblClickItem("|²¹ÙÇÝÇëïñ³ïáñÇ ²Þî 4.0|Ð³Ù³Ï³ñ·³ÛÇÝ ³ßË³ï³ÝùÝ»ñ|Ð³Ù³Ï³ñ·³ÛÇÝ ·áñÍÇùÝ»ñ|üáõÝÏóÇ³ÛÇ Ï³Ýã")
      Call Rekvizit_Fill("Dialog",1,"General","MODULENAME","Util")
      Call Rekvizit_Fill("Dialog",1,"General","SUBNAME","UpdateDataForRep24")
      Call ClickCmdButton(2,"Î³ï³ñ»É")
      Call Rekvizit_Fill("Dialog",1,"General","SDATE","010214")
      Call ClickCmdButton(2,"Î³ï³ñ»É")
      BuiltIn.Delay(5000)
      Call ClickCmdButton(5,"OK")
    
      'Արտահանել 24-րդ հաժվետվությունը
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
      Call ClickCmdButton(5,"OK")
      Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("FrmSpr").Close()
      Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").Close()
    
 
      'Անցում կատարել "Ենթահամակարգեր" ԱՇՏ
      Call ChangeWorkspace(c_Subsystems)
      Call wTreeView.DblClickItem("|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|Ð³ßí»ïíáõÃÛáõÝÝ»ñ,  Ù³ïÛ³ÝÝ»ñ|Î´ Ñ³ßí»ïíáõÃÛáõÝÝ»ñ|24 ´³ÝÏ»ñÇ ÙÇçáóáí ³ñï»ñÏñÇó Ùáõïù »Õ³Í ¨ ³ñï»ñÏÇñ áõÕ³ñÏí³Í ·áõÙ³ñÝ»ñÇ í»ñ³µ»ñÛ³É(ÑÇÝ)")
      dateS = "280214"
    
      'Լրացնում է  "Ժամանակահատված(սկիզբ)" դաշտը
      Call Rekvizit_Fill("Dialog",1 ,"General" ,"SDATE" ,dateS)
      'Լրացնում է  "Ժամանակահատված(վերջ)" դաշտը
      Call Rekvizit_Fill("Dialog",1 ,"General" ,"EDATE" ,dateS)
      'Սեղմել "Կատարել" կոճակը
      Call ClickCmdButton(2 ,"Î³ï³ñ»É")
    
      BuiltIn.Delay(8000)
    
      '24 հաշվետվության պահպանում
      Log.Message "24 հաշվետվության պահպանում", "", pmNormal, DivideColor
      actualFilePath = Project.Path & "Stores\CB\Actual\"
      actualFile = "24.txt"
      expectedFile = Project.Path & "Stores\CB\Expected\Expected 24.txt"

      Call SaveDoc(actualFilePath, actualFile) 
    
     ' Համեմատել ֆայլերը
      Log.Message "Համեմատել ֆայլերը", "", pmNormal, DivideColor
      Call Compare_Files(actualFilePath & actualFile, expectedFile, "")
      wMDIClient.VBObject("FrmSpr").Close()
      
      'Մուտք գործել 24 նոր հաշվետվություն նոր ամսաթվերով 
      Log.Message "Մուտք գործել 24-րդ հաշվետվություն", "", pmNormal, DivideColor
      startDate = "280214"
      endDate = "280214"
      foreignTransacts = 1
      internalTransacts = 0
      Call GoTo_Report24_New(startDate, endDate, foreignTransacts, internalTransacts)
    
      '24 հաշվետվության պահպանում
      Log.Message "2-րդ 24 հաշվետվության պահպանում", "", pmNormal, DivideColor
      actualFilePath = Project.Path & "Stores\CB\Actual\"
      actualFile = "24_2.txt"
      expectedFile = Project.Path & "Stores\CB\Expected\Expected_24_2.txt"

      Call SaveDoc(actualFilePath, actualFile) 
    
     ' Համեմատել ֆայլերը
      Log.Message "Համեմատել ֆայլերը", "", pmNormal, DivideColor
      Call Compare_Files(actualFilePath & actualFile, expectedFile, "")
      wMDIClient.VBObject("FrmSpr").Close()  
      'Փակել ՀԾ - Բանկ համակարգը
      Call Close_AsBank()
    
End Sub

