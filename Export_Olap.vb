'USEUNIT Library_Common
'USEUNIT OLAP_Library
'USEUNIT Constants

Sub Eport_OLAP_Test()
  
    Dim DateS,DateE,expOlap,expTXT,groupName,DateStart,DateEnd
    
    DateStart = "20120101"
    DateEnd = "20220101"
    groupName = "BAZEL CONTRACTS"
    DateS = "010214"
    DateE = "280214"   
    expOlap = 1
    expTXT = 0
    Call Initialize_AsBankQA(DateStart, DateEnd) 
 
    Call ChangeWorkspace(c_OLAPAdmin)
    Call wTreeView.DblClickItem("|OLAP ³¹ÙÇÝÇëïñ³ïáñÇ ²Þî|OLAP ËÙµ»ñÇ ï»Õ»Ï³ïáõ")
    aqPerformance.Start()
    Call Find_Group_and_Export(groupName,DateS,DateE,expOlap,expTXT)
    Log.Message(aqPerformance.Value)
    
    TestedApps.killproc.Run()
    
    Call Initialize_AsBankQA(DateStart, DateEnd) 
    Call ChangeWorkspace(c_OLAPAdmin)
    Call wTreeView.DblClickItem("|OLAP ³¹ÙÇÝÇëïñ³ïáñÇ ²Þî|OLAP ËÙµ»ñÇ ï»Õ»Ï³ïáõ")
    groupName = "COA1"
    DateE = "310116"
    aqPerformance.Start()
    Call Find_Group_and_Export(groupName,DateS,DateE,expOlap,expTXT)
    Log.Message(aqPerformance.Value)
    Sys.Process("Asbank").Window("ThunderRT6MDIForm", "ՀԾ-Բանկ 4.0 (bankTesting_QA)", 1).Window("MDIClient", "", 1).VBObject("frmPttel").Close()

    TestedApps.killproc.Run()    
    
    Call Initialize_AsBankQA(DateStart, DateEnd) 
    Call ChangeWorkspace(c_OLAPAdmin)
    Call wTreeView.DblClickItem("|OLAP ³¹ÙÇÝÇëïñ³ïáñÇ ²Þî|OLAP ËÙµ»ñÇ ï»Õ»Ï³ïáõ")
    aqPerformance.Start()
    groupName = "BNKLOANS"
    Call Find_Group_and_Export(groupName,DateS,DateE,expOlap,expTXT)
    Log.Message(aqPerformance.Value)
    Sys.Process("Asbank").Window("ThunderRT6MDIForm", "ՀԾ-Բանկ 4.0 (bankTesting_QA)", 1).Window("MDIClient", "", 1).VBObject("frmPttel").Close()
    'Սեղմել Կատարել կոճակը
    
    TestedApps.killproc.Run()

    DateStart = "20120101"
    DateEnd = "20200101"
    groupName = "LOWRISK_DEPOSITS"
    DateS = "310116"
    DateE = "310116"   
    expOlap = 1
    expTXT = 0
    
    Call Initialize_AsBankQA(DateStart, DateEnd) 
    Call ChangeWorkspace(c_OLAPAdmin)
    Call wTreeView.DblClickItem("|OLAP ³¹ÙÇÝÇëïñ³ïáñÇ ²Þî|OLAP ËÙµ»ñÇ ï»Õ»Ï³ïáõ")
    aqPerformance.Start()
    Call Find_Group_and_Export(groupName,DateS,DateE,expOlap,expTXT)
    Log.Message(aqPerformance.Value)
    Sys.Process("Asbank").Window("ThunderRT6MDIForm", "ՀԾ-Բանկ 4.0 (bankTesting_QA)", 1).Window("MDIClient", "", 1).VBObject("frmPttel").Close()
    'Սեղմել Կատարել կոճակը
    
    TestedApps.killproc.Run()
  
    DateStart = "20120101"
    DateEnd = "20200101"
    groupName = "SS14"
    DateS = "280214"
    DateE = "280214"   
    expOlap = 1
    expTXT = 0
    
    Call Initialize_AsBankQA(DateStart, DateEnd) 
    Call ChangeWorkspace(c_OLAPAdmin)
    Call wTreeView.DblClickItem("|OLAP ³¹ÙÇÝÇëïñ³ïáñÇ ²Þî|OLAP ËÙµ»ñÇ ï»Õ»Ï³ïáõ")
    aqPerformance.Start()
    Call Find_Group_and_Export(groupName,DateS,DateE,expOlap,expTXT)
    Log.Message(aqPerformance.Value)
    Sys.Process("Asbank").Window("ThunderRT6MDIForm", "ՀԾ-Բանկ 4.0 (bankTesting_QA)", 1).Window("MDIClient", "", 1).VBObject("frmPttel").Close()
    'Սեղմել Կատարել կոճակը
    
    TestedApps.killproc.Run()
    
    Call Initialize_AsBankQA(DateStart, DateEnd) 
    Call ChangeWorkspace(c_OLAPAdmin)
    Call wTreeView.DblClickItem("|OLAP ³¹ÙÇÝÇëïñ³ïáñÇ ²Þî|OLAP ËÙµ»ñÇ ï»Õ»Ï³ïáõ")
    groupName = "IncExpNW"
    aqPerformance.Start()
    Call Find_Group_and_Export(groupName,DateS,DateE,expOlap,expTXT)
    Log.Message(aqPerformance.Value)
    Sys.Process("Asbank").Window("ThunderRT6MDIForm", "ՀԾ-Բանկ 4.0 (bankTesting_QA)", 1).Window("MDIClient", "", 1).VBObject("frmPttel").Close()
    'Սեղմել Կատարել կոճակը

    TestedApps.killproc.Run()
    
    Call Initialize_AsBankQA(DateStart, DateEnd) 
    Call ChangeWorkspace(c_OLAPAdmin)
    Call wTreeView.DblClickItem("|OLAP ³¹ÙÇÝÇëïñ³ïáñÇ ²Þî|OLAP ËÙµ»ñÇ ï»Õ»Ï³ïáõ")
    aqPerformance.Start()
    groupName = "FORM6"
    Call Find_Group_and_Export(groupName,DateS,DateE,expOlap,expTXT)
    Log.Message(aqPerformance.Value)
    Sys.Process("Asbank").Window("ThunderRT6MDIForm", "ՀԾ-Բանկ 4.0 (bankTesting_QA)", 1).Window("MDIClient", "", 1).VBObject("frmPttel").Close()
    'Սեղմել Կատարել կոճակը
    
    TestedApps.killproc.Run()

    Call Initialize_AsBankQA(DateStart, DateEnd) 
    Call ChangeWorkspace(c_OLAPAdmin)
    Call wTreeView.DblClickItem("|OLAP ³¹ÙÇÝÇëïñ³ïáñÇ ²Þî|OLAP ËÙµ»ñÇ ï»Õ»Ï³ïáõ")
    aqPerformance.Start()
    groupName = "FORM9"
    Call Find_Group_and_Export(groupName,DateS,DateE,expOlap,expTXT)
    Log.Message(aqPerformance.Value)
    Sys.Process("Asbank").Window("ThunderRT6MDIForm", "ՀԾ-Բանկ 4.0 (bankTesting_QA)", 1).Window("MDIClient", "", 1).VBObject("frmPttel").Close()
    'Սեղմել Կատարել կոճակը
    
    TestedApps.killproc.Run()
    
     DateStart = "20120101"
    DateEnd = "20200101"
    groupName = "CashOtc"
    DateS = "010214"
    DateE = "280214"   
    expOlap = 1
    expTXT = 0 
    
    Call Initialize_AsBankQA(DateStart, DateEnd) 
    Call ChangeWorkspace(c_OLAPAdmin)
    Call wTreeView.DblClickItem("|OLAP ³¹ÙÇÝÇëïñ³ïáñÇ ²Þî|OLAP ËÙµ»ñÇ ï»Õ»Ï³ïáõ")
    aqPerformance.Start()
    Call Find_Group_and_Export(groupName,DateS,DateE,expOlap,expTXT)
    Log.Message(aqPerformance.Value)
    Sys.Process("Asbank").Window("ThunderRT6MDIForm", "ՀԾ-Բանկ 4.0 (bankTesting_QA)", 1).Window("MDIClient", "", 1).VBObject("frmPttel").Close()
    'Սեղմել Կատարել կոճակը
    
    TestedApps.killproc.Run()
    
    
    DateStart = "20120101"
    DateEnd = "20200101"
    groupName = "SS14"
    DateS = "280214"
    DateE = "280214"   
    expOlap = 1
    expTXT = 0
    
    Call Initialize_AsBankQA(DateStart, DateEnd) 
    Call ChangeWorkspace(c_OLAPAdmin)
    Call wTreeView.DblClickItem("|OLAP ³¹ÙÇÝÇëïñ³ïáñÇ ²Þî|OLAP ËÙµ»ñÇ ï»Õ»Ï³ïáõ")
    aqPerformance.Start()
    Call Find_Group_and_Export(groupName,DateS,DateE,expOlap,expTXT)
    Log.Message(aqPerformance.Value)
    Sys.Process("Asbank").Window("ThunderRT6MDIForm", "ՀԾ-Բանկ 4.0 (bankTesting_QA)", 1).Window("MDIClient", "", 1).VBObject("frmPttel").Close()
    'Սեղմել Կատարել կոճակը
    
    TestedApps.killproc.Run()

    Call Initialize_AsBankQA(DateStart, DateEnd) 
    Call ChangeWorkspace(c_OLAPAdmin)
    Call wTreeView.DblClickItem("|OLAP ³¹ÙÇÝÇëïñ³ïáñÇ ²Þî|OLAP ËÙµ»ñÇ ï»Õ»Ï³ïáõ")
    aqPerformance.Start()
    groupName = "FORM15"
    Call Find_Group_and_Export(groupName,DateS,DateE,expOlap,expTXT)
    Log.Message(aqPerformance.Value)
    Sys.Process("Asbank").Window("ThunderRT6MDIForm", "ՀԾ-Բանկ 4.0 (bankTesting_QA)", 1).Window("MDIClient", "", 1).VBObject("frmPttel").Close()
    'Սեղմել Կատարել կոճակը
    
    TestedApps.killproc.Run()
    
     DateStart = "20120101"
    DateEnd = "20200101"
    groupName = "DiAvPe"
    DateS = "010214"
    DateE = "280214"   
    expOlap = 1
    expTXT = 0
    Call Initialize_AsBankQA(DateStart, DateEnd) 
     'Test StartUp end
 
    Call ChangeWorkspace(c_OLAPAdmin)
    Call wTreeView.DblClickItem("|OLAP ³¹ÙÇÝÇëïñ³ïáñÇ ²Þî|OLAP ËÙµ»ñÇ ï»Õ»Ï³ïáõ")
    aqPerformance.Start()
    Call Find_Group_and_Export(groupName,DateS,DateE,expOlap,expTXT)
    Log.Message(aqPerformance.Value)
    
    TestedApps.killproc.Run()
    
    Call Initialize_AsBankQA(DateStart, DateEnd) 
    Call ChangeWorkspace(c_OLAPAdmin)
    Call wTreeView.DblClickItem("|OLAP ³¹ÙÇÝÇëïñ³ïáñÇ ²Þî|OLAP ËÙµ»ñÇ ï»Õ»Ï³ïáõ")
    groupName = "DiAvPe2"
    aqPerformance.Start()
    Call Find_Group_and_Export(groupName,DateE,DateE,expOlap,expTXT)
    Log.Message(aqPerformance.Value)
    Sys.Process("Asbank").Window("ThunderRT6MDIForm", "ՀԾ-Բանկ 4.0 (bankTesting_QA)", 1).Window("MDIClient", "", 1).VBObject("frmPttel").Close()
    'Սեղմել Կատարել կոճակը

    TestedApps.killproc.Run()
    
     DateStart = "20120101"
    DateEnd = "20200101"
    groupName = "FORM18_1"
    DateS = "310314"
    DateE = "310314"   
    expOlap = 1
    expTXT = 0
    Call Initialize_AsBankQA(DateStart, DateEnd) 
     'Test StartUp end
 
    Call ChangeWorkspace(c_OLAPAdmin)
    Call wTreeView.DblClickItem("|OLAP ³¹ÙÇÝÇëïñ³ïáñÇ ²Þî|OLAP ËÙµ»ñÇ ï»Õ»Ï³ïáõ")
    aqPerformance.Start()
    Call Find_Group_and_Export(groupName,DateS,DateE,expOlap,expTXT)
    Log.Message(aqPerformance.Value)
    
    TestedApps.killproc.Run()
    
    Call Initialize_AsBankQA(DateStart, DateEnd) 
    Call ChangeWorkspace(c_OLAPAdmin)
    Call wTreeView.DblClickItem("|OLAP ³¹ÙÇÝÇëïñ³ïáñÇ ²Þî|OLAP ËÙµ»ñÇ ï»Õ»Ï³ïáõ")
    groupName = "FORM18_2"
    aqPerformance.Start()
    Call Find_Group_and_Export(groupName,DateS,DateE,expOlap,expTXT)
    Log.Message(aqPerformance.Value)
    Sys.Process("Asbank").Window("ThunderRT6MDIForm", "ՀԾ-Բանկ 4.0 (bankTesting_QA)", 1).Window("MDIClient", "", 1).VBObject("frmPttel").Close()
    'Սեղմել Կատարել կոճակը

    TestedApps.killproc.Run()
    
    Call Initialize_AsBankQA(DateStart, DateEnd) 
     'Test StartUp end
 
    Call ChangeWorkspace(c_OLAPAdmin)
    Call wTreeView.DblClickItem("|OLAP ³¹ÙÇÝÇëïñ³ïáñÇ ²Þî|OLAP ËÙµ»ñÇ ï»Õ»Ï³ïáõ")
    groupName = "FORM18_4"
    aqPerformance.Start()
    Call Find_Group_and_Export(groupName,DateS,DateE,expOlap,expTXT)
    Log.Message(aqPerformance.Value)
    Sys.Process("Asbank").Window("ThunderRT6MDIForm", "ՀԾ-Բանկ 4.0 (bankTesting_QA)", 1).Window("MDIClient", "", 1).VBObject("frmPttel").Close()
    TestedApps.killproc.Run()
    
     DateStart = "20120101"
    DateEnd = "20200101"
    groupName = "CuPuSa1"
    DateS = "280214"
    DateE = "280214"   
    expOlap = 1
    expTXT = 0
    Call Initialize_AsBankQA(DateStart, DateEnd) 
     'Test StartUp end
 
    Call ChangeWorkspace(c_OLAPAdmin)
    Call wTreeView.DblClickItem("|OLAP ³¹ÙÇÝÇëïñ³ïáñÇ ²Þî|OLAP ËÙµ»ñÇ ï»Õ»Ï³ïáõ")
    aqPerformance.Start()
    Call Find_Group_and_Export(groupName,DateS,DateE,expOlap,expTXT)
    Log.Message(aqPerformance.Value)
    
    TestedApps.killproc.Run()
    
    Call Initialize_AsBankQA(DateStart, DateEnd) 
    Call ChangeWorkspace(c_OLAPAdmin)
    Call wTreeView.DblClickItem("|OLAP ³¹ÙÇÝÇëïñ³ïáñÇ ²Þî|OLAP ËÙµ»ñÇ ï»Õ»Ï³ïáõ")
    groupName = "CuPuSa2"
    aqPerformance.Start()
    Call Find_Group_and_Export(groupName,DateS,DateE,expOlap,expTXT)
    Log.Message(aqPerformance.Value)
    Sys.Process("Asbank").Window("ThunderRT6MDIForm", "ՀԾ-Բանկ 4.0 (bankTesting_QA)", 1).Window("MDIClient", "", 1).VBObject("frmPttel").Close()
    'Սեղմել Կատարել կոճակը

    TestedApps.killproc.Run()
    
    
    DateStart = "20120101"
    DateEnd = "20200101"
    groupName = "CliAcnt"
    DateS = "010214"
    DateE = "280214"   
    expOlap = 1
    expTXT = 0
    Call Initialize_AsBankQA(DateStart, DateEnd) 
     'Test StartUp end
 
    Call ChangeWorkspace(c_OLAPAdmin)
    Call wTreeView.DblClickItem("|OLAP ³¹ÙÇÝÇëïñ³ïáñÇ ²Þî|OLAP ËÙµ»ñÇ ï»Õ»Ï³ïáõ")
    aqPerformance.Start()
    Call Find_Group_and_Export(groupName,DateS,DateE,expOlap,expTXT)
    Log.Message(aqPerformance.Value)
    
    TestedApps.killproc.Run()

    Call Initialize_AsBankQA(DateStart, DateEnd) 
    Call ChangeWorkspace(c_OLAPAdmin)
    Call wTreeView.DblClickItem("|OLAP ³¹ÙÇÝÇëïñ³ïáñÇ ²Þî|OLAP ËÙµ»ñÇ ï»Õ»Ï³ïáõ")
    aqPerformance.Start()
    groupName = "FORM24"
    Call Find_Group_and_Export(groupName,DateS,DateE,expOlap,expTXT)
    Log.Message(aqPerformance.Value)
    Sys.Process("Asbank").Window("ThunderRT6MDIForm", "ՀԾ-Բանկ 4.0 (bankTesting_QA)", 1).Window("MDIClient", "", 1).VBObject("frmPttel").Close()
    'Սեղմել Կատարել կոճակը
    
    TestedApps.killproc.Run()
   
    DateStart = "20120101"
    DateEnd = "20200101"
    groupName = "FORM26"
    DateS = "280214"
    DateE = "280214"   
    expOlap = 1
    expTXT = 0
    
    Call Initialize_AsBankQA(DateStart, DateEnd) 
    Call ChangeWorkspace(c_OLAPAdmin)
    Call wTreeView.DblClickItem("|OLAP ³¹ÙÇÝÇëïñ³ïáñÇ ²Þî|OLAP ËÙµ»ñÇ ï»Õ»Ï³ïáõ")
    aqPerformance.Start()
    Call Find_Group_and_Export(groupName,DateS,DateE,expOlap,expTXT)
    Log.Message(aqPerformance.Value)
    Sys.Process("Asbank").Window("ThunderRT6MDIForm", "ՀԾ-Բանկ 4.0 (bankTesting_QA)", 1).Window("MDIClient", "", 1).VBObject("frmPttel").Close()
    'Սեղմել Կատարել կոճակը
    
    TestedApps.killproc.Run()
    
    Call Initialize_AsBankQA(DateStart, DateEnd) 
    Call ChangeWorkspace(c_OLAPAdmin)
    Call wTreeView.DblClickItem("|OLAP ³¹ÙÇÝÇëïñ³ïáñÇ ²Þî|OLAP ËÙµ»ñÇ ï»Õ»Ï³ïáõ")
    groupName = "FORM26PRC"
    DateS = "011213"
    aqPerformance.Start()
    Call Find_Group_and_Export(groupName,DateS,DateE,expOlap,expTXT)
    Log.Message(aqPerformance.Value)
    Sys.Process("Asbank").Window("ThunderRT6MDIForm", "ՀԾ-Բանկ 4.0 (bankTesting_QA)", 1).Window("MDIClient", "", 1).VBObject("frmPttel").Close()
    'Սեղմել Կատարել կոճակը

    TestedApps.killproc.Run()
    
     DateStart = "20120101"
    DateEnd = "20200101"
    groupName = "SOUSHI_Assets"
    DateS = "280214"
    DateE = "280214"   
    expOlap = 1
    expTXT = 0
    Call Initialize_AsBankQA(DateStart, DateEnd) 
     'Test StartUp end
 
    Call ChangeWorkspace(c_OLAPAdmin)
    Call wTreeView.DblClickItem("|OLAP ³¹ÙÇÝÇëïñ³ïáñÇ ²Þî|OLAP ËÙµ»ñÇ ï»Õ»Ï³ïáõ")
    aqPerformance.Start()
    Call Find_Group_and_Export(groupName,DateS,DateE,expOlap,expTXT)
    Log.Message(aqPerformance.Value)
    
    TestedApps.killproc.Run()
    
    Call Initialize_AsBankQA(DateStart, DateEnd) 
    Call ChangeWorkspace(c_OLAPAdmin)
    Call wTreeView.DblClickItem("|OLAP ³¹ÙÇÝÇëïñ³ïáñÇ ²Þî|OLAP ËÙµ»ñÇ ï»Õ»Ï³ïáõ")
    groupName = "SOUSHI_Liab"
    aqPerformance.Start()
    Call Find_Group_and_Export(groupName,DateS,DateE,expOlap,expTXT)
    Log.Message(aqPerformance.Value)
    Sys.Process("Asbank").Window("ThunderRT6MDIForm", "ՀԾ-Բանկ 4.0 (bankTesting_QA)", 1).Window("MDIClient", "", 1).VBObject("frmPttel").Close()
    'Սեղմել Կատարել կոճակը

    TestedApps.killproc.Run()

    DateStart = "20120101"
    DateEnd = "20200101"
    groupName = "SOUSHI_Assets_Flow"
    DateS = "010214"
    DateE = "280214"   
    expOlap = 1
    expTXT = 0
    Call Initialize_AsBankQA(DateStart, DateEnd) 
     'Test StartUp end
 
    Call ChangeWorkspace(c_OLAPAdmin)
    Call wTreeView.DblClickItem("|OLAP ³¹ÙÇÝÇëïñ³ïáñÇ ²Þî|OLAP ËÙµ»ñÇ ï»Õ»Ï³ïáõ")
    aqPerformance.Start()
    Call Find_Group_and_Export(groupName,DateS,DateE,expOlap,expTXT)
    Log.Message(aqPerformance.Value)
    
    TestedApps.killproc.Run()

    Call Initialize_AsBankQA(DateStart, DateEnd) 
     'Test StartUp end
 
    Call ChangeWorkspace(c_OLAPAdmin)
    Call wTreeView.DblClickItem("|OLAP ³¹ÙÇÝÇëïñ³ïáñÇ ²Þî|OLAP ËÙµ»ñÇ ï»Õ»Ï³ïáõ")
    aqPerformance.Start()
    groupName = "SOUSHI_Liab_Flow"
    Call Find_Group_and_Export(groupName,DateS,DateE,expOlap,expTXT)
    Log.Message(aqPerformance.Value)
    
    TestedApps.killproc.Run()
    
    

End Sub  