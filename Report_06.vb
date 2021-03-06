'USEUNIT SWIFT_International_Payorder_Library
'USEUNIT Library_Common
'USEUNIT OLAP_Library
'USEUNIT Constants

Option Explicit
'Test Case ID 122677

Sub Report_06_Test()
  
      Dim coaNum,dateE,DateStart,DateEnd,branch,nbTurn
      Dim file1, file2, param
    
      DateStart = "20120101"
      DateEnd = "20240101" 

      'Մուտք գործել ՀԾ- Բանկ համակարգ ARMSOFT օգտագործողով
      Call Initialize_AsBankQA(DateStart, DateEnd) 
 
      'Անցում կատարել "Ենթահամակարգեր" ԱՇՏ
      Call ChangeWorkspace(c_Subsystems)
      Call wTreeView.DblClickItem("|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|Ð³ßí»ïíáõÃÛáõÝÝ»ñ,  Ù³ïÛ³ÝÝ»ñ|Î´ Ñ³ßí»ïíáõÃÛáõÝÝ»ñ|06 ì³ñÏ³ÛÇÝ Ý»ñ¹ñáõÙÝ»ñÇ, ¹»µÇïáñ. å³ñïù. ¨ ÇÝí»ë.³ñÅ»Ã.")
    
      dateE = "280214"
      coaNum = 0 
      nbTurn = 1
    
      'Լրացնում է  "Ամսաթիվ" դաշտը
      Call Rekvizit_Fill("Dialog",1 ,"General" ,"DATEREP" ,dateE)
      'Լրացնում է "Գրասենյակ" դաշտը
      Call Rekvizit_Fill("Dialog",1 ,"General" ,"BRANCH" ,branch)
      'Լրացնում է "Հաշվել ատոմար ցուցանիշները" նշիչը
      Call Rekvizit_Fill("Dialog",1 ,"CheckBox" ,"CALCATMCOD" ,coaNum)
      'Լրացնում է " Ցույց տալ բաղադրիչ գումարները" նշիչը
      Call Rekvizit_Fill("Dialog",1 ,"CheckBox" ,"SHOWCOMMENT" ,coaNum)
      'Լրացնում է "Վարկային ներդրումներ" նշիչը
      Call Rekvizit_Fill("Dialog",1 ,"CheckBox" ,"CALC1" ,nbTurn)    
      'Լրացնում է "Վարկերի պահուստի չափը" նշիչը
      Call Rekvizit_Fill("Dialog",1 ,"CheckBox" ,"CALC2" ,nbTurn)
      'Լրացնում է "Դեբ, պարտքերի պահուստի չափը"  նշիչը
      Call Rekvizit_Fill("Dialog",1 ,"CheckBox" ,"CALC3" ,nbTurn)
      'Լրացնում է "Թղթակցային հաշիվեր" նշիչը 
      Call Rekvizit_Fill("Dialog",1 ,"CheckBox" ,"CALC7" ,nbTurn)
      'Լրացնում է "Հետհաշվ . հոդվածների պահուստի չափ"  նշիչը
      Call Rekvizit_Fill("Dialog",1 ,"CheckBox" ,"CALC4" ,nbTurn)
      'Լրացնում է "Ինվ. արժ. հնարավոր կորուստբերի պահուստ" նշիչը
      Call Rekvizit_Fill("Dialog",1 ,"CheckBox" ,"CALC5" ,nbTurn)
      'Լրացնում է "Չաշխատող և հետհ. դուրս գրված ակտիվներ"  նշիչը
      Call Rekvizit_Fill("Dialog",1 ,"CheckBox" ,"CALC6" ,nbTurn)
      'Սեղմել "Կատարել" կոճակը
      Call ClickCmdButton(2 ,"Î³ï³ñ»É")
    
      BuiltIn.Delay(900000)
    
      'Սեղմել "Հիշել որպես"
      Call wMainForm.MainMenu.Click(c_SaveAs)
      Sys.Process("Asbank").Window("#32770", "ÐÇß»É áñå»ë", 1).Window("DUIViewWndClassName", "", 1).Window("DirectUIHWND", "", 1).Window("FloatNotifySink", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1).Keys(Project.Path & "Stores\CB\Actual\06.txt")
      Sys.Process("Asbank").Window("#32770", "ÐÇß»É áñå»ë", 1).Window("Button", "&Save", 1).Click()
    
      'Համեմատել ֆայլերը
      file1 = Project.Path & "Stores\CB\Actual\06.txt"
      file2 = Project.Path & "Stores\CB\Expected\Expected 06.txt"
      Call Compare_Files(file1, file2,param)
    
      'Փայկել ՀԾ - Բանկ համակարգը
      Call Close_AsBank()
    
End Sub