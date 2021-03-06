Option Explicit

'USEUNIT Library_Common
'USEUNIT Group_Operations_Library
'USEUNIT Overdraft_NewCases_Library
'USEUNIT Constants
'USEUNIT Subsystems_SQL_Library
'USEUNIT Akreditiv_Library

'Test Case Id 165864

Sub Group_Calc_Fade_Overdraft_Test ()
  Utilities.ShortDateFormat = "yyyymmdd"
  Dim fDATE, sDATE 
  Dim Pttel, Typ, MesBox, FolderName, queryString, arrCheckbox
  Dim time, timeStart, timeEnd
  
  ''1.Համակարգ մուտք գործել ARMSOFT օգտագործողով
  fDATE = "20240101"
  sDATE = "20140101"
		Typ = "^A[Del]" & "[Tab]"
  MesBox = "1"
  Pttel = "_2"
		FolderName = "|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|"
  Call Initialize_AsBankQA(sDATE, fDATE)
  Login("ARMSOFT")
  Call Create_Connection()

  ''2.Անցում կատարել "Օվերդրաֆտ (տեղաբաշխված)"
  Call ChangeWorkspace(c_Overdraft)
  
  ''Ջնջել բոլոր գործողությունները
  Call wTreeView.DblClickItem("|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|¶áñÍáÕáõÃÛáõÝÝ»ñÇ ³Ù÷á÷áõÙ")
  Call Rekvizit_Fill("Dialog", 1, "General", "FDATE", "130314")
  Call Rekvizit_Fill("Dialog", 1, "General", "LDATE", "140314")
  Asbank.VBObject("frmAsUstPar").VBObject("CmdOK").ClickButton
  If wMDIClient.VBObject("frmPttel").VBObject("tdbgView").ApproxCount <> 0 Then
    wMDIClient.VBObject("frmPttel").Close
    Call LetterOfCredit_Filter_Fill(FolderName, 1, "TO2396")  
    Call DeleteActions("140314", Pttel, Typ, MesBox)
    wMDIClient.VBObject("frmPttel").Close
    Call LetterOfCredit_Filter_Fill(FolderName, 1, "TO5581")
    Call DeleteActions("130314", Pttel, Typ, MesBox)
    wMDIClient.VBObject("frmPttel").Close
  Else 
    wMDIClient.VBObject("frmPttel").Close  
  End If
  
  ''3.Մուտք գործել "Պայմանագրեր" թղթապանակ  
  Call wTreeView.DblClickItem("|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|ä³ÛÙ³Ý³·ñ»ñ")
		
  ''4.Լրացնել "Պայմանագրեր" դիալոգային պատուհանը
  Call Rekvizit_Fill("Dialog", 1, "General", "LEVEL", 1) 
  Call Rekvizit_Fill("Dialog", 1, "CheckBox", "CLOSE", 0) 
  Call ClickCmdButton(2, "Î³ï³ñ»É")
  
  ''5.Խմբային Տոկոսների հաշվարկում
  timeStart = aqDateTime.Now 
  ReDim arrCheckbox(2)          
  arrCheckbox = Array("CHG", "OPX")
  timeEnd = aqDateTime.Now 
		Call Group_Calculation("130314", arrCheckbox)
  time = aqDateTime.GetMinutes(aqDateTime.TimeInterval(timeStart, timeEnd))
  Log.Message("Հաշվարկման ժամանակը` " & time & " րոպե:")
		
		BuiltIn.delay(3000)
		wMDIClient.VBObject("frmPttel").Close
		
		''6.Կատարել SQL ստուգումներ:
  'Խմբային Տոկոսների հաշվարկում գործողությունից հետո ստուգել տողերի քանակը
  queryString = "SELECT COUNT(*) FROM HI WHERE fDATE  = '2014-03-13' AND fTYPE = '01'"
  Call CheckDB_Value(queryString, 15494, 0)
      
  queryString = "SELECT COUNT(*) FROM HI WHERE fDATE  = '2014-03-13' AND fTYPE = '02'"
  Call CheckDB_Value(queryString, 17, 0)
      
  BuiltIn.Delay(delay_middle) 
      
  queryString = "SELECT COUNT(*) FROM HIR WHERE fDATE  = '2014-03-13'"
  Call CheckDB_Value(queryString, 7303, 0)
      
  queryString = "SELECT COUNT(*) FROM HIF WHERE fDATE  = '2014-03-13'" 
  Call CheckDB_Value(queryString, 4053, 0)

  queryString = "SELECT COUNT(*) FROM HIT WHERE fDATE  = '2014-03-13'" 
  Call CheckDB_Value(queryString, 7315, 0)
      
  BuiltIn.Delay(delay_middle) 
      
  'Ստուգել գումարները
  queryString = "SELECT SUM(fSUM) FROM HI WHERE fDATE = '2014-03-13' AND fTYPE = '01'"
  Call CheckDB_Value(queryString, 4280262320.20, 0)
      
  BuiltIn.Delay(delay_middle) 
      
  queryString = "SELECT SUM(fSUM) FROM HI WHERE fDATE = '2014-03-13' AND fTYPE = '02'"
  Call CheckDB_Value(queryString, 125153900.30, 0)
          
  BuiltIn.Delay(delay_middle) 
      
  queryString = "SELECT round(SUM(fSUM),2)  FROM HIF WHERE fDATE = '2014-03-13'" 
  Call CheckDB_Value(queryString, 2000017100660.51, 0)
   
  ''Մուտք գործել "Օվերդրաֆտ (տեղաբաշխված)/Օվերդրաֆտ ունեցող հաշիվներ"   
  Call wTreeView.DblClickItem("|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|úí»ñ¹ñ³ýï áõÝ»óáÕ Ñ³ßÇíÝ»ñ")
  Call Rekvizit_Fill("Dialog", 1, "General", "ACCMASK", "^A[Del]")
  Call ClickCmdButton(2, "Î³ï³ñ»É")
      
  ''7.Օվերդրաֆտային պայմանագրի համար խմբային մարում
  timeStart = aqDateTime.Now
  timeEnd = aqDateTime.Now 
		Call Group_Payment("140314", "140314")
  time = aqDateTime.GetMinutes(aqDateTime.TimeInterval(timeStart, timeEnd))
  Log.Message("Պարտքերի մարման ժամանակը` " & time & " րոպե:")
		
  ''Փակել "Գերծախսի մարում" հաշվետվությունը
		BuiltIn.delay(3000)
  wMDIClient.VBObject("FrmSpr").Close 
      
  ''8.Կատարել SQL ստուգումներ:
  'Խմբային մարում գործողությունից հետո ստուգել տողերի քանակը
  queryString = "SELECT COUNT(*) FROM HI WHERE fDATE  = '2014-03-14' AND fTYPE = '01'"
  Call CheckDB_Value(queryString, 24, 0)
      
  queryString = "SELECT COUNT(*) FROM HI WHERE fDATE  = '2014-03-14' AND fTYPE = '02'"
  Call CheckDB_Value(queryString, 5, 0)
      
  queryString = "SELECT COUNT(*) FROM HIR WHERE fDATE  = '2014-03-14'" 
  Call CheckDB_Value(queryString, 171, 0)
      
  queryString = "SELECT COUNT(*) FROM HIF WHERE fDATE  = '2014-03-14'" 
  Call CheckDB_Value(queryString, 36, 0)
      
  'Ստուգել գումարները
  BuiltIn.Delay(delay_middle) 
  queryString = "SELECT SUM(fSUM) FROM HI WHERE fDATE  = '2014-03-14' AND fTYPE = '01'"
  Call CheckDB_Value(queryString, 1674246.8, 0)
      
  BuiltIn.Delay(delay_middle) 
      
  queryString = "SELECT SUM(fSUM) FROM HI WHERE fDATE  = '2014-03-14' AND fTYPE = '02'"
  Call CheckDB_Value(queryString, 1014979.1, 0)
     
  BuiltIn.Delay(delay_middle) 
      
  queryString = "SELECT SUM(fSUM) FROM HIF WHERE fDATE = '2014-03-14'" 
  Call CheckDB_Value(queryString, 600000.00, 0)

		''Փակել "Օվերդրաֆտ ունեցող հաշիվներ" պատուհանը
		wMDIClient.VBObject("frmPttel").Close
  
  FolderName = "|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|"
  Call LetterOfCredit_Filter_Fill(FolderName, 1, "TO2396")  

  Call DeleteActions("140314", Pttel, Typ, MesBox)
  wMDIClient.VBObject("frmPttel").Close
  
  Call LetterOfCredit_Filter_Fill(FolderName, 1, "TO5581")
  
  Call DeleteActions("130314", Pttel, Typ, MesBox)
		
  ''9.Դուրս գալ համակարգից:
  Call Close_AsBank()
End Sub