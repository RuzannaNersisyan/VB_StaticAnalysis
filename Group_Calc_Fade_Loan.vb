Option Explicit

'USEUNIT Library_Common
'USEUNIT Group_Operations_Library
'USEUNIT Constants
'USEUNIT Subsystems_SQL_Library

'Test Case Id 165866

Sub Group_Calc_Fade_Loan_Test()
  Utilities.ShortDateFormat = "yyyymmdd"
  Dim fDATE, sDATE 
  Dim CheckBox(0), queryString, sqlValue
  Dim time, timeStart, timeEnd
   
  ''1.Համակարգ մուտք գործել ARMSOFT օգտագործողով
  fDATE = "20240101"
  sDATE = "20140101"
  Call Initialize_AsBankQA(sDATE, fDATE)
  Login("ARMSOFT")
  Call Create_Connection()

  ''2.Անցում կատարել "Վարկեր (տեղաբաշխված)"
  Call ChangeWorkspace(c_Loans)
  
  ''3.Մուտք գործել "Պայմանագրեր" թղթապանակ 
  Call wTreeView.DblClickItem("|ì³ñÏ»ñ (ï»Õ³µ³ßËí³Í)|ä³ÛÙ³Ý³·ñ»ñ")
  
  ''4.Լրացնել "Պայմանագրեր" դիալոգային պատուհանը
		Call Rekvizit_Fill("Dialog", 1, "General", "LEVEL", 1)
		Call Rekvizit_Fill("Dialog", 1, "CheckBox", "CLOSE", 0)
		Call ClickCmdButton(2, "Î³ï³ñ»É")

  ''5.Խմբային Տոկոսների հաշվարկում
  CheckBox(0) = "CHG"
		timeStart = aqDateTime.Now 
  timeEnd = aqDateTime.Now 
		Call Group_Calculation("130314", CheckBox)
  time = aqDateTime.GetMinutes(aqDateTime.TimeInterval(timeStart, timeEnd))
  Log.Message("Հաշվարկման ժամանակը` " & time)
  
  ''6.Կատարել SQL ստուգումներ: 
  'Խմբային մարում գործողությունից հետո ստուգել տողերի քանակը
  queryString = "SELECT COUNT(*) FROM HI WHERE fDATE  = '2014-03-13' AND fTYPE = '01'"
  Call CheckDB_Value(queryString, 14460, 0)
      
  queryString = "SELECT COUNT(*) FROM HI WHERE fDATE  = '2014-03-13' AND fTYPE = '02'"
  Call CheckDB_Value(queryString, 11, 0)
      
  queryString = "SELECT COUNT(*) FROM HIR WHERE fDATE  = '2014-03-13'"
  Call CheckDB_Value(queryString, 7023, 0)
      
  queryString = "SELECT COUNT(*) FROM HIF WHERE fDATE  = '2014-03-13'" 
  Call CheckDB_Value(queryString, 3433, 0)

  queryString = "SELECT COUNT(*) FROM HIT WHERE fDATE  = '2014-03-13' "
  Call CheckDB_Value(queryString, 6925, 0)
      
  'Ստուգել գումարները
  queryString = "SELECT SUM(fSUM) FROM HI WHERE fDATE = '2014-03-13' AND fTYPE = '01'"
  Call CheckDB_Value(queryString, 4279602520.40, 0)
      
  queryString = "SELECT SUM(fSUM) FROM HI WHERE fDATE = '2014-03-13' AND fTYPE = '02'"
  Call CheckDB_Value(queryString, 125147985.90, 0)
  
  queryString = "SELECT SUM(fSUM) FROM HIF WHERE fDATE = '2014-03-13'" 
		sqlValue = CSTR(Get_Query_Result(queryString))
		if not sqlValue = "2000017100660.5111" then 
				Log.Error "Querystring = " & queryString & ":  Expected result = 2000017100660.5111, Query result = " & sqlValue, "", pmNormal, ErrorColor
		end if
		
  ''7.Վարկերում Խմբային պարտքերի մարում գործողության կատարում
  CheckBox(0) = "DBT"
  timeStart = aqDateTime.Now 
  timeEnd = aqDateTime.Now 
		Call Group_Calculation("140314", CheckBox)
  time = aqDateTime.GetMinutes(aqDateTime.TimeInterval(timeStart, timeEnd))
  Log.Message("Պարտքերի մարման ժամանակը` " & time)
  
		'8.Կատարել SQL ստուգումներ:
  'Խմբային մարում գործողությունից հետո ստուգել տողերի քանակը
  queryString = "SELECT COUNT(*) FROM HI WHERE fDATE  = '2014-03-14' AND fTYPE = '01'"
  Call CheckDB_Value(queryString, 44, 0)
      
  queryString = "SELECT COUNT(*) FROM HI WHERE fDATE  = '2014-03-14' AND fTYPE = '02'"
  Call CheckDB_Value(queryString, 2, 0)
      
  queryString = "SELECT COUNT(*) FROM HIR WHERE fDATE  = '2014-03-14' "
  Call CheckDB_Value(queryString, 200, 0)
      
  queryString = "SELECT COUNT(*) FROM HIF WHERE fDATE  = '2014-03-14' " 
  Call CheckDB_Value(queryString, 53, 0)
      
  'Ստուգել գումարները
  queryString = "SELECT SUM(fSUM) FROM HI WHERE fDATE = '2014-03-14' AND fTYPE = '01'"
  Call CheckDB_Value(queryString, 337720.80, 0)
      
  queryString = "SELECT SUM(fSUM) FROM HI WHERE fDATE = '2014-03-14' AND fTYPE = '02'"
  Call CheckDB_Value(queryString, 600000.00, 0)
     
  queryString = "SELECT SUM(fSUM) FROM HIF WHERE fDATE = '2014-03-14' " 
  Call CheckDB_Value(queryString, 600000.00, 0)
      
		'9.Դուրս գալ համակարգից:
		Call Close_AsBank()
End Sub
