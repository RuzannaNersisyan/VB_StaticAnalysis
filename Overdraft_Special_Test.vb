Option Explicit
'USEUNIT Library_Common
'USEUNIT Constants
'USEUNIT Subsystems_SQL_Library
'USEUNIT Subsystems_Special_Library
'USEUNIT Payment_Order_ConfirmPhases_Library

' Test Case Id 166636

Sub Overdraft_Special_Test()

      Dim fDATE, sDATE, confPath, confInput
        
      Dim  contType, fISN, docNum, clientCode, agType, curr, mAccacc, limitSumm,_
               isrGenerativ, allLim, autoCap, mDate, dateGive, dateAgr, valCheck, debtJPart, cardDebtType, _
               datesFilltype, agrBeg, agrFin, fixDays, agrPeriod, passDirection, summDateSelect,_
               summFillType, overRates, overRatesSect, unusedPortRate, unusedPortRateSec,_
               sect, sectNew, purpose, mShedule, mGuarantee, mCountry, lRegion, mRegion, paperCode
               
      Dim giveMoneyISN, docNumOut, calcfISN, orderISN, orderNum, insGrISN, param, status, paramN               
      Dim contractName, wMainForm, frmAsUstPar, tdbgViewn, frmPttel, tdbgView      
      Dim ovDate, overSumm, cashOrNo, accCorr, dateCharge,dateAction, percentMoney, ordDate, AccDb, AccCr, ordMoney      
      Dim startDate, endDate, calcDate, regDate, checkCount, dateType, startD, endD, userName, creatDate, aim      
      Dim queryString, sqlValue, colNum, sql_isEqual
         
      fDATE = "20250101"
      sDATE = "20120101"
      Call Initialize_AsBank("bank", sDATE, fDATE)
      Call Create_Connection()
      
      ' Մուտք գործել համակարգ ARMSOFT օգտագործողով
      Login("ARMSOFT")
      
      ' Մուտք գործել Օվերդրաֆտ (տեղաբաշխված) ԱՇՏ
      Call ChangeWorkspace(c_Overdraft)
      Call wTreeView.DblClickItem("|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|Üáñ å³ÛÙ³Ý³·ñÇ ëï»ÕÍáõÙ")
      
      aim = "Üå³ï³Ï"
      contType = "¶ñ³ýÇÏáí ûí»ñ¹ñ³ýïÇ å³ÛÙ³Ý³·Çñ"
      clientCode = "00034851"
      agType = "0001"
      curr = ""
      mAccacc = "30220042300"
      limitSumm = "100000"
'      isrGenerativ = ""
'      allLim = ""
'      autoCap = "" 
      mDate = "180714"
      dateGive = "180714"
      dateAgr = "180715"
'      valCheck = 0
      debtJPart = "1"
      cardDebtType = ""
'      datesFilltype = "2"
'      agrBeg = "180714"
'      agrFin = "180715"
'      fixDays = ""
'      agrPeriod = "1"
'      passDirection = "2"
'      summDateSelect = "2"
'      summFillType = "01"
'      overRates = "10"
'      overRatesSect = "365"
'      unusedPortRate = ""
'      unusedPortRateSec = ""
      sect = "A"
      sectNew = "01.001"
      purpose = "00"
      mShedule = "9"
      mGuarantee = "9"
      mCountry = "AM"
      lRegion = "001"
      mRegion = "010000008"
      paperCode = "1"
       
      Call CreatingOverdraftWithSchedule( contType, fISN, docNum, clientCode, agType, curr, mAccacc, limitSumm,_
                                                                           isrGenerativ, allLim, autoCap, mDate, dateGive, dateAgr, valCheck, debtJPart,_
                                                                           cardDebtType, datesFilltype, agrBeg, agrFin, fixDays, agrPeriod, passDirection,_
                                                                           summDateSelect, summFillType, overRates, overRatesSect, unusedPortRate,_
                                                                           unusedPortRateSec, sect, sectNew, purpose, mShedule, mGuarantee, mCountry,_
                                                                           lRegion, mRegion, paperCode)
                                                                           
      Log.Message(fISN)
      Log.Message(docNum)
      BuiltIn.Delay(10000)
      
            'CONTRACTS
             queryString = " SELECT COUNT(*) FROM CONTRACTS WHERE fDGISN = " & fISN & _
                                      " AND fDGSTATE = 206 AND  fDGSUMMA = 100000 AND fDGALLSUMMA = 0 " & _ 
                                      " AND fDGRISKDEGREE = 0 AND fDGRISKDEGNB = 0 AND fDGMPERCENTAGE = 0 "
              sqlValue = 1
              colNum = 0
              sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
              If Not sql_isEqual Then
                Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
              End If 
          
                                
              'FOLDERS
              queryString = " SELECT COUNT(*) FROM FOLDERS WHERE fISN = " & fISN 
              sqlValue = 3
              colNum = 0
              sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
              If Not sql_isEqual Then
                Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
              End If
          
              ' RESNUMBERS
              queryString = "SELECT COUNT(*)  FROM RESNUMBERS WHERE fISN =" & fISN & " And fTYPE = 'C' "
              sqlValue = 1
              colNum = 0
              sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
              If Not sql_isEqual Then
                Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
              End If
      
      BuiltIn.Delay(5000)
      Set frmPttel=  Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel")
      Set tdbgViewn = Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").VBObject("tdbgView")

      ' Ստուգում որ Պայմանագիրը ստեղծվել է
      If tdbgViewn.ApproxCount <> 1 Then
                Log.Error("Պայմանագրիը չի ստեղծվել")
                Exit Sub
      End If
      
      Set wMainForm = Sys.Process("Asbank").VBObject("MainForm")
      ' Կատարել բոլոր գործողությունները
      Call wMainForm.MainMenu.Click(c_AllActions)
      ' Մարման գրաֆիկի նշանակում
      Call wMainForm.PopupMenu.Click(c_RepaySchedule)    
       
      contractName = "¶ñ³ýÇÏáí ûí»ñ¹ñ³ýïÇ å³ÛÙ³Ý³·Çñ- " & Trim(docNum)  & " {öáË³ÝóÙ³Ý ëïáõ·Ù³Ý Ñ³×³Ëáñ¹ 1} "                                                               
      BuiltIn.Delay(1000)
      
      ' Օվերդրաֆտ պայմանագրի հաստատում
      status =  SendToApprove(contractName)
           
      ' Ստուգում որ Օվերդրաֆտ պայմանագրի հաստատումը տեղի է ունեցել
      If Not  status Then
              Log.Error("Օվերդրաֆտ պայմանագրի հաստատում տեղի չի ունեցել")
              Exit Sub
      End If
        
      ' Մուտք Հաստատող փաստաթղթեր 1
      Call wTreeView.DblClickItem("|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I")                                                                     
      BuiltIn.Delay(1000)
      Set frmAsUstPar = Sys.Process("Asbank").VBObject("frmAsUstPar")
       
      ' Ստուգում որ Օվերդրաֆտ (տեղաբաշխված) ՝ հաստատվող փաստաթղթեր 1 դիալոգը բացվել է
      If  Not frmAsUstPar.Exists Then
          Log.Error("Օվերդրաֆտ (տեղաբաշխված)՝ հաստատվող փաստաթղթեր 1 դիալոգը չի բացվել")
          Exit Sub
      End If
      
      ' Պայմանագրի համար դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "NUM", docNum )
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
      BuiltIn.Delay(4000)
      
      ' Ստուգում որ Վարկային պայմանագիրը ուղարկվել է հաստատվող փաստաթղթեր 1
      If tdbgViewn.ApproxCount <> 1 Then
                Log.Error("Վարկային պայմանագրիը չի ուղարկվել հաստատվող 1")
                Exit Sub
      End If
      
      ' Փաստաթղթի վավերացում 
      Call DocValidate(docNum)
      frmPttel.Close
      
                  'AGRSCHEDULEVALUES
                  queryString = " SELECT COUNT(*) FROM AGRSCHEDULEVALUES WHERE fAGRISN = " & fISN & _
                                           " AND fSUM = 0 " 
                  sqlValue = 24
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
      
                 'AGRSCHEDULE
                  queryString = " SELECT COUNT(*) FROM AGRSCHEDULE WHERE fAGRISN = " & fISN  & _
                                            " AND fKIND = 9 "
                  sqlValue = 1
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If    
     
                  'HI
                  queryString =  " SELECT SUM(fSUM)  FROM HI WHERE fBASE =  " & fISN & _
                                            " AND ( fSUM = 100000 AND fCURSUM = 100000 )"
                  sqlValue = 200000.00
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
                  'HIF
                  queryString = " SELECT SUM(fSUM) FROM HIF WHERE fBASE = " &  fISN 
                  sqlValue = 100032.4418
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
        
                  'CONTRACTS
                  queryString = " SELECT COUNT(*) FROM CONTRACTS WHERE fDGISN =  " & fISN  & _
                                           " AND fDGCUR = 0 AND  fDGSTATE = 7 AND  fDGSUMMA = 100000  " & _
                                           " AND fDGALLSUMMA = 0 AND fDGRISKDEGREE = 0 " & _
                                           " AND fDGRISKDEGNB = 0 AND fDGMPERCENTAGE = 0 "  
                  sqlValue = 1
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
        
                  'CAGRACCS
                  queryString = " SELECT COUNT(*) FROM CAGRACCS WHERE fAGRISN  = " &  fISN 
                  sqlValue = 1
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If

      ' Մուտք պայմանագրեր թղթապանակ
      Call wTreeView.DblClickItem("|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|ä³ÛÙ³Ý³·ñ»ñ")
      ' Պայմանագրի համար դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "NUM", docNum )
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
      
      ' Օվերդրաֆտի տրամադրում 8
      ovDate = "180714"
      overSumm = "50000"
      cashOrNo = "2"
      accCorr = "30220042300"
      Call GiveOverdraft(giveMoneyISN, docNumOut, ovDate, overSumm, cashOrNo, accCorr)
      Log.Message(giveMoneyISN)  
      BuiltIn.Delay(3000)   
      
                  'AGRSCHEDULEVALUES
                  queryString = " SELECT COUNT(*) FROM AGRSCHEDULEVALUES WHERE fAGRISN = " & fISN 
                  sqlValue = 48
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
      
                 'AGRSCHEDULE
                  queryString = " SELECT COUNT(*) FROM AGRSCHEDULE WHERE fAGRISN = " & fISN  & _
                                           " AND ( fKIND = 9 OR fKIND = 3 ) "
                  sqlValue = 2
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If    
     
                  'HI
                  queryString =  " SELECT COUNT(*)  FROM HI WHERE fBASE =  " & giveMoneyISN & _
                                            " AND fSUM = 50000 AND fCURSUM = 50000 " 
                  sqlValue = 3
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                 
                  'HIR 
                 queryString =  " SELECT fCURSUM  FROM HI WHERE fBASE =  " & giveMoneyISN
                  sqlValue = 50000
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
                  'HIF
                  queryString = " SELECT SUM(fSUM) FROM HIF WHERE fBASE = " &  fISN & _
                                            " AND (( fSUM = 0 AND fCURSUM = 0 ) " & _
                  													" OR ( fSUM = 100000 AND fCURSUM = 0 ) " & _
                  													" OR ( fSUM = 10 AND fCURSUM = 365 ) " & _
                  													" OR ( fSUM = 0 AND fCURSUM = 1 ) " & _
                  													" OR ( fSUM = 10.4709 AND fCURSUM = 365 ) " & _
                  													" OR ( fSUM = 0.50 AND fCURSUM = 365 ) " & _
                  													" OR ( fSUM = 0 AND fCURSUM = 1 ) " & _
                  													" OR ( fSUM = 1 AND fCURSUM = 0 ))"
                  sqlValue =100032.4418
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
        
                  'HIRREST
                  queryString = " SELECT COUNT(*) FROM HIRREST WHERE fOBJECT =  " & fISN & _
                                           " AND fPENULTREM = 0 AND fLASTREM = 50000 "
                  sqlValue = 1
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
      
      ' Օվերդրաֆտի տոկոսների հաշվարկ 9
      dateCharge = "140814"
      dateAction = "140814"      
      Call PercentCalculation(dateCharge, dateAction, percentMoney, calcfISN )  
      BuiltIn.Delay(700)
      frmPttel.Close
      Log.Message(calcfISN)     
      
                'HI
                queryString =  " SELECT SUM(fSUM) FROM HI WHERE fBASE = " & calcfISN & _ 
                                          " AND (( fSUM = 383.6 AND fCURSUM = 383.6 ) " & _
											                    " OR ( fSUM = 19.2 AND fCURSUM = 19.2 )) "
                sqlValue = 1572.80
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
                
                'HIF
                queryString = " SELECT COUNT(*) FROM HIF WHERE fBASE = " & calcfISN  & _
                                         " AND fSUM = 0 AND fCURSUM = 0 "
                sqlValue = 1
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
           
                'HIR
                queryString = " SELECT COUNT(*) FROM HIR WHERE fBASE = " & calcfISN  & _
                                         " AND ( fCURSUM = 383.60 OR fCURSUM = 19.2 ) "
                sqlValue = 2
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
           
                'HIRREST
                queryString = " SELECT COUNT(*)  FROM HIRREST WHERE fOBJECT =  " &  fISN & _
                                         " AND ( fLASTREM = 50000 AND fPENULTREM = 0 ) "& _
  											                 " OR ( fLASTREM = 383.6 AND fPENULTREM = 0 ) " & _
  											                 " OR ( fLASTREM = 19.2 AND fPENULTREM = 0 ) "
                sqlValue = 3
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
        
                'HIT
                queryString = " SELECT  COUNT(*) FROM HIT WHERE fBASE = " &  calcfISN & _
                                         " AND ( fCURSUM = 383.60 OR fCURSUM = 19.2 ) "
                sqlValue = 2
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
       
      ' Մուտք Հաճախորդի սպասարկում և դրամարկղ ԱՇՏ
      Call ChangeWorkspace(c_CustomerService)
      ' Մուտք Աշխատանքային փաստաթղթեր թղթապանակ
      Call wTreeView.DblClickItem("|Ð³×³Ëáñ¹Ç ëå³ë³ñÏáõÙ ¨ ¹ñ³Ù³ñÏÕ |²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ") 
      
      ' Ժամանակահատվածի սկիզբ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "PERN", mDate )
      ' Ժամանակահատվածի ավարտ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "PERK", dateAgr )
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
        
      ' Հիշարար օրդերի ստեղծում  11
      ordDate = "150814"
      AccDb = "000005200"
      AccCr = "30220042300"
      ordMoney = "4000"
      Call CreateMemOrders(orderISN, orderNum, ordDate, AccDb, AccCr, ordMoney, aim)        
      Log.Message(orderISN)
      BuiltIn.Delay(700)
         
                'HI
                queryString =  " SELECT COUNT(*) FROM HI WHERE fBASE = " & orderISN & _ 
                                          " AND fSUM = 4000 AND fCURSUM = 4000 "
                sqlValue = 2
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
           
                'HIRREST
                queryString = " SELECT COUNT(*)  FROM HIRREST WHERE fOBJECT =  " &  fISN & _
                                         " AND fPENULTREM = 0 AND fSTARTREM = 0 " & _
									                       " AND ( fLASTREM = 50000 OR fLASTREM = 383.60 OR fLASTREM = 19.2 ) "
                sqlValue = 3
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
        
                'MEMORDERS
                queryString = " SELECT fSUMMA FROM MEMORDERS WHERE fISN =  " &  orderISN 
                sqlValue = 4000
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
                
      ' Մուտք գործել Օվերդրաֆտ (տեղաբաշխված) ԱՇՏ
       Call ChangeWorkspace(c_Overdraft)
      ' Մուտք Օվերդրաֆտ ունեցող հաշիվներ
      Call wTreeView.DblClickItem("|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|úí»ñ¹ñ³ýï áõÝ»óáÕ Ñ³ßÇíÝ»ñ") 
      
      ' Ստուգում որ Օվերդրաֆտներ դիալոգը բացվել է
      If Not Sys.Process("Asbank").VBObject("frmAsUstPar").Exists Then
              Log.Error("Օվերդրաֆտներ դիալոգը չի բացվել")
              Exit Sub
      End If
        
      BuiltIn.Delay(700)
      ' Պայմանագրի համար դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "NUM", docNum)
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
        
      BuiltIn.Delay(4000)
      ' Պայմանագրի առկայության ստուգում Օվերդրաֆտներ թղթապանակում
      If tdbgViewn.ApproxCount <> 1 Then
              Log.Error("Պայմանագիրն առկա չէ Օվերդրաֆտներ թղթապանակում")
              Exit Sub
      End If
        
      ' Օվերդրաֆտի խմբային մարում 13
      endDate = "150814"
      Call OverdraftGroupRepayment(ordDate, endDate)
      ' Փակել Օվերդրաֆտներ թղթապանակը
      BuiltIn.Delay(700)
      frmPttel.Close
      
      ' Մուտք Հաճախորդի սպասարկում և դրամարկղ ԱՇՏ 
      Call ChangeWorkspace(c_CustomerService)
      ' Մուտք Աշխատանքային փաստաթղթեր թղթապանակ
      Call wTreeView.DblClickItem("|Ð³×³Ëáñ¹Ç ëå³ë³ñÏáõÙ ¨ ¹ñ³Ù³ñÏÕ |²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ") 
      
      ' Ժամանակահատվածի սկիզբ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "PERN", mDate )
      ' Ժամանակահատվածի ավարտ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "PERK", dateAgr )
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")            
      
      ' Հիշարար օրդերի ստեղծում  15
      ordDate = "150814"
      AccDb = "30220042300"
      AccCr = "000005200"
      ordMoney = "5000"
      Call CreateMemOrders(orderISN, orderNum, ordDate, AccDb, AccCr, ordMoney, aim)    
      Log.Message(orderISN)
      BuiltIn.Delay(700)
       
               'HI
                queryString =  " SELECT COUNT(*) FROM HI WHERE fBASE = " & orderISN & _ 
                                          " AND fSUM = 5000 AND fCURSUM = 5000 "
                sqlValue = 2
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
           
                'HIRREST
                queryString = " SELECT COUNT(*)  FROM HIRREST WHERE fOBJECT =  " &  fISN & _
                                         " AND fPENULTREM = 0 AND fSTARTREM =0 AND " & _ 
									                       " ( fLASTREM = 50000 OR fLASTREM = 383.60 OR fLASTREM = 19.2 ) "
                sqlValue = 3
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If 

      ' Մուտք գործել Օվերդրաֆտ (տեղաբաշխված) ԱՇՏ
      Call ChangeWorkspace(c_Overdraft)
      ' Մուտք Օերդրաֆտ ունեցող հաշիվներ
      Call wTreeView.DblClickItem("|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|úí»ñ¹ñ³ýï áõÝ»óáÕ Ñ³ßÇíÝ»ñ") 
      
      ' Ստուգում որ Օվերդրաֆտներ դիալոգը բացվել է
      If Not Sys.Process("Asbank").VBObject("frmAsUstPar").Exists Then
              Log.Error("Օվերդրաֆտներ դիալոգը չի բացվել")
              Exit Sub
      End If
      BuiltIn.Delay(700)
        
      ' Պայմանագրի համար դաշտի լրացում
      Call Rekvizit_Fill("Dialog" ,1, "General", "NUM", docNum)
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
        
      ' Օվերդրաֆտի խմբային տրամադրում  16
      startD = "150814"
      endD = "150814"
      Call GiveOverdraftGroup(startD, endD)
      BuiltIn.Delay(700)
      frmPttel.Close
      
      ' Մուտք պայմանագրեր թղթապանակ
      Call wTreeView.DblClickItem("|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|ä³ÛÙ³Ý³·ñ»ñ")
      ' Պայմանագրի համար դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "NUM", docNum )
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
      
      ' Խմբային տոկոսների հաշվարկի կատարում 17
      calcDate = "170814"
      regDate = "170814"
      checkCount = 1
      Call InterestGroupCalculationOverdraft (calcDate,regDate, checkCount)
      
      ' Խմբային տոկոսների հաշվարկ փաստաթղթի ISN - ի ստացում
      status = True
      paramN = c_ViewEdit & "|" & c_Other & "|" & c_CalcDates
      Call GetDocISN(paramN, calcDate, status, dateType, insGrISN)
      Log.Message(insGrISN)
        
        
                'HI
                queryString = " SELECT  COUNT(*) FROM HI WHERE fBASE = " &  insGrISN & _
                                         " AND (( fSUM = 41.1 AND fCURSUM = 41.1 ) " & _
										                     " OR ( fSUM = 2.1 AND fCURSUM = 2.1 )) "
                sqlValue = 6
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
 
                'HIF
                queryString = " SELECT COUNT(*) FROM HIF WHERE fBASE = " & insGrISN & _
                                         " AND fSUM = 0 AND fCURSUM = 0"
                sqlValue = 1
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If       

                'HIR
                queryString = " SELECT COUNT(*) FROM HIR WHERE fBASE = " &  insGrISN & _
                                          " AND ( fCURSUM = 41.1 OR fCURSUM = 2.1 OR fCURSUM = 424.7 " & _
                                          " OR fCURSUM = 21.3 OR fCURSUM = 4166.7 ) "
                sqlValue = 5
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
 
                'HIRREST
                queryString = " SELECT COUNT(*) FROM HIRREST WHERE fOBJECT = " & fISN & _
                                         " AND (( fLASTREM = 50000 AND fPENULTREM = 0 ) " & _
              												   " OR ( fLASTREM = 424.70 AND fPENULTREM = 383.60 ) " & _
              												   " OR ( fLASTREM = 21.30 AND fPENULTREM = 19.2 ) " & _
              												   " OR ( fLASTREM = 424.70 AND fPENULTREM = 0 ) " & _
              												   " OR ( fLASTREM = 21.30 AND fPENULTREM = 0 ) " & _
              												   " OR ( fLASTREM = 4166.70 AND fPENULTREM = 0 )) " 
                sqlValue = 6
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
                
                'HIT
                queryString = " SELECT  COUNT(*) FROM HIT WHERE fBASE = " &  insGrISN & _
                                         " AND ( fCURSUM = 41.1 OR fCURSUM = 2.1 ) "
                sqlValue = 2
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
      
      ' Խմբային տոկոսների հաշվարկի կատարում 18
      calcDate = "180814"
      regDate =  "180814"
      Call InterestGroupCalculationOverdraft (calcDate, regDate, checkCount)
      
      ' Խմբային տոկոսների հաշվարկ փաստաթղթի ISN - ի ստացում
      Call GetDocISN(paramN, calcDate, status, dateType, insGrISN)
      Log.Message(insGrISN)
        
                'HI
                queryString = " SELECT  COUNT(*) FROM HI WHERE fBASE = " &  insGrISN & _
                                         " AND (( fSUM = 12.60 AND fCURSUM = 12.60 AND fDBCR = 'C' ) " & _
										                     " OR ( fSUM = 0.7 AND fCURSUM = 0.7 AND fDBCR = 'C' ) " & _ 
										                     " OR ( fSUM = 12.60 AND fCURSUM = 12.60 AND fDBCR = 'D' ) " & _
										                     " OR ( fSUM = 0.7 AND fCURSUM = 0.7 AND fDBCR = 'D' )) "
                sqlValue = 6
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
 
                'HIF
                queryString = " SELECT COUNT(*) FROM HIF WHERE fBASE = " & insGrISN & _
                                         " AND fSUM = 0 AND fCURSUM = 0 "
                sqlValue = 6
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If       

                'HIR
                queryString = " SELECT COUNT(*) FROM HIR WHERE fBASE = " &  insGrISN & _
                                         " AND ( fCURSUM = 12.6 OR fCURSUM = 0.7 ) "
                sqlValue = 2
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
 
                'HIRREST
                queryString = " SELECT COUNT(*) FROM HIRREST WHERE fOBJECT = " & fISN & _
                                         " AND (( fLASTREM = 50000 AND fPENULTREM = 0 ) " & _
                												 " OR ( fLASTREM = 437.30 AND fPENULTREM = 424.70 ) " & _
                												 " OR ( fLASTREM = 22 AND fPENULTREM = 21.3 ) " & _
                												 " OR ( fLASTREM = 424.70 AND fPENULTREM = 0 ) " & _
                												 " OR ( fLASTREM = 21.30 AND fPENULTREM = 0 ) " & _
                												 " OR ( fLASTREM = 4166.70 AND fPENULTREM = 0 )) " 
                sqlValue = 6
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
                
                'HIT
                queryString = " SELECT  COUNT(*) FROM HIT WHERE fBASE = " &  insGrISN & _
                                         " AND (( fCURSUM = 12.6 AND fTYPE = 'N2' ) " & _
										                     " OR ( fCURSUM = 0.7 AND fTYPE = 'NH' )) "
                sqlValue = 2
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
      
      ' Խմբային տոկոսների հաշվարկի կատարում 19
      calcDate = "250814"
      regDate = "250814"
      Call InterestGroupCalculationOverdraft (calcDate, regDate, checkCount)
      
      ' Խմբային տոկոսների հաշվարկ փաստաթղթի ISN - ի ստացում
      Call GetDocISN(paramN, calcDate, status, dateType, insGrISN)
      Log.Message(insGrISN)
      frmPttel.Close
        
                'HI
                queryString = " SELECT  COUNT(*) FROM HI WHERE fBASE = " &  insGrISN & _
                                         " AND (( fSUM = 87.90 AND fCURSUM = 87.90 AND fDBCR = 'C' ) " & _
                    										 " OR ( fSUM = 4.80 AND fCURSUM = 4.80 AND fDBCR = 'C' ) " & _
                    										 " OR ( fSUM = 87.90 AND fCURSUM = 87.90 AND fDBCR = 'D' ) " & _
                    										 " OR ( fSUM = 4.80 AND fCURSUM = 4.80 AND fDBCR = 'D' )) "
                sqlValue = 6
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
 
                'HIF
                queryString = " SELECT COUNT(*) FROM HIF WHERE fBASE = " & insGrISN & _
                                         " AND fSUM = 0 AND fCURSUM = 0 "
                sqlValue = 1
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If       

                'HIR
                queryString = " SELECT COUNT(*) FROM HIR WHERE fBASE = " &  insGrISN & _
                                         " AND (( fCURSUM = 87.9 AND fTYPE = 'R2' ) " & _
										                     " OR  ( fCURSUM = 4.80 AND fTYPE = 'RH' )) "
                sqlValue = 2
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
 
                'HIRREST
                queryString = " SELECT COUNT(*) FROM HIRREST WHERE fOBJECT = " & fISN & _
                                         " AND (( fLASTREM = 50000 AND fPENULTREM = 0 ) " & _
                												 " OR ( fLASTREM = 525.20 AND fPENULTREM = 437.30 ) " & _
                												 " OR ( fLASTREM = 26.80 AND fPENULTREM = 22 ) " & _
                												 " OR ( fLASTREM = 424.70 AND fPENULTREM = 0 ) " & _
                												 " OR ( fLASTREM = 21.30 AND fPENULTREM = 0 ) " & _
                												 " OR ( fLASTREM = 4166.70 AND fPENULTREM = 0 ))" 
                sqlValue = 6
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
                
                'HIT
                queryString = " SELECT  COUNT(*) FROM HIT WHERE fBASE = " &  insGrISN & _
                                         " AND (( fCURSUM = 87.90 AND fTYPE = 'N2' ) " & _
										                     " OR ( fCURSUM = 4.8 AND fTYPE = 'NH' )) "
                sqlValue = 2
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
                        
      ' Մուտք Հաճախորդի սպասարկում և դրամարկղ ԱՇՏ 
      Call ChangeWorkspace(c_CustomerService)
      ' Մուտք Աշխատանքային փաստաթղթեր թղթապանակ
      Call wTreeView.DblClickItem("|Ð³×³Ëáñ¹Ç ëå³ë³ñÏáõÙ ¨ ¹ñ³Ù³ñÏÕ |²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ") 
      
      ' Ժամանակահատվածի սկիզբ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "PERN", mDate )
      ' Ժամանակահատվածի ավարտ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "PERK", dateAgr )
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")      
                
      ' Հիշարար օրդերի ստեղծում  20
      ordDate = "260814"
      AccDb = "000005200"
      AccCr = "30220042300"
      ordMoney = "3000"
      Call CreateMemOrders(orderISN, orderNum, ordDate, AccDb, AccCr, ordMoney, aim)      
      Log.Message(orderISN)
      BuiltIn.Delay(700)
         
                'HI
                queryString =  " SELECT COUNT(*) FROM HI WHERE fBASE = " & orderISN & _ 
                                          " AND (( fSUM = 3000 AND fCURSUM = 3000 AND fDBCR = 'D' ) " & _
											                    " OR ( fSUM = 3000 AND fCURSUM = 3000 AND fDBCR = 'C' )) "
                sqlValue = 2
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
           
                'HIRREST
                queryString = " SELECT COUNT(*)  FROM HIRREST WHERE fOBJECT =  " &  fISN & _
                                         " AND (( fLASTREM = 50000.00 AND fPENULTREM = 0 AND fSTARTREM = 0 ) " & _
                  											 " OR ( fLASTREM = 525.20 AND fPENULTREM = 437.30 AND fSTARTREM = 0 ) " & _
                  											 " OR ( fLASTREM = 26.80 AND fPENULTREM = 22 AND fSTARTREM = 0 ) " & _
                  											 " OR ( fLASTREM = 424.70 AND fPENULTREM = 0 AND fSTARTREM = 0 ) " & _
                  											 " OR ( fLASTREM = 21.30 AND fPENULTREM = 0 AND fSTARTREM = 0 ) " & _
                  											 " OR ( fLASTREM = 4166.70 AND fPENULTREM = 0 AND fSTARTREM = 0 )) "
                sqlValue = 6
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If 

      ' Մուտք գործել Օվերդրաֆտ (տեղաբաշխված) ԱՇՏ
      Call ChangeWorkspace(c_Overdraft)
      ' Մուտք Օվերդրաֆտ ունեցող հաշիվներ
      Call wTreeView.DblClickItem("|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|úí»ñ¹ñ³ýï áõÝ»óáÕ Ñ³ßÇíÝ»ñ") 
      
      ' Ստուգում որ Օվերդրաֆտներ դիալոգը բացվել է
      If Not Sys.Process("Asbank").VBObject("frmAsUstPar").Exists Then
              Log.Error("Օվերդրաֆտներ դիալոգը չի բացվել")
              Exit Sub
      End If
        
      BuiltIn.Delay(700)
      ' Պայմանագրի համար դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "NUM", docNum)
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
        
      BuiltIn.Delay(4000)
      ' Պայմանագրի առկայության ստուգում Օվերդրաֆտներ թղթապանակում
      If tdbgViewn.ApproxCount <> 1 Then
              Log.Error("Պայմանագիրն առկա չէ Օվերդրաֆտներ թղթապանակում")
              Exit Sub
      End If
        
      ' Օվերդրաֆտի խմբային մարում 21
      endDate = "260814"
      Call OverdraftGroupRepayment(ordDate, endDate)
      ' Փակել Օվերդրաֆտներ թղթապանակը
      BuiltIn.Delay(700)
      frmPttel.Close
      
      ' Մուտք պայմանագրեր թղթապանակ
      Call wTreeView.DblClickItem("|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|ä³ÛÙ³Ý³·ñ»ñ")
      ' Պայմանագրի համար դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "NUM", docNum )
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
      
      ' Խմբային տոկոսների հաշվարկի կատարում 22
      calcDate = "170914"
      regDate = "170914"
      Call InterestGroupCalculationOverdraft (calcDate,regDate, checkCount)
      
      ' Խմբային տոկոսների հաշվարկի փաստաթղթի ISN - ի ստացում
      Call GetDocISN(paramN, calcDate, status, dateType, insGrISN)
      Log.Message(insGrISN)
      
                'HI
                queryString = " SELECT  COUNT(*) FROM HI WHERE fBASE = " &  insGrISN & _
                                         " AND (( fSUM = 288.80 AND fCURSUM = 288.80 AND fDBCR = 'C' )  " & _
                    										 " OR ( fSUM = 15.80 AND fCURSUM = 15.80 AND fDBCR = 'C' ) " & _ 
                    										 " OR ( fSUM = 288.80 AND fCURSUM = 288.80 AND fDBCR = 'D' ) " & _
                    										 " OR ( fSUM = 15.80 AND fCURSUM = 15.80 AND fDBCR = 'D' )) "
                sqlValue = 6
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
 
                'HIF
                queryString = " SELECT COUNT(*) FROM HIF WHERE fBASE = " & insGrISN & _
                                          " AND fSUM = 0 AND fCURSUM = 0 "
                sqlValue = 1
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If       

                'HIR
                queryString = " SELECT COUNT(*) FROM HIR WHERE fBASE = " &  insGrISN & _
                                         " AND (( fCURSUM = 288.80 AND fTYPE = 'R2' ) " & _
                      									 " OR ( fCURSUM = 15.80 AND fTYPE = 'RH' ) " & _
                      									 " OR ( fCURSUM = 389.30 AND fTYPE = 'R¸' ) " & _
                      									 " OR ( fCURSUM = 21.30 AND fTYPE = 'RÂ' ) " & _
                      									 " OR ( fCURSUM = 4166.70 AND fTYPE = 'RÄ' )) "
                sqlValue = 5
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
 
                'HIRREST
                queryString = " SELECT COUNT(*) FROM HIRREST WHERE fOBJECT = " & fISN & _
                                         " AND (( fLASTREM = 50000 AND fPENULTREM = 0 ) " & _
                												 " OR ( fLASTREM = 814.00 AND fPENULTREM = 525.20 ) " & _
                												 " OR ( fLASTREM = 42.60 AND fPENULTREM = 26.80 ) " & _
                												 " OR ( fLASTREM = 814.00 AND fPENULTREM = 424.70 ) " & _
                												 " OR ( fLASTREM = 42.60 AND fPENULTREM = 21.30 ) " & _
                												 " OR ( fLASTREM = 8333.40 AND fPENULTREM = 4166.70 )) "
                sqlValue = 6
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
                
                'HIT
                queryString = " SELECT  COUNT(*) FROM HIT WHERE fBASE = " &  insGrISN & _
                                         " AND (( fCURSUM = 288.80 AND fTYPE = 'N2' ) " & _
										                     " OR ( fCURSUM = 15.80 AND fTYPE = 'NH' )) "
                sqlValue = 2
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
                      
      ' Խմբային տոկոսների հաշվարկի կատարում 23
      calcDate = "180914"
      regDate = "180914"
      Call InterestGroupCalculationOverdraft (calcDate, regDate, checkCount)
      
      ' Խմբային տոկոսների հաշվարկ փաստաթղթի ISN - ի ստացում
      Call GetDocISN(paramN, calcDate, status, dateType, insGrISN)
      Log.Message(insGrISN)
      
                'HI
                queryString = " SELECT  COUNT(*) FROM HI WHERE fBASE = " &  insGrISN & _
                                         " AND (( fSUM = 11.40 AND fCURSUM = 11.40 AND fDBCR = 'C' ) " & _
                    										 " OR ( fSUM = 0.7 AND fCURSUM = 0.7 AND fDBCR = 'C' ) " & _
                    										 " OR ( fSUM = 11.40 AND fCURSUM = 11.40 AND fDBCR = 'D' ) " & _
                    										 " OR ( fSUM = 0.7 AND fCURSUM = 0.7 AND fDBCR = 'D' )) "
                sqlValue = 6
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
 
                'HIF
                queryString = " SELECT COUNT(*) FROM HIF WHERE fBASE = " & insGrISN & _
                                         " AND fSUM = 0 AND fCURSUM = 0 "
                sqlValue = 4
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If       

                'HIR
                queryString = " SELECT COUNT(*) FROM HIR WHERE fBASE = " &  insGrISN & _
                                         " AND (( fCURSUM = 11.40 AND fTYPE = 'R2' ) " & _
										                     " OR  ( fCURSUM = 0.7 AND fTYPE = 'RH' )) "
                sqlValue = 2
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
 
                'HIRREST
                queryString = " SELECT COUNT(*) FROM HIRREST WHERE fOBJECT = " & fISN & _
                                         " AND (( fLASTREM = 50000 AND fPENULTREM = 0 ) " & _
                												 " OR ( fLASTREM = 825.40 AND fPENULTREM = 814.00 ) " & _
                												 " OR ( fLASTREM = 43.30 AND fPENULTREM = 42.60 ) " & _
                												 " OR ( fLASTREM = 814.00 AND fPENULTREM = 424.70 ) " & _
                												 " OR ( fLASTREM = 42.60 AND fPENULTREM = 21.30 ) " & _
                												 " OR ( fLASTREM = 8333.40 AND fPENULTREM = 4166.70 )) "
                sqlValue = 6
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
                
                'HIT
                queryString = " SELECT  COUNT(*) FROM HIT WHERE fBASE = " &  insGrISN & _
                                         " AND (( fCURSUM = 11.40 AND fTYPE = 'N2' ) " & _
										                     " OR ( fCURSUM = 0.70 AND fTYPE = 'NH' )) "
                sqlValue = 2
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
                     
      ' Խմբային տոկոսների հաշվարկի կատարում 24
      calcDate = "191014"
      regDate = "191014"
      Call InterestGroupCalculationOverdraft (calcDate, regDate, checkCount)
      
      ' Խմբային տոկոսների հաշվարկ փաստաթղթի ISN - ի ստացում
      Call GetDocISN(paramN, calcDate, status, dateType, insGrISN)
      Log.Message(insGrISN)
      frmPttel.Close
       
                'HI
                queryString = " SELECT  COUNT(*) FROM HI WHERE fBASE = " &  insGrISN & _
                                         " AND (( fSUM = 353.90 AND fCURSUM = 353.90 AND fDBCR = 'C' ) " & _
                    										 " OR ( fSUM = 21.20 AND fCURSUM = 21.20 AND fDBCR = 'C' ) " & _
                    										 " OR ( fSUM = 353.90 AND fCURSUM = 353.90 AND fDBCR = 'D' ) " & _
                    										 " OR ( fSUM = 21.20 AND fCURSUM = 21.20 AND fDBCR = 'D' )) "
                sqlValue = 6
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
 
                'HIF
                queryString = " SELECT COUNT(*) FROM HIF WHERE fBASE = " & insGrISN & _
                                         " AND fSUM = 0 AND fCURSUM = 0 "
                sqlValue = 1
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If       

                'HIR
                queryString = " SELECT SUM(fCURSUM) FROM HIR WHERE fBASE = " &  insGrISN 
                sqlValue = 4929
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
 
                'HIRREST
                queryString = " SELECT COUNT(*) FROM HIRREST WHERE fOBJECT = " & fISN & _
                                         " AND (( fLASTREM = 50000 AND fPENULTREM = 0 ) " & _
                												 " OR ( fLASTREM = 1179.30 AND fPENULTREM = 825.40 ) " & _
                												 " OR ( fLASTREM = 64.50 AND fPENULTREM = 43.30 ) " & _
                												 " OR ( fLASTREM = 1179.30 AND fPENULTREM = 814.00 ) " & _
                												 " OR ( fLASTREM = 64.50 AND fPENULTREM = 42.60 ) " & _
                												 " OR ( fLASTREM = 12500.10 AND fPENULTREM = 8333.40 )) "
                sqlValue = 6
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
                
                'HIT
                queryString = " SELECT  COUNT(*) FROM HIT WHERE fBASE = " &  insGrISN & _
                                         " AND (( fCURSUM = 353.90 AND fTYPE = 'N2' ) " & _
										                     " OR ( fCURSUM = 21.20 AND fTYPE = 'NH' )) "
                sqlValue = 2
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If 
                
      ' Մուտք Հաճախորդի սպասարկում և դրամարկղ
      Call ChangeWorkspace(c_CustomerService)
      ' Մուտք Աշխատանքային փաստաթղթեր թղթապանակ
      Call wTreeView.DblClickItem("|Ð³×³Ëáñ¹Ç ëå³ë³ñÏáõÙ ¨ ¹ñ³Ù³ñÏÕ |²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ") 
      
      ' Ժամանակահատվածի սկիզբ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "PERN", mDate )
      ' Ժամանակահատվածի ավարտ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "PERK", dateAgr )
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
        
      ' Հիշարար օրդերի ստեղծում  25
      ordDate = "201014"
      ordMoney = "4464.60"
      Call CreateMemOrders(orderISN, orderNum, ordDate, AccDb, AccCr, ordMoney, aim)      
      Log.Message(orderISN)
      BuiltIn.Delay(700)
       
                'HI
                queryString =  " SELECT COUNT(*) FROM HI WHERE fBASE = " & orderISN & _ 
                                          " AND (( fSUM = 4464.6 AND fCURSUM = 4464.6  AND fDBCR = 'D' ) " & _
											                    " OR	( fSUM = 4464.6 AND fCURSUM = 4464.6  AND fDBCR = 'C' )) "
                sqlValue = 2
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
           
                'HIRREST
                queryString = " SELECT COUNT(*)  FROM HIRREST WHERE fOBJECT =  " &  fISN & _
                                         " AND (( fLASTREM = 50000.00 AND fPENULTREM = 0 AND fSTARTREM = 0 ) " & _
                  											 " OR ( fLASTREM = 1179.30 AND fPENULTREM = 825.40 AND fSTARTREM = 0 ) " & _
                  											 " OR ( fLASTREM = 64.50 AND fPENULTREM = 43.30 AND fSTARTREM = 0 ) " & _
                  											 " OR ( fLASTREM = 1179.30 AND fPENULTREM = 814.00 AND fSTARTREM = 0 ) " & _
                  											 " OR ( fLASTREM = 64.50 AND fPENULTREM = 42.60 AND fSTARTREM = 0 ) " & _
                  											 " OR ( fLASTREM = 12500.10 AND fPENULTREM = 8333.40 AND fSTARTREM = 0 )) "
                sqlValue = 6
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If 
      
      ' Մուտք գործել Օվերդրաֆտ (տեղաբաշխված) ԱՇՏ
      Call ChangeWorkspace(c_Overdraft)
      ' Մուտք Օվերդրաֆտ ունեցող հաշիվներ թղթապանակ
      Call wTreeView.DblClickItem("|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|úí»ñ¹ñ³ýï áõÝ»óáÕ Ñ³ßÇíÝ»ñ") 
      
      ' Ստուգում որ Օվերդրաֆտներ դիալոգը բացվել է
      If Not Sys.Process("Asbank").VBObject("frmAsUstPar").Exists Then
              Log.Error("Օվերդրաֆտներ դիալոգը չի բացվել")
              Exit Sub
      End If
        
      BuiltIn.Delay(700)
      ' Պայմանագրի համար դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "NUM", docNum)
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
        
      BuiltIn.Delay(4000)
      ' Պայմանագրի առկայության ստուգում Օվերդրաֆտներ թղթապանակում
      If tdbgViewn.ApproxCount <> 1 Then
              Log.Error("Պայմանագիրն առկա չէ Օվերդրաֆտներ թղթապանակում")
              Exit Sub 
      End If
        
      ' Օվերդրաֆտի խմբային մարում 26
      endDate = "201014"
      Call OverdraftGroupRepayment(ordDate, endDate)
      BuiltIn.Delay(700)
      frmPttel.Close
                
      ' Մուտք Հաճախորդի սպասարկում և դրամարկղ 
      Call ChangeWorkspace(c_CustomerService)
      ' Մուտք Աշխատանքային փաստաթղթեր թղթապանակ
      Call wTreeView.DblClickItem("|Ð³×³Ëáñ¹Ç ëå³ë³ñÏáõÙ ¨ ¹ñ³Ù³ñÏÕ |²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ") 
      
      ' Ժամանակահատվածի սկիզբ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "PERN", mDate )
      ' Ժամանակահատվածի ավարտ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "PERK", dateAgr )
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
              
      ' Հիշարար օրդերի ստեղծում  27
      ordDate = "201014"
      ordMoney = "4464.60"
      Call CreateMemOrders(orderISN, orderNum, ordDate, AccDb, AccCr, ordMoney, aim)      
      Log.Message(orderISN)
      BuiltIn.Delay(700)
       
                'HI
                queryString =  " SELECT COUNT(*) FROM HI WHERE fBASE = " & orderISN & _ 
                                          " AND (( fSUM = 4464.6 AND fCURSUM = 4464.6  AND fDBCR = 'D' ) " & _
											                    " OR	( fSUM = 4464.6 AND fCURSUM = 4464.6  AND fDBCR = 'C' )) "
                sqlValue = 2
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
           
                'HIRREST
                queryString = " SELECT COUNT(*)  FROM HIRREST WHERE fOBJECT =  " &  fISN & _
                                         " AND (( fLASTREM = 50000.00 AND fPENULTREM = 0 AND fSTARTREM = 0 ) " & _
                  											 " OR ( fLASTREM = 1179.30 AND fPENULTREM = 825.40 AND fSTARTREM = 0 ) " & _
                  											 " OR ( fLASTREM = 64.50 AND fPENULTREM = 43.30 AND fSTARTREM = 0 ) " & _
                  											 " OR ( fLASTREM = 1179.30 AND fPENULTREM = 814.00 AND fSTARTREM = 0 ) " & _
                  											 " OR ( fLASTREM = 64.50 AND fPENULTREM = 42.60 AND fSTARTREM = 0 ) " & _
                  											 " OR ( fLASTREM = 12500.10 AND fPENULTREM = 8333.40 AND fSTARTREM = 0 )) "
                sqlValue = 6 
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If 
 
      ' Մուտք գործել Օվերդրաֆտ (տեղաբաշխված) ԱՇՏ
      Call ChangeWorkspace(c_Overdraft)
      ' Մուտք Օվերդրաֆտ ունեցող հաշիվներ
      Call wTreeView.DblClickItem("|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|úí»ñ¹ñ³ýï áõÝ»óáÕ Ñ³ßÇíÝ»ñ") 
      
      ' Ստուգում որ Օվերդրաֆտներ դիալոգը բացվել է
      If Not Sys.Process("Asbank").VBObject("frmAsUstPar").Exists Then
            Log.Error("Օվերդրաֆտներ դիալոգը չի բացվել")
            Exit Sub
      End If
      
      BuiltIn.Delay(700)
      ' Պայմանագրի համար դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "NUM", docNum)
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
        
      ' Օվերդրաֆտի խմբային տրամադրում 28
      startD = "201014"
      endD = "201014"
      Call GiveOverdraftGroup(startD, endD)
      BuiltIn.Delay(700)
      frmPttel.Close
      
      ' Մուտք պայմանագրեր թղթապանակ
      Call wTreeView.DblClickItem("|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|ä³ÛÙ³Ý³·ñ»ñ")
      ' Պայմանագրի համար դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "NUM", docNum )
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
      
      ' Խմբային տոկոսների հաշվարկի կատարում 29
      calcDate = "201014"
      regDate = "201014"
      Call InterestGroupCalculationOverdraft (calcDate, regDate, checkCount)
      
      ' Խմբային տոկոսների հաշվարկ փաստաթղթի ISN - ի ստացում
      Call GetDocISN(paramN, calcDate, status, dateType, insGrISN)
      Log.Message(insGrISN)
      
                'HI
                queryString = " SELECT  COUNT(*) FROM HI WHERE fBASE = " &  insGrISN & _
                                         " AND (( fSUM = 10.30 AND fCURSUM = 10.30 AND fDBCR = 'C' ) " & _
                    										 " OR ( fSUM = 0.70 AND fCURSUM = 0.70 AND fDBCR = 'C' ) " & _
                    										 " OR ( fSUM = 10.30 AND fCURSUM = 10.30 AND fDBCR = 'D' ) " & _
                    										 " OR ( fSUM = 0.70 AND fCURSUM = 0.70 AND fDBCR = 'D' )) "
                sqlValue = 6
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
 
                'HIF
                queryString = " SELECT COUNT(*) FROM HIF WHERE fBASE = " & insGrISN & _
                                         " AND fSUM = 0 AND fCURSUM = 0 "
                sqlValue = 4
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If       

                'HIR
                queryString = " SELECT COUNT(*) FROM HIR WHERE fBASE = " &  insGrISN & _
                                         " AND (( fCURSUM = 10.30 AND fTYPE = 'R2' ) " & _ 
										                     " OR  ( fCURSUM = 0.7 AND fTYPE = 'RH' )) "
                sqlValue = 2
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
 
                'HIRREST
                queryString = " SELECT COUNT(*) FROM HIRREST WHERE fOBJECT = " & fISN & _
                                         " AND (( fLASTREM = 50000 AND fPENULTREM = 0 ) " & _
                												 " OR ( fLASTREM = 1189.60 AND fPENULTREM = 1179.30 ) " & _
                												 " OR ( fLASTREM = 65.20 AND fPENULTREM = 64.50 ) " & _
                												 " OR ( fLASTREM = 1179.30 AND fPENULTREM = 814.00 ) " & _
                												 " OR ( fLASTREM = 64.50 AND fPENULTREM = 42.60 ) " & _
                												 " OR ( fLASTREM = 12500.10 AND fPENULTREM = 8333.40 )) "
                sqlValue = 6
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
                
                'HIT
                queryString = " SELECT  COUNT(*) FROM HIT WHERE fBASE = " &  insGrISN & _
                                         " AND (( fCURSUM = 10.3 AND fTYPE = 'N2' ) " & _
										                     " OR ( fCURSUM = 0.7 AND fTYPE = 'NH' )) "
                sqlValue = 2
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If 
      
      ' Խմբային տոկոսների հաշվարկի կատարում 30
      calcDate = "171114"
      regDate = "171114"
      Call InterestGroupCalculationOverdraft (calcDate, regDate, checkCount)
      
      ' Խմբային տոկոսների հաշվարկ փաստաթղթի ISN - ի ստացում
      Call GetDocISN(paramN, calcDate, status, dateType, insGrISN)
      Log.Message(insGrISN) 
      frmPttel.Close
        
                'HI
                queryString = " SELECT  COUNT(*) FROM HI WHERE fBASE = " &  insGrISN & _
                                         " AND (( fSUM = 287.60 AND fCURSUM = 287.60 AND fDBCR = 'C' ) " & _
                    										 " OR ( fSUM = 19.20 AND fCURSUM = 19.20 AND fDBCR = 'C' ) " & _
                    										 " OR ( fSUM = 287.60 AND fCURSUM = 287.60 AND fDBCR = 'D' ) " & _
                    										 " OR ( fSUM = 19.20 AND fCURSUM = 19.20 AND fDBCR = 'D' )) "
                sqlValue = 6
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
 
                'HIF
                queryString = " SELECT COUNT(*) FROM HIF WHERE fBASE = " & insGrISN & _
                                         " AND fSUM = 0 AND fCURSUM = 0 " 
                sqlValue = 1
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If       

                'HIR
                queryString = " SELECT COUNT(*) FROM HIR WHERE fBASE = " &  insGrISN & _
                                         " AND (( fCURSUM = 287.60 AND fTYPE = 'R2' ) " & _
                    										 " OR ( fCURSUM = 19.20 AND fTYPE = 'RH' ) " & _
                    										 " OR ( fCURSUM = 297.90 AND fTYPE = 'R¸' ) " & _
                    										 " OR ( fCURSUM = 19.90 AND fTYPE = 'RÂ' ) " & _ 
                    										 " OR ( fCURSUM = 4166.70 AND fTYPE = 'RÄ' )) "
                sqlValue = 5
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
 
                'HIRREST
                queryString = " SELECT COUNT(*) FROM HIRREST WHERE fOBJECT = " & fISN & _
                                         " AND (( fLASTREM = 50000 AND fPENULTREM = 0 ) " & _
                												 " OR ( fLASTREM = 1477.20 AND fPENULTREM = 1189.60 ) " & _
                												 " OR ( fLASTREM = 84.40 AND fPENULTREM = 65.20 ) " & _
                												 " OR ( fLASTREM = 1477.20 AND fPENULTREM = 1179.30 ) " & _
                												 " OR ( fLASTREM = 84.40 AND fPENULTREM = 64.50 ) " & _
                												 " OR ( fLASTREM = 16666.80 AND fPENULTREM = 12500.10 )) "
                sqlValue = 6
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
                
                'HIT
                queryString = " SELECT  COUNT(*) FROM HIT WHERE fBASE = " &  insGrISN & _
                                         " AND (( fCURSUM = 287.60 AND fTYPE = 'N2' ) " & _
										                     " OR ( fCURSUM = 19.20 AND fTYPE = 'NH' )) "
                sqlValue = 2
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If 
                
      Call ChangeWorkspace(c_CustomerService)
      ' Մուտք Հաճախորդի սպասարկում և դրամարկղ/ Աշխատանքային փաստաթղթեր
      Call wTreeView.DblClickItem("|Ð³×³Ëáñ¹Ç ëå³ë³ñÏáõÙ ¨ ¹ñ³Ù³ñÏÕ |²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ") 
      
      ' Ժամանակահատվածի սկիզբ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "PERN", mDate )
      ' Ժամանակահատվածի ավարտ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "PERK", dateAgr )
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
         
      ' Հիշարար օրդերի ստեղծում  31
      ordDate = "181114"
      ordMoney = "4166.70"
      Call CreateMemOrders(orderISN, orderNum, ordDate, AccDb, AccCr, ordMoney, aim)       
      Log.Message(orderISN)
       
                'HI
                queryString =  " SELECT COUNT(*) FROM HI WHERE fBASE = " & orderISN & _ 
                                          " AND (( fSUM = 4166.70 AND fCURSUM = 4166.70  AND fDBCR = 'D' ) " & _
											                    " OR ( fSUM = 4166.70 AND fCURSUM = 4166.70  AND fDBCR = 'C' )) "
                sqlValue = 2
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
           
                'HIRREST
                queryString = " SELECT COUNT(*)  FROM HIRREST WHERE fOBJECT =  " &  fISN & _
                                         " AND  (( fLASTREM = 50000.00 AND fPENULTREM = 0 AND fSTARTREM = 0 ) " & _
                  											 " OR ( fLASTREM = 1477.20 AND fPENULTREM = 1189.60 AND fSTARTREM = 0 ) " & _
                  											 " OR ( fLASTREM = 84.40 AND fPENULTREM = 65.20 AND fSTARTREM = 0 ) " & _
                  											 " OR ( fLASTREM = 1477.20 AND fPENULTREM = 1179.30 AND fSTARTREM = 0 ) " & _
                  											 " OR ( fLASTREM = 84.40 AND fPENULTREM = 64.50 AND fSTARTREM = 0 ) " & _
                  											 " OR ( fLASTREM = 16666.80 AND fPENULTREM = 12500.10 AND fSTARTREM = 0 )) "
                sqlValue = 6 
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If 
          
      ' Մուտք գործել Օվերդրաֆտ (տեղաբաշխված) ԱՇՏ
      Call ChangeWorkspace(c_Overdraft)
      ' Մուտք Օվերդրաֆտ ունեցող հաշիվներ թղթապանակ
      Call wTreeView.DblClickItem("|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|úí»ñ¹ñ³ýï áõÝ»óáÕ Ñ³ßÇíÝ»ñ") 
          
      ' Ստուգում որ Օվերդրաֆտներ դիալոգը չի բացվել
      If Not Sys.Process("Asbank").VBObject("frmAsUstPar").Exists Then
              Log.Error("Օվերդրաֆտներ դիալոգը չի բացվել")
              Exit Sub
      End If
        
      BuiltIn.Delay(700)
      ' Պայմանագրի համար դաշտի լրացում
      Call Rekvizit_Fill("Dialog" ,1, "General", "NUM", docNum)
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
      
      BuiltIn.Delay(4000) 
      ' Պայմանագրի առկայության ստուգում Օվերդրաֆտներ թղթապանակում
      If tdbgViewn.ApproxCount <> 1 Then
              Log.Error("Պայմանագիրն առկա չէ Օվերդրաֆտներ թղթապանակում")
              Exit Sub
      End If
       
      ' Օվերդրաֆտի խմբային մարում 32
      endDate = "181114"
      Call OverdraftGroupRepayment(ordDate, endDate)
      ' Փակել Օվերդրաֆտներ թղթապանակը
      BuiltIn.Delay(700)
      frmPttel.Close
       
      ' Մուտք պայմանագրեր թղթապանակ
      Call wTreeView.DblClickItem("|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|ä³ÛÙ³Ý³·ñ»ñ")
      ' Պայմանագրի համար դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "NUM", docNum )
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
      
      ' Խմբային տոկոսների հաշվարկի կատարում 33
      calcDate = "171214"
      regDate = "171214"
      Call InterestGroupCalculationOverdraft (calcDate,regDate, checkCount)
      
      ' Խմբային տոկոսների հաշվարկ փաստաթղթի ISN - ի ստացում
      Call GetDocISN(paramN, calcDate, status, dateType, insGrISN)
      Log.Message(insGrISN)
      frmPttel.Close
       
                'HI
                queryString = " SELECT  SUM(fSUM) FROM HI WHERE fBASE = " &  insGrISN 
                sqlValue = 1137.20
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
 
                'HIF
                queryString = " SELECT COUNT(*) FROM HIF WHERE fBASE = " & insGrISN & _
                                         " AND fSUM = 0 AND fCURSUM = 0 "
                sqlValue = 5
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If       

                'HIR
                queryString = " SELECT COUNT(*) FROM HIR WHERE fBASE = " &  insGrISN & _
                                         " AND (( fCURSUM = 9.10 AND fTYPE = 'R2' ) " & _
                    										 " OR  ( fCURSUM = 0.70 AND fTYPE = 'RH' ) " & _
                    										 " OR  ( fCURSUM = 264.90 AND fTYPE = 'R2' ) " & _
                    										 " OR  ( fCURSUM = 19.90 AND fTYPE = 'RH' ) " & _
                    										 " OR  ( fCURSUM = 274.00 AND fTYPE = 'R¸' ) " & _
                    										 " OR  ( fCURSUM = 20.60 AND fTYPE = 'RÂ' ) " & _
                    										 " OR  ( fCURSUM = 4166.70 AND fTYPE = 'RÄ' )) "
                sqlValue = 7
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
 
                'HIRREST
                queryString = " SELECT COUNT(*) FROM HIRREST WHERE fOBJECT = " & fISN & _
                                         " AND (( fLASTREM = 50000 AND fPENULTREM = 0 ) " & _
                												 " OR ( fLASTREM = 1751.20 AND fPENULTREM = 1477.20 ) " & _
                												 " OR ( fLASTREM = 105.00 AND fPENULTREM = 84.40 ) " & _
                												 " OR ( fLASTREM = 1751.20 AND fPENULTREM = 1477.20 ) " & _
                												 " OR ( fLASTREM = 105.00 AND fPENULTREM = 84.40 ) " & _
                												 " OR ( fLASTREM = 20833.50 AND fPENULTREM = 16666.80 )) "
                sqlValue = 6
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
                
                'HIT
                queryString = " SELECT  COUNT(*) FROM HIT WHERE fBASE = " &  insGrISN & _
                                         " AND (( fCURSUM = 9.10 AND fTYPE = 'N2' ) " & _
                    								     " OR ( fCURSUM = 0.70 AND fTYPE = 'NH' ) " & _
                    										 " OR ( fCURSUM = 264.90 AND fTYPE = 'N2' ) " & _
                    										 " OR ( fCURSUM = 19.90 AND fTYPE = 'NH' )) "
                sqlValue = 4
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If 
                   
      ' Մուտք Հաճախորդի սպասարկում և դրամարկղ
      Call ChangeWorkspace(c_CustomerService)
      ' Մուտք Աշխատանքային փաստաթղթեր թղթապանակ
      Call wTreeView.DblClickItem("|Ð³×³Ëáñ¹Ç ëå³ë³ñÏáõÙ ¨ ¹ñ³Ù³ñÏÕ |²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ") 
      
      ' Ժամանակահատվածի սկիզբ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "PERN", mDate )
      ' Ժամանակահատվածի ավարտ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "PERK", dateAgr )
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
        
      ' Հիշարար օրդերի ստեղծում  34
      ordDate = "181214 "
      ordMoney = "3000"
      Call CreateMemOrders(orderISN, orderNum, ordDate, AccDb, AccCr, ordMoney, aim)       
      Log.Message(orderISN)
      BuiltIn.Delay(700)
         
               'HI
                queryString =  " SELECT COUNT(*) FROM HI WHERE fBASE = " & orderISN & _ 
                                          " AND (( fSUM = 3000.00 AND fCURSUM = 3000.00  AND fDBCR = 'D' ) " & _
											                    " OR	( fSUM = 3000.00 AND fCURSUM = 3000.00  AND fDBCR = 'C' )) "
                sqlValue = 2
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
           
                'HIRREST
                queryString = " SELECT COUNT(*)  FROM HIRREST WHERE fOBJECT =  " &  fISN & _
                                         " AND (( fLASTREM = 50000.00 AND fPENULTREM = 0 AND fSTARTREM = 0 ) " & _
                  											 " OR ( fLASTREM = 1751.20 AND fPENULTREM = 1477.20 AND fSTARTREM = 0 ) " & _
                  											 " OR ( fLASTREM = 105.00 AND fPENULTREM = 84.40 AND fSTARTREM = 0 ) " & _
                  											 " OR ( fLASTREM = 1751.20 AND fPENULTREM = 1477.20 AND fSTARTREM = 0 ) " & _
                  											 " OR ( fLASTREM = 105.00 AND fPENULTREM = 84.40 AND fSTARTREM = 0 ) " & _
                  											 " OR ( fLASTREM = 20833.50 AND fPENULTREM = 16666.80 AND fSTARTREM = 0 )) "
                sqlValue = 6 
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If 
       
      ' Մուտք գործել Օվերդրաֆտ (տեղաբաշխված) ԱՇՏ
      Call ChangeWorkspace(c_Overdraft)
      ' Մուտք Օվերդրաֆտ ունեցող հաշիվներ թղթապանակ
      Call wTreeView.DblClickItem("|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|úí»ñ¹ñ³ýï áõÝ»óáÕ Ñ³ßÇíÝ»ñ") 
          
      ' Ստուգում որ Օվերդրաֆտներ դիալոգը բացվել է
      If Not Sys.Process("Asbank").VBObject("frmAsUstPar").Exists Then
              Log.Error("Օվերդրաֆտներ դիալոգը չի բացվել")
              Exit Sub
      End If
        
      BuiltIn.Delay(700)
      ' Պայմանագրի համար դաշտի լրացում
      Call Rekvizit_Fill("Dialog" ,1, "General", "NUM", docNum)
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
           
      BuiltIn.Delay(4000)
      ' Պայմանագրի առկայության ստուգում Օվերդրաֆտներ թղթապանակում
      If tdbgViewn.ApproxCount <> 1 Then
              Log.Error("Պայմանագիրն առկա չէ Օվերդրաֆտներ թղթապանակում")
              Exit Sub
      End If
        
      ' Օվերդրաֆտի խմբային մարում 35
      endDate = "181214"
      Call OverdraftGroupRepayment(ordDate, endDate)
      ' Փակել Օվերդրաֆտներ թղթապանակը
      BuiltIn.Delay(700)
      frmPttel.Close
      
      ' Մուտք պայմանագրեր թղթապանակ
      Call wTreeView.DblClickItem("|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|ä³ÛÙ³Ý³·ñ»ñ")
      ' Պայմանագրի համար դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "NUM", docNum )
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
      
      ' Խմբային տոկոսների հաշվարկի կատարում 36
      calcDate = "180115"
      regDate = "180115"
      Call InterestGroupCalculationOverdraft (calcDate, regDate, checkCount)
      
      ' Խմբային տոկոսների հաշվարկ փաստաթղթի ISN - ի ստացում
      Call GetDocISN(paramN, calcDate, status, dateType, insGrISN)
      Log.Message(insGrISN)
      
               'HI
                queryString = " SELECT  SUM(fSUM) FROM HI WHERE fBASE = " &  insGrISN 
                sqlValue = 1066.60
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
 
                'HIF
                queryString = " SELECT COUNT(*) FROM HIF WHERE fBASE = " & insGrISN & _
                                         " AND fSUM = 0 AND fCURSUM = 0 "
                sqlValue = 5
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If       

                'HIR
                queryString = " SELECT COUNT(*) FROM HIR WHERE fBASE = " &  insGrISN & _
                                         " AND (( fCURSUM = 8.00 AND fTYPE = 'R2' ) " & _
                      									 " OR ( fCURSUM = 0.70 AND fTYPE = 'RH' ) " & _
                      									 " OR ( fCURSUM = 247.70 AND fTYPE = 'R2' ) " & _
                      									 " OR ( fCURSUM = 21.20 AND fTYPE = 'RH' ) " & _
                      									 " OR ( fCURSUM = 255.70 AND fTYPE = 'R¸' ) " & _
                      									 " OR ( fCURSUM = 21.90 AND fTYPE = 'RÂ' ) " & _
                      									 " OR ( fCURSUM = 4166.70 AND fTYPE = 'RÄ' )) "
                sqlValue = 7
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
 
                'HIRREST
                queryString = " SELECT COUNT(*) FROM HIRREST WHERE fOBJECT = " & fISN & _
                                         " AND (( fLASTREM = 50000 AND fPENULTREM = 0 ) " & _
                												 " OR ( fLASTREM = 2006.90 AND fPENULTREM = 1751.20 ) " & _
                												 " OR ( fLASTREM = 126.90 AND fPENULTREM = 105.00 ) " & _
                												 " OR ( fLASTREM = 2006.90 AND fPENULTREM = 1751.20 ) " & _
                												 " OR ( fLASTREM = 126.90 AND fPENULTREM = 105.00 ) " & _
                												 " OR ( fLASTREM = 25000.20 AND fPENULTREM = 20833.50 )) "
                sqlValue = 6
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
                
                'HIT
                queryString = " SELECT  COUNT(*) FROM HIT WHERE fBASE = " &  insGrISN & _
                                         " AND (( fCURSUM = 8 AND fTYPE = 'N2' ) " & _
                    										 " OR ( fCURSUM = 0.70 AND fTYPE = 'NH' ) " & _
                    										 " OR ( fCURSUM = 247.70 AND fTYPE = 'N2' ) " & _
                    										 " OR ( fCURSUM = 21.20 AND fTYPE = 'NH' )) "
                sqlValue = 4
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If 
      
      ' Խմբային տոկոսների հաշվարկի կատարում 37
      calcDate = "190115"
      regDate = "190115"
      Call InterestGroupCalculationOverdraft (calcDate, regDate, checkCount)
      
      ' Խմբային տոկոսների հաշվարկ փաստաթղթի ISN - ի ստացում
      Call GetDocISN(paramN, calcDate, status, dateType, insGrISN)
      Log.Message(insGrISN)
      
               'HI
                queryString = " SELECT  SUM(fSUM) FROM HI WHERE fBASE = " &  insGrISN 
                sqlValue = 29
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
 
                'HIF
                queryString = " SELECT COUNT(*) FROM HIF WHERE fBASE = " & insGrISN & _
                                         " AND fSUM = 0 AND fCURSUM = 0 "
                sqlValue = 4
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If       

                'HIR
                queryString = " SELECT COUNT(*) FROM HIR WHERE fBASE = " &  insGrISN & _
                                         " AND (( fCURSUM = 6.90 AND fTYPE = 'R2' ) " & _
										                     " OR  ( fCURSUM = 0.70 AND fTYPE = 'RH' )) "
                sqlValue = 2
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
 
                'HIRREST
                queryString = " SELECT COUNT(*) FROM HIRREST WHERE fOBJECT = " & fISN & _
                                         " AND (( fLASTREM = 50000 AND fPENULTREM = 0 ) " & _
                												 " OR ( fLASTREM = 2013.80 AND fPENULTREM = 2006.90 ) " & _
                												 " OR ( fLASTREM = 127.60 AND fPENULTREM = 126.90 ) " & _
                												 " OR ( fLASTREM = 2006.90 AND fPENULTREM = 1751.20 ) " & _
                												 " OR ( fLASTREM = 126.90 AND fPENULTREM = 105.00 ) " & _
                												 " OR ( fLASTREM = 25000.20 AND fPENULTREM = 20833.50 )) "
                sqlValue = 6
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
                
                'HIT
                queryString = " SELECT  COUNT(*) FROM HIT WHERE fBASE = " &  insGrISN & _
                                         " AND (( fCURSUM = 6.90 AND fTYPE = 'N2' ) " & _
										                     " OR ( fCURSUM = 0.70 AND fTYPE = 'NH' )) "
                sqlValue = 2
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If 
      
      ' Խմբային տոկոսների հաշվարկի կատարում 38
      calcDate = "220115"
      regDate = "220115"
      Call InterestGroupCalculationOverdraft (calcDate, regDate, checkCount)
      
      ' Խմբային տոկոսների հաշվարկ փաստաթղթի ISN - ի ստացում
      Call GetDocISN(paramN, calcDate, status, dateType, insGrISN)
      Log.Message(insGrISN)
      frmPttel.Close
       
               'HI
                queryString = " SELECT  COUNT(*) FROM HI WHERE fBASE = " &  insGrISN  & _
                                         " AND (( fSUM = 20.5 AND fCURSUM = 20.5 AND fDBCR = 'C' ) " & _
                      									 " OR ( fSUM = 2.10 AND fCURSUM = 2.10 AND fDBCR = 'C' ) " & _
                      									 " OR ( fSUM = 20.5 AND fCURSUM = 20.5 AND fDBCR = 'D' ) " & _
                      									 " OR ( fSUM = 2.10 AND fCURSUM = 2.10 AND fDBCR = 'D' )) "
                sqlValue = 6
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
 
                'HIF
                queryString = " SELECT COUNT(*) FROM HIF WHERE fBASE = " & insGrISN & _
                                         " AND fSUM = 0 AND fCURSUM = 0 "
                sqlValue = 1
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If       

                'HIR
                queryString = " SELECT COUNT(*) FROM HIR WHERE fBASE = " &  insGrISN & _
                                         " AND (( fCURSUM = 20.50 AND fTYPE = 'R2' ) " & _
										                     " OR ( fCURSUM = 2.1 AND fTYPE = 'RH' )) "
                sqlValue = 2
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
 
                'HIRREST
                queryString = " SELECT COUNT(*) FROM HIRREST WHERE fOBJECT = " & fISN & _
                                         " AND (( fLASTREM = 50000 AND fPENULTREM = 0 ) " & _
                												 " OR ( fLASTREM = 2034.30 AND fPENULTREM = 2013.80 ) " & _
                												 " OR ( fLASTREM = 129.70 AND fPENULTREM = 127.60 ) " & _
                												 " OR ( fLASTREM = 2006.90 AND fPENULTREM = 1751.20 ) " & _
                												 " OR ( fLASTREM = 126.90 AND fPENULTREM = 105.00 ) " & _
                												 " OR ( fLASTREM = 25000.20 AND fPENULTREM = 20833.50 )) "
                sqlValue = 6
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
                
                'HIT
                queryString = " SELECT  COUNT(*) FROM HIT WHERE fBASE = " &  insGrISN & _
                                         " AND (( fCURSUM = 20.50 AND fTYPE = 'N2' ) " & _
										                     " OR ( fCURSUM = 2.1 AND fTYPE = 'NH' )) "
                sqlValue = 2
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If 
             
      ' Մուտք Հաճախորդի սպասարկում և դրամարկղ ԱՇՏ
      Call ChangeWorkspace(c_CustomerService)
      ' Մուտք Աշխատանքային փաստաթղթեր թղթապանակ
      Call wTreeView.DblClickItem("|Ð³×³Ëáñ¹Ç ëå³ë³ñÏáõÙ ¨ ¹ñ³Ù³ñÏÕ |²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ") 
      
      ' Ժամանակահատվածի սկիզբ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "PERN", mDate )
      ' Ժամանակահատվածի ավարտ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "PERK", dateAgr )
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
        
      ' Հիշարար օրդերի ստեղծում  39
      ordDate = "230115"
      ordMoney = "4166,70"
      Call CreateMemOrders(orderISN, orderNum, ordDate, AccDb, AccCr, ordMoney, aim)      
      Log.Message(orderISN)
      BuiltIn.Delay(700)
         
               'HI
                queryString =  " SELECT COUNT(*) FROM HI WHERE fBASE = " & orderISN & _ 
                                          " AND (( fSUM = 416670.00 AND fCURSUM = 416670.00  AND fDBCR = 'D' ) " & _
											                    " OR	( fSUM = 416670.00 AND fCURSUM = 416670.00  AND fDBCR = 'C' )) "
                sqlValue = 2
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
           
                'HIRREST
                queryString = " SELECT COUNT(*)  FROM HIRREST WHERE fOBJECT =  " &  fISN & _
                                         " AND (( fLASTREM = 50000.00 AND fPENULTREM = 0 AND fSTARTREM = 0 ) " & _
                  											 " OR ( fLASTREM = 2034.30 AND fPENULTREM = 2013.80 AND fSTARTREM = 0 ) " & _
                  											 " OR ( fLASTREM = 129.70 AND fPENULTREM = 127.60 AND fSTARTREM = 0 ) " & _
                  											 " OR ( fLASTREM = 2006.90 AND fPENULTREM = 1751.20 AND fSTARTREM = 0 ) " & _
                  											 " OR ( fLASTREM = 126.90 AND fPENULTREM = 105.00 AND fSTARTREM = 0 ) " & _
                  											 " OR ( fLASTREM = 25000.20 AND fPENULTREM = 20833.50 AND fSTARTREM = 0 )) "
                sqlValue = 6 
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If 
       
      ' Մուտք գործել Օվերդրաֆտ (տեղաբաշխված) ԱՇՏ
      Call ChangeWorkspace(c_Overdraft)
      ' Մուտք Օվեդրաֆտ ունեցող հաշիվներ թղթապանակ
      Call wTreeView.DblClickItem("|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|úí»ñ¹ñ³ýï áõÝ»óáÕ Ñ³ßÇíÝ»ñ") 
      
      ' Ստուգում որ Օվերդրաֆտներ դիալոգը բացվել է
      If Not Sys.Process("Asbank").VBObject("frmAsUstPar").Exists Then
              Log.Error("Օվերդրաֆտներ դիալոգը չի բացվել")
              Exit Sub
      End If
        
      BuiltIn.Delay(1000)
      ' Պայմանագրի համար դաշտի լրացում
      Call Rekvizit_Fill("Dialog" ,1, "General", "NUM", docNum)
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
        
      BuiltIn.Delay(4000)
      ' Պայմանագրի առկայության ստուգում Օվերդրաֆտներ թղթապանակում
      If tdbgViewn.ApproxCount <> 1 Then
              Log.Error("Պայմանագիրն առկա չէ Օվերդրաֆտներ թղթապանակում")
              Exit Sub
      End If
        
      ' Օվերդրաֆտի խմբային մարում 40
      endDate = "230115"
      Call OverdraftGroupRepayment(ordDate, endDate)
      ' Փակել Օվերդրաֆտներ թղթապանակը
      BuiltIn.Delay(2000)
      frmPttel.Close
              
      ' Մուտք պայմանագրեր թղթապանակ
      Call wTreeView.DblClickItem("|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|ä³ÛÙ³Ý³·ñ»ñ")
      ' Պայմանագրի համար դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "NUM", docNum )
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
      
      ' Խմբային տոկոսների հաշվարկի կատարում 41
      calcDate = "170215"
      regDate = "170215"
      Call InterestGroupCalculationOverdraft (calcDate, regDate, checkCount)
      
      ' Խմբային տոկոսների հաշվարկ փաստաթղթի ISN - ի ստացում
      Call GetDocISN(paramN, calcDate, status, dateType, insGrISN)
      Log.Message(insGrISN)
      
               'HI
                queryString = " SELECT  COUNT(*) FROM HI WHERE fBASE = " &  insGrISN  & _
                                         " AND (( fSUM = 178.10 AND fCURSUM = 178.10 AND fDBCR = 'C') " & _
                    										 " OR ( fSUM = 17.80 AND fCURSUM = 17.80 AND fDBCR = 'C' ) " & _
                    										 " OR ( fSUM = 178.10 AND fCURSUM = 178.10 AND fDBCR = 'D' ) " & _
                    										 " OR ( fSUM = 17.80 AND fCURSUM = 17.80 AND fDBCR = 'D' )) "
                sqlValue = 6
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
 
                'HIF
                queryString = " SELECT COUNT(*) FROM HIF WHERE fBASE = " & insGrISN & _
                                         " AND fSUM = 0 AND fCURSUM = 0 "
                sqlValue = 1
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If       

                'HIR
                queryString = " SELECT COUNT(*) FROM HIR WHERE fBASE = " &  insGrISN & _
                                         " AND (( fCURSUM = 178.10 AND fTYPE = 'R2' ) " & _
                      									 " OR  ( fCURSUM = 17.80 AND fTYPE = 'RH' ) " & _
                                         " OR  ( fCURSUM = 205.50 AND fTYPE = 'R¸' ) " & _
                      									 " OR  ( fCURSUM = 20.60 AND fTYPE = 'RÂ' ) " & _
                      									 " OR  ( fCURSUM = 4166.70 AND fTYPE = 'RÄ' )) "
                sqlValue = 5
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
 
                'HIRREST
                queryString = " SELECT COUNT(*) FROM HIRREST WHERE fOBJECT = " & fISN & _
                                         " AND (( fLASTREM = 50000 AND fPENULTREM = 0 ) " & _
                												 " OR ( fLASTREM = 2212.40 AND fPENULTREM = 2034.30 ) " & _
                											   " OR ( fLASTREM = 147.50 AND fPENULTREM = 129.70 ) " & _
                												 " OR ( fLASTREM = 2212.40 AND fPENULTREM = 2006.90 ) " & _
                												 " OR ( fLASTREM = 147.50 AND fPENULTREM = 126.90 ) " & _
                												 " OR ( fLASTREM = 29166.90 AND fPENULTREM = 25000.20 )) "
                sqlValue = 6
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
                
                'HIT
                queryString = " SELECT  COUNT(*) FROM HIT WHERE fBASE = " &  insGrISN & _
                                         " AND (( fCURSUM = 178.10 AND fTYPE = 'N2' ) " & _
										                     " OR ( fCURSUM = 17.80 AND fTYPE = 'NH' )) "
                sqlValue = 2
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If 
      
      ' Խմբային տոկոսների հաշվարկի կատարում 42
      calcDate = "180215"
      regDate = "180215"
      Call InterestGroupCalculationOverdraft (calcDate, regDate, checkCount)
      
      ' Խմբային տոկոսների հաշվարկ փաստաթղթի ISN - ի ստացում
      Call GetDocISN(paramN, calcDate, status, dateType, insGrISN)
      Log.Message(insGrISN)
      
               'HI
                queryString = " SELECT  COUNT(*) FROM HI WHERE fBASE = " &  insGrISN  & _
                                         " AND (( fSUM = 5.70 AND fCURSUM = 5.70 AND fDBCR = 'C' ) " & _
                      									 " OR ( fSUM = 0.70 AND fCURSUM = 0.70 AND fDBCR = 'C' ) " & _
                      									 " OR ( fSUM = 5.70 AND fCURSUM = 5.70 AND fDBCR = 'D' ) " & _
                      									 " OR ( fSUM = 0.70 AND fCURSUM = 0.70 AND fDBCR = 'D' )) "
                sqlValue = 6
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
 
                'HIF
                queryString = " SELECT COUNT(*) FROM HIF WHERE fBASE = " & insGrISN & _
                                         " AND fSUM = 0 AND fCURSUM = 0 "
                sqlValue = 4
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If       

                'HIR
                queryString = " SELECT COUNT(*) FROM HIR WHERE fBASE = " &  insGrISN & _
                                         " AND (( fCURSUM = 5.70 AND fTYPE = 'R2' ) " & _
										                     " OR ( fCURSUM = 0.70 AND fTYPE = 'RH' )) "
                sqlValue = 2
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
 
                'HIRREST
                queryString = " SELECT COUNT(*) FROM HIRREST WHERE fOBJECT = " & fISN & _
                                         " AND (( fLASTREM = 50000 AND fPENULTREM = 0 ) " & _
                												 " OR ( fLASTREM = 2218.10 AND fPENULTREM = 2212.40 ) " & _
                												 " OR ( fLASTREM = 148.20 AND fPENULTREM = 147.50 ) " & _
                												 " OR ( fLASTREM = 2212.40 AND fPENULTREM = 2006.90 ) " & _
                												 " OR ( fLASTREM = 147.50 AND fPENULTREM = 126.90 ) " & _
                												 " OR ( fLASTREM = 29166.90 AND fPENULTREM = 25000.20 )) "
                sqlValue = 6
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
                
                'HIT
                queryString = " SELECT  COUNT(*) FROM HIT WHERE fBASE = " &  insGrISN & _
                                         " AND (( fCURSUM = 5.70 AND fTYPE = 'N2' ) " & _
										                     " OR ( fCURSUM = 0.7 AND fTYPE = 'NH' )) "
                sqlValue = 2
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If 
      
      ' Խմբային տոկոսների հաշվարկի կատարում 43
      calcDate = "250215"
      regDate = "250215"
      Call InterestGroupCalculationOverdraft (calcDate, regDate, checkCount)
      
      ' Խմբային տոկոսների հաշվարկ փաստաթղթի ISN - ի ստացում
      Call GetDocISN(paramN, calcDate, status, dateType, insGrISN)
      Log.Message(insGrISN)
      frmPttel.Close
       
               'HI
                queryString = " SELECT  COUNT(*) FROM HI WHERE fBASE = " &  insGrISN  & _
                                         " AND (( fSUM = 40.00 AND fCURSUM = 40.00 AND fDBCR = 'C' ) " & _
                      									 " OR ( fSUM = 4.80 AND fCURSUM = 4.80 AND fDBCR = 'C' ) " & _
                      									 " OR ( fSUM = 40.00 AND fCURSUM = 40.00 AND fDBCR = 'D' ) " & _
                      									 " OR ( fSUM = 4.80 AND fCURSUM = 4.80 AND fDBCR = 'D' )) "
                sqlValue = 6
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
 
                'HIF
                queryString = " SELECT COUNT(*) FROM HIF WHERE fBASE = " & insGrISN & _
                                         " AND fSUM = 0 AND fCURSUM = 0 "
                sqlValue = 1
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If       

                'HIR
                queryString = " SELECT COUNT(*) FROM HIR WHERE fBASE = " &  insGrISN & _
                                         " AND (( fCURSUM = 40.00 AND fTYPE = 'R2' ) " & _
										                     " OR ( fCURSUM = 4.80 AND fTYPE = 'RH' )) "
                sqlValue = 2
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
 
                'HIRREST
                queryString = " SELECT COUNT(*) FROM HIRREST WHERE fOBJECT = " & fISN & _
                                         " AND (( fLASTREM = 50000 AND fPENULTREM = 0 ) " & _
                												 " OR ( fLASTREM = 2258.10 AND fPENULTREM = 2218.10 ) " & _
                												 " OR ( fLASTREM = 153.00 AND fPENULTREM = 148.20 ) " & _
                												 " OR ( fLASTREM = 2212.40 AND fPENULTREM = 2006.90 ) " & _
                												 " OR ( fLASTREM = 147.50 AND fPENULTREM = 126.90 ) " & _
                												 " OR ( fLASTREM = 29166.90 AND fPENULTREM = 25000.20 )) "
                sqlValue = 6
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
                
                'HIT
                queryString = " SELECT  COUNT(*) FROM HIT WHERE fBASE = " &  insGrISN & _
                                         " AND (( fCURSUM = 40 AND fTYPE = 'N2' ) " & _
										                     " OR ( fCURSUM = 4.80 AND fTYPE = 'NH' )) "
                sqlValue = 2
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If 
      
      ' Մուտք Հաճախորդի սպասարկում և դրամարկղ
      Call ChangeWorkspace(c_CustomerService)
      ' Մուտք  Աշխատանքային փաստաթղթեր թղթապանակ
      Call wTreeView.DblClickItem("|Ð³×³Ëáñ¹Ç ëå³ë³ñÏáõÙ ¨ ¹ñ³Ù³ñÏÕ |²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ") 
      
      ' Ժամանակահատվածի սկիզբ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "PERN", mDate )
      ' Ժամանակահատվածի ավարտ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "PERK", dateAgr )
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
        
      ' Հիշարար օրդերի ստեղծում  44
      ordDate = "260215"
      ordMoney = "5000"
      Call CreateMemOrders(orderISN, orderNum, ordDate, AccDb, AccCr, ordMoney, aim)       
      Log.Message(orderISN)
      BuiltIn.Delay(700)
         
                'HI
                queryString =  " SELECT COUNT(*) FROM HI WHERE fBASE = " & orderISN & _ 
                                          " AND (( fSUM = 5000.00 AND fCURSUM = 5000.00  AND fDBCR = 'D' ) " & _
										                      " OR ( fSUM = 5000.00 AND fCURSUM = 5000.00  AND fDBCR = 'C' )) "
                sqlValue = 2
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
           
                'HIRREST
                queryString = " SELECT COUNT(*)  FROM HIRREST WHERE fOBJECT =  " &  fISN & _
                                         " AND (( fLASTREM = 50000.00 AND fPENULTREM = 0 AND fSTARTREM = 0 ) " & _
                  											 " OR ( fLASTREM = 2258.10 AND fPENULTREM = 2218.10 AND fSTARTREM = 0 ) " & _
                  											 " OR ( fLASTREM = 153.00 AND fPENULTREM = 148.20 AND fSTARTREM = 0 ) " & _
                  											 " OR ( fLASTREM = 2212.40 AND fPENULTREM = 2006.90 AND fSTARTREM = 0 ) " & _
                  											 " OR ( fLASTREM = 147.50 AND fPENULTREM = 126.90 AND fSTARTREM = 0 ) " & _
                  											 " OR ( fLASTREM = 29166.90 AND fPENULTREM = 25000.20 AND fSTARTREM = 0 )) "
                sqlValue = 6 
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If 

      ' Մուտք գործել Օվերդրաֆտ (տեղաբաշխված) ԱՇՏ
      Call ChangeWorkspace(c_Overdraft)
      ' Մուտք Օվերդրաֆտ ունեցող հաշիվներ
      Call wTreeView.DblClickItem("|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|úí»ñ¹ñ³ýï áõÝ»óáÕ Ñ³ßÇíÝ»ñ") 
      
      ' Ստուգում որ Օվերդրաֆտներ դիալոգը բացվել է
      If Not Sys.Process("Asbank").VBObject("frmAsUstPar").Exists Then
              Log.Error("Օվերդրաֆտներ դիալոգը չի բացվել")
              Exit Sub
      End If
        
      BuiltIn.Delay(700)
      ' Պայմանագրի համար դաշտի լրացում
      Call Rekvizit_Fill("Dialog" ,1, "General", "NUM", docNum)
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
        
      BuiltIn.Delay(4000)
      ' Պայմանագրի առկայության ստուգում Օվերդրաֆտներ թղթապանակում
      If tdbgViewn.ApproxCount <> 1 Then
              Log.Error("Պայմանագիրն առկա չէ Օվերդրաֆտներ թղթապանակում")
              Exit Sub
      End If
        
      ' Օվերդրաֆտի խմբային մարում 45
      endDate = "260215"
      Call OverdraftGroupRepayment(ordDate, endDate)
      ' Փակել Օվերդրաֆտներ թղթապանակը
      BuiltIn.Delay(700)
      frmPttel.Close
      
      ' Մուտք պայմանագրեր թղթապանակ
      Call wTreeView.DblClickItem("|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|ä³ÛÙ³Ý³·ñ»ñ")
      ' Պայմանագրի համար դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "NUM", docNum )
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
      
      ' Խմբային տոկոսների հաշվարկի կատարում 46
      calcDate = "170315"
      regDate = "170315"
      Call InterestGroupCalculationOverdraft (calcDate, regDate, checkCount)
      
      ' Խմբային տոկոսների հաշվարկ փաստաթղթի ISN - ի ստացում
      Call GetDocISN(paramN, calcDate, status, dateType, insGrISN)
      Log.Message(insGrISN)
     
               'HI
                queryString = " SELECT  COUNT(*) FROM HI WHERE fBASE = " &  insGrISN  & _
                                         " AND (( fSUM = 114.10 AND fCURSUM = 114.10 AND fDBCR = 'C' ) " & _
                    										 " OR ( fSUM = 13.70 AND fCURSUM = 13.70 AND fDBCR = 'C' ) " & _
                    										 " OR ( fSUM = 114.10 AND fCURSUM = 114.10 AND fDBCR = 'D' ) " & _
                    										 " OR ( fSUM = 13.70 AND fCURSUM = 13.70 AND fDBCR = 'D' )) "
                sqlValue = 6
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
 
                'HIF
                queryString = " SELECT COUNT(*) FROM HIF WHERE fBASE = " & insGrISN & _
                                         " AND fSUM = 0 AND fCURSUM = 0 "
                sqlValue = 1
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If       

                'HIR
                queryString = " SELECT COUNT(*) FROM HIR WHERE fBASE = " &  insGrISN & _
                                         " AND (( fCURSUM = 114.10 AND fTYPE = 'R2' ) " & _
                      									 " OR ( fCURSUM = 13.70 AND fTYPE = 'RH' ) " & _
                      									 " OR ( fCURSUM = 159.80 AND fTYPE = 'R¸' ) " & _
                      									 " OR ( fCURSUM = 19.20 AND fTYPE = 'RÂ' ) " & _
                      									 " OR ( fCURSUM = 4166.70 AND fTYPE = 'RÄ' )) "
                sqlValue = 5
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
 
                'HIRREST
                queryString = " SELECT COUNT(*) FROM HIRREST WHERE fOBJECT = " & fISN & _
                                         " AND (( fLASTREM = 50000 AND fPENULTREM = 0 ) " & _
                												 " OR ( fLASTREM = 2372.20 AND fPENULTREM = 2258.10 ) " & _
                											   " OR ( fLASTREM = 166.70 AND fPENULTREM = 153.00 ) " & _
                												 " OR ( fLASTREM = 2372.20 AND fPENULTREM = 2212.40 ) " & _
                												 " OR ( fLASTREM = 166.70 AND fPENULTREM = 147.50 ) " & _
                												 " OR ( fLASTREM = 33333.60 AND fPENULTREM = 29166.90 )) "
                sqlValue = 6
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
                
                'HIT
                queryString = " SELECT  COUNT(*) FROM HIT WHERE fBASE = " &  insGrISN & _
                                         " AND (( fCURSUM = 114.10 AND fTYPE = 'N2' ) " & _
										                     " OR ( fCURSUM = 13.70 AND fTYPE = 'NH' )) "
                sqlValue = 2
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If 
      
      ' Խմբային տոկոսների հաշվարկի կատարում 47
      calcDate = "180315"
      regDate = "180315"
      Call InterestGroupCalculationOverdraft (calcDate, regDate, checkCount)
      
      ' Խմբային տոկոսների հաշվարկ փաստաթղթի ISN - ի ստացում
      Call GetDocISN(paramN, calcDate, status, dateType, insGrISN)
      Log.Message(insGrISN)
      
               'HI
                queryString = " SELECT  COUNT(*) FROM HI WHERE fBASE = " &  insGrISN  & _
                                         " AND (( fSUM = 4.60 AND fCURSUM = 4.60 AND fDBCR = 'C' ) " & _
                      									 " OR ( fSUM = 0.70 AND fCURSUM = 0.70 AND fDBCR = 'C' ) " & _
                      									 " OR ( fSUM = 4.60 AND fCURSUM = 4.60 AND fDBCR = 'D' ) " & _
                      									 " OR ( fSUM = 0.70 AND fCURSUM = 0.70 AND fDBCR = 'D' )) "
                sqlValue = 6
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
 
                'HIF
                queryString = " SELECT COUNT(*) FROM HIF WHERE fBASE = " & insGrISN & _
                                         " AND fSUM = 0 AND fCURSUM = 0 "
                sqlValue = 4
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If       

                'HIR
                queryString = " SELECT COUNT(*) FROM HIR WHERE fBASE = " &  insGrISN & _
                                         " AND (( fCURSUM = 4.60 AND fTYPE = 'R2' ) " & _
										                     " OR ( fCURSUM = 0.7 AND fTYPE = 'RH' )) "
                sqlValue = 2
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
 
                'HIRREST
                queryString = " SELECT COUNT(*) FROM HIRREST WHERE fOBJECT = " & fISN & _
                                         " AND (( fLASTREM = 50000 AND fPENULTREM = 0 ) " & _
                  											 " OR ( fLASTREM = 2376.80 AND fPENULTREM = 2372.20 ) " & _
                  											 " OR ( fLASTREM = 167.40 AND fPENULTREM = 166.70 ) " & _
                  											 " OR ( fLASTREM = 2372.20 AND fPENULTREM = 2212.40 ) " & _
                  											 " OR ( fLASTREM = 166.70 AND fPENULTREM = 147.50 ) " & _
                  											 " OR ( fLASTREM = 33333.60 AND fPENULTREM = 29166.90 )) "
                sqlValue = 6
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
                
                'HIT
                queryString = " SELECT  COUNT(*) FROM HIT WHERE fBASE = " &  insGrISN & _
                                         " AND (( fCURSUM = 4.60 AND fTYPE = 'N2' ) " & _
										                     " OR ( fCURSUM = 0.70 AND fTYPE = 'NH' )) "
                sqlValue = 2
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If 
    
      ' Խմբային տոկոսների հաշվարկի կատարում 48
      calcDate = "190415"
      regDate = "190415"
      Call InterestGroupCalculationOverdraft (calcDate, regDate, checkCount)
      
      ' Խմբային տոկոսների հաշվարկ փաստաթղթի ISN - ի ստացում
      Call GetDocISN(paramN, calcDate, status, dateType, insGrISN)
      Log.Message(insGrISN)
      frmPttel.Close
       
               'HI
                queryString = " SELECT  COUNT(*) FROM HI WHERE fBASE = " &  insGrISN  & _
                                         " AND (( fSUM = 146.10 AND fCURSUM = 146.10 AND fDBCR = 'C' ) " & _
                    										 " OR ( fSUM = 21.90 AND fCURSUM = 21.90 AND fDBCR = 'C' ) " & _
                    										 " OR ( fSUM = 146.10 AND fCURSUM = 146.10 AND fDBCR = 'D' ) " & _
                    										 " OR ( fSUM = 21.90 AND fCURSUM = 21.90 AND fDBCR = 'D' )) "
                sqlValue = 6
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
 
                'HIF
                queryString = " SELECT COUNT(*) FROM HIF WHERE fBASE = " & insGrISN & _
                                         " AND fSUM = 0 AND fCURSUM = 0 "
                sqlValue = 1
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If       

                'HIR
                queryString = " SELECT COUNT(*) FROM HIR WHERE fBASE = " &  insGrISN & _
                                         " AND (( fCURSUM = 146.10 AND fTYPE = 'R2' ) " & _
                      									 " OR ( fCURSUM = 21.90 AND fTYPE = 'RH' ) " & _
                      									 " OR ( fCURSUM = 150.70 AND fTYPE = 'R¸' ) " & _
                      									 " OR ( fCURSUM = 22.60 AND fTYPE = 'RÂ' ) " & _
                      									 " OR ( fCURSUM = 4166.70 AND fTYPE = 'RÄ' )) "
                sqlValue = 5
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
 
                'HIRREST
                queryString = " SELECT COUNT(*) FROM HIRREST WHERE fOBJECT = " & fISN & _
                                         " AND (( fLASTREM = 50000 AND fPENULTREM = 0) " & _
                												 " OR ( fLASTREM = 2522.90 AND fPENULTREM = 2376.80 ) " & _
                												 " OR ( fLASTREM = 189.30 AND fPENULTREM = 167.40 ) " & _
                												 " OR ( fLASTREM = 2522.90 AND fPENULTREM = 2372.20 ) " & _
                												 " OR ( fLASTREM = 189.30 AND fPENULTREM = 166.70 ) " & _
                												 " OR ( fLASTREM = 37500.30 AND fPENULTREM = 33333.60 )) "
                sqlValue = 6
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
                
                'HIT
                queryString = " SELECT  COUNT(*) FROM HIT WHERE fBASE = " &  insGrISN & _
                                         " AND (( fCURSUM = 146.10  AND fTYPE = 'N2' ) " & _
										                     " OR ( fCURSUM = 21.9 AND fTYPE = 'NH' )) "
                sqlValue = 2
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If 
       
      ' Մուտք Հաճախորդի սպասարկում և դրամարկղ ԱՇՏ
      Call ChangeWorkspace(c_CustomerService)
      ' Մուտք Աշխատանքային փաստաթղթեր թղթապանակ
      Call wTreeView.DblClickItem("|Ð³×³Ëáñ¹Ç ëå³ë³ñÏáõÙ ¨ ¹ñ³Ù³ñÏÕ |²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ") 
      
      ' Ժամանակահատվածի սկիզբ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "PERN", mDate )
      ' Ժամանակահատվածի ավարտ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "PERK", dateAgr )
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
        
      ' Հիշարար օրդերի ստեղծում  49
      ordDate = "200415"
      ordMoney = "4000"
      Call CreateMemOrders(orderISN, orderNum, ordDate, AccDb, AccCr, ordMoney, aim)       
      Log.Message(orderISN)
      BuiltIn.Delay(700)
       
                'HI
                queryString =  " SELECT COUNT(*) FROM HI WHERE fBASE = " & orderISN & _ 
                                          " AND (( fSUM = 4000.00 AND fCURSUM = 4000.00  AND fDBCR = 'D' ) " & _
										                      " OR ( fSUM = 4000.00 AND fCURSUM = 4000.00  AND fDBCR = 'C' )) "
                sqlValue = 2
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
           
                'HIRREST
                queryString = " SELECT COUNT(*)  FROM HIRREST WHERE fOBJECT =  " &  fISN & _
                                         " AND (( fLASTREM = 50000.00 AND fPENULTREM = 0 AND fSTARTREM = 0) " & _
                  											 " OR ( fLASTREM = 2522.90 AND fPENULTREM = 2376.80 AND fSTARTREM = 0 ) " & _
                  											 " OR ( fLASTREM = 189.30 AND fPENULTREM = 167.40 AND fSTARTREM = 0 ) " & _
                  											 " OR ( fLASTREM = 2522.90 AND fPENULTREM = 2372.20 AND fSTARTREM = 0 ) " & _
                  											 " OR ( fLASTREM = 189.30 AND fPENULTREM = 166.70 AND fSTARTREM = 0 ) " & _
                  											 " OR ( fLASTREM = 37500.30 AND fPENULTREM = 33333.60 AND fSTARTREM = 0 )) "
                sqlValue = 6 
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If 
       
      ' Մուտք գործել Օվերդրաֆտ (տեղաբաշխված) ԱՇՏ
      Call ChangeWorkspace(c_Overdraft)
      ' Մուտք Օվերդրաֆտ ունեցող հաշիվներ թղթապանակ
      Call wTreeView.DblClickItem("|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|úí»ñ¹ñ³ýï áõÝ»óáÕ Ñ³ßÇíÝ»ñ") 
      
      ' Ստուգում որ Օվերդրաֆտներ դիալոգը բացվել է
      If Not Sys.Process("Asbank").VBObject("frmAsUstPar").Exists Then
              Log.Error("Օվերդրաֆտներ դիալոգը չի բացվել")
              Exit Sub
      End If
        
      BuiltIn.Delay(700)
      ' Պայմանագրի համար դաշտի լրացում
      Call Rekvizit_Fill("Dialog" ,1, "General", "NUM", docNum)
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
        
      BuiltIn.Delay(4000)
      ' Պայմանագրի առկայության ստուգում օվերդրաֆտներ թղթապանակում
      If tdbgViewn.ApproxCount <> 1 Then
              Log.Error("Օվերդրաֆտներ թղթապանակում բացակայում է պայմանագիրը")
              Exit Sub
      End If
        
      ' Օվերդրաֆտի խմբային մարում 50
      endDate = "200415"
      Call OverdraftGroupRepayment(ordDate, endDate)
      ' Փակել Օվերդրաֆտներ թղթապանակը
      BuiltIn.Delay(700)
      frmPttel.Close
      
      ' Մուտք պայմանագրեր թղթապանակ
      Call wTreeView.DblClickItem("|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|ä³ÛÙ³Ý³·ñ»ñ")
      ' Պայմանագրի համար դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "NUM", docNum )
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
      
      ' Խմբային տոկոսների հաշվարկի կատարում 51
      calcDate = "110515"
      regDate = "110515"
      Call InterestGroupCalculationOverdraft (calcDate, regDate, checkCount)
      
      ' Խմբային տոկոսների հաշվարկ փաստաթղթի ISN - ի ստացում
      Call GetDocISN(paramN, calcDate, status, dateType, insGrISN)
      Log.Message(insGrISN)
      frmPttel.Close

                'HI
                queryString = " SELECT  SUM(fSUM) FROM HI WHERE fBASE = " &  insGrISN  
                sqlValue = 331.40
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
 
                'HIF
                queryString = " SELECT COUNT(*) FROM HIF WHERE fBASE = " & insGrISN & _
                                         " AND fSUM = 0 AND fCURSUM = 0 "
                sqlValue = 5
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If       

                'HIR
                queryString = " SELECT COUNT(*) FROM HIR WHERE fBASE = " &  insGrISN & _
                                         " AND (( fCURSUM = 3.40 AND fTYPE = 'R2' ) " & _
                      									 " OR ( fCURSUM = 0.70 AND fTYPE = 'RH' ) " & _
                      									 " OR ( fCURSUM = 71.90 AND fTYPE = 'R2' ) " & _
                      									 " OR ( fCURSUM = 14.40 AND fTYPE = 'RH' )) "
                sqlValue = 4
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
 
                'HIRREST
                queryString = " SELECT COUNT(*) FROM HIRREST WHERE fOBJECT = " & fISN & _
                                         " AND (( fLASTREM = 50000 AND fPENULTREM = 0 ) " & _
                												 " OR ( fLASTREM = 2598.20 AND fPENULTREM = 2522.90 ) " & _
                												 " OR ( fLASTREM = 204.40 AND fPENULTREM = 189.30 ) " & _
                												 " OR ( fLASTREM = 2522.90 AND fPENULTREM = 2372.20 ) " & _
                												 " OR ( fLASTREM = 189.30 AND fPENULTREM = 166.70 ) " & _
                												 " OR ( fLASTREM = 37500.30 AND fPENULTREM = 33333.60 )) "
                sqlValue = 6
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
                
                'HIT
                queryString = " SELECT  COUNT(*) FROM HIT WHERE fBASE = " &  insGrISN & _
                                         " AND (( fCURSUM = 3.40 AND fTYPE = 'N2' ) " & _
                    								     " OR ( fCURSUM = 0.70 AND fTYPE = 'NH' ) " & _
                    										 " OR ( fCURSUM = 71.90 AND fTYPE = 'N2' ) " & _
                    										 " OR ( fCURSUM = 14.40 AND fTYPE = 'NH' )) "
                sqlValue = 4
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If 
       
      ' Մուտք Հաճախորդի սպասարկում և դրամարկղ ԱՇՏ
      Call ChangeWorkspace(c_CustomerService)
      ' Մուտք Աշխատանքային փաստաթղթեր թղթապանակ
      Call wTreeView.DblClickItem("|Ð³×³Ëáñ¹Ç ëå³ë³ñÏáõÙ ¨ ¹ñ³Ù³ñÏÕ |²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ") 
      
      ' Ժամանակահատվածի սկիզբ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "PERN", mDate )
      ' Ժամանակահատվածի ավարտ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "PERK", dateAgr )
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
        
      ' Հիշարար օրդերի ստեղծում  52
      ordDate = "120515"
      ordMoney = "4262.60"
      Call CreateMemOrders(orderISN, orderNum, ordDate, AccDb, AccCr, ordMoney, aim)        
      Log.Message(orderISN)
      BuiltIn.Delay(700)
 
      ' Մուտք գործել Օվերդրաֆտ (տեղաբաշխված) ԱՇՏ
      Call ChangeWorkspace(c_Overdraft)
      ' Մուտք Օվերդրաֆտ ունեցող հաշիվներ
      Call wTreeView.DblClickItem("|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|úí»ñ¹ñ³ýï áõÝ»óáÕ Ñ³ßÇíÝ»ñ") 
      
      ' Ստուգում որ Օվերդրաֆտներ դիալոգը բացվել է
      If Not Sys.Process("Asbank").VBObject("frmAsUstPar").Exists Then
              Log.Error("Օվերդրաֆտներ դիալոգը չի բացվել")
              Exit Sub
      End If
        
      BuiltIn.Delay(700)
      ' Պայմանագրի համար դաշտի լրացում
      Call Rekvizit_Fill("Dialog" ,1, "General", "NUM", docNum)
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
        
      BuiltIn.Delay(4000)
      ' Պայմանագրի առկայության ստուգում օվերդրաֆտներ թղթապանակում
      If tdbgViewn.ApproxCount <> 1 Then
              Log.Error("Պայմանագիրն առկա չէ օվերդրաֆտներ թղթապանակում")
              Exit Sub
      End If
        
      ' Օվերդրաֆտի խմբային մարում 53
      BuiltIn.Delay(4000)
      endDate = "120515"
      Call OverdraftGroupRepayment(ordDate, endDate)
      BuiltIn.Delay(2000)
      ' Փակել Օվերդրաֆտներ թղթապանակը
      frmPttel.Close
      
      ' Մուտք պայմանագրեր թղթապանակ
      Call wTreeView.DblClickItem("|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|ä³ÛÙ³Ý³·ñ»ñ")
      ' Պայմանագրի համար դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "NUM", docNum )
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
      BuiltIn.Delay(1500)
      
      ' Խմբային տոկոսների հաշվարկի կատարում 54
      calcDate = "150615"
      regDate = "150615"
      Call InterestGroupCalculationOverdraft (calcDate, regDate, checkCount)
      
      ' Խմբային տոկոսների հաշվարկ փաստաթղթի ISN - ի ստացում
      Call GetDocISN(paramN, calcDate, status, dateType, insGrISN)
      Log.Message(insGrISN)
      frmPttel.Close
      
                'HI
                queryString = " SELECT  SUM(fSUM) FROM HI WHERE fBASE = " &  insGrISN  
                sqlValue = 395.20
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
 
                'HIF
                queryString = " SELECT COUNT(*) FROM HIF WHERE fBASE = " & insGrISN & _
                                         " AND fSUM = 0 AND fCURSUM = 0 "
                sqlValue = 6
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If       

                'HIR
                queryString = " SELECT SUM(fCURSUM) FROM HIR WHERE fBASE = " &  insGrISN 
                sqlValue = 4392.6
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
 
                'HIRREST
                queryString = " SELECT COUNT(*) FROM HIRREST WHERE fOBJECT = " & fISN & _
                                         " AND (( fLASTREM = 50000 AND fPENULTREM = 0 ) " & _
                												 " OR ( fLASTREM = 2685.00 AND fPENULTREM = 2598.20 ) " & _
                												 " OR ( fLASTREM = 228.40 AND fPENULTREM = 204.40 ) " & _
                												 " OR ( fLASTREM = 2618.80 AND fPENULTREM = 2522.90 ) " & _
                												 " OR ( fLASTREM = 208.50 AND fPENULTREM = 189.30 ) " & _
                												 " OR ( fLASTREM = 41667.00 AND fPENULTREM = 37500.30 )) "
                sqlValue = 6
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
                
                'HIT
                queryString = " SELECT  COUNT(*) FROM HIT WHERE fBASE = " &  insGrISN & _
                                         " AND (( fCURSUM = 20.60 AND fTYPE = 'N2' ) " & _
                    										 " OR ( fCURSUM = 4.10 AND fTYPE = 'NH' ) " & _
                    										 " OR ( fCURSUM = 2.30 AND fTYPE = 'N2' ) " & _
                    										 " OR ( fCURSUM = 0.70 AND fTYPE = 'NH' ) " & _
                    										 " OR ( fCURSUM = 63.90 AND fTYPE = 'N2' ) " & _
                    										 " OR ( fCURSUM = 19.20 AND fTYPE = 'NH' )) "
                sqlValue = 6
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If 
      
      ' Մուտք Հաճախորդի սպասարկում և դրամարկղ ԱՇՏ
      Call ChangeWorkspace(c_CustomerService)
      ' Մուտք Աշխատանքային փաստաթղթեր թղթապանակ
      Call wTreeView.DblClickItem("|Ð³×³Ëáñ¹Ç ëå³ë³ñÏáõÙ ¨ ¹ñ³Ù³ñÏÕ |²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ") 
      
      ' Ժամանակահատվածի սկիզբ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "PERN", mDate )
      ' Ժամանակահատվածի ավարտ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "PERK", dateAgr )
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
        
      ' Հիշարար օրդերի ստեղծում  55
      ordDate = "160615"
      ordMoney = "4262.60"
      Call CreateMemOrders(orderISN, orderNum, ordDate, AccDb, AccCr, ordMoney, aim)       
      Log.Message(orderISN)
      BuiltIn.Delay(700)
         
                'HI
                queryString =  " SELECT COUNT(*) FROM HI WHERE fBASE = " & orderISN & _ 
                                          " AND (( fSUM = 4262.60 AND fCURSUM = 4262.60  AND fDBCR = 'D' ) " & _
										                      " OR ( fSUM = 4262.60 AND fCURSUM = 4262.60  AND fDBCR = 'C' )) "
                sqlValue = 2
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If
           
                'HIRREST
                queryString = " SELECT COUNT(*)  FROM HIRREST WHERE fOBJECT =  " &  fISN & _
                                         " AND (( fLASTREM = 50000.00 AND fPENULTREM = 0 AND fSTARTREM = 0) " & _
                  											 " OR ( fLASTREM = 2685.00 AND fPENULTREM = 2598.20 AND fSTARTREM = 0) " & _
                  											 " OR ( fLASTREM = 228.40 AND fPENULTREM = 204.40 AND fSTARTREM = 0) " & _
                  											 " OR ( fLASTREM = 2618.80 AND fPENULTREM = 2522.90 AND fSTARTREM = 0) " & _
                  											 " OR ( fLASTREM = 208.50 AND fPENULTREM = 189.30 AND fSTARTREM = 0) " & _
                  											 " OR ( fLASTREM = 41667.00 AND fPENULTREM = 37500.30 AND fSTARTREM = 0 )) "
                sqlValue = 6 
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If 
 
      ' Մուտք գործել Օվերդրաֆտ (տեղաբաշխված) ԱՇՏ
      Call ChangeWorkspace(c_Overdraft)
      ' Մուտք Օվերդրաֆտ ունեցող հաշիվներ
      Call wTreeView.DblClickItem("|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|úí»ñ¹ñ³ýï áõÝ»óáÕ Ñ³ßÇíÝ»ñ") 
      
      ' Ստուգում որ Օվերդրաֆտներ դիալոգը բացվել է
      If Not Sys.Process("Asbank").VBObject("frmAsUstPar").Exists Then
              Log.Error("Օվերդրաֆտներ դիալոգը չի բացվել")
              Exit Sub
      End If
        
      BuiltIn.Delay(700)
      ' Պայմանագրի համար դաշտի լրացում
      Call Rekvizit_Fill("Dialog" ,1, "General", "NUM", docNum)
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
        
      BuiltIn.Delay(4000)
      ' Պայմանագրի առկայության ստուգում օվերդրաֆտներ թղթապանակում
      If tdbgViewn.ApproxCount <> 1 Then
              Log.Error("Պայմանագիրն առկա չէ օվերդրաֆտներ թղթապանակում")
              Exit Sub
      End If
        
      ' Օվերդրաֆտի խմբային մարում 56
      endDate = "160615"
      Call OverdraftGroupRepayment(ordDate, endDate)
      ' Փակել Օվերդրաֆտներ թղթապանակը
      BuiltIn.Delay(700)
      frmPttel.Close
       
      ' Մուտք պայմանագրեր թղթապանակ
      Call wTreeView.DblClickItem("|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|ä³ÛÙ³Ý³·ñ»ñ")
      ' Պայմանագրի համար դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "NUM", docNum )
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
        
      BuiltIn.Delay(4000)
      ' Պայմանագրի առկայության ստուգում օվերդրաֆտներ թղթապանակում
      If tdbgViewn.ApproxCount <> 1 Then
              Log.Error("Պայմանագիրն առկա չէ օվերդրաֆտներ թղթապանակում")
              Exit Sub
      End If
        
      ' խմբային տոկոսների հաշվարկի ջնջում
      param = c_ViewEdit & "|" & c_Other & "|" & c_CalcDates
      status = False
      dateGive = "150615 "
      dateAgr = "150615"
      Call DeleteActionOverdraft(param , dateGive, dateAgr, status, dateType )
        
      ' խմբային տոկոսների հաշվարկի ջնջում
      dateGive = "110515"
      dateAgr = "110515"
      Call DeleteActionOverdraft(param , dateGive, dateAgr, status, dateType )
        
      ' խմբային տոկոսների հաշվարկի ջնջում
      dateGive = "190415 "
      dateAgr = "190415"
      Call DeleteActionOverdraft(param , dateGive, dateAgr, status, dateType )
        
      ' խմբային տոկոսների հաշվարկի ջնջում
      dateGive = "180315"
      dateAgr = "180315"
      Call DeleteActionOverdraft(param , dateGive, dateAgr, status, dateType )
        
      ' խմբային տոկոսների հաշվարկի ջնջում
      dateGive = "170315"
      dateAgr = "170315"
      Call DeleteActionOverdraft(param , dateGive, dateAgr, status, dateType )
        
      ' խմբային տոկոսների հաշվարկի ջնջում
      dateGive = "250215"
      dateAgr = "250215"
      Call DeleteActionOverdraft(param , dateGive, dateAgr, status, dateType )
        
      ' խմբային տոկոսների հաշվարկի ջնջում
      dateGive = "180215"
      dateAgr = "180215"
      Call DeleteActionOverdraft(param , dateGive, dateAgr, status, dateType )
        
      ' խմբային տոկոսների հաշվարկի ջնջում
      dateGive = "170215"
      dateAgr = "170215"
      Call DeleteActionOverdraft(param , dateGive, dateAgr, status, dateType )
        
      ' խմբային տոկոսների հաշվարկի ջնջում
      dateGive = "220115"
      dateAgr = "220115"
      Call DeleteActionOverdraft(param , dateGive, dateAgr, status, dateType )
        
      ' խմբային տոկոսների հաշվարկի ջնջում
      dateGive = "190115"
      dateAgr = "190115"
      Call DeleteActionOverdraft(param , dateGive, dateAgr, status, dateType )
        
      ' խմբային տոկոսների հաշվարկի ջնջում
      dateGive = "180115"
      dateAgr = "180115"
      Call DeleteActionOverdraft(param , dateGive, dateAgr, status, dateType )
        
      ' խմբային տոկոսների հաշվարկի ջնջում
      dateGive = "171214"
      dateAgr = "171214"
      Call DeleteActionOverdraft(param , dateGive, dateAgr, status, dateType )
        
      ' խմբային տոկոսների հաշվարկի ջնջում
      dateGive = "171114"
      dateAgr = "171114"
      Call DeleteActionOverdraft(param , dateGive, dateAgr, status, dateType )
        
      ' խմբային տոկոսների հաշվարկի ջնջում
      dateGive = "201014"
      dateAgr = "201014"
      Call DeleteActionOverdraft(param , dateGive, dateAgr, status, dateType )
       
      ' խմբային տոկոսների հաշվարկի ջնջում
      dateGive = "191014"
      dateAgr = "191014"
      Call DeleteActionOverdraft(param , dateGive, dateAgr, status, dateType )
        
      ' խմբային տոկոսների հաշվարկի ջնջում
      dateGive = "180914"
      dateAgr = "180914"
      Call DeleteActionOverdraft(param , dateGive, dateAgr, status, dateType )
        
      ' խմբային տոկոսների հաշվարկի ջնջում
      dateGive = "170914"
      dateAgr = "170914"
      Call DeleteActionOverdraft(param , dateGive, dateAgr, status, dateType )
        
      ' խմբային տոկոսների հաշվարկի ջնջում
      dateGive = "250814"
      dateAgr = "250814"
      Call DeleteActionOverdraft(param , dateGive, dateAgr, status, dateType )
        
      ' խմբային տոկոսների հաշվարկի ջնջում
      dateGive = "180814"
      dateAgr = "180814"
      Call DeleteActionOverdraft(param , dateGive, dateAgr, status, dateType )
        
      ' խմբային տոկոսների հաշվարկի ջնջում
      dateGive = "170814"
      dateAgr = "170814"
      Call DeleteActionOverdraft(param , dateGive, dateAgr, status, dateType )
        
      ' խմբային տոկոսների հաշվարկի ջնջում
      dateGive = "140814"
      dateAgr = "140814"
      Call DeleteActionOverdraft(param , dateGive, dateAgr, status, dateType )

      ' Օվերդրաֆտի տրամադրման ջնջում        
      status = True
      param = c_OpersView
      dateGive = "180714 "
      dateAgr = "180714"
      Call DeleteActionOverdraft(param , dateGive, dateAgr, status, dateType )
      frmPttel.Close
          
      ' Մուտք Հաճախորդի սպասարկում և դրամարկղ
      Call ChangeWorkspace(c_CustomerService)
          
      ' Հիշարար օրդերի ջնջում 
      creatDate = "160615"
      Call DeleteMemOrderFromRegPayment(creatDate) 
        
      ' Հիշարար օրդերի ջնջում
      creatDate = "120515"
      Call DeleteMemOrderFromRegPayment(creatDate) 
         
      ' Հիշարար օրդերի ջնջում 
      creatDate = "200415"       
      Call DeleteMemOrderFromRegPayment(creatDate) 
        
      ' Հիշարար օրդերի ջնջում 
      creatDate = "260215"
      Call DeleteMemOrderFromRegPayment(creatDate) 
         
      ' Հիշարար օրդերի ջնջում 
      creatDate = "230115"
      Call DeleteMemOrderFromRegPayment(creatDate) 
        
      ' Հիշարար օրդերի ջնջում 
      creatDate = "181214"
      Call DeleteMemOrderFromRegPayment(creatDate)  
        
      ' Հիշարար օրդերի ջնջում 
      creatDate = "181114"
      Call DeleteMemOrderFromRegPayment(creatDate) 
        
      ' Հիշարար օրդերի ջնջում 
      creatDate = "201014"
      Call DeleteMemOrderFromRegPayment(creatDate) 
        
      ' Հիշարար օրդերի ջնջում 
      creatDate = "260814"
      Call DeleteMemOrderFromRegPayment(creatDate)
        
      ' Հիշարար օրդերի ջնջում 
      creatDate = "150814"
      Call DeleteMemOrderFromRegPayment(creatDate) 
        
      ' Մուտք գործել Օվերդրաֆտ (տեղաբաշխված) ԱՇՏ
      Call ChangeWorkspace(c_Overdraft)
      ' Մուտք Պայմանագրեր թղթապանակ
      Call wTreeView.DblClickItem("|úí»ñ¹ñ³ýï (ï»Õ³µ³ßËí³Í)|ä³ÛÙ³Ý³·ñ»ñ") 
      
      BuiltIn.Delay(1200)
      ' Ստուգում որ Օվերդրաֆտներ դիալոգը բացվել է   
      If Not Sys.Process("Asbank").VBObject("frmAsUstPar").Exists Then
              Log.Error("Օվերդրաֆտներ դիալոգը չի բացվել")
              Exit Sub
      End If
        
      ' Պայմանագրի համար դաշտի լրացում
      Call Rekvizit_Fill("Dialog" ,1, "General", "NUM", docNum)
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")       
       
      ' Ջնջել պայմանագիրը
      Call DelDoc()
      frmPttel.Close
                     
      ' Փակել ծրագիրը
      Call Close_AsBank()

End Sub