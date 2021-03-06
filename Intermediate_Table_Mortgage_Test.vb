'USEUNIT Subsystems_SQL_Library
'USEUNIT Library_Common
'USEUNIT Library_CheckDB 
'USEUNIT Constants
'USEUNIT Payment_Except_Library
'USEUNIT Intermediate_Table_Library

'Test case Id 166601
 
 Sub Intermediate_Table_Mortgage_Test()
 
      Dim contractN, outerCode, pladgeNumber, secType, secN, custName, gridCustName, gridCheckbox,_
              plCurrency, inSumma, inCount, plComment, dateSealing, plOffice, plSect, acsType, gridCustomer,_
              gridName, plSubject, plOther, accBalances, withPer, correlation, mortgage, existenceRes, riskWeight,_
              wNote, wNote2, wNote3, plACRA, plNewRV, pprCode, closeDate
      
      Dim status
      
      Dim queryString, sqlValue, colNum, sql_isEqual 
              
      Dim  startDate, fDate      
      
      startDate = "20100101"
      fDate = "20250101"
      Call Initialize_AsBank("bank", startDate, fDate)
               
      fileName1 = Project.Path & "Stores\Intermediate table\ActualMortgageErrorNew.txt"
      ' Ջնջել ֆայլը
      aqFile.Delete(fileName1)
            
      ' Մուտք համակարգ
      Call Create_Connection()
      Login("ARMSOFT")
        
              'CONTRACTS
              queryString = " SELECT COUNT(*) FROM NAGRACCS "
              sqlValue = 49
              colNum = 0
              sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
              If Not sql_isEqual Then
                Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
              End If 
            
      ' Մուտք Միջ. Աղյուսակներից ներմուծման ԱՇՏ
      Call ChangeWorkspace(c_ImpTable)
      ' Մուտք վարկային պայմանագրերի բացում
      Call wTreeView.DblClickItem("|ØÇç. ²ÕÛáõë³ÏÝ»ñÇó Ý»ñÙáõÍÙ³Ý ²Þî|imN1 êï³óí³Í ·ñ³íÝ»ñÇ µ³óáõÙ")
      ' Այո կոճակի սեղմում
      Call ClickCmdButton(5, "²Ûá")
      
      savePath = Project.Path & "Stores\Intermediate table\"
      fName = "ActualMortgageErrorNew.txt"
      fileName2 = Project.Path & "Stores\Intermediate table\ExpectedMortgageErrorNew.txt"
      
      If  wMDIClient.WaitVBObject("FrmSpr",30000).Exists Then
            ' Հիշել քաղվածքը
            Call SaveDoc(savePath, fName)

            If NOT Files.Compare(fileName1 , fileName2)  Then
                  Log.Warning("Ֆայլերը նման չեն")
            End If
            
            BuiltIn.Delay(1000)
            wMDIClient.WaitVBObject("FrmSpr",2000).Close
      Else
             Log.Error("Սխալի հաղորդագրության պատուհանը չի բացվել")
            Exit Sub
      End If
      
             'CONTRACTS
              queryString = " SELECT COUNT(*) FROM NAGRACCS "
              sqlValue = 84
              colNum = 0
              sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
              If Not sql_isEqual Then
                Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
              End If 

      ' Մուտք գործել վարկեր տեղաբաշխված/ Պայմանագրեր
      Call ChangeWorkspace(c_Loans)
      Call wTreeView.DblClickItem("|ì³ñÏ»ñ (ï»Õ³µ³ßËí³Í)|ä³ÛÙ³Ý³·ñ»ñ")

      If Not Sys.Process("Asbank").VBObject("frmAsUstPar").Exists Then
            Log.Error("Հաճախորդներ դիալոգը չի բացվել")
            Exit Sub 
      End If
      
      contractN = "A000048"
      outerCode ="645003003318L001"
      
      '  Պայմանագրի N դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "NUM", contractN)   
      '  Արտաքին N դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "OUTERCODE", outerCode)   
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
        
      If Not wMDIClient.WaitVBObject("frmPttel",2000).Exists Then
            Log.Error("Հաճախորդներ թղթապանակը չի բացվել")
            Exit Sub 
      ElseIf wMDIClient.VBObject("frmPttel").VBObject("tdbgView").ApproxCount <> 1 Then
              Log.Error("Նման համարով հաճախորդ գոյություն չունի")
              Exit Sub
      End If   
      
      ' Գրավի պայմանագրի տվյալների ստուգում սաբի կանչ
      pladgeNumber  = "¶ñ³íÇ å³ÛÙ³Ý³·Çñ- N10226         A000048"
      secType = "4"
      secN = "N10226"
      custName = "Ð³×³Ëáñ¹¨300331"
      gridCustName = "Ð³×³Ëáñ¹¨300331"          
      gridCheckbox = "0"
      plCurrency = "000"
      inSumma= "269,000.00"
      inCount = "1.00"
      plComment = ""
      dateSealing = "23/08/19"
      plOffice = "00"
      plSect = "1"
      acsType = "N10"
      gridCustomer = ""
      gridName = ""
      plSubject = "áëÏÇ"
      plOther = ""
      accBalances = "1"
      withPer = "1"
      correlation = "100.00"
      mortgage = "0"
      existenceRes = "0"
      riskWeight = "100.00"
      wNote = ""
      wNote2 = ""
      wNote3 = ""
      plACRA = "18"
      plNewRV = "1"
      pprCode = ""
      closeDate = "  /  /  "

      Call MortgageContract (contractN, pladgeNumber, secType, secN, custName, gridCustName, gridCheckbox,_ 
                                                plCurrency, inSumma, inCount, plComment, dateSealing, plOffice, plSect, acsType, gridCustomer,_
                                                gridName, plSubject, plOther, accBalances, withPer, correlation, mortgage, existenceRes, riskWeight,_
                                                wNote, wNote2, wNote3, plACRA, plNewRV, pprCode, closeDate)
        
      ' Մուտք գործել վարկեր տեղաբաշխված/ Պայմանագրեր                                        
      Call wTreeView.DblClickItem("|ì³ñÏ»ñ (ï»Õ³µ³ßËí³Í)|ä³ÛÙ³Ý³·ñ»ñ")

      If Not p1.WaitVBObject("frmAsUstPar",2000).Exists Then
            Log.Error("Հաճախորդներ դիալոգը չի բացվել")
            Exit Sub 
      End If
      
      contractN = "A000077"
      outerCode ="645003003391L001"
      
      '  Պայմանագրի N դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "NUM", contractN)   
      '  Արտաքին N դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "OUTERCODE", outerCode)   
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
        
      If Not wMDIClient.WaitVBObject("frmPttel",2000).Exists Then
            Log.Error("Հաճախորդներ թղթապանակը չի բացվել")
            Exit Sub 
      ElseIf wMDIClient.VBObject("frmPttel").VBObject("tdbgView").ApproxCount <> 1 Then
              Log.Error("Նման համարով հաճախորդ գոյություն չունի")
              Exit Sub
      End If 
                                                
      ' Գրավի պայմանագրի տվյալների ստուգում սաբի կանչ
      pladgeNumber  = "¶ñ³íÇ å³ÛÙ³Ý³·Çñ- N10258         A000077"
      secType = "5"
      secN = "N10258"
      custName = "Ð³×³Ëáñ¹¨300337"
      gridCustName = "Ð³×³Ëáñ¹¨300339"          
      gridCheckbox = "0"
      plCurrency = "000"
      inSumma= "149,300.00"
      inCount = "1.00"
      plComment = ""
      dateSealing = "26/08/19"
      plOffice = "00"
      plSect = "1"
      acsType = "N10"
      gridCustomer = ""
      gridName = ""
      plSubject = "Ù»ù»Ý³"
      plOther = "µ»éÝ³ï³ñ"
      accBalances = "0"
      withPer = "0"
      correlation = "99.00"
      mortgage = "0"
      existenceRes = "0"
      riskWeight = "100.00"
      wNote = ""
      wNote2 = ""
      wNote3 = ""
      plACRA = "06"
      plNewRV = "1"
      pprCode = "333333"
      closeDate = "  /  /  "

      Call MortgageContract (contractN, pladgeNumber, secType, secN, custName, gridCustName, gridCheckbox,_ 
                                                plCurrency, inSumma, inCount, plComment, dateSealing, plOffice, plSect, acsType, gridCustomer,_
                                                gridName, plSubject, plOther, accBalances, withPer, correlation, mortgage, existenceRes, riskWeight,_
                                                wNote, wNote2, wNote3, plACRA, plNewRV, pprCode, closeDate)
                                                
      ' Մուտք գործել վարկեր տեղաբաշխված/ Պայմանագրեր
      Call wTreeView.DblClickItem("|ì³ñÏ»ñ (ï»Õ³µ³ßËí³Í)|ä³ÛÙ³Ý³·ñ»ñ")

      If Not p1.WaitVBObject("frmAsUstPar",2000).Exists Then
            Log.Error("Հաճախորդներ դիալոգը չի բացվել")
            Exit Sub 
      End If
      
      contractN = "A000076"
      outerCode ="645003003383L001"
      
      '  Պայմանագրի N դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "NUM", contractN)   
      '  Արտաքին N դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "OUTERCODE", outerCode)   
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
        
      If Not wMDIClient.WaitVBObject("frmPttel",2000).Exists Then
            Log.Error("Հաճախորդներ թղթապանակը չի բացվել")
            Exit Sub 
      ElseIf wMDIClient.VBObject("frmPttel").VBObject("tdbgView").ApproxCount <> 1 Then
              Log.Error("Նման համարով հաճախորդ գոյություն չունի")
              Exit Sub
      End If 
                                                
      ' Գրավի պայմանագրի տվյալների ստուգում սաբի կանչ
      pladgeNumber  = "¶ñ³íÇ å³ÛÙ³Ý³·Çñ- N10259         A000076"
      secType = "3"
      secN = "N10259"
      custName = "Ð³×³Ëáñ¹¨300338"
      gridCustName = "Ð³×³Ëáñ¹¨300338"          
      gridCheckbox = "0"
      plCurrency = "000"
      inSumma= "300,100.00"
      inCount = "1.00"
      plComment = ""
      dateSealing = "26/08/19"
      plOffice = "00"
      plSect = "3"
      acsType = "N10"
      gridCustomer = ""
      gridName = ""
      plSubject = "·áõÛù"
      plOther = "ïáõÝ 526"
      accBalances = "0"
      withPer = "0"
      correlation = "100.00"
      mortgage = "0"
      existenceRes = "0"
      riskWeight = "100.00"
      wNote = ""
      wNote2 = ""
      wNote3 = ""
      plACRA = "03"
      plNewRV = "1"
      pprCode = "222222"
      closeDate = "  /  /  "

      Call MortgageContract (contractN, pladgeNumber, secType, secN, custName, gridCustName, gridCheckbox,_ 
                                                plCurrency, inSumma, inCount, plComment, dateSealing, plOffice, plSect, acsType, gridCustomer,_
                                                gridName, plSubject, plOther, accBalances, withPer, correlation, mortgage, existenceRes, riskWeight,_
                                                wNote, wNote2, wNote3, plACRA, plNewRV, pprCode, closeDate)
      
      ' Մուտք պայմանագրեր թղթապանակ                                  
      Call wTreeView.DblClickItem("|ì³ñÏ»ñ (ï»Õ³µ³ßËí³Í)|ä³ÛÙ³Ý³·ñ»ñ")

      If Not p1.WaitVBObject("frmAsUstPar",2000).Exists Then
            Log.Error("Հաճախորդներ դիալոգը չի բացվել")
            Exit Sub 
      End If
      
      contractN = "A000039"
      outerCode = "645003002807L001"
      
      '  Պայմանագրի N դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "NUM", contractN)   
      '  Արտաքին N դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "OUTERCODE", outerCode)   
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
        
      If Not wMDIClient.WaitVBObject("frmPttel",2000).Exists Then
            Log.Error("Հաճախորդներ թղթապանակը չի բացվել")
            Exit Sub 
      ElseIf wMDIClient.VBObject("frmPttel").VBObject("tdbgView").ApproxCount <> 1 Then
              Log.Error("Նման համարով հաճախորդ գոյություն չունի")
              Exit Sub
      End If 
      
      pladgeNumber = "N10239 (Ð³×³Ëáñ¹¨300280) ,     159000 -Ð³ÛÏ³Ï³Ý ¹ñ³Ù -            1.00 Ñ³ï"
      
      If Not CheckPladgeCount(pladgeNumber) Then
          Log.Warning("Պայմանագրում առկա չէ  գրավի պայմանագիր")
          Exit Sub
      End If
      
      pladgeNumber = "N10225 (Ð³×³Ëáñ¹¨300280) ,     253000 -Ð³ÛÏ³Ï³Ý ¹ñ³Ù -            2.00 Ñ³ï"
      
      If Not CheckPladgeCount(pladgeNumber) Then
          Log.Warning("Պայմանագրում առկա չէ երկրորդ գրավի պայմանագիրը")
          Exit Sub
      End If
      
      wMDIClient.VBObject("frmPttel").Close
      
      ' Փակել ծրագիրը
      Call Close_AsBank()
      
End Sub
