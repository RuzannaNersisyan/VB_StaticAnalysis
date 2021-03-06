'USEUNIT Library_Common 
'USEUNIT RemoteService_Library 
'USEUNIT Library_CheckDB
'USEUNIT Constants
'USEUNIT Subsystems_SQL_Library
'USEUNIT Payment_Order_ConfirmPhases_Library 
'USEUNIT BankMail_Library
'USEUNIT Library_Contracts

'Test Case ID 165385

Sub RemProc_Pay1ToMy_New_Test()
      
      Dim paramName, paramValue, confPath, confInput, queryStrin, queryStr
      Dim direction, system, dirName, wState, wChildISN
      Dim todayD, todayDMY, cliCode, payerAcc, taxCode, clCount, rowCount
      Dim regNum, row1, row2, amount, curISO, cur, aim, msgType, status
      Dim workEnvName, workEnv, stRekName, endRekName, wStatus, isnRekName
      Dim sqlValue, sql_isEqual, colNum, queryString, delayTime
      Dim colN, fISN, action, doNum, doActio, comment
      Dim startDate, fDate, verifyDocuments
      
      startDate = "20030101"
      fDate = "20250101"
      Call Initialize_AsBank("bank", startDate, fDate)
               
      ' Մուտք համակարգ ARMSOFT օգտագործողով
      Call Create_Connection()
      Login("ARMSOFT")
      
      ' Պարամետրերի արժեքների ճշգրտում   
      paramName = "CBDATEMAXDIFF"
      paramValue = "0"
      Call  SetParameter(paramName, paramValue)
      
      paramName = "IBCBPROCINTERVAL"
      paramValue = "20"
      Call  SetParameter(paramName, paramValue)
      
      ' Կարգավորումների ներմուծում
      confPath = "X:\Testing\RemAutoProc\Verify1_Pay1ToMy_New.txt"
      confInput = Input_Config(confPath)
      If Not confInput Then
          Log.Error("Կարգավորումները չեն ներմուծվել")
          Exit Sub
      End If
      
      todayD = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%Y%m%d")
      todayDMY = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%d/%m/%y")
      cliCode =  "00000680"
      payerAcc = "7770000068020100"
      taxCode = "1234567891"
      regNum = ""
      row1 = "          300.001660000000451000poxancum"
      row2 = "          400.001660033120090800poxancum"        
      amount = "700.00"
      curISO = "AMD"
      cur = "000"
      aim = "test"
      msgType = "IBMlPOrd"
            
      ' Ձևավորում ենք docNum-ը պատահական գեներացված թվի դիմացից ավելացնելով 0-ներ, այնպես որ լինի 6 նիշ
      Call Randomize()
      docNum = right(String(6, "0") + RTrim(Int(1000 * Rnd)), 6) 
      Log.Message("Պայամանագրի համարը՝ " & docNum)
      
      ' Տվյալ օրով տվյալների ջնջում
      queryStr = "Delete from  CB_MESSAGES where FORMAT (fDATE, 'dd/MM/yy') = '" & Trim(todayDMY) & "' "  
      Call Execute_SLQ_Query(queryStr)
      BuiltIn.Delay(2000)  
        
      queryStrin = " Insert into CB_MESSAGES (fSYSTEM, fSTATE, fCLIENT, " _
                         & "fMSGTYPE,fBODY,fSIGN1,fSIGN2) " _
                         & " values (20"  _ 
                         & "  , 8                   " _
                         & "  , '" & cliCode & "'  " _
                         & "  , '" & msgType & "'  " _
                         & "  , char(13)+char(10)" _
                         & "      + 'DOCNUM:" & docNum & "'      + char(13)+char(10) " _    
                         & "      + 'PAYDATE:" & Trim(todayD) & "'      + char(13)+char(10) " _
                         & "      + 'PAYERACC:" & payerAcc & "'  + char(13)+char(10) " _
                         & "      + 'TAXCODE:" & taxCode & "'    + char(13)+char(10) " _  
                         & "      + 'REGNUM:" & regNum & "'      + char(13)+char(10) " _ 
                         & "      + 'AMOUNT:" & amount & "'      + char(13)+char(10) " _
                         & "      + 'SUMMA:" & amount & "'       + char(13)+char(10)  " _ 
                         & "      + 'CURR:" & curISO & "'        + char(13)+char(10) " _ 
                         & "      + 'AIM:" & aim & "'            + char(13)+char(10) " _ 
                         & "      + 'ROW:" & row1 & "'           +CHAR(13)+char(10) " _
                         & "      + 'ROW:" & row2 & "'           +CHAR(13)+char(10) " _   
                         & " , Cast('' AS VARBINARY(MAX))" _
                         & " , Cast('' AS VARBINARY(MAX)))"
                       
      Call  Execute_SLQ_Query(queryStrin)
      
      ' Մուտք հեռահար համակարգեր ԱՇՏ
      Call ChangeWorkspace(c_RemoteSyss)

      ' Պայմանագրի առկայության ստուգումը մշակման ենթակա մուտքային հաղորդագրություններ (Ընդհանուր) թղթապանակում
      msgType = ""
      system = "20"
      wState = "êïáñ³·ñáõÃÛáõÝÝ»ñÁ ×Çßï »Ý"
      direction = "|Ð»é³Ñ³ñ Ñ³Ù³Ï³ñ·»ñ|Øß³ÏÙ³Ý »ÝÃ³Ï³ Ùáõïù³ÛÇÝ Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñ(ÀÝ¹Ñ³Ýáõñ)"
      dirName = "Մշակման ենթակա մուտքային հաղորդագրություններ (Ընդհանուր)"
      status = CheckContractRemoteSystems(direction, todayDMY, system, cliCode, msgType, amount, dirName, wState)
      If Not status Then
            Log.Error("Սխալ` Մշակման ենթակա մուտքային հաղորդագրություններ մուտք գործելիս")
            Exit Sub
      End If
      
      ' Փակել թղթապանակը
      BuiltIn.Delay(1000)
      wMDIClient.VBObject("frmPttel").Close
      
      ' Հաղորդագրությունների ավտոմատ մշակում
      delayTime = 8000
      Call AutoMessageProcessing(clCount, delayTime)
      
      ' ISN- ի ստացում
      queryString = " Select fISN  from CB_MESSAGES where fDATE > '" & Trim(todayD) & "' and substring(fBODY,10,6) = '" & docNum & "' " 
      fISN = Get_Query_Result(queryString)
      Log.Message("Պայամանագրի ISN` " & fISN)
      
      ' Մուտք համակարգ VERIFIER օգտագործողով
      Login("VERIFIER")
      
      ' Մուտք հաստատվող վճարային փաստաթղթեր թղթապանակ
   '   Call wTreeView.DblClickItem("|Ð³ëï³ïáÕ I ²Þî|Ð³ëï³ïíáÕ í×³ñ³ÛÇÝ ÷³ëï³ÃÕÃ»ñ")
      Set verifyDocuments = New_VerificationDocument()
      verifyDocuments.User = "^A[Del]"
      Call GoToVerificationDocument("|Ð³ëï³ïáÕ I ²Þî|Ð³ëï³ïíáÕ í×³ñ³ÛÇÝ ÷³ëï³ÃÕÃ»ñ",verifyDocuments)
      If Not wMDIClient.WaitVBObject("frmPttel", 3000).Exists Then
            Log.Error("Հաստատվող վճարային փաստաթղթեր թղթապանակը չի բացվել")
            Exit Sub
      End If
      
      ' Վավերացվել Բազմակի փոխանցման հանձնարարական փաստաթուղթը
      colN = 1
      action = c_ToConfirm
      doNum = 1
      doActio = "Ð³ëï³ï»É"
      status = ConfirmContractDoc(colN, fISN, action, doNum, doActio)
      If Not status Then
            Log.Error("Բազմակի փոխանցման հանձնարարական փաստաթուղթը չի վավերացվել")
            Exit Sub
      End If
      
      BuiltIn.Delay(1000)
      wMDIClient.VBObject("frmPttel").Close
      
      ' Մուտք համակարգ ARMSOFT օգտագործողով
      Login("ARMSOFT")
      
      ' Մուտք BankMail ԱՇՏ
      Call ChangeWorkspace(c_BM)
        
      ' Ուղարկվող փոխանցումներ թղթապանակի բացում
      workEnvName = "|BankMail ²Þî|àõÕ³ñÏíáÕ Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñ|àõÕ³ñÏíáÕ ÷áË³ÝóáõÙÝ»ñ"
      workEnv = "Ուղարկվող փոխանցումներ"
      stRekName = "PERN"
      wStatus = False
      endRekName = "PERK"
      status = AccessFolder(workEnvName, workEnv, stRekName, todayDMY, endRekName, todayDMY, wStatus, isnRekName, fISN)
      If Not status Then
            Log.Error("Սխալ՝ հաշվառված վճարային փաստաթղթեր թղթապանակում")
            Exit Sub
      End If
      
      ' Ծնող զավակ կապի ստուգում  
      queryString = "SELECT fISN FROM DOCP WHERE fNAME = 'HT202' AND fPARENTISN = " & fISN
       
      wChildISN = Get_Query_Result(queryString)
      Log.Message("Զավակ պայամանագրի ISN` " & wChildISN)
      
      ' Փաստաթղթի առկայության ստուգում
      colN = 2
      status = CheckContractDoc(colN, wChildISN)
      If Not status Then
            Log.Error("Փաստաթուղթն առկա չէ Ուղարկվող փոխանցումներ թղթապանակում")
            Exit Sub
      End If
 
      ' Տվյալների ստուգում CB_MESSAGES աղյուսակում
      queryString = "SELECT COUNT(*) FROM CB_MESSAGES WHERE fDATE > '" & Trim(todayD) & "' AND fSTATE = '9' AND fISN = " & fISN
      BuiltIn.Delay(1000)
      rowCount = Get_Query_Result(queryString)
      Log.Message("CB_MESSAGES աղյուսակում տողերի  քանակ՝ " & rowCount)
      If rowCount <> 1 Then
          Log.Error("CB_MESSAGES աղյուսակում SQL հարցումով միայն մեկ տող պետք է գտնվի")
          Exit Sub
      End If
      
      ' Տվյալների ստուգում BM_MESSAGES աղյուսակում
      queryString = "SELECT COUNT(*) FROM BM_MESSAGES WHERE fDATE = '" & Trim(todayD) & "' AND  fISN = " & wChildISN & " AND fREFERENCE = "& fISN
      BuiltIn.Delay(1000)
      rowCount = Get_Query_Result(queryString)
      Log.Message("BM_MESSAGES աղյուսակում տողերի քանակ՝ " & rowCount)
      If rowCount <> 1 Then
          Log.Error("BM_MESSAGES աղյուսակում SQL հարցումով միայն մեկ տող պետք է գտնվի")
          Exit Sub
      End If
      
      ' SQL ստուգում HI աղյուսակում
      queryString = " SELECT  COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                              " AND fDATE = '" & Trim(todayD) & "' AND fTYPE = '01'  AND fCUR = '000' AND fCURSUM = '300.00'  AND fOP = 'TRF' " &_
                              " AND fDBCR = 'C' AND fSUID = '77' AND fSUM = '300.00' AND fBASEBRANCH = '00' AND fBASEDEPART = '1' "
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 

      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                              " AND fDATE = '" & Trim(todayD) & "' AND fTYPE = '01'  AND fCUR = '000' AND fCURSUM = '300.00'  AND fOP = 'TRF' " &_
                              " AND fDBCR = 'D' AND fSUID = '77' AND fSUM = '300.00' AND fBASEBRANCH = '00' AND fBASEDEPART = '1' "
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 
              

      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                              " AND fDATE = '" & Trim(todayD) & "' AND fTYPE = '01'  AND fCUR = '000' AND fCURSUM = '400.00'  AND fOP = 'TRF' " &_
                              " AND fDBCR = 'D' AND fSUID = '77' AND fSUM = '400.00' AND fBASEBRANCH = '00' AND fBASEDEPART = '1' "
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 

              
      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                               " AND fDATE = '" & Trim(todayD) & "' AND fTYPE = '01'  AND fCUR = '000' AND fCURSUM = '400.00'  AND fOP = 'TRF' " &_
                              " AND fDBCR = 'D' AND fSUID = '77' AND fSUM = '400.00' AND fBASEBRANCH = '00' AND fBASEDEPART = '1' "
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 
              
              
      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                               " AND fDATE = '" & Trim(todayD) & "' AND fTYPE = '01'  AND fCUR = '000' AND fCURSUM = '500.00'  AND fOP = 'FEE' " &_
                               " AND fDBCR = 'C' AND fSUID = '77' AND fSUM = '500.00' AND fBASEBRANCH = '00' AND fBASEDEPART = '1' "
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 
              
              
      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                               " AND fDATE = '" & Trim(todayD) & "' AND fTYPE = '01'  AND fCUR = '000' AND fCURSUM = '500.00'  AND fOP = 'FEE' " &_
                               " AND fDBCR = 'D' AND fSUID = '77' AND fSUM = '500.00' AND fBASEBRANCH = '00' AND fBASEDEPART = '1' "
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 
              
      ' Փակել ՀԾ_Բանկ ծրագիրը
      Call Close_AsBank()
End Sub