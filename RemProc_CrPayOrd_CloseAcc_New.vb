'USEUNIT Library_Common 
'USEUNIT RemoteService_Library 
'USEUNIT Library_CheckDB
'USEUNIT Constants
'USEUNIT Subsystems_SQL_Library
'USEUNIT Payment_Order_ConfirmPhases_Library 
'USEUNIT BankMail_Library

'Test Case ID 165374

Sub RemProc_CrPayOrd_CloseAcc_New_Test()

      Dim paramName, paramValue, queryStr, queryStrin, queryString, acc, action
      Dim  todayD, todayDMY, cliCode, payerAcc, receiverAcc, receiver, areaCode, dbRes, dbJurStat, dbRegNum
      Dim debtor, dbPass, dbPassType, dbAddress, dbMng, dbInfo, reportCode, amount, curISO, cur, aim
      Dim direction, system, msgType, dirName, wState, wStatus, frmPttel, wComment, rowCount, clCount
      Dim startDate, fDate, delayTime, colNum
     
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
      
      ' Փակել հաշիվը
      acc = "33120090800"
      action = 0
      Call AccCloseOrOpen(acc, action)
      
      BuiltIn.Delay(1000)
      wMDIClient.VBObject("frmPttel").Close
      
      todayD = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%Y%m%d")
      todayDMY = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%d/%m/%y")
      cliCode =  "00001005"
      payerAcc = "7770033120090800"
      receiverAcc = "900000000100" 
      receiver = "Tranzitayin hashiv dramayin poxancumneri hamar" 
      areaCode = "99"
      dbRes = "1"
      dbJurStat = "21"
      dbRegNum = "1805901574"
      debtor = "Mnacakanyan Armen"
      dbPass = "AH0102654"
      dbPassType = "2"
      dbAddress = "Address" 
      dbMng = "Karen Vardanyan"
      dbInfo = "More info"
      reportCode = "PTD/OTM000000E000000OT/0/0  MESSAGE1:PND/01200154564           S"           
      amount = "900.00"
      curISO = "AMD"
      cur = "000"
      aim = "test"
      
      ' Ձևավորում ենք docNum-ը պատահական գեներացված թվի դիմացից ավելացնելով 0-ներ, այնպես որ լինի 6 նիշ
      Call Randomize()
      docNum = right(String(6, "0") + RTrim(Int(1000 * Rnd)), 6) 
      Log.Message("Փաստաթղթի համարը՝ " & docNum)
      
      ' Տվյալ օրով տվյալների ջնջում
      queryStr = "Delete from  CB_MESSAGES where FORMAT (fDATE, 'dd/MM/yy') = '" & Trim(todayDMY) & "' "  
      Call Execute_SLQ_Query(queryStr)
      BuiltIn.Delay(2000)
      
      queryStrin =  " Insert into CB_MESSAGES (fSYSTEM, fSTATE, fCLIENT, fMSGTYPE, " _
                                  & "fBODY,fSIGN1,fSIGN2) " _
                                  & " values (20             " _
                                  & "  , 8                   " _
                                  & "  , '" & cliCode & "'  " _
                                  & "  , 'IBPayOrd'   " _
                                  & "  , char(13)+char(10)" _
                                  & "      + 'DOCNUM:" & Trim(docNum) & "'              + char(13)+char(10) " _    
                                  & "      + 'PAYDATE:" & Trim(todayD) & "'              + char(13)+char(10) " _
                                  & "      + 'PAYERACC:" & payerAcc & "'          + char(13)+char(10) " _
                                  & "      + 'RECEIVERACC:" & receiverAcc & "'    + char(13)+char(10) " _  
                                  & "      + 'RECEIVER:" & receiver & "'          + char(13)+char(10) " _ 
                                  & "      + 'AREACODE:" & areaCode & "'          + CHAR(13)+char(10) " _
                                  & "      + 'DBRES:" & dbRes & "'                + CHAR(13)+char(10)" _
                                  & "      + 'DBJURSTAT:" & dbJurStat & "'        + CHAR(13)+char(10)" _
                                  & "      + 'DBREGNUM:" & dbRegNum & "'          + CHAR(13)+char(10)" _
                                  & "      + 'DEBTOR:" & debtor & "'              + CHAR(13)+char(10) " _ 
                                  & "      + 'DBPASS:" & dbPass & "'              + CHAR(13)+char(10) " _
                                  & "      + 'DBPASSTYPE:" & dbPassType & "'      + CHAR(13)+char(10) " _
                                  & "      + 'DBADDRESS:" & dbAddress & "'        + CHAR(13)+char(10) " _    
                                  & "      + 'DBMNG:" & dbMng & "'                + CHAR(13)+char(10) " _ 
                                  & "      + 'DBINFO:" & dbInfo & "'              + CHAR(13)+char(10) " _
                                  & "      + 'REPORTCODE:" & reportCode & "'      + char(13)+char(10) " _
                                  & "      + 'AMOUNT:" & amount & "'              + char(13)+char(10) " _
                                  & "      + 'CURR:" & curISO & "'                + char(13)+char(10) " _ 
                                  & "      + 'AIM:" & aim & "'                    + char(13)+char(10) " _    
                                  & " , Cast('' AS VARBINARY(MAX))" _
                                  & " , Cast('' AS VARBINARY(MAX)))" 
        
      Call  Execute_SLQ_Query(queryStrin)
      
      ' Մուտք հեռահար համակարգեր ԱՇՏ
      Call ChangeWorkspace(c_RemoteSyss)

      ' Պայմանագրի առկայության ստուգումը մշակման ենթակա մուտքային հաղորդագրություններ (Ընդհանուր) թղթապանակում
      system = "20"
'      msgType = "IBPayOrd"
      wState = "êïáñ³·ñáõÃÛáõÝÝ»ñÁ ×Çßï »Ý"
      direction = "|Ð»é³Ñ³ñ Ñ³Ù³Ï³ñ·»ñ|Øß³ÏÙ³Ý »ÝÃ³Ï³ Ùáõïù³ÛÇÝ Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñ(ÀÝ¹Ñ³Ýáõñ)"
      dirName = "Մշակման ենթակա մուտքային հաղորդագրություններ (Ընդհանուր)"
      wStatus = CheckContractRemoteSystems(direction, todayDMY, system, cliCode, msgType, amount, dirName, wState)
      If Not wStatus Then
            Log.Error("Սխալ՝ Մշակման ենթակա մուտքային հաղորդագրություններ մուտք գործելիս")
            Exit Sub
      End If
      
      ' Փակել թղթապանակը
      BuiltIn.Delay(1000)
      wMDIClient.VBObject("frmPttel").Close
      
      ' Հաղորդագրությունների ավտոմատ մշակում
      delayTime = 50000
      Call AutoMessageProcessing(clCount, delayTime)
      
      ' Մուտք Հեռահար համակարգեր ԱՇՏ
      Call ChangeWorkspace(c_RemoteSyss)
      
      ' Մուտք մուտքային հաղորդագրությունների դիտում (ընդհանուր) թղթապանակ
      wState = "Ø»ñÅí³Í ¿ µ³ÝÏÇ ÏáÕÙÇó"
      direction = ("|Ð»é³Ñ³ñ Ñ³Ù³Ï³ñ·»ñ|ÂÕÃ³å³Ý³ÏÝ»ñ|Øáõïù³ÛÇÝ Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñÇ ¹ÇïáõÙ|Øáõïù³ÛÇÝ Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñÇ ¹ÇïáõÙ(ÀÝ¹Ñ³Ýáõñ)")
      dirName = "Մուտքային հաղորդագրությունների դիտում (ընդհանուր)"
      wStatus = CheckContractRemoteSystems(direction, todayDMY, system, cliCode, msgType, amount, dirName, wState)
      If Not wStatus Then
            Log.Error("Սխալ՝ մուտքային հաղորդագրությունների դիտում թղթապանակ մուտք գործելիս")
            Exit Sub
      End If
      
      colNum = wMDIClient.VBObject("frmPttel").GetColumnIndex("fCOMMENT")
      wComment = Trim(wMDIClient.VBObject("frmPttel").VBObject("tdbgView").Columns.Item(colNum).Text)
      Log.Message("Մեկնաբանություն դաշտի արժեք՝ " & wComment)
      
      ' Փակել թղթապանակը
      BuiltIn.Delay(1000)
      wMDIClient.VBObject("frmPttel").Close
      
      ' Տվյալների ստուգում CB_MESSAGES աղյուսակում 
      queryString = " Select COUNT(*) From CB_MESSAGES Where fDATE > ' "& Trim(todayD) & "' And fSTATE = '10' And fCOMMENT = '"& wComment &"' "
      BuiltIn.Delay(1000)
      rowCount = Get_Query_Result(queryString)
      Log.Message("CB_MESSAGES աղյուսակում տողերի քանակ՝ " & rowCount)
      
      If rowCount <> 1 Then
            Log.Error("CB_MESSAGES աղյուսակում հարցման արդյունքում միայն մեկ տող պետք է գտնվի")
            Exit Sub
      End If
      
      ' Բացել հաշիվը
      acc = "33120090800"
      action = 1
      Call AccCloseOrOpen(acc, action)
      
      ' Փակել ծրագիրը
      Call Close_AsBank()   
End Sub