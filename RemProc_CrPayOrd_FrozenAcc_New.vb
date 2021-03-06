'USEUNIT Library_Common 
'USEUNIT RemoteService_Library 
'USEUNIT Library_CheckDB
'USEUNIT Constants
'USEUNIT Subsystems_SQL_Library
'USEUNIT Payment_Order_ConfirmPhases_Library 
'USEUNIT BankMail_Library

'Test Case ID 165371

Sub RemProc_CrPayOrd_FrozenAcc_New_Test()

      Dim paramName, paramValue, queryStrin, queryStr, queryString
      Dim accMask, frozen, direction, dirName
      Dim docNum, todayD, todayDMY, cliCode, payerAcc, receiverAcc, receiver, regNum, amount, curISO, cur, aim
      Dim fRow, rowCount, frmPttel, delayTime
      Dim startDate, fDate
     
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
      
      ' Մուտք հեռահար համակարգեր ԱՇՏ
      Call ChangeWorkspace(c_ChiefAcc)
      
      ' Սառեցնել հաշիվը
      accMask = "03485010100"
      frozen = "3"
      aim = "test"
      Call FreezeAccOrNo(accMask, frozen, aim)
      
      ' Ձևավորում ենք docNum-ը պատահական գեներացված թվի դիմացից ավելացնելով 0-ներ, այնպես որ լինի 6 նիշ
      Call Randomize()
      docNum = right(String(6, "0") + RTrim(Int(1000 * Rnd)), 6) 
      Log.Message("Փաստաթղթի համար՝ " & docNum)
      
      todayD = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%Y%m%d")
      todayDMY = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%d/%m/%y")
      cliCode =  "00034850"
      payerAcc = "7770003485010100"
      receiverAcc = "7770038210090300" 
      receiver = "Partavorutyunner petakan gandzaranin" 
      regNum = "8832132166"                    
      amount = "850.00"
      curISO = "AMD"
      cur = "000"
      aim = "test"
    
      ' Տվյալ օրով տվյալների ջնջում
      queryStr = "Delete from  CB_MESSAGES where FORMAT (fDATE, 'dd/MM/yy') = '" & Trim(todayDMY) & "' "  
      Call Execute_SLQ_Query(queryStr)
      BuiltIn.Delay(2000)
      
      queryStrin =  " Insert into CB_MESSAGES (fSYSTEM, fSTATE, fCLIENT, fMSGTYPE, " _
                          & " fBODY,fSIGN1,fSIGN2) " _
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
                          & "      + 'REGNUM:" & regNum & "'              + char(13)+char(10) " _
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
      BuiltIn.Delay(1000)
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
      wState = "ÊÝ¹ñ³Ñ³ñáõÛó"
      direction = ("|Ð»é³Ñ³ñ Ñ³Ù³Ï³ñ·»ñ|ÂÕÃ³å³Ý³ÏÝ»ñ|Øáõïù³ÛÇÝ Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñÇ ¹ÇïáõÙ|Øáõïù³ÛÇÝ Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñÇ ¹ÇïáõÙ(ÀÝ¹Ñ³Ýáõñ)")
      dirName = "Մուտքային հաղորդագրությունների դիտում (ընդհանուր)"
      wStatus = CheckContractRemoteSystems(direction, todayDMY, system, cliCode, msgType, amount, dirName, wState)
      BuiltIn.Delay(1000)
      If Not wStatus Then
            Log.Error("Սխալ՝ մուտքային հաղորդագրությունների դիտում թղթապանակ մուտք գործելիս")
            Exit Sub
      End If
      
      ' Փակել թղթապանակը
      BuiltIn.Delay(1000)
      wMDIClient.VBObject("frmPttel").Close
      
      ' fROW -  ի արժեքի ստացում
      queryString = " Select fROW From CB_MESSAGES Where fDATE > '" & Trim(todayD) & "' And fSTATE = '19' "
      BuiltIn.Delay(1000)
      fRow = Get_Query_Result(queryString)
      Log.Message("fROW սյան արժեք՝  "& fRow)
      
      ' Տվյալների ստուգում CB_MESSAGES  աղյուսակում
      queryString = " Select COUNT(*) From CB_MESSAGES Where fDATE > '" & Trim(todayD) & "' And fSTATE = '19' "
      BuiltIn.Delay(1000)
      rowCount = Get_Query_Result(queryString)
      Log.Message("CB_MESSAGES  աղյուսակում տողերի քանակ՝ " & rowCount)
      If rowCount <> 1 Then
          Log.Error("CB_MESSAGES աղյուսակում այսօրվա հարցումով միայն 1 տող պետք է գտնվի")
          Exit Sub
      End If
     
     ' Տվյալների ստուգում CB_PROCERRORS աղյուսակում
      queryString = " Select COUNT(*) From CB_PROCERRORS Where fROW = '"& fRow &"' And fPROCCOUNT = '3' And fERRORCODE = '2' "
      BuiltIn.Delay(1000)
      rowCount = Get_Query_Result(queryString)
      Log.Message("CB_PROCERRORS աղյուսակում տողերի քանակ՝ " & rowCount)
      
      If rowCount <> 1 Then
          Log.Error("CB_PROCERRORS աղյուսակում այսօրվա հարցումով միայն 1 տող պետք է գտնվի")
          Exit Sub
      End If
      
      ' Մուտք հեռահար համակարգեր ԱՇՏ
      Call ChangeWorkspace(c_ChiefAcc)
      
      ' Սառեցնել հաշիվը
      accMask = "03485010100"
      frozen = "0"
      wAim = "Test"
      Call FreezeAccOrNo(accMask, frozen, wAim)
      
      ' Փակել ծրագիրը
      Call Close_AsBank()   
End Sub