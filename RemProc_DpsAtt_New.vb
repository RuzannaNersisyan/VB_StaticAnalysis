'USEUNIT Library_Common 
'USEUNIT RemoteService_Library 
'USEUNIT Library_CheckDB
'USEUNIT Constants
'USEUNIT Subsystems_SQL_Library
'USEUNIT Payment_Order_ConfirmPhases_Library 
'USEUNIT BankMail_Library
'USEUNIT Deposit_Contract_Library
'USEUNIT Clients_Library

'Test Case ID 165381

Sub RemProc_DpsAtt_New_Test()
      
      Dim docNum, yesterday, calcInt
      Dim frmPttel, fAGRISN, fISN, clCount, rowCount
      Dim paramName, paramValue, queryStrin, queryStr, queryString, rowID
      Dim todayD, todayDMY, cliCode, dpsAttIsn, amount, curISO, aim, account, op, fData
      Dim status, direction, system, msgType, dirName, wState, delayTime
      Dim cap, ext, rep, per, close, folderDirect, folderName, rekvName
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
      
      ' Մուտք ավանդներ ներգրավված ԱՇՏ
      Call ChangeWorkspace(c_Deposits)

      ' Պայմանագրեր թղթապանակի բացում
      folderDirect = "|²í³Ý¹Ý»ñ (Ý»ñ·ñ³íí³Í)|ä³ÛÙ³Ý³·ñ»ñ"
      folderName = "Պայմանագրեր/Ավանդներ ներգրավված "
      rekvName = "NUM"
      docNum = "A-001149"
      Call OpenFolder(folderDirect, folderName, rekvName, docNum)
      
      ' Ավանդի համար խմբային հաշվարկ գործողության կատարում
      calcInt = 1
      yesterday = aqConvert.DateTimeToFormatStr(aqDateTime.Today-1,"%d/%m/%y")
      Call Group_Calculate(yesterday, yesterday, cap, ext, rep, per, calcInt, close)

      BuiltIn.Delay(2000)
      wMDIClient.VBObject("frmPttel").Close
      
      todayD = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%Y%m%d")
      todayDMY = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%d/%m/%y")
      cliCode =  "00000004"
      dpsAttIsn = 1238807189         
      amount = "500.00"
      curISO = "AMD"
      aim = "Avandi hamalrum"
      account = "00000450500"
      op = 4
      fData = "<Data><InternetBankData><documentID>743c63fb-e31f-486f-a050-4ad7e12657f7</documentID></InternetBankData><SumAMD>500.00</SumAMD></Data>"
         
      ' Տվյալ օրով տվյալների ջնջում
      queryStr = "Delete from  IB_AGREEMENT_PAYMENT where FORMAT (fDATE, 'dd/MM/yy') = '" & Trim(todayDMY) & "' "  
      Call Execute_SLQ_Query(queryStr)
      BuiltIn.Delay(2000)  
      
      queryStrin = " Insert into IB_AGREEMENT_PAYMENT ( " _
                                & "fSYSTEM, fDATE, fSTATE, fCLIENT, " _
                                & "fPASSPORT, fAGRISN, fTYPE, fPAYDATE, " _
                                & "fCURISO, fAMOUNT, " _
                                & "fACCOUNT,fISN, " _
                                & "fOP,fCOMMENT,fDATA) " _
                                & " values (20              " _
                                & "  , '" & Trim(todayD) & "'      " _     
                                & "  , 8                    " _
                                & "  , '" & cliCode & "'    " _
                                & "  , ''                   " _
                                & "  , " & dpsAttIsn          _
                                & "  , 1                    " _  
                                & "  , '" & Trim(todayD) & "'      " _ 
                                & "  , '" & curISO & "'     " _   
                                & "  , '" & amount & "'     " _ 
                                & "  , '" & account & "'    " _ 
                                & "  , -1                   " _
                                & "  , " & op                 _ 
                                & "  , '" & aim & "'        " _
                                & "  , '" & fData & "')     "  
      
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
            Log.Error("Սխալ` Մշակման ենթակա մուտքային հաղորդագրություններ թղթապանակ մուտք գործելիս")
            Exit Sub
      End If
      
      BuiltIn.Delay(2000)
      wMDIClient.VBObject("frmPttel").Close
      
      ' IB_AGREEMENT_PAYMENT աղյուսակից ստանում ենք  fROWID -ն
      queryString = " Select fROWID from IB_AGREEMENT_PAYMENT where fDATE > =  '" & Trim(todayD) & "' "
      rowID = Get_Query_Result(queryString)
      Log.Message("fROWID՝  " & rowID)
      
      ' Հաղորդագրությունների ավտոմատ մշակում
      delayTime = 8000
      Call AutoMessageProcessing(clCount, delayTime)
      
      ' Տվյալների ստուգում IB_AGREEMENT_PAYMENT աղյուսակում
      queryString = "SELECT COUNT(*) FROM IB_AGREEMENT_PAYMENT where fROWID = '"& rowID &"' AND fSTATE = '9' "
      BuiltIn.Delay(1000)
      rowCount = Get_Query_Result(queryString)
      Log.Message("Տողերի քանակը IB_AGREEMENT_PAYMENT աղյուսակում ՝ " & rowCount)
      If rowCount <> 1 Then
          Log.Error("IB_AGREEMENT_PAYMENT աղյուսակում տվյալ SQL հարցումով միայն մեկ տող պետք է գտնվի")
          Exit Sub
      End If
      
      ' Տվյալների ստուգում IB_AGREEMENT_PAYMENT աղյուսակում
      queryString = "Select fISN from IB_AGREEMENT_PAYMENT  where fDATE > =  '" & Trim(todayD) & "' and  fSTATE = 9 and fROWID = " & rowID
      BuiltIn.Delay(1000)
      fISN = Get_Query_Result(queryString)
      Log.Message("Պայմանագրի ISN` " & fISN)

      ' Տվյալների ստուգում IB_AGREEMENT_PAYMENT աղյուսակում
      queryString = " Select fAGRISN from IB_AGREEMENT_PAYMENT  where fDATE > =  '" & Trim(todayD) & "' and  fSTATE = 9 and fROWID = " & rowID
      BuiltIn.Delay(1000)
      fAGRISN = Get_Query_Result(queryString)
      Log.Message("fAGRISN` " & fAGRISN)
      
      ' SQL ստուգում HIR աղյուսակում
      queryString = " Select COUNT(*) from HIR where  fBASE = "& fISN &_
                               " and fCURSUM = '500.00' and fTYPE = 'R1' and fCUR = '000' and fOP = 'AGR' and fDBCR = 'D'  " &_
                               " and fSUID = '77' and fBASEBRANCH = '00' and fBASEDEPART = '1'	"
      
      rowCount = Get_Query_Result(queryString)
      Log.Message("HIR աղյուսակում գրանցումների քանակ՝ " & rowCount)
      If rowCount <> 1 Then
          Log.Error("HIR աղյուսակում տվյալ SQL հարցումով միայն մեկ տող պետք է գտնվի")
          Exit Sub
      End If
             
      ' Փակել ՀԾ_Բանկ ծրագիրը
      Call Close_AsBank()
End Sub