'USEUNIT Library_Common 
'USEUNIT RemoteService_Library 
'USEUNIT Library_CheckDB
'USEUNIT Constants
'USEUNIT Subsystems_SQL_Library
'USEUNIT BankMail_Library

'Test Case ID 165065

Sub RemoteProc_ComGrPay_Test_New()

      Dim paramName, paramValue, fISN
      Dim docNum, queryStr, queryStrin, queryString, wState, status
      Dim todayDMY, todayD, cliCode, summa, debt, accDB, name, address, Utility, location, code, jur, aim, cur, spec
      Dim workEnvName, workEnv, stRekName, endRekName, wStatus, isnRekName
      Dim system, cliMask, msgType, clCount, direction, dirName
      Dim wtdbgView, mDIClient, frmPttel, rowCount, delayTime
      Dim sqlValue, sql_isEqual, colNum, colN, docTypeName
      Dim startDate, fDate
     
      startDate = "20030101"
      fDate = "20250101"
      Call Initialize_AsBank("bank", startDate, fDate)
               
      ' Մուտք համակարգ ARMSOFT օգտագործողով
      Call Create_Connection()
      Login("ARMSOFT")
      
      ' Պարամետրերի արժեքների ճշգրտում   
      paramName = "CBDATEMAXDIFF "
      paramValue = "0"
      Call  SetParameter(paramName, paramValue)
      
      paramName = "IBCBPROCINTERVAL  "
      paramValue = "20"
      Call  SetParameter(paramName, paramValue)
      
      ' Ձևավորում ենք docNum-ը պատահական գեներացված թվի դիմացից ավելացնելով 0-ներ, այնպես որ լինի 6 նիշ
      Call Randomize()
      docNum = right(String(6, "0") + RTrim(Int(1000 * Rnd)), 6) 
      
      ' Տվյալների ներմուծում բազա  
      todayDMY = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%d/%m/%y")
      todayD = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%Y%m%d")
      cliCode =  "00000680"
      summa = "700.00"
      debt = "700.00"
      accDB = "7770000068020100"
      name = "Kaskade"
      address = "16 Â²Ô.  38 - 6 38 6"
      Utility = "WN"
      location = "01"
      code = "01-0029-38-6"
      jur = "0"
      aim = "æñÇ í³ñÓ"
      cur = "000"
      spec = "-520&0.00&0&0&0&0&172.80&31/05/07&0&0.00&"
      Log.Message("Այսօրվա ամսաթվիվ dd/mm/yy տեսքով՝ " & todayDMY)
      Log.Message("Այսօրվա ամսաթիվ՝ " & todayD)
      
      ' Տվյալ օրով տվյալների ջնջում
      queryStr = "Delete from  CB_MESSAGES where FORMAT (fDATE, 'dd/MM/yy') = '" & Trim(todayDMY) & "' "  
      Call Execute_SLQ_Query(queryStr)
      BuiltIn.Delay(2000)
      
      queryStrin = " Insert into CB_MESSAGES (fSYSTEM, fSTATE, fCLIENT, fMSGTYPE,fBODY,fSIGN1,fSIGN2) " _
                          & " values (20             " _
                          & "  , 8                   " _
                          & "  , '" & cliCode & "'  " _
                          & "  , 'IBPayCom'   " _
                          & "  , char(13)+char(10)" _
                          & "      + 'DOCNUM:" & Trim(docNum) & "'     + char(13)+char(10) " _    
                          & "      + 'PAYDATE:" & Trim(todayD) & "' + char(13)+char(10) " _
                          & "      + 'PAYERACC:" & accDB & "'     + char(13)+char(10) " _
                          & "      + 'UTILITY:" & Utility & "'    + char(13)+char(10) " _
                          & "      + 'LOCATION:" & location & "'  + char(13)+char(10) " _  
                          & "      + 'CODE:" & code & "'          + char(13)+char(10) " _  
                          & "      + 'NAME:" & name & "'          + char(13)+char(10) " _ 
                          & "      + 'ADDRESS:" & address & "'    + char(13)+char(10) " _ 
                          & "      + 'DEBT:" & debt & "'          + char(13)+char(10) " _     
                          & "      + 'JUR:" & jur & "'            + char(13)+char(10) " _
                          & "      + 'SPEC:" & spec & "'          + char(13)+char(10) " _ 
                          & "      + 'AMOUNT:" & summa & "'       + char(13)+char(10) " _
                          & "      + 'CURR:AMD'                   + char(13)+char(10) " _ 
                          & "      + 'AIM:" & aim & "'            + char(13)+char(10) " _    
                          & " , Cast('' AS VARBINARY(MAX))" _
                          & " , Cast('' AS VARBINARY(MAX)))" 
        
      ' Հաղորդագրության ներմուծում CB_MESSAGES աղյուսակում
      Call  Execute_SLQ_Query(queryStrin)
        
      ' Մուտք հեռահար համակարգեր ԱՇՏ
      Call ChangeWorkspace(c_RemoteSyss)
      
     ' Պայմանագրի առկայության ստուգումը մշակման ենթակա մուտքային հաղորդագրություններ (Ընդհանուր) թղթապանակում
      system = "20"
      cliMask = "00000680"
'      msgType = "IBPayCom"
      wState = "êïáñ³·ñáõÃÛáõÝÝ»ñÁ ×Çßï »Ý"
      direction = "|Ð»é³Ñ³ñ Ñ³Ù³Ï³ñ·»ñ|Øß³ÏÙ³Ý »ÝÃ³Ï³ Ùáõïù³ÛÇÝ Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñ(ÀÝ¹Ñ³Ýáõñ)"
      dirName = "Մշակման ենթակա մուտքային հաղորդագրություններ (Ընդհանուր)"
      status =  CheckContractRemoteSystems(direction, todayDMY, system, cliMask, msgType, summa, dirName, wState)
      
      If Not status Then
            Log.Error("Սխալ՝ Մշակման ենթակա մուտքային հաղորդագրություններ թղթապանակ մուտք գործելիս")
            Exit Sub
      End If
      
      ' Փակել թղթապանակը
      BuiltIn.Delay(1000)
      wMDIClient.VBObject("frmPttel").Close 
      
      ' Հաղորդագրությունների ավտոմատ մշակում
      delayTime = 8000
      Call AutoMessageProcessing(clCount, delayTime)
      
      ' Մուտք կոմունալ վճարումների ԱՇՏ
      Call ChangeWorkspace(c_ComPay) 
      
      ' Կոմունալ վճարումներ Թղթապանակ մուտք գործելուց դիալոգում տվյալների լրացում 
      workEnvName = "|ÎáÙáõÝ³É í×³ñáõÙÝ»ñÇ ²Þî|ÎáÙáõÝ³É í×³ñáõÙÝ»ñ"
      workEnv = "Կոմունալ վճարումներ"
      stRekName = "DSDATE"
      endRekName = "DEDATE"
      wStatus = False
      status = AccessFolder(workEnvName, workEnv, stRekName, todayDMY, endRekName, todayDMY, wStatus, isnRekName, fISN)
      If Not status Then
            Log.Error("Սխալ՝ Կոմունալ վճարումներ թղթապանակ մուտք գործելիս")
      End If
      
      ' Ստուգել Կոմունալ վճարման հանձնարարագրի առկայությունը  կոմունալ վճարումներ թղթապանակում
      colN = 10
      docTypeName = docNum & "/1"
      status = CheckContractDoc(colN, docTypeName)
      If Not status Then
            Log.Error("Կոմունալ վճարման հանձնարարագիրը առկա չէ կոմունալ վճարումներ թղթապանակում")
            Exit Sub 
      End If
      
'Function
'      ' Կատարել բոլոր գործողությունները
'      Call wMainForm.MainMenu.Click(c_AllActions)
'      ' Դիտել
'      Call wMainForm.PopupMenu.Click(c_View)
'        
'      ' Կոմունալ վճարման հանձնարարագրի ISN - ի ստացում
'      Set mDIClient = Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1)
'      fISN = mDIClient.VBObject("frmASDocForm").DocFormCommon.Doc.ISN
'      Log.Message("Փաստաթղթի ISN` " & fISN)
'      BuiltIn.Delay(1000)  
'       
'      ' OK կոճակի սեղմում
'      Call ClickCmdButton(1, "OK")
'End Function

      fISN = GetISN()
      If fISN = "" Then
            Log.Error("Կոմունալ վճարման հանձնարարագրի ISN - ի ստացման ձախողում")
            Exit Sub
      End If
      
      ' Անհրաժեշտ գրառման առկայության ստուգում COM_PAYMENTS աղյուսակում
      queryString = " Select COUNT(*) From COM_PAYMENTS Where  fISN = " & fISN
      rowCount = Get_Query_Result(queryString)
      If rowCount <>1 Then
            Log.Error("Չի կարող նույն ISN -ով 1ից ավել վճարման հանձնարարագիր լինել")
            Exit Sub
      End If 
      
      ' SQL ստուգում 
      queryString = " SELECT fSUM FROM HI WHERE fBASE= " & fISN & _
                              " AND fTYPE = '01' AND fOBJECT = '385340295' AND fCUR = '000'   AND fCURSUM = '700.00' AND fOP = 'MSC' " &_
                              " AND fDBCR = 'C' AND fSUID = '77' AND fTRANS = '0' AND fBASEBRANCH = '00' AND fBASEDEPART = '1' "
      sqlValue = 700.00
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 

      queryString = " SELECT fSUM FROM HI WHERE fBASE= " & fISN & _
                               "  AND fTYPE = '01' AND fOBJECT = '431808082' AND fCUR = '000'   AND fCURSUM = '700.00' AND fOP = 'MSC' " &_
                               " AND fDBCR = 'D' AND fSUID = '77' AND fTRANS = '0' AND fBASEBRANCH = '00' AND fBASEDEPART = '1' "
      sqlValue = 700.00
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 
             
      ' Տվյալների ստուգում CB_MESSAGES աղյուսակում
      queryString = " Select COUNT(*) From CB_MESSAGES Where FORMAT (fDATE, 'dd/MM/yy') >= '"& Trim(todayDMY) & _
                              "' And  fSTATE = '9' And fCLIENT ="& cliCode & " And fSYSTEM = " & system & " And fISN = " & fISN
      rowCount = Get_Query_Result(queryString)
      Log.Message("CB_MESSAGES աղյուսակում գտնված տողերի քանակ` " & rowCount)
      If Trim(rowCount) <> 1 Then
            Log.Error("CB_MESSAGES աղյուսակում SQL հարցումով միայն մեկ տող պետք է գտնվի")
      End If
      
      BuiltIn.Delay(1000)
      wMDIClient.VBObject("frmPttel").Close
      
      ' Փակել ծրագիրը
      Call Close_AsBank()
End Sub