'USEUNIT Library_Common 
'USEUNIT RemoteService_Library 
'USEUNIT Library_CheckDB 
'USEUNIT Constants
'USEUNIT Subsystems_SQL_Library
'USEUNIT Payment_Order_ConfirmPhases_Library 
'USEUNIT BankMail_Library
'USEUNIT Library_Contracts
'Test Case 85573

Sub RemProc_CrPayOrd_SendBack_From_PartEditing_20_Test()

      Dim cliCode, payerAcc, receiverAcc, receiver, areaCode, dbRes, dbJurStat, dbRegNum, debtor 
      Dim dbPass, dbPassType, dbAddress, dbMng, dbInfo, reportCode, amount, curISO, cur, aim
      Dim direction, todayDMY, system, msgType, dirName, wState, wStatus, status, colN, fISN
      Dim State, workEnvName, workEnv, stRekName, endRekName, isnRekName, todayD
      Dim  docNum, action, doNum, doActio, basis, refuse, memOrdfISN, ordDocNum
      Dim paramName, paramValue, confPath, confInput, queryStrin, queryStr
      Dim directFolder, folderName, wUser, wDocType, frmPttel, delayTime
      Dim queryString, rowCount, sqlValue, colNum, sql_isEqual
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
      confPath = "X:\Testing\RemAutoProc\Verify2_CrpayOrd_New.txt"
      confInput = Input_Config(confPath)
      If Not confInput Then
          Log.Error("Կարգավորումները չեն ներմուծվել")
          Exit Sub
      End If
    
      todayD = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%Y%m%d")
      todayDMY = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%d/%m/%y")
      cliCode =  "00000680"
      payerAcc = "7770000068020100"
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
      amount = "250.00"
      curISO = "AMD"
      cur = "000"
      aim = "test"
      msgType = "IBPayOrd"
      system = "20"
      
      ' Ձևավորում ենք docNum-ը պատահական գեներացված թվի դիմացից ավելացնելով 0-ներ, այնպես որ լինի 6 նիշ
      Call Randomize()
      docNum = right(String(6, "0") + RTrim(Int(1000 * Rnd)), 6) 
      Log.Message("Փաստաթղթի համար՝ " & docNum)
      
      ' Տվյալ օրով տվյալների ջնջում
      queryStr = "Delete from  CB_MESSAGES where FORMAT (fDATE, 'dd/MM/yy') = '" & Trim(todayDMY) & "' "  
      Call Execute_SLQ_Query(queryStr)
      BuiltIn.Delay(2000)
      
      queryStrin =  " Insert into CB_MESSAGES (fSYSTEM, fSTATE, fCLIENT, fMSGTYPE,fBODY,fSIGN1,fSIGN2) " _
                          & " values (" & system _ 
                          & "  , 8                   " _
                          & "  , '" & cliCode & "'  " _
                          & "  , '" & msgType & "'  " _
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
                          & "      + 'SUMMA:" & amount & "'              + char(13)+char(10)  " _ 
                          & "      + 'CURR:" & curISO & "'                + char(13)+char(10) " _ 
                          & "      + 'AIM:" & aim & "'                    + char(13)+char(10) " _    
                          & " , Cast('0x308206EA06092A864886F70D010702A08206DB308206D7020101310B300906052B0E03021A0500300B06092A864886F70D010701A08205AD308205A930820491A003020102020A12F7787F00000000000D300D06092A864886F70D0101050500305131123010060A0992268993F22C6401191602616D31173015060A0992268993F22C640119160761726D736F6674312230200603550403131941726D656E69616E20536F66747761726520526F6F74204341301E170D3034313030383035313634305A170D3035313030383035323634305A303531173015060355040A130E37373730302D3030303030323333311A301806035504031311416E61686974205368617368696B79616E30819F300D06092A864886F70D010101050003818D0030818902818100E41B109B1E9A7F5582AD3631831CC6E9EDB68408598439E53245D815198B5AF472CCC5D8F3FFA2413FAE18FF159B75A7415C5D98B7FC603BD0BAB2E4759A4F5D5CCD410893A92274939C789DC31E5D4B7C3B7FD962124AAAC92A06463F93E547DE89CDE85345054EF66DA2E203A1A36C4F9FE82C190CCC8E1E453B9DB79EEEE30203010001A38203213082031D300E0603551D0F0101FF0404030204F0304406092A864886F70D01090F04373035300E06082A864886F70D030202020080300E06082A864886F70D030402020080300706052B0E030207300A06082A864886F70D030730130603551D25040C300A06082B06010505070302301D0603551D0E0416041441998D5321E77A40E297EB511F5766692E856532301F0603551D230418301680149B946FBBC5063F443D23CFCAF2A313113D10C32E3082012D0603551D1F04820124308201203082011CA0820118A08201148681C66C6461703A2F2F2F434E3D41726D656E69616E253230536F667477617265253230526F6F7425323043412C434E3D7465726D696E616C2C434E3D4344502C434E3D5075626C69632532304B657925323053657276696365732C434E3D53657276696365732C434E3D436F6E66696775726174696F6E2C44433D61726D736F66742C44433D616D3F63657274696669636174655265766F636174696F6E4C6973743F626173653F6F626A656374436C6173733D63524C446973747269627574696F6E506F696E748649687474703A2F2F7465726D696E616C2E61726D736F66742E616D2F43657274456E726F6C6C2F41726D656E69616E253230536F667477617265253230526F6F7425323043412E63726C3082013D06082B060105050701010482012F3082012B3081BD06082B060105050730028681B06C6461703A2F2F2F434E3D41726D656E69616E253230536F667477617265253230526F6F7425323043412C434E3D4149412C434E3D5075626C69632532304B657925323053657276696365732C434E3D53657276696365732C434E3D436F6E66696775726174696F6E2C44433D61726D736F66742C44433D616D3F634143657274696669636174653F626173653F6F626A656374436C6173733D63657274696669636174696F6E417574686F72697479306906082B06010505073002865D687474703A2F2F7465726D696E616C2E61726D736F66742E616D2F43657274456E726F6C6C2F7465726D696E616C2E61726D736F66742E616D5F41726D656E69616E253230536F667477617265253230526F6F7425323043412E637274300D06092A864886F70D0101050500038201010056948359D9E1BB72F164B0159F8D89CB3AB3BA26E739F3F4AEAADCCE6DCF4FC8373ED5BC1C945686D7E7639ADF3FA0C81E3FDE71888D1F42235BA8F18DBAA73CDA0E140DD1A4B5C1366E7B44E32392A68B0BFCBBE08AF8958F66871171BFFCBE8947B0633CF09CEB4EBC94D59A0DB05F36063C6C0ADA541068BF5F30C71693B2BD0082ADD8211172E5AF9C40C12669D6ABD56EA8869D442861D52FA68EC619CDA3F63F97955906496D77FF0D7FEC264D738D660BE9DE7A827D0BE754B85AA9ECB092E0BFD498BD19E8872B6012264F4EBF9B88FFBBB812E50EBB9B03A376D325C8152D15BDBCB638AB5FF191B01D8BCFBB1884D8D3079D64E67991207C72B1563182010530820101020101305F305131123010060A0992268993F22C6401191602616D31173015060A0992268993F22C640119160761726D736F6674312230200603550403131941726D656E69616E20536F66747761726520526F6F74204341020A12F7787F00000000000D300906052B0E03021A0500300D06092A864886F70D010101050004818042A0B20247725B8580C78FCEA1412900999AF1473146B92F93E7CB917194D14744888222B3D732471EC430BF8B301C094D6E15E6C2841072ECA56169217F296C877826CE4EFE1E23C40D2C74CC9791255104743CAC2298CE174ABBCAE48619FB04F36FED9539A015663D3B90660660DC543167EA31FB421B20AB8FA4EAC75CD7' AS VARBINARY(MAX))" _
                          & " , Cast('0x308206EA06092A864886F70D010702A08206DB308206D7020101310B300906052B0E03021A0500300B06092A864886F70D010701A08205AD308205A930820491A003020102020A12F7787F00000000000D300D06092A864886F70D0101050500305131123010060A0992268993F22C6401191602616D31173015060A0992268993F22C640119160761726D736F6674312230200603550403131941726D656E69616E20536F66747761726520526F6F74204341301E170D3034313030383035313634305A170D3035313030383035323634305A303531173015060355040A130E37373730302D3030303030323333311A301806035504031311416E61686974205368617368696B79616E30819F300D06092A864886F70D010101050003818D0030818902818100E41B109B1E9A7F5582AD3631831CC6E9EDB68408598439E53245D815198B5AF472CCC5D8F3FFA2413FAE18FF159B75A7415C5D98B7FC603BD0BAB2E4759A4F5D5CCD410893A92274939C789DC31E5D4B7C3B7FD962124AAAC92A06463F93E547DE89CDE85345054EF66DA2E203A1A36C4F9FE82C190CCC8E1E453B9DB79EEEE30203010001A38203213082031D300E0603551D0F0101FF0404030204F0304406092A864886F70D01090F04373035300E06082A864886F70D030202020080300E06082A864886F70D030402020080300706052B0E030207300A06082A864886F70D030730130603551D25040C300A06082B06010505070302301D0603551D0E0416041441998D5321E77A40E297EB511F5766692E856532301F0603551D230418301680149B946FBBC5063F443D23CFCAF2A313113D10C32E3082012D0603551D1F04820124308201203082011CA0820118A08201148681C66C6461703A2F2F2F434E3D41726D656E69616E253230536F667477617265253230526F6F7425323043412C434E3D7465726D696E616C2C434E3D4344502C434E3D5075626C69632532304B657925323053657276696365732C434E3D53657276696365732C434E3D436F6E66696775726174696F6E2C44433D61726D736F66742C44433D616D3F63657274696669636174655265766F636174696F6E4C6973743F626173653F6F626A656374436C6173733D63524C446973747269627574696F6E506F696E748649687474703A2F2F7465726D696E616C2E61726D736F66742E616D2F43657274456E726F6C6C2F41726D656E69616E253230536F667477617265253230526F6F7425323043412E63726C3082013D06082B060105050701010482012F3082012B3081BD06082B060105050730028681B06C6461703A2F2F2F434E3D41726D656E69616E253230536F667477617265253230526F6F7425323043412C434E3D4149412C434E3D5075626C69632532304B657925323053657276696365732C434E3D53657276696365732C434E3D436F6E66696775726174696F6E2C44433D61726D736F66742C44433D616D3F634143657274696669636174653F626173653F6F626A656374436C6173733D63657274696669636174696F6E417574686F72697479306906082B06010505073002865D687474703A2F2F7465726D696E616C2E61726D736F66742E616D2F43657274456E726F6C6C2F7465726D696E616C2E61726D736F66742E616D5F41726D656E69616E253230536F667477617265253230526F6F7425323043412E637274300D06092A864886F70D0101050500038201010056948359D9E1BB72F164B0159F8D89CB3AB3BA26E739F3F4AEAADCCE6DCF4FC8373ED5BC1C945686D7E7639ADF3FA0C81E3FDE71888D1F42235BA8F18DBAA73CDA0E140DD1A4B5C1366E7B44E32392A68B0BFCBBE08AF8958F66871171BFFCBE8947B0633CF09CEB4EBC94D59A0DB05F36063C6C0ADA541068BF5F30C71693B2BD0082ADD8211172E5AF9C40C12669D6ABD56EA8869D442861D52FA68EC619CDA3F63F97955906496D77FF0D7FEC264D738D660BE9DE7A827D0BE754B85AA9ECB092E0BFD498BD19E8872B6012264F4EBF9B88FFBBB812E50EBB9B03A376D325C8152D15BDBCB638AB5FF191B01D8BCFBB1884D8D3079D64E67991207C72B1563182010530820101020101305F305131123010060A0992268993F22C6401191602616D31173015060A0992268993F22C640119160761726D736F6674312230200603550403131941726D656E69616E20536F66747761726520526F6F74204341020A12F7787F00000000000D300906052B0E03021A0500300D06092A864886F70D010101050004818042A0B20247725B8580C78FCEA1412900999AF1473146B92F93E7CB917194D14744888222B3D732471EC430BF8B301C094D6E15E6C2841072ECA56169217F296C877826CE4EFE1E23C40D2C74CC9791255104743CAC2298CE174ABBCAE48619FB04F36FED9539A015663D3B90660660DC543167EA31FB421B20AB8FA4EAC75CD7' AS VARBINARY(MAX)))"
        
      Call  Execute_SLQ_Query(queryStrin)
      
      ' Մուտք հեռահար համակարգեր ԱՇՏ
      Call ChangeWorkspace(c_RemoteSyss)

      ' Պայմանագրի առկայության ստուգումը մշակման ենթակա մուտքային հաղորդագրություններ (Ընդհանուր) թղթապանակում
      msgType = ""
      wState = "êïáñ³·ñáõÃÛáõÝÝ»ñÁ ×Çßï »Ý"
      direction = "|Ð»é³Ñ³ñ Ñ³Ù³Ï³ñ·»ñ|Øß³ÏÙ³Ý »ÝÃ³Ï³ Ùáõïù³ÛÇÝ Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñ(ÀÝ¹Ñ³Ýáõñ)"
      dirName = "Մշակման ենթակա մուտքային հաղորդագրություններ (Ընդհանուր)"
      wStatus = CheckContractRemoteSystems(direction, todayDMY, system, cliCode, msgType, amount, dirName, wState)
      If Not wStatus Then
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
      queryString = " Select fISN  from CB_MESSAGES where fDATE > '" & Trim(todayD) & "' and substring(fBODY,10,6) = '" & docNum & "'" 
      fISN = Get_Query_Result(queryString)
      Log.Message("Փաստաթղթի ISN` " & fISN)
      
      ' Մուտք համակարգ VERIFIER2 օգտագործողով
      Login("VERIFIER2")
      
      ' Մուտք հաստատվող վճարային փաստաթղթեր թղթապանակ
'      Call wTreeView.DblClickItem("|Ð³ëï³ïáÕ II ²Þî|Ð³ëï³ïíáÕ í×³ñ³ÛÇÝ ÷³ëï³ÃÕÃ»ñ")
      Set verifyDocuments = New_VerificationDocument()
      verifyDocuments.User = "^A[Del]"
      Call GoToVerificationDocument("|Ð³ëï³ïáÕ II ²Þî|Ð³ëï³ïíáÕ í×³ñ³ÛÇÝ ÷³ëï³ÃÕÃ»ñ",verifyDocuments)
      If Not wMDIClient.WaitVBObject("frmPttel", 2000).Exists Then
            Log.Error("Հաստատվող վճարային փաստաթղթեր թղթապանակը չի բացվել")
            Exit Sub
      End If
      
      ' Հաստատել վճարային փաստաթուղթը
      colN = 3
      action = c_ToConfirm
      doNum = 1
      doActio = "Ð³ëï³ï»É"
      status = ConfirmContractDoc(colN, docNum, action, doNum, doActio)
      If Not status Then
            Log.Error("Վճարման հանձնարարագիրը չի վավերացվել")
            Exit Sub
      End If
      
      BuiltIn.Delay(1000)
      wMDIClient.VBObject("frmPttel").Close
      
      ' Մուտք համակարգ ARMSOFT օգտագործողով
      Login("ARMSOFT")
      ' Մուտք արտաքին փոխանցումների ԱՇՏ
      Call ChangeWorkspace(c_ExternalTransfers)
      ' Մուտք BankMail թղթապանակ
      workEnvName = "|²ñï³ùÇÝ ÷áË³ÝóáõÙÝ»ñÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|àõÕ³ñÏí³Í  Ñ³ÝÓÝ³ñ³ñ³·ñ»ñ|àõÕ³ñÏí³Í BankMail"
      workEnv = "ՈՒղարկված BankMail"
      stRekName = "PERN"
      endRekName = "PERK"
      wStatus = False
      state =  AccessFolder(workEnvName, workEnv, stRekName, todayDMY, endRekName, todayDMY, wStatus, isnRekName, fISN)
      If Not state Then
          Log.Error("Սխալ՝ Ուղարկված BankMail  թղթապանակ մուտք գործելիս")
          Exit Sub
      End If
      
      ' Վճարման հանձնարաարգրի պայմանագրի առկայության ստուգում
      colN = 1
      action = c_SendToPartEd
      doNum = 2
      doActio = "Î³ï³ñ»É"
      refuse = "Üå³ï³Ï"
      basis = "REFUSE"
      status = ExcludeContractDoc(colN, docNum, action, basis, refuse, doNum, doActio)
      If Not status Then
            Log.Error("Վճարման հանձնարարագրի պայմանագիրն առկա չէ ուղարկված BankMail թղթապանակում")
            Exit Sub
      End If
      
      ' Փակել ուղարկված BankMail թղթապանակը
      BuiltIn.Delay(1000)
      wMDIClient.VBObject("frmPttel").Close
      
      ' Մուտք արտաքին փոփանցումների ԱՇՏ
      Call ChangeWorkspace(c_ExternalTransfers) 
      ' Մուտք Մասնակի խմբագրվող հանձնարարագրեր թղթապանակ
      workEnvName =  "|²ñï³ùÇÝ ÷áË³ÝóáõÙÝ»ñÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|àõÕ³ñÏíáÕ Ñ³ÝÓÝ³ñ³ñ³·ñ»ñ|Ø³ëÝ³ÏÇ ËÙµ³·ñíáÕ Ñ³ÝÓÝ³ñ³ñ³·ñ»ñ"
      workEnv =  "Մասնակի խմբագրվող հանձնարարագրեր "
      state = AccessFolder(workEnvName, workEnv, stRekName, todayDMY, endRekName, todayDMY, wStatus, isnRekName, fISN)
      If Not state Then
          Log.Error("Մասնակի խմբագրվող հանձնարարագրեր թղթապանակը չի բացվել")
          Exit Sub
      End If
      
      ' Մերժել պայմանագիրը
      colN = 1
      action =  c_ToRefuse
      doNum =  1
      basis = "AIM"
      refuse = "Üå³ï³Ï"
      doActio = "Î³ï³ñ»É"
      Call RejectPaymentOrder(colN, docNum, action, memOrdfISN, ordDocNum, basis, refuse, doNum, doActio)
      
      Log.Message("Մերժվող վճարման հանձնարարագրի ISN`  " & memOrdfISN)
      Log.Message("Մերժվող վճարման հանձնարարագրի համար՝  " & ordDocNum)
      
      BuiltIn.Delay(1000)
      wMDIClient.VBObject("frmPttel").Close
      
      ' Մուտք աշխատանքային փաստաթղթեր թղթապանակ
      directFolder = "|²ñï³ùÇÝ ÷áË³ÝóáõÙÝ»ñÇ ²Þî|²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ"
      folderName = "Աշխատանքային փաստաթղթեր"
      wUser = 77
      state = EnterFolder(directFolder, folderName, todayDMY, todayDMY, cur, wUser, wDocType)
      If Not state Then
            Log.Error("Սխալ` աշխատանքային փաստաթղթեր թղթապանակ մուտք գործելիս")
            Exit Sub
      End If
      
      ' Հաշվառել վճարային փաստաթուղթը
      colN = 2
      action = c_DoTrans
      doNum = 5
      doActio = "²Ûá"
      status = ConfirmContractDoc(colN, ordDocNum, action, doNum, doActio)
      If Not status Then
            Log.Error("Հիշարար օրդերը չի հաշվառվել")
            Exit Sub
      End If
      
      ' Փակել աշխատանքային փաստաթղթեր թղթապանակը
      BuiltIn.Delay(1000)
      wMDIClient.VBObject("frmPttel").Close
      
      ' Մուտք հեռահար համակարգերի ԱՇՏ
      Call ChangeWorkspace(c_RemoteSyss) 
      ' Մուտք Մուտքային հաղորդագրությունների դիտում թղթապանակ
      workEnvName =  "|Ð»é³Ñ³ñ Ñ³Ù³Ï³ñ·»ñ|ÂÕÃ³å³Ý³ÏÝ»ñ|Øáõïù³ÛÇÝ Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñÇ ¹ÇïáõÙ|Øáõïù³ÛÇÝ Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñÇ ¹ÇïáõÙ(ÀÝ¹Ñ³Ýáõñ)"
      workEnv =  "Մուտքային հաղորդագրությունների դիտում (Ընդհանուր) "
      stRekName = "SDATE"
      endRekName = "EDATE"
      state = AccessFolder(workEnvName, workEnv, stRekName, todayDMY, endRekName, todayDMY, wStatus, isnRekName, fISN)
      If Not state Then
            Log.Error("Սխալ՝ Մուտքային հաղորդագրությունների դիտում թղթապանակ մուտք գործելիս")
            Exit Sub
      End If
      
      BuiltIn.Delay(2000)
      ' Գործողություններ /  Բոլոր գործողություններ 
      Call wMainForm.MainMenu.Click(c_AllActions)
      ' Կապակցված բանկի փաստաթուղթ գործողության կատարում
      Call wMainForm.PopupMenu.Click(c_LinkedBankDoc)
      
      If Not wMDIClient.WaitVBObject("frmPttel", 2000).Exists Then
            Log.Error("Մուտքային հաղորդագրությունների դիտում թղթապանակը չի բացվել")
            Exit Sub 
      End If
      
      ' Ստուգում որ վճարման հանձնարարագրի կարգավիճակը 55 է
      wState = "55"
      If Trim(wMDIClient.VBObject("frmPttel_2").VBObject("tdbgView").Columns.Item(1).Text) <> Trim(wState) Then
            Log.Error("Փաստաթղթի սխալ վիճակ")
      End If
      
      BuiltIn.Delay(1000)
      wMDIClient.VBObject("frmPttel_2").Close
      
      ' SQL ստուգում CB_MESSAGES աղյուսակում
      queryString = " Select COUNT(*)  from CB_MESSAGES where fDATE > '" & Trim(todayD) & "' AND fSTATE = '9' AND fISN = "& fISN &" "
      BuiltIn.Delay(1000)
      rowCount = Get_Query_Result(queryString)
      Log.Message("CB_MESSAGES աղյուսակում տողերի քանակ՝ " & rowCount)
      If rowCount <> 1 Then
          Log.Error("CB_MESSAGES աղյուսակում այսօրվա հարցումով միայն 1 տող պետք է գտնվի")
          Exit Sub
      End If
      
      ' SQL ստուգում HI աղյուսակում
      queryString = " SELECT  COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                              " AND fTYPE = '01'  AND fCUR = '000' AND fCURSUM = '250.00'  AND fOP = 'TRF'  " &_
                              " AND fDBCR = 'C' AND fSUID = '77' AND fBASEBRANCH = '00' AND fBASEDEPART = '1' AND fSUM = '250.00'"
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 

      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                               " AND fTYPE = '01'  AND fCUR = '000' AND fCURSUM = '250.00'  AND fOP = 'TRF'  " &_
                              " AND fDBCR = 'D' AND fSUID = '77' AND fBASEBRANCH = '00' AND fBASEDEPART = '1' AND fSUM = '250.00'"
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 
              
      ' SQL ստուգում 
      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                              " AND fTYPE = '01'  AND fCUR = '000' AND fCURSUM = '500.00'  AND fOP = 'FEE'  " &_
                              " AND fDBCR = 'C' AND fSUID = '77' AND fBASEBRANCH = '00' AND fBASEDEPART = '1' AND fSUM = '500.00' "
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 

      queryString = " SELECT COUNT(*) FROM HI WHERE fBASE= " & fISN & _
                               " AND fTYPE = '01'  AND fCUR = '000' AND fCURSUM = '500.00'  AND fOP = 'FEE'  " &_
                               " AND fDBCR = 'D' AND fSUID = '77' AND fBASEBRANCH = '00' AND fBASEDEPART = '1' AND fSUM = '500.00' "
      sqlValue = 1
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
      If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
      End If 
     
     ' SQL ստուգում MEMORDERS աղյուսակում
      queryString = " Select COUNT(*) From MEMORDERS Where fDATE = '" & Trim(todayD) & "' And fSTATE = 5 " &_
                               " And fCOMPLETED = '1' And fDOCNUM = '" & ordDocNum & "' And fISN = " & memOrdfISN
      rowCount = Get_Query_Result(queryString)
      Log.Message("MEMORDERS աղյուսակում տողերի քանակ՝ " & rowCount)
      If rowCount <> 1 Then
          Log.Error("MEMORDERS աղյուսակում այսօրվա հարցումով միայն 1 տող պետք է գտնվի")
          Exit Sub
      End If
              
      ' Փակել ՀԾ-Բանկ ծրագիրը
      Call Close_AsBank()
End Sub