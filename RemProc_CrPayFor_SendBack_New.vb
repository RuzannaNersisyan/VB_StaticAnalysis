'USEUNIT Library_Common 
'USEUNIT RemoteService_Library 
'USEUNIT Library_CheckDB
'USEUNIT Constants
'USEUNIT Subsystems_SQL_Library
'USEUNIT Payment_Order_ConfirmPhases_Library 
'USEUNIT BankMail_Library
'USEUNIT Library_Contracts

'Test Case 165370

Sub RemoteProc_CrPayFor_SendBack_New_Test()

      Dim paramName, paramValue, fISN, confPath, confInput, docNum, rowCount
      Dim todayD, todayDMY, cliCode, payerAcc, payer, payerAddress, receiverAcc, receiver, receiverAddress, _
              receiverCity, receiverCountry, receiverBankBic, receiverBank, receiverBankAddress, amount, aim, curISO, cur
      Dim state, workEnvName, workEnv, stRekName, endRekName, wStatus, isnRekName
      Dim  directFolder, folderName, wUser, wDocType, delayTime
      Dim colN, action, basis, refuse, doNum, doActio
      Dim queryStr, queryStrin, queryString, frmPttel, clCount
      Dim  system, cliMask, wState, status, direction, dirName
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
      confPath = "X:\Testing\RemAutoProc\Verify13_CrPayFor.txt"
      confInput = Input_Config(confPath)
      If Not confInput Then
          Log.Error("Կարգավորումները չեն ներմուծվել")
         Exit Sub
      End If
      
      ' Ձևավորում ենք docNum-ը պատահական գեներացված թվի դիմացից ավելացնելով 0-ներ, այնպես որ լինի 6 նիշ
      Call Randomize()
      docNum = right(String(6, "0") + RTrim(Int(1000 * Rnd)), 6) 
      Log.Message("Փաստաթղթի համար՝  " & docNum)
      
      todayD = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%Y%m%d")
      todayDMY = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%d/%m/%y")
      cliCode =  "00000680"
      payerAcc = "7770000068020101"
      payer = "Kaskade"
      payerAddress = "Payers address"
      receiverAcc = "202047288096" 
      receiver = "Armen Avetisyan" 
      receiverAddress = "Receivers Address"
      receiverCity = "Receivers City"
      receiverCountry = "CA"
      receiverBankBic = "EBILAEADXXX"
      receiverBank = "EMIRATES BANK INTERNATIONAL PJSC (HEAD OFFICE)"
      receiverBankAddress = "UNITED ARAB EMIRATES,DUBAI,BENIYAS STREET, DEIRA" 
      amount = "100.00"
      aim = "TestAim"
      curISO = "USD"
      cur = "001"
    
      ' Տվյալ օրով տվյալների ջնջում
      queryStr = "Delete from  CB_MESSAGES where FORMAT (fDATE, 'dd/MM/yy') = '" & Trim(todayDMY) & "' "  
      Call Execute_SLQ_Query(queryStr)
      BuiltIn.Delay(2000)
      
      queryStrin = " Insert into CB_MESSAGES (fSYSTEM, fSTATE, fCLIENT, fMSGTYPE,fBODY,fSIGN1,fSIGN2) " _
                          & " values (20             " _
                          & "  , 8                   " _
                          & "  , '" & cliCode & "'  " _
                          & "  , 'IBPayFor'   " _
                          & "  , char(13)+char(10)" _
                          & "      + 'DOCNUM:" & Trim(docNum) & "'                    + char(13)+char(10) " _    
                          & "      + 'PAYDATE:" & Trim(todayD) & "'                    + char(13)+char(10) " _
                          & "      + 'PAYERACC:" & payerAcc & "'                + char(13)+char(10) " _
                          & "      + 'payerAddress:" & payerAddress & "'        + char(13)+char(10) " _
                          & "      + 'payer:" & payer & "'                      + char(13)+char(10) " _  
                          & "      + 'receiverAcc:" & receiverAcc & "'          + char(13)+char(10) " _  
                          & "      + 'receiver:" & receiver & "'                + char(13)+char(10) " _ 
                          & "      + 'receiverAddress:" & receiverAddress & "'  + CHAR(13)+char(10) " _
                          & "      + 'receiverCity:" & receiverCity & "'        + CHAR(13)+char(10) " _
                          & "      + 'receiverCountry:" & receiverCountry & "'  + CHAR(13)+char(10) " _
                          & "      + 'receiverBankBic:" & receiverBankBic & "'  + CHAR(13)+char(10) " _    
                          & "      + 'receiverBank:" & receiverBank & "'        + CHAR(13)+char(10) " _ 
                          & "      + 'receiverBankAddress:" & receiverBankAddress & "' + CHAR(13)+char(10) " _
                          & "      + 'AMOUNT:" & amount & "'                    + char(13)+char(10) " _
                          & "      + 'CURR:" & curISO & "'                      + char(13)+char(10) " _ 
                          & "      + 'AIM:" & aim & "'                          + char(13)+char(10) " _    
                          & " , Cast('0x308206EA06092A864886F70D010702A08206DB308206D7020101310B300906052B0E03021A0500300B06092A864886F70D010701A08205AD308205A930820491A003020102020A12F7787F00000000000D300D06092A864886F70D0101050500305131123010060A0992268993F22C6401191602616D31173015060A0992268993F22C640119160761726D736F6674312230200603550403131941726D656E69616E20536F66747761726520526F6F74204341301E170D3034313030383035313634305A170D3035313030383035323634305A303531173015060355040A130E37373730302D3030303030323333311A301806035504031311416E61686974205368617368696B79616E30819F300D06092A864886F70D010101050003818D0030818902818100E41B109B1E9A7F5582AD3631831CC6E9EDB68408598439E53245D815198B5AF472CCC5D8F3FFA2413FAE18FF159B75A7415C5D98B7FC603BD0BAB2E4759A4F5D5CCD410893A92274939C789DC31E5D4B7C3B7FD962124AAAC92A06463F93E547DE89CDE85345054EF66DA2E203A1A36C4F9FE82C190CCC8E1E453B9DB79EEEE30203010001A38203213082031D300E0603551D0F0101FF0404030204F0304406092A864886F70D01090F04373035300E06082A864886F70D030202020080300E06082A864886F70D030402020080300706052B0E030207300A06082A864886F70D030730130603551D25040C300A06082B06010505070302301D0603551D0E0416041441998D5321E77A40E297EB511F5766692E856532301F0603551D230418301680149B946FBBC5063F443D23CFCAF2A313113D10C32E3082012D0603551D1F04820124308201203082011CA0820118A08201148681C66C6461703A2F2F2F434E3D41726D656E69616E253230536F667477617265253230526F6F7425323043412C434E3D7465726D696E616C2C434E3D4344502C434E3D5075626C69632532304B657925323053657276696365732C434E3D53657276696365732C434E3D436F6E66696775726174696F6E2C44433D61726D736F66742C44433D616D3F63657274696669636174655265766F636174696F6E4C6973743F626173653F6F626A656374436C6173733D63524C446973747269627574696F6E506F696E748649687474703A2F2F7465726D696E616C2E61726D736F66742E616D2F43657274456E726F6C6C2F41726D656E69616E253230536F667477617265253230526F6F7425323043412E63726C3082013D06082B060105050701010482012F3082012B3081BD06082B060105050730028681B06C6461703A2F2F2F434E3D41726D656E69616E253230536F667477617265253230526F6F7425323043412C434E3D4149412C434E3D5075626C69632532304B657925323053657276696365732C434E3D53657276696365732C434E3D436F6E66696775726174696F6E2C44433D61726D736F66742C44433D616D3F634143657274696669636174653F626173653F6F626A656374436C6173733D63657274696669636174696F6E417574686F72697479306906082B06010505073002865D687474703A2F2F7465726D696E616C2E61726D736F66742E616D2F43657274456E726F6C6C2F7465726D696E616C2E61726D736F66742E616D5F41726D656E69616E253230536F667477617265253230526F6F7425323043412E637274300D06092A864886F70D0101050500038201010056948359D9E1BB72F164B0159F8D89CB3AB3BA26E739F3F4AEAADCCE6DCF4FC8373ED5BC1C945686D7E7639ADF3FA0C81E3FDE71888D1F42235BA8F18DBAA73CDA0E140DD1A4B5C1366E7B44E32392A68B0BFCBBE08AF8958F66871171BFFCBE8947B0633CF09CEB4EBC94D59A0DB05F36063C6C0ADA541068BF5F30C71693B2BD0082ADD8211172E5AF9C40C12669D6ABD56EA8869D442861D52FA68EC619CDA3F63F97955906496D77FF0D7FEC264D738D660BE9DE7A827D0BE754B85AA9ECB092E0BFD498BD19E8872B6012264F4EBF9B88FFBBB812E50EBB9B03A376D325C8152D15BDBCB638AB5FF191B01D8BCFBB1884D8D3079D64E67991207C72B1563182010530820101020101305F305131123010060A0992268993F22C6401191602616D31173015060A0992268993F22C640119160761726D736F6674312230200603550403131941726D656E69616E20536F66747761726520526F6F74204341020A12F7787F00000000000D300906052B0E03021A0500300D06092A864886F70D010101050004818042A0B20247725B8580C78FCEA1412900999AF1473146B92F93E7CB917194D14744888222B3D732471EC430BF8B301C094D6E15E6C2841072ECA56169217F296C877826CE4EFE1E23C40D2C74CC9791255104743CAC2298CE174ABBCAE48619FB04F36FED9539A015663D3B90660660DC543167EA31FB421B20AB8FA4EAC75CD7' AS VARBINARY(MAX))" _
                          & " , Cast('0x308206EA06092A864886F70D010702A08206DB308206D7020101310B300906052B0E03021A0500300B06092A864886F70D010701A08205AD308205A930820491A003020102020A12F7787F00000000000D300D06092A864886F70D0101050500305131123010060A0992268993F22C6401191602616D31173015060A0992268993F22C640119160761726D736F6674312230200603550403131941726D656E69616E20536F66747761726520526F6F74204341301E170D3034313030383035313634305A170D3035313030383035323634305A303531173015060355040A130E37373730302D3030303030323333311A301806035504031311416E61686974205368617368696B79616E30819F300D06092A864886F70D010101050003818D0030818902818100E41B109B1E9A7F5582AD3631831CC6E9EDB68408598439E53245D815198B5AF472CCC5D8F3FFA2413FAE18FF159B75A7415C5D98B7FC603BD0BAB2E4759A4F5D5CCD410893A92274939C789DC31E5D4B7C3B7FD962124AAAC92A06463F93E547DE89CDE85345054EF66DA2E203A1A36C4F9FE82C190CCC8E1E453B9DB79EEEE30203010001A38203213082031D300E0603551D0F0101FF0404030204F0304406092A864886F70D01090F04373035300E06082A864886F70D030202020080300E06082A864886F70D030402020080300706052B0E030207300A06082A864886F70D030730130603551D25040C300A06082B06010505070302301D0603551D0E0416041441998D5321E77A40E297EB511F5766692E856532301F0603551D230418301680149B946FBBC5063F443D23CFCAF2A313113D10C32E3082012D0603551D1F04820124308201203082011CA0820118A08201148681C66C6461703A2F2F2F434E3D41726D656E69616E253230536F667477617265253230526F6F7425323043412C434E3D7465726D696E616C2C434E3D4344502C434E3D5075626C69632532304B657925323053657276696365732C434E3D53657276696365732C434E3D436F6E66696775726174696F6E2C44433D61726D736F66742C44433D616D3F63657274696669636174655265766F636174696F6E4C6973743F626173653F6F626A656374436C6173733D63524C446973747269627574696F6E506F696E748649687474703A2F2F7465726D696E616C2E61726D736F66742E616D2F43657274456E726F6C6C2F41726D656E69616E253230536F667477617265253230526F6F7425323043412E63726C3082013D06082B060105050701010482012F3082012B3081BD06082B060105050730028681B06C6461703A2F2F2F434E3D41726D656E69616E253230536F667477617265253230526F6F7425323043412C434E3D4149412C434E3D5075626C69632532304B657925323053657276696365732C434E3D53657276696365732C434E3D436F6E66696775726174696F6E2C44433D61726D736F66742C44433D616D3F634143657274696669636174653F626173653F6F626A656374436C6173733D63657274696669636174696F6E417574686F72697479306906082B06010505073002865D687474703A2F2F7465726D696E616C2E61726D736F66742E616D2F43657274456E726F6C6C2F7465726D696E616C2E61726D736F66742E616D5F41726D656E69616E253230536F667477617265253230526F6F7425323043412E637274300D06092A864886F70D0101050500038201010056948359D9E1BB72F164B0159F8D89CB3AB3BA26E739F3F4AEAADCCE6DCF4FC8373ED5BC1C945686D7E7639ADF3FA0C81E3FDE71888D1F42235BA8F18DBAA73CDA0E140DD1A4B5C1366E7B44E32392A68B0BFCBBE08AF8958F66871171BFFCBE8947B0633CF09CEB4EBC94D59A0DB05F36063C6C0ADA541068BF5F30C71693B2BD0082ADD8211172E5AF9C40C12669D6ABD56EA8869D442861D52FA68EC619CDA3F63F97955906496D77FF0D7FEC264D738D660BE9DE7A827D0BE754B85AA9ECB092E0BFD498BD19E8872B6012264F4EBF9B88FFBBB812E50EBB9B03A376D325C8152D15BDBCB638AB5FF191B01D8BCFBB1884D8D3079D64E67991207C72B1563182010530820101020101305F305131123010060A0992268993F22C6401191602616D31173015060A0992268993F22C640119160761726D736F6674312230200603550403131941726D656E69616E20536F66747761726520526F6F74204341020A12F7787F00000000000D300906052B0E03021A0500300D06092A864886F70D010101050004818042A0B20247725B8580C78FCEA1412900999AF1473146B92F93E7CB917194D14744888222B3D732471EC430BF8B301C094D6E15E6C2841072ECA56169217F296C877826CE4EFE1E23C40D2C74CC9791255104743CAC2298CE174ABBCAE48619FB04F36FED9539A015663D3B90660660DC543167EA31FB421B20AB8FA4EAC75CD7' AS VARBINARY(MAX)))"
    
      Call  Execute_SLQ_Query(queryStrin)
        
      ' Մուտք հեռահար համակարգեր ԱՇՏ
      Call ChangeWorkspace(c_RemoteSyss)

      ' Պայմանագրի առկայության ստուգումը մշակման ենթակա մուտքային հաղորդագրություններ (Ընդհանուր) թղթապանակում
      system = "20"
      cliMask = "00000680"
'      msgType = "IBPayFor"
      wState = "êïáñ³·ñáõÃÛáõÝÝ»ñÁ ×Çßï »Ý"
      direction = "|Ð»é³Ñ³ñ Ñ³Ù³Ï³ñ·»ñ|Øß³ÏÙ³Ý »ÝÃ³Ï³ Ùáõïù³ÛÇÝ Ñ³Õáñ¹³·ñáõÃÛáõÝÝ»ñ(ÀÝ¹Ñ³Ýáõñ)"
      dirName = "Մշակման ենթակա մուտքային հաղորդագրություններ (Ընդհանուր)"
      wStatus =  CheckContractRemoteSystems(direction, todayDMY, system, cliMask, msgType, amount, dirName, wState)
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
      
      ' ISN -ի ստացում
      queryString =  " Select fISN from CB_MESSAGES where fDATE > '" & todayD & "' and substring(fBODY,10,6) = '" & docNum & "'"
      fISN = Get_Query_Result(queryString)
      Log.Message("Փաստաթղթի ISN` " & fISN)
            
      ' Մուտք ադմինիստրատորի ԱՇՏ
      Call ChangeWorkspace(c_Admin40)       
      
      ' Մուտք Աշխատանքային փաստաթղթեր թղթապանակ
      wUser = "77"
      wDocType = "CrPayFor"
      directFolder = "|²¹ÙÇÝÇëïñ³ïáñÇ ²Þî 4.0|ÂÕÃ³å³Ý³ÏÝ»ñ|²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ"
      folderName = "Աշխատանքային փաստաթղթեր"
      wStatus =  EnterFolder(directFolder, folderName, todayDMY, todayDMY, cur, wUser, wDocType)
      If Not wStatus Then
          Log.Error("Աշխատանքային փաստաթղթեր թղթապանակը չի բացվել")
          Exit Sub
      End If
      
      ' Պայմանագրի առկայության ստուգում
      colN = 2
      status = CheckContractDoc(colN, docNum)
      If Not status Then
            Log.Error("Վճարման հանձնարարագրի պայմանագրին առկա չէ աշխատանքային փաստաթղթեր թղթապանակում")
            Exit Sub
      End If
      
      BuiltIn.Delay(1000)
      wMDIClient.VBObject("frmPttel").Close
      
      ' Մուտք համակարգ VERIFIER օգտագործողով
      Login("VERIFIER")
      ' Մուտք հաստատվող վճարային փաստաթղթեր թղթապանակ
 '     Call wTreeView.DblClickItem("|Ð³ëï³ïáÕ I ²Þî|Ð³ëï³ïíáÕ í×³ñ³ÛÇÝ ÷³ëï³ÃÕÃ»ñ")
      Set verifyDocuments = New_VerificationDocument()
      verifyDocuments.User = "^A[Del]"
      Call GoToVerificationDocument("|Ð³ëï³ïáÕ I ²Þî|Ð³ëï³ïíáÕ í×³ñ³ÛÇÝ ÷³ëï³ÃÕÃ»ñ",verifyDocuments)
      If Not wMDIClient.WaitVBObject("frmPttel", 3000).Exists Then
            Log.Error("Հաստատվող վճարային փաստաթղթեր թղթապանակը չի բացվել")
            Exit Sub
      End If
      
      ' Հաստատել Միջազգ. վճարման հանձնարաարգիրը 
      colN = 3
      action = c_ToConfirm
      doNum = 1
      doActio = "Ð³ëï³ï»É"
      status = ConfirmContractDoc(colN, docNum, action, doNum, doActio)
      If Not status Then
            Log.Error("Միջազգ. վճարման հանձնարարագիրը չի վավերացվել")
      End If
      
      BuiltIn.Delay(1000)
      wMDIClient.VBObject("frmPttel").Close
      
      ' Մուտք համակարգ ARMSOFT օգտագործողով
      Login("ARMSOFT")
      ' Մուտք արտաքին փոխանցումների ԱՇՏ
      Call ChangeWorkspace(c_Admin40)
      
      ' Մուտք Աշխատանքային փաստաթղթեր թղթապանակ
      wStatus =  EnterFolder(directFolder, folderName, todayDMY, todayDMY, cur, wUser, wDocType)
      If Not wStatus Then
          Log.Error("Աշխատանքային փաստաթղթեր թղթապանակը չի բացվել")
          Exit Sub
      End If
      
      ' Միջազգ. վճարման հանձնարաարգրի առկայության ստուգում
      colN = 2
      status = CheckContractDoc(colN, docNum)
      If Not status Then
            Log.Error("Միջազգ. վճարման հանձնարարագրի պայմանագիրն առկա չէ Աշխատանքային փաստաթղթեր թղթապանակում")
      End If
      
      ' Ստուգում որ պայամանագրի վիճակը "Ուղարկված 3 հաստատ" է
      wState = "àõÕ³ñÏí³Í III Ñ³ëï³ïÙ³Ý"
      BuiltIn.Delay(4000)
      If Trim(wMDIClient.VBObject("frmPttel").VBObject("tdbgView").Columns.Item(10).Text) <> Trim(wState) Then
            Log.Error("Պայմանագրի վիճակը պետք է լինի ուղարկված 3 հաստատ")
      End If
      
      ' Մուտք համակարգ VERIFIER3 օգտագործողով
      Login("VERIFIER3")
      ' Մուտք հաստատվող վճարային փաստաթղթեր թղթապանակ
  '    Call wTreeView.DblClickItem("|Ð³ëï³ïáÕ III ²Þî|Ð³ëï³ïíáÕ í×³ñ³ÛÇÝ ÷³ëï³ÃÕÃ»ñ")
      Set verifyDocuments = New_VerificationDocument()
      verifyDocuments.User = "^A[Del]"
      Call GoToVerificationDocument("|Ð³ëï³ïáÕ III ²Þî|Ð³ëï³ïíáÕ í×³ñ³ÛÇÝ ÷³ëï³ÃÕÃ»ñ",verifyDocuments)
      If Not wMDIClient.WaitVBObject("frmPttel", 3000).Exists Then
            Log.Error("Հաստատվող վճարային փաստաթղթեր թղթապանակը չի բացվել")
            Exit Sub
      End If
      
      ' Հետ ուղարկել Միջազգ. վճարման հանձնարարագիրը 
      colN = 3
      action = c_SendBack
      doNum = 2
      doActio = "Î³ï³ñ»É"
      basis = "SENDBACK"
      refuse = "Refuse"
      status = ExcludeContractDoc(colN, docNum, action, basis, refuse, doNum, doActio)
      If Not status Then
            Log.Error("Միջազգ. վճարման հանձնարարագիրը չի գտնվել և ետ չի ուղարկվել")
      End If
      
      BuiltIn.Delay(1000)
      wMDIClient.VBObject("frmPttel").Close
      
      ' Մուտք համակարգ ARMSOFT օգտագործողով
      Login("ARMSOFT")
      ' Մուտք արտաքին փոխանցումների ԱՇՏ
      Call ChangeWorkspace(c_Admin40)
      
      ' Մուտք Աշխատանքային փաստաթղթեր թղթապանակ
      wStatus =  EnterFolder(directFolder, folderName, todayDMY, todayDMY, cur, wUser, wDocType)
      If Not wStatus Then
          Log.Error("Աշխատանքային փաստաթղթեր թղթապանակը չի բացվել")
          Exit Sub
      End If
      
      ' Վճարման հանձնարաարգրի պայմանագրի առկայության ստուգում
      colN = 2
      status = CheckContractDoc(colN, docNum)
      If Not status Then
            Log.Error("Վճարման հանձնարարագրի պայմանագիրն առկա չէ Աշխատանքային փաստաթղթեր թղթապանակում")
            Exit Sub
      End If
      
      BuiltIn.Delay(1000)
      wMDIClient.VBObject("frmPttel").Close

      ' Մուտք հեռահար համակարգերի ԱՇՏ
      Call ChangeWorkspace(c_RemoteSyss) 
      
      ' Մուտք Ճշտվող հանձնարարագրեր
      workEnvName =  "|Ð»é³Ñ³ñ Ñ³Ù³Ï³ñ·»ñ|ÖßïíáÕ Ñ³ÝÓÝ³ñ³ñ³·ñ»ñ"
      workEnv =  "Ճշտվող հանձնարարագրեր"
      stRekName = "PERN"
      endRekName = "PERK"
      wStatus = False
      state = AccessFolder(workEnvName, workEnv, stRekName, todayDMY, endRekName, todayDMY, wStatus, isnRekName, fISN)
      If Not state Then
          Log.Error("Ճշտվող հանձնարարագրեր թղթապանակը չի բացվել")
          Exit Sub
      End If
      
      ' Մերժել պայմանագիրը
      colN = 1
      action =  c_ToRefuse
      doNum =  2
      basis = "REFUSE"
      doActio = "Î³ï³ñ»É"
      status = ExcludeContractDoc(colN, docNum, action, basis, refuse, doNum, doActio)
      If Not status Then
            Log.Error("Պայմանագիրը չի գտնվել և չի մերժվել")
      End If
      
      ' CB_MESSAGES աղյուսակում տվյալների ստուգում
      queryString = " Select COUNT(*) From CB_MESSAGES where fDATE > '" & todayD & "' And fISN = " & fISN & " And fSTATE = '10' "
      BuiltIn.Delay(1000)
      rowCount = Get_Query_Result(queryString)
      Log.Message("CB_MESSAGES աղյուսակում գտնված տողերի քանակ՝ " & rowCount)
      If rowCount <> 1 Then
          Log.Error("CB_MESSAGES աղյուսակում  SQL հարցումով միայն մեկ տող պետք է լինի և վիճակն էլ լինի 10 ")
          Exit Sub
      End If
      
      ' Տվյալների ստուգում HI աղյուսակում, պետք է  տվյալներ չլինեն
      queryString = " Select COUNT(*) From HI Where fBASE = " & fISN 
      BuiltIn.Delay(1000)
      rowCount = Get_Query_Result(queryString)
      Log.Message("HI աղյուսակում գտնված տողերի քանակ՝ " & rowCount)
      If rowCount <> 0 Then
          Log.Error("HI աղյուսակում չպետք է ձևակերպումներ լինեն ")
          Exit Sub
      End If
      
      ' Փակել ՀԾ-Բանկ ծրագիրը
      Call Close_AsBank()   
End Sub