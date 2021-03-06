Option Explicit
'USEUNIT Library_Common
'USEUNIT Constants
'USEUNIT Subsystems_SQL_Library
'USEUNIT CashInput_Confirmphases_Library
'USEUNIT DAHK_Library_Filter
'USEUNIT Payment_Except_Library
'USEUNIT Library_CheckDB
'USEUNIT SWIFT_International_Payorder_Library
'USEUNIT BankMail_Library
'USEUNIT Library_Contracts
'USEUNIT Library_Colour
'USEUNIT Main_Accountant_Filter_Library
      
'Test Case ID 185140

 ' Խմբային Կանխիկ մուտք փաստաթղթի ստուգում (Անհատ ձեռներեց)
Sub Group_Cash_Entry_Private_Entrepreneur_Test()

      Dim newAcc, fDATE, sDATE, grCashInput, workingDocuments, addIntoCassa
      Dim folderDirect, accChartNum, balAcc, accMAsk, accCur, accType, accName, clName, clCode, incExp, showLimits, _
              oldAccMask, newAccMask, accNote, accNote2, accNote3, cashAcc, showCli, showOthInfo, opDate, endOpDAte,_
              acsBranch, acsDepart, acsType, selectView, exportExcel
      Dim fileName1, fileName2, param, savePath, fName, colN, workingDocs, editGrCashInp, fBODY, action, doNum, doActio
      Dim dbFOLDERS(3) ,  todayDateSQL, todayDateSQL2, dbPAYMENTS(1), ExpMess1, ExpMess2, verifyDocuments
      Dim todayDate, wUser, docType, wName, passNum, cliCode, paySysIn, paySysOut, docISN, dbHI(7)    
      
      fDATE = "20250101"
      sDATE = "20030101"
      Call Initialize_AsBank("bank", sDATE, fDATE)
      Call Create_Connection()
      Call SetParameter("CHECKCHRGINCACC ", "1")
      todayDate = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%d/%m/%y")
      todayDateSQL = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%Y-%m-%d")
      todayDateSQL2 = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%Y%m%d")
      
      ' Մուտք գործել համակարգ ARMSOFT  օգտագործողով 
      Login("ARMSOFT")
      
      ' Մուտք Գլխավոր հաշվապահի ԱՇՏ
      Call ChangeWorkspace(c_ChiefAcc)
      
      ' Մուտք հաշիվներ թղթապանակ, հաշիվներ դիալոգի լրացում
      folderDirect = "|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|Ð³ßÇíÝ»ñ"
      accType = "10"
      accChartNum = "1"
      selectView = "ACCS"
      exportExcel = "0"
      Call OpenAccauntsFolder(folderDirect, accChartNum, balAcc, accMAsk, accCur, accType, accName, clName, clCode, incExp, showLimits, _
                                                     oldAccMask, newAccMask, accNote, accNote2, accNote3, cashAcc, showCli, showOthInfo, opDate, endOpDAte,_
                                                     acsBranch, acsDepart, acsType, selectView, exportExcel )
                                                     
      ' Սպասում է այնքան մինչև "կատարման ընթացքը" վերջանա 
      Call  WaitForExecutionProgress()
                                  
      ' Ստուգում է Հաշիվներ թղթապանակը բացվել է թե ոչ
      If Not WaitForPttel("frmPttel")  Then
             Log.Error("Հաշիվներ թղթապանակը չի բացվել")
      End If
      
      Log.Message("Խմբային կանխիկ մուտք փաստաթղթի ստեղծում")
      Set grCashInput = New_GroupCashInput(2, 2, 1, 0)
      With grCashInput
                .generalTab.wOffice = "00"
                .generalTab.wDepartment = "1"
                .generalTab.wDate =  todayDate 
                .generalTab.cashRegister = "001"
                .generalTab.cashRegisterAcc =  "73030461000" 
                .generalTab.wCurr = "000"
                .generalTab.cashierChar = "021"
                .generalTab.wBase = "ÊÙµ³ÛÇÝ Ï³ÝËÇÏ Ùáõïù"
                .generalTab.wAcc(0)  = "73030201000" 
                .generalTab.wSum(0) = "1,000.00"
                .generalTab.wAim(0) = "Ð³ñÏ»ñÇ Ù³ñáõÙ"
                .generalTab.wAcc(1) = "73030461000" 
                .generalTab.wSum(1) = "3,000.00"
                .generalTab.wAim(1) = "ì³ñÏÇ Ù³ñáõÙ"
                .generalTab.wPayer = "00034855"
                .generalTab.payerLegalStatus = "ֆիզԱնձ"
                .generalTab.wName = "ì³ñ¹³Ý"
                .generalTab.surName = "²ñ³ÙÛ³Ý"
                .generalTab.wIdCheck = "AN123598745"
                .generalTab.wId = "AN123598745"
                .generalTab.idType = "01"
                .generalTab.idTypeCheck = "01"
                .generalTab.idGivenBy = "012"
                .generalTab.idGivenByCheck = "012"
                .generalTab.wCitizenship = "1"
                .generalTab.wCountry = "AM"
                .generalTab.wResidence = "010010338"
                .generalTab.wCity = "ºñ¨³Ý"
                .generalTab.wApartment = "Ñ³Ù³ñ 54"
                .generalTab.wStreet = "Ü³Éµ³Ý¹Û³Ý"
                .generalTab.wHouse = "îáõÝ"
                .generalTab.wEmail = "vardanaramyan@gmail.com"
                .generalTab.wEmailCheck = "vardanaramyan@gmail.com"
                .generalTab.wBirthDate = "01/01/1995"
                .generalTab.idGiveDate = "01/01/2020"
                .generalTab.idValidUntil = "01/01/2030"
                .generalTab.birthDateForCheck = "01/01/1995"
                .generalTab.idGiveDateForCheck = "01/01/2020"
                .generalTab.idValidUntilForCheck = "01/01/2030"
                .chargeTab.office = "00"
                .chargeTab.department = "1"
                .chargeTab.chargeAcc = "000001101"
                .chargeTab.chargeCurr = "001"
                .chargeTab.chargeCurrForCheck = "001"   
                .chargeTab.cbExchangeRate = "400.0000/1"
                .chargeTab.chargeType = "03"
                .chargeTab.chargeAmount = "0.03"
                .chargeTab.chargeAmoForCheck = "0.03"
                .chargeTab.chargePercent = "0.3000"
                .chargeTab.chargePerForCheck = "0.3000"
                .chargeTab.incomeAcc = "000434400"
                .chargeTab.incomeAccCurr = "000"
                .chargeTab.buyAndSell = "1"
                .chargeTab.buyAndSellForCheck = "1"
                .chargeTab.operType = "1"
                .chargeTab.operPlace = "3"
                .chargeTab.operArea = "7"
                .chargeTab.operAreaForCheck = "7"
                .chargeTab.nonResident = 1
                .chargeTab.nonResidentForCheck = 0
                .chargeTab.legalStatus = "22"
                .chargeTab.legalStatusForCheck = "22"
                .chargeTab.comment = "²ñï³ñÅ.ÙÇçí×. ·³ÝÓáõÙ"
                .chargeTab.commentForCheck = "²ñï³ñÅ.ÙÇçí×. ·³ÝÓáõÙ"
                .chargeTab.notGrCash = False
                .coinTab.coin = "0.00"
                .coinTab.coinForCheck = "0.00"
                .coinTab.coinExchangeRate = "0/0"
                .coinTab.coinCBExchangeRate = "0/0"
                .coinTab.coinPayAmount = "0.00"
                .coinTab.coinPayAmountForCheck = "0.00"
                .coinTab.amountCurrForCheck = "0.00"
                .coinTab.roundedAmountForCheck = "0.00"
                .attachedTab.addFiles(0) = Project.Path & "Stores\Attach file\excel.xlsx"
                .attachedTab.fileName(0) = "excel.xlsx"
                .attachedTab.addFiles(1) =  Project.Path & "Stores\Attach file\txtFile.txt"
                .attachedTab.fileName(1) = "txtFile.txt"
                .attachedTab.addLinks(0) =  Project.Path & "Stores\Attach file\Photo.jpg"
                .attachedTab.linkName(0) = "attachedLink_1"
      End With
      
      Call Create_Group_Cash_Input(grCashInput, "ê¨³·Çñ")
      Log.Message("Փաստաթղթի համարը " & grCashInput.generalTab.docNum)
      Log.Message("Փաստաթղթի ISN` " & grCashInput.fIsn)
      
      ' Փակել Հաշիվներ թղթապանակը
      Call Close_Window(wMDIClient, "frmPttel")
      
      ' Կանխիկ մուտք փաստաթղթի սևագրի ստեղծումից հետո SQL ստուգում
      Log.Message( "Կանխիկ մուտք փաստաթղթի սևագրի ստեղծումից հետո SQL ստուգում")
      
      ' DOCS
      fBODY = "" & vbCRLF _
            & "ACSBRANCH:00"& vbCRLF _
            & "ACSDEPART:1"& vbCRLF _
            & "TYPECODE:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28"& vbCRLF _
            & "USERID:  77"& vbCRLF _
            & "DOCNUM:"& grCashInput.generalTab.docNum & vbCRLF _
            & "DATE:"& todayDateSQL2 & vbCRLF _
            & "KASSA:001"& vbCRLF _
            & "ACCDB:73030461000"& vbCRLF _
            & "CUR:000"& vbCRLF _
            & "KASSIMV:021"& vbCRLF _
            & "BASE:ÊÙµ³ÛÇÝ Ï³ÝËÇÏ Ùáõïù"& vbCRLF _
            & "CLICODE:00034855"& vbCRLF _
            & "PAYER:ì³ñ¹³Ý"& vbCRLF _
            & "PAYERLASTNAME:²ñ³ÙÛ³Ý"& vbCRLF _
            & "PASSNUM:AN123598745"& vbCRLF _
            & "PASTYPE:01"& vbCRLF _
            & "PASBY:012"& vbCRLF _
            & "DATEPASS:20200101"& vbCRLF _
            & "DATEEXPIRE:20300101"& vbCRLF _
            & "DATEBIRTH:19950101"& vbCRLF _
            & "CITIZENSHIP:1"& vbCRLF _
            & "COUNTRY:AM"& vbCRLF _
            & "COMMUNITY:010010338"& vbCRLF _
            & "CITY:ºñ¨³Ý"& vbCRLF _
            & "APARTMENT:Ñ³Ù³ñ 54"& vbCRLF _
            & "ADDRESS:Ü³Éµ³Ý¹Û³Ý"& vbCRLF _
            & "BUILDNUM:îáõÝ"& vbCRLF _
            & "EMAIL:vardanaramyan@gmail.com"& vbCRLF _
            & "ACSBRANCHINC:00"& vbCRLF _
            & "ACSDEPARTINC:1"& vbCRLF _
            & "CHRGACC:000001101"& vbCRLF _
            & "TYPECODE2:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28"& vbCRLF _
            & "CHRGCUR:001"& vbCRLF _
            & "CHRGCBCRS:400.0000/1"& vbCRLF _
            & "PAYSCALE:03"& vbCRLF _
            & "CHRGSUM:0.03"& vbCRLF _
            & "PRSNT:0.3"& vbCRLF _
            & "CHRGINC:000434400"& vbCRLF _
            & "CUPUSA:1"& vbCRLF _
            & "CURTES:1"& vbCRLF _
            & "CURVAIR:3"& vbCRLF _
            & "TIME:"& grCashInput.chargeTab.timeForCheck & vbCRLF _
            & "VOLORT:7"& vbCRLF _
            & "NONREZ:1"& vbCRLF _
            & "JURSTAT:22"& vbCRLF _
            & "COMM:²ñï³ñÅ.ÙÇçí×. ·³ÝÓáõÙ"& vbCRLF _
            & ""

      fBODY = Replace(fBODY, "  ", "%")
      Call CheckDB_DOCS(grCashInput.fIsn,"PkCash  ","0",fBODY,1)
      Call CheckQueryRowCount("DOCS","fISN",grCashInput.fIsn,1)

	     ' DOCLOG
      Call CheckQueryRowCount("DOCLOG","fISN",grCashInput.fIsn,2)
      Call CheckDB_DOCLOG(grCashInput.fIsn,"77","N","0"," ",1)
      Call CheckDB_DOCLOG(grCashInput.fIsn,"77","F","0"," ",1)
          
    	 ' FOLDERS
      Set dbFOLDERS(0) = New_DB_FOLDERS()
      With dbFOLDERS(0)
          .fFOLDERID = ".D.GlavBux "
          .fNAME = "PkCash"
          .fKEY = grCashInput.fIsn
          .fISN = grCashInput.fIsn
          .fSTATUS = "1"
          .fCOM = "ÊÙµ³ÛÇÝ Ï³ÝËÇÏ Ùáõïù"
    		    .fDCBRANCH = "00"
    		    .fDCDEPART = "1"
      End With

     	Call CheckQueryRowCount("FOLDERS","fISN",grCashInput.fIsn,1)
      Call CheckDB_FOLDERS(dbFOLDERS(0),1)

    	 ' DOCSG
    	  Call CheckQueryRowCount("DOCSG","fISN",grCashInput.fIsn,8)
      
      ' DOCSATTACH
    	 Call CheckDB_DOCSATTACH(grCashInput.fIsn, Project.Path & "Stores\Attach file\Photo.jpg", 1, "attachedLink_1", 1)
    	 Call CheckDB_DOCSATTACH(grCashInput.fIsn, "excel.xlsx", 0,"" , 1)
    	 Call CheckDB_DOCSATTACH(grCashInput.fIsn, "txtFile.txt", 0, "", 1)
    	 Call CheckQueryRowCount("DOCSATTACH","fISN",GrCashInput.fIsn,3)
      
      ' Մուտք գործել Օգտագործողի Սևագրեր թղթապանակ
      Log.Message ("Մուտք գործել Օգտագործողի Սևագրեր թղթապանակ")
      Call wTreeView.DblClickItem("|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|ú·ï³·áñÍáÕÇ ë¨³·ñ»ñ")
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
    
      ' Ստեղծել Խմբային Կանխիկ մուտք փաստաթուղթը սևագրերից
      Log.Message( "Ստեղծել Խմբային Կանխիկ մուտք փաստաթուղթը սևագրերից")
      If SearchInPttel("frmPttel", 2, grCashInput.fIsn) Then
              BuiltIn.Delay(3000)
              Call wMainForm.MainMenu.Click(c_AllActions)
              Call wMainForm.PopupMenu.Click(c_ToEdit)
              If wMDIClient.WaitvbObject("frmASDocForm", 3000).Exists Then
                  Call Check_Group_Cash_Input(grCashInput)
                  Call ClickCmdButton(1, "Î³ï³ñ»É")
              Else 
                  Log.Error ("Խմբային Կանխիկ մուտք փաստաթուղթը առկա չէ Սևագրեր թղթապանակում")
              End If
      Else
              Log.Error("Օգտագործողի Սևագրեր թղթապանակը չի բացվել")
      End If
      
      savePath = Project.Path & "Stores\Cash_Input_Output\Actual\"
      fName = "GroupCashEntryPrivateEntAct.txt"
      fileName1 = Project.Path & "Stores\Cash_Input_Output\Actual\GroupCashEntryPrivateEntAct.txt"
      fileName2 = Project.Path & "Stores\Cash_Input_Output\Expected\GroupCashEntryPrivateEntExp.txt"
      
      If wMDIClient.WaitVBObject("FrmSpr",3000).Exists Then
           ' Հիշել քաղվածքը
            Call SaveDoc(savePath, fName)

            param = "(\d{2}[/]\d{2}[/]\d{2}.\d{2}[:]\d{2})|(N\s\d.*)"
            Call Compare_Files(fileName1, fileName2,param)
            
            BuiltIn.Delay(1000)
            Call Close_Window(wMDIClient, "FrmSpr")
      Else 
            Log.Error "Խմբային կանխիկ մուտքի քաղվածքը չի բացվել"  
      End If
      
      ' Փակել Օգտագործողսի Սևագրեր թղթապանակը
      Call Close_Window(wMDIClient, "frmPttel")
      
      ' DOCS
      fBODY = "" & vbCRLF _
            & "ACSBRANCH:00"& vbCRLF _
            & "ACSDEPART:1"& vbCRLF _
            & "BLREP:0"& vbCRLF _
            & "TYPECODE:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28"& vbCRLF _
            & "USERID:  77"& vbCRLF _
            & "DOCNUM:"& grCashInput.generalTab.docNum & vbCRLF _
            & "DATE:"& todayDateSQL2 & vbCRLF _
            & "KASSA:001"& vbCRLF _
            & "ACCDB:73030461000"& vbCRLF _
            & "CUR:000"& vbCRLF _
            & "KASSIMV:021"& vbCRLF _
            & "BASE:ÊÙµ³ÛÇÝ Ï³ÝËÇÏ Ùáõïù"& vbCRLF _
            & "CLICODE:00034855"& vbCRLF _
            & "PAYER:ì³ñ¹³Ý"& vbCRLF _
            & "PAYERLASTNAME:²ñ³ÙÛ³Ý"& vbCRLF _
            & "PASSNUM:AN123598745"& vbCRLF _
            & "PASTYPE:01"& vbCRLF _
            & "PASBY:012"& vbCRLF _
            & "DATEPASS:20200101"& vbCRLF _
            & "DATEEXPIRE:20300101"& vbCRLF _
            & "DATEBIRTH:19950101"& vbCRLF _
            & "CITIZENSHIP:1"& vbCRLF _
            & "COUNTRY:AM"& vbCRLF _
            & "COMMUNITY:010010338"& vbCRLF _
            & "CITY:ºñ¨³Ý"& vbCRLF _
            & "APARTMENT:Ñ³Ù³ñ 54"& vbCRLF _
            & "ADDRESS:Ü³Éµ³Ý¹Û³Ý"& vbCRLF _
            & "BUILDNUM:îáõÝ"& vbCRLF _
            & "EMAIL:vardanaramyan@gmail.com"& vbCRLF _
            & "ACSBRANCHINC:00"& vbCRLF _
            & "ACSDEPARTINC:1"& vbCRLF _
            & "CHRGACC:000001101"& vbCRLF _
            & "TYPECODE2:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28"& vbCRLF _
            & "CHRGCUR:001"& vbCRLF _
            & "CHRGCBCRS:400.0000/1"& vbCRLF _
            & "PAYSCALE:03"& vbCRLF _
            & "CHRGSUM:0.03"& vbCRLF _
            & "PRSNT:0.3"& vbCRLF _
            & "CHRGINC:000434400"& vbCRLF _
            & "CUPUSA:1"& vbCRLF _
            & "CURTES:1"& vbCRLF _
            & "CURVAIR:3"& vbCRLF _
            & "TIME:"& grCashInput.chargeTab.timeForCheck & vbCRLF _
            & "VOLORT:7"& vbCRLF _
            & "NONREZ:1"& vbCRLF _
            & "JURSTAT:22"& vbCRLF _
            & "COMM:²ñï³ñÅ.ÙÇçí×. ·³ÝÓáõÙ"& vbCRLF _
            & ""

      fBODY = Replace(fBODY, "  ", "%")
      Call CheckDB_DOCS(grCashInput.fIsn,"PkCash  ","2",fBODY,1)
      Call CheckQueryRowCount("DOCS","fISN",grCashInput.fIsn,1)

	     ' DOCLOG
      Call CheckQueryRowCount("DOCLOG","fISN",grCashInput.fIsn,3)
      Call CheckDB_DOCLOG(grCashInput.fIsn,"77","N","0"," ",1)
      Call CheckDB_DOCLOG(grCashInput.fIsn,"77","F","0"," ",1)
      Call CheckDB_DOCLOG(grCashInput.fIsn,"77","E","2"," ",1)
          
    	 ' FOLDERS
      With dbFOLDERS(0)
          .fFOLDERID = "C.1052440579"
          .fNAME = "PkCash"
          .fKEY = grCashInput.fIsn
          .fISN = grCashInput.fIsn
          .fSTATUS = "5"
          .fCOM = "ÊÙµ³ÛÇÝ Ï³ÝËÇÏ Ùáõïù"
          .fSPEC = "²Ùë³ÃÇí- "& todayDate &" N- "&grCashInput.generalTab.docNum &" ¶áõÙ³ñ-             4,000.00 ²ñÅ.- 000 [Üáñ]"
          .fECOM = "Grouped Cash Deposit Advice"
          .fDCDEPART = ""
          .fDCBRANCH = ""
      End With

     	Set dbFOLDERS(1) = New_DB_FOLDERS()
      With dbFOLDERS(1)
          .fFOLDERID = "Oper."&todayDateSQL2
          .fNAME = "PkCash"
          .fKEY = grCashInput.fIsn
          .fISN = grCashInput.fIsn
          .fSTATUS = "5"
          .fCOM = "ÊÙµ³ÛÇÝ Ï³ÝËÇÏ Ùáõïù"
          .fSPEC = grCashInput.generalTab.docNum &"7770073030461000                         4000.00000Üáñ                                                   77ì³ñ¹³Ý ²ñ³ÙÛ³Ý                  "&_
                            "AN123598745 012 01/01/2020                                      ÊÙµ³ÛÇÝ Ï³ÝËÇÏ Ùáõïù"
          .fECOM = "Grouped Cash Deposit Advice"
          .fDCDEPART = "1"
          .fDCBRANCH = "00"
      End With
          
     	Call CheckQueryRowCount("FOLDERS","fISN",grCashInput.fIsn,2)
      Call CheckDB_FOLDERS(dbFOLDERS(0),1)
      Call CheckDB_FOLDERS(dbFOLDERS(1),1)

    	  ' DOCSG
    	  Call CheckQueryRowCount("DOCSG","fISN",grCashInput.fIsn,10)

    	  ' HI
       Set dbHI(0) = New_DB_HI()
       With dbHI(0)
            .fBASE = grCashInput.fIsn
            .fDATE = todayDateSQL
            .fTYPE = "11"
            .fSUM = "1000.00"
            .fCUR = "000"
            .fCURSUM = "1000.00"
            .fOP = "MSC"
            .fDBCR = "D"
			         .fADB = "1406851809"
            .fACR = "1969111254"
            .fSPEC = grCashInput.generalTab.docNum & "                   Ð³ñÏ»ñÇ Ù³ñáõÙ                    1     1.0000    1"
      End With

      Set dbHI(1) = New_DB_HI()
      With dbHI(1)
            .fBASE = grCashInput.fIsn
            .fDATE = todayDateSQL
            .fTYPE = "11"
            .fSUM = "1000.00"
            .fCUR = "000"
            .fCURSUM = "1000.00"
            .fOP = "MSC"
            .fDBCR = "C"
			         .fADB = "1406851809"
            .fACR = "1969111254"
            .fSPEC = grCashInput.generalTab.docNum & "                   Ð³ñÏ»ñÇ Ù³ñáõÙ                    0     1.0000    1"
      End With

      Set dbHI(2) = New_DB_HI()
      With dbHI(2)
            .fBASE = grCashInput.fIsn
            .fDATE = todayDateSQL
            .fTYPE = "11"
            .fSUM = "3000.00"
            .fCUR = "000"
            .fCURSUM = "3000.00"
            .fOP = "MSC"
            .fDBCR = "D"
		          .fADB = "1406851809"
            .fACR = "1406851809"
            .fSPEC = grCashInput.generalTab.docNum & "                   ì³ñÏÇ Ù³ñáõÙ                      1     1.0000    1"
      End With

      Set dbHI(3) = New_DB_HI()
      With dbHI(3)
            .fBASE = grCashInput.fIsn
            .fDATE = todayDateSQL
            .fTYPE = "11"
            .fSUM = "3000.00"
            .fCUR = "000"
            .fCURSUM = "3000.00"
            .fOP = "MSC"
            .fDBCR = "C"
		          .fADB = "1406851809"
            .fACR = "1406851809"
            .fSPEC = grCashInput.generalTab.docNum & "                   ì³ñÏÇ Ù³ñáõÙ                      0     1.0000    1"
      End With

      Set dbHI(4) = New_DB_HI()
      With dbHI(4)
            .fBASE = grCashInput.fIsn
            .fDATE = todayDateSQL
            .fTYPE = "11"
            .fSUM = "12.00"
            .fCUR = "000"
            .fCURSUM = "12.00"
            .fOP = "FEX"
            .fDBCR = "C"
			         .fADB = "1630171"
            .fACR = "1629198"
            .fSPEC = grCashInput.generalTab.docNum & "                   ²ñï³ñÅ.ÙÇçí×. ·³ÝÓáõÙ             0     1.0000    1"
      End With

      Set dbHI(5) = New_DB_HI()
      With dbHI(5)
            .fBASE = grCashInput.fIsn
            .fDATE = todayDateSQL
            .fTYPE = "11"
            .fSUM = "12.00"
            .fCUR = "001"
            .fCURSUM = "0.03"
            .fOP = "FEX"
            .fDBCR = "D"
			         .fADB = "1630171"
            .fACR = "1629198"
            .fSPEC = grCashInput.generalTab.docNum & "                   ²ñï³ñÅ.ÙÇçí×. ·³ÝÓáõÙ             1   400.0000    1"
      End With
      
      Call Check_DB_HI(dbHI(0),1)
   	  Call Check_DB_HI(dbHI(1),1)
   	  Call Check_DB_HI(dbHI(2),1)
   	  Call Check_DB_HI(dbHI(3),1)
  	   Call Check_DB_HI(dbHI(4),1)
  	   Call Check_DB_HI(dbHI(5),1)
    	 Call CheckQueryRowCount("HI","fBASE",grCashInput.fIsn,6)

      ' Մուտք Աշխատանքային փաստաթղթեր
      Set workingDocs = New_MainAccWorkingDocuments()
      With workingDocs
            .startDate = todayDate
    			     .endDate = todayDate
      End With
   
      Call GoTo_MainAccWorkingDocuments("|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|", workingDocs)
      
      ' Գտնել խմբային կանխիկ մուտք փաստաթուղթը
      colN = 2
      If Not CheckContractDoc(colN, grCashInput.generalTab.docNum) Then
            Log.Error("խմբային կանխիկ մուտք փաստաթուղթն առկա չէ")
      End If
      
      Log.Message("Խմբային կանխիկ մուտք փաստաթղթի արժեքների ստուգում և խմբագրում")
      Set editGrCashInp = New_GroupCashInput(2,1, 1, 1)
      With editGrCashInp
                .generalTab.wOffice = "00"
                .generalTab.wDepartment = "1"
                .generalTab.wDate = todayDate 
                .generalTab.cashRegister = "001"
                .generalTab.cashRegisterAcc = "73030461000"
                .generalTab.wCurr = "000"
                .generalTab.cashierChar = "021"
                .generalTab.wBase = "Î³ÝËÇÏ Ùáõïù ËÙµ³ÛÇÝ "
                .generalTab.wAcc(0)  = "73030201000" 
                .generalTab.wSum(0) = "1,000.00"
                .generalTab.wAim(0) = "Ð³ñÏ»ñÇ Ù³ñáõÙ"
                .generalTab.wAcc(1) = "73030461000" 
                .generalTab.wSum(1) = "3,000.00"
                .generalTab.wAim(1) = "ì³ñÏÇ Ù³ñáõÙ"
                .generalTab.wPayer = "00034857"
                .generalTab.wName = "È¨áÝ"
                .generalTab.surName = "Ø»ÉùáÝÛ³Ý"
                .generalTab.wId = "AM95848546"
                .generalTab.idType = "02"
                .generalTab.idGivenBy = "125"
                .generalTab.wCitizenship = "1"
                .generalTab.wCountry = "AM"
                .generalTab.wResidence = "020030110"
                .generalTab.wCity = "Â³ÉÇÝ"
                .generalTab.wApartment = "µÝ. 9"
                .generalTab.wStreet = "Â³ÉÇÝÛ³Ý"
                .generalTab.wHouse = "Þ»Ýù"
                .generalTab.wEmail = "levonnew@gmail.com"
                .generalTab.wBirthDate = "01/01/1985"
                .generalTab.idGiveDate = "13/12/2019"
                .generalTab.idValidUntil = "13/12/2029"
                .chargeTab.office = "00"
                .chargeTab.department = "1"
                .chargeTab.chargeAcc = "000001101"
                .chargeTab.chargeAccForCheck = "000001101"
                .chargeTab.chargeCurr = "001"
                .chargeTab.chargeCurrForCheck = "001"   
                .chargeTab.cbExchangeRate = "400.0000/1"
                .chargeTab.chargeType = "03"
                .chargeTab.chargeAmount = "9.00"
                .chargeTab.chargeAmoForCheck = "9.00"
                .chargeTab.chargePercent = "90.0000"
                .chargeTab.chargePerForCheck = "90.0000"
                .chargeTab.incomeAcc = "000434400"
                .chargeTab.incomeAccCurr = "000"
                .chargeTab.buyAndSell = "1"
                .chargeTab.buyAndSellForCheck = "1"
                .chargeTab.operType = "1"
                .chargeTab.operPlace = "3"
                .chargeTab.operArea = "7"
                .chargeTab.operAreaForCheck = "7"
                .chargeTab.nonResident = 1
                .chargeTab.nonResidentForCheck = 0
                .chargeTab.legalStatus = "22"
                .chargeTab.legalStatusForCheck = "22"
                .chargeTab.comment = "²ñï³ñÅ. ·³ÝÓáõÙ"
                .chargeTab.commentForCheck = "²ñï³ñÅ.ÙÇçí×. ·³ÝÓáõÙ"
                .chargeTab.clientAgreeData = ""
                .chargeTab.notGrCash = False
                .coinTab.coin = "0.00"
                .coinTab.coinForCheck = "0.00"
                .coinTab.coinExchangeRate = "0/0"
                .coinTab.coinCBExchangeRate = "0/0"
                .coinTab.coinPayAmountForCheck = "0.00"
                .coinTab.amountCurrForCheck = "0.00"
                .coinTab.roundedAmountForCheck = "0.00"
                .attachedTab.addFiles(0) = Project.Path & "Stores\Attach file\Picture.png"
                .attachedTab.fileName(0) = "Picture.png"
                .attachedTab.addLinks(0) =  Project.Path & "Stores\Attach file\excel.xlsx"
                .attachedTab.delFiles(0) = "txtFile.txt"
                .attachedTab.linkName(0) = "attachedLink_1"
      End With
      Call Edit_Group_Cash_Input(grCashInput, editGrCashInp, "Î³ï³ñ»É")
      BuiltIn.Delay(5000)
      
      'DOCS
      fBODY = "" & vbCRLF _
                & "ACSBRANCH:00"& vbCRLF _
                & "ACSDEPART:1"& vbCRLF _
                & "BLREP:0"& vbCRLF _
                & "TYPECODE:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28"& vbCRLF _
                & "USERID:  77"& vbCRLF _
                & "DOCNUM:"& grCashInput.generalTab.docNum & vbCRLF _
                & "DATE:"& todayDateSQL2 & vbCRLF _
                & "KASSA:001"& vbCRLF _
                & "ACCDB:73030461000"& vbCRLF _
                & "CUR:000"& vbCRLF _
                & "KASSIMV:021"& vbCRLF _
                & "BASE:Î³ÝËÇÏ Ùáõïù ËÙµ³ÛÇÝ"& vbCRLF _
                & "CLICODE:00034857"& vbCRLF _
                & "PAYER:È¨áÝ"& vbCRLF _
                & "PAYERLASTNAME:Ø»ÉùáÝÛ³Ý"& vbCRLF _
                & "PASSNUM:AM95848546"& vbCRLF _
                & "PASTYPE:02"& vbCRLF _
                & "PASBY:125"& vbCRLF _
                & "DATEPASS:20191213"& vbCRLF _
                & "DATEEXPIRE:20291213"& vbCRLF _
                & "DATEBIRTH:19850101"& vbCRLF _
                & "CITIZENSHIP:1"& vbCRLF _
                & "COUNTRY:AM"& vbCRLF _
                & "COMMUNITY:020030110"& vbCRLF _
                & "CITY:Â³ÉÇÝ"& vbCRLF _
                & "APARTMENT:µÝ. 9"& vbCRLF _
                & "ADDRESS:Â³ÉÇÝÛ³Ý"& vbCRLF _
                & "BUILDNUM:Þ»Ýù"& vbCRLF _
                & "EMAIL:levonnew@gmail.com"& vbCRLF _
                & "ACSBRANCHINC:00"& vbCRLF _
                & "ACSDEPARTINC:1"& vbCRLF _
                & "CHRGACC:000001101"& vbCRLF _
                & "TYPECODE2:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28"& vbCRLF _
                & "CHRGCUR:001"& vbCRLF _
                & "CHRGCBCRS:400.0000/1"& vbCRLF _
                & "PAYSCALE:03"& vbCRLF _
                & "CHRGSUM:9"& vbCRLF _
                & "PRSNT:90"& vbCRLF _
                & "CHRGINC:000434400"& vbCRLF _
                & "CUPUSA:1"& vbCRLF _
                & "CURTES:1"& vbCRLF _
                & "CURVAIR:3"& vbCRLF _
                & "TIME:"& grCashInput.chargeTab.timeForCheck & vbCRLF _
                & "VOLORT:7"& vbCRLF _
                & "NONREZ:1"& vbCRLF _
                & "JURSTAT:22"& vbCRLF _
                & "COMM:²ñï³ñÅ. ·³ÝÓáõÙ"& vbCRLF _
                & ""

      Call CheckDB_DOCS(grCashInput.fIsn,"PkCash  ","2",fBODY,1)
      Call CheckQueryRowCount("DOCS","fISN",grCashInput.fIsn,1)

	     ' DOCLOG
      Call CheckQueryRowCount("DOCLOG","fISN",grCashInput.fIsn,4)
      Call CheckDB_DOCLOG(grCashInput.fIsn,"77","E","2"," ",2)
          

    	 ' FOLDERS
      With dbFOLDERS(0)
          .fFOLDERID = "C.753710355"
          .fNAME = "PkCash  "
          .fKEY = grCashInput.fIsn
          .fISN = grCashInput.fIsn
          .fSTATUS = "5"
          .fCOM = "ÊÙµ³ÛÇÝ Ï³ÝËÇÏ Ùáõïù"
          .fSPEC = "²Ùë³ÃÇí- "&todayDate&" N- "& grCashInput.generalTab.docNum &" ¶áõÙ³ñ-             4,000.00 ²ñÅ.- 000 [ÊÙµ³·ñíáÕ]"
          .fECOM = "Grouped Cash Deposit Advice"
          .fDCDEPART = ""
          .fDCBRANCH = ""
      End With

      With dbFOLDERS(1)
          .fFOLDERID = "Oper."&todayDateSQL2
          .fNAME = "PkCash  "
          .fKEY = grCashInput.fIsn
          .fISN = grCashInput.fIsn
          .fSTATUS = "5"
          .fCOM = "ÊÙµ³ÛÇÝ Ï³ÝËÇÏ Ùáõïù"
          .fSPEC = grCashInput.generalTab.docNum &"7770073030461000                         4000.00000ÊÙµ³·ñíáÕ                                             77È¨áÝ Ø»ÉùáÝÛ³Ý                  "&_
                           "AM95848546 125 13/12/2019                                       Î³ÝËÇÏ Ùáõïù ËÙµ³ÛÇÝ"
          .fECOM = "Grouped Cash Deposit Advice"
          .fDCDEPART = "1"
          .fDCBRANCH = "00"
      End With
          
     	Call CheckQueryRowCount("FOLDERS","fISN",grCashInput.fIsn,2)
      Call CheckDB_FOLDERS(dbFOLDERS(0),1)
      Call CheckDB_FOLDERS(dbFOLDERS(1),1)

    	  'DOCSATTACH
    	  Call CheckDB_DOCSATTACH(grCashInput.fIsn, Project.Path & "Stores\Attach file\excel.xlsx", 1, "attachedLink_1", 1)
       Call CheckDB_DOCSATTACH(grCashInput.fIsn, Project.Path & "Stores\Attach file\Photo.jpg", 1, "attachedLink_1", 1)
    	  Call CheckDB_DOCSATTACH(grCashInput.fIsn, "Picture.png", 0, "", 1)
    	  Call CheckDB_DOCSATTACH(grCashInput.fIsn, "excel.xlsx", 0, "", 1)
    	  Call CheckQueryRowCount("DOCSATTACH","fISN",grCashInput.fIsn,4)
      
      Log.Message("Խմբային Կանխիկ մուտք փաստաթուղթն ուղարկել հաստատման")
      colN = 2
      action = c_SendToVer
      doNum = 2
      doActio = "Î³ï³ñ»É"
      If Not ConfirmContractDoc(colN, grCashInput.generalTab.docNum, action, doNum, doActio) Then
            Log.Error("Խմբային կանխիկ մուտք փաստաթուղթը չի ուղարկվել հաստատման")
            Exit Sub
      End If
      
      ' Փակել աշխատանքային փաստաթղթեր թղթապանակը
      Call Close_Window(wMDIClient, "frmPttel")
      
	     ' DOCS
      fBODY = "" & vbCRLF _
                & "ACSBRANCH:00"& vbCRLF _
                & "ACSDEPART:1"& vbCRLF _
                & "BLREP:0"& vbCRLF _
                & "TYPECODE:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28"& vbCRLF _
                & "USERID:  77"& vbCRLF _
                & "DOCNUM:"& grCashInput.generalTab.docNum & vbCRLF _
                & "DATE:"& todayDateSQL2 & vbCRLF _
                & "KASSA:001"& vbCRLF _
                & "ACCDB:73030461000"& vbCRLF _
                & "CUR:000"& vbCRLF _
                & "ISTLLCREATED:1"& vbCRLF _
                & "KASSIMV:021"& vbCRLF _
                & "BASE:Î³ÝËÇÏ Ùáõïù ËÙµ³ÛÇÝ"& vbCRLF _
                & "CLICODE:00034857"& vbCRLF _
                & "PAYER:È¨áÝ"& vbCRLF _
                & "PAYERLASTNAME:Ø»ÉùáÝÛ³Ý"& vbCRLF _
                & "PASSNUM:AM95848546"& vbCRLF _
                & "PASTYPE:02"& vbCRLF _
                & "PASBY:125"& vbCRLF _
                & "DATEPASS:20191213"& vbCRLF _
                & "DATEEXPIRE:20291213"& vbCRLF _
                & "DATEBIRTH:19850101"& vbCRLF _
                & "CITIZENSHIP:1"& vbCRLF _
                & "COUNTRY:AM"& vbCRLF _
                & "COMMUNITY:020030110"& vbCRLF _
                & "CITY:Â³ÉÇÝ"& vbCRLF _
                & "APARTMENT:µÝ. 9"& vbCRLF _
                & "ADDRESS:Â³ÉÇÝÛ³Ý"& vbCRLF _
                & "BUILDNUM:Þ»Ýù"& vbCRLF _
                & "EMAIL:levonnew@gmail.com"& vbCRLF _
                & "ACSBRANCHINC:00"& vbCRLF _
                & "ACSDEPARTINC:1"& vbCRLF _
                & "CHRGACC:000001101"& vbCRLF _
                & "TYPECODE2:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28"& vbCRLF _
                & "CHRGCUR:001"& vbCRLF _
                & "CHRGCBCRS:400.0000/1"& vbCRLF _
                & "PAYSCALE:03"& vbCRLF _
                & "CHRGSUM:9"& vbCRLF _
                & "PRSNT:90"& vbCRLF _
                & "CHRGINC:000434400"& vbCRLF _
                & "CUPUSA:1"& vbCRLF _
                & "CURTES:1"& vbCRLF _
                & "CURVAIR:3"& vbCRLF _
                & "TIME:"& grCashInput.chargeTab.timeForCheck & vbCRLF _
                & "VOLORT:7"& vbCRLF _
                & "NONREZ:1"& vbCRLF _
                & "JURSTAT:22"& vbCRLF _
                & "COMM:²ñï³ñÅ. ·³ÝÓáõÙ"& vbCRLF _
                & ""
      fBODY = Replace(fBODY, "  ", "%")
      Call CheckDB_DOCS(GrCashInput.fIsn,"PkCash  ","101",fBODY,1)
      Call CheckQueryRowCount("DOCS","fISN",grCashInput.fIsn,1)

	     ' DOCLOG
      Call CheckQueryRowCount("DOCLOG","fISN",GrCashInput.fIsn,5)
	  Call CheckDB_DOCLOG(grCashInput.fIsn,"77","M","101","àõÕ³ñÏí»É ¿ Ñ³ëï³ïÙ³Ý",1)
      
    	 ' FOLDERS
      With dbFOLDERS(0)
          .fFOLDERID = "C.753710355"
          .fNAME = "PkCash  "
          .fKEY = grCashInput.fIsn
          .fISN = grCashInput.fIsn
          .fSTATUS = "0"
          .fCOM = "ÊÙµ³ÛÇÝ Ï³ÝËÇÏ Ùáõïù"
          .fSPEC = "²Ùë³ÃÇí- "&todayDate&" N- "& grCashInput.generalTab.docNum &" ¶áõÙ³ñ-             4,000.00 ²ñÅ.- 000 [àõÕ³ñÏí³Í I Ñ³ëï³ïÙ³Ý]"
          .fECOM = "Grouped Cash Deposit Advice"
      End With

      With dbFOLDERS(1)
          .fFOLDERID = "Oper."&todayDateSQL2
          .fNAME = "PkCash  "
          .fKEY = grCashInput.fIsn
          .fISN = grCashInput.fIsn
          .fSTATUS = "0"
          .fCOM = "ÊÙµ³ÛÇÝ Ï³ÝËÇÏ Ùáõïù"
          .fSPEC = grCashInput.generalTab.docNum &"7770073030461000                         4000.00000àõÕ³ñÏí³Í I Ñ³ëï³ïÙ³Ý                                 "&_
                            "77È¨áÝ Ø»ÉùáÝÛ³Ý                                                                                  Î³ÝËÇÏ Ùáõïù ËÙµ³ÛÇÝ"
          .fECOM = "Grouped Cash Deposit Advice"
          .fDCDEPART = "1"
          .fDCBRANCH = "00"
      End With
      
	     Set dbFOLDERS(2) = New_DB_FOLDERS()
      With dbFOLDERS(2)
          .fFOLDERID = "Oper."&todayDateSQL2
          .fNAME = "PkCash  "
          .fKEY = grCashInput.fIsn
          .fISN = grCashInput.fIsn
          .fSTATUS = "0"
          .fCOM = "ÊÙµ³ÛÇÝ Ï³ÝËÇÏ Ùáõïù"
          .fSPEC =  grCashInput.generalTab.docNum &"7770073030461000                         4000.00000àõÕ³ñÏí³Í I Ñ³ëï³ïÙ³Ý                                 "&_
                            "77È¨áÝ Ø»ÉùáÝÛ³Ý                                                                                  Î³ÝËÇÏ Ùáõïù ËÙµ³ÛÇÝ"           
          .fECOM = "Grouped Cash Deposit Advice"
          .fDCDEPART = "1"
          .fDCBRANCH = "00"
      End With
          
     	Call CheckQueryRowCount("FOLDERS","fISN",grCashInput.fIsn,3)
      Call CheckDB_FOLDERS(dbFOLDERS(0),1)
      Call CheckDB_FOLDERS(dbFOLDERS(1),1)
	       Call CheckDB_FOLDERS(dbFOLDERS(2),1)
      
      ' Մուտք համակարգ VERIFIER օգտագործողով
      Login("VERIFIER")
      
      Set verifyDocuments = New_VerificationDocument()
      verifyDocuments.User = "^A[Del]"
      Call GoToVerificationDocument("|Ð³ëï³ïáÕ I ²Þî|Ð³ëï³ïíáÕ í×³ñ³ÛÇÝ ÷³ëï³ÃÕÃ»ñ", verifyDocuments)

      If Not wMDIClient.WaitVBObject("frmPttel", 10000).Exists Then
            Log.Error("Հաստատվող փաստաթղթեր թղթապանակը չի բացվել")
      End If
      
      Log.Message("Վավերացնել Խմբային կանխիկ մուտքի փաստաթուղթը")
      colN = 3
      action = c_ToConfirm
      doNum = 1
      doActio = "Ð³ëï³ï»É"
      
      If Not ConfirmContractDoc(colN, grCashInput.generalTab.docNum, action, doNum, doActio) Then
            Log.Error("Խմբային կանխիկ մուտքի փաստաթուղթը չի վավերացվել")
      End If
      
      ' Փակել Հաստատվող վճարային փաստաթղթեր թղթապանակը
      Call Close_Window(wMDIClient, "frmPttel")

	     ' DOCS
      Call CheckDB_DOCS(grCashInput.fIsn,"PkCash  ","15",fBODY,1)
      Call CheckQueryRowCount("DOCS","fISN",grCashInput.fIsn,1)

	     ' DOCLOG
      Call CheckQueryRowCount("DOCLOG","fISN",grCashInput.fIsn,7)
      Call CheckDB_DOCLOG(grCashInput.fIsn,"81","W","102"," ",1)
	     Call CheckDB_DOCLOG(grCashInput.fIsn,"81","C","15"," ",1)
          
    	 ' FOLDERS
      With dbFOLDERS(0)
          .fFOLDERID = "C.753710355"
          .fNAME = "PkCash  "
          .fKEY = grCashInput.fIsn
          .fISN = grCashInput.fIsn
          .fSTATUS = "4"
          .fCOM = "ÊÙµ³ÛÇÝ Ï³ÝËÇÏ Ùáõïù"
          .fSPEC = "²Ùë³ÃÇí- "&todayDate&" N- "& grCashInput.generalTab.docNum &" ¶áõÙ³ñ-             4,000.00 ²ñÅ.- 000 [Ð³ëï³ïí³Í]"
          .fECOM = "Grouped Cash Deposit Advice"
          .fDCDEPART = ""
          .fDCBRANCH = ""
      End With

      With dbFOLDERS(1)
          .fFOLDERID = "Oper."&todayDateSQL2
          .fNAME = "PkCash  "
          .fKEY = grCashInput.fIsn
          .fISN = grCashInput.fIsn
          .fSTATUS = "4"
          .fCOM = "ÊÙµ³ÛÇÝ Ï³ÝËÇÏ Ùáõïù"
          .fSPEC = grCashInput.generalTab.docNum &"7770073030461000                         4000.00000Ð³ëï³ïí³Í                                             77È¨áÝ Ø»ÉùáÝÛ³Ý                  "&_
                          "AM95848546 125 13/12/2019                                       Î³ÝËÇÏ Ùáõïù ËÙµ³ÛÇÝ"
          .fECOM = "Grouped Cash Deposit Advice"
          .fDCDEPART = "1"
          .fDCBRANCH = "00"
      End With

     	Call CheckQueryRowCount("FOLDERS","fISN",grCashInput.fIsn,2)
      Call CheckDB_FOLDERS(dbFOLDERS(0),1)
      Call CheckDB_FOLDERS(dbFOLDERS(1),1)

      Login("ARMSOFT")
      ' Մուտք Գլխավոր հաշվապահի ԱՇՏ
      Call ChangeWorkspace(c_ChiefAcc)
      
      Set workingDocs = New_MainAccWorkingDocuments()
      With workingDocs
            .startDate = todayDate
    			     .endDate = todayDate
      End With
   
      Call GoTo_MainAccWorkingDocuments("|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|", workingDocs)
      
      Log.Message("Վավերացնել Խմբային Կանխիկ մուտքի փաստաթուղթը")
      colN = 2
      action = c_ToConfirm
      doNum = 1
      doActio = "Ð³ëï³ï»É"
      
      If Not ConfirmContractDoc(colN, grCashInput.generalTab.docNum, action, doNum, doActio) Then
            Log.Error("Խմբային Կանխիկ մուտքի փաստաթուղթը չի վավերացվել")
      End If
      
      ' Փակել Աշխատանքային փաստաթղթեր թղթապանակը
      Call Close_Window(wMDIClient, "frmPttel")
      
      ' DOCS
      Call CheckDB_DOCS(grCashInput.fIsn,"PkCash  ","11",fBODY,1)
      Call CheckQueryRowCount("DOCS","fISN",grCashInput.fIsn,1)

	     ' DOCLOG
      Call CheckQueryRowCount("DOCLOG","fISN",grCashInput.fIsn,9)
      Call CheckDB_DOCLOG(grCashInput.fIsn,"77","W","16"," ",1)
	     Call CheckDB_DOCLOG(grCashInput.fIsn,"77","M","11","¶ñ³Ýóí»É »Ý Ó¨³Ï»ñåáõÙÝ»ñÁ",1)
          
      'FOLDERS
      Call CheckQueryRowCount("FOLDERS","fISN",grCashInput.fIsn,0)
      
      'HI     
      With dbHI(0)
            .fBASE = grCashInput.fIsn
            .fDATE = todayDateSQL
            .fTYPE = "01"
            .fSUM = "1000.00"
            .fCUR = "000"
            .fCURSUM = "1000.00"
            .fOP = "MSC"
            .fDBCR = "D"
			         .fADB = "1406851809"
            .fACR = "1969111254"
            .fSPEC = grCashInput.generalTab.docNum & "021                Ð³ñÏ»ñÇ Ù³ñáõÙ                    1     1.0000    1"
			         .fBASEBRANCH = "00"
			         .fBASEDEPART = "1"
      End With

      With dbHI(1)
            .fBASE = grCashInput.fIsn
            .fDATE = todayDateSQL
            .fTYPE = "01"
            .fSUM = "1000.00"
            .fCUR = "000"
            .fCURSUM = "1000.00"
            .fOP = "MSC"
            .fDBCR = "C"
			         .fADB = "1406851809"
            .fACR = "1969111254"
            .fSPEC = grCashInput.generalTab.docNum & "                   Ð³ñÏ»ñÇ Ù³ñáõÙ                    0     1.0000    1                                                                        È¨áÝ Ø»ÉùáÝÛ³Ý                  "
			         .fBASEBRANCH = "00"
			         .fBASEDEPART = "1"
	     End With

      With dbHI(2)
            .fBASE = grCashInput.fIsn
            .fDATE = todayDateSQL
            .fTYPE = "01"
            .fSUM = "3000.00"
            .fCUR = "000"
            .fCURSUM = "3000.00"
            .fOP = "MSC"
            .fDBCR = "D"
			         .fADB = "1406851809"
            .fACR = "1406851809"
            .fSPEC = grCashInput.generalTab.docNum & "021                ì³ñÏÇ Ù³ñáõÙ                      1     1.0000    1"
			         .fBASEBRANCH = "00"
			         .fBASEDEPART = "1"
      End With

      With dbHI(3)
            .fBASE = grCashInput.fIsn
            .fDATE = todayDateSQL
            .fTYPE = "01"
            .fSUM = "3000.00"
            .fCUR = "000"
            .fCURSUM = "3000.00"
            .fOP = "MSC"
            .fDBCR = "C"
		          .fADB = "1406851809"
            .fACR = "1406851809"
            .fSPEC = grCashInput.generalTab.docNum & "                   ì³ñÏÇ Ù³ñáõÙ                      0     1.0000    1                                                                        È¨áÝ Ø»ÉùáÝÛ³Ý                  "
			         .fBASEBRANCH = "00"
			         .fBASEDEPART = "1"
      End With

      With dbHI(4)
            .fBASE = grCashInput.fIsn
            .fDATE = todayDateSQL
            .fTYPE = "01"
            .fSUM = "3600.00"
            .fCUR = "000"
            .fCURSUM = "3600.00"
            .fOP = "FEX"
            .fDBCR = "C"
			         .fADB = "1630171"
            .fACR = "1629198"
            .fSPEC = grCashInput.generalTab.docNum & "                   ²ñï³ñÅ. ·³ÝÓáõÙ                   0     1.0000    1"
			         .fBASEBRANCH = "00"
			         .fBASEDEPART = "1"
      End With

      With dbHI(5)
            .fBASE = grCashInput.fIsn
            .fDATE = todayDateSQL
            .fTYPE = "01"
            .fSUM = "3600.00"
            .fCUR = "001"
            .fCURSUM = "9.00"
            .fOP = "FEX"
            .fDBCR = "D"
			         .fADB = "1630171"
            .fACR = "1629198"
            .fSPEC = grCashInput.generalTab.docNum & "021                ²ñï³ñÅ. ·³ÝÓáõÙ                   1   400.0000    1"
			         .fBASEBRANCH = "00"
			         .fBASEDEPART = "1"
      End With

	     Set dbHI(6) = New_DB_HI()
      With dbHI(6)
            .fBASE = grCashInput.fIsn
            .fDATE = todayDateSQL
            .fTYPE = "CE"
            .fSUM = "3600.00"
            .fCUR = "001"
            .fCURSUM = "9.00"
            .fOP = "PUR"
            .fDBCR = "D"
			         .fADB = "-1"
            .fACR = "-1"
            .fSPEC = "% "
			         .fBASEBRANCH = "00"
			         .fBASEDEPART = "1"
      End With
      
      Call Check_DB_HI(dbHI(0),1)
   	  Call Check_DB_HI(dbHI(1),1)
   	  Call Check_DB_HI(dbHI(2),1)
   	  Call Check_DB_HI(dbHI(3),1)
  	   Call Check_DB_HI(dbHI(4),1)
  	   Call Check_DB_HI(dbHI(5),1)
  	   Call Check_DB_HI(dbHI(6),1)
	     Call CheckQueryRowCount("HI","fBASE",grCashInput.fIsn,7)

      'PAYMENTS
    	 Set dbPAYMENTS(0) = New_DB_PAYMENTS()
      With dbPAYMENTS(0)
            .fISN = grCashInput.fIsn
            .fDOCTYPE = "PkCash"
            .fDATE = todayDateSQL
            .fSTATE = "11"
            .fDOCNUM = grCashInput.generalTab.docNum
            .fCLIENT = "00034857"
            .fACCDB = "7770073030461000"
            .fPAYER = "È¨áÝ Ø»ÉùáÝÛ³Ý"
            .fCUR = "000"
            .fSUMMA = "4000.00"
            .fSUMMAAMD = "4000.00"
            .fSUMMAUSD = "10.00"
            .fCOM = "Î³ÝËÇÏ Ùáõïù ËÙµ³ÛÇÝ"
            .fPASSPORT = "AM95848546 125 13/12/2019"
            .fCOUNTRY = "AM"
            .fACSBRANCH = "00 "
            .fACSDEPART = "1  "
      End With
      Call CheckDB_PAYMENTS(dbPAYMENTS(0),1)
	     Call CheckQueryRowCount("PAYMENTS","fISN",grCashInput.fIsn,1)
      
      Log.Message("Մուտք Հաշվառված վճարային փաստաթղթեր թղթապանակ")
      folderDirect = "|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|Ð³ßí³éí³Í í×³ñ³ÛÇÝ ÷³ëï³ÃÕÃ»ñ"
      selectView = "Payments"
      cliCode = ""
      Call OpenAccPaymentDocFolder(folderDirect, todayDate, todayDate, wUser, docType, wName, passNum, cliCode,_
                                                                      paySysIn, paySysOut, acsBranch, acsDepart, docISN, selectView, exportExcel)
      
      ' Գտնել խմբային կանխիկ մուտք փաստաթուղթը
      colN = 2
      If CheckContractDoc(colN, grCashInput.generalTab.docNum) Then
          ' Կատարել բոլոր գործողությունները
          Call wMainForm.MainMenu.Click(c_AllActions)
          ' Խմբագրել
          Call wMainForm.PopupMenu.Click(c_View)
          
          If wMDIClient.WaitVBObject("frmASDocForm", 15000).Exists Then
                  Log.Message("Ստուգել Խմբային կանխիկ մուտք փաստաթղթի տվյալները")
                  editGrCashInp.generalTab.gridRowCount = 2
                  Call Check_Group_Cash_Input(editGrCashInp)
                  ' Կատարել կոճակի սեղմում
                  Call ClickCmdButton(1, "OK")
          
                  BuiltIn.Delay(3000)
                  Log.Message("Ջնջել Խմբային կանխիկ մուտք փաստաթուղթը")
                  ExpMess1 = "ö³ëï³ÃáõÕÃÁ çÝç»ÉÇë` ÏÑ»é³óí»Ý Ýñ³ Ñ»ï Ï³åí³Í ËÙµ³ÛÇÝ " & vbCrLf & "Ó¨³Ï»ñåáõÙÝ»ñÁ"
                  ExpMess2 = "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ"
          
                  If Not DeleteGroupAction(ExpMess1, ExpMess2) Then
                      Log.Error "Խմբային կանխիկ մուտք փաստաթուղթը չի ջնջվել",,, ErrorColor
                  End If
          Else
                  Log.Error("Խմբային կանխիկ մուտքի փաստաթուղթը չի բացվել")
          End If
      Else
            Log.Error("խմբային կանխիկ մուտք փաստաթուղթն առկա չէ Հաշվառված վճարային փաստաթղթեր թղթապանակում")
      End If

      ' Փակել Հաշվառված վճարային փաստաթղթեր թղթապանակը
      Call Close_Window(wMDIClient, "frmPttel")
      
      ' FOLDERS
       With dbFOLDERS(0)
          .fFOLDERID = ".R."&todayDateSQL2
          .fNAME = "PkCash  "
          .fKEY = grCashInput.fIsn
          .fISN = grCashInput.fIsn
          .fSTATUS = "0"
          .fCOM = ""
          .fSPEC = Left_Align(Get_Compname_DOCLOG(grCashInput.fIsn), 16) & "GlavBux ARMSOFT                       1111 "
          .fECOM = ""
          .fDCDEPART = "1"
          .fDCBRANCH = "00"
      End With

      Call CheckQueryRowCount("FOLDERS","fISN",grCashInput.fIsn,1)
      Call CheckDB_FOLDERS(dbFOLDERS(0),1)
      
	     ' DOCS
      Call CheckDB_DOCS(grCashInput.fIsn,"PkCash  ","999",fBODY,1)
      Call CheckQueryRowCount("DOCS","fISN",grCashInput.fIsn,1)
      
	     ' DOCLOG
      Call CheckDB_DOCLOG(grCashInput.fIsn,"77","D","999"," ",1)
      Call CheckQueryRowCount("DOCLOG","fISN",grCashInput.fIsn,10)

      ' Փակել ծրագիրը
      Call Close_AsBank()
      
End Sub