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
      
'Test Case ID 185142

 ' Խմբային Կանխիկ մուտք փաստաթղթի ստուգում (Ֆիզիկական անձ)
Sub Group_Cash_Entry_Physical_Person_Test()

      Dim newAcc, fDATE, sDATE, grCashInput, workingDocuments, addIntoCassa
      Dim folderDirect, accChartNum, balAcc, accMAsk, accCur, accType, accName, clName, clCode, incExp, showLimits, _
              oldAccMask, newAccMask, accNote, accNote2, accNote3, cashAcc, showCli, showOthInfo, opDate, endOpDAte,_
              acsBranch, acsDepart, acsType, selectView, exportExcel 
      Dim fileName1, fileName2, param, savePath, fName, colN, workingDocs, editGrCashInp, fBODY, action, doNum, doActio
      Dim dbFOLDERS(3) ,  todayDateSQL, dbPAYMENTS(1), dbHI(12), todayDateSQL2, verifyDocuments
      Dim todayDate, wUser, docType, wName, passNum, cliCode, paySysIn, paySysOut, docISN, ExpMess1, ExpMess2     

      fDATE = "20250101"
      sDATE = "20030101"
      Call Initialize_AsBank("bank", sDATE, fDATE)
      Call Create_Connection()
      Call SetParameter("CHECKCHRGINCACC ", "1")
      todayDate = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%d/%m/%y")
      todayDateSQL = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%Y%m%d")
      todayDateSQL2 = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%Y-%m-%d")

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
                .generalTab.cashRegisterAcc =  "73030121000" 
                .generalTab.wCurr = "001"
                .generalTab.cashierChar = "021"
                .generalTab.wBase = "ÊÙµ³ÛÇÝ Ï³ÝËÇÏ Ùáõïù"
                .generalTab.wAcc(0)  = "73030121000" 
                .generalTab.wSum(0) = "1,000.00"
                .generalTab.wAim(0) = "Ð³ñÏ»ñÇ Ù³ñáõÙ"
                .generalTab.wAcc(1) = "73030381000" 
                .generalTab.wSum(1) = "5,000.00"
                .generalTab.wAim(1) = "ì³ñÏÇ Ù³ñáõÙ"
                .generalTab.wPayer = "00034854"
                .generalTab.payerLegalStatus = "ֆիզԱնձ"
                .generalTab.wName = "²ñï³Ï"
                .generalTab.surName = "Ð³Ûñ³å»ïÛ³Ý"
                .generalTab.wId = "AN524685478"
                .generalTab.wIdCheck = "AN524685478"
                .generalTab.idType = "09"
                .generalTab.idGivenBy = "012"
                .generalTab.idGivenByCheck = "012"
                .generalTab.idTypeCheck = "09"
                .generalTab.wCitizenship = "1"
                .generalTab.wCountry = "AM"
                .generalTab.wResidence = "010010635"
                .generalTab.wCity = "ºñ¨³Ý"
                .generalTab.wApartment = "µÝ. 18 "
                .generalTab.wStreet = "²µáíÛ³Ý"
                .generalTab.wHouse = "Þ»Ýù"
                .generalTab.wEmail = "artakhayrapetyan@gmail.com"
                .generalTab.wEmailCheck = "artakhayrapetyan@gmail.com"
                .generalTab.birthDateForCheck = "01/01/1991"
                .generalTab.idGiveDateForCheck = "01/01/2020"
                .generalTab.idValidUntilForCheck = "01/01/2035"
                .generalTab.wBirthDate = "01/01/1991"
                .generalTab.idGiveDate = "01/01/2020"
                .generalTab.idValidUntil = "01/01/2035"
                .chargeTab.office = "00"
                .chargeTab.department = "1"
                .chargeTab.chargeAcc = "000001101"
                .chargeTab.chargeAccForCheck = ""
                .chargeTab.chargeCurr = "001"
                .chargeTab.chargeCurrForCheck = "001"   
                .chargeTab.cbExchangeRate = "400.0000/1"
                .chargeTab.chargeType = "03"
                .chargeTab.chargeAmount = "18.00"
                .chargeTab.chargeAmoForCheck = "18.00"
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
                .chargeTab.legalStatus = "21"
                .chargeTab.legalStatusForCheck = "21"
                .chargeTab.comment = "²ñï³ñÅ.ÙÇçí×. ·³ÝÓáõÙ"
                .chargeTab.commentForCheck = "²ñï³ñÅ.ÙÇçí×. ·³ÝÓáõÙ"
                .chargeTab.clientAgreeData = ""
                .chargeTab.notGrCash = False
                .coinTab.coin = "100.00"
                .coinTab.coinForCheck = "100.00"
                .coinTab.coinPayCurr = "000"
                .coinTab.coinPayAcc = "000001100"
                .coinTab.coinExchangeRate = "370.0000/1"
                .coinTab.coinCBExchangeRate = "400.0000/1"
                .coinTab.coinBuyAndSell = "2"
                .coinTab.coinPayAmount = "37,000.00"
                .coinTab.coinPayAmountForCheck = "37,000.00"
                .coinTab.amountWithMainCurr = "5,900.00"
                .coinTab.amountCurrForCheck = "5,900.00"
                .coinTab.incomeOutChange = "000931900"
                .coinTab.damagesOutChange = "001434300  "
                .coinTab.roundedAmount = "0.00"
                .coinTab.roundedAmountForCheck = "0.00"
                .attachedTab.addFiles(0) = Project.Path & "Stores\Attach file\excel.xlsx"
                .attachedTab.fileName(0) = "excel.xlsx"
                .attachedTab.addFiles(1) =  Project.Path & "Stores\Attach file\txtFile.txt"
                .attachedTab.fileName(1) = "txtFile.txt"
                .attachedTab.addLinks(0) =  Project.Path & "Stores\Attach file\Photo.jpg"
                .attachedTab.linkName(0) = "attachedLink_1"
      End With
      
      Call Create_Group_Cash_Input(grCashInput, "Î³ï³ñ»É")
      Log.Message("Փաստաթղթի համարը " & grCashInput.generalTab.docNum)
      Log.Message("Փաստաթղթի ISN` " & grCashInput.fIsn)
      
      savePath = Project.Path & "Stores\Cash_Input_Output\Actual\"
      fName = "GroupCashPhysicalPersonAct.txt"
      fileName1 = Project.Path & "Stores\Cash_Input_Output\Actual\GroupCashPhysicalPersonAct.txt"
      fileName2 = Project.Path & "Stores\Cash_Input_Output\Expected\GroupCashPhysicalPersonExp.txt"
      
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
      
      Call Close_Window(wMDIClient, "frmPttel")
      
      ' DOCS
      fBODY = "" & vbCRLF _
            & "ACSBRANCH:00" & vbCRLF _
            & "ACSDEPART:1" & vbCRLF _
            & "BLREP:0" & vbCRLF _
            & "TYPECODE:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28" & vbCRLF _
            & "USERID:  77" & vbCRLF _
            & "DOCNUM:"& grCashInput.generalTab.docNum & vbCRLF _
            & "DATE:"& todayDateSQL & vbCRLF _
            & "KASSA:001" & vbCRLF _
            & "ACCDB:73030121000" & vbCRLF _
            & "CUR:001" & vbCRLF _
            & "KASSIMV:021" & vbCRLF _
            & "BASE:ÊÙµ³ÛÇÝ Ï³ÝËÇÏ Ùáõïù" & vbCRLF _
            & "CLICODE:00034854" & vbCRLF _
            & "PAYER:²ñï³Ï" & vbCRLF _
            & "PAYERLASTNAME:Ð³Ûñ³å»ïÛ³Ý" & vbCRLF _
            & "PASSNUM:AN524685478" & vbCRLF _
            & "PASTYPE:09" & vbCRLF _
            & "PASBY:012" & vbCRLF _
            & "DATEPASS:20200101" & vbCRLF _
            & "DATEEXPIRE:20350101" & vbCRLF _
            & "DATEBIRTH:19910101" & vbCRLF _
            & "CITIZENSHIP:1" & vbCRLF _
            & "COUNTRY:AM" & vbCRLF _
            & "COMMUNITY:010010635" & vbCRLF _
            & "CITY:ºñ¨³Ý" & vbCRLF _
            & "APARTMENT:µÝ. 18" & vbCRLF _
            & "ADDRESS:²µáíÛ³Ý" & vbCRLF _
            & "BUILDNUM:Þ»Ýù" & vbCRLF _
            & "EMAIL:artakhayrapetyan@gmail.com" & vbCRLF _
            & "ACSBRANCHINC:00" & vbCRLF _
            & "ACSDEPARTINC:1" & vbCRLF _
            & "CHRGACC:000001101" & vbCRLF _
            & "TYPECODE2:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28" & vbCRLF _
            & "CHRGCUR:001" & vbCRLF _
            & "CHRGCBCRS:400.0000/1" & vbCRLF _
            & "PAYSCALE:03" & vbCRLF _
            & "CHRGSUM:18" & vbCRLF _
            & "PRSNT:0.3" & vbCRLF _
            & "CHRGINC:000434400" & vbCRLF _
            & "CUPUSA:1" & vbCRLF _
            & "CURTES:1" & vbCRLF _
            & "CURVAIR:3" & vbCRLF _
            & "TIME:"& grCashInput.chargeTab.timeForCheck & vbCRLF _
            & "VOLORT:7" & vbCRLF _
            & "NONREZ:1" & vbCRLF _
            & "JURSTAT:21" & vbCRLF _
            & "COMM:²ñï³ñÅ.ÙÇçí×. ·³ÝÓáõÙ" & vbCRLF _
            & "XSUM:100" & vbCRLF _
            & "XCUR:000" & vbCRLF _
            & "XACC:000001100" & vbCRLF _
            & "XDLCRS:370/1" & vbCRLF _
            & "XDLCRSNAME:000 / 001" & vbCRLF _
            & "XCBCRS:400.0000/1" & vbCRLF _
            & "XCBCRSNAME:000 / 001" & vbCRLF _
            & "XCUPUSA:2" & vbCRLF _
            & "XCURSUM:37000" & vbCRLF _
            & "XSUMMAIN:5900" & vbCRLF _
            & "XINC:000931900" & vbCRLF _
            & "XEXP:001434300"& vbCRLF _
            & ""
      fBODY = Replace(fBODY, "  ", "%")
      Call CheckDB_DOCS(grCashInput.fIsn,"PkCash  ","2",fBODY,1)
      Call CheckQueryRowCount("DOCS","fISN",grCashInput.fIsn,1)
                
	     ' DOCLOG
      Call CheckQueryRowCount("DOCLOG","fISN",grCashInput.fIsn,2)
      Call CheckDB_DOCLOG(grCashInput.fIsn,"77","N","1"," ",1)
      Call CheckDB_DOCLOG(grCashInput.fIsn,"77","C","2"," ",1)
          
    	 ' FOLDERS
      Set dbFOLDERS(0) = New_DB_FOLDERS()
      With dbFOLDERS(0)
          .fFOLDERID = "C.764513596"
          .fNAME = "PkCash"
          .fKEY = grCashInput.fIsn
          .fISN = grCashInput.fIsn
          .fSTATUS = "5"
          .fCOM = "ÊÙµ³ÛÇÝ Ï³ÝËÇÏ Ùáõïù"
          .fSPEC = "²Ùë³ÃÇí- "& todayDate &" N- "& grCashInput.generalTab.docNum &" ¶áõÙ³ñ-             6,000.00 ²ñÅ.- 001 [Üáñ]"
          .fECOM = "Grouped Cash Deposit Advice"
      End With

     	Set dbFOLDERS(1) = New_DB_FOLDERS()
      With dbFOLDERS(1)
          .fFOLDERID = "Oper."&todayDateSQL
          .fNAME = "PkCash"
          .fKEY = grCashInput.fIsn
          .fISN = grCashInput.fIsn
          .fSTATUS = "5"
          .fCOM = "ÊÙµ³ÛÇÝ Ï³ÝËÇÏ Ùáõïù"
          .fSPEC = grCashInput.generalTab.docNum &"7770073030121000                         6000.00001Üáñ                                                   77²ñï³Ï Ð³Ûñ³å»ïÛ³Ý               "&_
                            "AN524685478 012 01/01/2020                                      ÊÙµ³ÛÇÝ Ï³ÝËÇÏ Ùáõïù"
          .fECOM = "Grouped Cash Deposit Advice"
          .fDCDEPART = "1"
          .fDCBRANCH = "00"
      End With
          
     	Call CheckQueryRowCount("FOLDERS","fISN",grCashInput.fIsn,2)
      Call CheckDB_FOLDERS(dbFOLDERS(0),1)
      Call CheckDB_FOLDERS(dbFOLDERS(1),1)

   	  ' DOCSATTACH
   	  Call CheckDB_DOCSATTACH(grCashInput.fIsn, Project.Path & "Stores\Attach file\Photo.jpg", 1, "attachedLink_1", 1)
   	  Call CheckDB_DOCSATTACH(grCashInput.fIsn, "excel.xlsx", 0,"" , 1)
   	  Call CheckDB_DOCSATTACH(grCashInput.fIsn, "txtFile.txt", 0, "", 1)
   	  Call CheckQueryRowCount("DOCSATTACH","fISN",GrCashInput.fIsn,3)

   	  ' DOCSG
   	  Call CheckQueryRowCount("DOCSG","fISN",grCashInput.fIsn,10)

   	 ' HI
      Set dbHI(0) = New_DB_HI()
      With dbHI(0)
            .fBASE = grCashInput.fIsn
            .fDATE = todayDateSQL2
            .fTYPE = "11"
            .fSUM = "360000.00"
            .fCUR = "001"
            .fCURSUM = "900.00"
            .fOP = "MSC"
            .fDBCR = "D"
            .fADB = "1196072159"
            .fACR = "1196072159"
            .fSPEC = grCashInput.generalTab.docNum & "                   Ð³ñÏ»ñÇ Ù³ñáõÙ                    0   400.0000    1"
      End With

	     Set dbHI(1) = New_DB_HI()
      With dbHI(1)
            .fBASE = grCashInput.fIsn
            .fDATE = todayDateSQL2
            .fTYPE = "11"
            .fSUM = "360000.00"
            .fCUR = "001"
            .fCURSUM = "900.00"
            .fOP = "MSC"
            .fDBCR = "C"
            .fADB = "1196072159"
            .fACR = "1196072159"
            .fSPEC = grCashInput.generalTab.docNum & "                   Ð³ñÏ»ñÇ Ù³ñáõÙ                    1   400.0000    1"
      End With

	     Set dbHI(2) = New_DB_HI()
      With dbHI(2)
            .fBASE = grCashInput.fIsn
            .fDATE = todayDateSQL2
            .fTYPE = "11"
            .fSUM = "2000000.00"
            .fCUR = "001"
            .fCURSUM = "5000.00"
            .fOP = "MSC"
            .fDBCR = "D"
            .fADB = "1196072159"
            .fACR = "1426146440"
            .fSPEC = grCashInput.generalTab.docNum & "                   ì³ñÏÇ Ù³ñáõÙ                      0   400.0000    1"
      End With

	     Set dbHI(3) = New_DB_HI()
      With dbHI(3)
            .fBASE = grCashInput.fIsn
            .fDATE = todayDateSQL2
            .fTYPE = "11"
            .fSUM = "2000000.00"
            .fCUR = "001"
            .fCURSUM = "5000.00"
            .fOP = "MSC"
            .fDBCR = "C"
            .fADB = "1196072159"
            .fACR = "1426146440"
            .fSPEC = grCashInput.generalTab.docNum & "                   ì³ñÏÇ Ù³ñáõÙ                      1   400.0000    1"
      End With

	     Set dbHI(4) = New_DB_HI()
      With dbHI(4)
            .fBASE = grCashInput.fIsn
            .fDATE = todayDateSQL2
            .fTYPE = "11"
            .fSUM = "3000.00"
            .fCUR = "000"
            .fCURSUM = "3000.00"
            .fOP = "MSC"
            .fDBCR = "D"
            .fADB = "1629708"
            .fACR = "1196072159"
            .fSPEC = grCashInput.generalTab.docNum & "                   ìÝ³ëÝ»ñ ³ñï. ÷áË³Ý³ÏáõÙÇó         1     1.0000    1"
      End With


	     Set dbHI(5) = New_DB_HI()
      With dbHI(5)
            .fBASE = grCashInput.fIsn
            .fDATE = todayDateSQL2
            .fTYPE = "11"
            .fSUM = "3000.00"
            .fCUR = "001"
            .fCURSUM = "0.00"
            .fOP = "MSC"
            .fDBCR = "C"
            .fADB = "1629708"
            .fACR = "1196072159"
            .fSPEC = grCashInput.generalTab.docNum & "                   ìÝ³ëÝ»ñ ³ñï. ÷áË³Ý³ÏáõÙÇó         0   400.0000    1"
      End With

	     Set dbHI(6) = New_DB_HI()
      With dbHI(6)
            .fBASE = grCashInput.fIsn
            .fDATE = todayDateSQL2
            .fTYPE = "11"
            .fSUM = "37000.00"
            .fCUR = "000"
            .fCURSUM = "37000.00"
            .fOP = "CEX"
            .fDBCR = "D"
            .fADB = "1630170"
            .fACR = "1196072159"
            .fSPEC = grCashInput.generalTab.docNum & "                   Ð³ñÏ»ñÇ Ù³ñáõÙ                    0     1.0000    1"
      End With

	     Set dbHI(7) = New_DB_HI()
      With dbHI(7)
            .fBASE = grCashInput.fIsn
            .fDATE = todayDateSQL2
            .fTYPE = "11"
            .fSUM = "37000.00"
            .fCUR = "001"
            .fCURSUM = "100.00"
            .fOP = "CEX"
            .fDBCR = "C"
            .fADB = "1630170"
            .fACR = "1196072159"
            .fSPEC = grCashInput.generalTab.docNum & "                   Ð³ñÏ»ñÇ Ù³ñáõÙ                    1   370.0000    1"
      End With

	     Set dbHI(8) = New_DB_HI()
      With dbHI(8)
            .fBASE = grCashInput.fIsn
            .fDATE = todayDateSQL2
            .fTYPE = "11"
            .fSUM = "7200.00"
            .fCUR = "000"
            .fCURSUM = "7200.00"
            .fOP = "FEX"
            .fDBCR = "C"
            .fADB = "1630171"
            .fACR = "1629198"
            .fSPEC = grCashInput.generalTab.docNum & "                   ²ñï³ñÅ.ÙÇçí×. ·³ÝÓáõÙ             0     1.0000    1"
      End With

	     Set dbHI(9) = New_DB_HI()
      With dbHI(9)
            .fBASE = grCashInput.fIsn
            .fDATE = todayDateSQL2
            .fTYPE = "11"
            .fSUM = "7200.00"
            .fCUR = "001"
            .fCURSUM = "18.00"
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
  	   Call Check_DB_HI(dbHI(6),1)
  	   Call Check_DB_HI(dbHI(7),1)
  	   Call Check_DB_HI(dbHI(8),1)
  	   Call Check_DB_HI(dbHI(9),1)
     	Call CheckQueryRowCount("HI","fBASE",grCashInput.fIsn,10)
      

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
      Set editGrCashInp = New_GroupCashInput(2,1, 0, 1)
      With editGrCashInp
                .generalTab.payerLegalStatus = "ֆիզԱնձ"
                .generalTab.wOffice = "00"
                .generalTab.wDepartment = "1"
                .generalTab.wDate =  todayDate 
                .generalTab.cashRegister = "001"
                .generalTab.cashRegisterAcc =  "73030121000" 
                .generalTab.wCurr = "001"
                .generalTab.cashierChar = "021"
                .generalTab.wBase = "ÊÙµ³ÛÇÝ Ï³ÝËÇÏ Ùáõïù"
                .generalTab.wAcc(0)  = "73030121000" 
                .generalTab.wSum(0) = "1,000.00"
                .generalTab.wAim(0) = "Ð³ñÏ»ñÇ Ù³ñáõÙ"
                .generalTab.wAcc(1) = "73030381000" 
                .generalTab.wSum(1) = "5,000.00"
                .generalTab.wAim(1) = "ì³ñÏÇ Ù³ñáõÙ"
                .generalTab.wPayer = "00034854"
                .generalTab.wName = "²ñï³Ï"
                .generalTab.surName = "Ð³Ûñ³å»ïÛ³Ý"
                .generalTab.wIdCheck = "AN524685478"
                .generalTab.wId = "AN524685412"
                .generalTab.idType = "08"
                .generalTab.idTypeCheck = "09"
                .generalTab.idGivenBy = "013"
                .generalTab.idGivenByCheck = "012"
                .generalTab.wCitizenship = "1"
                .generalTab.wCountry = "AM"
                .generalTab.wResidence = "010010635"
                .generalTab.wCity = "ºñ¨³Ý"
                .generalTab.wApartment = "µÝ. 18 "
                .generalTab.wStreet = "²µáíÛ³Ý"
                .generalTab.wHouse = "Þ»Ýù"
                .generalTab.wEmailCheck = "artakhayrapetyan@gmail.com"
                .generalTab.wEmail = "artak@gmail.com"
                .generalTab.birthDateForCheck = "01/01/1991"
                .generalTab.idGiveDateForCheck = "01/01/2020"
                .generalTab.idValidUntilForCheck = "01/01/2035"
                .generalTab.wBirthDate = "01/01/1991"
                .generalTab.idGiveDate = "01/01/2021"
                .generalTab.idValidUntil = "01/01/2036"
                .chargeTab.office = "00"
                .chargeTab.department = "1"
                .chargeTab.chargeAcc = "000001101"   
                .chargeTab.chargeAccForCheck = "000001101"
                .chargeTab.chargeCurr = "001"
                .chargeTab.chargeCurrForCheck = "001"   
                .chargeTab.cbExchangeRate = "400.0000/1"
                .chargeTab.chargeType = "03"
                .chargeTab.chargeAmount = "23.00"
                .chargeTab.chargeAmoForCheck = "48.00"
                .chargeTab.chargePercent = "0.8000"
                .chargeTab.chargePerForCheck = "0.8000"
                .chargeTab.incomeAcc = "000436900"    
                .chargeTab.incomeAccCurr = "000"
                .chargeTab.buyAndSell = "1"
                .chargeTab.buyAndSellForCheck = "1"
                .chargeTab.operType = "1"
                .chargeTab.operPlace = "3"
                .chargeTab.operArea = "7"
                .chargeTab.operAreaForCheck = "7"
                .chargeTab.nonResident = 1
                .chargeTab.nonResidentForCheck = 1
                .chargeTab.legalStatus = "21"
                .chargeTab.legalStatusForCheck = "21"
                .chargeTab.comment = "²ñï³ñÅ.ÙÇçí×³ñ³ÛÇÝ ·³ÝÓáõÙ"
                .chargeTab.commentForCheck = "²ñï³ñÅ.ÙÇçí×. ·³ÝÓáõÙ"
                .chargeTab.clientAgreeData = ""
                .chargeTab.notGrCash = False
                .coinTab.coin = "120.00"
                .coinTab.coinForCheck = "120.00"
                .coinTab.coinPayCurr = "000"
                .coinTab.coinPayAcc = "000001100"
                .coinTab.coinExchangeRate = "370.0000/1"
                .coinTab.coinCBExchangeRate = "400.0000/1"
                .coinTab.coinBuyAndSell = "2"
                .coinTab.coinPayAmount = "44,400.00"
                .coinTab.coinPayAmountForCheck = "44,400.00"
                .coinTab.amountWithMainCurr = "5,880.00"
                .coinTab.amountCurrForCheck = "5,880.00"
                .coinTab.incomeOutChange = "000931900"
                .coinTab.damagesOutChange = "001434300  "
                .coinTab.roundedAmount = "0.00"
                .coinTab.roundedAmountForCheck = "0.00"
                .attachedTab.addFiles(0) = Project.Path & "Stores\Attach file\Photo.jpg"
                .attachedTab.fileName(0) = "Photo.jpg"
                .attachedTab.delFiles(0) = "txtFile.txt"
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
                & "DATE:"& todayDateSQL & vbCRLF _
                & "KASSA:001"& vbCRLF _
                & "ACCDB:73030121000"& vbCRLF _
                & "CUR:001"& vbCRLF _
                & "KASSIMV:021"& vbCRLF _
                & "BASE:ÊÙµ³ÛÇÝ Ï³ÝËÇÏ Ùáõïù"& vbCRLF _
                & "CLICODE:00034854"& vbCRLF _
                & "PAYER:²ñï³Ï"& vbCRLF _
                & "PAYERLASTNAME:Ð³Ûñ³å»ïÛ³Ý"& vbCRLF _
                & "PASSNUM:AN524685412"& vbCRLF _
                & "PASTYPE:08"& vbCRLF _
                & "PASBY:013"& vbCRLF _
                & "DATEPASS:20210101"& vbCRLF _
                & "DATEEXPIRE:20360101"& vbCRLF _
                & "DATEBIRTH:19910101"& vbCRLF _
                & "CITIZENSHIP:1"& vbCRLF _
                & "COUNTRY:AM"& vbCRLF _
                & "COMMUNITY:010010635"& vbCRLF _
                & "CITY:ºñ¨³Ý"& vbCRLF _
                & "APARTMENT:µÝ. 18"& vbCRLF _
                & "ADDRESS:²µáíÛ³Ý"& vbCRLF _
                & "BUILDNUM:Þ»Ýù"& vbCRLF _
                & "EMAIL:artak@gmail.com"& vbCRLF _
                & "ACSBRANCHINC:00"& vbCRLF _
                & "ACSDEPARTINC:1"& vbCRLF _
                & "CHRGACC:000001101"& vbCRLF _
                & "TYPECODE2:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28"& vbCRLF _
                & "CHRGCUR:001"& vbCRLF _
                & "CHRGCBCRS:400.0000/1"& vbCRLF _
                & "PAYSCALE:03"& vbCRLF _
                & "CHRGSUM:48"& vbCRLF _
                & "PRSNT:0.8"& vbCRLF _
                & "CHRGINC:000436900"& vbCRLF _
                & "CUPUSA:1"& vbCRLF _
                & "CURTES:1"& vbCRLF _
                & "CURVAIR:3"& vbCRLF _
                & "TIME:"& grCashInput.chargeTab.timeForCheck & vbCRLF _
                & "VOLORT:7"& vbCRLF _
                & "NONREZ:1"& vbCRLF _
                & "JURSTAT:21"& vbCRLF _
                & "COMM:²ñï³ñÅ.ÙÇçí×³ñ³ÛÇÝ ·³ÝÓáõÙ"& vbCRLF _
                & "XSUM:120"& vbCRLF _
                & "XCUR:000"& vbCRLF _
                & "XACC:000001100"& vbCRLF _
                & "XDLCRS:370/1"& vbCRLF _
                & "XDLCRSNAME:000 / 001"& vbCRLF _
                & "XCBCRS:400.0000/1"& vbCRLF _
                & "XCBCRSNAME:000 / 001"& vbCRLF _
                & "XCUPUSA:2"& vbCRLF _
                & "XCURSUM:44400"& vbCRLF _
                & "XSUMMAIN:5880"& vbCRLF _
                & "XINC:000931900"& vbCRLF _
                & "XEXP:001434300"& vbCRLF _
                & ""
      fBODY = Replace(fBODY, "  ", "%")
      Call CheckDB_DOCS(grCashInput.fIsn,"PkCash  ","2",fBODY,1)
      Call CheckQueryRowCount("DOCS","fISN",grCashInput.fIsn,1)

	     ' DOCLOG
      Call CheckQueryRowCount("DOCLOG","fISN",grCashInput.fIsn,3)
      Call CheckDB_DOCLOG(grCashInput.fIsn,"77","E","2"," ",1)
          

    	 ' FOLDERS
      With dbFOLDERS(0)
          .fFOLDERID = "C.764513596"
          .fNAME = "PkCash  "
          .fKEY = grCashInput.fIsn
          .fISN = grCashInput.fIsn
          .fSTATUS = "5"
          .fCOM = "ÊÙµ³ÛÇÝ Ï³ÝËÇÏ Ùáõïù"
          .fSPEC = "²Ùë³ÃÇí- "&todayDate&" N- "& grCashInput.generalTab.docNum &" ¶áõÙ³ñ-             6,000.00 ²ñÅ.- 001 [ÊÙµ³·ñíáÕ]"
          .fECOM = "Grouped Cash Deposit Advice"
      End With

      With dbFOLDERS(1)
          .fFOLDERID = "Oper."&todayDateSQL
          .fNAME = "PkCash  "
          .fKEY = grCashInput.fIsn
          .fISN = grCashInput.fIsn
          .fSTATUS = "5"
          .fCOM = "ÊÙµ³ÛÇÝ Ï³ÝËÇÏ Ùáõïù"
          .fSPEC = grCashInput.generalTab.docNum &"7770073030121000                         6000.00001ÊÙµ³·ñíáÕ                                             "&_
                   "77²ñï³Ï Ð³Ûñ³å»ïÛ³Ý               AN524685412 013 01/01/2021                                      ÊÙµ³ÛÇÝ Ï³ÝËÇÏ Ùáõïù"
          .fECOM = "Grouped Cash Deposit Advice"
          .fDCDEPART = "1"
          .fDCBRANCH = "00"
      End With
      
     	Call CheckQueryRowCount("FOLDERS","fISN",grCashInput.fIsn,2)
      Call CheckDB_FOLDERS(dbFOLDERS(0),1)
      Call CheckDB_FOLDERS(dbFOLDERS(1),1)

    	  'DOCSATTACH 
       Call CheckDB_DOCSATTACH(grCashInput.fIsn, Project.Path & "Stores\Attach file\Photo.jpg", 1, "attachedLink_1", 1)
    	 Call CheckDB_DOCSATTACH(grCashInput.fIsn, "Photo.jpg", 0, "", 1)
    	 Call CheckDB_DOCSATTACH(grCashInput.fIsn, "excel.xlsx", 0, "", 1)
    	 Call CheckQueryRowCount("DOCSATTACH","fISN",grCashInput.fIsn,3)
      
      Log.Message("Խմբային Կանխիկ մուտք փաստաթուղթն ուղարկել հաստատման")
      colN = 2
      action = c_SendToVer
      doNum = 2
      doActio = "Î³ï³ñ»É"
      If Not ConfirmContractDoc(colN, grCashInput.generalTab.docNum, action, doNum, doActio) Then
            Log.Error("Խմբային կանխիկ մուտք փաստաթուղթը չի ուղարկվել հաստատման")
            Exit Sub
      End If
      
      Call Close_Window(wMDIClient, "frmPttel")
      
	     ' DOCS
      fBODY = "" & vbCRLF _
                & "ACSBRANCH:00"& vbCRLF _
                & "ACSDEPART:1"& vbCRLF _
                & "BLREP:0"& vbCRLF _
                & "TYPECODE:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28"& vbCRLF _
                & "USERID:  77"& vbCRLF _
                & "DOCNUM:"& grCashInput.generalTab.docNum & vbCRLF _
                & "DATE:"& todayDateSQL & vbCRLF _
                & "KASSA:001"& vbCRLF _
                & "ACCDB:73030121000"& vbCRLF _
                & "CUR:001"& vbCRLF _
                & "ISTLLCREATED:1"& vbCRLF _
                & "KASSIMV:021"& vbCRLF _
                & "BASE:ÊÙµ³ÛÇÝ Ï³ÝËÇÏ Ùáõïù"& vbCRLF _
                & "CLICODE:00034854"& vbCRLF _
                & "PAYER:²ñï³Ï"& vbCRLF _
                & "PAYERLASTNAME:Ð³Ûñ³å»ïÛ³Ý"& vbCRLF _
                & "PASSNUM:AN524685412"& vbCRLF _
                & "PASTYPE:08"& vbCRLF _
                & "PASBY:013"& vbCRLF _
                & "DATEPASS:20210101"& vbCRLF _
                & "DATEEXPIRE:20360101"& vbCRLF _
                & "DATEBIRTH:19910101"& vbCRLF _
                & "CITIZENSHIP:1"& vbCRLF _
                & "COUNTRY:AM"& vbCRLF _
                & "COMMUNITY:010010635"& vbCRLF _
                & "CITY:ºñ¨³Ý"& vbCRLF _
                & "APARTMENT:µÝ. 18"& vbCRLF _
                & "ADDRESS:²µáíÛ³Ý"& vbCRLF _
                & "BUILDNUM:Þ»Ýù"& vbCRLF _
                & "EMAIL:artak@gmail.com"& vbCRLF _
                & "ACSBRANCHINC:00"& vbCRLF _
                & "ACSDEPARTINC:1"& vbCRLF _
                & "CHRGACC:000001101"& vbCRLF _
                & "TYPECODE2:-20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28"& vbCRLF _
                & "CHRGCUR:001"& vbCRLF _
                & "CHRGCBCRS:400.0000/1"& vbCRLF _
                & "PAYSCALE:03"& vbCRLF _
                & "CHRGSUM:48"& vbCRLF _
                & "PRSNT:0.8"& vbCRLF _
                & "CHRGINC:000436900"& vbCRLF _
                & "CUPUSA:1"& vbCRLF _
                & "CURTES:1"& vbCRLF _
                & "CURVAIR:3"& vbCRLF _
                & "TIME:"& grCashInput.chargeTab.timeForCheck & vbCRLF _
                & "VOLORT:7"& vbCRLF _
                & "NONREZ:1"& vbCRLF _
                & "JURSTAT:21"& vbCRLF _
                & "COMM:²ñï³ñÅ.ÙÇçí×³ñ³ÛÇÝ ·³ÝÓáõÙ"& vbCRLF _
                & "XSUM:120"& vbCRLF _
                & "XCUR:000"& vbCRLF _
                & "XACC:000001100"& vbCRLF _
                & "XDLCRS:370/1"& vbCRLF _
                & "XDLCRSNAME:000 / 001"& vbCRLF _
                & "XCBCRS:400.0000/1"& vbCRLF _
                & "XCBCRSNAME:000 / 001"& vbCRLF _
                & "XCUPUSA:2"& vbCRLF _
                & "XCURSUM:44400"& vbCRLF _
                & "XSUMMAIN:5880"& vbCRLF _
                & "XINC:000931900"& vbCRLF _
                & "XEXP:001434300"& vbCRLF _
                & ""
      Call CheckDB_DOCS(GrCashInput.fIsn,"PkCash  ","101",fBODY,1)
      Call CheckQueryRowCount("DOCS","fISN",grCashInput.fIsn,1)

	     ' DOCLOG
      Call CheckQueryRowCount("DOCLOG","fISN",GrCashInput.fIsn,4)
	  Call CheckDB_DOCLOG(grCashInput.fIsn,"77","M","101","àõÕ³ñÏí»É ¿ Ñ³ëï³ïÙ³Ý",1)

    	 ' FOLDERS
      With dbFOLDERS(0)
          .fFOLDERID = "C.764513596"
          .fNAME = "PkCash  "
          .fKEY = grCashInput.fIsn
          .fISN = grCashInput.fIsn
          .fSTATUS = "0"
          .fCOM = "ÊÙµ³ÛÇÝ Ï³ÝËÇÏ Ùáõïù"
          .fSPEC = "²Ùë³ÃÇí- "&todayDate&" N- "& grCashInput.generalTab.docNum &" ¶áõÙ³ñ-             6,000.00 ²ñÅ.- 001 [àõÕ³ñÏí³Í I Ñ³ëï³ïÙ³Ý]"
          .fECOM = "Grouped Cash Deposit Advice"
      End With

      With dbFOLDERS(1)
          .fFOLDERID = "Oper."&todayDateSQL
          .fNAME = "PkCash  "
          .fKEY = grCashInput.fIsn
          .fISN = grCashInput.fIsn
          .fSTATUS = "0"
          .fCOM = "ÊÙµ³ÛÇÝ Ï³ÝËÇÏ Ùáõïù"
          .fSPEC = grCashInput.generalTab.docNum &"7770073030121000                         6000.00001àõÕ³ñÏí³Í I Ñ³ëï³ïÙ³Ý                                 "&_
                            "77²ñï³Ï Ð³Ûñ³å»ïÛ³Ý                                                                               ÊÙµ³ÛÇÝ Ï³ÝËÇÏ Ùáõïù"
          .fECOM = "Grouped Cash Deposit Advice"
          .fDCDEPART = "1"
          .fDCBRANCH = "00"
      End With

	     Set dbFOLDERS(2) = New_DB_FOLDERS()
      With dbFOLDERS(2)
          .fFOLDERID = "Ver."&todayDateSQL&"001"
          .fNAME = "PkCash  "
          .fKEY = grCashInput.fIsn
          .fISN = grCashInput.fIsn
          .fSTATUS = "4"
          .fCOM = "ÊÙµ³ÛÇÝ Ï³ÝËÇÏ Ùáõïù"
          .fSPEC = grCashInput.generalTab.docNum &"7770073030121000                         6000.00001  77                                ÊÙµ³ÛÇÝ Ï³ÝËÇÏ Ùáõïù            ²ñï³Ï Ð³Ûñ³å»ïÛ³Ý"
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

      If Not wMDIClient.VBObject("frmPttel").Exists Then
            Log.Error("Հաստատվող փաստաթղթեր թղթապանակը չի բացվել")
            Exit Sub
      End If
      
      Log.Message("Վավերացնել Խմբային կանխիկ մուտքի փաստաթուղթը - VERIFIER")
      colN = 3
      action = c_ToConfirm
      doNum = 1
      doActio = "Ð³ëï³ï»É"

      If Not ConfirmContractDoc(colN, grCashInput.generalTab.docNum, action, doNum, doActio) Then
            Log.Error("Խմբային կանխիկ մուտքի փաստաթուղթը չի վավերացվել")
            Exit Sub
      End If
      
      Call Close_Window(wMDIClient, "frmPttel")

	     ' DOCS
      fBODY = Replace(fBODY, "  ", "%")
      Call CheckDB_DOCS(grCashInput.fIsn,"PkCash  ","15",fBODY,1)
      Call CheckQueryRowCount("DOCS","fISN",grCashInput.fIsn,1)

	     ' DOCLOG
      Call CheckQueryRowCount("DOCLOG","fISN",grCashInput.fIsn,6)
      Call CheckDB_DOCLOG(grCashInput.fIsn,"81","W","102"," ",1)
	     Call CheckDB_DOCLOG(grCashInput.fIsn,"81","C","15"," ",1)
          
    	 ' FOLDERS
      With dbFOLDERS(0)
          .fFOLDERID = "C.764513596"
          .fNAME = "PkCash  "
          .fKEY = grCashInput.fIsn
          .fISN = grCashInput.fIsn
          .fSTATUS = "4"
          .fCOM = "ÊÙµ³ÛÇÝ Ï³ÝËÇÏ Ùáõïù"
          .fSPEC = "²Ùë³ÃÇí- "&todayDate&" N- "& grCashInput.generalTab.docNum &" ¶áõÙ³ñ-             6,000.00 ²ñÅ.- 001 [Ð³ëï³ïí³Í]"
          .fECOM = "Grouped Cash Deposit Advice"
      End With

      With dbFOLDERS(1)
          .fFOLDERID = "Oper."&todayDateSQL
          .fNAME = "PkCash  "
          .fKEY = grCashInput.fIsn
          .fISN = grCashInput.fIsn
          .fSTATUS = "4"
          .fCOM = "ÊÙµ³ÛÇÝ Ï³ÝËÇÏ Ùáõïù"
          .fSPEC = grCashInput.generalTab.docNum &"7770073030121000                         6000.00001Ð³ëï³ïí³Í                                             77²ñï³Ï Ð³Ûñ³å»ïÛ³Ý               "&_
                            "AN524685412 013 01/01/2021                                      ÊÙµ³ÛÇÝ Ï³ÝËÇÏ Ùáõïù"
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
            Exit Sub
      End If
      
      Call Close_Window(wMDIClient, "frmPttel")
      
      ' DOCS
      fBODY = Replace(fBODY, "  ", "%")
      Call CheckDB_DOCS(grCashInput.fIsn,"PkCash  ","11",fBODY,1)
      Call CheckQueryRowCount("DOCS","fISN",grCashInput.fIsn,1)

	     ' DOCLOG
      Call CheckQueryRowCount("DOCLOG","fISN",grCashInput.fIsn,8)
      Call CheckDB_DOCLOG(grCashInput.fIsn,"77","W","16"," ",1)
	     Call CheckDB_DOCLOG(grCashInput.fIsn,"77","M","11","¶ñ³Ýóí»É »Ý Ó¨³Ï»ñåáõÙÝ»ñÁ",1)
          
      'FOLDERS
      Call CheckQueryRowCount("FOLDERS","fISN",grCashInput.fIsn,0)
      
      'HI
      With dbHI(0)
            .fBASE = grCashInput.fIsn
            .fDATE = todayDateSQL2
            .fTYPE = "01"
            .fSUM = "352000.00"
            .fCUR = "001"
            .fCURSUM = "880.00"
            .fOP = "MSC"
            .fDBCR = "D"
            .fADB = "1196072159"
            .fACR = "1196072159"
            .fSPEC = grCashInput.generalTab.docNum & "021                Ð³ñÏ»ñÇ Ù³ñáõÙ                    0   400.0000    1"
            .fBASEBRANCH = "00"
            .fBASEDEPART = "1"
      End With

      With dbHI(1)
            .fBASE = grCashInput.fIsn
            .fDATE = todayDateSQL2
            .fTYPE = "01"
            .fSUM = "352000.00"
            .fCUR = "001"
            .fCURSUM = "880.00"
            .fOP = "MSC"
            .fDBCR = "C"
            .fADB = "1196072159"
            .fACR = "1196072159"
            .fSPEC = grCashInput.generalTab.docNum & "                   Ð³ñÏ»ñÇ Ù³ñáõÙ                    1   400.0000    1                                                                        ²ñï³Ï Ð³Ûñ³å»ïÛ³Ý               "
            .fBASEBRANCH = "00"
            .fBASEDEPART = "1"
      End With

      With dbHI(2)
            .fBASE = grCashInput.fIsn
            .fDATE = todayDateSQL2
            .fTYPE = "01"
            .fSUM = "2000000.00"
            .fCUR = "001"
            .fCURSUM = "5000.00"
            .fOP = "MSC"
            .fDBCR = "D"
            .fADB = "1196072159"
            .fACR = "1426146440"
            .fSPEC = grCashInput.generalTab.docNum & "021                ì³ñÏÇ Ù³ñáõÙ                      0   400.0000    1"
            .fBASEBRANCH = "00"
            .fBASEDEPART = "1"
      End With

      With dbHI(3)
            .fBASE = grCashInput.fIsn
            .fDATE = todayDateSQL2
            .fTYPE = "01"
            .fSUM = "2000000.00"
            .fCUR = "001"
            .fCURSUM = "5000.00"
            .fOP = "MSC"
            .fDBCR = "C"
            .fADB = "1196072159"
            .fACR = "1426146440"
            .fSPEC = grCashInput.generalTab.docNum & "                   ì³ñÏÇ Ù³ñáõÙ                      1   400.0000    1                                                                        ²ñï³Ï Ð³Ûñ³å»ïÛ³Ý               "
            .fBASEBRANCH = "00"
            .fBASEDEPART = "1"
      End With

      With dbHI(4)
            .fBASE = grCashInput.fIsn
            .fDATE = todayDateSQL2
            .fTYPE = "01"
            .fSUM = "3600.00"
            .fCUR = "000"
            .fCURSUM = "3600.00"
            .fOP = "MSC"
            .fDBCR = "D"
            .fADB = "1629708"
            .fACR = "1196072159"
            .fSPEC = grCashInput.generalTab.docNum & "                   ìÝ³ëÝ»ñ ³ñï. ÷áË³Ý³ÏáõÙÇó         1     1.0000    1"
            .fBASEBRANCH = "00"
            .fBASEDEPART = "1"
      End With

      With dbHI(5)
            .fBASE = grCashInput.fIsn
            .fDATE = todayDateSQL2
            .fTYPE = "01"
            .fSUM = "3600.00"
            .fCUR = "001"
            .fCURSUM = "0.00"
            .fOP = "MSC"
            .fDBCR = "C"
            .fADB = "1629708"
            .fACR = "1196072159"
            .fSPEC = grCashInput.generalTab.docNum & "                   ìÝ³ëÝ»ñ ³ñï. ÷áË³Ý³ÏáõÙÇó         0   400.0000    1"
            .fBASEBRANCH = "00"
            .fBASEDEPART = "1"
      End With

      With dbHI(6)
            .fBASE = grCashInput.fIsn
            .fDATE = todayDateSQL2
            .fTYPE = "01"
            .fSUM = "44400.00"
            .fCUR = "000"
            .fCURSUM = "44400.00"
            .fOP = "CEX"
            .fDBCR = "D"
            .fADB = "1630170"
            .fACR = "1196072159"
            .fSPEC = grCashInput.generalTab.docNum & "021                Ð³ñÏ»ñÇ Ù³ñáõÙ                    0     1.0000    1"
            .fBASEBRANCH = "00"
            .fBASEDEPART = "1"
      End With

      With dbHI(7)
            .fBASE = grCashInput.fIsn
            .fDATE = todayDateSQL2
            .fTYPE = "01"
            .fSUM = "44400.00"
            .fCUR = "001"
            .fCURSUM = "120.00"
            .fOP = "CEX"
            .fDBCR = "C"
            .fADB = "1630170"
            .fACR = "1196072159"
            .fSPEC = grCashInput.generalTab.docNum & "                   Ð³ñÏ»ñÇ Ù³ñáõÙ                    1   370.0000    1"
            .fBASEBRANCH = "00"
            .fBASEDEPART = "1"
      End With

      With dbHI(8)
            .fBASE = grCashInput.fIsn
            .fDATE = todayDateSQL2
            .fTYPE = "CE"
            .fSUM = "120.00"
            .fCUR = "000"
            .fCURSUM = "44400.00"
            .fOP = "SAL"
            .fDBCR = "D"
            .fADB = "-1"
            .fACR = "-1"
            .fSPEC =  "%"& grCashInput.generalTab.docNum & "7 "
            .fBASEBRANCH = "00"
            .fBASEDEPART = "1"
      End With

      With dbHI(9)
            .fBASE = grCashInput.fIsn
            .fDATE = todayDateSQL2
            .fTYPE = "01"
            .fSUM = "19200.00"
            .fCUR = "000"
            .fCURSUM = "19200.00"
            .fOP = "FEX"
            .fDBCR = "C"
            .fADB = "1630171"
            .fACR = "1629200"
            .fSPEC = grCashInput.generalTab.docNum & "                   ²ñï³ñÅ.ÙÇçí×³ñ³ÛÇÝ ·³ÝÓáõÙ        0     1.0000    1"
            .fBASEBRANCH = "00"
            .fBASEDEPART = "1"
      End With

	     Set dbHI(10) = New_DB_HI()
      With dbHI(10)
            .fBASE = grCashInput.fIsn
            .fDATE = todayDateSQL2
            .fTYPE = "01"
            .fSUM = "19200.00"
            .fCUR = "001"
            .fCURSUM = "48.00"
            .fOP = "FEX"
            .fDBCR = "D"
            .fADB = "1630171"
            .fACR = "1629200"
            .fSPEC = grCashInput.generalTab.docNum & "021                ²ñï³ñÅ.ÙÇçí×³ñ³ÛÇÝ ·³ÝÓáõÙ        1   400.0000    1"
            .fBASEBRANCH = "00"
            .fBASEDEPART = "1"
      End With

	     Set dbHI(11) = New_DB_HI()
      With dbHI(11)
            .fBASE = grCashInput.fIsn
            .fDATE = todayDateSQL2
            .fTYPE = "CE"
            .fSUM = "19200.00"
            .fCUR = "001"
            .fCURSUM = "48.00"
            .fOP = "PUR"
            .fDBCR = "D"
            .fADB = "-1"
            .fACR = "-1"
            .fSPEC = "%"& grCashInput.generalTab.docNum &"7 "
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
  	   Call Check_DB_HI(dbHI(7),1)
  	   Call Check_DB_HI(dbHI(8),1)
  	   Call Check_DB_HI(dbHI(9),1)
  	   Call Check_DB_HI(dbHI(10),1)
  	   Call Check_DB_HI(dbHI(11),1)
      Call CheckQueryRowCount("HI","fBASE",grCashInput.fIsn,12)

     	Set dbPAYMENTS(0) = New_DB_PAYMENTS()
      With dbPAYMENTS(0)
            .fISN = grCashInput.fIsn
            .fDOCTYPE = "PkCash"
            .fDATE = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%Y-%m-%d")
            .fSTATE = "11"
            .fDOCNUM = grCashInput.generalTab.docNum
            .fCLIENT = "00034854"
            .fACCDB = "7770073030121000"
            .fPAYER = "²ñï³Ï Ð³Ûñ³å»ïÛ³Ý"
            .fCUR = "001"
            .fSUMMA = "6000.00"
            .fSUMMAAMD = "2400000.00"
            .fSUMMAUSD = "6000.00"
            .fCOM = "ÊÙµ³ÛÇÝ Ï³ÝËÇÏ Ùáõïù                                                                                                                        "
            .fPASSPORT = "AN524685412 013 01/01/2021"
            .fCOUNTRY = "AM"
            .fACSBRANCH = "00"
            .fACSDEPART = "1"
      End With
      Call CheckDB_PAYMENTS(dbPAYMENTS(0),1)
      
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

      Call Close_Window(wMDIClient, "frmPttel")
      
      ' FOLDERS
       With dbFOLDERS(0)
          .fFOLDERID = ".R."&todayDateSQL
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
      Call CheckQueryRowCount("DOCLOG","fISN",grCashInput.fIsn,9)
      
      ' Փակել ծրագիրը
      Call Close_AsBank()
      
End Sub