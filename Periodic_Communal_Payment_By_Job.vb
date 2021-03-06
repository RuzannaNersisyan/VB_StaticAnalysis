 Option Explicit
'USEUNIT Library_Common
'USEUNIT Constants
'USEUNIT Subsystems_SQL_Library
'USEUNIT Library_Colour
'USEUNIT Library_Periodic_Actions
'USEUNIT Subsystems_Special_Library
'USEUNIT BankMail_Library
'USEUNIT Payment_Except_Library
'USEUNIT SWIFT_International_Payorder_Library

'Test case ID 171090
   
' Պարբերական վճարումներ գործողության կատարում առաջադրանքի միջոցով
Sub Periodic_Communal_Payment_By_Job_Test()     
     
      Dim DateStart, DateEnd, isExists, param, fileName2, fileName1, savePath
      Dim folderName, communalPay, workingDocs, communalPayDoc, addJob, viewPayment
      Dim frmPttel, status, dateTimeNow, commISN, getJobISN, jobISNColN
      Dim queryString, sqlValue, colNum, sql_isEqual, commDocNum
      
      DateStart = "20120101"
      DateEnd = "20240101"
      Call Initialize_AsBankQA(DateStart, DateEnd) 
      Call Create_Connection()
      
      ' Մուտք գործել համակարգ ARMSOFT օգտագործողով 
      Login("ARMSOFT")
      dateTimeNow = aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%d%m%y")
     
      ' Պարբերական կոմունալ վճարումների հաճախորդին տեղեկացումների արտահանման ճանապարհ
      Call  SetParameter("PCPPATH", "\\host2\Sys\Testing\Comunal\Out")
      folderName = "|ä³ñµ»ñ³Ï³Ý ·áñÍáÕáõÃÛáõÝÝ»ñÇ ²Þî|"
      Call ChangeWorkspace(c_PeriodicActions)
      
      Set communalPay = New_CommunalPayment(2)
      With communalPay
      .general.office = "P00"
			.general.department = "08"
			.general.client = "00002248"
			.general.maxPrice = "50000"      
      .general.services(0).Num = "1"
			.general.services(0).service = "WN"
			.general.services(0).place = "01"
			.general.services(0).clientN = "7-62-0-0-101"
			.general.services(0).minPrice = ""
			.general.services(0).maxPrice = "2000" 
      .general.services(1).Num = "2"
			.general.services(1).service = "R"
			.general.services(1).place = "9"
			.general.services(1).clientN = "12402"
			.general.services(1).minPrice = ""
			.general.services(1).maxPrice = "7000"       
      .other.openDate = "010821"
			.other.lastDate = "010122"
      .other.payDays = "12"
      .other.payDays_to = "28"
			.other.informClient = 1
			.other.useClientEmail = 0
			.other.otherEmail = "sona.gyulamiryan@armsoft.am"
			.other.accsConnentScheme = "001"
			.other.useClientScheme = 1
			.other.useCardAccs = 0
			.other.addInfo = "For Test"
			.other.lastOpersDate = ""
			.other.lastCompletedDate = ""
			.other.closeDate = ""
      End With
      
      Call Create_CommunalPayment(folderName, communalPay)
      
      Log.Message(communalPay.fisn)
      Log.Message(communalPay.docNum)
      
                'DOCLOG
                queryString = " SELECT COUNT(*) FROM DOCLOG WHERE fISN = " & communalPay.fisn & _
                                         " and fSUID = '10' and fOP = 'N' and fSTATE = '1' and fSUIDCOR = '-1'"
                sqlValue = 1
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If 
                
                'FOLDERS
                queryString = " SELECT COUNT(*) FROM FOLDERS WHERE fISN = " & communalPay.fisn & _
                                          " and  fNAME = 'PCPAGR' " & _ 
                                          " and ((fCOM = 'ä³ñµ»ñ³Ï³Ý ÏáÙáõÝ³É í×³ñáõÙÝ»ñÇ å³ÛÙ³Ý³·Çñ' and fSPEC = '²Ùë³ÃÇí- 01/08/21 N- "&communalPay.docNum&" [Üáñ]' and fECOM = 'Periodic communal payments agreement')" & _ 
                                          " or (fCOM = 'ä³ñµ»ñ³Ï³Ý ÏáÙáõÝ³É í×³ñáõÙÝ»ñÇ å³ÛÙ³Ý³·Çñ' and fSPEC = '"&communalPay.docNum&"1660000224820100                            0.00000Üáñ                                                   10Ð³×³Ëáñ¹ 00002248                                                                               ä³ñµ»ñ³Ï³Ý ÏáÙáõÝ³É í×³ñáõÙÝ»ñÇ å³ÛÙ³Ý³·Çñ' and fECOM = 'Periodic communal payments agreement')" & _ 
                                          " or (fCOM = 'Ð³×³Ëáñ¹ 00002248' and fSPEC = '1   000022482021080120220101122810          50000.00100022482010000000011000000000000000000000000sona.gyulamiryan@armsoft.am' and fECOM = 'Client 00002248'))"
                sqlValue = 3
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If 
                
                'DOCSG
                queryString = " SELECT COUNT(*) FROM DOCSG WHERE fISN = " & communalPay.fisn
                sqlValue = 18
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If 
                
                'DOCS
                queryString = " SELECT COUNT(*) FROM DOCS WHERE fISN = " & communalPay.fisn & _
                                         " and fNAME = 'PCPAGR' and fSTATE = '1' and fNEXTTRANS = '1' and fBODY = '"& vbCRLF _ 
                                        & "ACSBRANCH:P00" & vbCRLF _ 
                                        & "ACSDEPART:08" & vbCRLF _ 
                                        & "USERID:10" & vbCRLF _ 
                                        & "CODE:"& communalPay.docNum & vbCRLF _ 
                                        & "CLICODE:00002248" & vbCRLF _ 
                                        & "NAME:Ð³×³Ëáñ¹ 00002248" & vbCRLF _ 
                                        & "ENAME:Client 00002248" & vbCRLF _ 
                                        & "FEEACC:00224820100" & vbCRLF _ 
                                        & "FEECUR:000" & vbCRLF _ 
                                        & "MAXSUM:50000" & vbCRLF _ 
                                        & "SDATE:20210801" & vbCRLF _ 
                                        & "EDATE:20220101" & vbCRLF _ 
                                        & "SDAY:12" & vbCRLF _ 
                                        & "LDAY:28" & vbCRLF _ 
                                        & "CLINOT:1" & vbCRLF _ 
                                        & "USECLIEMAIL:0" & vbCRLF _ 
                                        & "EMAIL:sona.gyulamiryan@armsoft.am" & vbCRLF _ 
                                        & "ACCCONNECT:001" & vbCRLF _ 
                                        & "USECLISCH:1" & vbCRLF _ 
                                        & "FEEFROMCARD:0" & vbCRLF _ 
                                        & "COMM:For Test" & vbCRLF _ 
                                        & "'"
                sqlValue = 1
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If 
                
      ' Մուտք 'Աշխատանքային փաստաթղթեր' թղթապանակ
      Set workingDocs = New_PeriodicWorkingDocuments()
      With workingDocs
        .startDate = dateTimeNow
				.endDate = dateTimeNow
				.curr = "000"
				.commonPaySys = ""
				.office = "P00"
				.section = "08"
      End With
      
      Call GoTo_PeriodicWorkingDocuments(folderName, workingDocs)
         
      ' Վավերացնել փաստաթուղթը
      Call DocValidate(communalPay.docNum)
      
      BuiltIn.Delay(2000)
      Set frmPttel = Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel")
      frmPttel.Close
      
               'DOCLOG
                queryString = " SELECT COUNT(*) FROM DOCLOG WHERE fISN = " & communalPay.fisn & _
                                          " and fSUID = '10' and fSUIDCOR = '-1'" & _ 
                                          " and ((fOP = 'N' and fSTATE = '1')" & _ 
                                          " or (fOP = 'W' and fSTATE = '2')" & _ 
                                          " or (fOP = 'C' and fSTATE = '7'))"
                sqlValue = 3
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If 
                
                'FOLDERS
                queryString = " SELECT COUNT(*) FROM FOLDERS WHERE fISN = " & communalPay.fisn & _
                                          " and fNAME = 'PCPAGR  ' and fSTATUS = '1'" & _ 
                                          " and ((fCOM = 'ä³ñµ»ñ³Ï³Ý ÏáÙáõÝ³É í×³ñáõÙÝ»ñÇ å³ÛÙ³Ý³·Çñ' and fSPEC = '²Ùë³ÃÇí- 01/08/21 N- "&communalPay.docNum&" [Ð³ëï³ïí³Í]' and fECOM = 'Periodic communal payments agreement')" & _ 
                                          " or(fCOM = 'Ð³×³Ëáñ¹ 00002248' and fSPEC = '7   000022482021080120220101122810          50000.00100022482010000000011000000000000000000000000sona.gyulamiryan@armsoft.am' and fECOM = 'Client 00002248'))"
                sqlValue = 2
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If 
                
                'DOCS
                queryString = " SELECT COUNT(*) FROM DOCS WHERE fISN = " & communalPay.fisn & _
                                         " and fNAME = 'PCPAGR' and fSTATE = '7' and fNEXTTRANS = '1'"
                sqlValue = 1
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If 
                
                'PERIODIC_COMMUNAL
                queryString = " SELECT COUNT(*) FROM PERIODIC_COMMUNAL WHERE fISN = " & communalPay.fisn & _
                                         " and ((fROWID = '2' and fSYS = 'R' and fLOCATION = '9' and fCODE = '12402' and fABONENT = 'Þ»ÏáÛ³Ý ìÉ³¹ÇÙÇñ' and fADDRESS = 'î. Ø»ÍÇ å. 79 37' and fMIN = '0.00' and fMAX = '7000.00' and fPAID = '0.00' and fCOMPLETED = '0' and fJUR = '0')" & _ 
                                         " or (fROWID = '1' and fSYS = 'WN' and fLOCATION = '01' and fCODE = '7-62-0-0-101' and fABONENT = '´¸àÚ²Ü ²Ü¸ð²ÜÆÎ' and fADDRESS = 'æðìºÄ 18ö  0 - 18/2 0 18/2' and fMIN = '0.00' and fMAX = '2000.00' and fPAID = '0.00' and fCOMPLETED = '0' and fJUR = '0'))"
                sqlValue = 2
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If 
      
      Call ChangeWorkspace(c_Admin40)
      
      ' Ավելացնել առաջադրանքի ձևանմուշ
  		Log.Message("Ավելացնել Առաջադրանքի ձևանմուշ")
  		Call Add_TaskTemplate("PerComm", "ä³ñµ»ñ³Ï³Ý ÏáÙáõÝ³É í×³ñáõÙÝ»ñ", "Periodic Communal Payments")
		
  		' Ստուգել ավելացրած ձևանմուշի առկայությունը
  		Log.Message("Ստուգել ավելացրած ձևանմուշի առկայությունը")
  		Call Search_In_EditTree("frmEditTree", "PerComm  ä³ñµ»ñ³Ï³Ý ÏáÙáõÝ³É í×³ñáõÙÝ»ñ")
		
  		' Ավելացնել ձևանմուշի տարր
  		Log.Message ("Ավելացնել Ձևանմուշի տարր")
  		Call Add_TemplateElement(1, "PCPJOB", "ä³ñµ»ñ³Ï³Ý ÏáÙáõÝ³É í×³ñáõÙÝ»ñ", "frmEditTree", "PerComm  ä³ñµ»ñ³Ï³Ý ÏáÙáõÝ³É í×³ñáõÙÝ»ñ")
      
      BuiltIn.Delay(1000)
      frmPttel.Close
		  wMDIClient.VBObject("frmEditTree").Close
      
      ' Ավելացնել առաջադրանք և կատարել անմիջապես կլաս
  		Log.Message("Ավելացնել Առաջադրանք")
      Set addJob = New_AddJobForPetComm()
      With addJob
            .folderName = "|²¹ÙÇÝÇëïñ³ïáñÇ ²Þî 4.0|²é³ç³¹ñ³ÝùÝ»ñ|²é³ç³¹ñ³ÝùÝ»ñ"
            .sDate = "010121"
            .eDate = "010124"
            .startTime = aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%d/%m/%y")
            .wHour = "1200"
            .JobDate = "230921"
            .jobGroup = "PerComm"
            .sendEmail = "àã"
      End With
    
      ' Ավելացնել առաջադրանք և կատարել անմիջապես
  		Call Fill_AddJobForPetComm(addJob)
      
      ' Ստանալ առաջադրանքի համարը
      jobISNColN = wMDIClient.VBObject("frmPttel").GetColumnIndex("fISN")
      getJobISN = frmPttel.VBObject("TDBGView").Columns.Item(jobISNColN)
      Log.Message("առաջադրանքի համարը` " & getJobISN)
      
  		BuiltIn.Delay(2000)
      frmPttel.Close
      
                'TREES
                queryString = " SELECT COUNT(*) FROM TREES WHERE fCODE = 'PerComm'" & _
                                         " and fNAME = 'ä³ñµ»ñ³Ï³Ý ÏáÙáõÝ³É í×³ñáõÙÝ»ñ' and fCODX = 'PerComm' and fENAME = 'Periodic Communal Payments'"
                sqlValue = 1
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If 
      
  		' Մուտք"Պարբերական գործողությունների ԱՇՏ
      Call ChangeWorkspace(c_PeriodicActions)
      
      ' Մուտք"Պարբերական գործողությունների պայմանագրեր/Կոմունալ վճ. պայմանագրեր" թղթապանակ
      Set communalPayDoc = New_CommunalPayDoc()
      With communalPayDoc
          .folderName = "|ä³ñµ»ñ³Ï³Ý ·áñÍáÕáõÃÛáõÝÝ»ñÇ ²Þî|ä³ñµ»ñ³Ï³Ý ·áñÍáÕáõÃÛáõÝÝ»ñÇ å³ÛÙ³Ý³·ñ»ñ|ÎáÙáõÝ³É í×. å³ÛÙ³Ý³·ñ»ñ"
          .wClient = "00002248"
          .wBranch = "P00"
          .wDepart = "08"
      End With
      
      Call Fill_CommunalPayDoc(communalPayDoc)
      
      ' Պայմանագրի առկայության ստուգում 
      status = CheckContractDoc(0, communalPay.docNum)
      
      If Not status Then
            Log.Error"Պայմանագիրն առկա չէ 'Պարբերական կոմունալ վճարումների պայմանագրեր' թղթապանակում " ,,,ErrorColor
            Exit Sub  
      End If 
      
    
		
       ' Պարբերական կոմունալ վճարումների դիտում
      Set viewPayment = New_ViewPayment()
      With viewPayment
              .stDate = "010821"
              .eDate = "010124"
              .wBranch = "P00"
              .wDepart = "08"
      End With
      
      Call Fill_ViewPayment(viewPayment)
      
      If  wMDIClient.VBObject("frmPttel_2").VBObject("TDBGView").ApproxCount <> 1 Then
             Log.Error("Վճարում փաստաթուղթն առկա չէ Պարբերական կոմունալ վճարումներ թղթապանակում")
      End If
      
      ' Ստանալ վճարման փաստաթղթի ISN-ը և Փաստաթղթի N-ը
      Call wMainForm.MainMenu.Click(c_AllActions)
      Call wMainForm.PopupMenu.Click(c_View)
      BuiltIn.Delay(1000)
      commISN = wMDIClient.VBObject("frmASDocForm").DocFormCommon.Doc.ISN
      commDocNum = wMDIClient.VBObject("frmASDocForm").vbObject("TabFrame").VBObject(GetVBObject("DOCNUM", wMDIClient.VBObject("frmASDocForm"))).Text
      
      Call ClickCmdButton(1, "OK")
      
                'COM_PAYMENTS
                queryString = " SELECT COUNT(*) FROM COM_PAYMENTS WHERE fISN = " & commISN & _
                                         " and ((fNUMBER = '0' and fTYPE = 'WN' and fLOCATION = '01' and fCODE = '7-62-0-0-101' and fAMOUNT = '2000.00' and fNAME = '´¸àÚ²Ü ²Ü¸ð²ÜÆÎ' and fADDRESS = 'æðìºÄ 18ö  0 - 18/2 0 18/2' and fBRANCH = 'P00' and fDEPART = '08' and fDEBT = '11870.00') "& _
                                         " or (fNUMBER = '1' and fTYPE = 'R' and fLOCATION = '9' and fCODE = '12402' and fAMOUNT = '2020.00' and fNAME = 'Þ»ÏáÛ³Ý ìÉ³¹ÇÙÇñ' and fADDRESS = 'î. Ø»ÍÇ å. 79 37' and fBRANCH = 'P00' and fDEPART = '08' and fDEBT = '2020.00'))"
                sqlValue = 2
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If 
                
                'DOCLOG
                queryString = " SELECT COUNT(*) FROM DOCLOG WHERE fISN = " & communalPay.fisn & _
                                          " and fSUID = '10' and fSUIDCOR = '-1'" & _
                                          " and ((fOP = 'N' and fSTATE = '1')" & _
                                          " or (fOP = 'W' and fSTATE = '2')" & _
                                          " or (fOP = 'C' and fSTATE = '7')" & _
                                          " or (fOP = 'E' and fSTATE = '7'))"
                sqlValue = 4
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If 
                
                'DOCS
                queryString = " SELECT COUNT(*) FROM DOCS WHERE fISN = " & communalPay.fisn & _
                                         " and fNAME = 'PCPAGR' and fSTATE = '7' and fNEXTTRANS = '1' and fBODY = '"& vbCRLF _
                                        & "ACSBRANCH:P00" & vbCRLF _
                                        & "ACSDEPART:08" & vbCRLF _
                                        & "USERID:10" & vbCRLF _
                                        & "CODE:"& communalPay.docNum & vbCRLF _
                                        & "CLICODE:00002248" & vbCRLF _
                                        & "NAME:Ð³×³Ëáñ¹ 00002248" & vbCRLF _
                                        & "ENAME:Client 00002248" & vbCRLF _
                                        & "FEEACC:00224820100" & vbCRLF _
                                        & "FEECUR:000" & vbCRLF _
                                        & "MAXSUM:50000" & vbCRLF _
                                        & "SDATE:20210801" & vbCRLF _
                                        & "EDATE:20220101" & vbCRLF _
                                        & "SDAY:12" & vbCRLF _
                                        & "LDAY:28" & vbCRLF _
                                        & "CLINOT:1" & vbCRLF _
                                        & "USECLIEMAIL:0" & vbCRLF _
                                        & "EMAIL:sona.gyulamiryan@armsoft.am" & vbCRLF _
                                        & "ACCCONNECT:001" & vbCRLF _
                                        & "USECLISCH:1" & vbCRLF _
                                        & "FEEFROMCARD:0" & vbCRLF _
                                        & "COMM:For Test" & vbCRLF _
                                        & "LASTOPDATE:20210923" & vbCRLF _
                                        & "LASTCOMPLETE:20210923" & vbCRLF _
                                        & "'"
                sqlValue = 1
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If 
                
                
                'DOCSG
                queryString = " SELECT COUNT(*) FROM DOCSG WHERE fISN = " & commISN 
                sqlValue = 125
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If 
                
                'FOLDERS
                queryString = " SELECT COUNT(*) FROM FOLDERS WHERE fISN = " & commISN & _
                                         " and fNAME = 'ComGrPay' and fSTATUS = '1'" & _
                                         " and fCOM = 'ÎáÙáõÝ³É í×³ñáõÙÝ»ñÇ Ñ³ÝÓÝ³ñ³ñ³·Çñ' and fECOM = '' "
                sqlValue = 1
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If 
                
                'HI
                queryString = " SELECT COUNT(*) FROM HI WHERE fBASE = " & commISN & _
                                         " and fCUR = '000' and fOP = 'MSC' and fSUID = '10' and fBASEBRANCH = 'P00' and fBASEDEPART = '08' and fTYPE = '01'" & _
                                        " and ((fSUM = '2000.00' and  fCURSUM = '2000.00' and fDBCR = 'D' and fSPEC = '"& commDocNum &"                   æñÇ í³ñÓ                          1     1.0000    1' and fTRANS = '0')" & _
                                        " or (fSUM = '2000.00' and  fCURSUM = '2000.00' and fDBCR = 'C' and fSPEC = '"& commDocNum &"                   æñÇ í³ñÓ                          0     1.0000    1                                                                        ´¸àÚ²Ü ²Ü¸ð²ÜÆÎ' and fTRANS = '0')" & _
                                        " or (fSUM = '2020.00' and  fCURSUM = '2020.00' and fDBCR = 'D' and fSPEC = '"& commDocNum &"                   ²Õµ³Ñ³ÝÙ³Ý í³ñÓ                   1     1.0000    1' and fTRANS = '1')" & _
                                        " or (fSUM = '2020.00' and  fCURSUM = '2020.00' and fDBCR = 'C' and fSPEC = '"& commDocNum &"                   ²Õµ³Ñ³ÝÙ³Ý í³ñÓ                   0     1.0000    1                                                                        Þ»ÏáÛ³Ý ìÉ³¹ÇÙÇñ' and fTRANS = '1'))"
                sqlValue = 4
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If 
                
                'PAYMENTS
                queryString = " SELECT COUNT(*) FROM PAYMENTS WHERE fISN = " & commISN & _
                                         " and fDOCTYPE = 'ComGrPay' and fCUR = '000' and fCOM = 'ÎáÙáõÝ³É í×³ñáõÙ/Utility Payment'"& _
                                        " and fCHARGESUM = '0.00' and fCHARGESUMAMD = '0.00' and fCHARGESUM2 = '0.00' and fCHARGESUMAMD2 = '0.00'"& _
                                        " and ((fSUMMA = '2000.00' and fSUMMAAMD = '2000.00' and fSUMMAUSD = '4.82')"& _
                                        " or (fSUMMA = '2020.00' and fSUMMAAMD = '2020.00' and fSUMMAUSD = '4.8682'))"
                sqlValue = 2
                colNum = 0
                sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                If Not sql_isEqual Then
                  Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                End If 
                
      ' Ջնջել վճարման փաստաթուղթը
      Call wMainForm.MainMenu.Click(c_AllActions)
      Call wMainForm.PopupMenu.Click(c_Delete )
      If  MessageExists(2, "Ð³×³Ëáñ¹ÇÝ Ï³ñáÕ ¿ áõÕ³ñÏí³Í ÉÇÝ»É ïíÛ³É í×³ñÙ³Ý Ù³ëÇÝ " & vbCrLf & "Ñ³Õáñ¹³·ñáõÃÛáõÝ:" & vbCrLf & "Þ³ñáõÝ³Ï»±É:") Then
           ' Սեղմել "OK" կոճակը
           Call ClickCmdButton(5, "²Ûá")  
            If  MessageExists(2, "ö³ëï³ÃáõÕÃÁ çÝç»ÉÇë` ÏÑ»é³óí»Ý Ýñ³ Ñ»ï Ï³åí³Í ËÙµ³ÛÇÝ " & vbCrLf &"Ó¨³Ï»ñåáõÙÝ»ñÁ") Then
                   ' Սեղմել "Կատարել" կոճակը
                   Call ClickCmdButton(5, "Î³ï³ñ»É")  
                         If  MessageExists(1, "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ") Then
                                ' Սեղմել "Այո" կոճակը
                                Call ClickCmdButton(3, "²Ûá")  
                         Else
                                 Log.Error"Հաղորդագրության պատուհանը չի բացվել" ,,,ErrorColor
                         End If
             Else
                      Log.Error"Հաղորդագրության պատուհանը չի բացվել" ,,,ErrorColor
             End If
         Else
            Log.Error"Հաղորդագրության պատուհանը չի բացվել" ,,,ErrorColor
      End If
          
      BuiltIn.Delay(1000) 
      ' Փակել ընթացիկ պատուհանը
      Call wMainForm.MainMenu.Click(c_Windows)
      Call wMainForm.PopupMenu.Click(c_ClCurrWindow)
      
      ' Ջնջել պայմանագիրը                
      Call DelDoc()
      BuiltIn.Delay(1000)
      frmPttel.Close
      
      Call ChangeWorkspace(c_Admin40)
      ' Ջնջել առաջադրանքի ձևանմուշն ու առաջադրանքը
      Call Delete_Task(getJobISN)
		
  		BuiltIn.Delay(3000)
  		frmPttel.Close
      
      ' Մուտք Առաջադրանքի ձևանմուշներ թղթապանակ
      wTreeView.DblClickItem("|²¹ÙÇÝÇëïñ³ïáñÇ ²Þî 4.0|²é³ç³¹ñ³ÝùÝ»ñ|²é³ç³¹ñ³ÝùÇ Ó¨³ÝÙáõßÝ»ñ")
      ' Ստուգել ավելացրած ձևանմուշի առկայությունը
  		Log.Message("Ստուգել ավելացրած ձևանմուշի առկայությունը")
  		status =  Search_In_EditTree("frmEditTree", "PerComm  ä³ñµ»ñ³Ï³Ý ÏáÙáõÝ³É í×³ñáõÙÝ»ñ")
      
      If status Then
          ' Դիտել ձևանմուշը
          Call wMainForm.MainMenu.Click(c_AllActions)
          Call wMainForm.PopupMenu.Click(c_ViewTemplate)
          BuiltIn.Delay(1500)
          ' Ջնջել ձևանմուշը
          Call DelDoc()
          
          ' Փակել ընթացիկ պատուհանը
          Call wMainForm.MainMenu.Click(c_Windows)
          Call wMainForm.PopupMenu.Click(c_ClCurrWindow)
          
          ' Ջնջել առաջադրանքի հանգույցը
          Call wMainForm.MainMenu.Click(c_AllActions)
          Call wMainForm.PopupMenu.Click(c_Delete)
          
          If MessageExists(2, "Ð³ëï³ï»ù Ñ³Ý·áõÛóÇ çÝç»ÉÁ") Then
                Call ClickCmdButton(5, "²Ûá") 
          Else
              Log.Error"Ջնջել պատուհանը չի բացվել",,, ErrorColor
          End If
      Else
            Log.Error"Ձևանմուշի առկա չէ Առաջադրանքի ձևանմուշներ թղթապանակում",,, ErrorColor
      End If
      
      ' Փակել ընթացիկ պատուհանը
      Call wMainForm.MainMenu.Click(c_Windows)
      Call wMainForm.PopupMenu.Click(c_ClCurrWindow)
      
      ' Փակել ծրագիրը
      Call Close_AsBank()   
      
End Sub