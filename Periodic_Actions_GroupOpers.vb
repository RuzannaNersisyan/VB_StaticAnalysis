'USEUNIT Library_Common
'USEUNIT Library_Colour
'USEUNIT Library_Contracts 
'USEUNIT Constants
'USEUNIT Library_CheckDB
'USEUNIT Library_Periodic_Actions
Option Explicit

'Test Case N 159171

Dim sDATE, fDATE, folderName, periodActions1, periodActions2
Dim Working_Docs, periodicAct, arrContractNums(2), groupEdit, chgReqs
Dim dbo_FOLDERS(5), fBODY, i, dbo_FOLDERS2(5)

Sub Periodic_Actions_GroupOpers()
		Call Test_Initialize()

		' Ð³Ù³Ï³ñ· Ùáõïù ·áñÍ»É ARMSOFT û·ï³·áñÍáÕáí
		Log.Message "Համակարգ մուտք գործել ARMSOFT օգտագործողով", "", pmNormal, DivideColor
  Call Test_StartUp()
		
		' êï»ÕÍ»É ä³ñµ»ñ³Ï³Ý ·áñÍáÕáõÃÛáõÝÝ»ñÇ å³ÛÙ³Ý³·Çñ
		Log.Message "Ստեղծել Պարբերական գործողությունների պայմանագիր", "", pmNormal, DivideColor
  Call Create_PeriodicActions(folderName, periodActions1, "create")
		
		' ä³ñµ»ñ³Ï³Ý ·áñÍáÕáõÃÛáõÝÝ»ñÇ å³ÛÙ³Ý³·ñÇ ëï»ÕÍáõÙÇó Ñ»ïá SQL ëïáõ·áõÙ
		Log.Message "Պարբերական գործողությունների պայմանագրի ստաղծումից հետո SQL ստուգում", "", pmNormal, SqlDivideColor
		Call DB_Initialize()
		Call Check_DB_PeriodicActionsCreation()
		
		' ì³í»ñ³óÝ»É å³ÛÙ³Ý³·ÇñÁ
		Log.Message "Վավերացնել պայմանագիրը", "", pmNormal, DivideColor
		Call GoTo_PeriodicWorkingDocuments(folderName, Working_Docs)
		Call SearchInPttel("frmPttel", 2, periodActions1.general.agreeN)
		Call Verify_Periodic_Actions()
		
		BuiltIn.Delay(3000)
		wMDIClient.VBObject("frmPttel").Close
		
		' ä³ÛÙ³Ý³·ñÇ í³í»ñ³óáõÙÇó Ñ»ïá SQL ëïáõ·áõÙ
		Log.Message "Պայմանագրի վավերացումից հետո SQL ստուգում", "", pmNormal, SqlDivideColor
		Call Check_DB_Confirm()
		
		' ä³ñµ»ñ³Ï³Ý ·áñÍáÕáõÃÛáõÝÝ»ñÇ å³ÛÙ³Ý³·ñ»ñ
		Log.Message "Պարբերական գործողությունների պայմանագրեր", "", pmNormal, DivideColor
		Call	Check_PeriodicExisting(folderName & "ä³ñµ»ñ³Ï³Ý ·áñÍáÕáõÃÛáõÝÝ»ñÇ å³ÛÙ³Ý³·ñ»ñ|", periodicAct, periodActions1.general.agreeN)
		
		' Î³ï³ñ»É í×³ñáõÙ
		Log.Message "Կատարել վճարում", "", pmNormal, DivideColor
		Call MakePayment(periodActions1.general.startDate, 1, 1)
		
		' Î³ï³ñ»É í×³ñáõÙÇó Ñ»ïá SQL ëïáõ·áõÙ
		Log.Message "Կատարել վճարումից հետո SQL ստուգում", "", pmNormal, SqlDivideColor
		Call Check_DB_MakePayment()
		
		' ì×³ñáõÙÝ»ñÇ ¹ÇïáõÙ
		Log.Message "Վճարումների դիտում", "", pmNormal, DivideColor
		Call PaymentView(periodActions1.general.startDate, periodActions1.general.startDate, 1)
		
		BuiltIn.Delay(3000)
		wMDIClient.VBObject("frmPttel_2").Close
		
		' ²í»É³óÝ»É ä³ñµ»ñ³Ï³Ý ·áñÍáÕáõÃÛáõÝÝ»ñÇ å³ÛÙ³Ý³·ñ»ñ
		Log.Message "Ավելացնել Պարբերական գործողությունների պայմանագրեր", "", pmNormal, DivideColor
		Call Create_PeriodicActions(folderName, periodActions2, "add")
		
		BuiltIn.Delay(3000)
		wMDIClient.VBObject("frmPttel").Close
		
		' ä³ñµ»ñ³Ï³Ý ·áñÍáÕáõÃÛáõÝÝ»ñÇ å³ÛÙ³Ý³·ñÇ ³í»É³óáõÙÇó Ñ»ïá SQL ëïáõ·áõÙ
		Log.Message "Պարբերական գործողությունների պայմանագրի ավելացումից հետո SQL ստուգում", "", pmNormal, SqlDivideColor
		Call Check_DB_PeriodicActionsAdd()
		
		' ì³í»ñ³óÝ»É å³ÛÙ³Ý³·ÇñÁ
		Log.Message "Վավերացնել պայմանագիրը", "", pmNormal, DivideColor
		Call GoTo_PeriodicWorkingDocuments(folderName, Working_Docs)
		Call SearchInPttel("frmPttel", 2, periodActions2.general.agreeN)
		Call Verify_Periodic_Actions()
		
		BuiltIn.Delay(3000)
		wMDIClient.VBObject("frmPttel").Close
		
		' ä³ÛÙ³Ý³·ñÇ í³í»ñ³óáõÙÇó Ñ»ïá SQL ëïáõ·áõÙ
		Log.Message "Պայմանագրի վավերացումից հետո SQL ստուգում", "", pmNormal, SqlDivideColor
		Call Check_DB_Confirm2()

			' ä³ñµ»ñ³Ï³Ý ·áñÍáÕáõÃÛáõÝÝ»ñÇ å³ÛÙ³Ý³·ñ»ñ
		Log.Message "Պարբերական գործողությունների պայմանագրեր", "", pmNormal, DivideColor
		Call	Check_PeriodicExisting(folderName & "ä³ñµ»ñ³Ï³Ý ·áñÍáÕáõÃÛáõÝÝ»ñÇ å³ÛÙ³Ý³·ñ»ñ|", periodicAct, periodActions2.general.agreeN)

			' Î³ï³ñ»É í×³ñáõÙ
		Log.Message "Կատարել վճարում", "", pmNormal, DivideColor
		Call MakePayment(periodActions2.general.startDate, 1, 1)
		
		' Î³ï³ñ»É í×³ñáõÙÇó Ñ»ïá SQL ëïáõ·áõÙ
		Log.Message "Կատարել վճարումից հետո SQL ստուգում", "", pmNormal, SqlDivideColor
		Call Check_DB_MakePayment2()
		
		' ì×³ñáõÙÝ»ñÇ ¹ÇïáõÙ
		Log.Message "Վճարումների դիտում", "", pmNormal, DivideColor
		Call PaymentView(periodActions2.general.startDate, periodActions2.general.startDate, 1)
		
		BuiltIn.Delay(3000)
		wMDIClient.VBObject("frmPttel_2").Close

		BuiltIn.Delay(4000)
		wMDIClient.VBObject("frmPttel").Close
				
		' ä³ñµ»ñ³Ï³Ý ·áñÍáÕáõÃÛáõÝÝ»ñÇ å³ÛÙ³Ý³·ñ»ñ
		Log.Message "Պարբերական գործողությունների պայմանագրեր", "", pmNormal, DivideColor
		Call	GoTo_PeriodicActionsAgree(folderName & "ä³ñµ»ñ³Ï³Ý ·áñÍáÕáõÃÛáõÝÝ»ñÇ å³ÛÙ³Ý³·ñ»ñ|", periodicAct)
		
		' Î³ï³ñ»É ¸³ßï»ñÇ ËÙµ³ÛÇÝ ËÙµ³·ñáõÙ
		Log.Message "Դաշտերի խմբային խմբագրում", "", pmNormal, DivideColor
		arrContractNums(0) = periodActions1.general.agreeN
		arrContractNums(1) = periodActions2.general.agreeN
		BuiltIn.Delay(6000)
		Call SelectRowByColumn(arrContractNums, 0)
		Call FieldsGroupEdit(groupEdit)
		
		BuiltIn.Delay(3000)
		wMDIClient.VBObject("frmPttel").Close
		
		' ¸³ßï»ñÇ ËÙµ³ÛÇÝ ËÙµ³·ñáõÙÇó Ñ»ïá SQL ëïáõ·áõÙ
		Log.Message "Դաշտերի խմբային խմբագրումից հետո SQL ստուգում", "", pmNormal, SqlDivideColor
		Call Check_DB_GroupEdit()
		
		' Ð³ëï³ï»É ÷á÷áËáõÃÛáõÝÝ»ñÁ öá÷áËÙ³Ý Ñ³Ûï»ñÇó
		Log.Message "Հաստատել փոփոխությունները Փոփոխման հայտերից", "", pmNormal, DivideColor
		Call GoTo_ChangeRequests(folderName & "ä³ñµ»ñ³Ï³Ý ·áñÍáÕáõÃÛáõÝÝ»ñÇ å³ÛÙ³Ý³·ñ»ñ|", chgReqs)
		Call Confirm_ChangeRequest(periodActions1.general.agreeN, 8)
		Call Confirm_ChangeRequest(periodActions2.general.agreeN, 8)
		
		BuiltIn.Delay(3000)
		wMDIClient.VBObject("frmPttel").Close
		
		' Ð³ëï³ï»É ÷á÷áËáõÃÛáõÝÝ»ñÁ öá÷áËÙ³Ý Ñ³Ûï»ñÇó Ñ»ïá SQL ëïáõ·áõÙ
		Log.Message "Հաստատել փոփոխությունները Փոփոխման հայտերից հետո SQL ստուգում", "", pmNormal, SqlDivideColor
		Call Check_DB_ConfirmGroupEdit()
		
		' æÝç»É ä³ñµ»ñ³Ï³Ý ·áñÍáÕáõÃÛáõÝÝ»ñÇ å³ÛÙ³Ý³·ñ»ñÁ
		Log.Message "Ջնջել Պարբերական գործողությունների պայմանագրերը", "", pmNormal, DivideColor
		Call	GoTo_PeriodicActionsAgree(folderName & "ä³ñµ»ñ³Ï³Ý ·áñÍáÕáõÃÛáõÝÝ»ñÇ å³ÛÙ³Ý³·ñ»ñ|", periodicAct)
		
		Log.Message "Վճարումների դիտում", "", pmNormal, DivideColor
		Call SearchInPttel("frmPttel", 0, periodActions1.general.agreeN)
		BuiltIn.Delay(3000)
		Call PaymentView(periodActions1.general.startDate, periodActions1.general.startDate, 1)
		Call SearchAndDelete("frmPttel_2", 1, periodActions1.general.agreeN, "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ")
		
		BuiltIn.Delay(3000)
		wMDIClient.VBObject("frmPttel_2").Close
		
		Call SearchAndDelete("frmPttel", 0, periodActions1.general.agreeN, "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ")
		
		Log.Message "Վճարումների դիտում", "", pmNormal, DivideColor
		Call SearchInPttel("frmPttel", 0, periodActions2.general.agreeN)
		Call PaymentView(periodActions2.general.startDate, periodActions2.general.startDate, 1)
		Call SearchAndDelete("frmPttel_2", 1, periodActions2.general.agreeN, "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ")
		
		BuiltIn.Delay(3000)
		wMDIClient.VBObject("frmPttel_2").Close
		
		Call SearchAndDelete("frmPttel", 0, periodActions2.general.agreeN, "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ")
		
		BuiltIn.Delay(3000)
		wMDIClient.VBObject("frmPttel").Close
		
		' æÝç»É ä³ñµ»ñ³Ï³Ý ·áñÍáÕáõÃÛáõÝÝ»ñÇ å³ÛÙ³Ý³·ñ»ñÇó Ñ»ïá SQL ëïáõ·áõÙ
		Log.Message "Ջնջել Պարբերական գործողությունների պայմանագրերից հետո SQL ստուգում", "", pmNormal, SqlDivideColor
		Call Check_DB_DeleteDocs()
		
		Call Close_AsBank()    
End Sub

Sub Test_StartUp()
		Call Initialize_AsBankQA(sDATE, fDATE)
  Login("ARMSOFT")
		Call ChangeWorkspace(c_PeriodicActions)
End	Sub

Sub Test_Initialize()
		folderName = "|ä³ñµ»ñ³Ï³Ý ·áñÍáÕáõÃÛáõÝÝ»ñÇ ²Þî|"
		sDATE = "20030101"
		fDATE = "20240101"  
		
		Set periodActions1 = New_PeriodicActions(1)
		with periodActions1
		  .general.office = "P00"
				.general.department = "08"
				.general.performer = "10"
				.general.client = "00000208"
				.general.doInEveryCall = 0
				.general.periodMounth = "0"
				.general.periodDay = "15"
				.general.implementDays_start = "1"
				.general.implementDays_end = "31"
				.general.bypassNonWorkDays = "1"
				.general.opersGridRowCount = 1 
				.general.operations(0).N_Edit = "1"
				.general.operations(0).operType = "01"
				.general.operations(0).calcMethod = "03"
				.general.operations(0).debitAccount = "00020880200"
				.general.operations(0).depositAccount  = "00020880100"
				.general.operations(0).percent = "2"
				.general.operations(0).price = ""
				.general.operations(0).curr = ""
				.general.operations(0).transactionRate = ""
				.general.operations(0).rateChange = ""
				.general.operations(0).daysCount = "0"
				.general.operations(0).minPrice = ""
				.general.operations(0).maxPrice = "" 
				.general.operations(0).aim = "Ñ³Ù³Ó³ÛÝ Ï³ñ·."
				.general.operations(0).addDocument = "0"
				.other.informToClient = 0
				.other.note = "002"
				.other.note2 = "102"
				.other.note3 = "01"
				.other.addInfo = "Èñ³óáõóÇã ÇÝýá"
		end with
		
		Set periodActions2 = New_PeriodicActions(1)
		with periodActions2
		  .general.office = "P00"
				.general.department = "08"
				.general.performer = "10"
				.general.client = "00009710"
				.general.doInEveryCall = 0
				.general.periodMounth = "1"
				.general.periodDay = "103"
				.general.bypassNonWorkDays = "0"
				.general.opersGridRowCount = 1 
				.general.operations(0).N_Edit = "1"
				.general.operations(0).operType = "04"
				.general.operations(0).calcMethod = "01"
				.general.operations(0).debitAccount = "00971040200"
				.general.operations(0).depositAccount  = "0123456"
				.general.operations(0).percent = "1"
				.general.operations(0).transactionRate = ""
				.general.operations(0).rateChange = ""
				.general.operations(0).daysCount = "0"
				.general.operations(0).opersAddDoc = true
				.general.operations(0).aim = "Ñ³Ù³Ó³ÛÝ å³ÛÙ³Ý³·ñÇ"
				.other.informToClient = 1
				.other.clientEmail = true
				.other.useClientEmail = 1
				.other.note = "001"
				.other.note2 = "193"
				.other.note3 = "08"
		end with

		Set Working_Docs = New_PeriodicWorkingDocuments()
		Working_Docs.performers = "10"
		
		Set periodicAct = New_PeriodicActionsAgree()
		periodicAct.performer = "10"
		
		Set groupEdit = New_FieldsGroup_Edit()
		with groupEdit
				.office = "P07"
				.department = "06"
				.emptyPerformer = 0
				.performer = ""
				.startDate = ""
				.endDate = ""
				.emptyEndDate = 0
				.doInEveryCall = 0
				.period_mounth = "3"
				.period_day = "16"
				.editImplementDays = 1
				.implementDays_start = "2"
				.implementDays_end = "9"
				.bypassNonWorkDays = "1"
				.informClient = "2"
				.emptyNote = 1
				.note = ""
				.emptyNote2 = 1
				.note2 = ""
				.emptyNote3 = 1
				.note3 = ""
				.emptyAddInfo = 0
				.addInfo = ""
		end with
		
		Set chgReqs = New_ChangeRequests()
		
End Sub

Sub DB_Initialize()
		for i = 0 to 4
    Set dbo_FOLDERS(i) = New_DB_FOLDERS()
    dbo_FOLDERS(i).fISN = periodActions1.fISN
    dbo_FOLDERS(i).fNAME = "PPAGR   "
  next
		dbo_FOLDERS(0).fKEY = periodActions1.fISN
		dbo_FOLDERS(0).fSTATUS = "5"
  dbo_FOLDERS(0).fFOLDERID = "C.102912"
  dbo_FOLDERS(0).fCOM = "ä³ñµ»ñ³Ï³Ý ·áñÍáÕáõÃÛáõÝÝ»ñÇ å³ÛÙ³Ý³·Çñ"
  dbo_FOLDERS(0).fSPEC = "²Ùë³ÃÇí- " & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%d/%m/%y") & " N- " & periodActions1.general.agreeN & " [Üáñ]"
		dbo_FOLDERS(0).fECOM = "Periodic payments agreement"
		dbo_FOLDERS(1).fKEY = periodActions1.fISN
		dbo_FOLDERS(1).fSTATUS = "5"
  dbo_FOLDERS(1).fFOLDERID = "Oper." & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%Y%m%d")
  dbo_FOLDERS(1).fCOM = "ä³ñµ»ñ³Ï³Ý ·áñÍáÕáõÃÛáõÝÝ»ñÇ å³ÛÙ³Ý³·Çñ"
  dbo_FOLDERS(1).fSPEC = periodActions1.general.agreeN & "16600                                       0.00   Üáñ                                                   10Ð³×³Ëáñ¹ " & periodActions1.general.client & "                                                                               ä³ñµ. ·áñÍ. å³ÛÙ³Ý³·Çñ                                                                                                                      "
		dbo_FOLDERS(1).fECOM = "Periodic payments agreement"
		dbo_FOLDERS(1).fDCBRANCH = "P00"
  dbo_FOLDERS(1).fDCDEPART = "08 "
		dbo_FOLDERS(2).fKEY = periodActions1.general.agreeN
		dbo_FOLDERS(2).fSTATUS = "1"
  dbo_FOLDERS(2).fFOLDERID = "PPAYMS"
  dbo_FOLDERS(2).fCOM = "Ð³×³Ëáñ¹ " & periodActions1.general.client
  dbo_FOLDERS(2).fSPEC = "1   0000020810  " & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%Y%m%d") & "000000000 0/ 15 131100                                   0000210201 Èñ³óáõóÇã ÇÝýá                    0000000000000000                                                                                                                                                                                                                            0     "
		dbo_FOLDERS(2).fECOM = "Client " & periodActions1.general.client
		dbo_FOLDERS(2).fDCBRANCH = "P00"
  dbo_FOLDERS(2).fDCDEPART = "08 "
		dbo_FOLDERS(3).fKEY = periodActions1.general.agreeN & "_1"
		dbo_FOLDERS(3).fSTATUS = "1"
  dbo_FOLDERS(3).fFOLDERID = "PPAYMSEXT" 
  dbo_FOLDERS(3).fCOM = "Ð³×³Ëáñ¹ " & periodActions1.general.client
  dbo_FOLDERS(3).fSPEC = "1   0000020810  " & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%Y%m%d") & "000000000 0/ 15 131100                                   0000210201 Èñ³óáõóÇã ÇÝýá                    00000000000000001 0103" & periodActions1.general.operations(0).debitAccount & "000" & periodActions1.general.operations(0).depositAccount & "                       000                                   2                                           0 0               0               Ñ³Ù³Ó³ÛÝ Ï³ñ·.                            000000000                                                  "
  dbo_FOLDERS(3).fECOM = "Client " & periodActions1.general.client
  dbo_FOLDERS(3).fDCBRANCH = "P00"
  dbo_FOLDERS(3).fDCDEPART = "08 "
End	Sub

Sub Check_DB_PeriodicActionsCreation()
		Dim i		
	 'SQL Ստուգում DOCLOG աղյուսակում համար
  Log.Message "SQL Ստուգում DOCLOG աղյուսակում", "", pmNormal, SqlDivideColor
  Call CheckQueryRowCount("DOCLOG", "fISN", periodActions1.fISN, 1)
  Call CheckDB_DOCLOG(periodActions1.fISN, "10", "N", "1", "", 1)
  
  'SQL Ստուգում DOCP աղյուսակում  
  Log.Message "SQL Ստուգում DOCP աղյուսակում", "", pmNormal, SqlDivideColor
  Call CheckQueryRowCount("DOCP", "fPARENTISN", periodActions1.fISN, 2)
  Call CheckDB_DOCP("416614044", "Acc     ", periodActions1.fISN, 1)
		Call CheckDB_DOCP("1008763443", "Acc     ", periodActions1.fISN, 1)
  
  'SQL Ստուգում DOCS աղյուսակում 
  Log.Message "SQL Ստուգում DOCS աղյուսակում", "", pmNormal, SqlDivideColor
 	fBODY = " ACSBRANCH:P00 ACSDEPART:08 USERID:10 CODE:" & periodActions1.general.agreeN & " CLICODE:00000208 NAME:Ð³×³Ëáñ¹ 00000208 ENAME:Client 00000208 CALCALWAYS:0 PERIODICITY:0/15 SDAY:1 LDAY:31 NONWORKDAYS:1 CLINOT:0 USECLIEMAIL:0 USECLISCH:0 FEEFROMCARD:0 NOTE1:002 NOTE2:102 NOTE3:01 COMM:Èñ³óáõóÇã ÇÝýá "
  fBODY = Replace(fBODY, " ", "%")
  Call CheckQueryRowCount("DOCS", "fISN", periodActions1.fISN, 1)
  Call CheckDB_DOCS(periodActions1.fISN, "PPAGR   ", "1", fBODY, 1)
  
  'SQL Ստուգում DOCSG աղյուսակում 
  Log.Message "SQL Ստուգում DOCSG աղյուսակում", "", pmNormal, SqlDivideColor
  Call CheckQueryRowCount("DOCSG", "fISN", periodActions1.fISN, 10)
  
  'SQL Ստուգում FOLDERS աղյուսակում 
  Log.Message "SQL Ստուգում FOLDERS աղյուսակում", "", pmNormal, SqlDivideColor
  Call CheckQueryRowCount("FOLDERS", "fISN", periodActions1.fISN, 4)
  for i = 0 to 3
    Call CheckDB_FOLDERS(dbo_FOLDERS(i), 1)
  next
End	Sub

Sub Check_DB_Confirm()
		Dim i		
	 'SQL Ստուգում DOCLOG աղյուսակում համար
  Log.Message "SQL Ստուգում DOCLOG աղյուսակում", "", pmNormal, SqlDivideColor
  Call CheckQueryRowCount("DOCLOG", "fISN", periodActions1.fISN, 3)
  Call CheckDB_DOCLOG(periodActions1.fISN, "10", "W", "2", "", 1)
		Call CheckDB_DOCLOG(periodActions1.fISN, "10", "C", "7", "", 1)
  
  'SQL Ստուգում DOCS աղյուսակում 
  Log.Message "SQL Ստուգում DOCS աղյուսակում", "", pmNormal, SqlDivideColor
  Call CheckQueryRowCount("DOCS", "fISN", periodActions1.fISN, 1)
  Call CheckDB_DOCS(periodActions1.fISN, "PPAGR   ", "7", fBODY, 1)
  
  'SQL Ստուգում DOCSG աղյուսակում 
  Log.Message "SQL Ստուգում DOCSG աղյուսակում", "", pmNormal, SqlDivideColor
  Call CheckQueryRowCount("DOCSG", "fISN", periodActions1.fISN, 10)
  
  'SQL Ստուգում FOLDERS աղյուսակում 
  Log.Message "SQL Ստուգում FOLDERS աղյուսակում", "", pmNormal, SqlDivideColor
  Call CheckQueryRowCount("FOLDERS", "fISN", periodActions1.fISN, 3)
		dbo_FOLDERS(0).fSTATUS = "1"
		dbo_FOLDERS(0).fSPEC = "²Ùë³ÃÇí- " & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%d/%m/%y") & " N- " & periodActions1.general.agreeN & " [Ð³ëï³ïí³Í]"
		dbo_FOLDERS(2).fSPEC = "7   0000020810  " & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%Y%m%d") & "000000000 0/ 15 131100                                   0000210201 Èñ³óáõóÇã ÇÝýá                    0000000000000000                                                                                                                                                                                                                            0     "
		dbo_FOLDERS(3).fSPEC = "7   0000020810  " & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%Y%m%d") & "000000000 0/ 15 131100                                   0000210201 Èñ³óáõóÇã ÇÝýá                    00000000000000001 0103" & periodActions1.general.operations(0).debitAccount & "000" & periodActions1.general.operations(0).depositAccount & "                       000                                   2                                           0 0               0               Ñ³Ù³Ó³ÛÝ Ï³ñ·.                            000000000                                                  "
  for i = 0 to 3
				if i <> 1 then 
		    Call CheckDB_FOLDERS(dbo_FOLDERS(i), 1)
				end if
  next
End	Sub

Sub Check_DB_MakePayment()
		Dim i		
	 'SQL Ստուգում DOCLOG աղյուսակում համար
  Log.Message "SQL Ստուգում DOCLOG աղյուսակում", "", pmNormal, SqlDivideColor
  Call CheckQueryRowCount("DOCLOG", "fISN", periodActions1.fISN, 4)
  Call CheckDB_DOCLOG(periodActions1.fISN, "10", "W", "2", "", 1)
		Call CheckDB_DOCLOG(periodActions1.fISN, "10", "C", "7", "", 1)
		
		'SQL Ստուգում DOCS աղյուսակում 
  Log.Message "SQL Ստուգում DOCS աղյուսակում", "", pmNormal, SqlDivideColor
  Call CheckQueryRowCount("DOCS", "fISN", periodActions1.fISN, 1)
  Call CheckDB_DOCS(periodActions1.fISN, "PPAGR   ", "7", fBODY, 1)
  
  'SQL Ստուգում DOCSG աղյուսակում 
  Log.Message "SQL Ստուգում DOCSG աղյուսակում", "", pmNormal, SqlDivideColor
  Call CheckQueryRowCount("DOCSG", "fISN", periodActions1.fISN, 11)
  
  'SQL Ստուգում FOLDERS աղյուսակում 
  Log.Message "SQL Ստուգում FOLDERS աղյուսակում", "", pmNormal, SqlDivideColor
  Call CheckQueryRowCount("FOLDERS", "fISN", periodActions1.fISN, 3)
		dbo_FOLDERS(0).fSTATUS = "1"
		dbo_FOLDERS(0).fSPEC = "²Ùë³ÃÇí- " & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%d/%m/%y") & " N- " & periodActions1.general.agreeN & " [Ð³ëï³ïí³Í]"
		dbo_FOLDERS(2).fSPEC = "7   0000020810  " & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%Y%m%d") & "000000000 0/ 15 131100                                   0000210201 Èñ³óáõóÇã ÇÝýá                    " & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%Y%m%d") & "00000000                                                                                                                                                                                                                            0     "
		dbo_FOLDERS(3).fSPEC = "7   0000020810  " & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%Y%m%d") & "000000000 0/ 15 131100                                   0000210201 Èñ³óáõóÇã ÇÝýá                    " & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%Y%m%d") & "000000001 0103" & periodActions1.general.operations(0).debitAccount & "000" & periodActions1.general.operations(0).depositAccount & "                       000                                   2                                           0 0               0               Ñ³Ù³Ó³ÛÝ Ï³ñ·.                            " & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%Y%m%d") & "0                                                  " 
  for i = 0 to 3
				if i <> 1 then 
		    Call CheckDB_FOLDERS(dbo_FOLDERS(i), 1)
				end if
  next
End	Sub

Sub Check_DB_PeriodicActionsAdd()
		Dim i, agrISN
	 'SQL Ստուգում DOCLOG աղյուսակում համար
  Log.Message "SQL Ստուգում DOCLOG աղյուսակում", "", pmNormal, SqlDivideColor
  Call CheckQueryRowCount("DOCLOG", "fISN", periodActions2.fISN, 2)
  Call CheckDB_DOCLOG(periodActions2.fISN, "10", "N", "1", "", 1)
		Call CheckDB_DOCLOG(periodActions2.fISN, "10", "E", "1", "", 1)
  
  'SQL Ստուգում DOCP աղյուսակում  
  Log.Message "SQL Ստուգում DOCP աղյուսակում", "", pmNormal, SqlDivideColor
		agrISN = Get_ColumnValueSQL("DOCP", "fISN", "fPARENTISN = " & periodActions2.fISN & " and fNAME = 'CrPayOrd'")
  Call CheckQueryRowCount("DOCP", "fPARENTISN", periodActions2.fISN, 2)
  Call CheckDB_DOCP("975067592", "Acc     ", periodActions2.fISN, 1)
		Call CheckDB_DOCP(agrISN, "CrPayOrd", periodActions2.fISN, 1)
  
  'SQL Ստուգում DOCS աղյուսակում 
  Log.Message "SQL Ստուգում DOCS աղյուսակում", "", pmNormal, SqlDivideColor
 	fBODY = " ACSBRANCH:P00 ACSDEPART:08 USERID:10 CODE:" & periodActions2.general.agreeN & " NAME:Ð³×³Ëáñ¹ 00009710 ENAME:Client 00009710 CALCALWAYS:0 PERIODICITY:1/103 NONWORKDAYS:0 CLINOT:1 USECLIEMAIL:1 USECLISCH:0 FEEFROMCARD:0 NOTE1:001 NOTE2:193 NOTE3:08 " 
  fBODY = Replace(fBODY, " ", "%")
  Call CheckQueryRowCount("DOCS", "fISN", periodActions2.fISN, 1)
  Call CheckDB_DOCS(periodActions2.fISN, "PPAGR   ", "1", fBODY, 1)
  
  'SQL Ստուգում DOCSG աղյուսակում 
  Log.Message "SQL Ստուգում DOCSG աղյուսակում", "", pmNormal, SqlDivideColor
  Call CheckQueryRowCount("DOCSG", "fISN", periodActions2.fISN, 10)
  
  'SQL Ստուգում FOLDERS աղյուսակում 
  Log.Message "SQL Ստուգում FOLDERS աղյուսակում", "", pmNormal, SqlDivideColor
		agrISN = Get_ColumnValueSQL("DOCP", "fISN", "fPARENTISN = " & periodActions2.fISN & " and fNAME = 'CrPayOrd'")
		for i = 0 to 3
    Set dbo_FOLDERS2(i) = New_DB_FOLDERS()
    dbo_FOLDERS2(i).fISN = periodActions2.fISN
    dbo_FOLDERS2(i).fNAME = "PPAGR   "
				dbo_FOLDERS2(i).fDCBRANCH = "P00"
		  dbo_FOLDERS2(i).fDCDEPART = "08"
  next
		dbo_FOLDERS2(0).fKEY = periodActions2.fISN
		dbo_FOLDERS2(0).fSTATUS = "5"
  dbo_FOLDERS2(0).fFOLDERID = "C.824431882"
  dbo_FOLDERS2(0).fCOM = "ä³ñµ»ñ³Ï³Ý ·áñÍáÕáõÃÛáõÝÝ»ñÇ å³ÛÙ³Ý³·Çñ"
  dbo_FOLDERS2(0).fSPEC = "²Ùë³ÃÇí- " & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%d/%m/%y") & " N- " & periodActions2.general.agreeN & " [ÊÙµ³·ñíáÕ]"
		dbo_FOLDERS2(0).fECOM = "Periodic payments agreement"
		dbo_FOLDERS2(0).fDCBRANCH = ""
		dbo_FOLDERS2(0).fDCDEPART = ""
		dbo_FOLDERS2(1).fKEY = periodActions2.fISN
		dbo_FOLDERS2(1).fSTATUS = "5"
  dbo_FOLDERS2(1).fFOLDERID = "Oper." & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%Y%m%d")
  dbo_FOLDERS2(1).fCOM = "ä³ñµ»ñ³Ï³Ý ·áñÍáÕáõÃÛáõÝÝ»ñÇ å³ÛÙ³Ý³·Çñ"
  dbo_FOLDERS2(1).fSPEC = periodActions2.general.agreeN & "16600                                       0.00   ÊÙµ³·ñíáÕ                                             10Ð³×³Ëáñ¹ 00009710                                                                               ä³ñµ. ·áñÍ. å³ÛÙ³Ý³·Çñ                                                                                                                      "
		dbo_FOLDERS2(1).fECOM = "Periodic payments agreement"
		dbo_FOLDERS2(2).fKEY = periodActions2.general.agreeN
		dbo_FOLDERS2(2).fSTATUS = "1"
  dbo_FOLDERS2(2).fFOLDERID = "PPAYMS"
  dbo_FOLDERS2(2).fCOM = "Ð³×³Ëáñ¹ " & periodActions2.general.client
  dbo_FOLDERS2(2).fSPEC = "1   " & periodActions2.general.client & "10  " & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%Y%m%d") & "000000000 1/103 0 0011                                   0000119308                                   0000000000000000                                                                                                                                                                                                                            0     "
		dbo_FOLDERS2(2).fECOM = "Client " & periodActions2.general.client
		dbo_FOLDERS2(3).fKEY = periodActions2.general.agreeN & "_1"
		dbo_FOLDERS2(3).fSTATUS = "1"
  dbo_FOLDERS2(3).fFOLDERID = "PPAYMSEXT" 
  dbo_FOLDERS2(3).fCOM = "Ð³×³Ëáñ¹ " & periodActions2.general.client
  dbo_FOLDERS2(3).fSPEC = "1   " & periodActions2.general.client & "10  " & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%Y%m%d") & "000000000 1/103 0 0011                                   0000119308                                   00000000000000001 0401" & periodActions2.general.operations(0).debitAccount & "00090000" & periodActions2.general.operations(0).depositAccount & "                      000                                   0                                           0 0               0               Ñ³Ù³Ó³ÛÝ å³ÛÙ³Ý³·ñÇ             " & Left_Align(agrISN, 10) & "000000000                                                  "
  dbo_FOLDERS2(3).fECOM = "Client " & periodActions2.general.client
  Call CheckQueryRowCount("FOLDERS", "fISN", periodActions2.fISN, 4)
  for i = 0 to 3
    Call CheckDB_FOLDERS(dbo_FOLDERS2(i), 1)
  next
End	Sub

Sub Check_DB_Confirm2()
		Dim i, agrISN	
	 'SQL Ստուգում DOCLOG աղյուսակում համար
  Log.Message "SQL Ստուգում DOCLOG աղյուսակում", "", pmNormal, SqlDivideColor
  Call CheckQueryRowCount("DOCLOG", "fISN", periodActions2.fISN, 4)
  Call CheckDB_DOCLOG(periodActions2.fISN, "10", "W", "2", "", 1)
		Call CheckDB_DOCLOG(periodActions2.fISN, "10", "C", "7", "", 1)
  
  'SQL Ստուգում DOCS աղյուսակում 
  Log.Message "SQL Ստուգում DOCS աղյուսակում", "", pmNormal, SqlDivideColor
  Call CheckQueryRowCount("DOCS", "fISN", periodActions2.fISN, 1)
  Call CheckDB_DOCS(periodActions2.fISN, "PPAGR   ", "7", fBODY, 1)
  
  'SQL Ստուգում DOCSG աղյուսակում 
  Log.Message "SQL Ստուգում DOCSG աղյուսակում", "", pmNormal, SqlDivideColor
  Call CheckQueryRowCount("DOCSG", "fISN", periodActions2.fISN, 10)
  
  'SQL Ստուգում FOLDERS աղյուսակում 
  Log.Message "SQL Ստուգում FOLDERS աղյուսակում", "", pmNormal, SqlDivideColor
  Call CheckQueryRowCount("FOLDERS", "fISN", periodActions2.fISN, 3)
		dbo_FOLDERS2(0).fSTATUS = "1"
		agrISN = Get_ColumnValueSQL("DOCP", "fISN", "fPARENTISN = " & periodActions2.fISN & " and fNAME = 'CrPayOrd'")
		dbo_FOLDERS2(0).fSPEC = "²Ùë³ÃÇí- " & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%d/%m/%y") & " N- " & periodActions2.general.agreeN & " [Ð³ëï³ïí³Í]"
		dbo_FOLDERS2(2).fSPEC = "7   " & periodActions2.general.client & "10  " & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%Y%m%d") & "000000000 1/103 0 0011                                   0000119308                                   0000000000000000                                                                                                                                                                                                                            0     "
		dbo_FOLDERS2(3).fSPEC = "7   " & periodActions2.general.client & "10  " & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%Y%m%d") & "000000000 1/103 0 0011                                   0000119308                                   00000000000000001 0401" & periodActions2.general.operations(0).debitAccount & "00090000" & periodActions2.general.operations(0).depositAccount & "                      000                                   0                                           0 0               0               Ñ³Ù³Ó³ÛÝ å³ÛÙ³Ý³·ñÇ             " & Left_Align(agrISN, 10) & "000000000                                                  "
		for i = 0 to 3
				if i <> 1 then 
		    Call CheckDB_FOLDERS(dbo_FOLDERS2(i), 1)
				end if
  next
End	Sub

Sub Check_DB_MakePayment2()
		Dim i, agrISN	
	 'SQL Ստուգում DOCLOG աղյուսակում համար
  Log.Message "SQL Ստուգում DOCLOG աղյուսակում", "", pmNormal, SqlDivideColor
  Call CheckQueryRowCount("DOCLOG", "fISN", periodActions2.fISN, 5)
  Call CheckDB_DOCLOG(periodActions2.fISN, "10", "W", "2", "", 1)
		Call CheckDB_DOCLOG(periodActions2.fISN, "10", "C", "7", "", 1)
  
  'SQL Ստուգում DOCS աղյուսակում 
  Log.Message "SQL Ստուգում DOCS աղյուսակում", "", pmNormal, SqlDivideColor
  Call CheckQueryRowCount("DOCS", "fISN", periodActions2.fISN, 1)
  Call CheckDB_DOCS(periodActions2.fISN, "PPAGR   ", "7", fBODY, 1)
  
  'SQL Ստուգում DOCSG աղյուսակում 
  Log.Message "SQL Ստուգում DOCSG աղյուսակում", "", pmNormal, SqlDivideColor
  Call CheckQueryRowCount("DOCSG", "fISN", periodActions2.fISN, 11)
  
  'SQL Ստուգում FOLDERS աղյուսակում 
  Log.Message "SQL Ստուգում FOLDERS աղյուսակում", "", pmNormal, SqlDivideColor
  Call CheckQueryRowCount("FOLDERS", "fISN", periodActions2.fISN, 3)
		agrISN = Get_ColumnValueSQL("DOCP", "fISN", "fPARENTISN = " & periodActions2.fISN & " and fNAME = 'CrPayOrd'")
		dbo_FOLDERS2(2).fSPEC = "7   " & periodActions2.general.client & "10  " & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%Y%m%d") & "000000000 1/103 0 0011                                   0000119308                                   " & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%Y%m%d") & "00000000                                                                                                                                                                                                                            0     "
		dbo_FOLDERS2(3).fSPEC = "7   " & periodActions2.general.client & "10  " & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%Y%m%d") & "000000000 1/103 0 0011                                   0000119308                                   " & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%Y%m%d") & "000000001 0401" & periodActions2.general.operations(0).debitAccount & "00090000" & periodActions2.general.operations(0).depositAccount & "                      000                                   0                                           0 0               0               Ñ³Ù³Ó³ÛÝ å³ÛÙ³Ý³·ñÇ             " & Left_Align(agrISN, 10) & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%Y%m%d") & "0                                                  " 
		for i = 0 to 3
				if i <> 1 then 
		    Call CheckDB_FOLDERS(dbo_FOLDERS2(i), 1)
				end if
  next
End	Sub

Sub Check_DB_GroupEdit()
		Dim i, dcrid1, dcrid2	
		'SQL Ստուգում DOCLOG աղյուսակում համար
  Log.Message "SQL Ստուգում DOCLOG աղյուսակում", "", pmNormal, SqlDivideColor
		dcrid1 = Get_ColumnValueSQL("DOCLOG", "fDCRID", "fISN = " & periodActions1.fISN & " and fOP = 'M' and fSTATE = '7'")
		dcrid2 = Get_ColumnValueSQL("DOCLOG", "fDCRID", "fISN = " & periodActions2.fISN & " and fOP = 'M' and fSTATE = '7'")
  Call CheckQueryRowCount("DOCLOG", "fISN", periodActions1.fISN, 5)
  Call CheckDB_DOCLOG(periodActions1.fISN, "10", "M", "7", dcrid1 & " ÷á÷áËÙ³Ý Ñ³ÛïÁ ëï»ÕÍí»É ¿", 1)
		Call CheckQueryRowCount("DOCLOG", "fISN", periodActions2.fISN, 6)
		Call CheckDB_DOCLOG(periodActions2.fISN, "10", "M", "7", dcrid2 & " ÷á÷áËÙ³Ý Ñ³ÛïÁ ëï»ÕÍí»É ¿", 1)
End	Sub

Sub Check_DB_ConfirmGroupEdit()
		Dim i, dcrid1, dcrid2, agrISN
	 'SQL Ստուգում DOCLOG աղյուսակում համար
  Log.Message "SQL Ստուգում DOCLOG աղյուսակում", "", pmNormal, SqlDivideColor
		dcrid1 = Get_ColumnValueSQL("DOCLOG", "fDCRID", "fISN = " & periodActions1.fISN & " and fOP = 'M' and fSTATE = '7'")
		dcrid2 = Get_ColumnValueSQL("DOCLOG", "fDCRID", "fISN = " & periodActions2.fISN & " and fOP = 'M' and fSTATE = '7'")
  Call CheckQueryRowCount("DOCLOG", "fISN", periodActions1.fISN, 6)
  Call CheckDB_DOCLOG(periodActions1.fISN, "10", "M", "7", dcrid1 & " ÷á÷áËÙ³Ý Ñ³ÛïÁ ëï»ÕÍí»É ¿", 1)
		Call CheckQueryRowCount("DOCLOG", "fISN", periodActions2.fISN, 7)
		Call CheckDB_DOCLOG(periodActions2.fISN, "10", "M", "7", dcrid2 & " ÷á÷áËÙ³Ý Ñ³ÛïÁ ëï»ÕÍí»É ¿", 1)
  
  'SQL Ստուգում DOCS աղյուսակում 
  Log.Message "SQL Ստուգում DOCS աղյուսակում", "", pmNormal, SqlDivideColor
 	fBODY = " ACSBRANCH:P07 ACSDEPART:06 USERID:10 CODE:" & periodActions1.general.agreeN & " CLICODE:00000208 NAME:Ð³×³Ëáñ¹ 00000208 ENAME:Client 00000208 CALCALWAYS:0 PERIODICITY:3/16 SDAY:2 LDAY:9 NONWORKDAYS:1 CLINOT:0 USECLIEMAIL:0 USECLISCH:0 FEEFROMCARD:0 COMM:Èñ³óáõóÇã ÇÝýá "
  fBODY = Replace(fBODY, " ", "%")
  Call CheckQueryRowCount("DOCS", "fISN", periodActions1.fISN, 1)
  Call CheckDB_DOCS(periodActions1.fISN, "PPAGR   ", "7", fBODY, 1)
		fBODY = " ACSBRANCH:P07 ACSDEPART:06 USERID:10 CODE:" & periodActions2.general.agreeN & " CLICODE:00009710 NAME:Ð³×³Ëáñ¹ 00009710 ENAME:Client 00009710 CALCALWAYS:0 PERIODICITY:3/16 SDAY:2 LDAY:9 NONWORKDAYS:1 CLINOT:0 USECLIEMAIL:1 USECLISCH:0 FEEFROMCARD:0 "
  fBODY = Replace(fBODY, " ", "%")
  Call CheckQueryRowCount("DOCS", "fISN", periodActions2.fISN, 1)
  Call CheckDB_DOCS(periodActions2.fISN, "PPAGR   ", "7", fBODY, 1)
  
  'SQL Ստուգում DOCSG աղյուսակում 
  Log.Message "SQL Ստուգում DOCSG աղյուսակում", "", pmNormal, SqlDivideColor
  Call CheckQueryRowCount("DOCSG", "fISN", periodActions1.fISN, 11)
		Call CheckQueryRowCount("DOCSG", "fISN", periodActions2.fISN, 11)
  
  'SQL Ստուգում FOLDERS աղյուսակում 
  Log.Message "SQL Ստուգում FOLDERS աղյուսակում", "", pmNormal, SqlDivideColor
  Call CheckQueryRowCount("FOLDERS", "fISN", periodActions2.fISN, 3)
		agrISN	=	Get_ColumnValueSQL("DOCP",	"fISN",	"fPARENTISN	=	"	&	periodActions2.fISN	&	"	and	fNAME	=	'CrPayOrd'")
		dbo_FOLDERS2(2).fSPEC = "7   0000971010  " & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%Y%m%d") & "000000000 3/ 16 2 9101                                   00                                           " & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%Y%m%d") & "00000000                                                                                                                                                                                                                            0     "
		dbo_FOLDERS2(3).fSPEC = "7   0000971010  " & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%Y%m%d") & "000000000 3/ 16 2 9101                                   00                                           " & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%Y%m%d") & "000000001 040100971040200000900000123456                      000                                   0                                           0 0               0               Ñ³Ù³Ó³ÛÝ å³ÛÙ³Ý³·ñÇ             " & Left_Align(agrISN, 10) & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%Y%m%d") & "0                                                  "
		dbo_FOLDERS(2).fSPEC = "7   0000020810  " & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%Y%m%d") & "000000000 3/ 16 2 9100                                   00         Èñ³óáõóÇã ÇÝýá                    " & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%Y%m%d") & "00000000                                                                                                                                                                                                                            0     "
		dbo_FOLDERS(3).fSPEC = "7   0000020810  " & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%Y%m%d") & "000000000 3/ 16 2 9100                                   00         Èñ³óáõóÇã ÇÝýá                    " & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%Y%m%d") & "000000001 01030002088020000000020880100                       000                                   2                                           0 0               0               Ñ³Ù³Ó³ÛÝ Ï³ñ·.                            " & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%Y%m%d") & "0                                                  "
  dbo_FOLDERS(2).fDCBRANCH = "P07"
		dbo_FOLDERS(2).fDCDEPART = "06"
		dbo_FOLDERS(3).fDCBRANCH = "P07"
  dbo_FOLDERS(3).fDCDEPART = "06"
		for i = 0 to 3
				if i <> 1 then
				  Call CheckDB_FOLDERS(dbo_FOLDERS(i), 1)
				end if
  next
		for i = 2 to 3
				dbo_FOLDERS2(i).fDCBRANCH = "P07"
		  dbo_FOLDERS2(i).fDCDEPART = "06"
				Call CheckDB_FOLDERS(dbo_FOLDERS2(i), 1)
		next
End	Sub

Sub Check_DB_DeleteDocs()
		'SQL Ստուգում DOCLOG աղյուսակում համար
  Log.Message "SQL Ստուգում DOCLOG աղյուսակում", "", pmNormal, SqlDivideColor
  Call CheckQueryRowCount("DOCLOG", "fISN", periodActions1.fISN, 8)
  Call CheckDB_DOCLOG(periodActions1.fISN, "10", "D", "999", "", 1)
		Call CheckQueryRowCount("DOCLOG", "fISN", periodActions2.fISN, 9)
		Call CheckDB_DOCLOG(periodActions2.fISN, "10", "D", "999", "", 1)
		
		'SQL Ստուգում DOCS աղյուսակում 
  Log.Message "SQL Ստուգում DOCS աղյուսակում", "", pmNormal, SqlDivideColor
 	fBODY = " ACSBRANCH:P07 ACSDEPART:06 USERID:10 CODE:" & periodActions1.general.agreeN & " CLICODE:00000208 NAME:Ð³×³Ëáñ¹ 00000208 ENAME:Client 00000208 CALCALWAYS:0 PERIODICITY:3/16 SDAY:2 LDAY:9 NONWORKDAYS:1 CLINOT:0 USECLIEMAIL:0 USECLISCH:0 FEEFROMCARD:0 COMM:Èñ³óáõóÇã ÇÝýá "
  fBODY = Replace(fBODY, " ", "%")
  Call CheckQueryRowCount("DOCS", "fISN", periodActions1.fISN, 1)
  Call CheckDB_DOCS(periodActions1.fISN, "PPAGR   ", "999", fBODY, 1)
		fBODY = " ACSBRANCH:P07 ACSDEPART:06 USERID:10 CODE:" & periodActions2.general.agreeN & " CLICODE:00009710 NAME:Ð³×³Ëáñ¹ 00009710 ENAME:Client 00009710 CALCALWAYS:0 PERIODICITY:3/16 SDAY:2 LDAY:9 NONWORKDAYS:1 CLINOT:0 USECLIEMAIL:1 USECLISCH:0 FEEFROMCARD:0 "
  fBODY = Replace(fBODY, " ", "%")
  Call CheckQueryRowCount("DOCS", "fISN", periodActions2.fISN, 1)
  Call CheckDB_DOCS(periodActions2.fISN, "PPAGR   ", "999", fBODY, 1)
		
		'SQL Ստուգում DOCSG աղյուսակում 
  Log.Message "SQL Ստուգում DOCSG աղյուսակում", "", pmNormal, SqlDivideColor
  Call CheckQueryRowCount("DOCSG", "fISN", periodActions1.fISN, 11)
		Call CheckQueryRowCount("DOCSG", "fISN", periodActions2.fISN, 10)
		
		'SQL Ստուգում FOLDERS աղյուսակում 
  Log.Message "SQL Ստուգում FOLDERS աղյուսակում", "", pmNormal, SqlDivideColor
  Call CheckQueryRowCount("FOLDERS", "fISN", periodActions1.fISN, 1)
		Call CheckQueryRowCount("FOLDERS", "fISN", periodActions2.fISN, 1)
  dbo_FOLDERS(0).fKEY = periodActions1.fISN
  dbo_FOLDERS(0).fISN = periodActions1.fISN
  dbo_FOLDERS(0).fNAME = "PPAGR   "
  dbo_FOLDERS(0).fSTATUS = "0"
		dbo_FOLDERS(0).fFOLDERID = ".R." & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%Y%m%d")
		dbo_FOLDERS(0).fCOM = ""
		dbo_FOLDERS(0).fSPEC = Left_Align(Get_Compname_DOCLOG(periodActions1.fISN), 16) &  "PERPAYS ARMSOFT                       007  "
		dbo_FOLDERS(0).fECOM = ""
		dbo_FOLDERS(0).fDCBRANCH = "P07"
		dbo_FOLDERS(0).fDCDEPART = "06 "
  Call CheckDB_FOLDERS(dbo_FOLDERS(0), 1)
		dbo_FOLDERS(0).fKEY = periodActions2.fISN
  dbo_FOLDERS(0).fISN = periodActions2.fISN
		Call CheckDB_FOLDERS(dbo_FOLDERS(0), 1)
End	Sub