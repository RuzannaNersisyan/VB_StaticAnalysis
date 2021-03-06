'USEUNIT Library_Common
'USEUNIT Library_Colour
'USEUNIT Library_Contracts 
'USEUNIT Constants
'USEUNIT Library_CheckDB
'USEUNIT Library_Periodic_Actions
'USEUNIT Overlimit_Library
Option Explicit

'Test Case N 170782

Dim sDATE, fDATE, folderName, periodActions, paymentDate, currMonth, currYear
Dim Working_Docs, periodicAct
Dim dbo_FOLDERS(5), fBODY, i, dbo_FOLDERS2(5), PayDate_ForSQL, dbo_FOLDERSOVER(1)
Dim riskiness, groupEdit, chgReqs, accWithOverlimit, overlimitIsn, contractFillter, accParentIsn

Sub Periodic_Actions_Test_Overlimit()
		Call Test_Initialize()
    
		' Ð³Ù³Ï³ñ· Ùáõïù ·áñÍ»É ARMSOFT û·ï³·áñÍáÕáí
		Log.Message "Համակարգ մուտք գործել ARMSOFT օգտագործողով", "", pmNormal, DivideColor
    Call Test_StartUp()
		
		' êï»ÕÍ»É ä³ñµ»ñ³Ï³Ý ·áñÍáÕáõÃÛáõÝÝ»ñÇ å³ÛÙ³Ý³·Çñ
		Log.Message "Ստեղծել Պարբերական գործողությունների պայմանագիր", "", pmNormal, DivideColor
    Call Create_PeriodicActions(folderName, periodActions, "create")
		
		' ä³ñµ»ñ³Ï³Ý ·áñÍáÕáõÃÛáõÝÝ»ñÇ å³ÛÙ³Ý³·ñÇ ëï»ÕÍáõÙÇó Ñ»ïá SQL ëïáõ·áõÙ
		Log.Message "Պարբերական գործողությունների պայմանագրի ստաղծումից հետո SQL ստուգում", "", pmNormal, SqlDivideColor
		Call DB_Initialize()
		Call Check_DB_PeriodicActionsCreation()
		
		' ì³í»ñ³óÝ»É å³ÛÙ³Ý³·ÇñÁ
		Log.Message "Վավերացնել պայմանագիրը", "", pmNormal, DivideColor
		Call GoTo_PeriodicWorkingDocuments(folderName, Working_Docs)
		Call SearchInPttel("frmPttel", 2, periodActions.general.agreeN)
		Call Verify_Periodic_Actions()
		
		BuiltIn.Delay(3000)
		wMDIClient.VBObject("frmPttel").Close
		
		' ä³ÛÙ³Ý³·ñÇ í³í»ñ³óáõÙÇó Ñ»ïá SQL ëïáõ·áõÙ
		Log.Message "Պայմանագրի վավերացումից հետո SQL ստուգում", "", pmNormal, SqlDivideColor
		Call Check_DB_Confirm()
		
		' ä³ñµ»ñ³Ï³Ý ·áñÍáÕáõÃÛáõÝÝ»ñÇ å³ÛÙ³Ý³·ñ»ñ
		Log.Message "Պարբերական գործողությունների պայմանագրեր", "", pmNormal, DivideColor
		Call	Check_PeriodicExisting(folderName & "ä³ñµ»ñ³Ï³Ý ·áñÍáÕáõÃÛáõÝÝ»ñÇ å³ÛÙ³Ý³·ñ»ñ|", periodicAct, periodActions.general.agreeN)
		
		' Î³ï³ñ»É í×³ñáõÙ
		Log.Message "Կատարել վճարում", "", pmNormal, DivideColor
    Call MakePayment_PeriodicActs()
		
		' Î³ï³ñ»É í×³ñáõÙÇó Ñ»ïá SQL ëïáõ·áõÙ
		Log.Message "Կատարել վճարումից հետո SQL ստուգում", "", pmNormal, SqlDivideColor
		Call Check_DB_MakePayment()
		
		' ì×³ñáõÙÝ»ñÇ ¹ÇïáõÙ
		Log.Message "Վճարումների դիտում", "", pmNormal, DivideColor
		Call PaymentView(paymentDate, paymentDate, 1)
		
		BuiltIn.Delay(3000)
		wMDIClient.VBObject("frmPttel_2").Close
		
		BuiltIn.Delay(3000)
		wMDIClient.VBObject("frmPttel").Close
		
		' Øáõïù ·áñÍ»É "¶»ñÍ³Ëë" ²Þî
    Call ChangeWorkspace(c_Overlimit) 
    
    Call wTreeView.DblClickItem("|¶»ñ³Í³Ëë|¶»ñ³Í³Ëë áõÝ»óáÕ Ñ³ßÇíÝ»ñ|")
    BuiltIn.Delay(1000)
    Call Fill_AccWithOverlimit(accWithOverlimit)

		' ¶»ñÍ³ËëÇ µ³óáõÙ (ËÙµ.)
    Log.Message "Գերծախսի բացում (խմբ.)", "", pmNormal, DivideColor
    overlimitIsn = OpenOverimitFromAccount(paymentDate)
		
		' ¶»ñÍ³ËëÇ µ³óáõÙ (ËÙµ.)-Çó Ñ»ïá SQL ëïáõ·áõÙ
		Log.Message "Վճարումների դիտում", "", pmNormal, DivideColor
		Call Check_DB_Overlimit()

		' ä³ÛÙ³Ý³·ñ»ñ ÃÕÃ³å³Ý³ÏáõÙ ÷³ëï³ÃÕÃÇ ³éÏ³ÛáõÃÛ³Ý ëïáõ·áõÙ
		Log.Message "Պայմանագրեր թղթապանակում փաստատթղթի առկայության ստուգում", "", pmNormal, DivideColor
    Call ExistsContract_Filter_Fill("|¶»ñ³Í³Ëë|", ContractFillter, 1)

		'êïáõ·áõÙ "ØÝ³óáñ¹" ëÛ³Ý ³ñÅ»ùÁ
    Call CompareFieldValue("frmPttel", "fAgrRem", "147.10") 
		
		BuiltIn.Delay(3000)
		wMDIClient.VBObject("frmPttel").Close
		
		' æÝç»É è. ¹³ë. ¨ å³Ñáõëï. ïáÏ. ÷á÷áËáõÃÛáõÝÝ»ñÁ
		Log.Message "Ջնջել Ռ. դաս. և պահուստ. տոկ. փոփոխությունները", "", pmNormal, DivideColor
		folderName = "|¶»ñ³Í³Ëë|Üáñ ÷³ëï³Ã., ÃÕÃ³å³Ý³ÏÝ»ñ, Ñ³ßí»ïíáõÃÛáõÝÝ»ñ|¶áñÍáÕáõÃÛáõÝÝ»ñ, ÷á÷áËáõÃÛáõÝÝ»ñ|èÇëÏ³ÛÝáõÃÛáõÝ|"
		riskiness.startDate = paymentDate
		riskiness.endDate = paymentDate
		Call GoTo_AgreementsCommomFilter(folderName, "è.¹³ë. ¨ å³Ñáõëï.ïáÏ. ÷á÷áËáõÃÛáõÝÝ»ñ", riskiness)
		Call SearchAndDelete("frmPttel", 1, "Ð³×³Ëáñ¹ " & periodActions.general.client, "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ")
		
		BuiltIn.Delay(3000)
		wMDIClient.VBObject("frmPttel").Close
		
		'æÝç»É ¶»ñÍ³ËëÇ å³ÛÙ³Ý³·ÇñÁ
    Log.Message "Ջնջել գերծախսի պայմանագիրը", "", pmNormal, DivideColor
    Call ExistsContract_Filter_Fill("|¶»ñ³Í³Ëë|", ContractFillter, 1)
    Call SearchAndDelete("frmPttel", 7, periodActions.general.client, "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ")
		
		BuiltIn.Delay(3000)
		wMDIClient.VBObject("frmPttel").Close
		
		' æÝç»É ä³ñµ»ñ³Ï³Ý ·áñÍáÕáõÃÛáõÝÝ»ñÇ å³ÛÙ³Ý³·ñ»ñÁ
		Log.Message "Ջնջել Պարբերական գործողությունների պայմանագրերը", "", pmNormal, DivideColor
		folderName = "|ä³ñµ»ñ³Ï³Ý ·áñÍáÕáõÃÛáõÝÝ»ñÇ ²Þî|"
		Call ChangeWorkspace(c_PeriodicActions)
		Call	GoTo_PeriodicActionsAgree(folderName & "ä³ñµ»ñ³Ï³Ý ·áñÍáÕáõÃÛáõÝÝ»ñÇ å³ÛÙ³Ý³·ñ»ñ|", periodicAct)
		
		Log.Message "Վճարումների դիտում", "", pmNormal, DivideColor
		Call SearchInPttel("frmPttel", 0, periodActions.general.agreeN)
		BuiltIn.Delay(3000)
		Call PaymentView(paymentDate, paymentDate, 1)
		Call SearchAndDelete("frmPttel_2", 1, periodActions.general.agreeN, "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ")
		
		BuiltIn.Delay(3000)
		wMDIClient.VBObject("frmPttel_2").Close
		
		Call SearchAndDelete("frmPttel", 0, periodActions.general.agreeN, "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ")
		
		BuiltIn.Delay(3000)
		wMDIClient.VBObject("frmPttel").Close
		
		' æÝç»É ä³ñµ»ñ³Ï³Ý ·áñÍáÕáõÃÛáõÝÝ»ñÇ å³ÛÙ³Ý³·ñ»ñÇó Ñ»ïá SQL ëïáõ·áõÙ
		Log.Message "Ջնջել Պարբերական գործողությունների պայմանագրերից հետո SQL ստուգում", "", pmNormal, SqlDivideColor
		Call Check_DB_DeleteDocs()
		
		Call Close_AsBank()    
End Sub

Sub MakePayment_PeriodicActs()
		if Day(Date) = 1 or Day(Date) = 3 or Day(Date) = 7 or Day(Date) = 8 or Day(Date) = 12 or Day(Date) = 16 or Day(Date) = 19 or Day(Date) = 22 or Day(Date) = 24 or Day(Date) = 25 or Day(Date) = 27 or Day(Date) = 30 then 
				Call MakePayment(periodActions.general.startDate, 1, 1)
        if Day(Date) < 10 then
            paymentDate = "0" & Day(Date) & "/" & currMonth & "/" & currYear
            PayDate_ForSQL = currYear + 2000 & currMonth & "0" & Day(Date)
        else 
            paymentDate = Day(Date) & "/" & currMonth & "/" & currYear
            PayDate_ForSQL = currYear + 2000 & currMonth & Day(Date)
        end if
		else 
				select case Day(Date)
						case 2
								paymentDate = "03/" & currMonth & "/" & currYear
								PayDate_ForSQL = currYear + 2000 & currMonth & "03"
								if Weekday(paymentDate) = 1 or Weekday(paymentDate) = 7 then 
										paymentDate = "07/" & currMonth & "/" & currYear
										PayDate_ForSQL = currYear + 2000 & currMonth & "07"
								end if
								Call MakePayment(paymentDate, 1, 1)
						case 4, 5, 6
								paymentDate = "07/" & currMonth & "/" & currYear
								PayDate_ForSQL = currYear + 2000 & currMonth & "07"
								if Weekday(paymentDate) = 7 then 
										paymentDate = "08/" & currMonth & "/" & currYear
										PayDate_ForSQL = currYear + 2000 & currMonth & "08"
								elseif Weekday(paymentDate) = 1 then
										paymentDate = "12/" & currMonth & "/" & currYear
										PayDate_ForSQL = currYear + 2000 & currMonth & "12"
								end if
								Call MakePayment(paymentDate, 1, 1)
						case 9, 10, 11
								paymentDate = "12/" & currMonth & "/" & currYear
								PayDate_ForSQL = currYear + 2000 & currMonth & "12"
								if Weekday(paymentDate) = 1 or Weekday(paymentDate) = 7 then 
											paymentDate = "16/" & currMonth & "/" & currYear
											PayDate_ForSQL = currYear + 2000 & currMonth & "16"
								end if
								Call MakePayment(paymentDate, 1, 1)
						case 13, 14, 15
								paymentDate = "16/" & currMonth & "/" & currYear
								PayDate_ForSQL = currYear + 2000 & currMonth & "16"
								if Weekday(paymentDate) = 1 or Weekday(paymentDate) = 7 then 
											paymentDate = "19/" & currMonth & "/" & currYear
											PayDate_ForSQL = currYear + 2000 & currMonth & "19"
								end if
								Call MakePayment(paymentDate, 1, 1)
						case 17, 18
								paymentDate = "19/" & currMonth & "/" & currYear
								PayDate_ForSQL = currYear + 2000 & currMonth & "19"
								if Weekday(paymentDate) = 1 or Weekday(paymentDate) = 7 then 
											paymentDate = "22/" & currMonth & "/" & currYear
											PayDate_ForSQL = currYear + 2000 & currMonth & "22"
								end if
								Call MakePayment(paymentDate, 1, 1)
						case 20, 21
								paymentDate = "22/" & currMonth & "/" & currYear
								PayDate_ForSQL = currYear + 2000 & currMonth & "22"
								if Weekday(paymentDate) = 1 or Weekday(paymentDate) = 7 then 
											paymentDate = "24/" & currMonth & "/" & currYear
											PayDate_ForSQL = currYear + 2000 & currMonth & "24"
								end if
								Call MakePayment(paymentDate, 1, 1)
						case 23
								paymentDate = "24/" & currMonth & "/" & currYear
								PayDate_ForSQL = currYear + 2000 & currMonth & "24"
								if Weekday(paymentDate) = 7 then 
										paymentDate = "25/" & currMonth & "/" & currYear
										PayDate_ForSQL = currYear + 2000 & currMonth & "25"
								elseif Weekday(paymentDate) = 1 then
										paymentDate = "27/" & currMonth & "/" & currYear
										PayDate_ForSQL = currYear + 2000 & currMonth & "27"
								end if
								Call MakePayment(paymentDate, 1, 1)
						case 26
								paymentDate = "27/" & currMonth & "/" & currYear
								PayDate_ForSQL = currYear + 2000 & currMonth & "27"
								if Weekday(paymentDate) = 1 or Weekday(paymentDate) = 7 then 
											paymentDate = "30/" & currMonth & "/" & currYear
											PayDate_ForSQL = currYear + 2000 & currMonth & "30"
								end if
								Call MakePayment(paymentDate, 1, 1)
						case 28, 29
								paymentDate = "30/" & currMonth & "/" & currYear
								PayDate_ForSQL = currYear + 2000 & currMonth & "30"
								if Month(Date) = 9 or Month(Date) = 10 or Month(Date) = 11 then 
										currMonth = Month(Date) + 1
								elseif Month(Date) = 12 then
										currMonth = "01"
										currYear = currYear + 1
								else
										currMonth = "0" & Month(Date) + 1
								end if
								if Weekday(paymentDate) = 7 then 
										paymentDate = "01/" & currMonth & "/" & currYear
										PayDate_ForSQL = currYear + 2000 & currMonth & "01"
								elseif Weekday(paymentDate) = 1 then
										paymentDate = "03/" & currMonth & "/" & currYear
										PayDate_ForSQL = currYear + 2000 & currMonth & "03"
								end if
								Call MakePayment(paymentDate, 1, 1)
						case 31
						  if Month(Date) = 9 or Month(Date) = 10 or Month(Date) = 11 then 
										currMonth = Month(Date) + 1
								elseif Month(Date) = 12 then
										currMonth = "01"
										currYear = currYear + 1
								else
										currMonth = "0" & Month(Date) + 1
								end if
								Call MakePayment("01/" & currMonth & "/" & currYear, 1, 1)
								paymentDate = "01/" & currMonth & "/" & currYear
								PayDate_ForSQL = currYear + 2000 & currMonth & "01"
								if Weekday(paymentDate) = 1 or Weekday(paymentDate) = 7 then 
											paymentDate = "03/" & currMonth & "/" & currYear
											PayDate_ForSQL = currYear + 2000 & currMonth & "03"
								end if
								Call MakePayment(paymentDate, 1, 1)
				end select 
		end if
End	Sub

Sub Test_StartUp()
		Call Initialize_AsBankQA(sDATE, fDATE)
  Login("ARMSOFT")
		Call ChangeWorkspace(c_PeriodicActions)
End	Sub

Sub Test_Initialize()
		folderName = "|ä³ñµ»ñ³Ï³Ý ·áñÍáÕáõÃÛáõÝÝ»ñÇ ²Þî|"
		
		sDATE = "20030101"
		fDATE = "20240101"
		
		if Month(Date) = 10 or Month(Date) = 11  or Month(Date) = 12 then 
				currMonth = Month(Date)
		else 
				currMonth = "0" & Month(Date)
		end if
		currYear = Year(Date) - 2000
		
		Set periodActions = New_PeriodicActions(1)
		with periodActions
		  .general.office = "P01"
				.general.department = "07"
				.general.performer = "10"
				.general.client = "00000392"
				.general.doInEveryCall = 0
				.general.mounthDays = "1,3,7,8,12,16,19,22,24,25,27,30"
				.general.bypassNonWorkDays = "2"
				.general.overlimit = 1
				.general.opersGridRowCount = 1 
				.general.operations(0).N_Edit = "1"
				.general.operations(0).operType = "01"
				.general.operations(0).calcMethod = "01"
				.general.operations(0).opersAddDoc = true
				.general.operations(0).debitAccount = "00039200100"
				.general.operations(0).depositAccount  = "4670294"
				.general.operations(0).percent = "4"
				.general.operations(0).price = ""
				.general.operations(0).curr = ""
				.general.operations(0).secID = ""
				.general.operations(0).debitAccountName = ""
				.general.operations(0).depositAccountName = ""
				.general.operations(0).minPrice = "3500"
				.general.operations(0).maxPrice = "8000" 
				.general.operations(0).aim = "äÎä-Ç ÷áË³ÝóÙ³Ý Ñ³Ù³ñ"
				.other.informToClient = 0
				.other.note = "015"
				.other.note2 = "193"
				.other.note3 = "08"
				.other.addInfo = ""
		end with
		
		Set Working_Docs = New_PeriodicWorkingDocuments()
		Working_Docs.performers = "10"
		
		Set periodicAct = New_PeriodicActionsAgree()
		periodicAct.performer = "10"
		
		Set chgReqs = New_ChangeRequests()
		
		Set accWithOverlimit = New_AccountsWithOverlimit()
    accWithOverlimit.Curr = "000"
    accWithOverlimit.Client = "00000392"
    accWithOverlimit.AccountMask = "00039200100"
		
		Set contractFillter = New_ContractOverlimit()
    contractFillter.AgreementLevel = "1"
    contractFillter.Curr = "000"
    contractFillter.Client = "00000392"
		
		Set riskiness = New_AgreementsCommomFilter()
		with riskiness
				.onlyChangesExists = true
		end with
		
End Sub

Sub DB_Initialize()
		for i = 0 to 4
      Set dbo_FOLDERS(i) = New_DB_FOLDERS()
      dbo_FOLDERS(i).fISN = periodActions.fISN
      dbo_FOLDERS(i).fNAME = "PPAGR   "
    next
		dbo_FOLDERS(0).fKEY = periodActions.fISN
		dbo_FOLDERS(0).fSTATUS = "5"
    dbo_FOLDERS(0).fFOLDERID = "C.103280"
    dbo_FOLDERS(0).fCOM = "ä³ñµ»ñ³Ï³Ý ·áñÍáÕáõÃÛáõÝÝ»ñÇ å³ÛÙ³Ý³·Çñ"
    dbo_FOLDERS(0).fSPEC = "²Ùë³ÃÇí- " & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%d/%m/%y") & " N- " & periodActions.general.agreeN & " [Üáñ]"
    dbo_FOLDERS(0).fECOM = "Periodic payments agreement"
    dbo_FOLDERS(1).fKEY = periodActions.fISN
    dbo_FOLDERS(1).fSTATUS = "5"
    dbo_FOLDERS(1).fFOLDERID = "Oper." & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%Y%m%d")
    dbo_FOLDERS(1).fCOM = "ä³ñµ»ñ³Ï³Ý ·áñÍáÕáõÃÛáõÝÝ»ñÇ å³ÛÙ³Ý³·Çñ"
    dbo_FOLDERS(1).fSPEC = periodActions.general.agreeN & "16600                                       0.00   Üáñ                                                   10Ð³×³Ëáñ¹ 00000392                                                                               ä³ñµ. ·áñÍ. å³ÛÙ³Ý³·Çñ                                                                                                                      "
    dbo_FOLDERS(1).fECOM = "Periodic payments agreement"
    dbo_FOLDERS(1).fDCBRANCH = "P01"
    dbo_FOLDERS(1).fDCDEPART = "07 "
    dbo_FOLDERS(2).fKEY = periodActions.general.agreeN
    dbo_FOLDERS(2).fSTATUS = "1"
    dbo_FOLDERS(2).fFOLDERID = "PPAYMS"
    dbo_FOLDERS(2).fCOM = "Ð³×³Ëáñ¹ " & periodActions.general.client
    dbo_FOLDERS(2).fSPEC = "1   0000039210  " & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%Y%m%d") & "000000000 0/  0 0 0200                                   0001519308                                   0000000000000000                                                                                                                                                                                                                            11,3,7,8,12,16,19,22,24,25,27,30                   "
    dbo_FOLDERS(2).fECOM = "Client " & periodActions.general.client
    dbo_FOLDERS(2).fDCBRANCH = "P01"
    dbo_FOLDERS(2).fDCDEPART = "07 "
    dbo_FOLDERS(3).fKEY = periodActions.general.agreeN & "_1"
    dbo_FOLDERS(3).fSTATUS = "1"
    dbo_FOLDERS(3).fFOLDERID = "PPAYMSEXT" 
    dbo_FOLDERS(3).fCOM = "Ð³×³Ëáñ¹ " & periodActions.general.client
    dbo_FOLDERS(3).fECOM = "Client " & periodActions.general.client
    dbo_FOLDERS(3).fSPEC = "1   0000039210  " & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%Y%m%d") & "000000000 0/  0 0 0200                                   0001519308                                   00000000000000001 0101000392001000004670294                           000                                   4                                           0 3500            8000            äÎä-Ç ÷áË³ÝóÙ³Ý Ñ³Ù³ñ                     0000000011,3,7,8,12,16,19,22,24,25,27,30                   "
    dbo_FOLDERS(3).fDCBRANCH = "P01"
    dbo_FOLDERS(3).fDCDEPART = "07 "
		
    Set dbo_FOLDERSOVER(0) = New_DB_FOLDERS()
		Set dbo_FOLDERSOVER(1) = New_DB_FOLDERS()
    dbo_FOLDERSOVER(0).fNAME = "Acc     "
    dbo_FOLDERSOVER(0).fSTATUS = "1"
    dbo_FOLDERSOVER(0).fFOLDERID = "C.103280"
    dbo_FOLDERSOVER(0).fCOM = "  Ð³ßÇí"
    dbo_FOLDERSOVER(0).fSPEC = "00039200100  ²ñÅ.- 000  îÇå- 01  Ð/Ð³ßÇí- 3030201   ²Ýí³ÝáõÙ-Ð³×³Ëáñ¹ 00000392"
    dbo_FOLDERSOVER(0).fECOM = "  Account"
    dbo_FOLDERSOVER(1).fNAME = "Acc     "
    dbo_FOLDERSOVER(1).fSTATUS = "0"
    dbo_FOLDERSOVER(1).fFOLDERID = "ACCOVERLIM"
    dbo_FOLDERSOVER(1).fCOM = "Ð³ßÇí"
    dbo_FOLDERSOVER(1).fSPEC = "            0.000000039200001  Ð³×³Ëáñ¹ 00000392                                           147.10      001"
    dbo_FOLDERSOVER(1).fECOM = "Account"
    dbo_FOLDERSOVER(1).fDCBRANCH = "P00"
    dbo_FOLDERSOVER(1).fDCDEPART = "02 "
End	Sub

Sub Check_DB_PeriodicActionsCreation()
		Dim i, agrISN
	 'SQL Ստուգում DOCLOG աղյուսակում համար
    Log.Message "SQL Ստուգում DOCLOG աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCLOG", "fISN", periodActions.fISN, 1)
    Call CheckDB_DOCLOG(periodActions.fISN, "10", "N", "1", "", 1)
  
    'SQL Ստուգում DOCP աղյուսակում  
    Log.Message "SQL Ստուգում DOCP աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCP", "fPARENTISN", periodActions.fISN, 2)
    Call CheckDB_DOCP("105504", "Acc     ", periodActions.fISN, 1)
    Call CheckDB_DOCP("767531023", "Acc     ", periodActions.fISN, 1)
  
    'SQL Ստուգում DOCS աղյուսակում 
    Log.Message "SQL Ստուգում DOCS աղյուսակում", "", pmNormal, SqlDivideColor
    fBODY = " ACSBRANCH:P01 ACSDEPART:07 USERID:10 CODE:" & periodActions.general.agreeN & " CLICODE:00000392 NAME:Ð³×³Ëáñ¹ 00000392 ENAME:Client 00000392 CALCALWAYS:0 DAYSOFMONTH:1,3,7,8,12,16,19,22,24,25,27,30 NONWORKDAYS:2 USEOVERLIMIT:1 CLINOT:0 USECLIEMAIL:0 USECLISCH:0 FEEFROMCARD:0 NOTE1:015 NOTE2:193 NOTE3:08 "
    fBODY = Replace(fBODY, " ", "%")
    Call CheckQueryRowCount("DOCS", "fISN", periodActions.fISN, 1)
    Call CheckDB_DOCS(periodActions.fISN, "PPAGR   ", "1", fBODY, 1)
  
    'SQL Ստուգում DOCSG աղյուսակում 
    Log.Message "SQL Ստուգում DOCSG աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCSG", "fISN", periodActions.fISN, 12)
  
    'SQL Ստուգում FOLDERS աղյուսակում 
    Log.Message "SQL Ստուգում FOLDERS աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("FOLDERS", "fISN", periodActions.fISN, 4)
    for i = 0 to 3
      Call CheckDB_FOLDERS(dbo_FOLDERS(i), 1)
    next
End	Sub

Sub Check_DB_Confirm()
		Dim i, agrISN
	 'SQL Ստուգում DOCLOG աղյուսակում համար
    Log.Message "SQL Ստուգում DOCLOG աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCLOG", "fISN", periodActions.fISN, 3)
    Call CheckDB_DOCLOG(periodActions.fISN, "10", "W", "2", "", 1)
		Call CheckDB_DOCLOG(periodActions.fISN, "10", "C", "7", "", 1)
  
    'SQL Ստուգում DOCS աղյուսակում 
    Log.Message "SQL Ստուգում DOCS աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCS", "fISN", periodActions.fISN, 1)
    Call CheckDB_DOCS(periodActions.fISN, "PPAGR   ", "7", fBODY, 1)
  
    'SQL Ստուգում DOCSG աղյուսակում 
    Log.Message "SQL Ստուգում DOCSG աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCSG", "fISN", periodActions.fISN, 12)
  
    'SQL Ստուգում FOLDERS աղյուսակում 
    Log.Message "SQL Ստուգում FOLDERS աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("FOLDERS", "fISN", periodActions.fISN, 3)
		dbo_FOLDERS(0).fSTATUS = "1"
		dbo_FOLDERS(0).fSPEC = "²Ùë³ÃÇí- " & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%d/%m/%y") & " N- " & periodActions.general.agreeN & " [Ð³ëï³ïí³Í]"
		dbo_FOLDERS(2).fSPEC = "7   0000039210  " & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%Y%m%d") & "000000000 0/  0 0 0200                                   0001519308                                   0000000000000000                                                                                                                                                                                                                            11,3,7,8,12,16,19,22,24,25,27,30                   "
		dbo_FOLDERS(3).fSPEC = "7   0000039210  " & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%Y%m%d") & "000000000 0/  0 0 0200                                   0001519308                                   00000000000000001 0101000392001000004670294                           000                                   4                                           0 3500            8000            äÎä-Ç ÷áË³ÝóÙ³Ý Ñ³Ù³ñ                     0000000011,3,7,8,12,16,19,22,24,25,27,30                   "
    for i = 0 to 3
				if i <> 1 then 
		    Call CheckDB_FOLDERS(dbo_FOLDERS(i), 1)
				end if
    next
End	Sub

Sub Check_DB_MakePayment()
		Dim i, agrISN
	 'SQL Ստուգում DOCLOG աղյուսակում համար
    Log.Message "SQL Ստուգում DOCLOG աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCLOG", "fISN", periodActions.fISN, 4)
		
		'SQL Ստուգում DOCS աղյուսակում 
    Log.Message "SQL Ստուգում DOCS աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCS", "fISN", periodActions.fISN, 1)
    Call CheckDB_DOCS(periodActions.fISN, "PPAGR   ", "7", fBODY, 1)
  
    'SQL Ստուգում DOCSG աղյուսակում 
    Log.Message "SQL Ստուգում DOCSG աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCSG", "fISN", periodActions.fISN, 13)
  
    'SQL Ստուգում FOLDERS աղյուսակում 
    Log.Message "SQL Ստուգում FOLDERS աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("FOLDERS", "fISN", periodActions.fISN, 3)
		dbo_FOLDERS(0).fSTATUS = "1"
		dbo_FOLDERS(0).fSPEC = "²Ùë³ÃÇí- " & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%d/%m/%y") & " N- " & periodActions.general.agreeN & " [Ð³ëï³ïí³Í]"
		dbo_FOLDERS(2).fSPEC = "7   0000039210  " & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%Y%m%d") & "000000000 0/  0 0 0200                                   0001519308                                   " & PayDate_ForSQL & "00000000                                                                                                                                                                                                                            11,3,7,8,12,16,19,22,24,25,27,30                   "
		dbo_FOLDERS(3).fSPEC = "7   0000039210  " & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%Y%m%d") & "000000000 0/  0 0 0200                                   0001519308                                   " & PayDate_ForSQL & "000000001 0101000392001000004670294                           000                                   4                                           0 3500            8000            äÎä-Ç ÷áË³ÝóÙ³Ý Ñ³Ù³ñ                     " & PayDate_ForSQL & "11,3,7,8,12,16,19,22,24,25,27,30                   "
    for i = 0 to 3
				if i <> 1 then 
		    Call CheckDB_FOLDERS(dbo_FOLDERS(i), 1)
				end if
    next
End	Sub

Sub Check_DB_Overlimit()
	 'SQL Ստուգում ACCOUNTS աղյուսակում համար
    Log.Message "SQL Ստուգում ACCOUNTS աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("ACCOUNTS", "fISN", overlimitIsn, 1)
		
		'SQL Ստուգում DOCP աղյուսակում  
    Log.Message "SQL Ստուգում DOCP աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCP", "fISN", overlimitIsn, 4)
		
		'SQL Ստուգում FOLDERS աղյուսակում 
    Log.Message "SQL Ստուգում FOLDERS աղյուսակում", "", pmNormal, SqlDivideColor
		dbo_FOLDERSOVER(0).fISN = overlimitIsn
		dbo_FOLDERSOVER(0).fKEY = overlimitIsn
    Call CheckQueryRowCount("FOLDERS", "fISN", overlimitIsn, 1)
    Call CheckDB_FOLDERS(dbo_FOLDERSOVER(0), 1)
		
		'SQL Ստուգում ACCHIRESTOUNTS աղյուսակում համար
    Log.Message "SQL Ստուգում HIREST աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("HIREST", "fOBJECT", overlimitIsn, 5)
End	Sub

Sub Check_DB_DeleteDocs()
		'SQL Ստուգում DOCLOG աղյուսակում համար
    Log.Message "SQL Ստուգում DOCLOG աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCLOG", "fISN", periodActions.fISN, 6)
    Call CheckDB_DOCLOG(periodActions.fISN, "10", "D", "999", "", 1)
		
		'SQL Ստուգում DOCS աղյուսակում 
    Log.Message "SQL Ստուգում DOCS աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCS", "fISN", periodActions.fISN, 1)
    Call CheckDB_DOCS(periodActions.fISN, "PPAGR   ", "999", fBODY, 1)
		
		'SQL Ստուգում DOCSG աղյուսակում 
    Log.Message "SQL Ստուգում DOCSG աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCSG", "fISN", periodActions.fISN, 13)
		
		'SQL Ստուգում FOLDERS աղյուսակում 
    Log.Message "SQL Ստուգում FOLDERS աղյուսակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("FOLDERS", "fISN", periodActions.fISN, 1)
		Call CheckQueryRowCount("FOLDERS", "fISN", overlimitIsn, 2)
    dbo_FOLDERS(0).fKEY = periodActions.fISN
    dbo_FOLDERS(0).fISN = periodActions.fISN
    dbo_FOLDERS(0).fNAME = "PPAGR   "
    dbo_FOLDERS(0).fSTATUS = "0"
		dbo_FOLDERS(0).fFOLDERID = ".R." & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%Y%m%d")
		dbo_FOLDERS(0).fCOM = ""
		dbo_FOLDERS(0).fSPEC = Left_Align(Get_Compname_DOCLOG(periodActions.fISN), 16) &  "PERPAYS ARMSOFT                       007  "
		dbo_FOLDERS(0).fECOM = ""
		dbo_FOLDERS(0).fDCBRANCH = "P01"
		dbo_FOLDERS(0).fDCDEPART = "07 "
    Call CheckDB_FOLDERS(dbo_FOLDERS(0), 1)
		dbo_FOLDERSOVER(1).fISN = overlimitIsn
		dbo_FOLDERSOVER(1).fKEY = "00039200100"
		Call CheckDB_FOLDERS(dbo_FOLDERSOVER(0), 1)
		Call CheckDB_FOLDERS(dbo_FOLDERSOVER(1), 1)
End	Sub