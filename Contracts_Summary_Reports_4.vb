'USEUNIT Subsystems_SQL_Library
'USEUNIT Library_Common
'USEUNIT Library_Contracts
'USEUNIT Constants
'USEUNIT Library_Colour
'USEUNIT OLAP_Library
'USEUNIT SWIFT_International_Payorder_Library
Option Explicit

'Test case ID 161646
'Test case ID 161657

Dim folderName, sDATE, fDATE, colName(5), param
Dim limitChanges, interestsCalcDates, priceAdjustDates, recalculateRates, deliveryDates
Dim actualFile1, actualFile2, actualFile3, actualFile4, actualFile5
Dim expectedFile1, expectedFile2, expectedFile3, expectedFile4, expectedFile5
Dim resultFile1, resultFile3

Sub Contracts_Summary_Reports_4(rowLimit)
		' ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|ä³ÛÙ³Ý³·ñ»ñ|²Ù÷á÷|î»Õ³µ³ßËí³Í ÙÇçáóÝ»ñ|²ÛÉ
		Call Test_Initialize()

		' Համակարգ մուտք գործել ARMSOFT օգտագործողով
		Log.Message "Համակարգ մուտք գործել ARMSOFT օգտագործողով", "", pmNormal, DivideColor
  Call Test_StartUp(rowLimit)
		
		'''''''''''''''''''''''''''''''''''''''''''''''''
		'''''''''''Սահմանաչափերի փոփոխություններ''''''''''''''
		
		colName(4) = "fLIM"
		
		' Լրացնել Սահմանաչափերի փոփոխություններ դիալոգային պատուհանը
		Log.Message "Սահմանաչափերի փոփոխություններ", "", pmNormal, DivideColor
		Call GoTo_AgreementsCommomFilter(folderName, "ê³ÑÙ³Ý³ã³÷»ñÇ ÷á÷áËáõÃÛáõÝÝ»ñ", limitChanges)
		
		if WaitForExecutionProgress() then		
				'êáñï³íáñ»É µ³óí³Í åïï»ÉÁ
				Call columnSorting(colName, 5, "frmPttel")
				' Արտահանել Excel
				Call ExportToExcel("frmPttel", actualFile1)
				' Ստուգել տողերի քանակը
				Call CheckPttel_RowCount("frmPttel", 7550)
				' Համեմատել Excel ֆայլերը
				Call CompareTwoExcelFiles(actualFile1, expectedFile1, resultFile1)
				' ö³Ï»É բոլոր Excel ֆայլերը
				Call CloseAllExcelFiles() 
				' ö³Ï»É åïï»ÉÁ
				BuiltIn.Delay(3000)
		  wMDIClient.VBObject("frmPttel").Close
		else																																	
						Log.Error "Can't open pttel window.", "", pmNormal, ErrorColor
		end if
		
		'''''''''''''''''''''''''''''''''''''''''''''''''
		'''''''''''''Տոկոսների հաշվարկման ամսաթվեր''''''''''''
		
		' Լրացնել Հաշվարկման ամսաթվեր դիալոգային պատուհանը
		Log.Message "Տոկոսների հաշվարկման ամսաթվեր", "", pmNormal, DivideColor
		Call GoTo_AgreementsCommomFilter(folderName, "îáÏáëÝ»ñÇ Ñ³ßí³ñÏÙ³Ý ³Ùë³Ãí»ñ", interestsCalcDates)
		
		if WaitForExecutionProgress() then		
				' Արտահանել, որպես txt ֆայլ
				Call ExportToTXTFromPttel("frmPttel", actualFile2)
				' Ստուգել տողերի քանակը
				Call CheckPttel_RowCount("frmPttel", 1)
				' Համեմատել txt ֆայլերը
				Call Compare_Files(actualFile2, expectedFile2, param)
				' ö³Ï»É åïï»ÉÁ
				BuiltIn.Delay(3000) 
		  wMDIClient.VBObject("frmPttel").Close
		else																																	
						Log.Error "Can't open pttel window.", "", pmNormal, ErrorColor
		end if
		
		'''''''''''''''''''''''''''''''''''''''''''''''''
		''''''''''''''''Գնի ճշտման ամսաթվեր'''''''''''''''''
		
		colName(4) = "fMDC"
		
		' Լրացնել Գնի ճշտման ամսաթվեր դիալոգային պատուհանը
		Log.Message "Գնի ճշտման ամսաթվեր", "", pmNormal, DivideColor
		Call GoTo_AgreementsCommomFilter(folderName, "¶ÝÇ ×ßïÙ³Ý ³Ùë³Ãí»ñ", priceAdjustDates)
		
		if WaitForExecutionProgress() then		
				'êáñï³íáñ»É µ³óí³Í åïï»ÉÁ
				Call columnSorting(colName, 5, "frmPttel")
				' Արտահանել Excel
				Call ExportToExcel("frmPttel", actualFile3)
				' Ստուգել տողերի քանակը
				Call CheckPttel_RowCount("frmPttel", 51218)
				' Համեմատել Excel ֆայլերը
				Call CompareTwoExcelFiles(actualFile3, expectedFile3, resultFile3)
				' ö³Ï»É բոլոր Excel ֆայլերը
				Call CloseAllExcelFiles() 
				'ö³Ï»É åïï»ÉÁ
				BuiltIn.Delay(3000)
		  wMDIClient.VBObject("frmPttel").Close
		else																																	
						Log.Error "Can't open pttel window.", "", pmNormal, ErrorColor
		end if
		
		'''''''''''''''''''''''''''''''''''''''''''''''''
		'''''''''''''Վերահաշվարկի տոկոսադրույքներ''''''''''''''
		
		' Լրացնել Վերահաշվարկի տոկոսադրույքներ դիալոգային պատուհանը
		Log.Message "Վերահաշվարկի տոկոսադրույքներ", "", pmNormal, DivideColor
		Call GoTo_AgreementsCommomFilter(folderName, "ì»ñ³Ñ³ßí³ñÏÇ ïáÏáë³¹ñáõÛùÝ»ñ", recalculateRates)
		
		if WaitForExecutionProgress() then		
				' Արտահանել, որպես txt ֆայլ
				Call ExportToTXTFromPttel("frmPttel", actualFile4)
				' Ստուգել տողերի քանակը
				Call CheckPttel_RowCount("frmPttel", 1)
				' Համեմատել txt ֆայլերը
				Call Compare_Files(actualFile4, expectedFile4, param)
				'ö³Ï»É åïï»ÉÁ
				BuiltIn.Delay(3000) 
		  wMDIClient.VBObject("frmPttel").Close
		else																																	
						Log.Error "Can't open pttel window.", "", pmNormal, ErrorColor
		end if
		
		'''''''''''''''''''''''''''''''''''''''''''''''''
		''''''''''''''''Տրամադրումների ամսաթվեր''''''''''''''
		
		' Լրացնել Տրամադրումների ամսաթվեր դիալոգային պատուհանը
		Log.Message "Տրամադրումների ամսաթվեր", "", pmNormal, DivideColor
		Call GoTo_AgreementsCommomFilter(folderName, "îñ³Ù³¹ñáõÙÝ»ñÇ ³Ùë³Ãí»ñ", deliveryDates)
		
		if WaitForExecutionProgress() then		
				' Արտահանել, որպես txt ֆայլ
				Call ExportToTXTFromPttel("frmPttel", actualFile5)
				' Ստուգել տողերի քանակը
				Call CheckPttel_RowCount("frmPttel", 1)
				' Համեմատել txt ֆայլերը
				Call Compare_Files(actualFile5, expectedFile5, param)
				'ö³Ï»É åïï»ÉÁ
				BuiltIn.Delay(3000) 
		  wMDIClient.VBObject("frmPttel").Close
		else																																	
						Log.Error "Can't open pttel window.", "", pmNormal, ErrorColor
		end if
		
		Call Close_AsBank()		
End	Sub

Sub Test_StartUp(rowLimit)
		Call Initialize_AsBank("bank_Report", sDATE, fDATE)
  Login("ARMSOFT")
		Call SaveRAM_RowsLimit(rowLimit)
		Call ChangeWorkspace(c_Subsystems)
End	Sub

Sub Test_Initialize()
		folderName = "|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|ä³ÛÙ³Ý³·ñ»ñ|²Ù÷á÷|î»Õ³µ³ßËí³Í ÙÇçáóÝ»ñ|²ÛÉ|"
	
		sDATE = "20030101"
		fDATE = "20260101"  
		
		colName(0) = "fKEY"
		colName(1) = "fCOM"
		colName(2) = "fDATE"
		colName(3) = "fSUID"
		
		' ê³ÑÙ³Ý³ã³÷»ñÇ ÷á÷áËáõÃÛáõÝÝ»ñ
		expectedFile1 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest4\Expected\expectedFile1.xlsx"
		' îáÏáëÝ»ñÇ Ñ³ßí³ñÏÙ³Ý ³Ùë³Ãí»ñ
		expectedFile2 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest4\Expected\expectedFile2.txt"
  ' ¶ÝÇ ×ßïÙ³Ý ³Ùë³Ãí»ñ
		expectedFile3 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest4\Expected\expectedFile3.xlsx"
		' ì»ñ³Ñ³ßí³ñÏÇ ïáÏáë³¹ñáõÛùÝ»ñ
		expectedFile4 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest4\Expected\expectedFile4.txt"
		' îñ³Ù³¹ñáõÙÝ»ñÇ ³Ùë³Ãí»ñ
		expectedFile5 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest4\Expected\expectedFile5.txt"
		
  ' ê³ÑÙ³Ý³ã³÷»ñÇ ÷á÷áËáõÃÛáõÝÝ»ñ
		actualFile1 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest4\Actual\actualFile1.xlsx"
		' îáÏáëÝ»ñÇ Ñ³ßí³ñÏÙ³Ý ³Ùë³Ãí»ñ
		actualFile2 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest4\Actual\actualFile2.txt"
  ' ¶ÝÇ ×ßïÙ³Ý ³Ùë³Ãí»ñ
		actualFile3 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest4\Actual\actualFile3.xlsx"
		' ì»ñ³Ñ³ßí³ñÏÇ ïáÏáë³¹ñáõÛùÝ»ñ
		actualFile4 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest4\Actual\actualFile4.txt"
		' îñ³Ù³¹ñáõÙÝ»ñÇ ³Ùë³Ãí»ñ
		actualFile5 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest4\Actual\actualFile5.txt"
		
  ' ê³ÑÙ³Ý³ã³÷»ñÇ ÷á÷áËáõÃÛáõÝÝ»ñ
		resultFile1 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest4\Result\resultFile1.xlsx"
		' ¶ÝÇ ×ßïÙ³Ý ³Ùë³Ãí»ñ
		resultFile3 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest4\Result\resultFile3.xlsx"
		
  ' ê³ÑÙ³Ý³ã³÷»ñÇ ÷á÷áËáõÃÛáõÝÝ»ñ
		Set limitChanges = New_AgreementsCommomFilter()
		limitChanges.onlyChangesExists = true
		
		' îáÏáëÝ»ñÇ Ñ³ßí³ñÏÙ³Ý ³Ùë³Ãí»ñ
		Set interestsCalcDates = New_AgreementsCommomFilter()
		with interestsCalcDates
				.startDate = "25/10/13"
				.endDate = "25/10/13"
				.agreeN = "TV8143"
				.performer = "17"
				.note = "00"
				.note2 = "00"
				.note3 = "01"
				.agreeOffice = "P00"
				.agreeSection = "05"
				.accessType = "C11"
		end with
		
		' ¶ÝÇ ×ßïÙ³Ý ³Ùë³Ãí»ñ
		Set priceAdjustDates = New_AgreementsCommomFilter()
		
		' ì»ñ³Ñ³ßí³ñÏÇ ïáÏáë³¹ñáõÛùÝ»ñ
		Set recalculateRates = New_AgreementsCommomFilter()
		with recalculateRates
				.startDate = "21/07/09"
				.endDate = "21/07/09"
				.agreeN = "TML22070901"
				.performer = "13"
				.agreeOffice = "P00"
				.agreeSection = "02"
				.accessType = "C21"
				.onlyChangesExists = true
		end with	
		
		' îñ³Ù³¹ñáõÙÝ»ñÇ ³Ùë³Ãí»ñ
		Set deliveryDates = New_AgreementsCommomFilter()
End Sub
