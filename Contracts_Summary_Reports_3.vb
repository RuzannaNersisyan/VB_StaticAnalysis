'USEUNIT Subsystems_SQL_Library
'USEUNIT Library_Common
'USEUNIT Library_Contracts
'USEUNIT Constants
'USEUNIT Library_Colour
'USEUNIT OLAP_Library
'USEUNIT SWIFT_International_Payorder_Library
Option Explicit

'Test case ID 161658
'Test case ID 161664

Dim folderName, sDATE, fDATE, colName(5), param
Dim rateChanges, subsidyRates, effectiveRateParent, effectiveRate, bankInterests
Dim actualFile1, actualFile2, actualFile3, actualFile4, actualFile5
Dim expectedFile1, expectedFile2, expectedFile3, expectedFile4, expectedFile5
Dim resultFile1, resultFile3, resultFile5

Sub Contracts_Summary_Reports_3(rowLimit)
		' ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|ä³ÛÙ³Ý³·ñ»ñ|²Ù÷á÷|î»Õ³µ³ßËí³Í ÙÇçáóÝ»ñ
		Call Test_Initialize()

		' Համակարգ մուտք գործել ARMSOFT օգտագործողով
		Log.Message "Համակարգ մուտք գործել ARMSOFT օգտագործողով", "", pmNormal, DivideColor
  Call Test_StartUp(rowLimit)
		
		'''''''''''''''''''''''''''''''''''''''''''''''''
		'''''''''''''Տոկոսադրույքի փոփոխություններ''''''''''''''
		
		' Լրացնել Տոկոսադրույքի փոփոխություններ դիալոգային պատուհանը
		Log.Message "Տոկոսադրույքի փոփոխություններ", "", pmNormal, DivideColor
		Call GoTo_AgreementsCommomFilter(folderName, "îáÏáë³¹ñáõÛùÇ ÷á÷áËáõÃÛáõÝÝ»ñ", rateChanges)
		
		if WaitForExecutionProgress() then		
				' Սորտավորել բացված պտտելը
				Call columnSorting(colName, 3, "frmPttel")
				' Ստուգել տողերի քանակը
				Call CheckPttel_RowCount("frmPttel", 21097)
				' Արտահանել Excel
				Call ExportToExcel("frmPttel", actualFile1)
				' Համեմատել Excel ֆայլերը
				Call CompareTwoExcelFiles(actualFile1, expectedFile1, resultFile1)
				' ö³Ï»É åïï»ÉÁ
				Call CloseAllExcelFiles()
				BuiltIn.Delay(3000) 
		  wMDIClient.VBObject("frmPttel").Close
		else																																	
						Log.Error "Can't open pttel window.", "", pmNormal, ErrorColor
		end if
		
		'''''''''''''''''''''''''''''''''''''''''''''''''
		''''''Սուբսիդավորման տոկոսադրույքների փոփոխություններ'''''''
		
		' Լրացնել Սուբսիդավորման տոկոսադրույք դիալոգային պատուհանը
		Log.Message "Սուբսիդավորման տոկոսադրույքների փոփոխություններ", "", pmNormal, DivideColor
		Call GoTo_AgreementsCommomFilter(folderName, "êáõµëÇ¹³íáñÙ³Ý ïáÏáë³¹ñáõÛùÝ»ñÇ ÷á÷áËáõÃÛáõÝÝ»ñ", subsidyRates)
		
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
		''''''''Արդյունավետ/փաստացի տոկոսադրույքներ (Ծնող)'''''''
		
		colName(4) = "fPRNDER"
		
		' Լրացնել Արդյունավետ/փաստացի տոկոսադրույքներ (Ծնող) դիալոգային պատուհանը
		Log.Message "Արդյունավետ/փաստացի տոկոսադրույքներ (Ծնող)", "", pmNormal, DivideColor
		Call GoTo_AgreementsCommomFilter(folderName, "²ñ¹ÛáõÝ³í»ï\÷³ëï³óÇ ïáÏáë³¹ñáõÛùÝ»ñ (ÌÝáÕ)", effectiveRateParent)
		
		if WaitForExecutionProgress() then		
				'êáñï³íáñ»É µ³óí³Í åïï»ÉÁ
				Call columnSorting(colName, 5, "frmPttel")
				' Արտահանել Excel
				Call ExportToExcel("frmPttel", actualFile3)
				' Ստուգել տողերի քանակը
				Call CheckPttel_RowCount("frmPttel", 24372)
				' Համեմատել Excel ֆայլերը
				Call CompareTwoExcelFiles(actualFile3, expectedFile3, resultFile3)
				' ö³Ï»É բոլոր Excel ֆայլերը
				Call CloseAllExcelFiles()
				' ö³Ï»É åïï»ÉÁ
				BuiltIn.Delay(3000) 
		  wMDIClient.VBObject("frmPttel").Close
		else																																	
						Log.Error "Can't open pttel window.", "", pmNormal, ErrorColor
		end if
		
		'''''''''''''''''''''''''''''''''''''''''''''''''
		'''''''''Արդյունավետ/փաստացի տոկոսադրույքներ''''''''''''
		
		' Լրացնել Արդյունավետ/փաստացի տոկոսադրույքներ դիալոգային պատուհանը
		Log.Message "Արդյունավետ/փաստացի տոկոսադրույքներ", "", pmNormal, DivideColor
		Call GoTo_AgreementsCommomFilter(folderName, "²ñ¹ÛáõÝ³í»ï\÷³ëï³óÇ ïáÏáë³¹ñáõÛùÝ»ñ", effectiveRate)
		
		if WaitForExecutionProgress() then		
				'êáñï³íáñ»É µ³óí³Í åïï»ÉÁ
				Call columnSorting(colName, 4, "frmPttel")
				' Արտահանել, որպես txt ֆայլ
				Call ExportToTXTFromPttel("frmPttel", actualFile4)
				' Ստուգել տողերի քանակը
				Call CheckPttel_RowCount("frmPttel", 1)
				' Համեմատել txt ֆայլերը
				Call Compare_Files(actualFile4, expectedFile4, param)
				' ö³Ï»É åïï»ÉÁ
				BuiltIn.Delay(3000) 
		  wMDIClient.VBObject("frmPttel").Close
		else																																	
						Log.Error "Can't open pttel window.", "", pmNormal, ErrorColor
		end if
		
		'''''''''''''''''''''''''''''''''''''''''''''''''
		''''''''''Բանկի արդյունավետ տոկոսադրույքներ'''''''''''''
		
		colName(4) = "fEFFRATE"
		
		' Լրացնել Բանկի արդյունավետ տոկոսադրույքներ դիալոգային պատուհանը
		Log.Message "Բանկի արդյունավետ տոկոսադրույքներ", "", pmNormal, DivideColor
		Call GoTo_AgreementsCommomFilter(folderName, "´³ÝÏÇ ³ñ¹ÛáõÝ³í»ï ïáÏáë³¹ñáõÛùÝ»ñ", bankInterests)
		
		if WaitForExecutionProgress() then		
				'êáñï³íáñ»É µ³óí³Í åïï»ÉÁ
				Call columnSorting(colName, 5, "frmPttel")
				' Արտահանել Excel
				Call ExportToExcel("frmPttel", actualFile5)
				' Ստուգել տողերի քանակը
				Call CheckPttel_RowCount("frmPttel", 17528)
				' Համեմատել Excel ֆայլերը
				Call CompareTwoExcelFiles(actualFile5, expectedFile5, resultFile5)
				'ö³Ï»É բոլոր Excel ֆայլերը
				Call CloseAllExcelFiles()
				BuiltIn.Delay(3000) 
				'ö³Ï»É åïï»ÉÁ
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
		folderName = "|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|ä³ÛÙ³Ý³·ñ»ñ|²Ù÷á÷|î»Õ³µ³ßËí³Í ÙÇçáóÝ»ñ|"
	
		sDATE = "20030101"
		fDATE = "20260101"  
		
		colName(0) = "fKEY"
		colName(1) = "fCOM"
		colName(2) = "fDATE"
		colName(3) = "fSUID"
		
		' îáÏáë³¹ñáõÛùÇ ÷á÷áËáõÃÛáõÝÝ»ñ
		expectedFile1 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest3\Expected\expectedFile1.xlsx"
		' êáõµëÇ¹³íáñÙ³Ý ïáÏáë³¹ñáõÛùÝ»ñÇ ÷á÷áËáõÃÛáõÝÝ»ñ
		expectedFile2 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest3\Expected\expectedFile2.txt"
  ' ²ñ¹ÛáõÝ³í»ï\÷³ëï³óÇ ïáÏáë³¹ñáõÛùÝ»ñ (ÌÝáÕ)
		expectedFile3 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest3\Expected\expectedFile3.xlsx"
		' ²ñ¹ÛáõÝ³í»ï\÷³ëï³óÇ ïáÏáë³¹ñáõÛùÝ»ñ
		expectedFile4 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest3\Expected\expectedFile4.txt"
		' ´³ÝÏÇ ³ñ¹ÛáõÝ³í»ï ïáÏáë³¹ñáõÛùÝ»ñ
		expectedFile5 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest3\Expected\expectedFile5.xlsx"
		
  ' îáÏáë³¹ñáõÛùÇ ÷á÷áËáõÃÛáõÝÝ»ñ
		actualFile1 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest3\Actual\actualFile1.xlsx"
		' êáõµëÇ¹³íáñÙ³Ý ïáÏáë³¹ñáõÛùÝ»ñÇ ÷á÷áËáõÃÛáõÝÝ»ñ
		actualFile2 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest3\Actual\actualFile2.txt"
  ' ²ñ¹ÛáõÝ³í»ï\÷³ëï³óÇ ïáÏáë³¹ñáõÛùÝ»ñ (ÌÝáÕ)
		actualFile3 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest3\Actual\actualFile3.xlsx"
		' ²ñ¹ÛáõÝ³í»ï\÷³ëï³óÇ ïáÏáë³¹ñáõÛùÝ»ñ
		actualFile4 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest3\Actual\actualFile4.txt"
		' ´³ÝÏÇ ³ñ¹ÛáõÝ³í»ï ïáÏáë³¹ñáõÛùÝ»ñ
		actualFile5 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest3\Actual\actualFile5.xlsx"
		
  ' îáÏáë³¹ñáõÛùÇ ÷á÷áËáõÃÛáõÝÝ»ñ
		resultFile1 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest3\Result\resultFile1.xlsx"
		' ²ñ¹ÛáõÝ³í»ï\÷³ëï³óÇ ïáÏáë³¹ñáõÛùÝ»ñ (ÌÝáÕ)
		resultFile3 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest3\Result\resultFile3.xlsx"
		' ´³ÝÏÇ ³ñ¹ÛáõÝ³í»ï ïáÏáë³¹ñáõÛùÝ»ñ
		resultFile5 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest3\Result\resultFile5.xlsx"
		
  ' îáÏáë³¹ñáõÛùÇ ÷á÷áËáõÃÛáõÝÝ»ñ
		Set rateChanges = New_AgreementsCommomFilter()
		rateChanges.onlyChangesExists = true
		
		' êáõµëÇ¹³íáñÙ³Ý ïáÏáë³¹ñáõÛùÝ»ñÇ ÷á÷áËáõÃÛáõÝÝ»ñ
		Set subsidyRates = New_AgreementsCommomFilter()
		with subsidyRates
				.startDate = "15/06/17"
				.endDate = "15/06/17"
				.agreeN = "TV22119"
				.performer = "253"
				.note = "005"
				.note2 = "06"
				.note3 = "03"
				.agreeOffice = "P00"
				.agreeSection = "08"
				.accessType = "C11"
				.onlyChangesExists = true
		end with
		
		' ²ñ¹ÛáõÝ³í»ï\÷³ëï³óÇ ïáÏáë³¹ñáõÛùÝ»ñ (ÌÝáÕ)
		Set effectiveRateParent = New_AgreementsCommomFilter()
		effectiveRateParent.onlyChangesExists = true
		
		' ²ñ¹ÛáõÝ³í»ï\÷³ëï³óÇ ïáÏáë³¹ñáõÛùÝ»ñ
		Set effectiveRate = New_AgreementsCommomFilter()
		with effectiveRate
				.startDate = "17/09/13"
				.endDate = "17/09/13"
				.agreeN = "TO5697"
				.performer = "31"
				.note = "00"
				.note3 = "03"
				.agreeOffice = "P00"
				.agreeSection = "061"
				.accessType = "C31"
				.onlyChangesExists = true
				.onlyChanges = 1
		end with
		
		' ´³ÝÏÇ ³ñ¹ÛáõÝ³í»ï ïáÏáë³¹ñáõÛùÝ»ñ
		Set bankInterests = New_AgreementsCommomFilter()
		bankInterests.onlyChangesExists = true
End Sub
