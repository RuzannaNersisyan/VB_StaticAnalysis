'USEUNIT Subsystems_SQL_Library
'USEUNIT Library_Common
'USEUNIT Library_Contracts
'USEUNIT Constants
'USEUNIT Library_Colour
'USEUNIT OLAP_Library
'USEUNIT SWIFT_International_Payorder_Library
Option Explicit

'Test case ID 161457
'Test case ID 161487

Dim folderName, sDATE, fDATE, colName(5), param
Dim contract1, operations1, agreeRepayTerms1, interestRepayTerms1, riskiness1, objRisk1
Dim actualFile1, actualFile2, actualFile3, actualFile4, actualFile5, actualFile6
Dim expectedFile1, expectedFile2, expectedFile3, expectedFile4, expectedFile5, expectedFile6
Dim resultFile1, resultFile2, resultFile3, resultFile4, resultFile5, resultFile6

Sub Contracts_Summary_Reports_1(rowLimit)
		' ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|ä³ÛÙ³Ý³·ñ»ñ|²Ù÷á÷|î»Õ³µ³ßËí³Í ÙÇçáóÝ»ñ
		Call Test_Initialize()

		' Համակարգ մուտք գործել ARMSOFT օգտագործողով
		Log.Message "Համակարգ մուտք գործել ARMSOFT օգտագործողով", "", pmNormal, DivideColor
  Call Test_StartUp(rowLimit)
		
		'''''''''''''''''''''''''''''''''''''''''''''''''
		''''''''''''''''''ä³ÛÙ³Ý³·ñ»ñ''''''''''''''''''''
		
		' Լրացնել Պայմանագրեր դիալոգային պատուհանը
		Log.Message "Պայմանագրեր 1 տարբերակ", "", pmNormal, DivideColor
		Call GoTo_Contracts(folderName, contract1)
		
		if WaitForExecutionProgress() then		
				' êáñï³íáñ»É µ³óí³Í åïï»ÉÁ
				Call columnSorting(colName, 3, "frmPttel")
				' Ստուգել տողերի քանակը
				Call CheckPttel_RowCount("frmPttel", 23760)
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
		'''''''''''''''¶áñÍáÕáõÃÛáõÝÝ»ñ''''''''''''''''''
		
		' Լրացնել Գործողությունների դիտում դիալոգային պատուհանը
		Log.Message "Գործողություններ 1 տարբերակ", "", pmNormal, DivideColor
		Call GoTo_AllocFundsOperations(folderName, operations1)
		
		' ä³Ñå³Ý»É Excel
		Call SaveExcelFile(actualFile2)
		' Համեմատել Excel ֆայլերը
		Call CompareTwoExcelFiles(actualFile2, expectedFile2, resultFile2)
		' ö³Ï»É Excel ֆայլերը
		Call CloseAllExcelFiles()
		
		'''''''''''''''''''''''''''''''''''''''''''''''''
		''''ä³ÛÙ. Ù³ñÙ³Ý(í»ñ³ýÇÝ³Ýë³íáñÙ³Ý) Å³ÙÏ»ïÝ»ñ''''
		
		' Լրացնել Ժամկետներ դիալոգային պատուհանը
		Log.Message "Պայմ. մարման(վերաֆինանսավորման) ժամկետներ 1 տարբերակ", "", pmNormal, DivideColor
		Call GoTo_AgreementsCommomFilter(folderName, "Ä³ÙÏ»ïÝ»ñ|ä³ÛÙ.Ù³ñÙ³Ý(í»ñ³ýÇÝ³Ýë³íáñÙ³Ý) Å³ÙÏ»ïÝ»ñ", agreeRepayTerms1)
		
		if WaitForExecutionProgress() then		
				' êáñï³íáñ»É µ³óí³Í åïï»ÉÁ
				Call columnSorting(colName, 5, "frmPttel")
				' Ստուգել տողերի քանակը
				Call CheckPttel_RowCount("frmPttel", 27646)
				' Արտահանել Excel
				Call ExportToExcel("frmPttel", actualFile3)
				' Համեմատել Excel ֆայլերը
				Call CompareTwoExcelFiles(actualFile3, expectedFile3, resultFile3)
				' ö³Ï»É åïï»ÉÁ
				Call CloseAllExcelFiles()
				BuiltIn.Delay(3000) 
		  wMDIClient.VBObject("frmPttel").Close
		else																																	
						Log.Error "Can't open pttel window.", "", pmNormal, ErrorColor
		end if
		
		'''''''''''''''''''''''''''''''''''''''''''''''''
		''''''''''''îáÏáëÝ»ñÇ Ù³ñÙ³Ý Å³ÙÏ»ïÝ»ñ'''''''''''
		
		' Լրացնել Ժամկետներ դիալոգային պատուհանը
		Log.Message "Տոկոսների մարման ժամկետներ 1 տարբերակ", "", pmNormal, DivideColor
		Call GoTo_AgreementsCommomFilter(folderName, "Ä³ÙÏ»ïÝ»ñ|îáÏáëÝ»ñÇ Ù³ñÙ³Ý Å³ÙÏ»ïÝ»ñ", interestRepayTerms1)
		
		if WaitForExecutionProgress() then		
				' êáñï³íáñ»É µ³óí³Í åïï»ÉÁ
				Call columnSorting(colName, 5, "frmPttel")
				' Ստուգել տողերի քանակը
				Call CheckPttel_RowCount("frmPttel", 24621)
				' Արտահանել Excel
				Call ExportToExcel("frmPttel", actualFile4)
				' Համեմատել Excel ֆայլերը
				Call CompareTwoExcelFiles(actualFile4, expectedFile4, resultFile4)
				'ö³Ï»É åïï»ÉÁ
				Call CloseAllExcelFiles()
				BuiltIn.Delay(3000) 
		  wMDIClient.VBObject("frmPttel").Close
		else																																	
						Log.Error "Can't open pttel window.", "", pmNormal, ErrorColor
		end if
		
		'''''''''''''''''''''''''''''''''''''''''''''''''
		'''''è. ¹³ë. ¨ å³Ñáõëï. ïáÏ. ÷á÷áËáõÃÛáõÝÝ»ñ'''''
		
		colName(1) = "fCRRISK"
		colName(2) = "fPRREZ"
		
		' Լրացնել Ռիսկի դասիչներ և պահուստավորման տոկոսներ դիալոգային պատուհանը
		Log.Message "Ռիսկի դասիչներ և պահուստավորման տոկոսներ 1 տարբերակ", "", pmNormal, DivideColor
		Call GoTo_AgreementsCommomFilter(folderName, "èÇëÏ³ÛÝáõÃÛáõÝ|è.¹³ë. ¨ å³Ñáõëï.ïáÏ. ÷á÷áËáõÃÛáõÝÝ»ñ", riskiness1)
		
		if WaitForExecutionProgress() then		
				colName(2) = "fPERRES"
				' êáñï³íáñ»É µ³óí³Í åïï»ÉÁ
				Call columnSorting(colName, 5, "frmPttel")
				' Ստուգել տողերի քանակը
				Call CheckPttel_RowCount("frmPttel", 27064)
				' Արտահանել Excel
				Call ExportToExcel("frmPttel", actualFile5)
				' Համեմատել Excel ֆայլերը
				Call CompareTwoExcelFiles(actualFile5, expectedFile5, resultFile5)
				'ö³Ï»É åïï»ÉÁ
				Call CloseAllExcelFiles()
				BuiltIn.Delay(3000) 
		  wMDIClient.VBObject("frmPttel").Close
		else																																	
						Log.Error "Can't open pttel window.", "", pmNormal, ErrorColor
		end if
		
		'''''''''''''''''''''''''''''''''''''''''''''''''
		'''''''úµÛ»ÏïÇí éÇëÏÇ ¹³ëÇãÇ ÷á÷ËáõÃÛáõÝÝ»ñ''''''
		
		colName(1) = "fCOM"
		
		' Լրացնել Օբյեկտիվ ռիսկի դասիչ դիալոգային պատուհանը
		Log.Message "Օբյեկտիվ ռիսկի դասիչ 1 տարբերակ", "", pmNormal, DivideColor
		Call GoTo_AgreementsCommomFilter(folderName, "èÇëÏ³ÛÝáõÃÛáõÝ|úµÛ»ÏïÇí éÇëÏÇ ¹³ëÇãÇ ÷á÷áËáõÃÛáõÝÝ»ñ", objRisk1)
		
		if WaitForExecutionProgress() then		
				' êáñï³íáñ»É µ³óí³Í åïï»ÉÁ
				Call columnSorting(colName, 5, "frmPttel")
				' Ստուգել տողերի քանակը
				Call CheckPttel_RowCount("frmPttel", 23562)
				' Արտահանել Excel
				Call ExportToExcel("frmPttel", actualFile6)
				' Համեմատել Excel ֆայլերը
				Call CompareTwoExcelFiles(actualFile6, expectedFile6, resultFile6)
				'ö³Ï»É åïï»ÉÁ
				Call CloseAllExcelFiles()
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
		folderName = "|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|ä³ÛÙ³Ý³·ñ»ñ|²Ù÷á÷|î»Õ³µ³ßËí³Í ÙÇçáóÝ»ñ|"
	
		sDATE = "20030101"
		fDATE = "20260101"  
		
		colName(0) = "fKEY"
		colName(1) = "fCOM"
		colName(2) = "fCURRENCY"
		colName(3) = "fSUID"
		colName(4) = "fDATE"
		
		' ä³ÛÙ³Ý³·ñ»ñ
		expectedFile1 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest1\Expected\expectedFile1.xlsx"
		' ¶áñÍáÕáõÃÛáõÝÝ»ñ
		expectedFile2 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest1\Expected\expectedFile2.xlsx"
  ' ä³ÛÙ³Ý³·ñÇ Ù³ñÙ³Ý(í»ñ³ýÇÝ³Ýë³íáñÙ³Ý) Å³ÙÏ»ïÝ»ñ
		expectedFile3 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest1\Expected\expectedFile3.xlsx"
		' îáÏáëÝ»ñÇ Ù³ñÙ³Ý Å³ÙÏ»ïÝ»ñ
		expectedFile4 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest1\Expected\expectedFile4.xlsx"
		' è. ¹³ë. ¨ å³Ñáõëï. ïáÏ. ÷á÷áËáõÃÛáõÝÝ»ñ
		expectedFile5 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest1\Expected\expectedFile5.xlsx"
		' úµÛ»ÏïÇí éÇëÏÇ ¹³ëÇãÇ ÷á÷áËáõÃÛáõÝÝ»ñ
		expectedFile6 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest1\Expected\expectedFile6.xlsx"
	
  ' ä³ÛÙ³Ý³·ñ»ñ
		actualFile1 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest1\Actual\actualFile1.xlsx"
		' ¶áñÍáÕáõÃÛáõÝÝ»ñ
		actualFile2 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest1\Actual\actualFile2.xlsx"
  ' ä³ÛÙ³Ý³·ñÇ Ù³ñÙ³Ý(í»ñ³ýÇÝ³Ýë³íáñÙ³Ý) Å³ÙÏ»ïÝ»ñ
		actualFile3 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest1\Actual\actualFile3.xlsx"
		' îáÏáëÝ»ñÇ Ù³ñÙ³Ý Å³ÙÏ»ïÝ»ñ
		actualFile4 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest1\Actual\actualFile4.xlsx"
		' è. ¹³ë. ¨ å³Ñáõëï. ïáÏ. ÷á÷áËáõÃÛáõÝÝ»ñ
		actualFile5 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest1\Actual\actualFile5.xlsx"
		' úµÛ»ÏïÇí éÇëÏÇ ¹³ëÇãÇ ÷á÷áËáõÃÛáõÝÝ»ñ
		actualFile6 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest1\Actual\actualFile6.xlsx"
		
  ' ä³ÛÙ³Ý³·ñ»ñ
		resultFile1 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest1\Result\resultFile1.xlsx"
  '¶áñÍáÕáõÃÛáõÝÝ»ñ
		resultFile2 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest1\Result\resultFile2.xlsx"
		' ä³ÛÙ³Ý³·ñÇ Ù³ñÙ³Ý(í»ñ³ýÇÝ³Ýë³íáñÙ³Ý) Å³ÙÏ»ïÝ»ñ
		resultFile3 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest1\Result\resultFile3.xlsx"
		' îáÏáëÝ»ñÇ Ù³ñÙ³Ý Å³ÙÏ»ïÝ»ñ
		resultFile4 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest1\Result\resultFile4.xlsx"
		' è. ¹³ë. ¨ å³Ñáõëï. ïáÏ. ÷á÷áËáõÃÛáõÝÝ»ñ
		resultFile5 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest1\Result\resultFile5.xlsx"
		' úµÛ»ÏïÇí éÇëÏÇ ¹³ëÇãÇ ÷á÷áËáõÃÛáõÝÝ»ñ
		resultFile6 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest1\Result\resultFile6.xlsx"
		
  ' ä³ÛÙ³Ý³·ñ»ñ
		Set contract1 = New_ContractsFilter()
		contract1.AgreementLevel = "1"
		contract1.ShowClosed = 1
		contract1.NotFullClosedExist = true
		contract1.ShowNotFullClosedAgr = 1
	
		'¶áñÍáÕáõÃÛáõÝÝ»ñ
		Set operations1 = New_AllocFundsOperations()
		with operations1 
				.startDate = "31/01/14"
				.agreeType = "22"
				.curr = "001"
				.performer = "17"
    .fill = "1"
				.operationType = "e9"
		end with
		
		' ä³ÛÙ³Ý³·ñÇ Ù³ñÙ³Ý(í»ñ³ýÇÝ³Ýë³íáñÙ³Ý) Å³ÙÏ»ïÝ»ñ
		Set agreeRepayTerms1 = New_AgreementsCommomFilter()
		agreeRepayTerms1.onlyChangesExists = true
		
		' îáÏáëÝ»ñÇ Ù³ñÙ³Ý Å³ÙÏ»ïÝ»ñ
		Set interestRepayTerms1 = New_AgreementsCommomFilter()
		interestRepayTerms1.onlyChangesExists = true
		
		' è. ¹³ë. ¨ å³Ñáõëï. ïáÏ. ÷á÷áËáõÃÛáõÝÝ»ñ
		Set riskiness1 = New_AgreementsCommomFilter()
		riskiness1.onlyChangesExists = true
		
		' úµÛ»ÏïÇí éÇëÏÇ ¹³ëÇãÇ ÷á÷áËáõÃÛáõÝÝ»ñ
		Set objRisk1 = New_AgreementsCommomFilter()
		objRisk1.onlyChangesExists = true
		
End Sub