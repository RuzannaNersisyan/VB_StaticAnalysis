'USEUNIT Subsystems_SQL_Library
'USEUNIT Library_Common
'USEUNIT Library_Contracts
'USEUNIT Constants
'USEUNIT Library_Colour
'USEUNIT OLAP_Library
'USEUNIT SWIFT_International_Payorder_Library
Option Explicit

'Test case ID 161489

Dim folderName, sDATE, fDATE, colName(5), param
Dim contract1, operations1, agreeRepayTerms1, interestRepayTerms1, riskiness1, objRisk1 
Dim actualFile1, actualFile2, actualFile3, actualFile4, actualFile5, actualFile6
Dim expectedFile1, expectedFile2, expectedFile3, expectedFile4, expectedFile5, expectedFile6, resultFile1

Sub Contracts_Summary_Reports_2()
		' ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|ä³ÛÙ³Ý³·ñ»ñ|²Ù÷á÷|î»Õ³µ³ßËí³Í ÙÇçáóÝ»ñ
		Call Test_Initialize()

		' Համակարգ մուտք գործել ARMSOFT օգտագործողով
		Log.Message "Համակարգ մուտք գործել ARMSOFT օգտագործողով", "", pmNormal, DivideColor
  Call Test_StartUp()
		
		'''''''''''''''''''''''''''''''''''''''''''''''''
		''''''''''''''''''ä³ÛÙ³Ý³·ñ»ñ''''''''''''''''''''
		
		' Լրացնել Պայմանագրեր դիալոգային պատուհանը
		Log.Message "Պայմանագրեր 1 տարբերակ", "", pmNormal, DivideColor
		Call GoTo_Contracts(folderName, contract1)
		
		if WaitForExecutionProgress() then		
				'êáñï³íáñ»É µ³óí³Í åïï»ÉÁ
				Call columnSorting(colName, 3, "frmPttel")
				' Ստուգել տողերի քանակը
				Call CheckPttel_RowCount("frmPttel", 9)
				' Արտահանել, որպես txt ֆայլ
				Call ExportToTXTFromPttel("frmPttel", actualFile1)
				' Համեմատել txt ֆայլերը
				Call Compare_Files(actualFile1, expectedFile1, param)
				' ö³Ï»É åïï»ÉÁ
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
	
		if WaitForExecutionProgress() then		
				'êáñï³íáñ»É µ³óí³Í åïï»ÉÁ
				Call columnSorting(colName, 5, "frmPttel")
				' Ստուգել տողերի քանակը
				Call CheckPttel_RowCount("frmPttel", 208190)
				' Արտահանել Excel
				Call ExportToExcel("frmPttel", actualFile2)
				' Համեմատել Excel ֆայլերը
				Call CompareTwoExcelFiles(actualFile2, expectedFile2, resultFile1)
				'ö³Ï»É åïï»ÉÁ
				Call CloseAllExcelFiles()
				BuiltIn.Delay(3000) 
		  wMDIClient.VBObject("frmPttel").Close
		else																																	
						Log.Error "Can't open pttel window.", "", pmNormal, ErrorColor
		end if
		
		'''''''''''''''''''''''''''''''''''''''''''''''''
		''''ä³ÛÙ. Ù³ñÙ³Ý(í»ñ³ýÇÝ³Ýë³íáñÙ³Ý) Å³ÙÏ»ïÝ»ñ''''
		
		' Լրացնել Ժամկետներ դիալոգային պատուհանը
		Log.Message "Պայմ. մարման(վերաֆինանսավորման) ժամկետներ 1 տարբերակ", "", pmNormal, DivideColor
		Call GoTo_AgreementsCommomFilter(folderName, "Ä³ÙÏ»ïÝ»ñ|ä³ÛÙ.Ù³ñÙ³Ý(í»ñ³ýÇÝ³Ýë³íáñÙ³Ý) Å³ÙÏ»ïÝ»ñ", agreeRepayTerms1)
		
		if WaitForExecutionProgress() then		
				'êáñï³íáñ»É µ³óí³Í åïï»ÉÁ
				Call columnSorting(colName, 5, "frmPttel")
				' Ստուգել տողերի քանակը
				Call CheckPttel_RowCount("frmPttel", 93)
				' Արտահանել, որպես txt ֆայլ
				Call ExportToTXTFromPttel("frmPttel", actualFile3)
				' Համեմատել txt ֆայլերը
				Call Compare_Files(actualFile3, expectedFile3, param)
				' ö³Ï»É åïï»ÉÁ
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
				'êáñï³íáñ»É µ³óí³Í åïï»ÉÁ
				Call columnSorting(colName, 5, "frmPttel")
				' Ստուգել տողերի քանակը
				Call CheckPttel_RowCount("frmPttel", 10)
				' Արտահանել, որպես txt ֆայլ
				Call ExportToTXTFromPttel("frmPttel", actualFile4)
				' Համեմատել txt ֆայլերը
				Call Compare_Files(actualFile4, expectedFile4, param)
				'ö³Ï»É åïï»ÉÁ
				BuiltIn.Delay(3000) 
		  wMDIClient.VBObject("frmPttel").Close
		else																																	
						Log.Error "Can't open pttel window.", "", pmNormal, ErrorColor
		end if
		
		'''''''''''''''''''''''''''''''''''''''''''''''''
		'''''è. ¹³ë. ¨ å³Ñáõëï. ïáÏ. ÷á÷áËáõÃÛáõÝÝ»ñ'''''

		' Լրացնել Ռիսկի դասիչներ և պահուստավորման տոկոսներ դիալոգային պատուհանը
		Log.Message "Ռիսկի դասիչներ և պահուստավորման տոկոսներ 1 տարբերակ", "", pmNormal, DivideColor
		Call GoTo_AgreementsCommomFilter(folderName, "èÇëÏ³ÛÝáõÃÛáõÝ|è.¹³ë. ¨ å³Ñáõëï.ïáÏ. ÷á÷áËáõÃÛáõÝÝ»ñ", riskiness1)
		
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
		
		'''''''''''''''''''''''''''''''''''''''''''''''''
		'''''''úµÛ»ÏïÇí éÇëÏÇ ¹³ëÇãÇ ÷á÷ËáõÃÛáõÝÝ»ñ''''''
		
		' Լրացնել Օբյեկտիվ ռիսկի դասիչ դիալոգային պատուհանը
		Log.Message "Օբյեկտիվ ռիսկի դասիչ 1 տարբերակ", "", pmNormal, DivideColor
		Call GoTo_AgreementsCommomFilter(folderName, "èÇëÏ³ÛÝáõÃÛáõÝ|úµÛ»ÏïÇí éÇëÏÇ ¹³ëÇãÇ ÷á÷áËáõÃÛáõÝÝ»ñ", objRisk1)
		
		if WaitForExecutionProgress() then		
				' Արտահանել, որպես txt ֆայլ
				Call ExportToTXTFromPttel("frmPttel", actualFile6)
				' Ստուգել տողերի քանակը
				Call CheckPttel_RowCount("frmPttel", 1)
				' Համեմատել txt ֆայլերը
				Call Compare_Files(actualFile6, expectedFile6, param)
				'ö³Ï»É åïï»ÉÁ
				BuiltIn.Delay(3000) 
		  wMDIClient.VBObject("frmPttel").Close
		else																																	
						Log.Error "Can't open pttel window.", "", pmNormal, ErrorColor
		end if
		
		Call Close_AsBank()		
End	Sub

Sub Test_StartUp()
		Call Initialize_AsBank("bank_Report", sDATE, fDATE)
  Login("ARMSOFT")
		Call SaveRAM_RowsLimit("100")
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
		expectedFile1 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest2\Expected\expectedFile1.txt"
		' ¶áñÍáÕáõÃÛáõÝÝ»ñ
		expectedFile2 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest2\Expected\expectedFile2.xlsx"
  ' ä³ÛÙ³Ý³·ñÇ Ù³ñÙ³Ý(í»ñ³ýÇÝ³Ýë³íáñÙ³Ý) Å³ÙÏ»ïÝ»ñ
		expectedFile3 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest2\Expected\expectedFile3.txt"
		' îáÏáëÝ»ñÇ Ù³ñÙ³Ý Å³ÙÏ»ïÝ»ñ
		expectedFile4 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest2\Expected\expectedFile4.txt"
		' è. ¹³ë. ¨ å³Ñáõëï. ïáÏ. ÷á÷áËáõÃÛáõÝÝ»ñ
		expectedFile5 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest2\Expected\expectedFile5.txt"
		' úµÛ»ÏïÇí éÇëÏÇ ¹³ëÇãÇ ÷á÷áËáõÃÛáõÝÝ»ñ
		expectedFile6 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest2\Expected\expectedFile6.txt"
	
  ' ä³ÛÙ³Ý³·ñ»ñ
		actualFile1 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest2\Actual\actualFile1.txt"
		' ¶áñÍáÕáõÃÛáõÝÝ»ñ
		actualFile2 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest2\Actual\actualFile2.xlsx"
  ' ä³ÛÙ³Ý³·ñÇ Ù³ñÙ³Ý(í»ñ³ýÇÝ³Ýë³íáñÙ³Ý) Å³ÙÏ»ïÝ»ñ
		actualFile3 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest2\Actual\actualFile3.txt"
		' îáÏáëÝ»ñÇ Ù³ñÙ³Ý Å³ÙÏ»ïÝ»ñ
		actualFile4 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest2\Actual\actualFile4.txt"
		' è. ¹³ë. ¨ å³Ñáõëï. ïáÏ. ÷á÷áËáõÃÛáõÝÝ»ñ
		actualFile5 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest2\Actual\actualFile5.txt"
		' úµÛ»ÏïÇí éÇëÏÇ ¹³ëÇãÇ ÷á÷áËáõÃÛáõÝÝ»ñ
		actualFile6 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest2\Actual\actualFile6.txt"
		
		'¶áñÍáÕáõÃÛáõÝÝ»ñ
		resultFile1 = Project.Path & "Stores\Reports\Subsystems\Summary\AllocatedFundsTest2\Result\resultFile2.xlsx"
		
  ' ä³ÛÙ³Ý³·ñ»ñ
		Set contract1 = New_ContractsFilter()
		contract1.AgreementLevel = "2"
		contract1.AgreementSpecies = "1"
		
		'¶áñÍáÕáõÃÛáõÝÝ»ñ
		Set operations1 = New_AllocFundsOperations()
		operations1.endDate = "31/03/08"
		
		' ä³ÛÙ³Ý³·ñÇ Ù³ñÙ³Ý(í»ñ³ýÇÝ³Ýë³íáñÙ³Ý) Å³ÙÏ»ïÝ»ñ		
		Set agreeRepayTerms1 = New_AgreementsCommomFilter()
		with agreeRepayTerms1
				.startDate = "08/01/14"
				.endDate = "08/01/14"
				.agreeN = "TV4253"
				.performer = "189"
				.note = "00"
				.note2 = "02"
				.agreeOffice = "P00"
				.agreeSection = "05"
				.accessType = "C11"
				.onlyChangesExists = true
				.onlyChanges = 1
				.showInOpFormExists = true
				.showInOpForm = 1
		end with
		
		' îáÏáëÝ»ñÇ Ù³ñÙ³Ý Å³ÙÏ»ïÝ»ñ
		Set interestRepayTerms1 = New_AgreementsCommomFilter()
		with interestRepayTerms1
				.startDate = "14/12/17"
				.endDate = "14/12/17"
				.agreeN = "TV22128"
				.performer = "253"
				.note = "002"
				.note2 = "02"
				.note3 = "01"
				.agreeOffice = "P00"
				.agreeSection = "08"
				.accessType = "C11"
				.onlyChangesExists = true
				.onlyChanges = 1
				.showInOpFormExists = true
				.showInOpForm = 1
		end with
		
		' è. ¹³ë. ¨ å³Ñáõëï. ïáÏ. ÷á÷áËáõÃÛáõÝÝ»ñ
		Set riskiness1 = New_AgreementsCommomFilter()
		with riskiness1
				.endDate = "16/05/19"
				.agreeN = "TV19577"
				.agreeOffice = "P10"
				.accessType = "C12"
				.onlyChangesExists = true
				.onlyChanges = 1
		end with
		
		' úµÛ»ÏïÇí éÇëÏÇ ¹³ëÇãÇ ÷á÷áËáõÃÛáõÝÝ»ñ
		Set objRisk1 = New_AgreementsCommomFilter()
		with objRisk1
				.startDate = "16/01/14"
				.agreeN = "TV15179"
				.agreeOffice = "P02"
				.accessType = "C12"
				.onlyChangesExists = true
				.onlyChanges = 1
		end with
		
End Sub