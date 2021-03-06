'USEUNIT Library_Common
'USEUNIT Library_Contracts
'USEUNIT Library_Colour
'USEUNIT BankMail_Library
'USEUNIT Payment_Except_Library
'USEUNIT Constants
Option Explicit

'Test Case ID 183987

Dim sDate, eDate, folderName, expectedFile, actualFilePath, actualFile, dictPatterns
Dim column_number, actual_value, time, hours, mins

Sub Olap_Job_Result_Comparing_Test()
    Call Test_Inintialize()

				' Համակարգ մուտք գործել ARMSOFT օգտագործողով
				Log.Message "Համակարգ մուտք գործել ARMSOFT օգտագործողով", "", pmNormal, DivideColor
				Call Test_StartUp()
    
    ' Մուտք գործել Առաջադրանքներ թղթապանակ 
    Call GoTo_Tasks("010219", "010219")
    
    ' Հաշվետվություն կատարման մասին
    Log.Message "Հաշվետվություն կատարման մասին", "", pmNormal, DivideColor
    BuiltIn.Delay(3000)
    Call SearchInPttel("frmPttel", 0, "4852")
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_PerformanceReport)
    If wMDIClient.VBObject("FrmSpr").Exists Then
        ' Հաշվետվության պահպանում 
        Call SaveDoc(actualFilePath, actualFile)
        ' Փակել Հաշվետվության պատուհանը 
    				Call Close_Window(wMDIClient, "FrmSpr")
    End If
    
    ' Փաստացի Հաշվետվության համեմատում սպասվողի հետ
				Log.Message "Փաստացի Հաշվետվության համեմատում սպասվողի հետ", "", pmNormal, DivideColor
    Set dictPatterns = CreateObject("Scripting.Dictionary")
    dictPatterns.Add 1, "\d{1,2}.\d{1,2}.\d{1,2}\s\d{1,2}:\d{1,2}:\d{1,2}"
				dictPatterns.Add 2, "\d{1,2}:\d{1,2}:\d{1,2}"
    Call Compare_Files_With_Patterns_Array(actualFilePath & actualFile, expectedFile, dictPatterns)
    
    ' Ստուգել կատարման ժամանակը
    column_number = wMDIClient.VBObject("frmPttel").GetColumnIndex("DURATION")
    actual_value = wMDIClient.VBObject("frmPttel").VBObject("tdbgView").Columns.Item(column_number).Text
    time = Split(actual_value, ":")
    hours = time(0)
    mins = time(1)
    
    Select Case hours
    Case 3
         If mins < "50" Then
            Log.Error "Executing time is 3 houres and " & mins & " minutes!", "", pmNormal, ErrorColor
         End If
    Case 4
        If mins > "10" Then
            Log.Error "Executing time is 4 houres and " & mins & " minutes!", "", pmNormal, ErrorColor
        End If
    Case Else
        Log.Error "Executing time is " & hours &" houres and " & mins & "minutes!", "", pmNormal, ErrorColor
    End Select 
    
    ' Փակել Առաջադրանքներ թղթապանակը 
				Call Close_Window(wMDIClient, "frmPttel") 
    
    ' Փակել ծրագիրը
				Call Close_AsBank()
End Sub

Sub Test_StartUp()
				Call Initialize_AsBankQA(sDate, eDate)   
				Login("ARMSOFT")
				' Մուտք Գլխավոր հաշվապահի ԱՇՏ
				Call ChangeWorkspace(c_Admin40)
End Sub

Sub Test_Inintialize()
				sDate = "20030101"
				eDate = "20250101"
    
    expectedFile = Project.Path &  "Stores\Reports\Tasks\Expected\Expected_Performance_Report.txt"
				actualFilePath = Project.Path &  "Stores\Reports\Tasks\Actual\"
    actualFile = "Actual_Performance_Report.txt"
		
				folderName = "|²¹ÙÇÝÇëïñ³ïáñÇ ²Þî|"
End Sub