  Option Explicit
'USEUNIT Subsystems_SQL_Library
'USEUNIT Library_Common
'USEUNIT Card_Library
'USEUNIT Mortgage_Library
'USEUNIT OLAP_Library
'USEUNIT Constants
'USEUNIT Library_Colour

'Test Case ID 165988

Sub Cards_Registr_SMS_Test()

    Dim DateStart,DateEnd
    Dim SMS_Messages,Path1,Path2,resultWorksheet
    Dim queryString,sql_Value, colNum,sql_isEqual
    
    DateStart = "20010101"
    DateEnd = "20240101"
  
    Set SMS_Messages = New_SMS_Messages()
    With SMS_Messages
        .FileDate_1 = "121017"
        .FileDate_2 = "121017"
        .ShowProcessed = 1
        .View = "vACSMS\2"
        .FillInto = "0"
    End With
    
    'queryString = "update statistics HI DELETE FROM HI WHERE fDATE = '2018-05-23'"
    
    'Test StartUp start
    Call Initialize_AsBankQA(DateStart, DateEnd)
    
    Call Create_Connection()    
    'Call Execute_SLQ_Query(queryString) 
  
    Call ChangeWorkspace(c_CardsSV)
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''--- Թղթապանակի ստուգում ֆայլերի ընդունելուց հետո ---''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''' ''''''''''''''
    Log.Message "--- Check SMS_Messages After Add Files ---" ,,, DivideColor 

    Call GoToSMS_Messages_PlasticCarts(SMS_Messages)
    Call CheckPttel_RowCount("frmPttel", 822)

    Path1 = Project.Path & "Stores\ExpectedReports\PlasticCards\FilesRegistr\Actual\Actual_SMS_Messages.xlsx"
    Path2 = Project.Path & "Stores\ExpectedReports\PlasticCards\FilesRegistr\Expected\Expected_SMS_Messages.xlsx"
    resultWorksheet = Project.Path & "Stores\ExpectedReports\PlasticCards\FilesRegistr\Result\Result_SMS_Messages.xlsx"
    
    'Արտահանել և Ð³Ù»Ù³ï»É »ñÏáõ EXCEL ý³ÛÉ»ñ
    Call ExportToExcel("frmPttel",Path1)
    Call CompareTwoExcelFiles(Path1, Path2, resultWorksheet)  
    Call CloseAllExcelFiles()
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''--- Հաշվառել փաստաթղթերը ---''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 
    Log.Message "--- Registr Cards Files ---" ,,, DivideColor 

    Call Registr_Cards_Total("230518")
    Call Close_Pttel("frmPttel")   
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''--- Թղթապանակի ստուգում ֆայլերի Հաշվառելուց հետո ---'''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''' ''''''''''''''
    Log.Message "--- Check SMS Messages After Registr Cards ---" ,,, DivideColor 

    SMS_Messages.FileDate_2 = "230518"
    Call GoToSMS_Messages_PlasticCarts(SMS_Messages)
    Call CheckPttel_RowCount("frmPttel", 822)

    Path1 = Project.Path & "Stores\ExpectedReports\PlasticCards\FilesRegistr\Actual\Actual_SMSMessagesAfterRegistr.xlsx"
    Path2 = Project.Path & "Stores\ExpectedReports\PlasticCards\FilesRegistr\Expected\Expected_SMSMessagesAfterRegistr.xlsx"
    resultWorksheet = Project.Path & "Stores\ExpectedReports\PlasticCards\FilesRegistr\Result\Result_SMSMessagesAfterRegistr.xlsx"
    
    'Արտահանել և Ð³Ù»Ù³ï»É »ñÏáõ EXCEL ý³ÛÉ»ñ
    Call ExportToExcel("frmPttel",Path1)
    Call CompareTwoExcelFiles(Path1, Path2, resultWorksheet)  
    Call CloseAllExcelFiles()
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''--- Կատարում է SQL ստուգում ---''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 
    Log.Message "SQL Check For SMS Messages",,,SqlDivideColor
      
      'Կատարում է SQL ստուգում
      queryString = "select Count(*) from HI where fDATE = '2018-05-23' "
      sql_Value = 14
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If
       
      queryString = "select Sum(fSUM) from HI where fDATE = '2018-05-23' and fTYPE = '01' and fDBCR = 'C' "
      sql_Value = 164
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If
  
      queryString = "select Sum(fCURSUM) from HI where fDATE = '2018-05-23' and fTYPE = '01' and fDBCR = 'C'"
      sql_Value = 164
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If
       
      queryString = "select Sum(fSUM) from HI where fDATE = '2018-05-23' and fTYPE = '01' and fDBCR = 'D'"
      sql_Value = 164
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If 
       
      queryString = "select Sum(fCURSUM) from HI where fDATE = '2018-05-23' and fTYPE = '01' and fDBCR = 'D'"
      sql_Value = 164
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If      
       
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''--- Ջնջում է բոլոր ներմուծած ֆայլերը ---''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "--- Delete All Contracts Total ---" ,,, DivideColor       

    'Ջնջում է բոլոր ներմուծած ֆայլերը
    Call Delete_All_Contracts_Total()
    Call Close_Pttel("frmPttel")
  
      'Կատարում է SQL ստուգում
      queryString = "select Count(*) from HI where fDATE = '2018-05-23' "
      sql_Value = 0
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If
       
    'Փակել ASBANK-ը
    Call Close_AsBank()
End Sub