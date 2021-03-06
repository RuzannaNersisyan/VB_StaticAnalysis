Option Explicit
'USEUNIT Subsystems_SQL_Library
'USEUNIT Library_Common
'USEUNIT Card_Library
'USEUNIT Mortgage_Library
'USEUNIT OLAP_Library
'USEUNIT Constants
'USEUNIT Library_Colour

'Test Case Id 166001

' Պլաստիկ Քարտեր ԱՇՏ/Ստացված գործողություններ թղթապանակ
Sub Cards_Registr_Actions_Test()
 
    Dim DateStart,DateEnd
    Dim ReceivedTrans,Path1,Path2,resultWorksheet
    Dim queryString,sql_Value, colNum,sql_isEqual
  
    DateStart = "20010101"
    DateEnd = "20240101"
    
    Set ReceivedTrans = New_ReceivedTransactions()
    With ReceivedTrans
        .FileDate_1 = "121017"
        .FileDate_2 = "121017"
        .CardsTransactions = 1
        .MerchantPointTransactions = 1
        .ShowMadeTransactions = 1
        .ShowAllRows = 1
        .ShowArchivedOpers = 0
        .View = "VRecTrns\2"
        .FillInto = "0"
    End With

    queryString = " update statistics HI  DELETE FROM HI WHERE fDATE = '2018-05-27'"
  
    'Test StartUp start
    Call Initialize_AsBankQA(DateStart, DateEnd) 

    Call Create_Connection()
    Call Execute_SLQ_Query(queryString)
    Call ChangeWorkspace(c_CardsSV)
  
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''--- Թղթապանակի ստուգում ֆայլերի ընդունելուց հետո ---''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''' ''''''''''''''
    Log.Message "--- Check Received Transaction After Add Files ---" ,,, DivideColor 

    Call GoToReceivedTrans_PlasticCarts(ReceivedTrans)  
    Call CheckPttel_RowCount("frmPttel", 3070)

    Path1 = Project.Path & "Stores\ExpectedReports\PlasticCards\FilesRegistr\Actual\Actual_ReceivedTrans.xlsx"
    Path2 = Project.Path & "Stores\ExpectedReports\PlasticCards\FilesRegistr\Expected\Expected_ReceivedTrans.xlsx"
    resultWorksheet = Project.Path & "Stores\ExpectedReports\PlasticCards\FilesRegistr\Result\Result_ReceivedTrans.xlsx"
    
    'Արտահանել և Ð³Ù»Ù³ï»É »ñÏáõ EXCEL ý³ÛÉ»ñ
'    Call ExportToExcel("frmPttel",Path1)
'    Call CompareTwoExcelFiles(Path1, Path2, resultWorksheet)  
'    Call CloseAllExcelFiles()    
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''--- Հաշվառել փաստաթղթերը ---''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 
    Log.Message "--- Registr Cards Files ---" ,,, DivideColor 

    Call Registr_Cards_Total("270518")
    Call Close_Pttel("frmPttel")   
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''--- Թղթապանակի ստուգում ֆայլերի Հաշվառելուց հետո ---'''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''' ''''''''''''''
    Log.Message "--- Check Received Transaction After Registr Cards ---" ,,, DivideColor 

    ReceivedTrans.FileDate_2 = "270518"
    Call GoToReceivedTrans_PlasticCarts(ReceivedTrans) 
    Call CheckPttel_RowCount("frmPttel", 3070)

    Path1 = Project.Path & "Stores\ExpectedReports\PlasticCards\FilesRegistr\Actual\Actual_ReceivedTransAfterRegistr.xlsx"
    Path2 = Project.Path & "Stores\ExpectedReports\PlasticCards\FilesRegistr\Expected\Expected_ReceivedTransAfterRegistr.xlsx"
    resultWorksheet = Project.Path & "Stores\ExpectedReports\PlasticCards\FilesRegistr\Result\Result_ReceivedTransAfterRegistr.xlsx"
    
    'Արտահանել և Ð³Ù»Ù³ï»É »ñÏáõ EXCEL ý³ÛÉ»ñ
'    Call ExportToExcel("frmPttel",Path1)
'    Call CompareTwoExcelFiles(Path1, Path2, resultWorksheet)  
'    Call CloseAllExcelFiles()    
        
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''--- Կատարում է SQL ստուգում ---''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 
    Log.Message "SQL Check For Rec Clearing Transation",,,SqlDivideColor    
    
      'Կատարում ենք SQL ստուգում
      queryString = "select Count(*) from HI where fDATE = '2018-05-27' "
      sql_Value = 6
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If
       
      queryString = "select Sum(fSUM) from HI where fDATE = '2018-05-27' and fTYPE = '01' and fDBCR = 'C' "
      sql_Value = 29634
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If
  
      queryString = "select Sum(fCURSUM) from HI where fDATE = '2018-05-27' and fTYPE = '01' and fDBCR = 'C'"
      sql_Value = 29634
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If
       
      queryString = "select Sum(fSUM) from HI where fDATE = '2018-05-27' and fTYPE = '01' and fDBCR = 'D'"
      sql_Value = 29634
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If 
       
      queryString = "select Sum(fCURSUM) from HI where fDATE = '2018-05-27' and fTYPE = '01' and fDBCR = 'D'"
      sql_Value = 29634
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
    
    'Ջնջում է ներմուծման փաստաթղթերը  
    Call Delete_PCTrans("121017","10","PCTrans")
    
      'Կատարում ենք SQL ստուգում
      queryString = "select Count(*) from HI where fDATE = '2018-05-27' "
      sql_Value = 0
      colNum = 0
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If
     
    'Փակել ASBANK-ը
    Call Close_AsBank()
End Sub