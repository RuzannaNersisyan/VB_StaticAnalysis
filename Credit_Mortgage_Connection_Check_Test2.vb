'USEUNIT Library_Common
'USEUNIT Library_CheckDB
'USEUNIT Mortgage_Library
'USEUNIT Credit_Mortgage_Connection_Check_Library
'USEUNIT Payment_Order_ConfirmPhases_Library
'USEUNIT Online_PaySys_Library
'USEUNIT Loan_Agreements_Library
'USEUNIT Subsystems_SQL_Library
'USEUNIT Constants

'Test case ID 165673

Sub Credit_Mortgage_Connection_Check_Test2()
    
    Dim startDATE, fDATE
    Dim pType, pNumber, cliCode, mortName, fillGrid, sDate, fBASE, docNumber
    Dim mortCurr , mortSumma, mortCount , mortComment , queryString,mortageItemNew, MortSubject
    
    Utilities.ShortDateFormat = "yyyymmdd"
    CurrentDate = aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%d%m%y")
    startDATE = "20030101"
    fDATE = "20250101"
    
    'Test StartUp 
    Call Initialize_AsBank("bank", startDATE, fDATE)
    Login ("MORTGAGEOPERATOR")           
    Call Create_Connection ()
    Call Initialize_Arrays(1, 1, 4, 1)
    
    loanAgrNum(1) = "V-002520"
    loanAgrType(1) = "2"
    partnerCode(1) = "00000008"
    partnerCode(2) = "00000009"
    partnerCode(3) = "00000010"
    partnerCode(4) = "00034853"    
    
    'êï»ÕÍ»É Ýáñ å³ÛÙ³Ý³·Çñ ` "¶ñ³í (³ÛÉ)"
    agrType = "¶ñ³í(³ÛÉ)"
    cliCode = "00034851"
    mortCurr = "000"
    mortSumma = "10000"
    mortCount = "1"
    fillGrid = True
    MortSubject = 0
    Call Mortgage_Doc_Fill (agrType , pType, pNumber, cliCode, mortName, fillGrid, _
                            loanAgrType , loanAgrClient , loanAgrNum, mortCurr , mortSumma, mortCount , mortComment, _
                            sDate, fBASE, docNumber,mortageItemNew, MortSubject)
    
    queryString = "SELECT COUNT(*) FROM DOCS WHERE fBODY LIKE '%" & docNumber & "%' and fSTATE='1' and fNAME='N1Mort' "
    sql_Value = 1
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Call Log_Error_My()
        Log.Error "Querystring = " & queryString & ":  Expected result = " & sql_Value , "", pmNormal, attr
    End If
    BuiltIn.Delay(delay_small)
    
    queryString = "SELECT COUNT(*) FROM FOLDERS WHERE fISN= '" & fBASE & "'"
    sql_Value = 9
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Call Log_Error_My()
        Log.Error "Querystring = " & queryString & ":  Expected result = " & sql_Value , "", pmNormal, attr
    End If
    
    objCount = "10"
    objSum = "1000"
    Call Create_New_Object_Other(objCount , objSum)
    
    BuiltIn.Delay(1000)
    wMDIClient.vbObject("frmPttel").Close()
    
    'êï»ÕÍ»É Ýáñ å³ÛÙ³Ý³·Çñ ` "¶ñ³í (³ÛÉ)"
    agrType = "¶ñ³í(³ÛÉ)ª³íïáÙ³ï µ³óíáÕ"
    cliCode = "00034851"
    mortCurr = "000"
    mortageItemNew = "0"
    Call Mortgage_Doc_Fill (agrType , pType, pNumber, cliCode, mortName, fillGrid, _
                            loanAgrType , loanAgrClient , loanAgrNum, mortCurr , mortSumma, mortCount , mortComment, _
                            sDate, fBASE1, docNumber1,mortageItemNew, MortSubject)
    
    queryString = "SELECT COUNT(*) FROM DOCS WHERE fBODY LIKE '%" & docNumber1 & "%' and fSTATE='1' and fNAME='N1Mort' "
    sql_Value = 1
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Call Log_Error_My()
        Log.Error "Querystring = " & queryString & ":  Expected result = " & sql_Value , "", pmNormal, attr
    End If
    BuiltIn.Delay(delay_small)
    
    queryString = "SELECT COUNT(*) FROM FOLDERS WHERE fISN= '" & fBASE1 & "'"
    sql_Value = 9
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Call Log_Error_My()
        Log.Error "Querystring = " & queryString & ":  Expected result = " & sql_Value , "", pmNormal, attr
    End If
    
    BuiltIn.Delay(1000)
    wMDIClient.vbObject("frmPttel").Close()
 
    'êï»ÕÍ»É Ýáñ å³ÛÙ³Ý³·Çñ ` "¶ñ³í (³Ýß³ñÅ ·áõÛù)"
    agrType = "¶ñ³í(³Ýß³ñÅ ·áõÛù)"
    cliCode = "00034851"
    mortCurr = "003"
    Call Mortgage_Doc_Fill (agrType , pType, pNumber, cliCode, mortName, fillGrid, _
                            loanAgrType , loanAgrClient , loanAgrNum, mortCurr , mortSumma, mortCount , mortComment, _
                            sDate, fBASE2, docNumber2,mortageItemNew, MortSubject)
    
    queryString = "SELECT COUNT(*) FROM DOCS WHERE fBODY LIKE '%" & docNumber2 & "%' and fSTATE='1' and fNAME='N1Mort' "
    sql_Value = 1
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Call Log_Error_My()
        Log.Error "Querystring = " & queryString & ":  Expected result = " & sql_Value , "", pmNormal, attr
    End If
    BuiltIn.Delay(delay_small)
    
    queryString = "SELECT COUNT(*) FROM FOLDERS WHERE fISN= '" & fBASE2 & "'"
    sql_Value = 9
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Call Log_Error_My()
        Log.Error "Querystring = " & queryString & ":  Expected result = " & sql_Value , "", pmNormal, attr
    End If
    BuiltIn.Delay(delay_small)
    
    AmountObject = "1000"
    Call Create_Object_Car(AmountObject)
    
    BuiltIn.Delay(1000)
    wMDIClient.vbObject("frmPttel").Close()
    
    'êï»ÕÍ»É Ýáñ å³ÛÙ³Ý³·Çñ ` "¶ñ³í (³Ýß³ñÅ ·áõÛù)`³íïáÙ³ï µ³óíáÕ"
    agrType = "¶ñ³í(³Ýß³ñÅ ·áõÛù)ª³íïáÙ³ï µ³óíáÕ"
    cliCode = "00034851"
    mortCurr = "003"
    Call Mortgage_Doc_Fill (agrType , pType, pNumber, cliCode, mortName, fillGrid, _
                            loanAgrType , loanAgrClient , loanAgrNum, mortCurr , mortSumma, mortCount , mortComment, _
                            sDate, fBASE3, docNumber3,mortageItemNew, MortSubject)
    
    queryString = "SELECT COUNT(*) FROM DOCS WHERE fBODY LIKE '%" & docNumber3 & "%' and fSTATE='1' and fNAME='N1Mort' "
    sql_Value = 1
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Call Log_Error_My()
        Log.Error "Querystring = " & queryString & ":  Expected result = " & sql_Value , "", pmNormal, attr
    End If
    BuiltIn.Delay(delay_small)
    
    queryString = "SELECT COUNT(*) FROM FOLDERS WHERE fISN= '" & fBASE3 & "'"
    sql_Value = 9
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Call Log_Error_My()
        Log.Error "Querystring = " & queryString & ":  Expected result = " & sql_Value , "", pmNormal, attr
    End If
    
    BuiltIn.Delay(1000)
    wMDIClient.VBObject("frmPttel").Close()
    
    'êï»ÕÍ»É Ýáñ å³ÛÙ³Ý³·Çñ ` "¶ñ³í (áëÏÇ)"
    agrType = "¶ñ³í(áëÏÇ)"
    cliCode = "00034851"
    mortCurr = "006"
    Call Mortgage_Doc_Fill (agrType , pType, pNumber, cliCode, mortName, fillGrid, _
                            loanAgrType , loanAgrClient , loanAgrNum, mortCurr , mortSumma, mortCount , mortComment, _
                            sDate, fBASE4, docNumber4,mortageItemNew, MortSubject)
    
    queryString = "SELECT COUNT(*) FROM DOCS WHERE fBODY LIKE '%" & docNumber4 & "%' and fSTATE='1' and fNAME='N1Mort' "
    sql_Value = 1
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Call Log_Error_My()
        Log.Error "Querystring = " & queryString & ":  Expected result = " & sql_Value , "", pmNormal, attr
    End If
    BuiltIn.Delay(delay_small)
    
    queryString = "SELECT COUNT(*) FROM FOLDERS WHERE fISN= '" & fBASE4 & "'"
    sql_Value = 9
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Call Log_Error_My()
        Log.Error "Querystring = " & queryString & ":  Expected result = " & sql_Value , "", pmNormal, attr
    End If
    
    'êï»ÕÍ»É Ýáñ ³é³ñÏ³` ¶ñ³í(àëÏÇ)
    NameObject = "2"
    CountObject = "5"
    AmountObject = "1252"
    Call Create_Object_Gold(NameObject, CountObject, AmountObject)
    
    BuiltIn.Delay(1000)
    wMDIClient.vbObject("frmPttel").Close()
    
    'êï»ÕÍ»É Ýáñ å³ÛÙ³Ý³·Çñ ` "¶ñ³í (áëÏÇ)` ³íïáÙ³ï µ³óíáÕ"
    agrType = "¶ñ³í(áëÏÇ)ª³íïáÙ³ï µ³óíáÕ"
    cliCode = "00034851"
    mortCurr = "006"
    Call Mortgage_Doc_Fill (agrType , pType, pNumber, cliCode, mortName, fillGrid, _
                            loanAgrType , loanAgrClient , loanAgrNum, mortCurr , mortSumma, mortCount , mortComment, _
                            sDate, fBASE5, docNumber5,mortageItemNew, MortSubject)
    
    queryString = "SELECT COUNT(*) FROM DOCS WHERE fBODY LIKE '%" & docNumber5 & "%' and fSTATE='1' and fNAME='N1Mort' "
    sql_Value = 1
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Call Log_Error_My()
        Log.Error "Querystring = " & queryString & ":  Expected result = " & sql_Value , "", pmNormal, attr
    End If
    BuiltIn.Delay(delay_small)
    
    queryString = "SELECT COUNT(*) FROM FOLDERS WHERE fISN= '" & fBASE5 & "'"
    sql_Value = 9
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Call Log_Error_My()
        Log.Error "Querystring = " & queryString & ":  Expected result = " & sql_Value , "", pmNormal, attr
    End If
    
    BuiltIn.Delay(1000)
    wMDIClient.vbObject("frmPttel").Close()
    
    'êï»ÕÍ»É Ýáñ å³ÛÙ³Ý³·Çñ ` "¶ñ³í (÷áË³¹ñ³ÙÇçáó)"
    agrType = "¶ñ³í(÷áË³¹ñ³ÙÇçáó)"
    cliCode = "00034851"
    mortCurr = "001"
    MortSubject = 0
    Call Mortgage_Doc_Fill (agrType , pType, pNumber, cliCode, mortName, fillGrid, _
                            loanAgrType , loanAgrClient , loanAgrNum, mortCurr , mortSumma, mortCount , mortComment, _
                            sDate, fBASE6, docNumber6,mortageItemNew, MortSubject)
    
    queryString = "SELECT COUNT(*) FROM DOCS WHERE fBODY LIKE '%" & docNumber6 & "%' and fSTATE='1' and fNAME='N1Mort' "
    sql_Value = 1
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Call Log_Error_My()
        Log.Error "Querystring = " & queryString & ":  Expected result = " & sql_Value , "", pmNormal, attr
    End If
    BuiltIn.Delay(delay_small)
    
    queryString = "SELECT COUNT(*) FROM FOLDERS WHERE fISN= '" & fBASE6 & "'"
    sql_Value = 9
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Call Log_Error_My()
        Log.Error "Querystring = " & queryString & ":  Expected result = " & sql_Value , "", pmNormal, attr
    End If
    
    'Üáñ ³é³ñÏ³ ` ÷áË³¹ñ³ÙÇçáó
    AmountObject = "1000"
    Call Create_Object_Car(AmountObject)
    BuiltIn.Delay(delay_small)
    
    BuiltIn.Delay(1000)
    wMDIClient.vbObject("frmPttel").Close()
    
    'êï»ÕÍ»É Ýáñ å³ÛÙ³Ý³·Çñ ` "¶ñ³í (áëÏÇ)"
    agrType = "¶ñ³í(÷áË³¹ñ³ÙÇçáó)ª³íïáÙ³ï µ³óíáÕ"
    cliCode = "00034851"
    mortCurr = "001"
    MortSubject = 0
    Call Mortgage_Doc_Fill (agrType , pType, pNumber, cliCode, mortName, fillGrid, _
                            loanAgrType , loanAgrClient , loanAgrNum, mortCurr , mortSumma, mortCount , mortComment, _
                            sDate, fBASE7, docNumber7,mortageItemNew, MortSubject)
    
    queryString = "SELECT COUNT(*) FROM DOCS WHERE fBODY LIKE '%" & docNumber7 & "%' and fSTATE='1' and fNAME='N1Mort' "
    sql_Value = 1
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Call Log_Error_My()
        Log.Error "Querystring = " & queryString & ":  Expected result = " & sql_Value , "", pmNormal, attr
    End If
    BuiltIn.Delay(delay_small)
    
    queryString = "SELECT COUNT(*) FROM FOLDERS WHERE fISN= '" & fBASE7 & "'"
    sql_Value = 9
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Call Log_Error_My()
        Log.Error "Querystring = " & queryString & ":  Expected result = " & sql_Value , "", pmNormal, attr
    End If
    
    BuiltIn.Delay(1000)
    wMDIClient.vbObject("frmPttel").Close()
    
    'êï»ÕÍ»É Ýáñ å³ÛÙ³Ý³·Çñ ` "¶ñ³í ` ³íïáÙ³ï µ³óíáÕ"
    agrType = "¶ñ³í`³íïáÙ³ï µ³óíáÕ"
    cliCode = "00034851"
    mortCurr = "009"
    pType = "11"
    Call Mortgage_Doc_Fill (agrType , pType, pNumber, cliCode, mortName, fillGrid, _
                            loanAgrType , loanAgrClient , loanAgrNum, mortCurr , mortSumma, mortCount , mortComment, _
                            sDate, fBASE8, docNumber8,mortageItemNew, MortSubject)
    
    queryString = "SELECT COUNT(*) FROM DOCS WHERE fBODY LIKE '%" & docNumber8 & "%' and fSTATE='1' and fNAME='N1Mort' "
    sql_Value = 1
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Call Log_Error_My()
        Log.Error "Querystring = " & queryString & ":  Expected result = " & sql_Value , "", pmNormal, attr
        Exit Sub
    End If
    BuiltIn.Delay(delay_small)
    
    queryString = "SELECT COUNT(*) FROM FOLDERS WHERE fISN= '" & fBASE8 & "'"
    sql_Value = 9
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Call Log_Error_My()
        Log.Error "Querystring = " & queryString & ":  Expected result = " & sql_Value , "", pmNormal, attr
        Exit Sub
    End If
    
    BuiltIn.Delay(1000)
    wMDIClient.vbObject("frmPttel").Close()
    
    '1-ÇÝ å³ÛÙ³Ý³·ñÇ áõÕ³ñÏáõÙ Ñ³ëï³ïÙ³Ý "²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ" ÃÕÃ³å³Ý³ÏÇó
    isExist = Search_Mortgage_In_WorkPapers(docNumber)
    If Not isExist Then
        Call Log_Error_My()
        Log.Error "Document with number" & docNumber & " doesn't exist in workpapers forder" , "" , pmNormal, attr
        Exit Sub
    End If
    
    'ä³ÛÙ³Ý³·ÇñÁ áõÕ³ñÏ»É Ñ³ëï³ïÙ³Ý
    Call PaySys_Send_To_Verify()
    
    BuiltIn.Delay(1000)
    wMDIClient.vbObject("frmPttel").Close()
    
    '2-ñ¹  å³ÛÙ³Ý³·ñÇ áõÕ³ñÏáõÙ Ñ³ëï³ïÙ³Ý "²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ" ÃÕÃ³å³Ý³ÏÇó
    isExist = Search_Mortgage_In_WorkPapers(docNumber1)
    If Not isExist Then
        Call Log_Error_My()
        Log.Error "Document with number" & docNumber1 & " doesn't exist in workpapers forder" , "" , pmNormal, attr
        Exit Sub
    End If
    
    'ä³ÛÙ³Ý³·ÇñÁ áõÕ³ñÏ»É Ñ³ëï³ïÙ³Ý
    Call PaySys_Send_To_Verify()
    
    BuiltIn.Delay(1000)
    wMDIClient.vbObject("frmPttel").Close()
    
    '3-ñ¹ å³ÛÙ³Ý³·ñÇ áõÕ³ñÏáõÙ Ñ³ëï³ïÙ³Ý "²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ" ÃÕÃ³å³Ý³ÏÇó
    isExist = Search_Mortgage_In_WorkPapers(docNumber2)
    If Not isExist Then
        Call Log_Error_My()
        Log.Error "Document with number" & docNumber2 & " doesn't exist in workpapers forder" , "" , pmNormal, attr
        Exit Sub
    End If
    
    'ä³ÛÙ³Ý³·ÇñÁ áõÕ³ñÏ»É Ñ³ëï³ïÙ³Ý
    Call PaySys_Send_To_Verify()
    
    BuiltIn.Delay(1000)
    wMDIClient.vbObject("frmPttel").Close()
    
    '4-ñ¹ å³ÛÙ³Ý³·ñÇ áõÕ³ñÏáõÙ Ñ³ëï³ïÙ³Ý "²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ" ÃÕÃ³å³Ý³ÏÇó
    isExist = Search_Mortgage_In_WorkPapers(docNumber3)
    If Not isExist Then
        Call Log_Error_My()
        Log.Error "Document with number" & docNumber3 & " doesn't exist in workpapers forder" , "" , pmNormal, attr
        Exit Sub
    End If
    
    'ä³ÛÙ³Ý³·ÇñÁ áõÕ³ñÏ»É Ñ³ëï³ïÙ³Ý
    Call PaySys_Send_To_Verify()
    
    BuiltIn.Delay(1000)
    wMDIClient.vbObject("frmPttel").Close()
    
    '5-ñ¹ å³ÛÙ³Ý³·ñÇ áõÕ³ñÏáõÙ Ñ³ëï³ïÙ³Ý "²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ" ÃÕÃ³å³Ý³ÏÇó
    isExist = Search_Mortgage_In_WorkPapers(docNumber4)
    If Not isExist Then
        Call Log_Error_My()
        Log.Error "Document with number" & docNumber4 & " doesn't exist in workpapers forder" , "" , pmNormal, attr
        Exit Sub
    End If
    
    'ä³ÛÙ³Ý³·ÇñÁ áõÕ³ñÏ»É Ñ³ëï³ïÙ³Ý
    Call PaySys_Send_To_Verify()
    
    BuiltIn.Delay(1000)
    wMDIClient.vbObject("frmPttel").Close()
    
    '6-ñ¹ å³ÛÙ³Ý³·ñÇ áõÕ³ñÏáõÙ Ñ³ëï³ïÙ³Ý "²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ" ÃÕÃ³å³Ý³ÏÇó
    isExist = Search_Mortgage_In_WorkPapers(docNumber5)
    If Not isExist Then
        Call Log_Error_My()
        Log.Error "Document with number" & docNumber5 & " doesn't exist in workpapers forder" , "" , pmNormal, attr
        Exit Sub
    End If
    
    'ä³ÛÙ³Ý³·ÇñÁ áõÕ³ñÏ»É Ñ³ëï³ïÙ³Ý
    Call PaySys_Send_To_Verify()
    
    BuiltIn.Delay(1000)
    wMDIClient.vbObject("frmPttel").Close()
    
    '7-ñ¹ å³ÛÙ³Ý³·ñÇ áõÕ³ñÏáõÙ Ñ³ëï³ïÙ³Ý "²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ" ÃÕÃ³å³Ý³ÏÇó
    isExist = Search_Mortgage_In_WorkPapers(docNumber6)
    If Not isExist Then
        Call Log_Error_My()
        Log.Error "Document with number" & docNumber6 & " doesn't exist in workpapers forder" , "" , pmNormal, attr
        Exit Sub
    End If
    
    'ä³ÛÙ³Ý³·ÇñÁ áõÕ³ñÏ»É Ñ³ëï³ïÙ³Ý
    Call PaySys_Send_To_Verify()
    
    BuiltIn.Delay(1000)
    wMDIClient.vbObject("frmPttel").Close()
    
    '8-ñ¹ å³ÛÙ³Ý³·ñÇ áõÕ³ñÏáõÙ Ñ³ëï³ïÙ³Ý "²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ" ÃÕÃ³å³Ý³ÏÇó
    isExist = Search_Mortgage_In_WorkPapers(docNumber7)
    If Not isExist Then
        Call Log_Error_My()
        Log.Error "Document with number" & docNumber7 & " doesn't exist in workpapers forder" , "" , pmNormal, attr
        Exit Sub
    End If
    
    'ä³ÛÙ³Ý³·ÇñÁ áõÕ³ñÏ»É Ñ³ëï³ïÙ³Ý
    Call PaySys_Send_To_Verify()
    
    BuiltIn.Delay(1000)
    wMDIClient.vbObject("frmPttel").Close()
    
    '9-ñ¹ å³ÛÙ³Ý³·ñÇ áõÕ³ñÏáõÙ Ñ³ëï³ïÙ³Ý "²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ" ÃÕÃ³å³Ý³ÏÇó
    isExist = Search_Mortgage_In_WorkPapers(docNumber8)
    If Not isExist Then
        Call Log_Error_My()
        Log.Error "Document with number" & docNumber8 & " doesn't exist in workpapers forder" , "" , pmNormal, attr
        Exit Sub
    End If
    
    'ä³ÛÙ³Ý³·ÇñÁ áõÕ³ñÏ»É Ñ³ëï³ïÙ³Ý
    Call PaySys_Send_To_Verify()
    
    BuiltIn.Delay(1000)
    wMDIClient.vbObject("frmPttel").Close()
    
    '²ÝóáõÙ ¹»åÇ  "Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I"-Ç ÃÕÃ³å³Ý³Ï
    FolderName = "|êï³óí³Í ·ñ³í|Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I"
    Call GoTo_Folders(FolderName, docNumber)
    
    'ä³ÛÙ³Ý³·ñÇ ³éÏ³ÛáõÃÛ³Ý ëïáõ·áõÙ  "Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I"-Ç ÃÕÃ³å³Ý³ÏáõÙ
    ColNum = 2
    Is_Exist = False
    Is_Exist = Is_Agr_Exist(docNumber, ColNum)
    If Is_Exist Then
        TextMSG = "Agreement  is Exist in the Verifier's Folder " & docNumber
        Call Log_Print_My()
        Log.Message TextMSG , "", pmNormal, attr
    Else
        TextMSG = "Agreement  is'n Exist in the Verifier's Folder " & docNumber
        Call Log_Error_My()
        Log.Error TextMSG , "", pmNormal, attr
    End If
    
    'ä³ÛÙ³Ý³·ñÇ Ñ³ëï³ïáõÙ
    Call PaySys_Verify(True)
    
    BuiltIn.Delay(1000)
    wMDIClient.vbObject("frmPttel").Close()
    
    queryString = "SELECT COUNT(fBASE) FROM LINKEDAGRS WHERE fMORTISN= '" & fBASE & "'"
    sql_Value = 1
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Call Log_Error_My()
        Log.Error "Querystring = " & queryString & ":  Expected result = " & sql_Value , "", pmNormal, attr
    End If
    BuiltIn.Delay(delay_small)
    
    queryString = "SELECT COUNT(fISN) from FOLDERS WHERE fISN='" & fBASE & "' and fFOLDERID='NADDITINFO'"
    sql_Value = 1
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Call Log_Error_My()
        Log.Error "Querystring = " & queryString & ":  Expected result = " & sql_Value , "", pmNormal, attr
    End If
    BuiltIn.Delay(delay_small)
    
    queryString = "SELECT COUNT(fISN) FROM DOCS WHERE fBODY LIKE '%" & docNumber & "%' and fSTATE='7'"
    sql_Value = 2
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Call Log_Error_My()
        Log.Error "Querystring = " & queryString & ":  Expected result = " & sql_Value , "", pmNormal, attr
    End If
    
    '²ÝóáõÙ ¹»åÇ  "Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I"-Ç ÃÕÃ³å³Ý³Ï
    FolderName = "|êï³óí³Í ·ñ³í|Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I"
    Call GoTo_Folders(FolderName, docNumber1)
    
    'ä³ÛÙ³Ý³·ñÇ ³éÏ³ÛáõÃÛ³Ý ëïáõ·áõÙ  "Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I"-Ç ÃÕÃ³å³Ý³ÏáõÙ
    ColNum = 2
    Is_Exist = False
    Is_Exist = Is_Agr_Exist(docNumber1, ColNum)
    If Is_Exist Then
        TextMSG = "Agreement  is Exist in the Verifier's Folder " & docNumber1
        Call Log_Print_My()
        Log.Message TextMSG , "", pmNormal, attr
    Else
        TextMSG = "Agreement  is'n Exist in the Verifier's Folder " & docNumber1
        Call Log_Error_My()
        Log.Error TextMSG , "", pmNormal, attr
    End If
    
    'ä³ÛÙ³Ý³·ñÇ Ñ³ëï³ïáõÙ
    Call PaySys_Verify(True)
    
    BuiltIn.Delay(1000)
    wMDIClient.vbObject("frmPttel").Close()
    
    queryString = "SELECT COUNT(fBASE) FROM LINKEDAGRS WHERE fMORTISN= '" & fBASE1 & "'"
    sql_Value = 1
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Call Log_Error_My()
        Log.Error "Querystring = " & queryString & ":  Expected result = " & sql_Value , "", pmNormal, attr
    End If
    BuiltIn.Delay(delay_small)
    
    queryString = "SELECT COUNT(fISN) from FOLDERS WHERE fISN='" & fBASE1 & "' and fFOLDERID='NADDITINFO'"
    sql_Value = 1
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Call Log_Error_My()
        Log.Error "Querystring = " & queryString & ":  Expected result = " & sql_Value , "", pmNormal, attr
    End If
    BuiltIn.Delay(delay_small)
    
    queryString = "SELECT COUNT(fISN) FROM DOCS WHERE fBODY LIKE '%" & docNumber1 & "%' and fSTATE='7'"
    sql_Value = 2
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Call Log_Error_My()
        Log.Error "Querystring = " & queryString & ":  Expected result = " & sql_Value , "", pmNormal, attr
    End If
    
    '²ÝóáõÙ ¹»åÇ  "Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I"-Ç ÃÕÃ³å³Ý³Ï
    FolderName = "|êï³óí³Í ·ñ³í|Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I"
    Call GoTo_Folders(FolderName, docNumber2)
    
    'ä³ÛÙ³Ý³·ñÇ ³éÏ³ÛáõÃÛ³Ý ëïáõ·áõÙ  "Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I"-Ç ÃÕÃ³å³Ý³ÏáõÙ
    ColNum = 2
    Is_Exist = False
    Is_Exist = Is_Agr_Exist(docNumber2, ColNum)
    If Is_Exist Then
        TextMSG = "Agreement  is Exist in the Verifier's Folder " & docNumber2
        Call Log_Print_My()
        Log.Message TextMSG , "", pmNormal, attr
    Else
        TextMSG = "Agreement  is'n Exist in the Verifier's Folder " & docNumber2
        Call Log_Error_My()
        Log.Error TextMSG , "", pmNormal, attr
    End If
    
    'ä³ÛÙ³Ý³·ñÇ Ñ³ëï³ïáõÙ
    Call PaySys_Verify(True)
    
    BuiltIn.Delay(1000)
    wMDIClient.vbObject("frmPttel").Close()
    
    queryString = "SELECT COUNT(fBASE) FROM LINKEDAGRS WHERE fMORTISN= '" & fBASE2 & "'"
    sql_Value = 1
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Call Log_Error_My()
        Log.Error "Querystring = " & queryString & ":  Expected result = " & sql_Value , "", pmNormal, attr
    End If
    BuiltIn.Delay(delay_small)
    
    queryString = "SELECT COUNT(fISN) from FOLDERS WHERE fISN='" & fBASE2 & "' and fFOLDERID='NADDITINFO'"
    sql_Value = 1
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Call Log_Error_My()
        Log.Error "Querystring = " & queryString & ":  Expected result = " & sql_Value , "", pmNormal, attr
    End If
    BuiltIn.Delay(delay_small)
    
    queryString = "SELECT COUNT(fISN) FROM DOCS WHERE fBODY LIKE '%" & docNumber2 & "%' and fSTATE='7'"
    sql_Value = 2
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Call Log_Error_My()
        Log.Error "Querystring = " & queryString & ":  Expected result = " & sql_Value , "", pmNormal, attr
    End If
    
    '²ÝóáõÙ ¹»åÇ  "Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I"-Ç ÃÕÃ³å³Ý³Ï
    FolderName = "|êï³óí³Í ·ñ³í|Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I"
    Call GoTo_Folders(FolderName, docNumber3)
    
    'ä³ÛÙ³Ý³·ñÇ ³éÏ³ÛáõÃÛ³Ý ëïáõ·áõÙ  "Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I"-Ç ÃÕÃ³å³Ý³ÏáõÙ
    ColNum = 2
    Is_Exist = False
    Is_Exist = Is_Agr_Exist(docNumber3, ColNum)
    If Is_Exist Then
        TextMSG = "Agreement  is Exist in the Verifier's Folder " & docNumber3
        Call Log_Print_My()
        Log.Message TextMSG , "", pmNormal, attr
    Else
        TextMSG = "Agreement  is'n Exist in the Verifier's Folder " & docNumber3
        Call Log_Error_My()
        Log.Error TextMSG , "", pmNormal, attr
    End If
    
    'ä³ÛÙ³Ý³·ñÇ Ñ³ëï³ïáõÙ
    Call PaySys_Verify(True)
    
    BuiltIn.Delay(1000)
    wMDIClient.vbObject("frmPttel").Close()
    
    queryString = "SELECT COUNT(fBASE) FROM LINKEDAGRS WHERE fMORTISN= '" & fBASE3 & "'"
    sql_Value = 1
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Call Log_Error_My()
        Log.Error "Querystring = " & queryString & ":  Expected result = " & sql_Value , "", pmNormal, attr
    End If
    BuiltIn.Delay(delay_small)
    
    queryString = "SELECT COUNT(fISN) from FOLDERS WHERE fISN='" & fBASE3 & "' and fFOLDERID='NADDITINFO'"
    sql_Value = 1
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Call Log_Error_My()
        Log.Error "Querystring = " & queryString & ":  Expected result = " & sql_Value , "", pmNormal, attr
    End If
    BuiltIn.Delay(delay_small)
    
    queryString = "SELECT COUNT(fISN) FROM DOCS WHERE fBODY LIKE '%" & docNumber3 & "%' and fSTATE='7'"
    sql_Value = 2
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Call Log_Error_My()
        Log.Error "Querystring = " & queryString & ":  Expected result = " & sql_Value , "", pmNormal, attr
    End If
    
    '²ÝóáõÙ ¹»åÇ  "Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I"-Ç ÃÕÃ³å³Ý³Ï
    FolderName = "|êï³óí³Í ·ñ³í|Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I"
    Call GoTo_Folders(FolderName, docNumber4)
    
    'ä³ÛÙ³Ý³·ñÇ ³éÏ³ÛáõÃÛ³Ý ëïáõ·áõÙ  "Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I"-Ç ÃÕÃ³å³Ý³ÏáõÙ
    ColNum = 2
    Is_Exist = False
    Is_Exist = Is_Agr_Exist(docNumber4, ColNum)
    If Is_Exist Then
        TextMSG = "Agreement  is Exist in the Verifier's Folder " & docNumber4
        Call Log_Print_My()
        Log.Message TextMSG , "", pmNormal, attr
    Else
        TextMSG = "Agreement  is'n Exist in the Verifier's Folder " & docNumber4
        Call Log_Error_My()
        Log.Error TextMSG , "", pmNormal, attr
    End If
    
    'ä³ÛÙ³Ý³·ñÇ Ñ³ëï³ïáõÙ
    Call PaySys_Verify(True)
    
    BuiltIn.Delay(1000)
    wMDIClient.vbObject("frmPttel").Close()
    
    queryString = "SELECT COUNT(fBASE) FROM LINKEDAGRS WHERE fMORTISN= '" & fBASE4 & "'"
    sql_Value = 1
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Call Log_Error_My()
        Log.Error "Querystring = " & queryString & ":  Expected result = " & sql_Value , "", pmNormal, attr
    End If
    BuiltIn.Delay(delay_small)
    
    queryString = "SELECT COUNT(fISN) from FOLDERS WHERE fISN='" & fBASE4 & "' and fFOLDERID='NADDITINFO'"
    sql_Value = 1
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Call Log_Error_My()
        Log.Error "Querystring = " & queryString & ":  Expected result = " & sql_Value , "", pmNormal, attr
    End If
    BuiltIn.Delay(delay_small)
    
    queryString = "SELECT COUNT(fISN) FROM DOCS WHERE fBODY LIKE '%" & docNumber4 & "%' and fSTATE='7'"
    sql_Value = 2
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Call Log_Error_My()
        Log.Error "Querystring = " & queryString & ":  Expected result = " & sql_Value , "", pmNormal, attr
    End If
    
    '²ÝóáõÙ ¹»åÇ  "Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I"-Ç ÃÕÃ³å³Ý³Ï
    FolderName = "|êï³óí³Í ·ñ³í|Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I"
    Call GoTo_Folders(FolderName, docNumber5)
    
    'ä³ÛÙ³Ý³·ñÇ ³éÏ³ÛáõÃÛ³Ý ëïáõ·áõÙ  "Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I"-Ç ÃÕÃ³å³Ý³ÏáõÙ
    ColNum = 2
    Is_Exist = False
    Is_Exist = Is_Agr_Exist(docNumber5, ColNum)
    If Is_Exist Then
        TextMSG = "Agreement  is Exist in the Verifier's Folder " & docNumber5
        Call Log_Print_My()
        Log.Message TextMSG , "", pmNormal, attr
    Else
        TextMSG = "Agreement  is'n Exist in the Verifier's Folder " & docNumber5
        Call Log_Error_My()
        Log.Error TextMSG , "", pmNormal, attr
    End If
    
    'ä³ÛÙ³Ý³·ñÇ Ñ³ëï³ïáõÙ
    Call PaySys_Verify(True)
    
    BuiltIn.Delay(1000)
    wMDIClient.vbObject("frmPttel").Close()
    
    queryString = "SELECT COUNT(fBASE) FROM LINKEDAGRS WHERE fMORTISN= '" & fBASE5 & "'"
    sql_Value = 1
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Call Log_Error_My()
        Log.Error "Querystring = " & queryString & ":  Expected result = " & sql_Value , "", pmNormal, attr
    End If
    BuiltIn.Delay(delay_small)
    
    queryString = "SELECT COUNT(fISN) from FOLDERS WHERE fISN='" & fBASE5 & "' and fFOLDERID='NADDITINFO'"
    sql_Value = 1
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Call Log_Error_My()
        Log.Error "Querystring = " & queryString & ":  Expected result = " & sql_Value , "", pmNormal, attr
    End If
    BuiltIn.Delay(delay_small)
    
    queryString = "SELECT COUNT(fISN) FROM DOCS WHERE fBODY LIKE '%" & docNumber5 & "%' and fSTATE='7'"
    sql_Value = 2
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Call Log_Error_My()
        Log.Error "Querystring = " & queryString & ":  Expected result = " & sql_Value , "", pmNormal, attr
    End If
    
    '²ÝóáõÙ ¹»åÇ  "Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I"-Ç ÃÕÃ³å³Ý³Ï
    FolderName = "|êï³óí³Í ·ñ³í|Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I"
    Call GoTo_Folders(FolderName, docNumber6)
    
    'ä³ÛÙ³Ý³·ñÇ ³éÏ³ÛáõÃÛ³Ý ëïáõ·áõÙ  "Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I"-Ç ÃÕÃ³å³Ý³ÏáõÙ
    ColNum = 2
    Is_Exist = False
    Is_Exist = Is_Agr_Exist(docNumber6, ColNum)
    If Is_Exist Then
        TextMSG = "Agreement  is Exist in the Verifier's Folder " & docNumber6
        Call Log_Print_My()
        Log.Message TextMSG , "", pmNormal, attr
    Else
        TextMSG = "Agreement  is'n Exist in the Verifier's Folder " & docNumber6
        Call Log_Error_My()
        Log.Error TextMSG , "", pmNormal, attr
    End If
    
    'ä³ÛÙ³Ý³·ñÇ Ñ³ëï³ïáõÙ
    Call PaySys_Verify(True)
    
    BuiltIn.Delay(1000)
    wMDIClient.vbObject("frmPttel").Close()
    
    queryString = "SELECT COUNT(fBASE) FROM LINKEDAGRS WHERE fMORTISN= '" & fBASE6 & "'"
    sql_Value = 1
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Call Log_Error_My()
        Log.Error "Querystring = " & queryString & ":  Expected result = " & sql_Value , "", pmNormal, attr
    End If
    BuiltIn.Delay(delay_small)
    
    queryString = "SELECT COUNT(fISN) from FOLDERS WHERE fISN='" & fBASE6 & "' and fFOLDERID='NADDITINFO'"
    sql_Value = 1
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Call Log_Error_My()
        Log.Error "Querystring = " & queryString & ":  Expected result = " & sql_Value , "", pmNormal, attr
    End If
    BuiltIn.Delay(delay_small)
    
    queryString = "SELECT COUNT(fISN) FROM DOCS WHERE fBODY LIKE '%" & docNumber6 & "%' and fSTATE='7'"
    sql_Value = 2
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Call Log_Error_My()
        Log.Error "Querystring = " & queryString & ":  Expected result = " & sql_Value , "", pmNormal, attr
    End If
    
    '²ÝóáõÙ ¹»åÇ  "Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I"-Ç ÃÕÃ³å³Ý³Ï
    FolderName = "|êï³óí³Í ·ñ³í|Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I"
    Call GoTo_Folders(FolderName, docNumber7)
    
    'ä³ÛÙ³Ý³·ñÇ ³éÏ³ÛáõÃÛ³Ý ëïáõ·áõÙ  "Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I"-Ç ÃÕÃ³å³Ý³ÏáõÙ
    ColNum = 2
    Is_Exist = False
    Is_Exist = Is_Agr_Exist(docNumber7, ColNum)
    If Is_Exist Then
        TextMSG = "Agreement  is Exist in the Verifier's Folder " & docNumber7
        Call Log_Print_My()
        Log.Message TextMSG , "", pmNormal, attr
    Else
        TextMSG = "Agreement  is'n Exist in the Verifier's Folder " & docNumber7
        Call Log_Error_My()
        Log.Error TextMSG , "", pmNormal, attr
    End If
    
    'ä³ÛÙ³Ý³·ñÇ Ñ³ëï³ïáõÙ
    Call PaySys_Verify(True)
    
    BuiltIn.Delay(1000)
    wMDIClient.vbObject("frmPttel").Close()
    
    queryString = "SELECT COUNT(fBASE) FROM LINKEDAGRS WHERE fMORTISN= '" & fBASE7 & "'"
    sql_Value = 1
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Call Log_Error_My()
        Log.Error "Querystring = " & queryString & ":  Expected result = " & sql_Value , "", pmNormal, attr
    End If
    BuiltIn.Delay(delay_small)
    
    queryString = "SELECT COUNT(fISN) from FOLDERS WHERE fISN='" & fBASE7 & "' and fFOLDERID='NADDITINFO'"
    sql_Value = 1
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Call Log_Error_My()
        Log.Error "Querystring = " & queryString & ":  Expected result = " & sql_Value , "", pmNormal, attr
    End If
    BuiltIn.Delay(delay_small)
    
    queryString = "SELECT COUNT(fISN) FROM DOCS WHERE fBODY LIKE '%" & docNumber7 & "%' and fSTATE='7'"
    sql_Value = 2
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Call Log_Error_My()
        Log.Error "Querystring = " & queryString & ":  Expected result = " & sql_Value , "", pmNormal, attr
    End If
    
    '²ÝóáõÙ ¹»åÇ  "Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I"-Ç ÃÕÃ³å³Ý³Ï
    FolderName = "|êï³óí³Í ·ñ³í|Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I"
    Call GoTo_Folders(FolderName, docNumber8)
    
    'ä³ÛÙ³Ý³·ñÇ ³éÏ³ÛáõÃÛ³Ý ëïáõ·áõÙ  "Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I"-Ç ÃÕÃ³å³Ý³ÏáõÙ
    ColNum = 2
    Is_Exist = False
    Is_Exist = Is_Agr_Exist(docNumber8, ColNum)
    If Is_Exist Then
        TextMSG = "Agreement  is Exist in the Verifier's Folder " & docNumber8
        Call Log_Print_My()
        Log.Message TextMSG , "", pmNormal, attr
    Else
        TextMSG = "Agreement  is'n Exist in the Verifier's Folder " & docNumber8
        Call Log_Error_My()
        Log.Error TextMSG , "", pmNormal, attr
    End If
    
    'ä³ÛÙ³Ý³·ñÇ Ñ³ëï³ïáõÙ
    Call PaySys_Verify(True)
    
    BuiltIn.Delay(1000)
    wMDIClient.vbObject("frmPttel").Close()
    
    queryString = "SELECT COUNT(fBASE) FROM LINKEDAGRS WHERE fMORTISN= '" & fBASE8 & "'"
    sql_Value = 1
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Call Log_Error_My()
        Log.Error "Querystring = " & queryString & ":  Expected result = " & sql_Value , "", pmNormal, attr
    End If
    BuiltIn.Delay(delay_small)
    
    queryString = "SELECT COUNT(fISN) from FOLDERS WHERE fISN='" & fBASE8 & "' and fFOLDERID='NADDITINFO'"
    sql_Value = 1
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Call Log_Error_My()
        Log.Error "Querystring = " & queryString & ":  Expected result = " & sql_Value , "", pmNormal, attr
    End If
    BuiltIn.Delay(delay_small)
    
    queryString = "SELECT COUNT(fISN) FROM DOCS WHERE fBODY LIKE '%" & docNumber8 & "%' and fSTATE='7'"
    sql_Value = 2
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Call Log_Error_My()
        Log.Error "Querystring = " & queryString & ":  Expected result = " & sql_Value , "", pmNormal, attr
    End If
    
    Call Login("ARMSOFT")
    Call ChangeWorkspace(c_Loans)
    
    ' ì³ñÏ³ÛÇÝ å³ÛÙ³Ý³·ñÇ ÷ÝïñáõÙ "ä³ÛÙ³Ý³·ñ»ñ" ÃÕÃ³å³Ý³ÏáõÙ
    docExist = Contracts_Filter_Fill("2", loanAgrNum(1), "|ì³ñÏ»ñ (ï»Õ³µ³ßËí³Í)|ä³ÛÙ³Ý³·ñ»ñ")
    If Not docExist Then
        Call Log_Error_My()
        Log.Error "Document with number " & loanAgrNum(1) & "does'n exist" , "" , pmNormal, attr
    End If
    
    '²é³çÇÝ ·ñ³íÇ å³ÛÙ³Ý³·ñÇ ÷ÝïñáõÙ "Ð³×³Ëáñ¹Ç ÃÕÃ³å³Ý³Ï"  - áõÙ
    docExist = Search_Morgage(docNumber)
    If Not docExist Then
        Call Log_Error_My()
        Log.Error "MortGage with number " & docNumber & "does'n exist in clients workpapers" , "" , pmNormal, attr
    End If
    
    'ºñÏñáñ¹ ·ñ³íÇ å³ÛÙ³Ý³·ñÇ ÷ÝïñáõÙ "Ð³×³Ëáñ¹Ç ÃÕÃ³å³Ý³Ï"  - áõÙ
    docExist = Search_Morgage(docNumber1)
    If Not docExist Then
        Call Log_Error_My()
        Log.Error "MortGage with number " & docNumber1 & "does'n exist in clients workpapers" , "" , pmNormal, attr
    End If
    
    'ºññáñ¹ ·ñ³íÇ å³ÛÙ³Ý³·ñÇ ÷ÝïñáõÙ "Ð³×³Ëáñ¹Ç ÃÕÃ³å³Ý³Ï"  - áõÙ
    docExist = Search_Morgage(docNumber2)
    If Not docExist Then
        Call Log_Error_My()
        Log.Error "MortGage with number " & docNumber2 & "does'n exist in clients workpapers" , "" , pmNormal, attr
    End If
    
    'âáññáñ¹ ·ñ³íÇ å³ÛÙ³Ý³·ñÇ ÷ÝïñáõÙ "Ð³×³Ëáñ¹Ç ÃÕÃ³å³Ý³Ï"  - áõÙ
    docExist = Search_Morgage(docNumber3)
    If Not docExist Then
        Call Log_Error_My()
        Log.Error "MortGage with number " & docNumber3 & "does'n exist in clients workpapers" , "" , pmNormal, attr
    End If
    
    'ÐÇÝ·»ñáñ¹ ·ñ³íÇ å³ÛÙ³Ý³·ñÇ ÷ÝïñáõÙ "Ð³×³Ëáñ¹Ç ÃÕÃ³å³Ý³Ï"  - áõÙ
    docExist = Search_Morgage(docNumber4)
    If Not docExist Then
        Call Log_Error_My()
        Log.Error "MortGage with number " & docNumber4 & "does'n exist in clients workpapers" , "" , pmNormal, attr
    End If
    
    'ì»ó»ñáñ¹ ·ñ³íÇ å³ÛÙ³Ý³·ñÇ ÷ÝïñáõÙ "Ð³×³Ëáñ¹Ç ÃÕÃ³å³Ý³Ï"  - áõÙ
    docExist = Search_Morgage(docNumber5)
    If Not docExist Then
        Call Log_Error_My()
        Log.Error "MortGage with number " & docNumber5 & "does'n exist in clients workpapers" , "" , pmNormal, attr
    End If
    
    'ÚáÃ»ñáñ¹ ·ñ³íÇ å³ÛÙ³Ý³·ñÇ ÷ÝïñáõÙ "Ð³×³Ëáñ¹Ç ÃÕÃ³å³Ý³Ï"  - áõÙ
    docExist = Search_Morgage(docNumber6)
    If Not docExist Then
        Call Log_Error_My()
        Log.Error "MortGage with number " & docNumber6 & "does'n exist in clients workpapers" , "" , pmNormal, attr
    End If
    
    'àõÃ»ñáñ¹ ·ñ³íÇ å³ÛÙ³Ý³·ñÇ ÷ÝïñáõÙ "Ð³×³Ëáñ¹Ç ÃÕÃ³å³Ý³Ï"  - áõÙ
    docExist = Search_Morgage(docNumber7)
    If Not docExist Then
        Call Log_Error_My()
        Log.Error "MortGage with number " & docNumber7 & "does'n exist in clients workpapers" , "" , pmNormal, attr
    End If
    
    'ÆÝ»ñáñ¹ ·ñ³íÇ å³ÛÙ³Ý³·ñÇ ÷ÝïñáõÙ "Ð³×³Ëáñ¹Ç ÃÕÃ³å³Ý³Ï"  - áõÙ
    docExist = Search_Morgage(docNumber8)
    If Not docExist Then
        Call Log_Error_My()
        Log.Error "MortGage with number " & docNumber8 & "does'n exist in clients workpapers" , "" , pmNormal, attr
    End If
    
    'Test CleanUp
    Call Close_AsBank()
End Sub