'USEUNIT Library_Common
'USEUNIT Payment_Order_ConfirmPhases_Library
'USEUNIT Payment_Except_Library
'USEUNIT Online_PaySys_Library
'USEUNIT Loan_Agreements_Library
'USEUNIT Loan_Agreemnts_With_Schedule_Library
'USEUNIT Contract_Summary_Report_Library
'USEUNIT Subsystems_SQL_Library
'USEUNIT Constants

'Test case ID 165643

Sub Credit_With_Schedule_Rent_Test()
    
    Dim fDATE, data, startDATE , arrFOP, arrVal, percSumma, work, aCon, aCmd1, calcPRBase1, fadeBase, calcPRBase, queryString, giveCrBase, fBaseCP, fDate1, isExists, docNumber, fISN, actionCount, dateStart, dateEnd
    Dim clientCode, tmpltype, curr, accacc, summ, date_arg, dateFillType, fadeDate, finishFadeDate
    Dim passDirection, sumDates, i, sumFill, pCalcDate, agrIntRate, agrIntRatePart, branch, sector, schedule
    Dim guarante, startFadeDate, district, paperCode, fBASE, docExist, isEqual, round, percent
    Dim dategive, dateconcl, calcPRBase2, calcPRBase3, calcPRBase4, rpBase , ccalcBase, fBase1
    Dim wrBase, opTp, fBase2 , yldBase, note, aim, beforeTerm, fadeBase1, newSchedule,sum
    
    Utilities.ShortDateFormat = "yyyymmdd"
    startDATE = "20030101"
    fDATE = "20250101"
    clientCode = "00034851"
    curr = Null
    accacc = "03485190101"
    summ = "10,000.00"
    dateconcl = "25/12/12"
    data = "25/12/12"
    dategive = "25/12/12"
    date_arg = "25/12/13"
    dateFillType = "1"
    fadeDate = "25"
    startFadeDate = "25/12/12"
    finishFadeDate = "25/12/13"
    passDirection = "2"
    sumDates = "1"
    sumFill = "04"
    round = "2"
    agrIntRate = "19"
    agrIntRatePart = "365"
    branch = "9"
    sector = "U2"
    aim = "00"
    schedule = "9"
    guarante = "9"
    district = "001"
    paperCode = "12"
    percent = "1,059.01"
    pCalcDate = "19/01/13"
    fDate1 = "20/01/13"
    percSumma = Null
    opTp = "22"
    note = "01"
    arrFOP = Array("AGR", "DBT" , "INC" , "LET" , "OUT" , "PAY" , "PER", "RAC", "RES", "RET", "RTP")
    arrVal = Array("100000.00", "22088.66", "5248552.66", "170201.20", "39455416.56", "40100.00" , "18521.86", "5.00", "39357024.00", "5000.00", "4000.00" )
    newSchedule = True
    
    'Test StartUp 
    Call Initialize_AsBank("bank", startDATE, fDATE)
    Call Login("CREDITOPERATOR")
    Call Create_Connection()
    
    '¶ñ³ýÇÏáí í³ñÏ³ÛÇÝ å³ÛÙ³Ý³·ñÇ ëï»ÕÍáõÙ
    Call Select_Credit_Type("¶ñ³ýÇÏáí í³ñÏ³ÛÇÝ å³ÛÙ³Ý³·Çñ")
    Call Credit_With_Schedule_Doc_Fill(clientCode, tmpltype, curr, accacc, summ, dateconcl, dategive, date_arg, dateFillType, fadeDate, _
                                       finishFadeDate, startFadeDate, passDirection, sumDates, sumFill, round, agrIntRate, _
                                       agrIntRatePart, pcnotchoose , pcGrant , pcPenAgr, pcPenPer , part, _
                                       branch, sector, aim, schedule, guarante, district, note, paperCode, fBASE, docNumber)
    'Ø³ñÙ³Ý ·ñ³ýÇÏÇ Ýß³Ý³ÏáõÙ
    BuiltIn.Delay(2000)
    docExist = Fade_Schedule()
    If Not docExist Then
        Log.Error("Cannot create fade schedule")
        Exit Sub
    End If
    
    'Ø³ñÙ³Ý ·ñ³ýÇÏÇ ·áõÙ³ñ ¨ ïáÏáë ¹³ßï»ñÇ ³ñÅ»ùÝ»ñÇ ëïáõ·áõÙ
    isEqual = Compare_FadeSchedule_Values (summ, percent, newSchedule)
    If Not isEqual Then
        Log.Error("Fading schedule values are wrong")
    End If
    
    'ä³ÛÙ³Ý³·ñÇ ÷ÝïñáõÙ å³ÛÙ³Ý³·ñÇ ÃÕÃ³å³Ý³Ï³áõÙ
    wMDIClient.vbObject("frmPttel").vbObject("tdbgView").MoveFirst
    Do Until wMDIClient.vbObject("frmPttel").vbObject("tdbgView").EOF
        If Left(Trim(wMDIClient.vbObject("frmPttel").vbObject("tdbgView").Text), 28) = "¶ñ³ýÇÏáí í³ñÏ³ÛÇÝ å³ÛÙ³Ý³·Çñ" Then
            Exit Do
        Else
            Call wMDIClient.vbObject("frmPttel").vbObject("tdbgView").MoveNext
        End If
    Loop
    
    'ì³ñÓ³í×³ñÇ ·ñ³ýÇÏÇ Ýß³Ý³ÏáõÙ
    docExist = Rent_Schedule()
    If Not docExist Then
        Log.Error("Cannot create fade schedule")
        Exit Sub
    End If
    
    'ä³ÛÙ³Ý³·ñÇ ÷ÝïñáõÙ å³ÛÙ³Ý³·ñÇ ÃÕÃ³å³Ý³Ï³áõÙ
    wMDIClient.vbObject("frmPttel").vbObject("tdbgView").MoveFirst
    Do Until wMDIClient.vbObject("frmPttel").vbObject("tdbgView").EOF
        If Left(Trim(wMDIClient.vbObject("frmPttel").vbObject("tdbgView").Text), 28) = "¶ñ³ýÇÏáí í³ñÏ³ÛÇÝ å³ÛÙ³Ý³·Çñ" Then
            Exit Do
        Else
            Call wMDIClient.vbObject("frmPttel").vbObject("tdbgView").MoveNext
        End If
    Loop
    
    '²ÛÉ í×³ñáõÙÝ»ñÇ ·ñ³ýÇÏÇ Ýß³Ý³ÏáõÙ
    docExist = Other_Payment_Schedule(date_arg, "1000")
    If Not docExist Then
        Log.Error("Cannot create payment schedule")
        Exit Sub
    End If
    
    'ä³ÛÙ³Ý³·ñÇ ÷ÝïñáõÙ å³ÛÙ³Ý³·ñÇ ÃÕÃ³å³Ý³Ï³áõÙ
    wMDIClient.vbObject("frmPttel").vbObject("tdbgView").MoveFirst
    Do Until wMDIClient.vbObject("frmPttel").vbObject("tdbgView").EOF
        If Left(Trim(wMDIClient.vbObject("frmPttel").vbObject("tdbgView").Text), 28) = "¶ñ³ýÇÏáí í³ñÏ³ÛÇÝ å³ÛÙ³Ý³·Çñ" Then
            Exit Do
        Else
            Call wMDIClient.vbObject("frmPttel").vbObject("tdbgView").MoveNext
        End If
    Loop
    
    'ä³ÛÙ³Ý³·ñÇ áõÕ³ñÏáõÙ Ñ³ëï³ïÙ³Ý
    Call PaySys_Send_To_Verify()
    
    'ä³ÛÙ³Ý³·ñÇ ³éÏ³ÛáõÃÛ³Ý ëïáõ·áõÙ Ñ³ëï³ïíáÕ ÷³ëïÃÕÃ»ñ 1 ÃÕÃ³å³Ý³ÏáõÙ
    Call Login("ARMSOFT")
    Call ChangeWorkspace(c_Loans)
    docExist = Verify_Credit(docNumber)
    If Not docExist Then
        Log.Error("The document doesn't exist in verifier folder")
    End If
    
    'ö³ëï³ÃÕÃÇ Ñ³ëï³ïáõÙ
    Call PaySys_Verify(True)
    
    BuiltIn.Delay(1000)
    wMDIClient.vbObject("frmPttel").Close()

    'ö³ëï³ÃÕÃÇ ³éÏ³ÛáõÃÛ³Ý ëïáõ·áõÙ "ä³ÛÙ³Ý³·ñ»ñ" ÃÕÃ³å³Ý³ÏáõÙ
    docExist = Contracts_Filter_Fill("1", docNumber, "|ì³ñÏ»ñ (ï»Õ³µ³ßËí³Í)|ä³ÛÙ³Ý³·ñ»ñ")
    If Not docExist Then
        Log.Error("The document doesn't exist in payments folder ")
        Exit Sub
    End If
    
    ' ¶³ÝÓáõÙ ïñ³Ù³¹ñáõÙÇó ·áñÍáÕáõÃÛ³Ý Ï³ï³ñáõÙ
    Call Collect_From_Provision(data, sum, "2", Null, fBaseCP)
    
    'ì³ñÏÇ ïñ³Ù³¹ñáõÙ ·áñÍáÕáõÃÛ³Ý Ï³ï³ñáõÙ
    Call Give_Credit(data, summ, "2", accacc, giveCrBase)
    
    'îáÏáëÝ»ñÇ Ñ³ßí³ñÏáõÙ ·áñÍáÕáõÃÛ³Ý Ï³ï³ñáõÙ
    calcPRBase1 = Calculate_Percents(pCalcDate, pCalcDate, False)
        
    'ÊÙµ³ÛÇÝ í³ñÓ³í×³ñÇ Ñ³ßí³ñÏáõÙ ·áñÍáÕáõÃÛ³Ý Ï³ï³ñáõÙ
    pCalcDate = "19/01/13"
    Call Percent_Group_Calculate(pCalcDate, pCalcDate, True, False)
    
    date1 = "25/01/13"
    beforeTerm = True
    'ì³ñÏÇ å³ñïù»ñÇ Ù³ñáõÙ ·áñÍáÕáõÃÛ³Ý Ï³ï³ñáõÙ
    Call Fade_Debt(fDate1, fadeBase1, date1, null, percSumma, beforeTerm)
    
    BuiltIn.Delay(1000)
    wMDIClient.vbObject("frmPttel").Close()
    
    Call wTreeView.DblClickItem("|ì³ñÏ»ñ (ï»Õ³µ³ßËí³Í)|ä³ÛÙ³Ý³·ñ»ñÇ ³Ù÷á÷áõÙ")
    Call Contract_Sammary_Report_Fill(fDate1, Null, Null, Null, docNumber, Null, Null, Null, _
                                      Null, Null, Null, Null, Null, Null, Null, _
                                      Null, Null, Null, Null, Null, Null, Null, False, False, _
                                      Null, False, False, False, _
                                      False, False, False, False, False, _
                                      True, False, True, False, False, False, _
                                      False, False, False, False, False, False, False, 1)
    
    'îáÏáë ëÛ³Ý ³ñÅ»ùÇ ëïáõ·áõÙ
    If Trim(wMDIClient.vbObject("frmPttel").vbObject("tdbgView").Columns.Item(5)) <> "-24.05" Then
        Log.Error("Wrong  percent")
    End If
    
    'ì³ñÓ³í×³ñ ëÛ³Ý ³ñÅ»ùÇ ëïáõ·áõÙ
    If Trim(wMDIClient.vbObject("frmPttel").vbObject("tdbgView").Columns.Item(9)) <> "-100.00" Then
        Log.Error("Wrong value of rent : actual = "  & Sys.Process("Asbank").vbObject("MainForm").Window("MDIClient", "", 1).vbObject("frmPttel").vbObject("tdbgView").Columns.Item(8))
        Exit Sub
    End If
    
    'îáÏáëÝ»ñÇ Ñ³ßí³ñÏáõÙ ·áñÍáÕáõÃÛ³Ý Ï³ï³ñáõÙ
    pCalcDate = "24/01/13"
    calcPRBase1 = Calculate_Percents(pCalcDate, pCalcDate, False)
    
    'ÊÙµ³ÛÇÝ í³ñÓ³í×³ñÇ Ñ³ßí³ñÏáõÙ ·áñÍáÕáõÃÛ³Ý Ï³ï³ñáõÙ
    Call Percent_Group_Calculate(pCalcDate, pCalcDate, True, False)
    
    BuiltIn.Delay(1000)
    wMDIClient.vbObject("frmPttel").Close()
    
    Call wTreeView.DblClickItem("|ì³ñÏ»ñ (ï»Õ³µ³ßËí³Í)|ä³ÛÙ³Ý³·ñ»ñÇ ³Ù÷á÷áõÙ")
    pCalcDate  = "25/01/13"
    Call Contract_Sammary_Report_Fill(pCalcDate, Null, Null, Null, docNumber, Null, Null, Null, _
                                      Null, Null, Null, Null, Null, Null, Null, _
                                      Null, Null, Null, Null, Null, Null, Null, False, False, _
                                      Null, False, False, False, _
                                      False, False, False, False, False, _
                                      True, False, True, False, False, False, _
                                      False, False, False, False, False, False, False, 1)
    
    'îáÏáë ëÛ³Ý ³ñÅ»ùÇ ëïáõ·áõÙ
    If Trim(wMDIClient.vbObject("frmPttel").vbObject("tdbgView").Columns.Item(5)) <> "0.00" Then
        Log.Error("Wrong  percent")
    End If
    
    'îáÏáëÝ»ñÇ Ñ³ßí³ñÏáõÙ ·áñÍáÕáõÃÛ³Ý Ï³ï³ñáõÙ
    pCalcDate = "24/02/13"
    calcPRBase1 = Calculate_Percents(pCalcDate, pCalcDate, False)
    
    'ÊÙµ³ÛÇÝ í³ñÓ³í×³ñÇ Ñ³ßí³ñÏáõÙ ·áñÍáÕáõÃÛ³Ý Ï³ï³ñáõÙ
    Call Percent_Group_Calculate(pCalcDate, pCalcDate, True, False)
    
    'ì³ñÏÇ å³ñïù»ñÇ Ù³ñáõÙ ·áñÍáÕáõÃÛ³Ý Ï³ï³ñáõÙ
    pCalcDate = "25/02/13"
    Call Fade_Debt(pCalcDate, fadeBase, Null, null, percSumma, False)
    
    'îáÏáëÝ»ñÇ Ñ³ßí³ñÏáõÙ ·áñÍáÕáõÃÛ³Ý Ï³ï³ñáõÙ
    pCalcDate = "20/03/13"
    calcPRBase1 = Calculate_Percents(pCalcDate, pCalcDate, False)
    
    'ÊÙµ³ÛÇÝ í³ñÓ³í×³ñÇ Ñ³ßí³ñÏáõÙ ·áñÍáÕáõÃÛ³Ý Ï³ï³ñáõÙ
    pCalcDate = "21/03/13"
    Call Percent_Group_Calculate(pCalcDate, pCalcDate, True, False)
    
    'ì³ñÏÇ å³ñïù»ñÇ Ù³ñáõÙ ·áñÍáÕáõÃÛ³Ý Ï³ï³ñáõÙ
    date1 = "25/12/13"
    Call Fade_Debt(pCalcDate, fadeBase, date1, null, percSumma, True)
    
    BuiltIn.Delay(1000)
    wMDIClient.vbObject("frmPttel").Close()
    
    Call wTreeView.DblClickItem("|ì³ñÏ»ñ (ï»Õ³µ³ßËí³Í)|ä³ÛÙ³Ý³·ñ»ñÇ ³Ù÷á÷áõÙ")
    Call Contract_Sammary_Report_Fill(pCalcDate, Null, Null, Null, docNumber, Null, Null, Null, _
                                      Null, Null, Null, Null, Null, Null, Null, _
                                      Null, Null, Null, Null, Null, Null, Null, False, False, _
                                      Null, False, False, False, _
                                      False, False, False, False, False, _
                                      True, False, True, False, False, False, _
                                      False, False, False, False, False, False, False, 1)
    
    'îáÏáë ëÛ³Ý ³ñÅ»ùÇ ëïáõ·áõÙ
    If Trim(wMDIClient.vbObject("frmPttel").vbObject("tdbgView").Columns.Item(5)) <> "0.00" Then
        Log.Error("Wrong  percent")
    End If
    
    'ì³ñÓ³í×³ñ ëÛ³Ý ³ñÅ»ùÇ ëïáõ·áõÙ
    If Trim(wMDIClient.vbObject("frmPttel").vbObject("tdbgView").Columns.Item(9)) <> "0.00" Then
        Log.Error("Wrong value of rent")
    End If
    
    BuiltIn.Delay(1000)
    
     queryString = "select COUNT(*) from AGRSCHEDULE where fAGRISN= '" & fBASE & "'"
    sql_Value = 6
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    queryString = "select SUM(fINC), SUM(fKIND) from AGRSCHEDULE where fAGRISN='" & fBASE & "'"
    sql_Value = 16
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    sql_Value = 22
    colNum = 1
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    queryString = "select COUNT(*) from .AGRSCHEDULEVALUES where fAGRISN= '" & fBASE & "'"
    sql_Value = 104
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    queryString = "select SUM(fSUM), SUM(fVALUETYPE)  from .AGRSCHEDULEVALUES where fAGRISN= '" & fBASE & "'"
    sql_Value = 41909.12
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    sql_Value = 222
    colNum = 1
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    queryString = "select COUNT(*) from  dbo.HIF where fOBJECT='" & fBASE & "'"
    sql_Value = 26
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    queryString = "select SUM(fSUM), SUM(fCURSUM) from  dbo.HIF where fOBJECT= '" & fBASE & "'"
    sql_Value = 10088.4326 
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    sql_Value = 1100.00
    colNum = 1
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    queryString = "select COUNT(*) from  dbo.HIR where fOBJECT='" & fBASE & "'" 
    Log.Message(fBASE)
    sql_Value =  31
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    queryString =  "select SUM(fLASTREM) from  dbo.HIRREST where fTYPE = 'R1' and fOBJECT= '" & fBASE & "'"
    sql_Value = 0.00
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    queryString =  "select SUM(fLASTREM) from  dbo.HIRREST where fOBJECT= '" & fBASE & "'"
    sql_Value = 4000
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If

    queryString =  "select SUM(fLASTREM) from  dbo.HIRREST where fTYPE = 'R2' and fOBJECT= '" & fBASE & "'"
    sql_Value = 0.00
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    queryString = "select SUM(fCURSUM) from  dbo.HIR where fOBJECT= '" & fBASE & "'"
    sql_Value = 27891.74 ' 19163.98
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    queryString = "select COUNT(*) from  dbo.HIT where fOBJECT= '" & fBASE & "' and fTYPE='N2' and fOP='PER'"
    sql_Value = 4
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    queryString = "select COUNT(*) from HI where fDATE='2013-01-20' and fBASE='" & fadeBase1 & "'and fSUM=40000.00 and fCURSUM=40000.00 and fADB=1629496 and fACR=230416894 and fOP='MSC'"
    sql_Value = 2
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    queryString = "select COUNT(*) from HI where fDATE='2013-01-20' and fBASE='" & fadeBase1 & "'and fSUM=40000.00 and fCURSUM=40000.00 and fADB=82335686 and fOP='FEX' "
    sql_Value = 1
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    queryString = "select COUNT(*) from HI where fDATE='2013-01-20' and fBASE='" & fadeBase1 & "'and fSUM=100 and fCURSUM=40000.00 and fOP='SAL' and fTYPE='CE'"
    sql_Value = 1
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    queryString = "select COUNT(*) from HI where fDATE='2013-01-20' and fBASE='" & fadeBase1 & "'and fSUM=40000.00 and fCURSUM=100.00 and fOP='FEX' and fADB=82335686 "
    sql_Value = 1
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    queryString = "select COUNT(*) from HI where fDATE='2013-03-21' and fBASE='" & fadeBase & "'and fSUM=40000.00 and fCURSUM=40000.00 and fADB=1629496 and fACR=230416894 and fOP='MSC'"
    sql_Value = 2
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    queryString = "select COUNT(*) from HI where fDATE='2013-03-21' and fBASE='" & fadeBase & "'and fSUM=40000.00 and fCURSUM=40000.00 and fADB=82335686  and fOP='FEX'"
    sql_Value = 1
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    queryString = "select COUNT(*) from HI where fDATE='2013-03-21' and fBASE='" & fadeBase & "'and fSUM=100 and fCURSUM=40000.00 and fOP='SAL' and fTYPE='CE'"
    sql_Value = 1
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    queryString = "select COUNT(*) from HI where fDATE='2013-03-21' and fBASE='" & fadeBase & "'and fSUM=40000.00 and fCURSUM=100.00 and fOP='FEX' and fADB=82335686 "
    sql_Value = 1
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    queryString = "select COUNT(*) from HI where fDATE='2013-03-21' and fBASE='" & fadeBase & "'and fSUM=58128.00 and fCURSUM=145.32 and fOP='MSC' and fADB=1629496 "
    sql_Value = 1
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    queryString = "select COUNT(*) from HI where fDATE='2013-03-21' and fBASE='" & fadeBase & "'and fSUM=58128.00 and fCURSUM=58128.00 and fOP='MSC' and fADB=1629496 "
    sql_Value = 1
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    'ì³ñÏÇ Ù³ñáõÙ ·áñÍáÕáõÃÛ³Ý Ñ»é³óáõÙ
    optype = "22"
    opdate = "21/03/13"
    group = False
    fdDoc = False
    Call DeleteOP(optype, opdate, group, fdDoc)
    
    'ïáÏáëÝ»ñÇ Ñ³ßí³ñÏ ÷³ëï³ÃÕÃÇ Ñ»é³óáõÙ
    optype = "511"
    opdate = "20/03/13"
    group = False
    Call DeleteOP(optype, opdate, group, fdDoc)
    
    'ì³ñÏÇ Ù³ñáõÙ ·áñÍáÕáõÃÛ³Ý Ñ»é³óáõÙ
    optype = "22"
    opdate = "25/02/13"
    group = False
    fdDoc = False
    Call DeleteOP(optype, opdate, group, fdDoc)
    
    '¶³ÝÓÙ³Ý Ïáõï³ÏáõÙ ÷³ëï³ÃÕÃÇ Ñ»é³óáõÙ
    optype = "H1"
    opdate = "25/02/13"
    group = True
    Call DeleteOP(optype, opdate, group, fdDoc)
    
     'ïáÏáëÝ»ñÇ Ñ³ßí³ñÏ ÷³ëï³ÃÕÃÇ Ñ»é³óáõÙ
    optype = "511"
    opdate = "24/02/13"
    group = False
    Call DeleteOP(optype, opdate, group, fdDoc)
    
    '¶³ÝÓÙ³Ý Ïáõï³ÏáõÙ ÷³ëï³ÃÕÃÇ Ñ»é³óáõÙ
    optype = "H1"
    opdate = "25/01/13"
    group = True
    Call DeleteOP(optype, opdate, group, fdDoc)
    
    'ïáÏáëÝ»ñÇ Ñ³ßí³ñÏ ÷³ëï³ÃÕÃÇ Ñ»é³óáõÙ
    optype = "511"
    opdate = "24/01/13"
    group = False
    Call DeleteOP(optype, opdate, group, fdDoc)
    
    'ØÝ³ó³Í ·áñÍáÕáõÃÛáõÝÝ»ñÇ Ñ»é³óáõÙ
    actionCount = Delete_Operations_From_OperationsView_Folder(9)
    If Not actionCount Then
        Log.Error("Wrong count of actions")
    End If
    
    'ì³ñÏ³ÛÇÝ å³ÛÙ³Ý³·ñÇ çÝçáõÙ
    Call Online_PaySys_Delete_Agr()
    
    'Test CleanUp
    Call Close_AsBank()      
End Sub