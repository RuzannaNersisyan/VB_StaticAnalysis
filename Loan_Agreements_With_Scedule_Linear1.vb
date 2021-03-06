'USEUNIT Library_Common
'USEUNIT Payment_Order_ConfirmPhases_Library
'USEUNIT Payment_Except_Library
'USEUNIT Online_PaySys_Library
'USEUNIT Loan_Agreements_Library
'USEUNIT Loan_Agreemnts_With_Schedule_Library
'USEUNIT Contract_Summary_Report_Library
'USEUNIT Library_CheckDB
'USEUNIT Subsystems_SQL_Library
'USEUNIT Loan_Agreements_With_Schedule_Linear_Library
'USEUNIT Constants

'Test case ID  165063

Sub Credit_With_Schedule_WithLimit_Test()
    
    Dim fDATE, data, startDATE , calcPRBase1, fadeBase, calcPRBase, queryString, giveCrBase, fBaseCP, fDate1, isExists, docNumber, fISN, actionCount, dateStart, dateEnd
    Dim clientCode, tmpltype, curr, accacc, summ, date_arg, dateFillType, fadeDate, finishFadeDate
    Dim passDirection, sumDates, sumFill, pCalcDate, agrIntRate, agrIntRatePart, branch, sector, schedule
    Dim guarante, startFadeDate, district, note, paperCode, fBASE, docExist, isEqual, round, percent
    Dim dategive, dateconcl, newSchedule, date_perm , allWithLimit , Count, mainSumma, groupOpIsn, fadePeriod, sum
    
    Utilities.ShortDateFormat = "yyyymmdd"
    startDATE = "20030101"          
    fDATE = "20250101"
    clientCode = "00034851"
    curr = "000"
    accacc = "30220042300"
    summ = "100,000.00"
    dateconcl = "05/12/12"
    data = "05/12/12"
    dategive = "05/12/12"
    date_arg = "05/12/13"
    dateFillType = "9"
    fadeDate = Null
    startFadeDate = "05/12/12"
    finishFadeDate = "05/12/13"
    passDirection = "2"
    sumDates = Null
    sumFill = Null
    round = "2"
    agrIntRate = "19"
    agrIntRatePart = "365"
    branch = "9"
    sector = "U2"
    schedule = "9"
    guarante = "9"
    district = "001"
    paperCode = "12"
    percent = "10,038.70"
    pCalcDate = "05/12/12"
    newSchedule = True
    date_perm = "05/12/13"
    allWithLimit = True
    restore = True
    period = "3"
    mainSumma = "26000"
    fadePeriod = Null
    aim = "00"
    
    'Test StartUp 
    Call Initialize_AsBank("bank", startDATE, fDATE)
    Call Login("CREDITOPERATOR")
    Call Create_Connection()
    
    '¶ñ³ýÇÏáí í³ñÏ³ÛÇÝ å³ÛÙ³Ý³·ñÇ ëï»ÕÍáõÙ
    Call Select_Credit_Type("¶ñ³ýÇÏáí í³ñÏ³ÛÇÝ å³ÛÙ³Ý³·Çñ (·Í³ÛÇÝ)")
    Call Credit_With_Schedule_Linear_Doc_Fill(clientCode, tmpl_type, curr, accacc, summ, dateconcl, dategive, date_arg , _
                                              allWithLimit , date_perm , restore, dateFillType, fadeDate, fadePeriod, _
                                              finishFadeDate, startFadeDate, passDirection, sumDates, sumFill, round, agrIntRate, _
                                              agrIntRatePart, pcnotchoose , pcGrant , pcPenAgr, pcPenPer , part, _
                                              branch, sector,aim, schedule, guarante, district, note, paperCode, fBASE, docNumber)
    
    'ê³ÑÙ³Ý³ã³÷Ç ·ñ³ýÇÏÇ Ýß³Ý³ÏáõÙ
    docExist = Limit_Schedule(period, fdsumm)
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
    
    'Ø³ñÙ³Ý ·ñ³ýÇÏÇ Ýß³Ý³ÏáõÙ
    docExist = Fade_Schedule()
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
    
    '¶³ÝÓáõÙ ïñ³Ù³¹ñáõÙÇó ·áñÍáÕáõÃÛ³Ý Ï³ï³ñáõÙ
    Call Collect_From_Provision(data, sum, "2", accacc, fBaseCP)
    
    'ì³ñÏÇ ïñ³Ù³¹ñáõÙ ·áñÍáÕáõÃÛ³Ý Ï³ï³ñáõÙ
    Call Give_Credit(data, summ, "2", accacc, giveCrBase)
    
    'îáÏáëÝ»ñÇ Ñ³ßí³ñÏáõÙ ·áñÍáÕáõÃÛ³Ý Ï³ï³ñáõÙ
    calcPRBase = Calculate_Percents(pCalcDate, pCalcDate, False)
    
    queryString = "select COUNT(*) from DOCS where fSTATE=7 and fNEXTTRANS=2  and fISN= '" & fBASE & "'"
    sql_Value = 1
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    queryString = "select COUNT(*) from AGRSCHEDULE where fAGRISN= '" & fBASE & "'"
    sql_Value = 3
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    queryString = "select COUNT(*) from .AGRSCHEDULEVALUES where fVALUETYPE=6 and fAGRISN='" & fBASE & "'"
    sql_Value = 4
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    queryString = "select COUNT(*) from .AGRSCHEDULEVALUES where fVALUETYPE<>6 and fAGRISN='" & fBASE & "'"
    sql_Value = 10
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    queryString = "select COUNT(*) from  dbo.HIF where fDATE='2012-12-05' and fOBJECT='" & fBASE & "'"
    sql_Value = 16
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    queryString = "select COUNT(*) from  dbo.HIF where fDATE='2012-12-04' and fOBJECT='" & fBASE & "'"
    sql_Value = 6
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    queryString = "select COUNT(*) from  dbo.HIR where fOBJECT='" & fBASE & "'"
    Log.Message(fBASE)
    sql_Value = 3
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    queryString =  "select SUM(fLASTREM) from  dbo.HIRREST where fTYPE = 'R1' and fOBJECT= '" & fBASE & "'"
    sql_Value = 100000.00
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If  

    queryString =  "select SUM(fLASTREM) from  dbo.HIRREST where fTYPE = 'R2' and fOBJECT= '" & fBASE & "'"
    sql_Value = 52.10
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If

    queryString =  "select SUM(fLASTREM) from  dbo.HIRREST where fOBJECT= '" & fBASE & "'"
    sql_Value = 100152.10
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If

    queryString = "select COUNT(*) from  dbo.HIT where  fDATE='2012-12-05' and fTYPE='N2' and fOP='PER' and fOBJECT='" & fBASE & "'"
    sql_Value = 1
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    'îáÏáëÝ»ñÇ Ñ³ßí³ñÏáõÙ ·áñÍáÕáõÃÛ³Ý Ï³ï³ñáõÙ
    pCalcDate = "04/03/13"
    calcPRBase = Calculate_Percents(pCalcDate, pCalcDate, False)
    
    '¸ÇïáõÙ ¨ ËÙµ³·ñáõÙ Ù»ÝÛáõÇó ê³ÑÙ³Ý³ã³÷Ç ïáÕ»ñÇ ù³Ý³ÏÇ ïáõ·áõÙ
    Count = 1
    isEqual = Check_Limit_Count(Count)
    If Not isEqual Then
        Log.Error("Wrong count of limits")
    End If
    
    'ê³ÑÙ³Ý³ã³÷ ëÛ³Ý ëïáõ·áõÙ
    If Trim(wMDIClient.VBObject("frmPttel_2").VBObject("tdbgView").Columns.Item(3))<>"100,000.00" Then
        Log.Error("Wrong value of limit")
    End If
    
    BuiltIn.Delay(1000)
    wMDIClient.vbObject("frmPttel_2").Close()
    
    'ì³ñÏÇ å³ñïù»ñÇ Ù³ñáõÙ ·áñÍáÕáõÃÛ³Ý Ï³ï³ñáõÙ
    fDate1 = "05/03/13"
    Call Fade_Debt(fDate1, fadeBase, Null, mainSumma, null, False)
    
    'ê³ÑÙ³Ý³ã³÷»ñÇ ÷á÷áËáõÙ Áëï ·ñ³ýÇÏÝ»ñÇ ·áñÍáÕáõÃÛ³Ý Ï³ï³ñáõÙ
    calcDate = "05/03/13"
    Call Percent_Group_Calculate(calcDate, calcDate, False, True)
    Count = 2
    isEqual = Check_Limit_Count(Count)
    If Not isEqual Then
        Log.Error("Wrong count of limits")
    End If
    
    'ê³ÑÙ³Ý³ã³÷Ç ÷á÷áËáõÙ ÷³ëï³ÃÕÃÇ ISN-Ç ëï³óáõÙ
    wMDIClient.vbObject("frmPttel_2").vbObject("tdbgView").MoveLast
    'ê³ÑÙ³Ý³ã³÷ ëÛ³Ý ëïáõ·áõÙ
    If Trim(wMDIClient.vbObject("frmPttel_2").vbObject("tdbgView").Columns.Item(3))<>"75,000.00" Then    
        Log.Error("Wrong value of limit")
    End If
    
    BuiltIn.Delay(3000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_View)
    BuiltIn.Delay(3000)
    groupOpIsn = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
    Call ClickCmdButton(1, "OK")
    
    BuiltIn.Delay(1000)
    wMDIClient.vbObject("frmPttel_2").Close
    
    'âû·ï. Ù³ëÇ å³Ñáõëï³íáñáõÙ ·áñÍáÕáõÃÛ³Ý Ï³ï³ñáõÙ
    Call FillDoc_Store_UnusedPart(calcDate)
    
    'ä³ÛÙ³Ý³·ñÇ ¹ÇïáõÙ ·áñÍáÕáõÃÛ³Ý Ï³ï³ñáõÙ
    isExists = View_Contract()
    If Not isExists Then
        Log.Error("The document view doesn't exist")
    End If
    
    queryString = "select COUNT(*) from AGRSCHEDULE where fAGRISN=  '" & fBASE & "'"
    sql_Value = 4
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    queryString = "select COUNT(*) from .AGRSCHEDULEVALUES where fAGRISN='" & fBASE & "'and fVALUETYPE=6"
    sql_Value = 4
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    queryString = "select COUNT(*) from .AGRSCHEDULEVALUES where fAGRISN='" & fBASE & "'and fVALUETYPE<>6"
    sql_Value = 16
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    queryString = "select SUM(fSUM) from .AGRSCHEDULEVALUES where fAGRISN=  '" & fBASE & "'and fVALUETYPE<>6"
    sql_Value = 192978.00
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    queryString = "select COUNT(*) from  dbo.CONTRACTS where fDGISN=  '" & fBASE & "'"
    sql_Value = 1
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    queryString = "select COUNT(*) from  dbo.HIF where fOBJECT='" & fBASE & "'and fDATE>'2012-12-05'"
    sql_Value = 2
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
'        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    queryString = "select SUM(fSUM), SUM(fCURSUM) from  dbo.HIF where fOBJECT='" & fBASE & "'and fDATE>'2012-12-05'"
    sql_Value = 75000.00
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    sql_Value = 0
    colNum = 1
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
       
    queryString = "select COUNT(*) from  dbo.HIR where fOBJECT= '" & fBASE & "'"
    sql_Value = 11
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
'        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    queryString = "select COUNT(*) from  dbo.HIT where fOBJECT='" & fBASE & "'and fTYPE='N2' and fOP='PER'"
    sql_Value = 2
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    queryString = "select COUNT(*) from  dbo.HIR where fOBJECT= '" & fBASE & "' and fOP='RES'"
    Log.Message(fBASE)
    sql_Value = 1
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    queryString = "select COUNT(*) from HI where fDATE='2012-12-05' and fBASE='" & giveCrBase & "'and fTYPE='01'  and fACR=82335686"
    sql_Value = 2
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    queryString = "select COUNT(*) from HI where fDATE='2012-12-05' and  fBASE='" & giveCrBase & "'and fTYPE='02' and fSUM=100000.00 and fCURSUM=100000.00"
    sql_Value = 1
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    queryString = "select COUNT(*) from HI where fDATE='2013-03-05' and  fBASE='" & fadeBase & "'and fTYPE='01' and fADB=82335686  and fSUM=26000.00 "
    sql_Value = 2
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    queryString = "select COUNT(*) from HI where fDATE='2013-03-05' and  fBASE='" & fadeBase & "'and fTYPE='01'  and fSUM=4684.90"
    sql_Value = 2
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    queryString = "select COUNT(*) from HI where fDATE='2013-03-05' and  fBASE='" & fadeBase & "'and fTYPE='02'  and fSUM=26000.00"
    sql_Value = 1
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    queryString = "select COUNT(*) from HI where fDATE='2013-03-05' and  fBASE='" & groupOpIsn & "'and fTYPE='02' and fSUM=25000.00 and fDBCR='C'"
    sql_Value = 2
    colNum = 0
    sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
    If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
    End If
    
    'Test CleanUp start
    'âû·ï. Ù³ëÇ å³Ñáõëï³íáñáõÙ ·áñÍáÕáõÃÛ³Ý Ñ»é³óáõÙ
   ' optype = "Q1"
   ' opdate = Null
    'group = False
    'fdDoc = False
    'Call DeleteOP(optype, opdate, group, fdDoc)
    
'    'âû·ï. Ù³ëÇ å³Ñáõëï³íáñáõÙ ·áñÍáÕáõÃÛ³Ý Ñ»é³óáõÙ
'    optype = "Q1"
'    opdate = Null
'    group = False
'    fdDoc = False
'    Call DeleteOP(optype, opdate, group, fdDoc)
'    
'    
    'ê³ÑÙ³Ý³ã³÷Ç Ñ»é³óáõÙ
   ' isEqual = Check_Limit_Count(Count)
   ' If Not isEqual Then
   '     Log.Error("Wrong count of limits")
   '     Exit Sub
        
  '  End If
  '  Sys.Process("Asbank").vbObject("MainForm").Window("MDIClient", "", 1).vbObject("frmPttel_2").vbObject("tdbgView").MoveLast
  '  BuiltIn.Delay(delay_middle)
  '  Call wMainForm.MainMenu.Click("¶áñÍáÕáõÃÛáõÝÝ»ñ|´áÉáñ ·áñÍáÕáõÃÛáõÝÝ»ñÁ . . .")
  '  BuiltIn.Delay(delay_middle)
   ' Call wMainForm.PopupMenu.Click("æÝç»É")
  '  BuiltIn.Delay(delay_middle)
   ' Sys.Process("Asbank").vbObject("frmAsMsgBox").vbObject("cmdButton").Click()
  '  Call Sys.Process("Asbank").vbObject("frmDeleteDoc").vbObject("YesButton").Click()
   ' Sys.Process("Asbank").vbObject("MainForm").Window("MDIClient", "", 1).vbObject("frmPttel_2").Close()
    
    '¶áñÍáÕáõÃÛáõÝÝ»ñÇ Ñ»é³óáõÙ ¶áñÍáÕáõÃÛáõÝÝ»ñÇ ¹ÇïáõÙÇó
'    actionCount = Delete_Operations_From_OperationsView_Folder(6)
'    If Not actionCount Then
'        Log.Error("Wrong count of actions")
'    End If
'    
'    'ì³ñÏ³ÛÇÝ å³ÛÙ³Ý³·ñÇ çÝçáõÙ
'    Call Online_PaySys_Delete_Agr()
'    Call Close_AsBank()
    'Test CleanUp end
    
End Sub