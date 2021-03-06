Option Explicit
'USEUNIT Payment_Order_ConfirmPhases_Library
'USEUNIT Derivative_Tools_Library
'USEUNIT Subsystems_SQL_Library
'USEUNIT Derivatives_Library
'USEUNIT Akreditiv_Library
'USEUNIT Library_Common
'USEUNIT Constants
'USEUNIT BankMail_Library

'Test Case -  102921

Sub Persentage_Swap_Test()
  
    Dim startDATE,fDATE ,CurrSwap,FolderPath,fBASE,actionExist,contr
    Dim date,revDate,sumRevl,repDate,extention,actionEndDate,docType
    Dim per,part, sPer,Calculate_Date,dateStart, summperc,calcDate,CloseDate
    Dim repType, time, sold, place,actionDate,actionExists,actionType
    Dim queryString,sql_Value,colNum,sql_isEqual,docAcc,docISN,fOBJECT
    startDATE = "20120101"
    fDATE = "20250101"    
    
    'Test StartUp start
    Call Initialize_AsBank("bank", startDATE, fDATE)
    
    Call Create_Connection()
    
    Call ChangeWorkspace(c_Derivatives)
    'Ստեղծել Ածանցյալ գործիք/ Արժույթային գործիկ տեսակի փաստաթուղթ
    Set CurrSwap = New_DerivativeDoc()  
    With CurrSwap
    .Client = "00000668"
    .BuyAcc = "10310070100"
    .RepayAcc = "77786271031"
    .Date = "140316"
    .Term = "140317 "
    .BaseSum = 1000000
    .PaperCode = 123
    .FirstDate = "030216"
    .Paragraph = 1
    .Direction = 2
  
    Call .CreateDerivative("|²Í³ÝóÛ³É ·áñÍÇùÝ»ñ|Üáñ å³ÛÙ³Ý³·ñÇ ëï»ÕÍáõÙ", "îáÏáë³¹ñáõÛù³ÛÇÝ ëíá÷") 
    WMDIClient.VBObject("frmPttel").Close
    BuiltIn.Delay(8000)
    
    'Կատարում ենք SQL ստուգում
        queryString = "select fSTATE from DOCS where fISN = '" & CurrSwap.fBASE & "'"
        sql_Value = 1
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
      
    'Պայմանագիրը ուղարկել հաստատման
    .SendToVerify("|²Í³ÝóÛ³É ·áñÍÇùÝ»ñ|²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ")
     BuiltIn.Delay(10000)

     Log.Message("Հաստատել պայմանագիրը")
     Call wTreeView.DblClickItem("|²Í³ÝóÛ³É ·áñÍÇùÝ»ñ|Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I")
     BuiltIn.Delay(1000)
     Call ClickCmdButton(2, "Î³ï³ñ»É")
     BuiltIn.Delay(2000)
     
     If Not ConfirmContractDoc(2, CurrSwap.DocNum, c_ToConfirm, 1, "Ð³ëï³ï»É") Then
            Log.Error("փաստաթուղթը չի վավերացվել")
            Exit Sub
      End If
      BuiltIn.Delay(2000)
      wMDIClient.VBObject("frmPttel").Close
      
      BuiltIn.Delay(2000)
      FolderPath = "|²Í³ÝóÛ³É ·áñÍÇùÝ»ñ|ä³ÛÙ³Ý³·ñ»ñ"
      .OpenInFolder(FolderPath)
      End with
      Log.Message( CurrSwap.fBASE)
    
        'Կատարում ենք SQL ստուգում
        queryString = "select fSTATE from DOCS where fISN = '" & CurrSwap.fBASE & "'"
        sql_Value = 7
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
    
        queryString = "select fSUM from HIF where fBASE = '" & CurrSwap.fBASE & "' and fOP = 'PAG'"
        sql_Value = 12.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If

        queryString = "select fSUM from HIF where fBASE = '" & CurrSwap.fBASE & "' and fOP = 'PCR'"
        sql_Value = 5.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If

        queryString = "select fCURSUM from HIF where fBASE = '" & CurrSwap.fBASE & "' and fOP = 'PAG'"
        sql_Value = 365.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
    
        queryString = "select fCURSUM from HIF where fBASE = '" & CurrSwap.fBASE & "' and fOP = 'PCR'"
        sql_Value = 365.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If 
      
    '" Տոկոսի հաշվարկ" գործողության կատարում
    Calculate_Date = "140316"
    Call Calculate_Percent(fBASE, Calculate_Date , Calculate_Date)
    
    BuiltIn.Delay(2000)
        'Կատարում ենք SQL ստուգում
        queryString = "select fCURSUM from HIR where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'R1'"
        sql_Value = 1000000.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fCURSUM from HIR where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'R2'"
        sql_Value = 328.80
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fCURSUM from HIR where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'Rì' and fOP = 'PER'"
        sql_Value = 328.80
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fCURSUM from HIR where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'Rì' and fOP = 'PCR'"
        sql_Value = 137.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fCURSUM from HIR where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'R¸'"
        sql_Value = 328.80
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fCURSUM from HIR where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'RÆ'"
        sql_Value = 137.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'R1'"
        sql_Value = 1000000.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'R2'"
        sql_Value = 328.80
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'R¸'"
        sql_Value = 328.80
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'RÆ'"
        sql_Value = 137.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'Rì'"
        sql_Value = 191.80
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'Rî'"
        sql_Value = 137.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
    
    ' "Տոկոսադրույք" գործողության կատարում
    revDate = "150316"
    per = "14"
    part = "365"
    sPer = "8"
    Call Set_Persentage(fBase,revDate,per,part, sPer,part)
    
        'Կատարում ենք SQL ստուգում
        queryString = "select fSUM from HIF where fBASE= '" & fBase & "' and fOP = 'PAG'"
        sql_Value = 14.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
    
        queryString = "select fSUM from HIF where fBASE = '" & fBase & "' and fOP = 'PCR'"
        sql_Value = 8.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
    'Խմբային տոկոսի հաշվարկ
    calcDate = "151216"
    Call Group_Percent_Calculate_Overdraft(calcDate , calcDate)
    
    BuiltIn.Delay(2000)
        'Կատարում ենք SQL ստուգում
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'R2'"
        sql_Value = 106192.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'R¸'"
        sql_Value = 105808.40
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'RÆ'"
        sql_Value = 60411.10
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'Rì'"
        sql_Value = 45561.70
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'Rî'"
        sql_Value = 60630.30
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'R2'"
        sql_Value = 328.80
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'R¸'"
        sql_Value = 94301.50
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select SUM(fCURSUM) from HIR where fOBJECT = '" & CurrSwap.fBASE & "'and  fTYPE = 'Rî'"
        sql_Value = 60630.30
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select SUM(fCURSUM) from HIR where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'R¸'"
        sql_Value = 105808.4
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select SUM(fCURSUM) from HIR where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'RÆ'"
        sql_Value = 60411.1
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select SUM(fCURSUM) from HIR where fOBJECT = '" & CurrSwap.fBASE & "' and  fTYPE = 'Rì'"
        sql_Value = 166822.3
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select SUM(fCURSUM) from HIR where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'R2'"
        sql_Value = 106192
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'RÆ'"
        sql_Value = 53835.70
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'Rì'"
        sql_Value = 191.80
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'Rî'"
        sql_Value = 137.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select SUM(fCURSUM) from HIR where fOBJECT = '" & CurrSwap.fBASE & "'and  fOP = 'PRJ' and fDBCR = 'D'"
        sql_Value = 166219.50
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select SUM(fCURSUM) from HIR where fOBJECT = '" & CurrSwap.fBASE & "' and fOP = 'PER' and fDBCR = 'D'"
        sql_Value = 273014.30
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select SUM(fCURSUM) from HIR where fOBJECT = '" & CurrSwap.fBASE & "' and fOP = 'PCR' and fDBCR = 'C'"
        sql_Value = 60630.30
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
    ' " Մարում" գործողության կատարում
    date = "161216"
    repType = "1"
    time = "1"
    sold = "1"
    place = "1"
    docType = 1
    contr = "îáÏáë³¹ñáõÛù³ÛÇÝ ëíá÷"
    Call Repayments(date, repType, time, sold, place,docType,contr )
    
    BuiltIn.Delay(2000)
        'Կատարում ենք SQL ստուգում
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'R2'"
        sql_Value = 0.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'R¸'"
        sql_Value = 0.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'RÆ'"
        sql_Value = 0.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'Rì'"
        sql_Value = 0.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'Rî'"
        sql_Value = 0.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'R2'"
        sql_Value = 106192.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'R¸'"
        sql_Value = 105808.40
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'RÆ'"
        sql_Value = 60411.10
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'Rì'"
        sql_Value = 45561.70
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'Rî'"
        sql_Value = 60630.30
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
    '" Ժամկետների վերանայում" գործողության կատարում
    repDate = "161217"
    extention = "1"
    Call ReviewTerms(date,repDate,extention)
    
    '" Տոկոսի հաշվարկ" գործողության կատարում
    Calculate_Date = "150117"
    Call Calculate_Percent(fBASE, Calculate_Date , Calculate_Date)
    
    BuiltIn.Delay(2000)
        'Կատարում ենք SQL ստուգում
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'R2'"
        sql_Value = 11890.40
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'R¸'"
        sql_Value = 11890.40
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'RÆ'"
        sql_Value = 6794.50
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'Rì'"
        sql_Value = 5095.90
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'Rî'"
        sql_Value = 6794.50
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select SUM(fCURSUM) from HIR where fOBJECT = '" & CurrSwap.fBASE & "'and  fOP = 'PRJ' and fDBCR = 'D'"
        sql_Value = 184904.40
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select SUM(fCURSUM) from HIR where fOBJECT = '" & CurrSwap.fBASE & "' and fOP = 'PER' and fDBCR = 'D'"
        sql_Value = 303589.60
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select SUM(fCURSUM) from HIR where fOBJECT = '" & CurrSwap.fBASE & "' and fOP = 'PCR' and fDBCR = 'C'"
        sql_Value = 67424.80
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If      
        
    '" Տոկոսի հաշվարկ" գործողության կատարում
    Calculate_Date = "160117"
    Call Calculate_Percent(fBASE, Calculate_Date , Calculate_Date)  
    
     BuiltIn.Delay(2000)
        'Կատարում ենք SQL ստուգում
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'R2'"
        sql_Value = 12274.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'R¸'"
        sql_Value = 11890.40
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'RÆ'"
        sql_Value = 6794.50
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'Rì'"
        sql_Value = 5260.30
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'Rî'"
        sql_Value = 7013.70
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'R2'"
        sql_Value = 11890.40
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'R¸'"
        sql_Value = 0.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'RÆ'"
        sql_Value = 0.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'Rì'"
        sql_Value = 5095.90
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'Rî'"
        sql_Value = 6794.50
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select SUM(fCURSUM) from HIR where fOBJECT = '" & CurrSwap.fBASE & "' and fOP = 'PER' and fDBCR = 'D'"
        sql_Value = 304576.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select SUM(fCURSUM) from HIR where fOBJECT = '" & CurrSwap.fBASE & "' and fOP = 'PCR' and fDBCR = 'C'"
        sql_Value = 67644.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If 
     
        queryString = "select fPENULTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'Rî'"
        sql_Value = 6794.50
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select SUM(fCURSUM) from HIR where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'Rî'"
        sql_Value = 128274.30
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select SUM(fCURSUM) from HIR where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'R¸'"
        sql_Value = 223507.20
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select SUM(fCURSUM) from HIR where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'RÆ'"
        sql_Value = 127616.70
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select SUM(fCURSUM) from HIR where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'Rì'"
        sql_Value = 231671.70
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select SUM(fCURSUM) from HIR where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'R2'"
        sql_Value = 224658.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select SUM(fCURSUM) from HIR where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'RÆ'"
        sql_Value = 127616.7
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select SUM(fCURSUM) from HIR where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'Rì'"
        sql_Value = 231671.7
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select SUM(fCURSUM) from HIR where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'R2'"
        sql_Value = 224658
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
    ' " Մարում" գործողության կատարում
    date = "170117"
    repType = "2"
    time = "1"
    sold = "1"
    place = "1"
    contr = "îáÏáë³¹ñáõÛù³ÛÇÝ ëíá÷"
    Call Repayments(date, repType, time, sold, place,docType,contr )
    
    BuiltIn.Delay(2000)
        'Կատարում ենք SQL ստուգում
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'R2'"
        sql_Value = 0.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'R¸'"
        sql_Value = 0.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'RÆ'"
        sql_Value = 0.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'Rì'"
        sql_Value = 0.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'Rî'"
        sql_Value = 0.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'R2'"
        sql_Value = 12274.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'R¸'"
        sql_Value = 11890.40
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'RÆ'"
        sql_Value = 6794.50
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'Rì'"
        sql_Value = 5260.30
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM from HIRREST where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'Rî'"
        sql_Value = 7013.70
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select SUM(fCURSUM) from HIR where fOBJECT = '" & CurrSwap.fBASE & "'and  fTYPE = 'Rî'"
        sql_Value = 135288.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select SUM(fCURSUM) from HIR where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'R¸' "
        sql_Value = 235397.60
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select SUM(fCURSUM) from HIR where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'RÆ'"
        sql_Value = 134411.20
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select SUM(fCURSUM) from HIR where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'Rì'"
        sql_Value = 236932.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select SUM(fCURSUM) from HIR where fOBJECT = '" & CurrSwap.fBASE & "' and fTYPE = 'R2'"
        sql_Value = 236932.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
    
    actionDate = "161216"
    actionEndDate = "170117"
    actionExists = True
    actionType = Null
    
    'Ջնջում է Գործողությունների դիտում թղթապանակի բոլոր փաստաթղթերը
    Call Delete_Actions(actionDate,actionEndDate,actionExists,actionType,c_OpersView)

    'Ջնջում է Դիտում և խմբագրում/Ժամկետներ/Պայմ.մարման ժամկետներ թղթապանակի բոլոր փաստաթղթերը
    actionExist = False 
    Call Delete_Actions(actionDate,actionDate,actionExist,actionType,c_ViewEdit & "|" & c_Dates & "|" & c_PerDates )

    'Ջնջում է Գործողությունների դիտում թղթապանակի բոլոր փաստաթղթերը
    actionDate = "150316"
    actionEndDate = "151216"
    Call Delete_Actions(actionDate,actionEndDate,actionExists,actionType,c_OpersView)

    'Ջնջում է Դիտում և խմբագրում/Տոկոսադրույքներ/Տոկոսադրույքներ թղթապանակի բոլոր փաստաթղթերը
    Call Delete_Actions(revDate,revDate,actionExist,actionType,c_ViewEdit & "|" & c_Percentages & "|" & c_Percentages)

    'Ջնջում է գլխավոր պայմանագիրը
    Call Delete_Doc()
    wMDIClient.VBObject("frmPttel").Close()    

    Call Close_AsBank()
     
End Sub