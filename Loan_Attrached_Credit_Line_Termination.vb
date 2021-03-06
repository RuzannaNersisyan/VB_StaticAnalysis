Option Explicit
'USEUNIT Deposit_Contract_Library
'USEUNIT Loan_Agreements_Library 
'USEUNIT Subsystems_SQL_Library
'USEUNIT Credit_Line_Library
'USEUNIT Akreditiv_Library
'USEUNIT Library_Common  
'USEUNIT Constants

'Test Case Id 166691

Sub Loan_Attrached_Credit_Line_Termination_Test()

    Dim fDATE, sDATE, Count, i, DocNum, DocLevel, FolderName,colNum,sql_isEqual
    Dim Date, isExists, fBASE, Sum, calcDate, Param, dategive, CLSate, tabN
    Dim CreditLine, CashOrNo, acc, Climit,isEqual,actionExists,actionType
    Dim attr, mainSum,perSum, capData , summa,closeDate,queryString,sql_Value
      
    ''Համակարգ մուտք գործել ARMSOFT օգտագործողով
    fDATE = "20250101"
    sDATE = "20010101"
    Call Initialize_AsBank("bank", sDATE, fDATE)
    Call Create_Connection()
    Login("ARMSOFT")
    Call ChangeWorkspace(c_LoanAttrached)
  
    Call Log.Message("Ներգրավված վարկեր/Վարկային գիծ",,,attr)
    Set CreditLine = New_LoanDocument()
    With CreditLine
        .CalcAcc = "00000113032"                                    
        .Limit = 2500000
        .Date = "130117" 
        .GiveDate = "130117"
        .Term = "130118"
        .FirstDate = "130117"
        .PaperCode = 555
    
        .DocType = "ì³ñÏ³ÛÇÝ ·ÇÍ"
        Call .CreateAttrLoan(FolderName & "Üáñ å³ÛÙ³Ý³·ñÇ ëï»ÕÍáõÙ")
  
        Log.Message(.DocNum)

        wMDIClient.VBObject("frmPttel").Close
  
        'Պայմանագրին ուղղարկել հաստատման
        .SendToVerify("|Ü»ñ·ñ³íí³Í í³ñÏ»ñ|²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ")
        'Հաստատել
        Call wTreeView.DblClickItem("|Ü»ñ·ñ³íí³Í í³ñÏ»ñ|Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I")
         'Call ClickCmdButton(2,"Î³ï³ñ»É")
        'Կատարում է ստուգում , եթե գլխավոր պայմանագիրը առկա է ,ապա ուղարկում է հաստատման, հակառակ դեպքում դուրս է բերում սխալ
        isExists = Find_Doc_ByNum(CreditLine.DocNum,2)
        If Not isExists Then
            Log.Error("The document does՚t exist")
            Exit sub
        End If
        'Վավերացնում է փաստաթուղթը
        Call Validate_Doc()
        Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").Close
        Log.Message(CreditLine.fBASE)
    
        Call ChangeWorkspace(c_LoanAttrached)
        Call LetterOfCredit_Filter_Fill(FolderName, .DocLevel, .DocNum)
    
    End With
        
    BuiltIn.Delay(1000)
        'Կատարում ենք SQL ստուգում
        queryString = "select fSTATE from DOCS where fISN = '" & CreditLine.fBASE & "'"
        sql_Value = 7
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select count(*) from DAGRACCS where fAGRISN = '" & CreditLine.fBASE & "' "
        sql_Value = 1
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fSUM from HI where fBASE = '" & CreditLine.fBASE & "'"
        sql_Value = 2500000.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select count(*) from HIF where fBASE = '" & CreditLine.fBASE & "'"
        sql_Value =26
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
    'Գանձում ներգավվումից գործողության կատարում 
    Date = "130117"
    Sum = "25000"
    Call ChargeForAttraction(fBASE,Date, Sum, CashOrNo,acc)
    'Վարկի ներգրավվում գործողության կատարում 
    Call Loan_Attraction(fBASE,Date,Sum,CashOrNo,acc)
    BuiltIn.Delay(6000)
    
        'Կատարում ենք SQL ստուգում
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R1'"
        sql_Value = 25000.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R^'"
        sql_Value = 25000.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fCURSUM from HIR where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R1'"
        sql_Value = 25000.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fCURSUM from HIR where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R^' and fOP = 'PAY'"
        sql_Value = 22500.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fCURSUM from HIR where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R^' and fOP = 'TAX' "
        sql_Value = 2500.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
    'Տոկոսի հաշվարկ
    calcDate = "120217" 
    Call Calculate_Percent(fBASE , calcDate , calcDate)
    Log.Message(fBASE)
    BuiltIn.Delay(2000)
    
        'Կատարում ենք SQL ստուգում
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R2'"
        sql_Value = 246.60
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'RH'"
        sql_Value = 16274.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R¸'"
        sql_Value = 246.60
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'RÂ'"
        sql_Value = 16274.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fCURSUM from HIR where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R2'"
        sql_Value = 246.60
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fCURSUM from HIR where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'RH'"
        sql_Value = 16274.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fCURSUM from HIR where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R¸'"
        sql_Value = 246.60
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fCURSUM from HIR where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'RÂ'"
        sql_Value = 16274.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
    'Վարկային գծի դադարցում/Վերականգնում գործողության կատարում
    CLSate = "1"
    dategive = "130217"
    Param  = "|Գծայնության դադարեցում"
    Call Credit_Termination_Restoration(Param,dategive,CLSate)
    
    'Ստուգում է ,արդյոք սահմանաչափը փոխվել է
    Climit = "0.00"
    isEqual = Check_Changed_Limit(dategive,dategive,CLimit)    
    If isEqual Then
      wMDIClient.VBObject("frmPttel_2").Close
    Else
      Log.Error("The limit doesn't change")
      wMDIClient.VBObject("frmPttel_2").Close 
    End If
    
    'Պարտքերի մարում
    tabN = 2
    Call Debt_Repayment(fBASE,dategive, mainSum,perSum,cashORno,Acc,docNum, tabN)
    BuiltIn.Delay(2000)
    
        queryString = "select fCURSUM from HIR where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R2' and fOP = 'DBT'"
        sql_Value = 221.90
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fCURSUM from HIR where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R2' and fOP = 'TXD'"
        sql_Value = 24.70
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fCURSUM from HIR where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'RH' and fOP = 'DBT'"
        sql_Value = 16274.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fCURSUM from HIR where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R¸'  and fOP = 'DBT'"
        sql_Value = 221.90
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fCURSUM from HIR where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R¸' and fOP = 'TXD'"
        sql_Value = 24.70
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fCURSUM from HIR where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'RÂ'"
        sql_Value = 16274.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R2'"
        sql_Value = 246.60
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'RH'"
        sql_Value = 16274.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
    'Տոկոսի հաշվարկ
    Call Calculate_Percent(fBASE , dategive , dategive)
    Log.Message(fBASE)
    BuiltIn.Delay(2000)
    
         'Կատարում ենք SQL ստուգում
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R2'"
        sql_Value = 8.20
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fCURSUM from HIR where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R2' and fOP = 'PER' and fDATE = '2017-02-13'"
        sql_Value = 8.20
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
    'Տոկոսի հաշվարկ
    calcDate = "120317" 
    Call Calculate_Percent(fBASE , calcDate , calcDate)
    Log.Message(fBASE)
    BuiltIn.Delay(1000)
    
         'Կատարում ենք SQL ստուգում
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R2'"
        sql_Value = 230.10
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fCURSUM from HIR where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R2' and fOP = 'PER' and fDATE = '2017-03-12'"
        sql_Value = 221.90
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R2'"
        sql_Value = 8.20
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
                
    'Վարկային գծի վերականգնում/դադարեցում
    dategive = "130317"
    CLSate = "2"
    Param = "|Գծայնության վերականգնում"
    Call Credit_Termination_Restoration(Param,dategive,CLSate) 
    
    'Ստուգում է ,արդյոք սահմանաչափը փոխվել է
    Climit = "2,500,000.00"
    isEqual = Check_Changed_Limit(dategive,dategive,CLimit)    
    If isEqual Then
      wMDIClient.VBObject("frmPttel_2").Close
    Else
      Log.Error("The limit doesn't change") 
      wMDIClient.VBObject("frmPttel_2").Close
    End If
    
    'Տոկոսի հաշվարկ
    Call Calculate_Percent(fBASE , dategive , dategive)
    Log.Message(fBASE)

    'Տոկոսի հաշվարկ
    calcDate = "120417" 
    Call Calculate_Percent(fBASE , calcDate , calcDate)
    Log.Message(fBASE)
    BuiltIn.Delay(1000)
    
         'Կատարում ենք SQL ստուգում
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R2'"
        sql_Value = 484.90
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'RÂ'"
        sql_Value = 16816.50
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R¸'"
        sql_Value = 484.90
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'RH'"
        sql_Value = 16816.50
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R2'"
        sql_Value = 238.30
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and  fTYPE = 'RH'"
        sql_Value = 542.50
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
    'Տոկոսների կապիտալացում գործողղության կատարում 
    capData = "130417"
    summa = "484.90"
    Call Percent_Capitalization(fBASE , capData , summa)
    
         'Կատարում ենք SQL ստուգում
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R2'"
        sql_Value = 0.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R¸'"
        sql_Value = 0.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R2'"
        sql_Value = 484.90
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
    'Տոկոսի խմբային հաշվարկ գորշողության կատարում 
    closeDate = "140118"
    Call Group_Persent_Calculate(closeDate,closeDate)
    
        'Կատարում ենք SQL ստուգում
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R2'"
        sql_Value = 2299.60
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R¸'"
        sql_Value = 2299.60
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R1'"
        sql_Value = 25436.40
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'RH'"
        sql_Value = 165968.50
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'RÂ'"
        sql_Value = 165968.50
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
                
        queryString = "select fPENULTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'RH'"
        sql_Value = 16816.50
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "'and  fTYPE = 'R¸'"
        sql_Value = 2040.40
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and  fTYPE = 'RÂ'"
        sql_Value = 149155.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If        
        
    'Պարտքերի մարում
    dategive = "150118"
    Call Debt_Repayment(fBASE,dategive, mainSum,perSum,cashORno,Acc,docNum, tabN)
    BuiltIn.Delay(2000)    
    
        'Կատարում ենք SQL ստուգում
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R2'"
        sql_Value = 0.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R¸'"
        sql_Value = 0.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R1'"
        sql_Value = 0.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'RH'"
        sql_Value = 0.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'RÂ'"
        sql_Value = 0.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
                
        queryString = "select fPENULTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'RH'"
        sql_Value = 165968.50
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R¸'"
        sql_Value = 2299.60
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'RÂ'"
        sql_Value = 165968.50
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If    
        
    'Ջնջում է Պարտքերի մարման փաստաթուղթը
    dategive = "150118"
    actionExists = True
    actionType = "23"
    Call Delete_Actions(dategive,dategive,actionExists,actionType,c_OpersView)
    'Ջնջում է Տոկոսի հաշվարկ փաստաթուղթը
    dategive = "140118"
    actionType = "211"
    Call Delete_Actions(dategive,dategive,actionExists,actionType,c_OpersView)
    'Ջնջում է Տոկոսի կապիտալացում փաստաթուղթը
    dategive = "130417"
    actionType = "13"
    Call Delete_Actions(dategive,dategive,actionExists,actionType,c_OpersView)
    'Ջնջում է Տոկոսի հաշվարկ փաստաթուղթը
    dategive = "120417"
    actionType = "211"
    Call Delete_Actions(dategive,dategive,actionExists,actionType,c_OpersView)
    'Ջնջում է Տոկոսի հաշվարկ փաստաթուղթը
    dategive = "130317"
    actionType = "211"
    Call Delete_Actions(dategive,dategive,actionExists,actionType,c_OpersView)
    'Ջնջում է Տոկոսի հաշվարկ փաստաթուղթը   
    dategive = "120317"
    actionType = "211"
    Call Delete_Actions(dategive,dategive,actionExists,actionType,c_OpersView)
    'Ջնջում է Տոկոսի հաշվարկ փաստաթուղթը
    dategive = "130217"
    actionType = "211"
    Call Delete_Actions(dategive,dategive,actionExists,actionType,c_OpersView)
    'Ջնջում է Պարտքերի մարման փաստաթուղթը
    dategive = "130217"
    actionType = "23"
    Call Delete_Actions(dategive,dategive,actionExists,actionType,c_OpersView)
    'Ջնջում է Տոկոսի հաշվարկ փաստաթուղթը
    dategive = "120217"
    actionType = "211"
    Call Delete_Actions(dategive,dategive,actionExists,actionType,c_OpersView)
    'Ջնջում է Սահամաչափերը
    dategive = "130317"
    actionExists = False
    Call Delete_Actions(Date,dategive,actionExists,actionType,c_ViewEdit & "|" & c_Other & "|" & c_Limits)
    'Ջնջում է Վարկի մարում փաստաթուղթը
    dategive = "130117"
    actionExists = True
    actionType = "11"
    Call Delete_Actions(dategive,dategive,actionExists,actionType,c_OpersView)
    'Ջնջում է Գանձման փաստաթուղթը
    actionType = ""
    Call Delete_Actions(dategive,dategive,actionExists,actionType,c_OpersView)
    'Ջնջում է գլխավոր պայմանագիրը
    Call Delete_Doc()
    Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").Close()
    'Փակում է ASBANK - ը
    Call Close_AsBank()   
  
End Sub