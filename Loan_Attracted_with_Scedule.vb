Option Explicit
'USEUNIT Payment_Order_ConfirmPhases_Library
'USEUNIT Group_Operations_Library
'USEUNIT Deposit_Contract_Library
'USEUNIT Loan_Agreements_Library 
'USEUNIT Subsystems_SQL_Library
'USEUNIT Credit_Line_Library
'USEUNIT Akreditiv_Library
'USEUNIT Library_Common  
'USEUNIT Constants

'Test Case Id 166693

Sub Loan_Attracted_with_Scedule_Test()

    Dim fDATE, sDATE, Count, i, DocNum, DocLevel, FolderName,colNum,sql_isEqual
    Dim Date, isExists, fBASE, Sum, calcDate, Param, dategive, CLSate, Workspace
    Dim CreditLine, CashOrNo, acc, Climit,isEqual,actionExists,actionType, DocType
    Dim attr, mainSum,perSum, capData ,docExist, perS, date_arg, summa,closeDate,queryString,sql_Value
    Dim griddate,Period,Direction, state, tabN
    
    ''Համակարգ մուտք գործել ARMSOFT օգտագործողով
    fDATE = "20250101"
    sDATE = "20010101"
    Date = "060417"
    Call Initialize_AsBank("bank", sDATE, fDATE)
    Call Create_Connection()
    Login("ARMSOFT")
    Call ChangeWorkspace(c_LoanAttrached)
  
    Call Log.Message("Ներգրավված վարկեր/Գրաֆիկով վարկային պայմանագիր",,,attr)
    Set CreditLine = New_LoanDocument()
    With CreditLine
        .CalcAcc = "77787753818"                                    
        .Limit = 1550000
        .Date = "060417" 
        .GiveDate = "060417"
        .Term = "060418"
        .FirstDate = "060417"
        .PaperCode = 555    
        .DocType = "¶ñ³ýÇÏáí í³ñÏ³ÛÇÝ å³ÛÙ³Ý³·Çñ"
        Call .CreateAttrLoan(FolderName & "Üáñ å³ÛÙ³Ý³·ñÇ ëï»ÕÍáõÙ")
  
        Log.Message(.DocNum)

        'wMDIClient.VBObject("frmPttel").Close
        'Մարման գրաֆիկի նշանակում
        
        BuiltIn.Delay(2000)
        Call wMainForm.MainMenu.Click(c_AllActions)
        Call wMainForm.PopupMenu.Click(c_RepaySchedule)
        BuiltIn.Delay(2000)
        
        'Այլ վճարումների գրաֆիկի նշանակում
        param = "¶ñ³ýÇÏáí í³ñÏ³ÛÇÝ å³ÛÙ³Ý³·Çñ- "& Trim(.DocNum) & " {äáÕáëÛ³Ý äáÕáë}"
        griddate = "060418"
        period = 1
        direction = 2
        If Not Other_Payment_Schedule_AllTypes(param,Date,Date,griddate,Period,Direction) Then 
            Log.Error("There was no document")
            Exit Sub
        End If
        'wMDIClient.VBObject("frmPttel").Close
        Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").VBObject("tdbgView").Keys("[Up]")
        Call PaySys_Send_To_Verify() 
        BuiltIn.Delay(2000) 
        Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").Close
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
        
        queryString = "select count(*) from HIF where fBASE = '" & CreditLine.fBASE & "'"
        sql_Value = 13
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
    'Գանձում ներգավվումից գործողության կատարում 
    Date = "060417"
    Sum = "1550000"
    Call ChargeForAttraction(fBASE,Date, Sum, CashOrNo,acc)
    'Վարկի ներգրավվում գործողության կատարում 
    Call Loan_Attraction(fBASE,Date,Sum,CashOrNo,acc)
    
        'Կատարում ենք SQL ստուգում
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R^'"
        sql_Value = 1550000.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fCURSUM from HIR where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R1'"
        sql_Value = 1550000.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fCURSUM from HIR where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R^' and fOP = 'PAY'"
        sql_Value = 1395000.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fCURSUM from HIR where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R^' and fOP = 'TAX' "
        sql_Value = 155000.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If 
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R¾' "
        sql_Value = -1550000.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fCURSUM from HIR where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R¾' "
        sql_Value = 1550000.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If

        
    'Տոկոսի հաշվարկ
    calcDate = "070517" 
    Call Calculate_Percent(fBASE , calcDate , calcDate)
    Log.Message(fBASE)
    BuiltIn.Delay(1000)
    
        'Կատարում ենք SQL ստուգում
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R2'"
        sql_Value = 15797.3
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R¸'"
        sql_Value = 15797.3
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fCURSUM from HIR where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R2'"
        sql_Value = 15797.3
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fCURSUM from HIR where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R¸'"
        sql_Value = 15797.3
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
     'Պարտքերի մարում
    dategive = "080517"
    perSum = "20778.10"
    tabN =2
    Call Debt_Repayment(fBASE,dategive, mainSum,perSum,cashORno,Acc,docNum, tabN)
    BuiltIn.Delay(1000)
    
        'Կատարում ենք SQL ստուգում
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R2'"
        sql_Value = -4980.8
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
        
        queryString = "select fCURSUM from HIR where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R2' and fOP = 'TXD'"
        sql_Value = 2077.8
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R2'"
        sql_Value = 15797.3
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R¾' "
        sql_Value = -1565797.30
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R¾' "
        sql_Value = -1550000.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fCURSUM from HIR where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R¾'  and fOP = 'PER'"
        sql_Value = -15797.30
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
    
    'Կանխավ վճարված տոկոսի հաշվարկ
    state = True
    Call Return_Payed_Percent(dategive, "3000","2",acc,fBASE, state)
    Log.Message(fBASE)
    
        'Կատարում ենք SQL ստուգում
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R2'"
        sql_Value = -1980.8
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
    
    'Գումարի մարում տոկոսների հաշվին
    Call Fadeing_LeasingSumma_From_PayedPercents(c_FadeLoanFromPercent, dategive,"SUMMA", "1980.80")
    BuiltIn.Delay(1000)
    
        'Կատարում ենք SQL ստուգում
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R2'"
        sql_Value = 0.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R1'"
        sql_Value = 1418852.5
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R1'"
        sql_Value = 1550000.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
    'Տոկոսի հաշվարկ
    calcDate = "080517" 
    Call Calculate_Percent(fBASE , calcDate , calcDate)
    Log.Message(fBASE)
    BuiltIn.Delay(1000)
    
        'Կատարում ենք SQL ստուգում
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R2'"
        sql_Value = 466.5
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
    
    'Տոկոսի հաշվարկ
    calcDate = "050617" 
    Call Calculate_Percent(fBASE , calcDate , calcDate)
    Log.Message(fBASE)
    BuiltIn.Delay(1000)
        
         'Կատարում ենք SQL ստուգում
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R2'"
        sql_Value = 13527.7
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R¸'"
        sql_Value = 13527.7
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        queryString = "select fPENULTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R2'"
        sql_Value = 466.5
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If

    'Հաշվարկների ճշգրտում
    dategive = "060617"
    perSum = "2000"
    Call Correction_Calculation(dategive, perSum, fBASE)
    Log.Message(fBASE)
    
         'Կատարում ենք SQL ստուգում
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R2'"
        sql_Value = 15527.7
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R¸'"
        sql_Value = 15527.7
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R2'"
        sql_Value = 13527.7
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        

    'Տոկոսի խմբային հաշվարկ գորշողության կատարում 
    closeDate = "050418"
    Call Group_Persent_Calculate(closeDate,closeDate)
    
         'Կատարում ենք SQL ստուգում
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R2'"
        sql_Value = 86870.1
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R¸'"
        sql_Value = 86870.1
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R2'"
        sql_Value = 15527.7
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R¸'"
        sql_Value = 85553.7
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
                
    'Պարտքերի մարում
    dategive = "060418"
    Call Debt_Repayment(fBASE,dategive, mainSum,perS,cashORno,Acc,docNum, tabN)
    Log.Message(fBASE)
    BuiltIn.Delay(5000)
    
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
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R^'"
        sql_Value = 1550000.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R2'"
        sql_Value = 86870.1
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R¾' "
        sql_Value = 0.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM from HIRREST where fOBJECT = '" & CreditLine.fBASE & "' and fTYPE = 'R¾' "
        sql_Value = -1003977.70
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
    'Ջնջում է Պարտքերի մարման փաստաթուղթը
    dategive = "060418"
    actionExists = True
    actionType = "12"
    Call Delete_Actions(dategive,dategive,actionExists,actionType,c_OpersView)
    'Ջնջում է Տոկոսի հաշվարկ փաստաթուղթը
    dategive = "050418"
    actionType = "211"
    Call Delete_Actions(dategive,dategive,actionExists,actionType,c_OpersView)
    'Ջնջում է Տոկոսի կապիտալացում փաստաթուղթը
    dategive = "060617"
    actionType = "73"
    Call Delete_Actions(dategive,dategive,actionExists,actionType,c_OpersView)
    'Ջնջում է Տոկոսի հաշվարկ փաստաթուղթը
    dategive = "050617"
    actionType = "211"
    Call Delete_Actions(dategive,dategive,actionExists,actionType,c_OpersView)

    'Ջնջում է Պարտքերի մարման փաստաթուղթը     
    dategive = "080517"
    actionType = "211"
    Call Delete_Actions(dategive,dategive,actionExists,actionType,c_OpersView)

    Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").Close()
    'Ջնջում է Տոկոսի հաշվարկ փաստաթուղթը
    Workspace = "|Ü»ñ·ñ³íí³Í í³ñÏ»ñ|"
    dategive = "^A[Del]"
    actionType = ""
    Call GroupDelete(Workspace, DocType, CreditLine.DocNum, dategive, dategive, actionType)
    'Ջնջում է գլխավոր պայմանագիրը
    Log.Message(CreditLine.DocNum)
    Call wTreeView.DblClickItem(Workspace & "ä³ÛÙ³Ý³·ñ»ñ")
    Call Rekvizit_Fill("Dialog", 1, "General", "LEVEL", DocType) 
    Call Rekvizit_Fill("Dialog", 1, "General", "NUM",  CreditLine.DocNum) 
  	Call ClickCmdButton(2, "Î³ï³ñ»É")
    'Ջնջում է գլխավոր պայմանագիրը
    Call Delete_Doc()
    Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").Close()
    'Փակում է ASBANK - ը
    Call Close_AsBank()   
  
End Sub