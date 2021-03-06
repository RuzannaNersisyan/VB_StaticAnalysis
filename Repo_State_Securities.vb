Option Explicit
'USEUNIT Loan_Agreements_With_Schedule_Linear_Library
'USEUNIT Loan_Agreemnts_With_Schedule_Library
'USEUNIT Payment_Order_ConfirmPhases_Library
'USEUNIT Overdraft_NewCases_Library
'USEUNIT Deposit_Contract_Library
'USEUNIT Group_Operations_Library
'USEUNIT Derivative_Tools_Library
'USEUNIT Loan_Agreements_Library 
'USEUNIT Subsystems_SQL_Library
'USEUNIT Credit_Line_Library
'USEUNIT Akreditiv_Library
'USEUNIT Library_Common 
'USEUNIT Repo_Library
'USEUNIT Constants
'USEUNIT Mortgage_Library
 
'Test Case Id 166633 

'Ռեպո համաձայնագրի համար տեստային սցենար(ՀՀ պետական արժեթղթեր)
Sub Repo_State_Securities_Test()

    Dim fDATE, sDATE, attr, frmAsMsgBox, FrmSpr, opDate, exTerm, MainSum, PerSum, Prc, NonUsedPrc
    Dim EffRete, ActRete,calcDate,fOBJECT,mainSumm,perSumm,CloseDate
    Dim Client, Curr, CalcAcc, Summa, Date, SecState, SecName, Nominal, Price, Kindscale
    Dim Percent, GiveDate, Term, DateFill, CheckPayDates, PayDates, Paragraph, Direction
    Dim Sector, Aim, Country,District, Region, PaperCode, fBASE, DocNum, Workspace
    Dim actionDate,actionEndDate,actionExists,actionType, tabN
    Dim queryString, sql_Value, colNum, sql_isEqual,fDocBASE
    
    fDATE = "20250101"
    sDATE = "20030101"
    Call Initialize_AsBank("bank", sDATE, fDATE)
    Call Create_Connection()
    Call ChangeWorkspace(c_RepoAgrs) 

    'Ռեպոի ստեղծում
    CalcAcc = "30220042300"
    Summa = 1200000
    Date = "010118"
    SecState = 1
    SecName = "AMGN36294202"
    Nominal = "800000" 
    Price = "1000000"
    Percent = "12"
    GiveDate = "010118"
    Term = "281218"
    DateFill = 1
    CheckPayDates = 0
    Direction = 2
    Paragraph = 1
    Sector = "U2"
    Aim = "00"
    District = "001"
    Region = "010000008"
    Country = "AM"
    PaperCode = 111
    
    Call Repo_Create(Client, Curr, CalcAcc, Summa, Date, SecState, SecName, Nominal, Price,_
               "", Percent, GiveDate, Term, DateFill, CheckPayDates, PayDates, Paragraph,_
               Direction, Sector, Aim, Country, District, Region, PaperCode, fBASE, DocNum)
  
       'Կատարում ենք SQL ստուգում
        queryString = "select fSTATE from DOCS where fISN = '" & fBASE & "'"
        sql_Value = 1
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
    'Պայմանագրը ուղարկել հաստատման                               
    Call PaySys_Send_To_Verify()
    Log.Message(fBASE)                
    BuiltIn.Delay(2000)   
    'Հաստատել պայմանագիրը
    Call Close_Pttel("frmPttel")  
    Call wTreeView.DblClickItem("|è»åá Ñ³Ù³Ó³ÛÝ³·ñ»ñ|Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I")
    'Լրացնել "Պայմանագարի համար"   
    Call Rekvizit_Fill("Dialog",1,"General","NUM",DocNum)
    'Սեղմել "Կատարել" կոճակը
    Sys.Process("Asbank").VBObject("frmAsUstPar").VBObject("CmdOK").Click()
    BuiltIn.Delay(4000)
    'Հաստատել Հաստատող փաստաթղթեր 1- ում
    Call PaySys_Verify(True)
    Call Close_Pttel("frmPttel")
    BuiltIn.Delay(1000)
    
        'Կատարում ենք SQL ստուգում
        queryString = "select fSTATE from DOCS where fISN = '" & fBASE & "'"
        sql_Value = 7
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
  
    Call LetterOfCredit_Filter_Fill("|è»åá Ñ³Ù³Ó³ÛÝ³·ñ»ñ|", 1, DocNum)
  
    'Ստանում է Ռեպոյով գնված արժեթղթեր փաստաթղթի ISN-ը
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_Folders & "|" & c_SecsFromRepo)
    BuiltIn.Delay(2000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_View)
    fDocBASE = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
    Log.Message(fDocBASE)
    wMDIClient.vbObject("frmASDocForm").Close()
    Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel_2").Close()
    
    '"Թղթապանակներ/Ռեպոյով գնված արժեթղթեր" թղթապանակում դաշտերի ստուգում
    Call Repo_Check_Col_Value("0.00","0.00","0.00","0.00","0.00","0.00")

    'Ռեպոյի տրամադրում
    Call Repo_Provide(Date)
    
         'Կատարում ենք SQL ստուգում
        queryString = "select count(*)  from HIF where fOBJECT = '" & fBASE & "'"
        sql_Value = 28
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fCURSUM  from HIR where fOBJECT = '" & fBASE & "' and fTYPE = 'R1'"
        sql_Value = 1200000.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If

        queryString = "select fCURSUM  from HIR where fOBJECT = '" & fBASE & "' and fTYPE = 'R9' "
        sql_Value = 1000000.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
    
        queryString = "select fLASTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'R1'"
        sql_Value = 1200000.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If

        queryString = "select fLASTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'R9' "
        sql_Value = 1000000.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fCURSUM  from HIR where fOBJECT = '" & fDocBASE & "' and fTYPE = 'HN'"
        sql_Value = 800000.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If

        queryString = "select fCURSUM  from HIR where fOBJECT = '" & fDocBASE & "' and fTYPE = 'HP' "
        sql_Value = 1000000.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
    
        queryString = "select fLASTREM  from HIRREST  where fOBJECT = '" & fDocBASE & "' and fTYPE = 'HN'"
        sql_Value = 800000.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If

        queryString = "select fLASTREM  from HIRREST  where fOBJECT = '" & fDocBASE & "' and fTYPE = 'HP' "
        sql_Value = 1000000.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
    Call LetterOfCredit_Filter_Fill("|è»åá Ñ³Ù³Ó³ÛÝ³·ñ»ñ|", 1, DocNum)
    
    '"Թղթապանակներ/Ռեպոյով գնված արժեթղթեր" թղթապանակում դաշտերի ստուգում
    Call Repo_Check_Col_Value("800,000.00", "0.00", "0.00", "1,000,000.00", "0.00", "0.00")
    
    'Տոկոսի հաշվարկ
    calcDate = "310118"
    Call Calculate_Percent(fOBJECT , calcDate , calcDate)
    Log.Message(fOBJECT)
    
        'Կատարում ենք SQL ստուգում
        queryString = "select count(*)  from HIF where fOBJECT = '" & fBASE & "'"
        sql_Value = 29
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fCURSUM  from HIR where fOBJECT = '" & fBASE & "' and fTYPE = 'R2'"
        sql_Value = 12230.10
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fCURSUM  from HIR where fOBJECT = '" & fBASE & "' and fTYPE = 'R¾'"
        sql_Value = 1.50
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If

        queryString = "select fCURSUM  from HIR where fOBJECT = '" & fBASE & "' and fTYPE = 'R¸' "
        sql_Value = 12230.10
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
    
        queryString = "select fLASTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'R2'"
        sql_Value = 12230.10
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'R¾'"
        sql_Value = 1.50
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If

        queryString = "select fLASTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'R¸' "
        sql_Value = 12230.10
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fCURSUM  from HIT where fOBJECT = '" & fBASE & "' and fTYPE = 'N2'"
        sql_Value = 12230.10
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If

        queryString = "select fCURSUM  from HIT where fOBJECT = '" & fBASE & "' and fTYPE = 'N¾' "
        sql_Value = 1.50
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
    
    'Պահուստավորում
    Date = "010218"
    Call FillDoc_Store(Date,fOBJECT)
    Log.Message(fOBJECT)
    
        'Կատարում ենք SQL ստուգում
        queryString = "select fCURSUM from HIR where fOBJECT = '" & fBASE & "'  and fTYPE = 'R4'"
        sql_Value =2322.3
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM  from HIRREST where fOBJECT = '" & fBASE & "' and fTYPE = 'R4'"
        sql_Value = 2322.3
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
    'Պարտքերի մարում
    Date = "010218"
    perSum = "12,230.10"
    tabN = 2
    Call Debt_Repayment(fOBJECT,Date, mainSum,perSum,2,CalcAcc,docNum, tabN)
    Log.Message(fOBJECT)    
    
         'Կատարում ենք SQL ստուգում
        queryString = "select fCURSUM  from HIR where fOBJECT = '" & fBASE & "' and fTYPE = 'R2' and fOP = 'PER'"
        sql_Value = 12230.10
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fCURSUM  from HIR where fOBJECT = '" & fBASE & "' and fTYPE = 'R2' and fOP = 'DBT'"
        sql_Value = 12230.10
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If

        queryString = "select fCURSUM  from HIR where fOBJECT = '" & fBASE & "' and fTYPE = 'R¸' and fOP = 'DBT' "
        sql_Value = 12230.10
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
         queryString = "select fCURSUM  from HIR where fOBJECT = '" & fBASE & "' and fTYPE = 'R¸' and fOP = 'PRJ' "
        sql_Value = 12230.10
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
    
        queryString = "select fLASTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'R2'"
        sql_Value = 0.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If

        queryString = "select fLASTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'R¸' "
        sql_Value = 0.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'R2' "
        sql_Value = 12230.10
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
    
    'Տոկոսի հաշվարկ
    calcDate = "010218"
    Call Calculate_Percent(fOBJECT , calcDate , calcDate)
    
         'Կատարում ենք SQL ստուգում
        queryString = "select fCURSUM  from HIT where fOBJECT = '" & fBASE & "' and fTYPE = 'N2' and fDATE = '2018-02-01'"
        sql_Value = 394.50
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fCURSUM  from HIT where fOBJECT = '" & fBASE & "' and fTYPE = 'N¾' and fDATE = '2018-02-01'"
        sql_Value = -1.80
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If

        queryString = "select fCURSUM  from HIR where fOBJECT = '" & fBASE & "' and fTYPE = 'R2' and fDATE = '2018-02-01' "
        sql_Value = 12230.1
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If        
        
        queryString = "select fLASTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'R¾'"
        sql_Value = -0.30
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If

        queryString = "select fLASTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'R2' "
        sql_Value = 394.50
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'R¾' "
        sql_Value = 1.50
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
    'Տոկոսի հաշվարկ
    calcDate = "280218"
    Call Calculate_Percent(fOBJECT , calcDate , calcDate)
    
        'Կատարում ենք SQL ստուգում
        queryString = "select fLASTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'R¾'"
        sql_Value =  -2.50
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If

        queryString = "select fLASTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'R2' "
        sql_Value = 11046.60
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'R¾' "
        sql_Value = -0.30
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
    'Ռեպո համ-ով գնված արժ.գնի ճշտում
    Date = "020218"
    Call Repo_Secur_Cost_Correction(Date,Date)
    
        'Կատարում ենք SQL ստուգում
        queryString = "select count(*)  from HIF where fOBJECT = '" & fBASE & "'"
        sql_Value = 32
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fCURSUM  from HIR where fOBJECT = '" & fBASE & "' and fTYPE = 'RQ'"
        sql_Value = -171878.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If        
        
        queryString = "select fLASTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'RQ'"
        sql_Value = -171878.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fCURSUM  from HIR where fOBJECT = '" & fDocBASE & "' and fTYPE = 'HQ' "
        sql_Value = -171878.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
    
        queryString = "select fLASTREM  from HIRREST  where fOBJECT = '" & fDocBASE & "' and fTYPE = 'HQ'"
        sql_Value = -171878.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
    
    '"Թղթապանակներ/Ռեպոյով գնված արժեթղթեր" թղթապանակում դաշտերի ստուգում
    BuiltIn.Delay(3000)
    Call Repo_Check_Col_Value("800,000.00", "0.00", "0.00", "828,122.00", "0.00", "0.00")
    
    'Տոկոսների խմբային հաշվարկ
    Date = "030918"
    Call Group_Persent_Calculate(Date,Date)
    BuiltIn.Delay(1000)
    
        'Կատարում ենք SQL ստուգում        
        queryString = "select count(*)  from HIR where fOBJECT = '" & fBASE & "'"
        sql_Value = 46
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If        
        
        queryString = "select fLASTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'R2'"
        sql_Value =  84821.80
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If

        queryString = "select fLASTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'R¾' "
        sql_Value = 5.30
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'R¸' "
        sql_Value = 84427.30
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'R2' "
        sql_Value = 11046.6
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'R¾' "
        sql_Value = -2.5
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'R¸' "
        sql_Value = 71408.10
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
    'Ռեպո համ-ով գնված արժ.գնի ճշտում
    Date = "040918"
    Call Repo_Secur_Cost_Correction(Date,Date)
    
        'Կատարում ենք SQL ստուգում        
        queryString = "select count(*)  from HIF where fOBJECT = '" & fBASE & "'"
        sql_Value = 53
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If        
        
        queryString = "select fCURSUM  from HIR where fOBJECT = '" & fBASE & "' and fTYPE = 'RQ'  and fDATE = '2018-09-04'"
        sql_Value = 3974.20
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'RQ'"
        sql_Value =  -167903.80
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fCURSUM  from HIR where fOBJECT = '" & fDocBASE & "' and fTYPE = 'HQ'  and fDATE = '2018-09-04'"
        sql_Value = 3974.20
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
    
        queryString = "select fLASTREM  from HIRREST  where fOBJECT = '" & fDocBASE & "' and fTYPE = 'HQ'"
        sql_Value = -167903.80
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM  from HIRREST  where fOBJECT = '" & fDocBASE & "' and fTYPE = 'HQ'"
        sql_Value = -171878.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
    
    BuiltIn.Delay(5000)
    '"Թղթապանակներ/Ռեպոյով գնված արժեթղթեր" թղթապանակում դաշտերի ստուգում
    Call Repo_Check_Col_Value("800,000.00", "0.00", "0.00", "832,096.20", "0.00", "0.00")
    
    'Արժեթղթերի վաճառք
    Date = "040918"
    Call Repo_Sell_Security(Date, mainSum, 2, CalcAcc)
    
    Call LetterOfCredit_Filter_Fill("|è»åá Ñ³Ù³Ó³ÛÝ³·ñ»ñ|", 1, DocNum)
    
    '"Թղթապանակներ/Ռեպոյով գնված արժեթղթեր" թղթապանակում դաշտերի ստուգում
    Call Repo_Check_Col_Value("0.00", "800,000.00", "0.00", "0.00", "832,096.20", "0.00")
    
    'Տոկոսի հաշվարկ
    calcDate = "300918"
    Call Calculate_Percent(fOBJECT , calcDate , calcDate)
    Log.Message(fOBJECT)
         'Կատարում ենք SQL ստուգում
        queryString = "select count(*)  from HIR where fOBJECT = '" & fBASE & "'"
        sql_Value = 54
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If  
        
        queryString = "select fLASTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'R2'"
        sql_Value =  95473.90
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If

        queryString = "select fLASTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'R¾' "
        sql_Value = 3.20
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'R¸'"
        sql_Value =  95473.90
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If

        queryString = "select fLASTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'RS' "
        sql_Value = 1000000.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'RR' "
        sql_Value = -167903.80
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'R¾' "
        sql_Value = 5.30
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
    
        queryString = "select fPENULTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'R2' "
        sql_Value = 84821.80
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'R¸' "
        sql_Value = 84427.30
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'R9' "
        sql_Value = 1000000.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'RQ' "
        sql_Value = -171878.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fCURSUM  from HIR where fOBJECT = '" & fDocBASE & "' and fTYPE = 'HL'"
        sql_Value = 1000000.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fCURSUM  from HIR where fOBJECT = '" & fDocBASE & "' and fTYPE = 'HN' and fOP = 'AGR'"
        sql_Value = 800000.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fCURSUM  from HIR where fOBJECT = '" & fDocBASE & "' and fTYPE = 'HN' and fOP = 'SEL' "
        sql_Value = 800000.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fCURSUM  from HIR where fOBJECT = '" & fDocBASE & "' and fTYPE = 'HP' and fOP = 'AGR'"
        sql_Value = 1000000.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fCURSUM  from HIR where fOBJECT = '" & fDocBASE & "' and fTYPE = 'HP' and fOP = 'SEL'"
        sql_Value = 1000000.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fCURSUM  from HIR where fOBJECT = '" & fDocBASE & "' and fTYPE = 'HQ' and fOP = 'SEL'"
        sql_Value = -167903.8 '''0
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fCURSUM  from HIR where fOBJECT = '" & fDocBASE & "' and fTYPE = 'HR'"
        sql_Value = -167903.80
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fCURSUM  from HIR where fOBJECT = '" & fDocBASE & "' and fTYPE = 'HS'"
        sql_Value = 800000.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
    
        queryString = "select fLASTREM  from HIRREST  where fOBJECT = '" & fDocBASE & "' and fTYPE = 'HL'"
        sql_Value = 1000000.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM  from HIRREST  where fOBJECT = '" & fDocBASE & "' and fTYPE = 'HN'"
        sql_Value = 0.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM  from HIRREST  where fOBJECT = '" & fDocBASE & "' and fTYPE = 'HP'"
        sql_Value = 0.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM  from HIRREST  where fOBJECT = '" & fDocBASE & "' and fTYPE = 'HQ'"
        sql_Value = 0.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM  from HIRREST  where fOBJECT = '" & fDocBASE & "' and fTYPE = 'HR'"
        sql_Value = -167903.80
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM  from HIRREST  where fOBJECT = '" & fDocBASE & "' and fTYPE = 'HS'"
        sql_Value = 800000.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select count(*) from HIR where fOBJECT = '" & fDocBASE & "'"
        sql_Value = 10
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
    'Տոկոսի հաշվարկ
    calcDate = "011018"
    Call Calculate_Percent(fOBJECT , calcDate , calcDate)
    Log.Message(fOBJECT)
    BuiltIn.Delay(2000)
    
         'Կատարում ենք SQL ստուգում
        queryString = "select count(*)  from HIR where fOBJECT = '" & fBASE & "'"
        sql_Value = 56
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If  
    
        queryString = "select fLASTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'R2' "
        sql_Value = 95868.40
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'R¾' "
        sql_Value = 1.30
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'R¾' "
        sql_Value = 3.20
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
    
        queryString = "select fPENULTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'R2' "
        sql_Value = 95473.90
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
    'Արժեթղթերի գնում
    Date = "040918"
    mainSum = 500000
    Call Repo_Buy_Security(Date, mainSum, 2, CalcAcc)
    
        'Կատարում ենք SQL ստուգում
        queryString = "select count(*)  from HIR where fOBJECT = '" & fBASE & "'"
        sql_Value = 60
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If  
    
        queryString = "select fLASTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'RS' "
        sql_Value = 0.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'RR' "
        sql_Value = 0.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If        
        
        queryString = "select fLASTREM  from HIRREST where fOBJECT = '" & fDocBASE & "' and fTYPE = 'HL'"
        sql_Value = 0.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM  from HIRREST where fOBJECT = '" & fDocBASE & "' and fTYPE = 'HN'"
        sql_Value = 800000.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM  from HIRREST where fOBJECT = '" & fDocBASE & "' and fTYPE = 'HP'"
        sql_Value = 1000000.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM  from HIRREST where fOBJECT = '" & fDocBASE & "' and fTYPE = 'HQ'"
        sql_Value = -167903.80
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM  from HIRREST where fOBJECT = '" & fDocBASE & "' and fTYPE = 'HR'"
        sql_Value = 0.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM  from HIRREST where fOBJECT = '" & fDocBASE & "' and fTYPE = 'HS'"
        sql_Value = 0.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM  from HIRREST where fOBJECT = '" & fDocBASE & "' and fTYPE = 'HL'"
        sql_Value = 0.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM  from HIRREST where fOBJECT = '" & fDocBASE & "' and fTYPE = 'HN'"
        sql_Value = 800000.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
    
        queryString = "select fPENULTREM  from HIRREST  where fOBJECT = '" & fDocBASE & "' and fTYPE = 'HP'"
        sql_Value = 1000000.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM  from HIRREST  where fOBJECT = '" & fDocBASE & "' and fTYPE = 'HQ'"
        sql_Value = -171878.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM  from HIRREST  where fOBJECT = '" & fDocBASE & "' and fTYPE = 'HR'"
        sql_Value = 0.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM  from HIRREST  where fOBJECT = '" & fDocBASE & "' and fTYPE = 'HS'"
        sql_Value = 0.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select count(*) from HIR where fOBJECT = '" & fDocBASE & "'"
        sql_Value = 16
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
            
    Call LetterOfCredit_Filter_Fill("|è»åá Ñ³Ù³Ó³ÛÝ³·ñ»ñ|", 1, DocNum)
    
    '"Թղթապանակներ/Ռեպոյով գնված արժեթղթեր" թղթապանակում դաշտերի ստուգում
    Call Repo_Check_Col_Value("800,000.00", "0.00", "0.00", "832,096.20", "0.00", "0.00")
    
    'Տոկոսների խմբային հաշվարկ
    Date = "271218"
    Call Group_Persent_Calculate(Date,Date)
    BuiltIn.Delay(1000)
    
        'Կատարում ենք SQL ստուգում
        queryString = "select count(*)  from HIR where fOBJECT = '" & fBASE & "'"
        sql_Value = 74
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If  
    
        queryString = "select fLASTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'R2' "
        sql_Value = 130191.60
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'R¾' "
        sql_Value = 0.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'RÄ' "
        sql_Value = 1200000.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'R¾' "
        sql_Value = 1.30
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
    
        queryString = "select fPENULTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'R2' "
        sql_Value = 95868.40
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'R¸' "
        sql_Value = 120328.60
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
    
    'Ռեպո համ-ով գնված արժ.գնի ճշտում
    Call Repo_Secur_Cost_Correction(Date,Date)
    
        'Կատարում ենք SQL ստուգում
        queryString = "select fLASTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'RQ' "
        sql_Value = -181387.70
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'RQ' "
        sql_Value = -167903.80
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM  from HIRREST  where fOBJECT = '" & fDocBASE & "' and fTYPE = 'HQ'"
        sql_Value = -181387.70
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM  from HIRREST  where fOBJECT = '" & fDocBASE & "' and fTYPE = 'HQ'"
        sql_Value = -167903.80
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
    
        queryString = "select count(*) from HIR where fOBJECT = '" & fDocBASE & "'"
        sql_Value = 17
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
    '"Թղթապանակներ/Ռեպոյով գնված արժեթղթեր" թղթապանակում դաշտերի ստուգում
    Call Repo_Check_Col_Value("800,000.00", "0.00", "0.00", "818,612.30", "0.00", "0.00")
    
    'Պարտքերի մարում
    Date = "281218"
    Call Debt_Repayment(fOBJECT,Date, mainSumm,perSumm,2,CalcAcc,docNum, tabN)
    Log.Message(fOBJECT)
    BuiltIn.Delay(1000)
    
         'Կատարում ենք SQL ստուգում    
        queryString = "select fLASTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'R1' "
        sql_Value = 0'''--1200000.00 
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'R2' "
        sql_Value = 0.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'R9' "
        sql_Value = 0.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'RQ'"
        sql_Value = 0.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'R¸' "
        sql_Value = 0.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'R¾' "
        sql_Value = 0.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'RÄ' "
        sql_Value = 0.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'R1' "
        sql_Value = 1200000.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
    
        queryString = "select fPENULTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'R2' "
        sql_Value = 130191.60
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'R9' "
        sql_Value = 1000000.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'RQ' "
        sql_Value = -181387.70
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
    
        queryString = "select fPENULTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'R¸' "
        sql_Value = 120328.60
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'R¾' "
        sql_Value = 1.30
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM  from HIRREST  where fOBJECT = '" & fBASE & "' and fTYPE = 'RÄ' "
        sql_Value = 0.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM  from HIRREST  where fOBJECT = '" & fDocBASE & "' and fTYPE = 'HN'"
        sql_Value = 0.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM  from HIRREST  where fOBJECT = '" & fDocBASE & "' and fTYPE = 'HP'"
        sql_Value = 0.00
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fLASTREM  from HIRREST  where fOBJECT = '" & fDocBASE & "' and fTYPE = 'HQ'"
        sql_Value = 0.30
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
        
        queryString = "select fPENULTREM  from HIRREST  where fOBJECT = '" & fDocBASE & "' and fTYPE = 'HQ'"
        sql_Value = -181387.70
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
    
        queryString = "select count(*) from HIR where fOBJECT = '" & fDocBASE & "'"
        sql_Value = 20
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
    
    '"Թղթապանակներ/Ռեպոյով գնված արժեթղթեր" թղթապանակում դաշտերի ստուգում
    Call Repo_Check_Col_Value("0.00", "0.00", "0.00", "0.30", "0.00", "0.00")
    
    'Տոկոսի հաշվարկ
    calcDate = "281218"
    Call Calculate_Percent(fOBJECT , calcDate , calcDate)
    
    'Պայմանագրի փակում
    CloseDate = "010119"
    Call Close_Contract(CloseDate)
    BuiltIn.Delay(2000)
    
    'Ստուգում է փակման ամսաթիվ սյունը
    colNum =	wMDIClient.VBObject("frmPttel").GetColumnIndex("fDATECLOSE")
    If Not Trim(wMDIClient.vbObject("frmPttel").vbObject("tdbgView").Columns.Item(colNum).Text) = Trim("01/01/19") Then
          Log.Error("Don't match")
    End If
   
    'Բացում է պայմանագիրը
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_AgrOpen)
    Call ClickCmdButton(5,"²Ûá")
  
    actionDate = "281218"
    actionEndDate = "281218"
    actionExists = False
    
    'Ջնջում է Դիտում և խմբագրում/Այլ/Հաշվարկման ամսաթվեր թղթապանակի բոլոր փաստաթղթերը
    Call Delete_Actions(actionDate,actionEndDate,actionExists,actionType,c_ViewEdit & "|" & c_Other & "|" & c_CalcDates)

    'Ջնջում է Ռեպոյի մարում գործողությունը
    actionExists = true
    actionType = "22"
    Call Delete_Actions(actionDate,actionDate,actionExists,actionType,c_OpersView )

    'Ջնջում է Ռեպո համ-ով ձեռք բերված արժեթղթերի գնի ճշտում փաստաթուղթը
    actionDate = "271218"
    actionType = "C6"
    Call Delete_Actions(actionDate,actionDate,actionExists,actionType,c_OpersView )

    'Ջնջում է Տոկոսի հաշվարկ փաստաթութը
    actionDate = "300918"
    actionEndDate = "271218"
    actionType = 511
    Call Delete_Actions(actionDate,actionEndDate,actionExists,actionType,c_OpersView )

    'Ջնջում է Ռեպո համ-ով ձեռք բերված արժեթղթերի գնի ճշտում փաստաթուղթը
    actionDate = "040918"
    actionType = "C6"
    Call Delete_Actions(actionDate,actionDate,actionExists,actionType,c_OpersView )

    'Ջնջում է Ռեպո համ-ով ձեռք բերված արժեթղթերի վաճառված մասի առք փաստաթուղթը
    actionType = "B4"
    Call Delete_Actions(actionDate,actionDate,actionExists,actionType,c_OpersView )

    'Ջնջում է Ռեպո համ-ով ձեռք բերված արժեթղթերի վաճառք փաստաթուղթը
    actionType = "B3"
    Call Delete_Actions(actionDate,actionDate,actionExists,actionType,c_OpersView )

    'Ջնջում է Տոկոսի հաշվարկ փաստաթութը
    actionDate = "010218"
    actionEndDate = "030918"
    actionType = 511
    Call Delete_Actions(actionDate,actionEndDate,actionExists,actionType,c_OpersView )

    'Ջնջում է Ռեպո համ-ով ձեռք բերված արժեթղթերի գնի ճշտում փաստաթուղթը
    actionDate = "020218"
    actionEndDate = "040918"
    actionType = "C6"
    Call Delete_Actions(actionDate,actionDate,actionExists,actionType,c_OpersView )

    'Ջնջում է Տոկոսի մարում փաստաթութը
    actionDate = "010218"
    actionType = "53"
    Call Delete_Actions(actionDate,actionDate,actionExists,actionType,c_OpersView )

    'Ջնջում է Տոկոսի հաշվարկ փաստաթութը
    actionDate = "310118"
    actionType = 511
    Call Delete_Actions(actionDate,actionEndDate,actionExists,actionType,c_OpersView )

    'Ջնջում է Գործողությունների դիտում թղթապանակի մնացած գործողությունները
    actionDate = "010118"
    actionEndDate = "030918"
    actionType = Null
    Call Delete_Actions(actionDate,actionEndDate,actionExists,actionType,c_OpersView )

    'Ջնջում է գլխավոր պայմանագիրը
    Call Delete_Doc()

    'Փակել ASBANK համակարգը
    Call Close_AsBank()
    
End Sub