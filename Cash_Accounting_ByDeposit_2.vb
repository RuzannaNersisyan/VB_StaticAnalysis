'USEUNIT Library_Common
'USEUNIT Library_Colour
'USEUNIT Library_CheckDB
'USEUNIT Mortgage_Library
'USEUNIT Library_Contracts
'USEUNIT Card_Library
'USEUNIT Loan_Agreements_Library 
'USEUNIT Akreditiv_Library
'USEUNIT CashInput_Confirmphases_Library
'USEUNIT Main_Accountant_Filter_Library
'USEUNIT Deposit_Contract_Library
'USEUNIT Payment_Order_ConfirmPhases_Library

'USEUNIT Constants
Option Explicit

'Test Case Id - 176136

Sub Cash_Accounting_ByDeposit_2()
    
    Dim sDATE,eDATE
    Dim Acc,FolderName,CashIn,CashInIsn
    Dim VerificationDoc,CashAccountingFilter,verifyFilter1
    Dim fBASE,fBODY,dbHI2
    
    'Համակարգ մուտք գործել ARMSOFT օգտագործողով
    sDATE = "20030101"
    eDATE = "20260101"
    Call Initialize_AsBank("bank", sDATE, eDATE)
    Login("ARMSOFT")
    
    'Մուտք գործել "Ենթահամակրգեր(ՀԾ)"
    Call ChangeWorkspace(c_Subsystems)
    wTreeView.DblClickItem("|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|²¹ÙÇÝÇëïñ³ïÇí Ù³ë|Ð³Ù³Ï³ñ·³ÛÇÝ ³ßË³ï³ÝùÝ»ñ|²ÛÉ|²¹ÙÇÝÇëïñ³ïÇí ÷á÷áËáõÃÛáõÝÝ»ñ »ÝÃ³Ñ³Ù³Ï³ñ·»ñáõÙ")
    Call Rekvizit_Fill("Dialog",1,"General","FUNCTIONS","²í³Ý¹Ý»ñáõÙ Ï³ÝËÇÏ Ñ³ßí³éÙ³Ý Ó¨³íáñáõÙ")
  
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    Call ClickCmdButton(5, "²Ûá")
    BuiltIn.Delay(2000)
    
    If p1.VBObject("frmAsMsgBox").VBObject("lblMessage").NativeVBObject = "üáõÝÏóÇ³ÛÇ ³ßË³ï³ÝùÁ ³í³ñïí³Í ¿" Then
        Call ClickCmdButton(5, "OK")
    Else
        Call MessageExists(2, "àõÝÇÏ³É ¹³ßïÇ ÏñÏÝáõÃÛáõÝ")
        Call ClickCmdButton(5, "OK")
    End If
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''-- Կարգավորումներում փոփոխել "Միայն տոկոսներ" նշիչը դրած վիճակով --''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Կարգավորումներում փոփոխել (Միայն տոկոսներ) նշիչը դրած վիճակով --" ,,, DivideColor     
        
    Call wTreeView.DblClickItem("|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|²¹ÙÇÝÇëïñ³ïÇí Ù³ë|Î³ñ·³íáñáõÙÝ»ñ ¨ ¹ñáõÛÃÝ»ñ|²ÝÏ³ÝËÇÏ ·áñÍ³ñùÝ»ñÇó Ï³ÝËÇÏÇ Ñ³ßí³éÙ³Ý ¹ñáõÛÃÝ»ñ|²í³Ý¹Ý»ñ (Ý»ñ·ñ³íí³Í)")
    BuiltIn.Delay(2000)
    With wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame").VBObject("DocGrid")
      .Row = 0
      .Col = 0
      .Keys("^A[Del]")
      
      .Row = 0
      .Col = 14
      .Text = "-1"
    End With
    Call ClickCmdButton(1, "Î³ï³ñ»É")  
    
    'SQL Ստուգում DOCSG աղուսյակում 
    Call CheckQueryRowCount("DOCSG","fISN","131889730",2)
    Call CheckDB_DOCSG("131889730","GRIDINST","0","ONLYCASHATTRPART","0",1)
    Call CheckDB_DOCSG("131889730","GRIDINST","0","ONLYPER","1",1)
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''-- Հաշիվներ թղթապանակից կատարել "Ավելացնել" գործողությունը --''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Հաշիվներ թղթապանակից կատարել Ավելացնել գործողությունը --" ,,, DivideColor    
    
    'Մուտք Գլխավոր հաշվապահի ԱՇՏ
    Call ChangeWorkspace(c_ChiefAcc)
    
    Call wTreeView.DblClickItem("|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|Ð³ßÇíÝ»ñ") 
    BuiltIn.Delay(1000)
    'Կանխիկ հաշվառման և Բացման ամսաթիվ դաշտի լրացում
    Call Rekvizit_Fill("Dialog", 1, "CheckBox", "CASHAC", 1)
    Call Rekvizit_Fill("Dialog", 1, "General", "DATOTKN", "010120" &"[Tab]"& "010120")
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    BuiltIn.Delay(2000)

    Set Acc = New_Account()
    With Acc         
      .BalanceAccount = "3022000"
      .AccountHolder = "00000678"
      .Name = ""
      .EnglishName = ""
      .RemainderType = ""
      .Curr = "000"
      .AccountType = "01"
      .OpenDate = "010120"
      .Account = ""
      .AccessType = "01"
      .CashAccounting = 1
    End With
    
    Call Create_Account(Acc)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''-- Հաշիվներ թղթապանակից կատարել "Կանխիկ մուտք" գործողությունը --'''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Հաշիվներ թղթապանակից կատարել Կանխիկ մուտք գործողությունը --" ,,, DivideColor 
    Log.Message Acc.Account
    Log.Message Acc.Isn,,,SqlDivideColor
    
    Call SearchInPttel("frmPttel",0, Acc.Account)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_InnerOpers &"|"& c_Cashin)             
      
    Set CashIn = New_CashIn()  
    With CashIn
        .Date = "010120"
        .Amount = "100000"
        .CashLabel = "022"
        .Base = "Ð³Ù³Ó³ÛÝ å³ÛÙ³Ý³·ñÇ"
        .Aim = "Ð³Ù³Ó³ÛÝ Ã. Ñ³ßíÇ"
        .Depositor = "00000678"
        .FirstName = "master"
    End With 

    CashInIsn = Fill_CashIn(CashIn)
    
    'Եթե քաղվածքի պատուհանը հայտնվել է, ապա փակում է
    If wMDIClient.VBObject("FrmSpr").Exists Then
        wMDIClient.VBObject("FrmSpr").Close
    Else
        Log.Error "Statement window doesn't exist!",,,ErrorColor
    End If
    Call Close_Pttel("frmPttel")
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''-- "Աշխատանքային փաստաթղթեր" թղթապանակից կատարել "Ուղարկել հաստատման" գործողությունը --'''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Կատարել Ուղարկել հաստատման գործողությունը --",,,DivideColor       
    
    wTreeView.DblClickItem("|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ") 
    'Լրացնել "Ամսաթիվ" դաշտերը
    Call Rekvizit_Fill("Dialog",1,"General","PERN", "010120")
    Call Rekvizit_Fill("Dialog",1,"General","PERK", "010120")
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    
    If WaitForPttel("frmPttel") Then
        Call wMainForm.MainMenu.Click(c_AllActions)
        Call wMainForm.PopupMenu.Click(c_SendToVer)
        BuiltIn.Delay(2000)
        Call MessageExists(2,"àõÕ³ñÏ»É Ñ³ëï³ïÙ³Ý")
        Call ClickCmdButton(2, "Î³ï³ñ»É")
        Call Close_Pttel("frmPttel")
    Else
        Log.Error "Can Not Open Աշխատանքային փաստաթղթեր pttel",,,ErrorColor      
    End If  

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''-- Գլխավոր հաշվապահ/Հաստատվող փաստաթղթեր(|) թղթապանակից կատարել "Վավերացնել" գործողությունը --''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Կատարել Վավերացնել գործողությունը --",,,DivideColor   

    Set VerificationDoc = New_VerificationDocument()
        VerificationDoc.DocType = "KasPrOrd"
    
    Call GoToVerificationDocument("|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ (I)",VerificationDoc) 
    
    If WaitForPttel("frmPttel") Then
        If SearchInPttel("frmPttel",7, "100000") Then
            Call wMainForm.MainMenu.Click(c_AllActions)
            Call wMainForm.PopupMenu.Click(c_ToConfirm)
            BuiltIn.Delay(2000)
            Call ClickCmdButton(1, "Ð³ëï³ï»É")
        Else 
            Log.Error "Տողը չի գտնվել Հաստատվող փաստաթղթեր(|) թղթապանակում" ,,,ErrorColor
        End If
        Call Close_Pttel("frmPttel")
     Else
        Log.Error "Can Not Open Հաստատվող փաստաթղթեր(|) Window",,,ErrorColor      
     End If   
     
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''-- "Աշխատանքային փաստաթղթեր" թղթապանակից կատարել "Վավերացնել" գործողությունը --'''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Աշխատանքային փաստաթղթեր թղթապանակից կատարել Վավերացնել գործողությունը --",,,DivideColor       
    
    wTreeView.DblClickItem("|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ") 
    'Լրացնել "Ամսաթիվ" դաշտերը
    Call Rekvizit_Fill("Dialog",1,"General","PERN", "010120")
    Call Rekvizit_Fill("Dialog",1,"General","PERK", "010120")
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    
    If WaitForPttel("frmPttel") Then
        Call wMainForm.MainMenu.Click(c_AllActions)
        Call wMainForm.PopupMenu.Click(c_ToConfirm)
        BuiltIn.Delay(2000)
        Call ClickCmdButton(1, "Ð³ëï³ï»É")
        Call Close_Pttel("frmPttel")
    Else
        Log.Error "Can Not Open Աշխատանքային փաստաթղթեր pttel",,,ErrorColor
    End If
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''-- Պարտքերի մարում (Կարգավորումներում միայն տոկոսներ-ով ) --''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Պարտքերի մարում (Կարգավորումներում միայն տոկոսներ-ով) --", "", pmNormal, DivideColor   

    'Մուտք գործել "Ենթահամակրգեր(ՀԾ)"
    Call ChangeWorkspace(c_Subsystems)
    
    FolderName = "|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|ä³ÛÙ³Ý³·ñ»ñ|Ü»ñ·ñ³íí³Í ÙÇçáóÝ»ñ|²í³Ý¹Ý»ñ (Ý»ñ·ñ³íí³Í)|"
    Call LetterOfCredit_Filter_Fill(FolderName, 1, "A-001087")
    Call Debt_Repayment(fBASE,"010907", "19999","","2",Acc.Account,"",2)    
    Call Close_Pttel("frmPttel")
    
    Log.Message fBASE,,,SqlDivideColor
    
    'SQL Ստուգում DOCS աղուսյակում
    fBODY = "  CODE:A-001087  DATE:20070901  CORROFFPER:0  AMDSUMDBT:22292.1  SUMAGR:19999  NOTAXEDSUM:22292.1  SUMPER:2547.9  ALLTAX:254.8  TAXPER:254.8  SUMMA:22546.9  CASHORNO:2  ISPUSA:0  ACCCORR:"&Acc.Account&"  COMMENT:²í³Ý¹Ç å³ñïù»ñÇ í×³ñáõÙ  ACSBRANCH:00  ACSDEPART:1  ACSTYPE:D10  USERID:  77  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",fBASE,1)
    Call CheckDB_DOCS(fBASE,"D1DSDebt","5",fBODY,1)
        
    'SQL Ստուգում HI աղուսյակում  
    Call CheckQueryRowCount("HI","fBASE",fBASE,6)
    Call Check_HI_CE_accounting ("20070901",fBASE, "01", "1630187","254.80", "000", "254.80", "MSC", "C")
    Call Check_HI_CE_accounting ("20070901",fBASE, "01", "241020717","19999.00", "000", "19999.00", "MSC", "D")
    Call Check_HI_CE_accounting ("20070901",fBASE, "01", "877353073","2547.90", "000", "2547.90", "MSC", "D")
    Call Check_HI_CE_accounting ("20070901",fBASE, "01", Acc.Isn,"19999.00", "000", "19999.00", "MSC", "C")
    Call Check_HI_CE_accounting ("20070901",fBASE, "01", Acc.Isn,"2547.90", "000", "2547.90", "MSC", "C")
    Call Check_HI_CE_accounting ("20070901",fBASE, "01", Acc.Isn,"254.80", "000", "254.80", "MSC", "D")
    
    'SQL Ստուգում HI2 աղուսյակում 
    Call CheckQueryRowCount("HI2","fBASE",fBASE,1)
    Set dbHI2 = New_DB_HI2()
    With dbHI2
        .fDATE = "2007-09-01"
        .fTYPE = "10"
        .fOBJECT = "737994605"
        .fGLACC = "1559631"
        .fSUM = "0.00"
        .fCUR = "000"
        .fCURSUM = "2293.10"
        .fOP = "MSC"
        .fBASE = fBASE
        .fDBCR = "D"
    End With
    Call CheckDB_HI2(dbHI2,1)
    
    'SQL Ստուգում HIR աղուսյակում
    Call CheckQueryRowCount("HIR","fBASE",fBASE,6)
    Call Check_HIR("2007-09-01", "R1", "53912814", "000", "19999.00", "DBT", "C")
    Call Check_HIR("2007-09-01", "R2", "53912814", "000", "2293.10", "DBT", "C")
    Call Check_HIR("2007-09-01", "R2", "53912814", "000", "254.80", "TXD", "C")
    Call Check_HIR("2007-09-01", "R¸", "53912814", "000", "2293.10", "DBT", "C")
    Call Check_HIR("2007-09-01", "R¸", "53912814", "000", "254.80", "TXD", "C")
    Call Check_HIR("2007-09-01", "RÄ", "53912814", "000", "19999.00", "DBT", "C")
    
    'SQL Ստուգում HIREST2 աղուսյակում
    Call CheckDB_HIREST2("10","737994605","1559631","0.00","000","102293.10", 1)
    
    'SQL Ստուգում HIREST  աղուսյակում  
    Call CheckDB_HIREST("01", "1630187","-237002.10","000","-237002.10",1)
    Call CheckDB_HIREST("01", Acc.Isn,"-122292.10","000","-122292.10",1)    
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''-- Կանխիկ միջոցների հաշվառում թղթապանակում ստուգել Մնացորդը --''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Կանխիկ միջոցների հաշվառում թղթապանակում ստուգել Մնացորդը --" ,,, DivideColor 
    
    'Մուտք Գլխավոր հաշվապահի ԱՇՏ
    Call ChangeWorkspace(c_ChiefAcc)
    
    Set CashAccountingFilter = New_CashAccounting()
    With CashAccountingFilter
      .ClientCode = "00000678"
      .Curr = "000"
    End With
    Call GoTo_CashAccounting(CashAccountingFilter)     
    
    Call CompareFieldValue("frmPttel", "FKEY", "00000678")
    Call CompareFieldValue("frmPttel", "FCUR", "000")   
    Call CompareFieldValue("frmPttel", "SUM", "102,293.10")   
    Call CompareFieldValue("frmPttel", "FCOM", "KERAMIKA Ê³ã³ïãÛ³Ý ìÉ³¹»Ý ì³ÝÇÏÇ")
    Call Close_Pttel("frmPttel")
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''-- Գործողությունների դիտում թղթապանակից հեռացնել Ավանդի մարում գործողությունը --''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Գործողությունների դիտում թղթապանակից հեռացնել Ավանդի մարում գործողությունը --", "", pmNormal, DivideColor   

    'Մուտք գործել "Ենթահամակրգեր(ՀԾ)"
    Call ChangeWorkspace(c_Subsystems) 
    
    Call LetterOfCredit_Filter_Fill(FolderName, 1, "A-001087")
    Call Delete_Actions("010907","010907",True,"12",c_OpersView)
     
    Call Close_Pttel("frmPttel")
     
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''-- Կարգավորումներում փոփոխել "Ներգրավված գումարը = 80%" վիճակով --''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Կարգավորումներում փոփոխել <Ներգրավված գումարը = 80%> վիճակով --" ,,, DivideColor   
        
    Call wTreeView.DblClickItem("|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|²¹ÙÇÝÇëïñ³ïÇí Ù³ë|Î³ñ·³íáñáõÙÝ»ñ ¨ ¹ñáõÛÃÝ»ñ|²ÝÏ³ÝËÇÏ ·áñÍ³ñùÝ»ñÇó Ï³ÝËÇÏÇ Ñ³ßí³éÙ³Ý ¹ñáõÛÃÝ»ñ|²í³Ý¹Ý»ñ (Ý»ñ·ñ³íí³Í)")
    BuiltIn.Delay(2000)
    With wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame").VBObject("DocGrid")
      .Row = 0
      .Col = 14
      .Text = "0"
      .keys("[Enter]")
      
      .Row = 0
      .Col = 16
      .Keys("80")
    End With
    
    Call ClickCmdButton(1, "Î³ï³ñ»É")
    
    'SQL Ստուգում DOCSG աղուսյակում 
    Call CheckQueryRowCount("DOCSG","fISN","131889730",3)
    Call CheckDB_DOCSG("131889730","GRIDINST","0","ATTRSUMREMMAX","80",1)
    Call CheckDB_DOCSG("131889730","GRIDINST","0","ONLYCASHATTRPART","0",1)
    Call CheckDB_DOCSG("131889730","GRIDINST","0","ONLYPER","0",1)
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''-- Պարտքերի մարում "Ներգրավված գումարը = 80%" --''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Պարտքերի մարում (Ներգրավված գումարը = 80%) --", "", pmNormal, DivideColor   

    Call LetterOfCredit_Filter_Fill(FolderName, 1, "A-001087")
    Call Debt_Repayment(fBASE,"010907", "18400","","2",Acc.Account,"",2)    
    Call Close_Pttel("frmPttel")    
    
    Log.Message fBASE,,,SqlDivideColor
    
    'SQL Ստուգում DOCS աղուսյակում
    fBODY = "  CODE:A-001087  DATE:20070901  CORROFFPER:0  AMDSUMDBT:20693.1  SUMAGR:18400  NOTAXEDSUM:20693.1  SUMPER:2547.9  ALLTAX:254.8  TAXPER:254.8  SUMMA:20947.9  CASHORNO:2  ISPUSA:0  ACCCORR:"&Acc.Account&"  COMMENT:²í³Ý¹Ç å³ñïù»ñÇ í×³ñáõÙ  ACSBRANCH:00  ACSDEPART:1  ACSTYPE:D10  USERID:  77  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",fBASE,1)
    Call CheckDB_DOCS(fBASE,"D1DSDebt","5",fBODY,1)
        
    'SQL Ստուգում HI աղուսյակում  
    Call CheckQueryRowCount("HI","fBASE",fBASE,6)
    Call Check_HI_CE_accounting ("20070901",fBASE, "01", "1630187","254.80", "000", "254.80", "MSC", "C")
    Call Check_HI_CE_accounting ("20070901",fBASE, "01", "241020717","18400.00", "000", "18400.00", "MSC", "D")
    Call Check_HI_CE_accounting ("20070901",fBASE, "01", "877353073","2547.90", "000", "2547.90", "MSC", "D")
    Call Check_HI_CE_accounting ("20070901",fBASE, "01", Acc.Isn,"18400", "000", "18400.00", "MSC", "C")
    Call Check_HI_CE_accounting ("20070901",fBASE, "01", Acc.Isn,"2547.90", "000", "2547.90", "MSC", "C")
    Call Check_HI_CE_accounting ("20070901",fBASE, "01", Acc.Isn,"254.80", "000", "254.80", "MSC", "D")
    
    'SQL Ստուգում HI2 աղուսյակում 
    Call CheckQueryRowCount("HI2","fBASE",fBASE,2)
    Set dbHI2 = New_DB_HI2()
    With dbHI2
        .fDATE = "2007-09-01"
        .fTYPE = "10"
        .fOBJECT = "737994605"
        .fGLACC = "1559631"
        .fSUM = "0.00"
        .fCUR = "000"
        .fCURSUM = "2293.10"
        .fOP = "MSC"
        .fBASE = fBASE
        .fDBCR = "D"
    End With
    Call CheckDB_HI2(dbHI2,1)
    dbHI2.fCURSUM = "18400.00"
    Call CheckDB_HI2(dbHI2,1)
    
    'SQL Ստուգում HIR աղուսյակում
    Call CheckQueryRowCount("HIR","fBASE",fBASE,6)
    Call Check_HIR("2007-09-01", "R1", "53912814", "000", "18400.00", "DBT", "C")
    Call Check_HIR("2007-09-01", "R2", "53912814", "000", "2293.10", "DBT", "C")
    Call Check_HIR("2007-09-01", "R2", "53912814", "000", "254.80", "TXD", "C")
    Call Check_HIR("2007-09-01", "R¸", "53912814", "000", "2293.10", "DBT", "C")
    Call Check_HIR("2007-09-01", "R¸", "53912814", "000", "254.80", "TXD", "C")
    Call Check_HIR("2007-09-01", "RÄ", "53912814", "000", "18400.00", "DBT", "C")
    
    'SQL Ստուգում HIREST2 աղուսյակում
    Call CheckDB_HIREST2("10","737994605","1559631","0.00","000","120693.10", 1)
    
    'SQL Ստուգում HIREST  աղուսյակում  
    Call CheckDB_HIREST("01", "1630187","-237002.10","000","-237002.10",1)
    Call CheckDB_HIREST("01", Acc.Isn,"-120693.10","000","-120693.10",1)    
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''-- Կանխիկ միջոցների հաշվառում թղթապանակում ստուգել Մնացորդը --''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Կանխիկ միջոցների հաշվառում թղթապանակում ստուգել Մնացորդը --" ,,, DivideColor
    
    'Մուտք Գլխավոր հաշվապահի ԱՇՏ
    Call ChangeWorkspace(c_ChiefAcc)
    
    Set CashAccountingFilter = New_CashAccounting()
    With CashAccountingFilter
      .ClientCode = "00000678"
      .Curr = "000"
    End With
    Call GoTo_CashAccounting(CashAccountingFilter)
    
    Call CompareFieldValue("frmPttel", "FKEY", "00000678")
    Call CompareFieldValue("frmPttel", "FCUR", "000")
    Call CompareFieldValue("frmPttel", "SUM", "120,693.10")
    Call CompareFieldValue("frmPttel", "FCOM", "KERAMIKA Ê³ã³ïãÛ³Ý ìÉ³¹»Ý ì³ÝÇÏÇ")
    Call Close_Pttel("frmPttel")
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''-- Գործողությունների դիտում թղթապանակից հեռացնել Ավանդի մարում գործողությունը --''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Գործողությունների դիտում թղթապանակից հեռացնել Ավանդի մարում գործողությունը --", "", pmNormal, DivideColor   

    'Մուտք գործել "Ենթահամակրգեր(ՀԾ)"
    Call ChangeWorkspace(c_Subsystems) 
    
    Call LetterOfCredit_Filter_Fill(FolderName, 1, "A-001087")
    Call Delete_Actions("010907","010907",True,"12",c_OpersView)
     
    Call Close_Pttel("frmPttel")

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''-- Կարգավորումներում փոփոխել "Պայմանագրի տևողությունը = 300օր" վիճակով --''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Կարգավորումներում փոփոխել (Պայմանագրի տևողություն = 300օր) վիճակով --" ,,, DivideColor      
        
    Call wTreeView.DblClickItem("|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|²¹ÙÇÝÇëïñ³ïÇí Ù³ë|Î³ñ·³íáñáõÙÝ»ñ ¨ ¹ñáõÛÃÝ»ñ|²ÝÏ³ÝËÇÏ ·áñÍ³ñùÝ»ñÇó Ï³ÝËÇÏÇ Ñ³ßí³éÙ³Ý ¹ñáõÛÃÝ»ñ|²í³Ý¹Ý»ñ (Ý»ñ·ñ³íí³Í)")
    BuiltIn.Delay(2000)
    With wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame").VBObject("DocGrid")
      
      .Row = 0
      .Col = 16
      .Keys("^A[Del]")
      
      .Row = 0
      .Col = 15
      .Keys("300")
    End With
    
    Call ClickCmdButton(1, "Î³ï³ñ»É")
    
    'SQL Ստուգում DOCSG աղուսյակում 
    Call CheckQueryRowCount("DOCSG","fISN","131889730",3)
    Call CheckDB_DOCSG("131889730","GRIDINST","0","AGRDURMIN","300",1)
    Call CheckDB_DOCSG("131889730","GRIDINST","0","ONLYCASHATTRPART","0",1)
    Call CheckDB_DOCSG("131889730","GRIDINST","0","ONLYPER","0",1)
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''-- Պարտքերի մարում (Պայմանագրի տևողությունը = 300) --''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Պարտքերի մարում <Պայմանագրի տևողություն = 300> --", "", pmNormal, DivideColor   

    Call LetterOfCredit_Filter_Fill(FolderName, 1, "A-001087")
    Call Debt_Repayment(fBASE,"010907", "24590.50","","2",Acc.Account,"",2)    
    Call Close_Pttel("frmPttel")    
    
    Log.Message fBASE,,,SqlDivideColor
    
    'SQL Ստուգում DOCS աղուսյակում
    fBODY = "  CODE:A-001087  DATE:20070901  CORROFFPER:0  AMDSUMDBT:26883.6  SUMAGR:24590.5  NOTAXEDSUM:26883.6  SUMPER:2547.9  ALLTAX:254.8  TAXPER:254.8  SUMMA:27138.4  CASHORNO:2  ISPUSA:0  ACCCORR:"&Acc.Account&"  COMMENT:²í³Ý¹Ç å³ñïù»ñÇ í×³ñáõÙ  ACSBRANCH:00  ACSDEPART:1  ACSTYPE:D10  USERID:  77  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",fBASE,1)
    Call CheckDB_DOCS(fBASE,"D1DSDebt","5",fBODY,1)
        
    'SQL Ստուգում HI աղուսյակում  
    Call CheckQueryRowCount("HI","fBASE",fBASE,6)
    Call Check_HI_CE_accounting ("20070901",fBASE, "01", "1630187","254.80", "000", "254.80", "MSC", "C")
    Call Check_HI_CE_accounting ("20070901",fBASE, "01", "241020717","24590.50", "000", "24590.50", "MSC", "D")
    Call Check_HI_CE_accounting ("20070901",fBASE, "01", "877353073","2547.90", "000", "2547.90", "MSC", "D")
    Call Check_HI_CE_accounting ("20070901",fBASE, "01", Acc.Isn,"24590.50", "000", "24590.50", "MSC", "C")
    Call Check_HI_CE_accounting ("20070901",fBASE, "01", Acc.Isn,"2547.90", "000", "2547.90", "MSC", "C")
    Call Check_HI_CE_accounting ("20070901",fBASE, "01", Acc.Isn,"254.80", "000", "254.80", "MSC", "D")
    
    'SQL Ստուգում HI2 աղուսյակում 
    Call CheckQueryRowCount("HI2","fBASE",fBASE,2)
    Set dbHI2 = New_DB_HI2()
    With dbHI2
        .fDATE = "2007-09-01"
        .fTYPE = "10"
        .fOBJECT = "737994605"
        .fGLACC = "1559631"
        .fSUM = "0.00"
        .fCUR = "000"
        .fCURSUM = "2293.10"
        .fOP = "MSC"
        .fBASE = fBASE
        .fDBCR = "D"
    End With
    Call CheckDB_HI2(dbHI2,1)
    dbHI2.fCURSUM = "24590.50"
    Call CheckDB_HI2(dbHI2,1)
    
    'SQL Ստուգում HIR աղուսյակում
    Call CheckQueryRowCount("HIR","fBASE",fBASE,6)
    Call Check_HIR("2007-09-01", "R1", "53912814", "000", "24590.50", "DBT", "C")
    Call Check_HIR("2007-09-01", "R2", "53912814", "000", "2293.10", "DBT", "C")
    Call Check_HIR("2007-09-01", "R2", "53912814", "000", "254.80", "TXD", "C")
    Call Check_HIR("2007-09-01", "R¸", "53912814", "000", "2293.10", "DBT", "C")
    Call Check_HIR("2007-09-01", "R¸", "53912814", "000", "254.80", "TXD", "C")
    Call Check_HIR("2007-09-01", "RÄ", "53912814", "000", "24590.50", "DBT", "C")
    
    'SQL Ստուգում HIREST2 աղուսյակում
    Call CheckDB_HIREST2("10","737994605","1559631","0.00","000","126883.60", 1)
    
    'SQL Ստուգում HIREST  աղուսյակում  
    Call CheckDB_HIREST("01", "1630187","-237002.10","000","-237002.10",1)
    Call CheckDB_HIREST("01", Acc.Isn,"-126883.60","000","-126883.60",1)   
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''-- Կանխիկ միջոցների հաշվառում թղթապանակում ստուգել Մնացորդը --''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Կանխիկ միջոցների հաշվառում թղթապանակում ստուգել Մնացորդը --" ,,, DivideColor 
    
    'Մուտք Գլխավոր հաշվապահի ԱՇՏ
    Call ChangeWorkspace(c_ChiefAcc)
    
    Set CashAccountingFilter = New_CashAccounting()
    With CashAccountingFilter
      .ClientCode = "00000678"
      .Curr = "000"
    End With
    Call GoTo_CashAccounting(CashAccountingFilter)     
    
    Call CompareFieldValue("frmPttel", "FKEY", "00000678")
    Call CompareFieldValue("frmPttel", "FCUR", "000")   
    Call CompareFieldValue("frmPttel", "SUM", "126,883.60")
    Call CompareFieldValue("frmPttel", "FCOM", "KERAMIKA Ê³ã³ïãÛ³Ý ìÉ³¹»Ý ì³ÝÇÏÇ")
    Call Close_Pttel("frmPttel")
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''-- Գործողությունների դիտում թղթապանակից հեռացնել Ավանդի մարում գործողությունը --''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Գործողությունների դիտում թղթապանակից հեռացնել Ավանդի մարում գործողությունը --", "", pmNormal, DivideColor   

    'Մուտք գործել "Ենթահամակրգեր(ՀԾ)"
    Call ChangeWorkspace(c_Subsystems) 
    
    Call LetterOfCredit_Filter_Fill(FolderName, 1, "A-001087")
    Call Delete_Actions("010907","010907",True,"12",c_OpersView)
     
    Call Close_Pttel("frmPttel")

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''-- Կարգավորումներում փոփոխել "Միայն կանխիկ ներգրավված" նշիչը դրած վիճակով --''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Կարգավորումներում փոփոխել (Միայն կանխիկ ներգրավված) նշիչը դրած վիճակով --" ,,, DivideColor    
        
    Call wTreeView.DblClickItem("|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|²¹ÙÇÝÇëïñ³ïÇí Ù³ë|Î³ñ·³íáñáõÙÝ»ñ ¨ ¹ñáõÛÃÝ»ñ|²ÝÏ³ÝËÇÏ ·áñÍ³ñùÝ»ñÇó Ï³ÝËÇÏÇ Ñ³ßí³éÙ³Ý ¹ñáõÛÃÝ»ñ|²í³Ý¹Ý»ñ (Ý»ñ·ñ³íí³Í)")
    BuiltIn.Delay(2000)
    With wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame").VBObject("DocGrid")
      .Row = 0
      .Col = 15
      .Keys("^A[Del]")
      
      .Row = 0
      .Col = 17
      .Text = "-1"
    End With
    
    Call ClickCmdButton(1, "Î³ï³ñ»É")   
    
    'SQL Ստուգում DOCSG աղուսյակում 
    Call CheckQueryRowCount("DOCSG","fISN","131889730",2)
    Call CheckDB_DOCSG("131889730","GRIDINST","0","ONLYCASHATTRPART","1",1)
    Call CheckDB_DOCSG("131889730","GRIDINST","0","ONLYPER","0",1)
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''-- Պարտքերի մարում (Միայն կանխիկ ներգրավված-ով) --'''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Պարտքերի մարում (Միայն կանխիկ ներգրավված-ով) --", "", pmNormal, DivideColor   

    Call LetterOfCredit_Filter_Fill(FolderName, 1, "A-001087")
    Call Debt_Repayment(fBASE,"010907", "17750","","2",Acc.Account,"",2)    
    Call Close_Pttel("frmPttel")  
    
    Log.Message fBASE,,,SqlDivideColor  
    
    'SQL Ստուգում DOCS աղուսյակում
    fBODY = "  CODE:A-001087  DATE:20070901  CORROFFPER:0  AMDSUMDBT:20043.1  SUMAGR:17750  NOTAXEDSUM:20043.1  SUMPER:2547.9  ALLTAX:254.8  TAXPER:254.8  SUMMA:20297.9  CASHORNO:2  ISPUSA:0  ACCCORR:"&Acc.Account&"  COMMENT:²í³Ý¹Ç å³ñïù»ñÇ í×³ñáõÙ  ACSBRANCH:00  ACSDEPART:1  ACSTYPE:D10  USERID:  77  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",fBASE,1)
    Call CheckDB_DOCS(fBASE,"D1DSDebt","5",fBODY,1)
        
    'SQL Ստուգում HI աղուսյակում  
    Call CheckQueryRowCount("HI","fBASE",fBASE,6)
    Call Check_HI_CE_accounting ("20070901",fBASE, "01", "241020717","17750.00", "000", "17750.00", "MSC", "D")
    Call Check_HI_CE_accounting ("20070901",fBASE, "01", Acc.Isn,"17750.00", "000", "17750.00", "MSC", "C")
    
    'SQL Ստուգում HI2 աղուսյակում 
    Call CheckQueryRowCount("HI2","fBASE",fBASE,1)
    Set dbHI2 = New_DB_HI2()
    With dbHI2
        .fDATE = "2007-09-01"
        .fTYPE = "10"
        .fOBJECT = "737994605"
        .fGLACC = "1559631"
        .fSUM = "0.00"
        .fCUR = "000"
        .fCURSUM = "2293.10"
        .fOP = "MSC"
        .fBASE = fBASE
        .fDBCR = "D"
    End With
    Call CheckDB_HI2(dbHI2,1)
    
    'SQL Ստուգում HIR աղուսյակում
    Call CheckQueryRowCount("HIR","fBASE",fBASE,6)
    Call Check_HIR("2007-09-01", "R1", "53912814", "000", "17750.00", "DBT", "C")
    Call Check_HIR("2007-09-01", "R2", "53912814", "000", "254.80", "TXD", "C")
    Call Check_HIR("2007-09-01", "R¸", "53912814", "000", "2293.10", "DBT", "C")
    Call Check_HIR("2007-09-01", "R¸", "53912814", "000", "254.80", "TXD", "C")
    Call Check_HIR("2007-09-01", "RÄ", "53912814", "000", "17750.00", "DBT", "C")
    
    'SQL Ստուգում HIREST2 աղուսյակում
    Call CheckDB_HIREST2("10","737994605","1559631","0.00","000","102293.10", 1)
    
    'SQL Ստուգում HIREST  աղուսյակում  
    Call CheckDB_HIREST("01", Acc.Isn,"-120043.10","000","-120043.10",1)  
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''-- Կանխիկ միջոցների հաշվառում թղթապանակում ստուգել Մնացորդը --''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Կանխիկ միջոցների հաշվառում թղթապանակում ստուգել Մնացորդը --" ,,, DivideColor 
    
    'Մուտք Գլխավոր հաշվապահի ԱՇՏ
    Call ChangeWorkspace(c_ChiefAcc)
    
    Set CashAccountingFilter = New_CashAccounting()
    With CashAccountingFilter
      .ClientCode = "00000678"
      .Curr = "000"
    End With
    Call GoTo_CashAccounting(CashAccountingFilter)     
    
    Call CompareFieldValue("frmPttel", "FKEY", "00000678")
    Call CompareFieldValue("frmPttel", "FCUR", "000")   
    Call CompareFieldValue("frmPttel", "SUM", "102,293.10")
    Call CompareFieldValue("frmPttel", "FCOM", "KERAMIKA Ê³ã³ïãÛ³Ý ìÉ³¹»Ý ì³ÝÇÏÇ")
    Call Close_Pttel("frmPttel")
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''-- Հեռացնել տրված Կարգավորումները --''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Հեռացնել տրված Կարգավորումներները --" ,,, DivideColor
        
    'Մուտք գործել "Ենթահամակրգեր(ՀԾ)"
    Call ChangeWorkspace(c_Subsystems) 
    
    Call wTreeView.DblClickItem("|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|²¹ÙÇÝÇëïñ³ïÇí Ù³ë|Î³ñ·³íáñáõÙÝ»ñ ¨ ¹ñáõÛÃÝ»ñ|²ÝÏ³ÝËÇÏ ·áñÍ³ñùÝ»ñÇó Ï³ÝËÇÏÇ Ñ³ßí³éÙ³Ý ¹ñáõÛÃÝ»ñ|²í³Ý¹Ý»ñ (Ý»ñ·ñ³íí³Í)")
    BuiltIn.Delay(2000)
    With wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame").VBObject("DocGrid") 
      .Row = 0
      .Col = 17
      .Text = "0"
      .keys("[Enter]")
    End With
    Call ClickCmdButton(1, "Î³ï³ñ»É")   
    
    'SQL Ստուգում DOCSG աղուսյակում 
    Call CheckQueryRowCount("DOCSG","fISN","131889730",0)
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''-- Գործողությունների դիտում թղթապանակից հեռացնել Ավանդի մարում գործողությունը --''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Գործողությունների դիտում թղթապանակից հեռացնել Ավանդի մարում գործողությունը --", "", pmNormal, DivideColor   

    Call LetterOfCredit_Filter_Fill(FolderName, 1, "A-001087")
    Call Delete_Actions("010907","010907",True,"12",c_OpersView)
     
    Call Close_Pttel("frmPttel")
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''-"Հաշվառված վճարային փաստաթղթերից" հեռացնել "Կանխիք մուտք" գործողությունները-'''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "--Հեռացնել Կանխիք մուտք գործողություննը --",,,DivideColor     
    
    Call ChangeWorkspace(c_ChiefAcc)
    wTreeView.DblClickItem("|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|Ð³ßí³éí³Í í×³ñ³ÛÇÝ ÷³ëï³ÃÕÃ»ñ")
    'Լրացնել "Ամսաթիվ" դաշտը
    Call Rekvizit_Fill("Dialog",1,"General","PERN", "010120")
    Call Rekvizit_Fill("Dialog",1,"General","PERK", "010121")

    Call ClickCmdButton(2, "Î³ï³ñ»É")
    
    If WaitForPttel("frmPttel") Then
        Call SearchAndDelete("frmPttel", 1, "Î³ÝËÇÏ Ùáõïù", "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ") 
        BuiltIn.Delay(2000)
        Call Close_Pttel("frmPttel")
     Else
        Log.Error "Can Not Open Հաշվառված վճարային փաստաթղթեր Window",,,ErrorColor      
     End If     
     
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''-- "Հաշիվներ" թղթապանակից հեռացնել ստաղծված հաշիվը --''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Հաշիվներ թղթապանակից հեռացնել ստաղծված հաշիվը --",,,DivideColor     
    
    Call wTreeView.DblClickItem("|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|Ð³ßÇíÝ»ñ") 
    BuiltIn.Delay(1000)
    'Կանխիկ հաշվառման և Բացման ամսաթիվ դաշտի լրացում
    Call Rekvizit_Fill("Dialog", 1, "CheckBox", "CASHAC", 1)
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    BuiltIn.Delay(2000)
    
    If WaitForPttel("frmPttel") Then
        Call SearchAndDelete("frmPttel", 1, Acc.Account, "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ") 
        BuiltIn.Delay(2000)
        Call Close_Pttel("frmPttel")
     Else
        Log.Error "Can Not Open Հաշիվներ Window",,,ErrorColor      
     End If   
     
    Call Close_AsBank()  
End Sub  

