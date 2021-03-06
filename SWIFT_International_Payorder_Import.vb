Option Explicit
'USEUNIT International_PayOrder_Receive_Confirmphases_Library
'USEUNIT International_PayOrder_ConfirmPhases_Library
'USEUNIT PayOrder_Receive_ConfirmPhases_Library
'USEUNIT SWIFT_International_Payorder_Library
'USEUNIT Payment_Order_ConfirmPhases_Library
'USEUNIT Payment_Except_Library
'USEUNIT Subsystems_SQL_Library
'USEUNIT Online_PaySys_Library
'USEUNIT Akreditiv_Library
'USEUNIT Library_Common
'USEUNIT Constants

'Test case Id 166763

Sub SWIFT_Internatioanal_Payorder_Import_Test()

    Dim max,min,rand, startDATE, fDATE,DocNum,cashOutN
    Dim fileFrom,fileTo,what,fWith,isExists,fBASE
    Dim queryString,sql_Value, colNum,sql_isEqual,result,fOBJECT
    
    max=100
    min=999
    Randomize
    rand = Int((max-min+1)*Rnd+min)
    fileFrom = Project.Path & "Stores\SWIFTtest\Expected\IA000385.RJE"
    fileTo = Project.Path & "Stores\SWIFTtest\Import\IA000387.RJE"
    what = "UBSWCHZHXXXX901"
    fWith = "UBSWCHZHXXXX" & rand
    DocNum = "951394"
    startDATE = "20010101"
    fDATE = "20250101"
    Log.Message(fWith)
    
    'Վերարտագրում է նախօրոք տրված ֆայլը մեկ այլ ֆայլի մեջ` կատարելով փոփոխություն
    Call Read_Write_File_With_replace(fileFrom,fileTo,what,fWith)
    
   'Test StartUp start
    Call Initialize_AsBank("bank", startDATE, fDATE)
    BuiltIn.Delay(1000)
    Call ChangeWorkspace(c_Admin)
    Call Create_Connection()
    Call SetParameter("SWIN", Project.Path& "Stores\SWIFTtest\Actual\")
    Call SetParameter("SWFAIN", Project.Path& "Stores\SWIFTtest\Actual\FileAct\")
    Call SetParameter("SWFAOUT", Project.Path& "Stores\SWIFTtest\Import\FileAct\")
    Call SetParameter("SWOUT", Project.Path& "Stores\SWIFTtest\Import\")
    Call SetParameter("SWTMPDIR", "\\host2\Sys\Testing\SWIFT\tmp\")
      
    'Դնում է ուղարկել SWIFT նշիչը
    Call Change_User_Permission_SWIFT()
    Call Login("ARMSOFT")
   
    Call ChangeWorkspace(c_SWIFT)
    Call wTreeView.DblClickItem("|S.W.I.F.T. ²Þî                  |Ð³Õáñ¹³·ñáõÃÛáõÝÝ»ñÇ ÁÝ¹áõÝáõÙ|ÀÝ¹áõÝ»É S.W.I.F.T. Ñ³Ù³Ï³ñ·Çó")
    Call ClickCmdButton(5, "OK")
      
    Call ChangeWorkspace(c_ExternalTransfers)
    Call wTreeView.DblClickItem("|²ñï³ùÇÝ ÷áË³ÝóáõÙÝ»ñÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|êï³óí³Í Ñ³ÝÓÝ³ñ³ñ³·ñ»ñ|Ð³ßí³éÙ³Ý »ÝÃ³Ï³")
    Call Rekvizit_Fill("Dialog",1,"General", "PERN",aqDateTime.Today)
    Call Rekvizit_Fill("Dialog",1,"General", "PERK",aqDateTime.Today)
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    
    'Ստուգում է փաստաթղթի առկայությունը
    BuiltIn.Delay(4000) 
    If Not Trim(wMainForm.Window("MDIClient", "", 1).vbObject("frmPttel").vbObject("tdbgView").Columns.Item(2).Text) = Trim(DocNum) Then
        Log.Error("The  document  does't exist")
    End If          
    'Վերցնում է հանձնարարգրի ISN-ը
    Call wMainForm.MainMenu.Click(c_AllActions)    
    Call wMainForm.PopupMenu.Click(c_View)
    fBASE = Sys.Process("Asbank").vbObject("MainForm").Window("MDIClient", "", 1).vbObject("frmASDocForm").DocFormCommon.Doc.isn
    Call ClickCmdButton(1, "OK")
    Log.Message(fBASE)
   
       'Կատարում ենք SQL ստուգում
       queryString = "select fTRANS from  HI where fBASE = '" & fBASE & "'  and fDBCR = 'C'"
       sql_Value = 0
       colNum = 0
       sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
       If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
       End If
       
       queryString = "select fTRANS from  HI where fBASE = '" & fBASE & "'  and fDBCR = 'D'"
       sql_Value = 0
       colNum = 0
       sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
       If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
       End If
  
    'Կատարում է "Հաշվառել" գործողությունը
    BuiltIn.Delay(2000)
    Call wMainForm.MainMenu.Click(c_AllActions)    
    Call wMainForm.PopupMenu.Click(c_DoTrans)
    Call ClickCmdButton(1, "Î³ï³ñ»É")
    BuiltIn.Delay(2000)
    Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").Close()
    Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("FrmSpr").Close()
    
   
        'Կատարում ենք SQL ստուգում
       queryString = "select fTRANS from  HI where fBASE = '" & fBASE & "'  and fDBCR = 'C'"
       sql_Value = 0
       colNum = 0
       sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
       If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
       End If
       
       queryString = "select fTRANS from  HI where fBASE = '" & fBASE & "'  and fDBCR = 'D'"
       sql_Value = 0
       colNum = 0
       sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
       If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
       End If
       
       queryString = "select fCORBANK from SW_MESSAGES where fUNIQUEID like '%" & fWith & "%'"
       sql_Value = "UBSWCHZHXXX"
       colNum = 0
       sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
       If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
       End If
       
       queryString = "select fSTATE from PAYMENTS where fISN = " & fBASE 
       sql_Value = 7
       colNum = 0
       sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
       If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
       End If
       
    'Ստուգումէ պայմանագրի առկայությունը
    Call wTreeView.DblClickItem("|²ñï³ùÇÝ ÷áË³ÝóáõÙÝ»ñÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|Ð³ßí³éí³Í ëï³óí³Í ÷áË³ÝóáõÙÝ»ñ")
    Call Rekvizit_Fill("Dialog",1,"General", "PERN",aqDateTime.Today)
    Call Rekvizit_Fill("Dialog",1,"General", "PERK",aqDateTime.Today)
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    
    BuiltIn.Delay(2500) 
    If Not Trim(wMainForm.Window("MDIClient", "", 1).vbObject("frmPttel").vbObject("tdbgView").Columns.Item(2).Text) = Trim(DocNum) Then
        Log.Error("The  document  does't exist")
    End If          
    Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").Close()
   
    Call Login("ARMSOFT")
    
    Randomize
    rand = Int((max-min+1)*Rnd+min)
    what = "UBSWCHZHXXXX901"
    fWith = "UBSWCHZHXXXX" & rand
    Log.Message(fWith)
    
    'Վերարտագրում է նախօրոք տրված ֆայլը մեկ այլ ֆայլի մեջ` կատարելով փոփոխություն
    Call Read_Write_File_With_replace(fileFrom,fileTo,what,fWith)
    BuiltIn.Delay(1000) 
    Call ChangeWorkspace(c_SWIFT)
    Call wTreeView.DblClickItem("|S.W.I.F.T. ²Þî                  |Ð³Õáñ¹³·ñáõÃÛáõÝÝ»ñÇ ÁÝ¹áõÝáõÙ|ÀÝ¹áõÝ»É S.W.I.F.T. Ñ³Ù³Ï³ñ·Çó")
    Call ClickCmdButton(5, "OK")
      
    Call ChangeWorkspace(c_ExternalTransfers)
    Call wTreeView.DblClickItem("|²ñï³ùÇÝ ÷áË³ÝóáõÙÝ»ñÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|êï³óí³Í Ñ³ÝÓÝ³ñ³ñ³·ñ»ñ|Ð³ßí³éÙ³Ý »ÝÃ³Ï³")
    Call Rekvizit_Fill("Dialog",1,"General", "PERN",aqDateTime.Today)
    Call Rekvizit_Fill("Dialog",1,"General", "PERK",aqDateTime.Today)
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    BuiltIn.Delay(2000)
       
        'Կատարում ենք SQL ստուգում
        queryString = "select fTRANS from  HI where fBASE = '" & fBASE & "'  and fDBCR = 'C'"
        sql_Value = 0
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If
       
        queryString = "select fTRANS from  HI where fBASE = '" & fBASE & "'  and fDBCR = 'D'"
        sql_Value = 0
        colNum = 0
        sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
        If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
        End If

    If Not Trim(wMainForm.Window("MDIClient", "", 1).vbObject("frmPttel").vbObject("tdbgView").Columns.Item(2).Text) = Trim(DocNum) Then
        Log.Error("The  document  does't exist")
    End If
   
    'Խմբագրում է հանձնարարգիրը
    Call wMainForm.MainMenu.Click(c_AllActions)    
    Call wMainForm.PopupMenu.Click(c_ToEdit)
    'Լրացնում է "Տարանցիկ հաշիվ" դաշտը
    Call Rekvizit_Fill("Document",2,"General","TCORRACC","000548101  ")
    Call ClickCmdButton(1, "Î³ï³ñ»É")
    
    'Կատարում է "Հաշվառել" գործողությունը
    BuiltIn.Delay(2000)
    Call wMainForm.MainMenu.Click(c_AllActions)    
    Call wMainForm.PopupMenu.Click(c_DoTrans)
    Call ClickCmdButton(1, "Î³ï³ñ»É")
    BuiltIn.Delay(2000)
    Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").Close()
    Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("FrmSpr").Close()
      
      'Կատարում ենք SQL ստուգում
       queryString = "select fTRANS from  HI where fBASE = '" & fBASE & "'  and fDBCR = 'C'"
       sql_Value = 0
       colNum = 0
       sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
       If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
       End If
       
       queryString = "select fTRANS from  HI where fBASE = '" & fBASE & "'  and fDBCR = 'D'"
       sql_Value = 0
       colNum = 0
       sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
       If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
       End If
       
       queryString = "select fCORBANK from SW_MESSAGES where fUNIQUEID like '%" & fWith & "%'"
       sql_Value = "UBSWCHZHXXX"
       colNum = 0
       sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
       If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
       End If
       
       queryString = "select fSTATE from PAYMENTS where fISN = " & fBASE 
       sql_Value = 7
       colNum = 0
       sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
       If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
       End If
       
    Call wTreeView.DblClickItem("|²ñï³ùÇÝ ÷áË³ÝóáõÙÝ»ñÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|êï³óí³Í Ñ³ÝÓÝ³ñ³ñ³·ñ»ñ|î³ñ³ÝóÇÏ")
    Call Rekvizit_Fill("Dialog",1,"General", "PERN",aqDateTime.Today)
    Call Rekvizit_Fill("Dialog",1,"General", "PERK",aqDateTime.Today)
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    BuiltIn.Delay(2000)
    'Կատարում է "Մարել" գործողությունը
    Call wMainForm.MainMenu.Click(c_AllActions)    
    Call wMainForm.PopupMenu.Click(c_ToFade)
    Call ClickCmdButton(5, "²Ûá")
    BuiltIn.Delay(2000)
    Call ClickCmdButton(1, "Î³ï³ñ»É")
    BuiltIn.Delay(2000)
    Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("FrmSpr").Close()
    Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").Close()
     
      'Կատարում ենք SQL ստուգում        
       queryString = "select fSTATE from PAYMENTS where fISN = " & fBASE 
       sql_Value = 7
       colNum = 0
       sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
       If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
       End If
       
    BuiltIn.Delay(1000) 
    Call wTreeView.DblClickItem("|²ñï³ùÇÝ ÷áË³ÝóáõÙÝ»ñÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|Ð³ßí³éí³Í ëï³óí³Í ÷áË³ÝóáõÙÝ»ñ")
    Call Rekvizit_Fill("Dialog",1,"General", "PERN",aqDateTime.Today)
    Call Rekvizit_Fill("Dialog",1,"General", "PERK",aqDateTime.Today)
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    
    BuiltIn.Delay(2000) 
    If Not Trim(wMainForm.Window("MDIClient", "", 1).vbObject("frmPttel").vbObject("tdbgView").Columns.Item(2).Text) = Trim(DocNum) Then
        Log.Error("The  document  does't exist")
    End If          
    
    
    Call Login("ARMSOFT")
    
    Randomize
    rand = Int((max-min+1)*Rnd+min)
    what = "UBSWCHZHXXXX901"
    fWith = "UBSWCHZHXXXX" & rand
    Log.Message(fWith)
    
    'Վերարտագրում է նախօրոք տրված ֆայլը մեկ այլ ֆայլի մեջ` կատարելով փոփոխություն
    Call Read_Write_File_With_replace(fileFrom,fileTo,what,fWith)
    
    Call ChangeWorkspace(c_SWIFT)
    Call wTreeView.DblClickItem("|S.W.I.F.T. ²Þî                  |Ð³Õáñ¹³·ñáõÃÛáõÝÝ»ñÇ ÁÝ¹áõÝáõÙ|ÀÝ¹áõÝ»É S.W.I.F.T. Ñ³Ù³Ï³ñ·Çó")
    Sys.Process("Asbank").VBObject("frmAsMsgBox").VBObject("cmdButton").Click()
      
    Call ChangeWorkspace(c_ExternalTransfers)
    Call wTreeView.DblClickItem("|²ñï³ùÇÝ ÷áË³ÝóáõÙÝ»ñÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|êï³óí³Í Ñ³ÝÓÝ³ñ³ñ³·ñ»ñ|Ð³ßí³éÙ³Ý »ÝÃ³Ï³")
    Call Rekvizit_Fill("Dialog",1,"General", "PERN",aqDateTime.Today)
    Call Rekvizit_Fill("Dialog",1,"General", "PERK",aqDateTime.Today)
    Sys.Process("Asbank").VBObject("frmAsUstPar").VBObject("CmdOK").Click()
    BuiltIn.Delay(2000)
       
        'Կատարում ենք SQL ստուգում
       queryString = "select fTRANS from  HI where fBASE = '" & fBASE & "'  and fDBCR = 'C'"
       sql_Value = 0
       colNum = 0
       sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
       If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
       End If
       
       queryString = "select fTRANS from  HI where fBASE = '" & fBASE & "'  and fDBCR = 'D'"
       sql_Value = 0
       colNum = 0
       sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
       If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
       End If

    If Not Trim(wMainForm.Window("MDIClient", "", 1).vbObject("frmPttel").vbObject("tdbgView").Columns.Item(2).Text) = Trim(DocNum) Then
        Log.Error("The  document  does't exist")
    End If
   
    'Խմբագրում է հանձնարարգիրը
    Call wMainForm.MainMenu.Click(c_AllActions)    
    Call wMainForm.PopupMenu.Click(c_ToEdit)
    'Լրացնում է "Տարանցիկ հաշիվ" դաշտը
    Call Rekvizit_Fill("Document",2,"General","TCORRACC","000548101  ")
    Call ClickCmdButton(1, "Î³ï³ñ»É")
    BuiltIn.Delay(2000)
    
    'Կատարում է "Հաշվառել" գործողությունը
    Call wMainForm.MainMenu.Click(c_AllActions)    
    Call wMainForm.PopupMenu.Click(c_DoTrans)
    Call ClickCmdButton(1, "Î³ï³ñ»É")
    BuiltIn.Delay(2000)
    Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").Close()
    Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("FrmSpr").Close()
      
       'Կատարում ենք SQL ստուգում
       queryString = "select fTRANS from  HI where fBASE = '" & fBASE & "'  and fDBCR = 'C'"
       sql_Value = 0
       colNum = 0
       sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
       If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
       End If
       
       queryString = "select fTRANS from  HI where fBASE = '" & fBASE & "'  and fDBCR = 'D'"
       sql_Value = 0
       colNum = 0
       sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
       If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
       End If
       
       queryString = "select fCORBANK from SW_MESSAGES where fUNIQUEID like '%" & fWith & "%'"
       sql_Value = "UBSWCHZHXXX"
       colNum = 0
       sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
       If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
       End If
       
       queryString = "select fSTATE from PAYMENTS where fISN = " & fBASE 
       sql_Value = 7
       colNum = 0
       sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
       If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
       End If
       
    Call wTreeView.DblClickItem("|²ñï³ùÇÝ ÷áË³ÝóáõÙÝ»ñÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|êï³óí³Í Ñ³ÝÓÝ³ñ³ñ³·ñ»ñ|î³ñ³ÝóÇÏ")
    Call Rekvizit_Fill("Dialog",1,"General", "PERN",aqDateTime.Today)
    Call Rekvizit_Fill("Dialog",1,"General", "PERK",aqDateTime.Today)
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    BuiltIn.Delay(2000)
    Call wMainForm.MainMenu.Click(c_AllActions)    
    Call wMainForm.PopupMenu.Click(c_CashOut)
    BuiltIn.Delay(1000)
    cashOutN = Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmASDocForm").VBObject("TabFrame").VBObject("TextC").Text
    Log.Message(cashOutN)
    Call Rekvizit_Fill("Document",1,"General", "AIM","test")
    Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmASDocForm").VBObject("CmdOk_2").Click()
    Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("FrmSpr").Close()
    Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").Close()
    BuiltIn.Delay(1000)
    
    Call ChangeWorkspace(c_CustomerService)
    Call wTreeView.DblClickItem("|Ð³×³Ëáñ¹Ç ëå³ë³ñÏáõÙ ¨ ¹ñ³Ù³ñÏÕ |²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ")
    Call Rekvizit_Fill("Dialog",1,"General", "PERN",aqDateTime.Today)
    Call Rekvizit_Fill("Dialog",1,"General", "PERK",aqDateTime.Today)
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    
    BuiltIn.Delay(5000) 
    If Not Trim(wMainForm.Window("MDIClient", "", 1).vbObject("frmPttel").vbObject("tdbgView").Columns.Item(2).Text) = Trim(cashOutN) Then
        Log.Error("The  document  does't exist")
    End If
    
    'Վավերացնում է փաստաթուղթը
    Call Validate_Doc()
    Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").Close()
    
    Call wTreeView.DblClickItem("|Ð³×³Ëáñ¹Ç ëå³ë³ñÏáõÙ ¨ ¹ñ³Ù³ñÏÕ |Ð³ßí³éí³Í í×³ñ³ÛÇÝ ÷³ëï³ÃÕÃ»ñ")
    Call Rekvizit_Fill("Dialog",1,"General", "PERN",aqDateTime.Today)
    Call Rekvizit_Fill("Dialog",1,"General", "PERK",aqDateTime.Today)
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    BuiltIn.Delay(2000)

    Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").VBObject("tdbgView").MoveLast
    Do Until Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").VBObject("tdbgView").EOF
    BuiltIn.Delay(2000)
       If Not Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").VBObject("tdbgView").ApproxCount = 0  then
             BuiltIn.Delay(3000)
             Sys.Process("Asbank").Refresh
             'Կատարում է ջնջել գործողությունը
             Call wMainForm.MainMenu.Click(c_AllActions)
             Call wMainForm.PopupMenu.Click(c_Delete)
             If Sys.Process("Asbank").WaitVBObject("frmAsMsgBox", 1500).Exists Then
                Call ClickCmdButton(5, "Î³ï³ñ»É")
                Call ClickCmdButton(3, "²Ûá")
             Else
                Call ClickCmdButton(3, "²Ûá")
                If Sys.Process("Asbank").WaitVBObject("frmAsMsgBox", 1500).Exists Then
                   Call ClickCmdButton(5, "Î³ï³ñ»É")
                   Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").VBObject("tdbgView").MoveNext
                End If
             End If       
        Else
          Exit Do
       End If
    Loop
    BuiltIn.Delay(2000)
    Call wMainForm.MainMenu.Click(c_Windows)
    Call wMainForm.PopupMenu.Click(c_ClCurrWindow)
   
    Call Close_AsBank()
      
End Sub