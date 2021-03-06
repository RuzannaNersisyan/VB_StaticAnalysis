Option Explicit
'USEUNIT Payment_Order_ConfirmPhases_Library
'USEUNIT CashInput_Confirmphases_Library
'USEUNIT Subsystems_SQL_Library
'USEUNIT Online_PaySys_Library
'USEUNIT Online_PaySys_Library
'USEUNIT Akreditiv_Library
'USEUNIT Library_Common
'USEUNIT DAHK_Libraray
'USEUNIT Repo_Library
'USEUNIT Constants
'USEUNIT Mortgage_Library
'USEUNIT Card_Library

'Test Case ID 166614
Sub DAHK_Recalculate_Test()

      Dim startDATE, fDATE, date, SqlQuery, SqlQuery2, clientCode, blockSum, debt, acc, summ
      Dim blockID, currDate, docType, messType, docISN, client
      Dim summa, docN, confAcc, colNum, opType, Doc_ISN, Name
      Dim queryString, sql_Value, sql_isEqual, docExist, my_vbobj
    
'-------------------------- Արգելանք -----------------------------------------------------------------------------------------------------------------   
      SqlQuery = "SET IDENTITY_INSERT DAHKATTACH ON" 
      SqlQuery2 = "Insert into DAHKATTACH (fID,	fIMPDATE,	fMESSAGEID,	fDECISIONNUM,	fDECISIONDATE,	fDECISIONPLACE,	fDECISIONOWNER,	fINQUESTNUMBER,	fINQUESTID,	fINQUESTDATE,	fBRANCH,	fBRANCHSUB,	fDEBTORID" _
      			  &  "	, fDEBTORNAME,	fDEBTORPASSPORT, fDEBTORADDRESS,	fDEBTORTYPE,	fISSUM,	fBBLOCKOTHER,	fBBLOCKSUM1,	fBBLOCKCUR1,	fBBLOCKSUM2,	fBBLOCKCUR2,	fBBLOCKSUM3" _
      			  &  "	, fBBLOCKCUR3,	fBBLOCKSUM4,	fBBLOCKCUR4,	fBBLOCKSUM5,	fBBLOCKCUR5,	fBBLOCKSUM6,	fBBLOCKCUR6,	fBBLOCKSUM7,	fBBLOCKCUR7,	fORDERTEXT,	fCOURT,	fCLICODE" _
      			  &  "	,	fBLCODE,	fRESPONSEISN,	fEMPLOYERACC1,	fEMPLOYERACC2,	fEMPLOYERACC3,	fEMPLOYERACC4,	fEMPLOYERACC5,	fDUPLICATE,	fSSN,	fPROCESSED,	fBBLOCKPERCENT,	fRESPONSENUMBER" _
      			  &  "	,	fBBLOCKEDACCOUNTPERCENT,	fBBLOCKEDACCOUNT,	fBLCODEUNVER)" _
           &  "   Values ('15','2018-01-22 00:00:00',	'º01000239589',	'à00166-00001/10',	'2018-11-16 00:00:00',	'ù. ºñ¨³Ý',	'²í³· Ñ³ñÏ³¹Çñ Ï³ï³ñáÕ ³ñ¹³ñ³¹³ïáõÃÛ³Ý ³í. É»Ûï»Ý³Ýï Î³ñ»Ý Ê³Ý½³¹Û³Ý',	'01/03-02753/10'" _
           &  "  ,'00023198',	'2018-11-16 00:00:00',	'Ø³É³ÃÇ³-ê»µ³ëïÇ³ µ³ÅÇÝ',	NULL,	'1111111112',	'êáÏáÉ-¶ñáõå êäÀ',	'01826746','ù.ºñ¨³Ý, Ø³É³ÃÇ³-ê»µ³ëïÇ³, Þñç³Ý³ÛÇÝ 2/4-23',	0,	1,	NULL" _
           &  "  ,	'102253200.00',	'AMD',	'0.00',	NULL,	'0.00',	NULL,	'97384000.00',	'AMD',	'0.00',	NULL,	'0.00',	NULL,	'4869200.00',	'AMD',	'§êáÏáÉ ¶ñáõå¦ êäÀ-Çó Ñû·áõï §Ð³ñ³íÏáíÏ³ëÛ³Ý »ñÏ³ÃáõÕÇ¦ ö´À-Ç µéÝ³·³ÝÓ»É 94.000.000 ÐÐ ¹ñ³Ù'" _
           &  "  ,	'Ø³É³ÃÇ³ - ê»µ³ëïÇ³ Ñ³Ù³ÛÝùÇ ÁÝ¹Ñ³Ýáõñ Çñ³í³ëáõÃÛ³Ý',	'00000024',' ', ' ',NULL,	NULL,	NULL,	NULL,	NULL,	NULL,	NULL,	1,	NULL,	NULL,	NULL,	NULL,	NULL)" _

      startDATE = "20100101"
      fDATE = "20250101"
      date = "220118"   
      messType = "01" 
      blockID = "º01000239589" 
    
      'Test StartUp start
      Call Initialize_AsBank("bank", startDATE, fDATE)

      Call Create_Connection()
      Login("ARMSOFT")
      'Ներմուծել "Արգելանք" տեսակի հաղորդագորւթյունը
      Call Execute_SLQ_Query(SqlQuery)
      Call Execute_SLQ_Query(SqlQuery2)   
   
      'Կատարում ենք SQL ստուգում
      queryString = "select dbo.asfb_GetRemHI('01','443871031','2020-01-01 00:00:00' ) as Acc"
      sql_Value = -10000000000.00
      colNum = 0 
      sql_isEqual = CheckDB_Value(queryString, sql_Value, colNum)
      If Not sql_isEqual Then
         Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
      End If     
        
      ' Մուտք ադմինիստրատորի ԱՇՏ4.0
      Call ChangeWorkspace(c_Admin40)
      Call wTreeView.DblClickItem("|²¹ÙÇÝÇëïñ³ïáñÇ ²Þî 4.0|î»Õ»Ï³ïáõÝ»ñ|Ð³×³Ëáñ¹Ý»ñ")
       
      ' Խմբագրել հաճախորդի տվյալները
      Set client = New_Clients()  
      client.ClientsCode = "00000024"
      
      Call Fill_Clients(client)  
      Call WaitForExecutionProgress()
      Call CheckPttel_RowCount("frmPttel", 1)
       
      ' Կատարել բոլոր գործողությունները
      Call wMainForm.MainMenu.Click(c_AllActions)
      ' Վավերացնել գործողության կատարում
      Call wMainForm.PopupMenu.Click(c_ToEdit)

      If wMDIClient.WaitVBObject("frmASDocForm",10000).Exists Then
           ' Անձը հաստ. փաստաթ. կոդ դաշտի լրացում
            Call Rekvizit_Fill("Document", 1, "General", "PASCODE", "AR21354218")
            ' Անձը հաստ. փաստաթ. կոդ տիպ դաշտի լրացում
            Call Rekvizit_Fill("Document", 1, "General", "PASTYPE", "13")        
            ' Տրված դաշտի լրացում
            Call Rekvizit_Fill("Document", 1, "General", "PASBY", "012")
            ' Տրված ժամանակ դաշտի լրացում
            Call Rekvizit_Fill("Document", 1, "General", "DATEPASS", "01012020")
            ' Վավեր է մինչև դաշտի լրացում
            Call Rekvizit_Fill("Document", 1, "General", "DATEEXPIRE", "01012030")  
            
            ' Սեռ դաշտի լրացում
            Call Rekvizit_Fill("Document", 2, "General", "GENDER", "M")
            ' Ծննդյան ամսաթիվ դաշտի լրացում
            Call Rekvizit_Fill("Document", 2, "General", "DATEBIRTH", "01011995")
            ' Ծննդավայր դաշտի լրացում
            Call Rekvizit_Fill("Document", 2, "General", "BIRTHPLACE", "AM")
            
            ' Բնակավայր (գրանցման հասցե) դաշտի լրացում
            Call Rekvizit_Fill("Document", 3, "General", "COMMUNITY", "010010635")
            ' Քաղաք (գրանցման հասցե) դաշտի լրացում
            Call Rekvizit_Fill("Document", 3, "General", "CITY",  "^A[Del]" & "ºñ¨³Ý")      
            ' Երկիր (փաստացի հասցե) դաշտի լրացում
            Call Rekvizit_Fill("Document", 3, "General", "COUNTRY2", "AM")
            ' Բնակավայր (փաստացի հասցե) դաշտի լրացում
            Call Rekvizit_Fill("Document", 3, "General", "COMMUNITY2", "010010635")
            ' Քաղաք (փաստացի հասցե) դաշտի լրացում
            Call Rekvizit_Fill("Document", 3, "General", "CITY2",  "^A[Del]" & "ºñ¨³Ý")
            ' Փողոց (փաստացի հասցե) դաշտի լրացում
            Call Rekvizit_Fill("Document", 3, "General", "STREET2", "ºñ¨³Ý, ´³ÕÝÇù N 4")
            Call ClickCmdButton(1, "Î³ï³ñ»É")
            
            If  MessageExists(2, "ÐÇß»±É ÑÇÝ ³ÝÓÁ Ñ³ëï³ïáÕ ÷³ëï³ÃÕÃÇ Ñ³Ù³ñÁ") Then
                Call ClickCmdButton(5, "àã")  
            End If
            Call Close_Pttel("frmPttel")
      End If
      
    BuiltIn.Delay(2000) 
    Call ChangeWorkspace(c_DAHK)
    
    'Մուտք գործել "Ընդունված հաղորդագրություններ" թղթապանակ
    If Not Enter_Recieved_Messages(date,date,messType,blockID) Then
            Log.Error("Փաստաթուղթը չի գտնվել")
            Exit Sub
    End If
    
    'Ստուգել հաճախորդի կոդ դաշտը ճիշտ ներմուծված լինի
    colNum =	wMDIClient.VBObject("frmPttel").GetColumnIndex("fCLICODE")
    If Not Trim(wMDIClient.VBObject("frmPttel").VBObject("tdbgView").Columns.Item(colNum).Value) = Trim("00000024") Then
            Log.Error("Հաճախորդի կոդը սխալ է")
            Exit Sub
    End If
    
    'Կատարել Գումարների արգելադրում
    blockSum = "102253200.00"
    debt = "102,252,746.10"
    Call Blocking_Money(blockSum,debt)  
    BuiltIn.Delay(4000)  
    
    'Ստուգել Հաշվի ստորին սահմանը
    acc = "01046803311  ²ñÅ.- 000  îÇå- 01  Ð/Ð³ßÇí- 3022000   ²Ýí³ÝáõÙ-ØÇß³ï ´³ÝÛ³Ý"
    summ = "453.9"
    Call Check_Account_Low_Border(acc,summ)
    
    'Փակել պտտելը
    BuiltIn.Delay(4000)
    Call Close_Pttel("frmPttel")
    BuiltIn.Delay(4000)
    
    Call ChangeWorkspace(c_CustomerService)
    BuiltIn.Delay(2000)
    
    currDate  = aqConvert.DateTimeToFormatStr(aqDateTime.Today(), "%d%m%y")
    Log.Message(currDate)
    
    'Կանխիկ մուտք փաստաթղթի ստեղծում
    confAcc = "01046803311"
    summa = "100000000"
    Name = "Պետրոսյան Պետրոս"
    Call CashInputOutput(currDate, "CashInput", confAcc, summa, docN, Name, docISN)
    wMDIClient.VBObject("FrmSpr").Close()
    
    'Ուղարկել փաստաթուղթը հաստատման
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_SendToVer)
    Sys.Process("Asbank").VBObject("frmAsUstPar").VBObject("CmdOK").Click()
    Call Close_Pttel("frmPttel")
    
    'Անցում կատարել "Հաստատող 1" ԱՇՏ
    Call ChangeWorkspace(c_Verifier1)
    
    'Փաստաթղթի առկայության ստուգում 1-ին հաստատողի մոտ
    docExist = Online_PaySys_Check_Doc_In_Verifier(docN, currDate, currDate)
    If Not docExist Then
          Log.Error(docN & " համարի փաստաթութը առկա չէ 'Հաստաատող 1'-ում")
          Exit Sub
    End If
   
    'Փաստաթղթի վավերացում 1-ին հաստատողի կողմից
    Call PaySys_Verify(True)
    Call Close_Pttel("frmPttel")
      
    Call ChangeWorkspace(c_CustomerService)
    Call Online_PaySys_Go_To_Agr_WorkPapers("|Ð³×³Ëáñ¹Ç ëå³ë³ñÏáõÙ ¨ ¹ñ³Ù³ñÏÕ |²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ", currDate, currDate)
    BuiltIn.Delay(5000)
    Set my_vbobj = wMDIClient.WaitVBObject("frmPttel", delay_middle)
    colNum =	wMDIClient.VBObject("frmPttel").GetColumnIndex("DOCNUM")
    If my_vbobj.Exists Then
        Do Until wMDIClient.vbObject("frmPttel").vbObject("tdbgView").EOF
            If Trim(wMDIClient.vbObject("frmPttel").vbObject("tdbgView").Columns.Item(colNum).text) = docN Then
                Call PaySys_Verify(True)
            Else                                                                         
                Call wMDIClient.vbObject("frmPttel").vbObject("tdbgView").MoveNext
            End If
        Loop
    Else
        Log.Error("Թղթապանակը չի բացվել")
    End If
    Call Close_Pttel("frmPttel")
            
    'Արտաքին փոխանցումներում հանձնարարգրի առկայության ստուգում
    Call ChangeWorkspace(c_CustomerService)
     Call Online_PaySys_Check_Doc_In_Registered_Payment_Documents(docN, currDate, currDate)
    
    'Փակել պտըտելը
    Call Close_Pttel("frmPttel")

    'Անցում կատարել ԴԱՀԿ ԱՇՏ
    Call ChangeWorkspace(c_DAHK)
    BuiltIn.Delay(1500)
    
    'Ստուգել Արգելանքի առկայությունը "Գումարների արգելադրում" թղթապանակում
    If Not Enter_Money_Blockings(currDate,currDate,blockID) Then
            Log.Error("Արգելանքը չի գտնվել Գումարների արգելադրում թղթապանակում")
    End If
    'Ստուգել Սկզբնական գումար սյան արժեքը
    colNum =	wMDIClient.VBObject("frmPttel").GetColumnIndex("fSTARTSUM")
    If Not Trim(wMDIClient.VBObject("frmPttel").VBObject("tdbgView").Columns.Item(colNum)) = "102,253,200.00" Then 
            Log.Error("Սկզբնական գումարը սխալ է")
    End If
    'Ստուգել Արգելադրվող գումար սյան արժեքը
    colNum =	wMDIClient.VBObject("frmPttel").GetColumnIndex("fBLOCKSUM")
    If Not Trim(wMDIClient.VBObject("frmPttel").VBObject("tdbgView").Columns.Item(colNum)) = "102,253,200.00" Then 
            Log.Error("Արգելադրվող գումարը սխալ է")
    End If
    'Ստուգել Պարտք դաշտի արժեքը
    colNum =	wMDIClient.VBObject("frmPttel").GetColumnIndex("fDEBT")
    If Not Trim(wMDIClient.VBObject("frmPttel").VBObject("tdbgView").Columns.Item(colNum)) = "102,252,746.10" Then 
            Log.Error("Պարտքը սխալ է")
    End If
    
    'Կատարել Վերահաշվարկ գործողություն
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_Recalc)
    'Սեղմել Կատարել կոճակը
    Call ClickCmdButton(5, "OK") 
    
    'Թարմացնել Պատուհանը
    Call wMainForm.VBObject("tbToolBar").Window("ToolbarWindow32", "", 1).ClickItem(10)
    wMainForm.Refresh
    BuiltIn.Delay(2000)
    
    'Ստուգել Պարտք դաշտի արժեքը
    colNum =	wMDIClient.VBObject("frmPttel").GetColumnIndex("fDEBT")
    If Not Trim(wMDIClient.VBObject("frmPttel").VBObject("tdbgView").Columns.Item(colNum)) = "2,252,746.10" Then 
            Log.Error("Պարտքը չի փոխվել")
    End If
    
     'Պարտքի ապաակտիվացում
    Call wMainForm.MainMenu.Click(c_AllActions) 
    Call wMainForm.PopupMenu.Click(c_DebtDeactivation)
    Call ClickCmdButton(5, "²Ûá")
    BuiltIn.Delay(2000)
  
    'Ստուգել որ "Ապաակտիվացված է" սյան արժեքը լինի -1(Նշիչը լինի դրված)
    colNum =	wMDIClient.VBObject("frmPttel").GetColumnIndex("fDEBTDEACT")
    If Not Trim(wMDIClient.VBObject("frmPttel").VBObject("tdbgView").Columns.Item(colNum).Value) = "-1" Then
            Log.Error("Պարտքը չի ապաակտիվացվել")
            Exit Sub
    End If
    
    'Պարտքի ակտիվացում
    Call wMainForm.MainMenu.Click(c_AllActions) 
    Call wMainForm.PopupMenu.Click(c_DebtActivation)
    Call ClickCmdButton(5, "²Ûá")
    
    'Ստուգել որ "Ապաակտիվացված է" սյան արժեքը լինի 0(Նշիչը չլինի դրված)
    If Not Trim(wMDIClient.VBObject("frmPttel").VBObject("tdbgView").Columns.Item(colNum).Value) = "0" Then
            Log.Error("Պարտքը չի ակտիվացվել")
            Exit Sub
    End If    
    Call Close_Pttel("frmPttel")
    
    Call ChangeWorkspace(c_CustomerService)
    
    'Կանխիկ մուտք փաստաթղթի ստեծում
    confAcc = "01046803311"
    summa = "2252746.10"
    Call CashInputOutput(currDate, "CashInput", confAcc, summa, docN, Name, docISN)
    wMDIClient.VBObject("FrmSpr").Close()
    
    'Ուղարկել փաստաթուղթը հաստատման
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_SendToVer)
    Sys.Process("Asbank").VBObject("frmAsUstPar").VBObject("CmdOK").Click()
    Call Close_Pttel("frmPttel")
    BuiltIn.Delay(3000)
    
    'Անցում կատարել "Հաստատաող 1" ԱՇՏ
    Call ChangeWorkspace(c_Verifier1)
    
    'Փաստաթղթի առկայության ստուգում 1-ին հաստատողի մոտ
    docExist = Online_PaySys_Check_Doc_In_Verifier(docN, currDate, currDate)
    If Not docExist Then
          Log.Error(docN & " համարի փաստաթութը առկա չէ 'Հաստաատող 1'-ում")
          Exit Sub
    End If
   
    'Փաստաթղթի վավերացում 1-ին հաստաոտղի կողմից
    Call PaySys_Verify(True)
    Call Close_Pttel("frmPttel")
      
    Call ChangeWorkspace(c_CustomerService)
    Call Online_PaySys_Go_To_Agr_WorkPapers("|Ð³×³Ëáñ¹Ç ëå³ë³ñÏáõÙ ¨ ¹ñ³Ù³ñÏÕ |²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ", currDate, currDate)
    BuiltIn.Delay(5000)
    Set my_vbobj = wMDIClient.WaitVBObject("frmPttel", delay_middle)
    colNum =	wMDIClient.VBObject("frmPttel").GetColumnIndex("DOCNUM")
    If my_vbobj.Exists Then
        Do Until wMDIClient.vbObject("frmPttel").vbObject("tdbgView").EOF
            If Trim(wMDIClient.vbObject("frmPttel").vbObject("tdbgView").Columns.Item(colNum).text) = docN Then
                Call PaySys_Verify(True)
            Else                                                                         
                Call wMDIClient.vbObject("frmPttel").vbObject("tdbgView").MoveNext
            End If
        Loop
    Else
        Log.Error("Թղթապանակը չի բացվել")
    End If
    Call Close_Pttel("frmPttel")
            
    'Արտաքին փոխանցումներում հանձնարարգրի առկայության ստուգում
    Call Online_PaySys_Check_Doc_In_Registered_Payment_Documents(docN, currDate, currDate)
    
    'Փակել պտըտելը
    Call Close_Pttel("frmPttel")

    'Անցում կատարել ԴԱՀԿ ԱՇՏ
    Call ChangeWorkspace(c_DAHK)
    
    'Ստուգել Արգելանքի առկայությունը "Գումարների արգելադրում" թղթապանակում
    If Not Enter_Money_Blockings(currDate,currDate,blockID) Then
            Log.Error("Արգելանքը չի գտնվել Գումարների արգելադրում թղթապանակում")
    End If
    'Ստուգել Սկզբնական գումար սյան արժեքը
    colNum =	wMDIClient.VBObject("frmPttel").GetColumnIndex("fSTARTSUM")
    If Not Trim(wMDIClient.VBObject("frmPttel").VBObject("tdbgView").Columns.Item(colNum)) = "102,253,200.00" Then 
            Log.Error("Սկզբնական գումարը սխալ է")
    End If
    'Ստուգել Արգելադրվող գումար սյան արժեքը
    colNum =	wMDIClient.VBObject("frmPttel").GetColumnIndex("fBLOCKSUM")
    If Not Trim(wMDIClient.VBObject("frmPttel").VBObject("tdbgView").Columns.Item(colNum)) = "102,253,200.00" Then 
            Log.Error("Արգելադրվող գումարը սխալ է")
    End If
    'Ստուգել Պարտք դաշտի արժեքը
    colNum =	wMDIClient.VBObject("frmPttel").GetColumnIndex("fDEBT")
    If Not Trim(wMDIClient.VBObject("frmPttel").VBObject("tdbgView").Columns.Item(colNum)) = "2,252,746.10" Then 
            Log.Error("Պարտքը սխալ է")
    End If
    
    'Կատարել Վերահաշվարկ գործողություն
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_Recalc)
    'Սեղմել Կատարել կոճակը
    Call ClickCmdButton(5, "OK") 
    
    'Թարմացնել Պատուհանը
    Call wMainForm.VBObject("tbToolBar").Window("ToolbarWindow32", "", 1).ClickItem(10)
    wMainForm.Refresh
    BuiltIn.Delay(2000)
    
    'Ստուգել Պարտք դաշտի արժեքը
    colNum =	wMDIClient.VBObject("frmPttel").GetColumnIndex("fDEBT")
    If Not Trim(wMDIClient.VBObject("frmPttel").VBObject("tdbgView").Columns.Item(colNum)) = "0.00" Then 
            Log.Error("Պարտքը չի փոխվել")
    End If
    
    'Ջնջել գումարների արգելադրում փաստաթուղթը
    Call wMainForm.MainMenu.Click(c_Opers & "|" & c_Delete)
    Sys.Process("Asbank").VBObject("frmDeleteDoc").VBObject("YesButton").ClickButton
    Call Close_Pttel("frmPttel")
    
    Call ChangeWorkspace(c_CustomerService)
    'Ջնջել Հաշվառված վճարային փաստաթղթեր թղթապանակի փաստաթղթերը
    opType = "KasPrOrd"
    Call DeletePayingDoc(currDate, opType, Doc_ISN)
    
    'Ջնջել ներմուծված հաղորդագորությունները
    SqlQuery2 = " Delete from DAHKFREEATTACH "_
              & " Delete from DAHKCATCH "_
              & " Delete from DAHKATTACH " 
              
    Call Execute_SLQ_Query(SqlQuery2) 
    
    Call Close_AsBank()   
End Sub