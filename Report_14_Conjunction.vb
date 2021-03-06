'USEUNIT Library_Common
'USEUNIT Library_Contracts
'USEUNIT Library_CheckDB  
'USEUNIT Library_Colour
'USEUNIT Constants
'USEUNIT Card_Library
Option Explicit

'Test Case ID 184538

Dim sDate, eDate, exportOLAP, report14
Dim majorClient, minorClient, conjuction, clients(1)

Sub Report_14_Conjunction_Test()
    Call Test_Inintialize()

    ' Համակարգ մուտք գործել ARMSOFT օգտագործողով
    Log.Message "Համակարգ մուտք գործել ARMSOFT օգտագործողով", "", pmNormal, DivideColor
    Call Test_StartUp()
    
    ' Մուտք գործել OLAP խմբերի տեղեկատու թղթապանակ 
    wTreeView.DblClickItem("|OLAP ³¹ÙÇÝÇëïñ³ïáñÇ ²Þî|OLAP ËÙµ»ñÇ ï»Õ»Ï³ïáõ")
    If wMDIClient.WaitVBObject("frmPttel", 2000).Exists Then
        ' Արտահանել BAZEL CONTRACTS-ը
        Log.Message "Արտահանել BAZEL CONTRACTS-ը", "", pmNormal, DivideColor
        If SearchInPttel("frmPttel", 0, "BAZEL CONTRACTS") Then 
            Call OLAP_Group_Export(exportOLAP)
            ' Փակել OLAP խմբերի տեղեկատու թղթապանակը
            Call Close_Window(wMDIClient, "frmPttel")
        Else
            Log.Error "Can't fined searched row with BAZEL CONTRACTS value.", "", pmNormal, ErrorColor
        End If
    Else 
        Log.Error "Can't open frmPttel widow.", "", pmNormal, ErrorColor
    End If
    
    ' Անցում կատարել Ենթահամակարգեր (§ՀԾ¦) ԱՇՏ
    p1.Terminate()
    Call Initialize_AsBankQA(sDate, eDate)
    Call ChangeWorkspace(c_Subsystems)
    
    ' Մուտք գործել 14 խոշոր փոխառուներ հաշվետվություն
    Call GoTo_Report14(report14)
    
    ' Ստուգել, որ Խոշոր փոխառու հաճախորդը գտնվել է
    Log.Message "Ստուգել, որ Խոշոր փոխառու հաճախորդը գտնվել է", "", pmNormal, DivideColor
    If Find_Word(wMDIClient.VBObject("FrmSpr"), majorClient.ClientsCode) Then
        Log.Message "Major client " & majorClient.ClientsCode & " was found.", "", pmNormal, MessageColor
    Else 
        Log.Error "Major client " & majorClient.ClientsCode & " wasn't found.", "", pmNormal, ErrorColor
    End If
    
    ' Ստուգել, որ կապակցված ակտիվ չունեցող հաճախորդը առկա չէ 14 հաշվետվությունում
    Log.Message "Ստուգել, որ կապակցված ակտիվ չունեցող հաճախորդը առկա չէ 14 հաշվետվությունում", "", pmNormal, DivideColor
    If Not Find_Word(wMDIClient.VBObject("FrmSpr"), minorClient) Then
        Log.Message "Client without actives " & minorClient & " wasn't found.", "", pmNormal, MessageColor
    Else 
        Log.Error "Client without actives " & minorClient & " was found.", "", pmNormal, ErrorColor
    End If
    
    ' Փակել 14 խոշոր փոխառուներ հաշվետվությունը
    Call Close_Window(wMDIClient, "FrmSpr")
    
    ' Անցում կատարել Գլխավոր հաշվապահի ԱՇՏ
    p1.Terminate()
    Call Initialize_AsBankQA(sDate, eDate)
    Call ChangeWorkspace(c_ChiefAcc)  
    
    ' Մուտք գործել Հաճախորդներ թղթապանակ
    Call wTreeView.DblClickItem("|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|Ð³×³Ëáñ¹Ý»ñ")
    Call Fill_Clients(majorClient)
    
    ' Կապակցել ակտիվ չունեցող հաճախորդ խոշոր փոխառուի հետ 
    Log.Message "Կապակցել ակտիվ չունեցող հաճախորդ խոշոր փոխառուի հետ ", "", pmNormal, DivideColor
    Call Cilent_Conjunction(conjuction)
    
    ' Կատարել SQL ստուգում Հաճախորդների կապակցումից հետո
    Log.Message "Կատարել SQL ստուգում Հաճախորդների կապակցումից հետո", "", pmNormal, DivideColor
    Call CheckDB_Conjuction()
    
    ' Փակել Հաճախորդներ թղթապանակը
    Call Close_Window(wMDIClient, "frmPttel")
    
    ' Անցում կատարել Ենթահամակարգեր (§ՀԾ¦) ԱՇՏ
    Call ChangeWorkspace(c_Subsystems)
    
    ' Մուտք գործել 14 խոշոր փոխառուներ հաշվետվություն
    Call GoTo_Report14(report14)
    
    ' Ստուգել, որ Խոշոր փոխառու հաճախորդը գտնվել է
    Log.Message "Ստուգել, որ Խոշոր փոխառու հաճախորդը գտնվել է", "", pmNormal, DivideColor
    If Find_Word(wMDIClient.VBObject("FrmSpr"), majorClient.ClientsCode) Then
        Log.Message "Major client " & majorClient.ClientsCode & " was found.", "", pmNormal, MessageColor
    Else 
        Log.Error "Major client " & majorClient.ClientsCode & " wasn't found.", "", pmNormal, ErrorColor
    End If
    
    ' Ստուգել, որ կապակցված ակտիվ չունեցող հաճախորդը առկա չէ 14 հաշվետվությունում
    Log.Message "Ստուգել, որ կապակցված ակտիվ չունեցող հաճախորդը առկա չէ 14 հաշվետվությունում", "", pmNormal, DivideColor
    If Not Find_Word(wMDIClient.VBObject("FrmSpr"), minorClient) Then
        Log.Message "Client without actives " & minorClient & " wasn't found.", "", pmNormal, MessageColor
    Else 
        Log.Error "Client without actives " & minorClient & " was found.", "", pmNormal, ErrorColor
    End If
    
    ' Փակել 14 խոշոր փոխառուներ հաշվետվությունը
    Call Close_Window(wMDIClient, "FrmSpr")
    
    ' Անցում կատարել Գլխավոր հաշվապահի ԱՇՏ
    Call ChangeWorkspace(c_ChiefAcc) 
    
    ' Մուտք գործել Հաճախորդներ թղթապանակ
    Call wTreeView.DblClickItem("|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|Ð³×³Ëáñ¹Ý»ñ")
    Call Fill_Clients(majorClient)
    
    ' Ջնջել հաճախորդների կապակցումը
    Log.Message "Ջնջել հաճախորդների կապակցումը", "", pmNormal, DivideColor
    Call Delete_Conjunction(1, "frmPttel_2", conjuction.comment(0), "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ")
    
    ' Կատարել SQL ստուգում Հաճախորդների կապակցումը ջնջելուց հետո
    Log.Message "Կատարել SQL ստուգում Հաճախորդների կապակցումը ջնջելուց հետո", "", pmNormal, DivideColor
    Call CheckDB_Delete_Conjuction()
    
    ' Փակել Հաճախորդներ թղթապանակը
    Call Close_Window(wMDIClient, "frmPttel")
      
    ' Փակել ՀԾ - Բանկ համակարգգը
    Call Close_AsBank()
End Sub

Sub Test_StartUp()
				Call Initialize_AsBankQA(sDate, eDate) 
				Login("ARMSOFT")
				' Մուտք OLAP ադմինիստրատորի ԱՇՏ
				Call ChangeWorkspace(c_OLAPAdmin)
End Sub

Sub Test_Inintialize()
				sDate = "20030101"
				eDate = "20250101"
    
    Set majorClient = New_Clients()
    majorClient.ClientsCode = "00006296"
    
    minorClient = "00006803"
    clients(0) = "00006803"
		
    Set exportOLAP = New_OLAP_Export()
    With exportOLAP
        .dateStart = aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%d%m%y")
        .dateEnd = aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%d%m%y")
        .exportOLAP = 1
    End With 
    
    Set report14 = New_Report14()
    With report14
        .startDate = aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%d%m%y")
        .endDate = aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%d%m%y")
        .calculateFund = 1
    End With
    
    Set conjuction = New_Conjuction(1)
    With conjuction
        .client(0) = "00006803"
        .name(0) = "Ð³×³Ëáñ¹ 00006803"
        .conjuctType(0) = "102"
        .conjuctName(0) = "öáËÏ³å³Ïóí³Í ÁÝÏ»ñáõÃÛáõÝ"
        .comment(0) = "²ÏïÇí ãáõÝ»óáÕ Ñ³×³Ëáñ¹"
    End With
End Sub

Sub CheckDB_Conjuction()
    Dim i, dbo_FOLDERS(3)
    
    Log.Message "fISN = " & conjuction.fIsn, "", pmNormal, MessageColor
    
    'SQL Ստուգում DOCLOG աղուսյակում համար
    Log.Message "SQL Ստուգում DOCLOG աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCLOG", "fISN", conjuction.fIsn, 2)
    Call CheckDB_DOCLOG(conjuction.fIsn, "10", "N", "1", "", 1)
    Call CheckDB_DOCLOG(conjuction.fIsn, "10", "C", "2", "", 1)
  
    'SQL Ստուգում DOCS աղուսյակում 
    Log.Message "SQL Ստուգում DOCS աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCS", "fISN", conjuction.fIsn, 1)
    Call CheckDB_DOCS(conjuction.fIsn, "CliRel", "2", "%CLICODE:00006296%", 1)
    
    'SQL Ստուգում DOCSG աղուսյակում 
    Log.Message "SQL Ստուգում DOCSG աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCSG", "fISN", conjuction.fIsn, 3)
    Call CheckDB_DOCSG(conjuction.fIsn, "RELCLIENTS", "0", "CLICODE", "00006803", 1)
    Call CheckDB_DOCSG(conjuction.fIsn, "RELCLIENTS", "0", "COMMENT", "²ÏïÇí ãáõÝ»óáÕ Ñ³×³Ëáñ¹                                                                                                                     ", 1)
    Call CheckDB_DOCSG(conjuction.fIsn, "RELCLIENTS", "0", "RELTYPE", "102", 1)
  
    'SQL Ստուգում FOLDERS աղուսյակում
    For i = 0 To 2
        Set dbo_FOLDERS(i) = New_DB_FOLDERS()
        With dbo_FOLDERS(i)
            .fISN = conjuction.fIsn
            .fNAME = "CliRel"
            .fSTATUS = "1"
            .fCOM = "Î³å³Ïóí³Í Ñ³×³Ëáñ¹Ý»ñ"
        End With 
    Next
    With dbo_FOLDERS(0)
        .fKEY = "00006296407584625"
        .fFOLDERID = "C.407584625"
        .fSPEC = "Ø»ÏÝ. - "
        .fECOM = "Related Clients"
    End With 
    With dbo_FOLDERS(1)
        .fKEY = "00006296"
        .fFOLDERID = "CliRelCode"
    End With 
    With dbo_FOLDERS(2)
        .fKEY = "00006296/00006803"
        .fFOLDERID = "CliRels"
        .fSPEC = "102²ÏïÇí ãáõÝ»óáÕ Ñ³×³Ëáñ¹"
    End With 
    Log.Message "SQL Ստուգում FOLDERS աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("FOLDERS", "fISN", conjuction.fIsn, 3)
    For i = 0 To 2
        Call CheckDB_FOLDERS(dbo_FOLDERS(i), 1)
    Next
End Sub

Sub CheckDB_Delete_Conjuction()
    Dim dbo_FOLDERS
    
    'SQL Ստուգում DOCLOG աղուսյակում համար
    Log.Message "SQL Ստուգում DOCLOG աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCLOG", "fISN", conjuction.fIsn, 3)
    Call CheckDB_DOCLOG(conjuction.fIsn, "10", "D", "999", "", 1)
    
    'SQL Ստուգում DOCS աղուսյակում 
    Log.Message "SQL Ստուգում DOCS աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCS", "fISN", conjuction.fIsn, 1)
    Call CheckDB_DOCS(conjuction.fIsn, "CliRel", "999", "%CLICODE:00006296%", 1)
    
    'SQL Ստուգում DOCSG աղուսյակում 
    Log.Message "SQL Ստուգում DOCSG աղուսյակում", "", pmNormal, SqlDivideColor
    Call CheckQueryRowCount("DOCSG", "fISN", conjuction.fIsn, 3)
    
    'SQL Ստուգում FOLDERS աղուսյակում 
    Log.Message "SQL Ստուգում FOLDERS աղուսյակում", "", pmNormal, SqlDivideColor
    Set dbo_FOLDERS = New_DB_FOLDERS()
    With dbo_FOLDERS
        .fISN = conjuction.fIsn
        .fKEY = conjuction.fIsn
        .fNAME = "CliRel"
        .fSTATUS = "0"
        .fFOLDERID = ".R." & aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%Y%m%d")
        .fSPEC = Left_Align(Get_Compname_DOCLOG(conjuction.fIsn), 16) & "GlavBux ARMSOFT                       002  "
        .fCOM = ""
        .fECOM = ""
    End With 
    Call CheckQueryRowCount("FOLDERS", "fISN", conjuction.fIsn, 1)
    Call CheckDB_FOLDERS(dbo_FOLDERS, 1)
End Sub