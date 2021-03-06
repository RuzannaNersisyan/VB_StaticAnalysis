'USEUNIT Library_Common
'USEUNIT Constants
'USEUNIT Mem_Order_Library
'USEUNIT Library_Colour
'USEUNIT Library_CheckDB
'USEUNIT DAHK_Library_Filter
'USEUNIT Currency_Exchange_Confirmphases_Library
'USEUNIT Online_PaySys_Library
'USEUNIT Library_Contracts
'USEUNIT Percentage_Calculation_Filter_Library


'Test case ID 180662 
Option Explicit
Dim dbFOLDERS(6)

Sub Memorial_Order_AccEdit_Test ()
    Dim sDate, eDate, Memorder(4), docPath, startDate, workingDocs
    Dim fBODY, dAccIsnEdit, cAccIsnEdit, dAccIsn, cAccIsn
    Dim stDate , enDate, docsP, wISN, accRow, wUser, acsBranch, acsDepart, folderDirect

' ---------------------------------------------------------
' -------------Ստեղծել 3 հիշարար օրդեր -----------------------
' ---------------------------------------------------------
 
    Log.Message  "Հիշարար օրդերիների ստեղծում",,,DivideColor
    Set Memorder(1) = New_Memorder()
    With Memorder(1)  
        .Div = "01"
        .Dep = "4"
        .MDate = "191222"
        .AccD = "01078573313"
        .AccC = "000011003  "
        .Curr = "003"
        .Sum = "1.00"
        .Aim = "0"
        .paysys = "^A[Del]"
    End With
    Set Memorder(2) = New_Memorder()
    With Memorder(2)
        .Div = "00"
        .Dep = "3"
        .MDate = "260622"
        .AccD = "01077663311"
        .AccC = "000488001  "
        .Curr = "001"
        .Sum = "22,222,222,222.22"
        .Aim = "1234567890+Npatak:;1234567890*Npatak1234567890_Npatak)(1234567890/Npatak1234567890Npatak1234567890?Npatak1234567890+Npatak1234567890Npatak-" '140 նիշ
    End With
    Set Memorder(3) = New_Memorder()
    With Memorder(3)
        .Div = "01"
        .Dep = "4"
        .MDate = "191222"
        .AccD = "01078573313"
        .AccC = "000011003  "
        .Curr = "003"
        .Sum = "1.00"
        .Aim = "0"
        .paysys = "^A[Del]"
    End With
    
    Set Memorder(4) = New_Memorder()
    With Memorder(4)
        .Div = "01"
        .Dep = "4"
        .MDate = "191222"
        .AccD = "01078573313"
        .AccC = "000011003  "
        .Curr = "003"
        .Sum = "1.00"
        .Aim = "0"
        .paysys = "^A[Del]"
    End With
    
    docPath = Project.Path &  "Stores\MemorialOrder\Memorder_With_Acc_Edit_0_Exp.txt" 

     'Մուտք գործել ծրագիր ARMSOFT օգտագործողով
    sDate = "20050101"
    eDate = "20250101"
    Call Initialize_AsBank("bank", sDATE, eDATE)
    Call Login ("ARMSOFT")
    cAccIsn = GetAccountISN(Memorder(1).AccC)
    dAccIsn = GetAccountISN(Memorder(1).AccD)
        
    cAccIsnEdit = GetAccountISN(Memorder(2).AccC)
    dAccIsnEdit = GetAccountISN(Memorder(2).AccD)
    'Մուտք Հաճախորդների սպասարկում և դրամարկղ (ընդլայված)
    Call ChangeWorkspace(c_CustomerService)
    'Անցնել Նոր փաստաթղթեր/ Վճարային փաստաթղթեր/ Հիշարար օրդեր
    Call wTreeView.DblClickItem("|Ð³×³Ëáñ¹Ç ëå³ë³ñÏáõÙ ¨ ¹ñ³Ù³ñÏÕ |Üáñ ÷³ëï³ÃÕÃ»ñ|ì×³ñ³ÛÇÝ ÷³ëï³ÃÕÃ»ñ|ÐÇß³ñ³ñ ûñ¹»ñ")
    
    'Ստեղծել 1 վավեր ("Հաջորդը" գործողության միջոցով) և 2 սևագիր հիշարար օրդեր
    BuiltIn.Delay(2000)
    Call Fill_Memorder(Memorder(1))
    Call ClickCmdButton (1,"ê¨³·Çñ")
    
    
    'SQL Database Checks for Memorder(1)
    Log.Message  "SQL Database Checks for Memorder(1)",,,SqlDivideColor
    Call Intitialize_DB_Memorder (Memorder(1).Isn,Memorder(1).DocN)
    'DOCLOG
    Call CheckQueryRowCount("DOCLOG","fISN",Memorder(1).Isn,2)
    Call CheckDB_DOCLOG(Memorder(1).Isn,"77","F","0"," ",1)
    Call CheckDB_DOCLOG(Memorder(1).Isn,"77","N","0"," ",1)
    'DOCS
    fBODY = "  TYPECODE1:-10 20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28  TYPECODE2:-10 20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28  "_
            & "USERID:  77  ACSBRANCH:01  ACSDEPART:4  DOCNUM:"&Memorder(1).DocN&"  DATE:20221219  ACCDB:01078573313  ACCCR:000011003  CUR:003  SUMMA:1  "_
            &"AIM:0  USEOVERLIMIT:0  NOTSENDABLECR:0  NOTSENDABLEDB:0  SBQENABLED:1  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",Memorder(1).Isn,1)
    Call CheckDB_DOCS(Memorder(1).Isn,"MemOrd  ","0",fBODY,1)
    
    'FODLERS
    Call CheckQueryRowCount("FOLDERS","fISN",Memorder(1).Isn,1) 
    Call CheckDB_FOLDERS(dbFOLDERS(1),1)
   
   
    BuiltIn.Delay(2000)
    Call Fill_Memorder(Memorder(3))
    Call ClickCmdButton (1,"Ð³çáñ¹Á")
    If wMDIClient.WaitVBObject("FrmSpr",1000).Exists Then
       Call wMainForm.MainMenu.Click (c_Windows&"|Բոլոր պատուհանները")
       Sys.Process("Asbank").VBObject("frmWindowList").VBObject("ListWindows").DblClickItem("§ÐÇß³ñ³ñ ûñ¹»ñ¦ ÷³ëï³ÃÕÃÇ ïå»Éáõ Ó¨Á")
       Call Memorder_Doc_Check(docPath,1) 
       Call Close_Window(wMDIClient, "FrmSpr" )   
    Else 
        Log.Error "Order print Sample doesn't exists",,,ErrorColor
    End If
     
     
     'SQL Database Checks for Memorder(3)
     Log.Message  "SQL Database Checks for Memorder(3)",,,SqlDivideColor
    Call Intitialize_DB_Memorder (Memorder(3).Isn,Memorder(3).DocN)
    'DOCLOG
    Call CheckQueryRowCount("DOCLOG","fISN",Memorder(3).Isn,2)
    Call CheckDB_DOCLOG(Memorder(3).Isn,"77","C","10"," ",1)
    Call CheckDB_DOCLOG(Memorder(3).Isn,"77","N","1"," ",1)
    'DOCS
    fBODY = "  TYPECODE1:-10 20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28  TYPECODE2:-10 20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28  "_
            & "USERID:  77  ACSBRANCH:01  ACSDEPART:4  DOCNUM:"&Memorder(3).DocN&"  DATE:20221219  ACCDB:01078573313  ACCCR:000011003  CUR:003  SUMMA:1  "_
            &"AIM:0  USEOVERLIMIT:0  NOTSENDABLECR:0  NOTSENDABLEDB:0  SBQENABLED:1  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",Memorder(3).Isn,1)
    Call CheckDB_DOCS(Memorder(3).Isn,"MemOrd  ","10",fBODY,1)
    'FODLERS
    Call CheckQueryRowCount("FOLDERS","fISN",Memorder(3).Isn,1) 
    Call CheckDB_FOLDERS(dbFOLDERS(2),1)
    'HI
    Call Check_HI_CE_accounting ("20221219",Memorder(3).Isn, "11", cAccIsn ,"450.00", "003", "1.00", "MSC", "C")
    Call Check_HI_CE_accounting ("20221219",Memorder(3).Isn, "11", dAccIsn ,"450.00", "003", "1.00", "MSC", "D")
    Call CheckQueryRowCount("HI","fBASE",Memorder(3).Isn,2) 
    'HIREST
    Call CheckDB_HIREST("11", cAccIsn , "-450.00" ,"003", "-1.00", 1)
    Call CheckDB_HIREST("11", dAccIsn , "450.00" ,"003", "1.00", 1)
    
    
    BuiltIn.Delay(2000)
    Call Fill_Memorder(Memorder(4))
    Call ClickCmdButton (1, "ê¨³·Çñ")
    BuiltIn.Delay(2000)
    
    'SQL Database Checks for Memorder(4)
    Log.Message  "SQL Database Checks for Memorder(4)",,,SqlDivideColor
    Call Intitialize_DB_Memorder (Memorder(4).Isn,Memorder(4).DocN)
    'DOCLOG
    Call CheckQueryRowCount("DOCLOG","fISN",Memorder(4).Isn,2)
    Call CheckDB_DOCLOG(Memorder(4).Isn,"77","F","0"," ",1)
    Call CheckDB_DOCLOG(Memorder(4).Isn,"77","N","0"," ",1)
    'DOCS
    fBODY = "  TYPECODE1:-10 20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28  TYPECODE2:-10 20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28  "_
            & "USERID:  77  ACSBRANCH:01  ACSDEPART:4  DOCNUM:"&Memorder(4).DocN&"  DATE:20221219  ACCDB:01078573313  ACCCR:000011003  CUR:003  SUMMA:1  "_
            &"AIM:0  USEOVERLIMIT:0  NOTSENDABLECR:0  NOTSENDABLEDB:0  SBQENABLED:1  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",Memorder(4).Isn,1)
    Call CheckDB_DOCS(Memorder(4).Isn,"MemOrd  ","0",fBODY,1)
    'FODLERS
    Call CheckQueryRowCount("FOLDERS","fISN",Memorder(4).Isn,1) 
    Call CheckDB_FOLDERS(dbFOLDERS(1),1)

    Call ClickCmdButton(1,"¸³¹³ñ»óÝ»É")
    docPath = Project.Path &  "Stores\MemorialOrder\Memorder_With_Acc_Edit_0_Exp.txt"

    'Բացել Օգտագործողի սևագրեր թղթապանակը
    Memorder(1).MDate = "19/12/22" 
    Memorder(1).paysys = ""
    Memorder(3).MDate = "19/12/22" 
    Memorder(3).paysys = ""   
    Memorder(4).MDate = "19/12/22" 
    Memorder(4).paysys = ""   
        
    Log.Message  "Բացել Օգտագործողի սևագրեր թղթապանակը",,,DivideColor
    If Online_PaySys_Check_Doc_In_Drafts(Memorder(1).Isn) Then
        'Խմբագրել փաստաթուղթը և դադարեցնել,ստուգել որ այն չի փոփոխվել
        BuiltIn.Delay(2000)
        Call wMainForm.MainMenu.Click (c_Allactions)
        Call wMainForm.PopupMenu.Click (c_ToEdit)
        BuiltIn.Delay(2000)
        Memorder(1).DocN = Get_Rekvizit_Value("Document",1,"General","DOCNUM")
        Call Memorder_Window_Check(Memorder(1))
        Call Rekvizit_Fill("Document",1,"General","SUMMA","6666")
        Call ClickCmdButton(1,"¸³¹³ñ»óÝ»É")
        
        Call SearchInPttel("frmPttel",2,Memorder(1).Isn)
        Call wMainForm.MainMenu.Click (c_Allactions)
        Call wMainForm.PopupMenu.Click (c_ToEdit)      
        Call Memorder_Window_Check(Memorder(1))
        'Հաստատել փաստաթուղթը Սևագրեր թղթապանակից
        Call ClickCmdButton(1,"Î³ï³ñ»É")
        BuiltIn.Delay(2000)
        If wMDIClient.WaitVBObject("FrmSpr",1000).Exists Then
        '   Ստուգել փաստաթղթի համապատասխանությունը օրինակի հետ 
            Call Memorder_Doc_Check(docPath,3)    
            Call Close_Window(wMDIClient, "FrmSpr" )
        Else
            Log.Error "Order print Sample doesn't exists",,,ErrorColor
        End If
    End If
     
    
    'SQL Database Checks for Memorder(1)
    Log.Message  "SQL Database Checks for Memorder(1)",,,SqlDivideColor
    Call Intitialize_DB_Memorder (Memorder(1).Isn,Memorder(1).DocN)
    'DOCLOG
    Call CheckQueryRowCount("DOCLOG","fISN",Memorder(1).Isn,3)
    Call CheckDB_DOCLOG(Memorder(1).Isn,"77","N","0"," ",1)
    Call CheckDB_DOCLOG(Memorder(1).Isn,"77","F","0"," ",1)
    Call CheckDB_DOCLOG(Memorder(1).Isn,"77","E","10"," ",1)
    'DOCS
    fBODY = "  TYPECODE1:-10 20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28  TYPECODE2:-10 20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28  "_
            & "USERID:  77  ACSBRANCH:01  ACSDEPART:4  DOCNUM:"&Memorder(1).DocN&"  DATE:20221219  ACCDB:01078573313  ACCCR:000011003  CUR:003  SUMMA:1  "_
            &"AIM:0  USEOVERLIMIT:0  NOTSENDABLECR:0  NOTSENDABLEDB:0  SBQENABLED:1  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",Memorder(1).Isn,1)
    Call CheckDB_DOCS(Memorder(1).Isn,"MemOrd  ","10",fBODY,1)
    'FODLERS
    Call CheckQueryRowCount("FOLDERS","fISN",Memorder(1).Isn,1) 
    Call CheckDB_FOLDERS(dbFOLDERS(2),1)
    'HI
    Call Check_HI_CE_accounting ("20221219",Memorder(3).Isn, "11", cAccIsn ,"450.00", "003", "1.00", "MSC", "C")
    Call Check_HI_CE_accounting ("20221219",Memorder(3).Isn, "11", dAccIsn ,"450.00", "003", "1.00", "MSC", "D")
    Call Check_HI_CE_accounting ("20221219",Memorder(1).Isn, "11", cAccIsn ,"450.00", "003", "1.00", "MSC", "C")
    Call Check_HI_CE_accounting ("20221219",Memorder(1).Isn, "11", dAccIsn ,"450.00", "003", "1.00", "MSC", "D")
    Call CheckQueryRowCount("HI","fBASE",Memorder(3).Isn,2) 
    Call CheckQueryRowCount("HI","fBASE",Memorder(1).Isn,2) 
    'HIREST
    Call CheckDB_HIREST("11", cAccIsn , "-900.00" ,"003", "-2.00", 1)
    Call CheckDB_HIREST("11", dAccIsn , "900.00" ,"003", "2.00", 1)
 
    Call Close_Window(wMDIClient, "frmPttel")
    
    'Ջնջել փաստաթուղթը Սևագրեր թղթապանակից
    If Online_PaySys_Check_Doc_In_Drafts(Memorder(4).Isn) Then
       Call SearchAndDelete( "frmPttel", 2, Memorder(4).Isn , "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ" ) 
    End If
        
    'SQL Database Checks for Memorder(4)
    Log.Message  "SQL Database Checks for Memorder(4)",,,SqlDivideColor
    Call Intitialize_DB_Memorder (Memorder(4).Isn,Memorder(4).DocN)
    'DOCLOG
    Call CheckQueryRowCount("DOCLOG","fISN",Memorder(4).Isn,3)
    Call CheckDB_DOCLOG(Memorder(4).Isn,"77","F","0"," ",1)
    Call CheckDB_DOCLOG(Memorder(4).Isn,"77","N","0"," ",1)
    Call CheckDB_DOCLOG(Memorder(4).Isn,"77","D","999"," ",1)
    'DOCS
    fBODY = "  TYPECODE1:-10 20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28  TYPECODE2:-10 20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28  "_
            & "USERID:  77  ACSBRANCH:01  ACSDEPART:4  DOCNUM:"&Memorder(4).DocN&"  DATE:20221219  ACCDB:01078573313  ACCCR:000011003  CUR:003  SUMMA:1  "_
            &"AIM:0  USEOVERLIMIT:0  NOTSENDABLECR:0  NOTSENDABLEDB:0  SBQENABLED:1  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",Memorder(4).Isn,1)
    Call CheckDB_DOCS(Memorder(4).Isn,"MemOrd  ","999",fBODY,1)
    'FODLERS
    Call CheckQueryRowCount("FOLDERS","fISN",Memorder(4).Isn,1) 
    Call CheckDB_FOLDERS(dbFOLDERS(4),1)
    
    'Փակել Օգտագործողի Սևագրեր թղթապանակը
    Call Close_Window(wMDIClient, "frmPttel")
    
    ' Բացել Աշխատանքային փաստաթղթեր թղթապանակը        
    Log.Message  "Բացել Աշխատանքային փաստաթղթեր թղթապանակը",,,DivideColor 
    Set workingDocs = New_MainAccWorkingDocuments()
    With workingDocs
         .startDate = "010122"
				     .endDate = "010123"
    End With
    Call GoTo_MainAccWorkingDocuments("|Ð³×³Ëáñ¹Ç ëå³ë³ñÏáõÙ ¨ ¹ñ³Ù³ñÏÕ |ÂÕÃ³å³Ý³ÏÝ»ñ|", workingDocs)

' ----------------------------------------------------------
' ------------Ջնջել ստեղծված հիշարար օրդերը առանց հաշվառելու-------
' ----------------------------------------------------------                             
    Call View_Memorder(Memorder(3),"frmPttel")
   Log.Message  "Հիշարար օրդերի ջնջում առանց հաշվառելու",,,DivideColor
    If WaitForPttel("frmPttel") Then
        Call SearchAndDelete( "frmPttel", 2, Memorder(3).DocN , "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ" )
    Else
        Log.Error "Աշխատանքային փաստաթղթեր թղթապանակը չի բացվել",,,ErrorColor
    End If
    
    
    'SQL Database Checks for Memorder(3)
     Log.Message  "SQL Database Checks for Memorder(3)",,,SqlDivideColor
    Call Intitialize_DB_Memorder (Memorder(3).Isn,Memorder(3).DocN)
    'DOCLOG
    Call CheckQueryRowCount("DOCLOG","fISN",Memorder(3).Isn,3)
    Call CheckDB_DOCLOG(Memorder(3).Isn,"77","C","10"," ",1)
    Call CheckDB_DOCLOG(Memorder(3).Isn,"77","N","1"," ",1)
    Call CheckDB_DOCLOG(Memorder(3).Isn,"77","D","999"," ",1)
    'DOCS
    fBODY = "  TYPECODE1:-10 20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28  TYPECODE2:-10 20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28  "_
            & "USERID:  77  ACSBRANCH:01  ACSDEPART:4  DOCNUM:"&Memorder(3).DocN&"  DATE:20221219  ACCDB:01078573313  ACCCR:000011003  CUR:003  SUMMA:1  "_
            &"AIM:0  USEOVERLIMIT:0  NOTSENDABLECR:0  NOTSENDABLEDB:0  SBQENABLED:1  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",Memorder(3).Isn,1)
    Call CheckDB_DOCS(Memorder(3).Isn,"MemOrd  ","999",fBODY,1)
    'FODLERS
    Call CheckQueryRowCount("FOLDERS","fISN",Memorder(3).Isn,1) 
    Call CheckDB_FOLDERS(dbFOLDERS(5),1) 
    'HI
    Call Check_HI_CE_accounting ("20221219",Memorder(1).Isn, "11", cAccIsn ,"450.00", "003", "1.00", "MSC", "C")
    Call Check_HI_CE_accounting ("20221219",Memorder(1).Isn, "11", dAccIsn ,"450.00", "003", "1.00", "MSC", "D")
    Call CheckQueryRowCount("HI","fBASE",Memorder(1).Isn,2) 
    'HIREST
    Call CheckDB_HIREST("11", cAccIsn , "-450.00" ,"003", "-1.00", 1)
    Call CheckDB_HIREST("11", dAccIsn , "450.00" ,"003", "1.00", 1)
    
' ----------------------------------------------------------
' -------Դիտել և Խմբագրել ստեղծված հիշարար օրդերը-----------------
' ----------------------------------------------------------
    
    Call View_Memorder(Memorder(1),"frmPttel")
    Call Edit_Memorder(Memorder(1),Memorder(2),0,"frmPttel")
    
    
    'SQL Database Checks for Memorder(1)
     Log.Message  "SQL Database Checks for Memorder(1)",,,SqlDivideColor
    Call Intitialize_DB_Memorder (Memorder(1).Isn,Memorder(1).DocN)
    'DOCLOG
    Call CheckQueryRowCount("DOCLOG","fISN",Memorder(1).Isn,4)
    Call CheckDB_DOCLOG(Memorder(1).Isn,"77","E","10"," ",2)
    Call CheckDB_DOCLOG(Memorder(1).Isn,"77","F","0"," ",1)
    Call CheckDB_DOCLOG(Memorder(1).Isn,"77","N","0"," ",1)
    'DOCS
    fBODY = "  TYPECODE1:-10 20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28  TYPECODE2:-10 20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28  USERID:  77  "_
            &"ACSBRANCH:00  ACSDEPART:3  BLREP:0  DOCNUM:"&Memorder(1).DocN&"  DATE:20220626  ACCDB:01077663311  ACCCR:000488001  CUR:001  SUMMA:22222222222.22  "_
            &"AIM:1234567890+Npatak:;1234567890*Npatak1234567890_Npatak)(1234567890/Npatak1234567890Npatak1234567890?Npatak1234567890+Npatak1234567890Npatak-  "_
            &"USEOVERLIMIT:0  NOTSENDABLECR:0  NOTSENDABLEDB:0  SBQENABLED:1  "      
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",Memorder(1).Isn,1)
    Call CheckDB_DOCS(Memorder(1).Isn,"MemOrd  ","10",fBODY,1)
    'FOLDERS
    Call CheckQueryRowCount("FOLDERS","fISN",Memorder(1).Isn,1) 
    Call CheckDB_FOLDERS(dbFOLDERS(6),1)
    'HI
    Call Check_HI_CE_accounting ("20220626",Memorder(1).Isn, "11", cAccIsnEdit ,"8888888888888.00", "001", "22222222222.22", "MSC", "C")
    Call Check_HI_CE_accounting ("20220626",Memorder(1).Isn, "11", dAccIsnEdit ,"8888888888888.00", "001", "22222222222.22", "MSC", "D")
    Call CheckQueryRowCount("HI","fBASE",Memorder(1).Isn,2) 
    'HIREST
    Call CheckDB_HIREST("11", cAccIsnEdit , "-8888888888888.00" ,"001", "-22222222222.22", 1)
    Call CheckDB_HIREST("11", dAccIsnEdit , "8888888888888.00" ,"001", "22222222222.22", 1)
    
'   Կատարել "Դիտել փաստաթսուղթը" գործողությունը և ստուգել դրա համապատասխանությունը օրինակի հետ    
    Memorder(2).MDate = "26/06/22"
    Memorder(2).paysys = "" 
    
    Call View_Memorder(Memorder(2),"frmPttel")
    BuiltIn.Delay(2000)
    Call wMainForm.MainMenu.Click (c_Allactions)
    Call wMainForm.PopupMenu.Click (c_ViewDoc)
    BuiltIn.Delay(delay_middle)
    docPath = Project.Path &  "Stores\MemorialOrder\Memorder_With_Acc_Edit_1_Exp.txt"
   
    Call Memorder_Doc_Check(docPath,4)
    
    Call Close_Window(wMDIClient, "FrmSpr" ) 
    
    With Memorder (2)
        .MDate = "26/06/22"
        .Aim = "1234567890+Npatak:;1234567890*Npatak1234567890_Npatak)(1234567890/Npatak1234567890Npatak1234567890?Npatak1234567890+Npatak1234567890Npatak-" '140 նիշ
    End With

'   Հաշվառել հիշարար օրդերը
    Call Register_Payment()
   
    
    'SQL Database Checks for Memorder(1)
     Log.Message  "SQL Database Checks for Memorder(1)",,,SqlDivideColor
    Call Intitialize_DB_Memorder (Memorder(2).Isn,Memorder(2).DocN)
    'DOCLOG
    Call CheckQueryRowCount("DOCLOG","fISN",Memorder(2).Isn,5)
    Call CheckDB_DOCLOG(Memorder(2).Isn,"77","E","10"," ",2)
    Call CheckDB_DOCLOG(Memorder(2).Isn,"77","F","0"," ",1)
    Call CheckDB_DOCLOG(Memorder(2).Isn,"77","N","0"," ",1)
    Call CheckDB_DOCLOG(Memorder(2).Isn,"77","M","5","¶ñ³Ýóí»É »Ý Ó¨³Ï»ñåáõÙÝ»ñÁ",1)
    'DOCS
    fBODY = "  OPERTYPE:MSC  TYPECODE1:-10 20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28  TYPECODE2:-10 20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28  USERID:  77  "_
            &"ACSBRANCH:00  ACSDEPART:3  BLREP:0  DOCNUM:"&Memorder(2).DocN&"  DATE:20220626  ACCDB:01077663311  ACCCR:000488001  CUR:001  SUMMA:22222222222.22  "_
            &"AIM:1234567890+Npatak:;1234567890*Npatak1234567890_Npatak)(1234567890/Npatak1234567890Npatak1234567890?Npatak1234567890+Npatak1234567890Npatak-  "_
            &"USEOVERLIMIT:0  NOTSENDABLECR:0  NOTSENDABLEDB:0  SBQENABLED:1  "      
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",Memorder(2).Isn,1)
    Call CheckDB_DOCS(Memorder(2).Isn,"MemOrd  ","5",fBODY,1)
    'FOLDERS
    Call CheckQueryRowCount("FOLDERS","fISN",Memorder(2).Isn,0)
    'MEMORDERS
    Call CheckDB_MEMORDERS(Memorder(2).Isn,"MemOrd  ","1","20220626","5","22222222222.22","001",1)
    'HI
    Call Check_HI_CE_accounting ("20220626",Memorder(2).Isn, "01", cAccIsnEdit ,"8888888888888.00", "001", "22222222222.22", "MSC", "C")
    Call Check_HI_CE_accounting ("20220626",Memorder(2).Isn, "01", dAccIsnEdit ,"8888888888888.00", "001", "22222222222.22", "MSC", "D")
    Call CheckQueryRowCount("HI","fBASE",Memorder(2).Isn,2) 
    'HIREST
    Call CheckDB_HIREST("01", cAccIsnEdit , "-8888888888888.00" ,"001", "-22222222222.22", 1)
    Call CheckDB_HIREST("01", dAccIsnEdit , "8888300888472.20" ,"001", "22220822221.23", 1)
    
    Call Close_Window(wMDIClient, "frmPttel")
    
    Log.Message  "Բացել Հաշվառված վճարային փաստաթղթեր թղթապանակը",,,DivideColor 
    'Մուտք հաշվառված վճարային փաստաթղթեր թղթապանակ  
    Call wTreeview.DblclickItem("|Ð³×³Ëáñ¹Ç ëå³ë³ñÏáõÙ ¨ ¹ñ³Ù³ñÏÕ |ÂÕÃ³å³Ý³ÏÝ»ñ|Ð³ßí³éí³Í í×³ñ³ÛÇÝ ÷³ëï³ÃÕÃ»ñ")
    Call Rekvizit_Fill ("Dialog",1,"General","PERN","010122")
    Call Rekvizit_Fill ("Dialog",1,"General","PERK","010123")
    Call ClickCmdButton (2, "Î³ï³ñ»É")
    
     If WaitForPttel("frmPttel") Then
        Call SearchAndDelete( "frmPttel", 2, Memorder(2).DocN , "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ" )
    Else
        Log.Error "Աշխատանքային փաստաթղթեր թղթապանակը չի բացվել ",,,ErrorColor
    End If
    
    
    'SQL Database Checks for Memorder(1)
     Log.Message  "SQL Database Checks for Memorder(1)",,,SqlDivideColor
    Call Intitialize_DB_Memorder (Memorder(1).Isn,Memorder(1).DocN)
    'DOCLOG
    Call CheckQueryRowCount("DOCLOG","fISN",Memorder(2).Isn,6)
    Call CheckDB_DOCLOG(Memorder(2).Isn,"77","E","10"," ",2)
    Call CheckDB_DOCLOG(Memorder(2).Isn,"77","F","0"," ",1)
    Call CheckDB_DOCLOG(Memorder(2).Isn,"77","N","0"," ",1)
    Call CheckDB_DOCLOG(Memorder(2).Isn,"77","M","5","¶ñ³Ýóí»É »Ý Ó¨³Ï»ñåáõÙÝ»ñÁ",1)
    Call CheckDB_DOCLOG(Memorder(2).Isn,"77","D","999"," ",1)
    'DOCS
    fBODY = "  OPERTYPE:MSC  TYPECODE1:-10 20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28  TYPECODE2:-10 20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28  USERID:  77  "_
            &"ACSBRANCH:00  ACSDEPART:3  BLREP:0  DOCNUM:"&Memorder(2).DocN&"  DATE:20220626  ACCDB:01077663311  ACCCR:000488001  CUR:001  SUMMA:22222222222.22  "_
            &"AIM:1234567890+Npatak:;1234567890*Npatak1234567890_Npatak)(1234567890/Npatak1234567890Npatak1234567890?Npatak1234567890+Npatak1234567890Npatak-  "_
            &"USEOVERLIMIT:0  NOTSENDABLECR:0  NOTSENDABLEDB:0  SBQENABLED:1  "      
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",Memorder(2).Isn,1)
    Call CheckDB_DOCS(Memorder(2).Isn,"MemOrd  ","999",fBODY,1)
    'FOLDERS
    Call CheckQueryRowCount("FOLDERS","fISN",Memorder(2).Isn,1)
    Call CheckDB_FOLDERS(dbFOLDERS(3),1)
    'MEMORDERS
    Call CheckDB_MEMORDERS(Memorder(2).Isn,"MemOrd  ","1","20220626","5","22222222222.22","001",0)
    'HIREST
    Call CheckDB_HIREST("01", cAccIsnEdit , "0.00" ,"001", "0.00", 7)
    Call CheckDB_HIREST("11", cAccIsnEdit , "0.00" ,"001", "0.00", 2)
    Call CheckDB_HIREST("11", dAccIsnEdit , "0.00" ,"001", "0.00", 2)
    Call CheckDB_HIREST("01", dAccIsnEdit , "0.00" ,"001", "0.00", 1)
    Call CheckDB_HIREST("01", dAccIsnEdit , "-620550438.80" ,"001", "-1400000.99", 2)
    Call CheckDB_HIREST("01", dAccIsnEdit , "-588896416.40" ,"001", "-1400000.99", 3)   
    Call CheckDB_HIREST("01", dAccIsnEdit , "-588000415.80" ,"001", "-1400000.99", 1)  

    Call Close_Window(wMDIClient, "frmPttel")

' ----------------------------------------------------------
' -----Ստուգել փաստաթղթերի առկայությունը Ջնջված փաստաթղթեր-------
' --------------------թղթապանակում---------------------------   
    Log.Message  "Բացել Ջնջված փաստաթղթեր թղթապանակը",,,DivideColor
    startDate = aqConvert.DateTimeToStr(aqDateTime.Today) 
    ' Մուտք  Ջնջված փաստաթղթեր դիալոգ և արժեքների լրացում
      folderDirect = "|Ð³×³Ëáñ¹Ç ëå³ë³ñÏáõÙ ¨ ¹ñ³Ù³ñÏÕ |ÂÕÃ³å³Ý³ÏÝ»ñ|æÝçí³Í ÷³ëï³ÃÕÃ»ñ"
      stDate = startDate
      enDate = startDate
      docsP = "MemOrd"
      wISN = ""
      accRow = 0
      wUser = ""
      acsBranch = ""
      acsDepart = ""
      Call DeletedDocFilter(folderDirect, stDate, enDate, docsP, wISN, accRow, wUser, acsBranch, acsDepart)
    
    If WaitForPttel("frmPttel") Then
        Call SearchInPttel("frmPttel",2,Memorder(2).Isn)
        Call SearchInPttel("frmPttel",2,Memorder(3).Isn)
        Call SearchInPttel("frmPttel",2,Memorder(4).Isn)
    Else
        Log.Error "Ջնջված փաստաթղթեր թղթապանակը չի բացվել",,,ErrorColor
    End If
    Call Close_AsBank()
    
End Sub


Sub Intitialize_DB_Memorder (fISN,fDOCN)
    Dim tday
    
    tday = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"20%y%m%d")
    Set dbFOLDERS(1) = New_DB_FOLDERS()
    With dbFOLDERS(1) 
        .fFOLDERID = ".D.TellerX "
        .fNAME = "MemOrd  "
        .fKEY = fISN
        .fISN = fISN
        .fSTATUS = "1"
        .fCOM = "ÐÇß³ñ³ñ ûñ¹»ñ"
        .fSPEC = ""
        .fECOM = ""
        .fDCBRANCH = "01 "
        .fDCDEPART = "4  "
    End With
      
    Set dbFOLDERS(2) = New_DB_FOLDERS()
    With dbFOLDERS(2) 
        .fFOLDERID = "Oper.20221219"
        .fNAME = "MemOrd  "
        .fKEY = fISN
        .fISN = fISN
        .fSTATUS = "1"
        .fCOM = "ÐÇß³ñ³ñ ûñ¹»ñ"
        .fSPEC = fDOCN & "777000107857331377700000011003              1.00003Üáñ                                                   "_
                 &"77                                                                                                "_
                 &"0                                                                                                                                           "
        .fECOM = "Memorial Order"
        .fDCBRANCH = "01 "
        .fDCDEPART = "4  "
    End With
      
    Set dbFOLDERS(3) = New_DB_FOLDERS()
    With dbFOLDERS(3) 
        .fFOLDERID = ".R."&tday
        .fNAME = "MemOrd  "
        .fKEY = fISN
        .fISN = fISN
        .fSTATUS = "0"
        .fCOM = ""
        .fSPEC = Left_Align(Get_Compname_DOCLOG(fISN), 16)&"TellerX ARMSOFT                       115  "
        .fECOM = ""
        .fDCBRANCH = "00 "
        .fDCDEPART = "3  "
    End With
    
    Set dbFOLDERS(4) = New_DB_FOLDERS()
    With dbFOLDERS(4) 
        .fFOLDERID = ".R."&tday
        .fNAME = "MemOrd  "
        .fKEY = fISN
        .fISN = fISN
        .fSTATUS = "0"
        .fCOM = ""
        .fSPEC = Left_Align(Get_Compname_DOCLOG(fISN), 16)&"TellerX ARMSOFT                       000  "
        .fECOM = ""
        .fDCBRANCH = "01 "
        .fDCDEPART = "4  "
    End With
    
    Set dbFOLDERS(5) = New_DB_FOLDERS()
    With dbFOLDERS(5) 
        .fFOLDERID = ".R."&tday
        .fNAME = "MemOrd  "
        .fKEY = fISN
        .fISN = fISN
        .fSTATUS = "0"
        .fCOM = ""
        .fSPEC = Left_Align(Get_Compname_DOCLOG(fISN), 16)&"TellerX ARMSOFT                       0110 "
        .fECOM = ""
        .fDCBRANCH = "01 "
        .fDCDEPART = "4  "
    End With
    
    Set dbFOLDERS(6) = New_DB_FOLDERS()
    With dbFOLDERS(6) 
        .fFOLDERID = "Oper.20220626"
        .fNAME = "MemOrd  "
        .fKEY = fISN
        .fISN = fISN
        .fSTATUS = "1"
        .fCOM = "ÐÇß³ñ³ñ ûñ¹»ñ"
        .fSPEC = fDOCN & "777000107766331177700000488001    22222222222.22001ÊÙµ³·ñíáÕ                                             "_
                 &"77                                                                                                "_
                 &"1234567890+Npatak:;1234567890*Npatak1234567890_Npatak)(1234567890/Npatak1234567890Npatak1234567890?Npatak1234567890+Npatak1234567890Npatak-"
        .fECOM = "Memorial Order"
        .fDCBRANCH = "00 "
        .fDCDEPART = "3  "
    End With
    
End Sub