Option Explicit
'USEUNIT SWIFT_International_Payorder_Library
'USEUNIT Library_Common
'USEUNIT Library_Colour
'USEUNIT Library_CheckDB
'USEUNIT Constants
'USEUNIT CashOutput_Confirmpases_Library

'Test case Id 183646

Dim dbFOLDERS(2)
    
Sub SWIFT_Import_And_Refuse_Test()

    Dim sDATE,fDATE
    Dim docNum,fBODY
    Dim max,min,rand,fileFrom,fileTo,what,fWith
    Dim SwiftIsn
    
    'Համակարգ մուտք գործել ARMSOFT օգտագործողով
    sDATE = "20020101"
    fDATE = "20260101"
    Call Initialize_AsBank("bank", sDATE, fDATE)
    Login("ARMSOFT")

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''--- "S.W.I.F.T. ԱՇՏ/Պարամետրեր"-ում կատարել համապատասխան փոփոխությունները---''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- S.W.I.F.T. ԱՇՏ/Պարամետրեր-ում կատարել համապատասխան փոփոխությունները --",,,DivideColor
        
    'Մուտք գործել "S.W.I.F.T. ԱՇՏ"
    Call ChangeWorkspace(c_SWIFT)
    Call wTreeView.DblClickItem("|S.W.I.F.T. ²Þî                  |ä³ñ³Ù»ïñ»ñ")
    BuiltIn.Delay(3000)
    
    max=100
    min=999
    Randomize
    rand = Int((max-min+1)*Rnd+min)
    fileFrom = Project.Path &"Stores\SWIFT\HT103\ImportFile\IA000385.RJE"
    fileTo = Project.Path &"Stores\SWIFT\HT103\ImportFile\Import\IA000387.RJE"
    aqFileSystem.DeleteFile(Project.Path &"Stores\SWIFT\HT103\ImportFile\Import\*")
    what = "UBSWCHZHXXXX901"
    fWith = "UBSWCHZHXXXX" & rand
    
    Log.Message(fWith)
    
    'Վերարտագրում է նախօրոք տրված ֆայլը մեկ այլ ֆայլի մեջ` կատարելով փոփոխություն
    Call Read_Write_File_With_replace(fileFrom,fileTo,what,fWith)
    
    Call SetParameter_InPttel("SWOUT",Project.Path & "Stores\SWIFT\HT103\ImportFile\Import\")
    
    Login("ARMSOFT")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''--- Կատարել Ընդունել SWIFT համակարգից գործողությունը ---''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Կատարել Ընդունել SWIFT համակարգից գործողությունը --",,,DivideColor
    
    'Մուտք գործել "S.W.I.F.T. ԱՇՏ"
    Call ChangeWorkspace(c_SWIFT)
    Call wTreeView.DblClickItem("|S.W.I.F.T. ²Þî                  |Ð³Õáñ¹³·ñáõÃÛáõÝÝ»ñÇ ÁÝ¹áõÝáõÙ|ÀÝ¹áõÝ»É S.W.I.F.T. Ñ³Ù³Ï³ñ·Çó")
    Call ClickCmdButton(5, "OK")
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''--- Ստուգում է փաստաթղթի առկայությունը ---''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Ստուգում է փաստաթղթի առկայությունը --",,,DivideColor       
    
    'Մուտք գործել "Արտաքին փոխանցումների ԱՇՏ"
    Call ChangeWorkspace(c_ExternalTransfers)
    Call wTreeView.DblClickItem("|²ñï³ùÇÝ ÷áË³ÝóáõÙÝ»ñÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|êï³óí³Í Ñ³ÝÓÝ³ñ³ñ³·ñ»ñ|Ð³ßí³éÙ³Ý »ÝÃ³Ï³")
    Call Rekvizit_Fill("Dialog",1,"General", "PERN",aqDateTime.Today)
    Call Rekvizit_Fill("Dialog",1,"General", "PERK",aqDateTime.Today)
    Call ClickCmdButton(2, "Î³ï³ñ»É") 
    BuiltIn.Delay(4000) 
    
    'Ստուգում է փաստաթղթի առկայությունը
    docNum = "951394"
    Call SearchInPttel("frmPttel",2, DocNum)
    
    'Վերցնում է հանձնարարգրի isn-ը
    SwiftIsn = GetIsn()
    Log.Message "SWIFT fISN = "& SwiftIsn,,,SqlDivideColor
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''-- Հանձնարարգրում լրացնում է "Տարանցիկ հաշիվ" դաշտը --''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Հանձնարարգրում լրացնում է Տարանցիկ հաշիվ դաշտը --",,,DivideColor    
    
    'Խմբագրում է հանձնարարգիրը
    Call wMainForm.MainMenu.Click(c_AllActions)    
    Call wMainForm.PopupMenu.Click(c_ToEdit)
    'Լրացնում է "Տարանցիկ հաշիվ" դաշտը
    Call Rekvizit_Fill("Document",2,"General","TCORRACC","000548101")
    Call ClickCmdButton(1, "Î³ï³ñ»É")
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''-- Կատարում է "Հաշվառել" գործողությունը --''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Կատարում է Հաշվառել գործողությունը--",,,DivideColor        
    
    'Կատարում է "Հաշվառել" գործողությունը
    BuiltIn.Delay(2000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_DoTrans)
    
    BuiltIn.Delay(2000)
    Call ClickCmdButton(1, "Î³ï³ñ»É")
    BuiltIn.Delay(2000)
    Call Close_Window(wMDIClient, "frmPttel")
    Call Close_Window(wMDIClient, "FrmSpr")
    
    Call wTreeView.DblClickItem("|²ñï³ùÇÝ ÷áË³ÝóáõÙÝ»ñÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|êï³óí³Í Ñ³ÝÓÝ³ñ³ñ³·ñ»ñ|î³ñ³ÝóÇÏ")
    Call Rekvizit_Fill("Dialog",1,"General", "PERN",aqDateTime.Today)
    Call Rekvizit_Fill("Dialog",1,"General", "PERK",aqDateTime.Today)
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    BuiltIn.Delay(2000)
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''-- Կատարում է "Մերժել (ստացվածը)" գործողությունը --'''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Կատարում է Մերժել (ստացվածը) գործողությունը--",,,DivideColor        
    
    Call wMainForm.MainMenu.Click(c_AllActions)    
    Call wMainForm.PopupMenu.Click(c_RefuseRecieved)

    BuiltIn.Delay(4000)
    Call ClickCmdButton(1, "Î³ï³ñ»É")    
    
    BuiltIn.Delay(2000)
    Call wMainForm.MainMenu.Click(c_Windows)
    Call wMainForm.PopupMenu.Click(c_ClAllWindows)
    
    'SQL Ստուգում DOCS աղուսյակում
    fBODY = "  USERID:  77  ACSBRANCH:00  ACSDEPART:1  BLREP:0  DOCNUM:ART046  ACCCR:7770003485010101  RECEIVER:1/asdfasdf  RECADDR:2/sdfgsdfg      "&_
    "                   3/BE                               3/asdf  REALACC:7770003485010101  ACCDB:223041172402  PAYER:1/ABCD  PAYADDR:2/HASCE,POXOC"&_
    "                      3/AM/QAXAQ  SUMMA:6560  CUR:001  AIM:/ACC TO THE INVOICE                //N88/20 DD 13.09.2009  CLITRANS:1  PAYSYSIN:5  "&_
    "TOTAL:6560  XTOTAL:6560  OCUR:001  BMDOCNUM:600091016ART046  TCORRACC:000548101  CORRACC:01080463012  EXPTYPE:BEN  FORTRADE:0  ACC2ACC:0  "&_
    "EPSSTATE:Received  INITDATE:20220131  CLICODE:00034850  REFUSEPAR:-1  TYPECODE:-10 20 21 22 23 24 30 31 32 25 26 93 11 27 33 28  COVER:1  "&_
    "DUPLICATE:0  RCORBANK:UBSWCHZHXXX  PINSTOP:A  PAYINST:CITIUS33  COUNTRY:US  INCHARGE:0  ACCTYPE:C  CORTYPE:3  SNDREC:UBSWCHZHXXX  MT:103  "&_
    "GRPBMDOCNUM:600091016ART046  GRPSUMMA:6560  NOTSENDABLE:0  "
    fBODY = Replace(fBODY, "  ", "%")
    Call CheckQueryRowCount("DOCS","fISN",SwiftIsn,1)
    Call CheckDB_DOCS(SwiftIsn,"DbPayFor","4",fBODY,1)
    
    'SQL Ստուգում DOCLOG աղուսյակում
    Call CheckQueryRowCount("DOCLOG","fISN",SwiftIsn,6)
    Call CheckDB_DOCLOG(SwiftIsn,"77","N","10","",1)
    Call CheckDB_DOCLOG(SwiftIsn,"77","T","10","",2)
    Call CheckDB_DOCLOG(SwiftIsn,"77","T","2","",1)
    Call CheckDB_DOCLOG(SwiftIsn,"77","M","2","REFUSE PROCESSED",1)
    Call CheckDB_DOCLOG(SwiftIsn,"77","C","4","",1)
    
    'SQL Ստուգում FOLDERS աղուսյակում
    Call SQL_Initialize_For_Actions(SwiftIsn,"")
    Call CheckQueryRowCount("FOLDERS","fISN",SwiftIsn,2)
    Call CheckDB_FOLDERS_With_Like(dbFOLDERS(1),1)
    Call CheckDB_FOLDERS_With_Like(dbFOLDERS(2),1)
    
    Call Close_AsBank()      
End Sub

Sub SQL_Initialize_For_Actions(fISN,docNum)

    Set dbFOLDERS(1) = New_DB_FOLDERS()
    With dbFOLDERS(1)
        .fFOLDERID = "EPS."&fISN
        .fNAME = "DbPayFor"
        .fKEY = "600091016ART046"
        .fISN = fISN
        .fSTATUS = "0"
        .fCOM = "ØÇç³½·. í×³ñÙ³Ý Ñ³ÝÓÝ³ñ³ñ³·Çñ (ëï.)"
        .fSPEC = "%00034850101011/asdfasdf                         01080463012CITIUS33XXX                                       US            "&_
                 "0.00Ø»ñÅí³Í              BEN   77700            0.00 /ACC TO THE INVOICE                             600091016ART046         "&_
                 " 6560.00      77                                              5%"
        .fECOM = "Foreign Payment Order (received)"
        .fDCBRANCH = "00"
        .fDCDEPART = "1"
    End With  
    
    Set dbFOLDERS(2) = New_DB_FOLDERS()
    With dbFOLDERS(2)
        .fFOLDERID = "PayR."&aqConvert.DateTimeToFormatStr(aqDateTime.Now(),"%Y%m%d")
        .fNAME = "DbPayFor"
        .fKEY = fISN
        .fISN = fISN
        .fSTATUS = "1"
        .fCOM = "ØÇç³½·. í×³ñÙ³Ý Ñ³ÝÓÝ³ñ³ñ³·Çñ (ëï.)"
        .fSPEC = "%ART04677700010804630127770003485010101         6560.00001Ø»ñÅí³Í               77/ACC TO THE INVOICE             1/ABCD        "&_
                 "                  1/asdfasdf                         5 %"
        .fECOM = "Foreign Payment Order (received)"
        .fDCBRANCH = "00"
        .fDCDEPART = "1"
    End With 
End Sub