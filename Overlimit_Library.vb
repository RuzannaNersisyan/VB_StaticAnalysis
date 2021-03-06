Option Explicit
'USEUNIT Library_Common 
'USEUNIT Library_Colour
'USEUNIT Constants
'USEUNIT Akreditiv_Library
'USEUNIT Library_Contracts

Dim DocForm,NewOverlimitDoc,VerifyOverlimit1,ContractNew,RcOptionOverlimit,CalcPercents,CalcPercents_2
Dim NewOverlimitRepay,NewStore,NewWriteOut,NewWriteOff,NewOverlimitRepay2,dbFOLDERS_ForRate
Dim AccountIsn,AccountParentIsn,ExpectedMessage
Dim fISN,fBODY,fVALUES,Query,dbCONTRACT
Dim fOBJECT(15),dbFOLDERS(18),ActionIsn(4),fBASE(2)
Dim AccWithOverlimit,ContractFillter,PaymentfISN,MemOrdfISN
Dim DocNum,dbFOLDERS_ForDelete,OverlimitAccountIsn,AccIsn(3) 

'---------------------------------------------------------------------------------------------
'OverlimitDoc - "Գերածախս" փաստաթղթի Class
'---------------------------------------------------------------------------------------------
Class OverlimitDoc
    Public Isn
    Public DocType
    Public GeneralTab
    Public ScheduleOrganizingTab
    Public InterestsTab
    Public AdditionalTab
    Public NotificationTab
    Public NotesTab
    Public LoanRegTab
    Public AttachmentsTab
    
    Private Sub Class_Initialize
        Isn = ""
        Set GeneralTab = New_GeneralOverlimit()
        Set ScheduleOrganizingTab = New_ScheduleOrgOverlimit()
        Set InterestsTab = New_InterestsOverlimit()
        Set AdditionalTab = New_AdditionalOverlimit()
        Set NotificationTab = New_NotificationOverlimit()
        Set NotesTab = New_NotesOverlimit()
        Set LoanRegTab = New_LoanRegOverlimit()
        Set AttachmentsTab = New_AttachmentsOverlimit()
    End Sub  
End Class

Function New_OverlimitDoc()
    Set New_OverlimitDoc = NEW OverlimitDoc      
End Function

'---------------------------------------------------------------------------------------------
'Overlimit/General - "Գերածախս/Ընդանուր" tab-ի Class
'---------------------------------------------------------------------------------------------
Class GeneralOverLimit
    Public FillTab
    Public AgreementN
    Public GenerateButton
    Public ExpectedCreditCode
    Public ExpectedClient
    Public Client
    Public ClientComment
    Public ExpectedName
    Public Name
    Public ExpectedCurrency
    Public Curr
    Public CurrencyComment
    Public ExpectedRepaymentCurrency
    Public RepaymentCurrency
    Public RepaymentCurrencyComment
    Public SettlementAccount
    Public SettlementAccountComment
    Public Comment
    Public ExpectedAutomaticallyPaym
    Public AutomaticallyPaym
    Public ExpectedUseOtherAccountRemainders
    Public UseOtherAccountRemainders
    Public UseOtherAccountRemaindersComment
    Public UseClientsSchema
    Public AccountConnectionSchema
    Public SigningDate
    Public SigningDateComment
    Public DisbursementDate
    Public DisbursementDateComment
    Public Division
    Public DivisionComment
    Public Department
    Public AccessType
    Public AccessTypeComment
    Private Sub Class_Initialize
        FillTab = True
        AgreementN = ""
        GenerateButton = False
        ExpectedCreditCode = ""
        ExpectedClient = ""
        Client = ""
        ClientComment = ""
        Name = ""
        ExpectedCurrency = ""
        Curr = ""
        CurrencyComment = ""
        ExpectedRepaymentCurrency = ""
        RepaymentCurrency = ""
        RepaymentCurrencyComment = ""
        SettlementAccount = ""
        SettlementAccountComment = ""
        Comment = ""
        ExpectedAutomaticallyPaym = "1"
        AutomaticallyPaym = ""
        ExpectedUseOtherAccountRemainders = ""
        UseOtherAccountRemainders = ""
        UseOtherAccountRemaindersComment = ""
        UseClientsSchema = "0"
        AccountConnectionSchema = ""
        SigningDate = ""
        SigningDateComment = ""
        DisbursementDate = ""
        DisbursementDateComment = ""
        Division = ""
        DivisionComment = ""
        Department = ""
        AccessType = ""
        AccessTypeComment = ""
    End Sub  
End Class

Function New_GeneralOverLimit()
    Set New_GeneralOverLimit = NEW GeneralOverLimit      
End Function

'---------------------------------------------------------------------------------------------
'Overlimit/ScheduleOrg - "Գերածախս/Գրաֆիկի լրացման ձև" tab-ի Class
'---------------------------------------------------------------------------------------------
Class ScheduleOrgOverlimit
    Public FillTab
    Public ScheduleDateOrgMode
    Public ScheduleDateOrgModeComment
    Public RepaymentDays
    Public PerioducityMonts
    Public PerioducityDays
    Public DayOverpassingMethod
    Public DayOverpassingMethodComment
    Public OvrDays
    Public OvrDaysComment
    
    Private Sub Class_Initialize
        FillTab = False
        ScheduleDateOrgMode = ""
        ScheduleDateOrgModeComment = ""
        RepaymentDays = ""
        PerioducityMonts = ""
        PerioducityDays = ""
        DayOverpassingMethod = ""
        DayOverpassingMethodComment = ""
        OvrDays = ""
        OvrDaysComment = ""
    End Sub  
End Class

Function New_ScheduleOrgOverlimit()
    Set New_ScheduleOrgOverLimit = NEW ScheduleOrgOverlimit      
End Function

'---------------------------------------------------------------------------------------------
'Overlimit/Interests - "Գերածախս/Տոկոսներ" tab-ի Class
'---------------------------------------------------------------------------------------------
Class InterestsOverlimit
    Public FillTab
    Public ExpectedKindOfScale
    Public KindOfScale
    Public FineOnPastDueSum
    Public Div
    
    Private Sub Class_Initialize
        FillTab = False
        ExpectedKindOfScale = ""
        KindOfScale = ""
        FineOnPastDueSum = "0.0000"
        Div = "0"
    End Sub  
End Class

Function New_InterestsOverlimit()
    Set New_InterestsOverlimit = NEW InterestsOverlimit      
End Function

'---------------------------------------------------------------------------------------------
'Overlimit/Additional - "Գերածախս/Լրացուցիչ" tab-ի Class
'---------------------------------------------------------------------------------------------
Class AdditionalOverlimit
    Public FillTab
    Public Sector
    Public SectorComment
    Public UsageField
    Public UsageFieldComment
    Public Aim
    Public AimComment
    Public InternationalOrganization
    Public ProjectName
    Public ProjectComment
    Public Guarantee
    Public GuaranteeComment
    Public Country
    Public CountryComment
    Public Region
    Public RegionComment
    Public RegionNewLR
    Public RegionNewLRComment
    Public Note
    Public Note2
    Public Note3
    Public AgreemPaperN
    Public ClosingDate
    Public FullyClosed
    Public SubjectiveCategorized
    
    Private Sub Class_Initialize
        FillTab = False
        Sector = ""
        SectorComment = ""
        UsageField = ""
        UsageFieldComment = ""
        Aim = ""
        AimComment = ""
        InternationalOrganization = ""
        ProjectName = ""
        ProjectComment = ""
        Guarantee = ""
        GuaranteeComment = ""
        Country = ""
        CountryComment = ""
        Region = ""
        RegionComment = ""
        RegionNewLR = ""
        RegionNewLRComment = ""
        Note = ""
        Note2 = ""
        Note3 = ""
        AgreemPaperN = ""
        ClosingDate = ""
        FullyClosed = ""
        SubjectiveCategorized = ""
    End Sub  
End Class

Function New_AdditionalOverlimit()
    Set New_AdditionalOverlimit = NEW AdditionalOverlimit      
End Function

'---------------------------------------------------------------------------------------------
'Overlimit/Notification - "Գերածախս/Ծանուցում" tab-ի Class
'---------------------------------------------------------------------------------------------

Class NotificationOverlimit
    Public FillTab
    Public NotifyMode
    Public SendNotificationAddress

    
    Private Sub Class_Initialize
        FillTab = False
        NotifyMode = ""
        SendNotificationAddress = ""
    End Sub  
End Class

Function New_NotificationOverlimit()
    Set New_NotificationOverlimit = NEW NotificationOverlimit      
End Function

'---------------------------------------------------------------------------------------------
'Overlimit/Notes - "Գերածախս/Նշումներ" tab-ի Class
'---------------------------------------------------------------------------------------------
Class NotesOverlimit
    Public FillTab
    Public RegisteredNum
    Public RegisteredName
    Public RegisteredValue
    Public RegisteredValue2
    Public NotesName
    Public NotesValue
    Public NotesFill
    Public NotesFillValue
    
    Private Sub Class_Initialize
        FillTab = False
        RegisteredNum = ""
        RegisteredName = ""
        RegisteredValue = ""
        RegisteredValue2  = ""
        NotesName = ""
        NotesValue = ""
        NotesFill = False
        NotesFillValue = ""
    End Sub  
End Class

Function New_NotesOverlimit()
    Set New_NotesOverlimit = NEW NotesOverlimit      
End Function

'---------------------------------------------------------------------------------------------
'Overlimit/LoanReg - "Գերածախս/Վարկ. ռեգ" tab-ի Class
'---------------------------------------------------------------------------------------------

Class LoanRegOverlimit
    Public FillTab
    Public AccumulateInLoanReg
    Public LoanRegisterCode
    Public CountOfChanges
    Public AdditionalInformation
    Public PledgeCurrency
    Public RemnantOfPladge
    Public PladgeObject
    Public PladgeAdditionalInfo
    Public PladgeObjectArca
    Public PladgeObjectArcaComment
    Public NotClassifiable
    Public LRCodeNew
    Public RevisionReason
    Public RepaymentSourse
    Public PledgeObjectNew
    Public GuaranteedByOtherCallateral
    
    Private Sub Class_Initialize
        FillTab = False
        AccumulateInLoanReg = "0"
        LoanRegisterCode = ""
        CountOfChanges = ""
        AdditionalInformation = ""
        PledgeCurrency = ""
        RemnantOfPladge = ""
        PladgeObject = ""
        PladgeAdditionalInfo = ""
        PladgeObjectArca = ""
        PladgeObjectArcaComment = ""
        NotClassifiable = "0"
        LRCodeNew = ""
        RevisionReason = ""
        RepaymentSourse = ""
        PledgeObjectNew = ""
        GuaranteedByOtherCallateral = "0"
    End Sub  
End Class

Function New_LoanRegOverlimit()
    Set New_LoanRegOverlimit = NEW LoanRegOverlimit      
End Function

'---------------------------------------------------------------------------------------------
'Overlimit/Attachments - "Գերածախս/Կցված" tab-ի Class
'---------------------------------------------------------------------------------------------
Class AttachmentsOverlimit
    Public FillTab
    Public AddFile
    Public FilePath
    Public AddLink
    Public Link
    Public Description
    
    Private Sub Class_Initialize
        FillTab = False
        AddFile = False
        FilePath = ""
        AddLink = False
        Link = ""
        Description = ""
    End Sub  
End Class

Function New_AttachmentsOverlimit()
    Set New_AttachmentsOverlimit = NEW AttachmentsOverlimit      
End Function

'---------------------------------------------------------------------------------------------
'Overlimit Verify Doc - "Գերածախս հաստատող փաստաթղթեր" filter-ի Class
'---------------------------------------------------------------------------------------------
Class VerifyOverlimitDoc1
    Public ConFirmationGroup
    Public TermsStates
    Public AgreementOperations
    Public AgreementN
    Public AgreemPaperN
    Public Curr
    Public Client
    Public ClientName
    Public Note
    Public Note2
    Public Note3
    Public Executors
    Public Division
    Public Department
    Public AccessType
    
    Private Sub Class_Initialize
       ConFirmationGroup = ""
       TermsStates = ""
       AgreementOperations = ""
       AgreementN = ""
       AgreemPaperN = ""
       Curr = ""
       Client = ""
       ClientName = ""
       Note = ""
       Note2 = ""
       Note3 = ""
       Executors = ""
       Division = ""
       Department = ""
       AccessType = ""
    End Sub  
End Class

Function New_VerifyOverlimitDoc1()
    Set New_VerifyOverlimitDoc1 = NEW VerifyOverlimitDoc1      
End Function

'---------------------------------------------------------------------------------------------
' RcOverlimit  - ¶»ñ³Í³Ëë ·áñÍáÕáõÃÛ³Ý ³ñ¹ÛáõÝùáõÙ µ³óí³Í ÷³ëï³ÃÕÃÇ Éրացման Class  
'---------------------------------------------------------------------------------------------
Class RcOverlimit
    Public Isn
    Public ExpectedAgreementN
    Public Date
    Public Sum
    Public CashOrNo
    Public CalcAcc
    Public Comment
    Public Division
    Public Department
    
    Private Sub Class_Initialize
        Isn = ""
        ExpectedAgreementN = ""
        Date = ""
        Sum = ""
        CashOrNo = ""
        CalcAcc = ""
        Comment = ""
        Division = ""
        Department = ""
    End Sub  
End Class

Function New_RcOverlimit()
    Set New_RcOverlimit = NEW RcOverlimit      
End Function

'------------------------------------------------------------------------------------
' Գերածախսի(Overlimit) պայմանագրի ստեղծում
'------------------------------------------------------------------------------------
Function CreateOverlimitDoc(Overlimit)
    Dim frmModalBrowser, wTabStrip, TabN
    TabN = 2
    Set frmModalBrowser = p1.WaitVBObject("frmModalBrowser", 500)	
		Do Until p1.frmModalBrowser.VBObject("tdbgView").EOF
			If RTrim(p1.frmModalBrowser.VBObject("tdbgView").Columns.Item(col_item).Text) = Overlimit.DocType  Then
  			Call p1.frmModalBrowser.VBObject("tdbgView").Keys("[Enter]")
  			Exit do
			Else
  			Call p1.frmModalBrowser.VBObject("tdbgView").MoveNext
			End If
		Loop 
    'ISN-ի վերագրում փոփոխականին
    Overlimit.Isn = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
    
    'Լրացնել "Ընդանուր/General" Tab-ի արժեքները
    Call Fill_GeneralTab(Overlimit.GeneralTab)
    
    'Լրացնել "Գրաֆիկի լրացման ձև/ScheduleOrganizing" Tab-ի արժեքները
    If Overlimit.DocType  = "¶»ñ³Í³Ëë (·ñ³ýÇÏáí å³ÛÙ.)" Then
        Call Fill_ScheduleOrganizingTab(Overlimit.ScheduleOrganizingTab)
        TabN = TabN + 1
    End If
    
    'Լրացնել "Տոկոսներ/Interests" Tab-ի արժեքները
    Call Fill_InterestsTab(Overlimit.InterestsTab,TabN)
    
    'Լրացնել "Լրացուցիչ/Additional" Tab-ի արժեքները
    TabN = TabN + 1
    Call Fill_AdditionalTab(Overlimit.AdditionalTab,TabN)
    
    'Լրացնել "Ծանուցում/Notification" Tab-ի արժեքները
    TabN = TabN + 1
    Call Fill_NotificationTab(Overlimit.NotificationTab,TabN)
    
    'Լրացնել "Վարկ.ռեգ/LoanReg" Tab-ի արժեքները
    TabN = TabN + 1
    Call Fill_LoanRegTab(Overlimit.LoanRegTab,TabN)
    
    'Սեղմել "Կատարել"
    Call ClickCmdButton(1, "Î³ï³ñ»É")
End Function    

'------------------------------------------------------------------------------------
' 'Լրացնել "Ընդանուր/General" Tab-ի արժեքները
'------------------------------------------------------------------------------------
Sub Fill_GeneralTab(GeneralTab)

    If GeneralTab.FillTab Then
        Log.Message "Fill General Tab",,,MessageColor
        'Ստուգում "Մարման Արժույթ" դաշտի համար 
        Call Compare_Two_Values("Մարման Արժույթ",Get_Rekvizit_Value("Document",1,"Mask","CURRENCY"),GeneralTab.ExpectedRepaymentCurrency)

        'Լրացնել "Հաշվարկային հաշիվ" դաշտը
        Call Rekvizit_Fill("Document", 1, "General", "ACCACC", GeneralTab.SettlementAccount) 
      
        'Ստուգել "Հաշվարկային հաշիվ" դաշտի Լրացման արդյունքում ավտոմատ լրացվող դաշտերը
        'Ստուգում "Պայմանագրի N" դաշտի համար
        Call Compare_Two_Values("Փաստաթղթի N",Get_Rekvizit_Value("Document",1,"General","CODE"),GeneralTab.AgreementN)
        Call Check_ReadOnly("Document",1,"General","CODE",True)
        'Ստուգում "Հաճախորդ" դաշտի համար
        Call Compare_Two_Values("Հաճախորդ",Get_Rekvizit_Value("Document",1,"Mask","CLICOD"),GeneralTab.ExpectedClient)
        'Ստուգում "Անվանում" դաշտի համար  
        Call Compare_Two_Values("Անվանում",Get_Rekvizit_Value("Document",1,"General","NAME"),GeneralTab.ExpectedName)
        'Ստուգում "Արժույթ" դաշտի համար  
        Call Compare_Two_Values("Արժույթ",Get_Rekvizit_Value("Document",1,"Mask","CURRENCY"),GeneralTab.ExpectedCurrency)
        'Ստուգում "Մարման Արժույթ" դաշտի համար  
        If Not Check_ReadOnly("Document",1,"Mask","REPAYCURR",True) Then
            Call Rekvizit_Fill("Document", 1, "Mask", "REPAYCURR", GeneralTab.RepaymentCurrency) 
        End If
      
        'Լրացնել "Մեկնաբանություն" դաշտը
        Call Rekvizit_Fill("Document", 1, "General", "COMMENT", GeneralTab.Comment) 
        'Ստուգում "Պարտքերի ավտոմատ մարում" դաշտի համար 
        Call Compare_Two_Values("Պարտքերի ավտոմատ մարում",Get_Rekvizit_Value("Document",1,"CheckBox","AUTODEBT"),GeneralTab.ExpectedAutomaticallyPaym)
        If Not Check_ReadOnly("Document",1,"CheckBox","AUTODEBT",True) Then
            'Լրացնել "Պարտքերի ավտոմատ մարում" դաշտը
            Call Rekvizit_Fill("Document", 1, "CheckBox", "AUTODEBT", GeneralTab.AutomaticallyPaym) 
        End If

        'Ստուգում "Այլ հաշիվների մնացորդների օգտագործում" դաշտի համար 
        Call Compare_Two_Values("Այլ հաշիվների մնացորդների օգտագործում",Get_Rekvizit_Value("Document",1,"Mask","ACCCONNMODE"),GeneralTab.ExpectedUseOtherAccountRemainders)
        'Լրացնել "Այլ հաշիվների մնացորդների օգտագործում" դաշտը
        Call Rekvizit_Fill("Document", 1, "General", "ACCCONNMODE", GeneralTab.UseOtherAccountRemainders)
    
        'Լրացնել "Օգտագործել հաճախորդային սխեմա" դաշտը
        Call Rekvizit_Fill("Document",1,"CheckBox","USECLICONNSCH",GeneralTab.UseClientsSchema)
        'Լրացնել "Հածիվների փոխկապակցման սխեմա" դաշտը
        Call Rekvizit_Fill("Document",1,"General","ACCCONNSCH",GeneralTab.AccountConnectionSchema)
    
        'Լրացնել "Կնքման ամսաթիվ" դաշտը
        Call Rekvizit_Fill("Document",1,"General","DATE",GeneralTab.SigningDate)
        'Լրացնել "Հատկացման ամսաթիվ" դաշտը
        Call Rekvizit_Fill("Document",1,"General","DATEGIVE",GeneralTab.DisbursementDate)
        
        'Լրացնել "Գրասենյակ" դաշտը
        Call Rekvizit_Fill("Document",1,"General","ACSBRANCH",GeneralTab.Division)
        'Լրացնել "Բաժին" դաշտը
        Call Rekvizit_Fill("Document",1,"General","ACSDEPART",GeneralTab.Department)
        'Լրացնել "Հասան-ն տիպ" դաշտը
        Call Rekvizit_Fill("Document",1,"General","ACSTYPE",GeneralTab.AccessType)
        
        'Ստուգում "Վարկային կոդ" դաշտի համար 
        Call Check_ReadOnly("Document",1,"General","CRDTCODE",True)
        If GeneralTab.GenerateButton Then
            Call ClickCmdButton(1, "¶»Ý»ñ³óÝ»É")
            GeneralTab.ExpectedCreditCode = Trim(Get_Rekvizit_Value("Document",1,"General","CRDTCODE"))
'            Call Compare_Two_Values("Վարկային կոդ",Get_Rekvizit_Value("Document",1,"General","CRDTCODE"),GeneralTab.ExpectedCreditCode)
        End If
    End If
End Sub

'------------------------------------------------------------------------------------
'Լրացնել "Գրաֆիկի լրացման ձև/ScheduleOrganizing" Tab-ի արժեքները
'------------------------------------------------------------------------------------
Sub Fill_ScheduleOrganizingTab(ScheduleOrganizingTab)

    If ScheduleOrganizingTab.FillTab Then
        Log.Message "Fill ScheduleOrganizing Tab",,,MessageColor
        'Ստուգում "Մարման օրեր" դաշտի խմբագրելիությունը 
        Call Check_ReadOnly("Document",2,"General","FIXEDDAYS",True)
        'Ստուգում "Պարբերություն" դաշտի խմբագրելիությունը 
        Call Check_ReadOnly("Document",2,"General","AGRPERIOD",True)
        'Ստուգում "Ուղղ. օր" դաշտի խմբագրելիությունը 
        Call Check_ReadOnly("Document",2,"General","PASSOVTYPE",True)
        
        'Լրացնել "Ամսաթվերի լրացման ձև" դաշտը
        Call Rekvizit_Fill("Document",2,"General","DATESFILLTYPE",ScheduleOrganizingTab.ScheduleDateOrgMode)
        If Get_Rekvizit_Value("Document",2,"Mask","DATESFILLTYPE") = "1" Then
            'Ստուգում "Մարման օրեր" դաշտի խմբագրելիությունը 
            Call Check_ReadOnly("Document",2,"General","FIXEDDAYS",False)
            'Լրացնել "Մարման օրեր" դաշտը
            Call Rekvizit_Fill("Document",2,"General","FIXEDDAYS",ScheduleOrganizingTab.RepaymentDays)
        End If
        If Get_Rekvizit_Value("Document",2,"Mask","DATESFILLTYPE") = "2" Then
            'Ստուգում "Պարբերություն" դաշտի խմբագրելիությունը 
            Call Check_ReadOnly("Document",2,"General","AGRPERIOD",False)
            'Լրացնել "Պարբերություն" երկու դաշտերը
             Call Rekvizit_Fill("Document",2,"General","AGRPERIOD",ScheduleOrganizingTab.PerioducityMonts & "[Tab]" & ScheduleOrganizingTab.PerioducityDays)
        End If

        'Լրացնել "Շրջանցման ուղղություն" դաշտը
        Call Rekvizit_Fill("Document",2,"General","PASSOVDIRECTION",ScheduleOrganizingTab.DayOverpassingMethod)
        
        If Get_Rekvizit_Value("Document",2,"Mask","PASSOVDIRECTION") = "0" Then
            'Ստուգում "Ուղղ. օր" դաշտի խմբագրելիությունը 
            Call Check_ReadOnly("Document",2,"General","PASSOVTYPE",True)
        Else
            'Լրացնել "Ուղղ. օր" դաշտը
            Call Check_ReadOnly("Document",2,"General","PASSOVTYPE",False)
            Call Rekvizit_Fill("Document",2,"General","PASSOVTYPE",ScheduleOrganizingTab.OvrDays)
        End If
    End If
End Sub

'------------------------------------------------------------------------------------
'Լրացնել "Տոկոսներ/Interests" Tab-ի արժեքները
'------------------------------------------------------------------------------------
Sub Fill_InterestsTab(InterestsTab,tabN)

    If InterestsTab.FillTab Then
        Log.Message "Fill Interests Tab",,,MessageColor
        'Ստուգում "Օրացույցի հաշվարկման ձև" դաշտի խմբագրելիությունը 
        Call Check_ReadOnly("Document",tabN,"Mask","KINDSCALE",True)
        
        'Լրացնել "Օրացույցի հաշվարկման ձև" դաշտը
        If Not Check_ReadOnly("Document",tabN,"Mask","KINDSCALE",True) Then 
            Call Rekvizit_Fill("Document",tabN,"General","KINDSCALE",InterestsTab.KindOfScale)
        End If
        'Լրացնել "Ժամկետանց գումարի տույժ" երկու դաշտերը
        Call Rekvizit_Fill("Document", TabN, "General", "PCPENAGR", InterestsTab.FineOnPastDueSum &"[Tab]"& InterestsTab.Div)
    End If
End Sub

'------------------------------------------------------------------------------------
'Լրացնել "Լրացուցիչ/Additional" Tab-ի արժեքները
'------------------------------------------------------------------------------------
Sub Fill_AdditionalTab(AdditionalTab,tabN)

    Log.Message "Fill Additional Tab",,,MessageColor
    If AdditionalTab.FillTab Then
        'Լրացնել "Ճյուղայնություն" դաշտը
        Call Rekvizit_Fill("Document",tabN,"General","SECTOR",AdditionalTab.Sector)
        'Լրացնել "Օգտագործման ոլորտ" դաշտը
        Call Rekvizit_Fill("Document",tabN,"General","USAGEFIELD",AdditionalTab.UsageField)
        'Լրացնել "Նպատակ" դաշտը
        Call Rekvizit_Fill("Document",tabN,"General","AIM",AdditionalTab.Aim)
        'Լրացնել "Միջազգային կազմակերպություն" դաշտը
        Call Rekvizit_Fill("Document",tabN,"General","INTERORG",AdditionalTab.InternationalOrganization)
        'Լրացնել "Ծրագիր" դաշտը
        Call Rekvizit_Fill("Document",tabN,"General","SCHEDULE",AdditionalTab.ProjectName)
        'Լրացնել "Երաշխավորություն" դաշտը
        Call Rekvizit_Fill("Document",tabN,"General","GUARANTEE",AdditionalTab.Guarantee)
        'Լրացնել "Երկիր" դաշտը
        Call Rekvizit_Fill("Document",tabN,"General","COUNTRY",AdditionalTab.Country)
        'Լրացնել "Մարզ" դաշտը
        Call Rekvizit_Fill("Document",tabN,"General","LRDISTR",AdditionalTab.Region)
        'Լրացնել "Մարզ(նոր ՎՌ)" դաշտը
        Call Rekvizit_Fill("Document",tabN,"General","REGION",AdditionalTab.RegionNewLR)
        'Լրացնել "Նմուշ" դաշտը
        Call Rekvizit_Fill("Document",tabN,"General","NOTE",AdditionalTab.Note)
        'Լրացնել "Նմուշ 2" դաշտը
        Call Rekvizit_Fill("Document",tabN,"General","NOTE2",AdditionalTab.Note2)
        'Լրացնել "Նմուշ 3" դաշտը
        Call Rekvizit_Fill("Document",tabN,"General","NOTE3",AdditionalTab.Note3)
        'Լրացնել "Պայմ.թղթային N" դաշտը
        Call Rekvizit_Fill("Document",tabN,"General","PPRCODE",AdditionalTab.AgreemPaperN)
        'Լրացնել "Փակման ամսաթիվ" դաշտը
        Call Check_ReadOnly("Document",tabN,"General","DATECLOSE",True)
        'Լրացնել "Լրիվ փակված" դաշտը
        Call Check_ReadOnly("Document",tabN,"General","CANCELED",True)
        'Լրացնել "Սուբեկտիվ Դասակարգված" դաշտը
        Call Rekvizit_Fill("Document",tabN,"CheckBox","SUBJRISK",AdditionalTab.SubjectiveCategorized)
    End If
End Sub
        
'------------------------------------------------------------------------------------
'Լրացնել "Ծանուցում/Notification" Tab-ի արժեքները
'------------------------------------------------------------------------------------
Sub Fill_NotificationTab(NotificationTab,tabN)

    If NotificationTab.FillTab Then
        Log.Message "Fill Notification Tab",,,MessageColor
        'Լրացնել "Ծանուցման ձև" դաշտը
        Call Rekvizit_Fill("Document",tabN,"General","NTFMODE",NotificationTab.NotifyMode)
        'Լրացնել "Ծանուցման ուղղարկման հասցե" դաշտը
        Call Rekvizit_Fill("Document",tabN,"General","SENDSTMADRS",NotificationTab.SendNotificationAddress)
    End If
End Sub
   
'------------------------------------------------------------------------------------
'Լրացնել "Վարկ. ռեգ/LoanReg" Tab-ի արժեքները
'------------------------------------------------------------------------------------
Sub Fill_LoanRegTab(LoanRegTab,tabN)

    If LoanRegTab.FillTab Then
        Log.Message "Fill Loan Reg Tab",,,MessageColor
        'Լրացնել "Հաշվառել վարկային ռեգիստրում" դաշտը
        Call Rekvizit_Fill("Document",tabN,"CheckBox","PUTINLR",LoanRegTab.AccumulateInLoanReg)
        
        If Get_Rekvizit_Value("Document",tabN,"CheckBox","PUTINLR") = "0" Then
            'Ստուգում "Վարկային ռեգիստրի կոդ" դաշտի խմբագրելիությունը 
            Call Check_ReadOnly("Document",tabN,"General","LRCODE",True)
            'Ստուգում "Փոփոխ.քանակ" դաշտի խմբագրելիությունը 
            Call Check_ReadOnly("Document",tabN,"General","CHGSCNT",True)
            'Ստուգում "Լրացուցիչ ինֆորմացիա" դաշտի խմբագրելիությունը 
            Call Check_ReadOnly("Document",tabN,"General","OTHER",True)
            'Ստուգում "Գրավի գումար" դաշտի խմբագրելիությունը 
            Call Check_ReadOnly("Document",tabN,"General","LRMRTSUM",True)
            'Ստուգում "Գրավի առարկա" դաշտի խմբագրելիությունը 
            Call Check_ReadOnly("Document",tabN,"General","LRMRTOBJ",True)
            'Ստուգում "Գրավ(Լրացուցիչ ինֆորմացիա)" դաշտի խմբագրելիությունը 
            Call Check_ReadOnly("Document",tabN,"General","LRMRTOTHER",True)
        Else
            'Լրացնել "Վարկային ռեգիստրի կոդ" դաշտը
            Call Rekvizit_Fill("Document",tabN,"General","LRCODE",LoanRegTab.LoanRegisterCode)
            'Լրացնել "Փոփոխ.քանակ" դաշտը
            Call Rekvizit_Fill("Document",tabN,"General","CHGSCNT",LoanRegTab.CountOfChanges)
            'Լրացնել "Լրացուցիչ ինֆորմացիա" դաշտը
            Call Rekvizit_Fill("Document",tabN,"General","OTHER",LoanRegTab.AdditionalInformation)
            'Լրացնել "Գրավի գումար" դաշտը
            Call Rekvizit_Fill("Document",tabN,"General","LRMRTSUM",LoanRegTab.RemnantOfPladge)
            'Լրացնել "Գրավի առարկա" դաշտը
            Call Rekvizit_Fill("Document",tabN,"General","LRMRTOBJ",LoanRegTab.PladgeObject)
            'Լրացնել "Գրավ(Լրացուցիչ ինֆորմացիա)" դաշտը
            Call Rekvizit_Fill("Document",tabN,"General","LRMRTOTHER",LoanRegTab.PladgeAdditionalInfo)
        End If
        
        'Լրացնել "Գրավի արժույթ" դաշտը
        Call Rekvizit_Fill("Document",tabN,"General","LRMRTCUR",LoanRegTab.PledgeCurrency)
        'Լրացնել "Գրավի առարկա ARCA" դաշտը
        Call Rekvizit_Fill("Document",tabN,"General","ACRANOTE",LoanRegTab.PladgeObjectArca)
        'Լրացնել "Չդասակարգող" դաշտը
        Call Rekvizit_Fill("Document",tabN,"CheckBox","NOTCLASS",LoanRegTab.NotClassifiable)
        
        'Ստուգում "ՎՌ կոդ(Նոր)" դաշտի խմբագրելիությունը 
        If Not Check_ReadOnly("Document",tabN,"General","NEWLRCODE",True) Then
            Call Rekvizit_Fill("Document",tabN,"General","NEWLRCODE",LoanRegTab.LRCodeNew)
        End If
        
        'Լրացնել "Պայմանագրի վերանայման պատճառ" դաշտը
        Call Rekvizit_Fill("Document",tabN,"General","REVISIONREASON",LoanRegTab.RevisionReason)
        'Լրացնել "Մարման աղբյուր" դաշտը
        Call Rekvizit_Fill("Document",tabN,"General","REPSOURCE",LoanRegTab.RepaymentSourse)
        'Լրացնել "Գրավի առարկա(նոր ՎՌ)" դաշտը
        Call Rekvizit_Fill("Document",tabN,"General","MORTSUBJECT",LoanRegTab.PledgeObjectNew)
        'Լրացնել "Ապահովված է այլ ապահովվածությամբ" դաշտը
        Call Rekvizit_Fill("Document",tabN,"CheckBox","OTHERCOLLATERAL",LoanRegTab.GuaranteedByOtherCallateral)
    End If
End Sub

'--------------------------------------------------------------------------------------
'Տոկոսների հաշվարկում գործողության կատարում
'--------------------------------------------------------------------------------------
'CalcPercents - օբեկտի անունը
'beforeTerm - true եթե կատարվել է ժամկետից շուտ մարում
Function CalculatePercents(CalcPercents,ExpectedMessage,beforeTerm)
    
    Set wMDIClient = wMainForm.Window("MDIClient", "", 1)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_Opers & "|" & c_Interests & "|" & c_PrcAccruing)
    wMDIClient.Refresh

    Set DocForm = wMDIClient.WaitVBObject("frmASDocForm", delay_middle)
    
    If DocForm.Exists Then
        'ISN-ի վերագրում փոփոխականին
        CalcPercents.Isn = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
    
        Call Fill_CalculatePercents(CalcPercents)
        If beforeTerm Then
            Call MessageExists(2,ExpectedMessage)
            p1.vbObject("frmAsMsgBox").vbObject("cmdButton").click()
        End If
    Else
        Log.Error "Can Not Open Rc(CalculatePercents/Տոկոսների հաշվարկում) Window",,,ErrorColor        
    End If
    BuiltIn.Delay(1500)
    If DocForm.Exists Then
        Log.Error "Can Not Close Rc(CalculatePercents/Տոկոսների հաշվարկում) Window",,,ErrorColor
    End If
End Function

'---------------------------------------------------------------------------------------------
' Լրացնել "Գործողություններ/Տոկոսների հաշվարկում" պատուհանի դաշտերը
'---------------------------------------------------------------------------------------------
Sub Fill_CalculatePercents(CalcPercents)

    'Ստուգում "Պայմանագրի N" դաշտի խմբագրելիությունը և արժեքը
    Call Check_ReadOnly("Document",1,"General","CODE",True) 
    Call Compare_Two_Values("Պայմանագրի N",Get_Rekvizit_Value("Document",1,"Mask","CODE"),CalcPercents.ExpectedAgreementN)
    'Լրացնել "Հաշվարկման ամսաթիվ" դաշտը
    Call Rekvizit_Fill("Document", 1, "General", "DATECHARGE", CalcPercents.CalculationDate)
    'Լրացնել "Գործողության ամսաթիվ" դաշտը
    Call Rekvizit_Fill("Document", 1, "General", "DATE", CalcPercents.OperationDate)
    'Լրացնել "Ժամկետանց գումարի տույժ" և "Որից դուրս գրված" դաշտերը
    Call Rekvizit_Fill("Document", 1, "General", "SUMAGRPEN", CalcPercents.FineOnPastDueSum &"[Tab]"&CalcPercents.FineOnPastDueSum2 )
    
    'Ստուգում "Ընդամենը տույժ"  և "Որից դուրս գրված" դաշտերի խմբագրելիությունը և արժեքները
    Call Check_ReadOnly("Document",1,"Course1","SUMALLPEN",True) 
    Call Check_ReadOnly("Document",1,"Course2","SUMALLPEN",True) 
    Call Compare_Two_Values("Ընդամենը տույժ",Get_Rekvizit_Value("Document",1,"Course","SUMALLPEN"),CalcPercents.TotalPenalty & "/" & CalcPercents.TotalPenalty2)
    
    'Լրացնել "Մեկանաբանություն" դաշտը
    Call Rekvizit_Fill("Document",1,"General","COMMENT","![End][Del]" & CalcPercents.Comment)
    'Լրացնել "Գրասենյակ" դաշտերը
    Call Rekvizit_Fill("Document",1,"General","ACSBRANCH",CalcPercents.Division)
    'Լրացնել "´աժին" դաշտերը
    Call Rekvizit_Fill("Document",1,"General","ACSDEPART",CalcPercents.Department)
    'ê»ÕÙ»É "Կատարել" կոճակÁ
     Call ClickCmdButton(1, "Î³ï³ñ»É")
End Sub

'------------------------------------------------------------------------------------
' "Գերածախս" գործողության կատարում
'------------------------------------------------------------------------------------
Sub Give_Overlimit(RcOptionOverlimit)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_Opers & "|" & c_GiveAndBack & "|" & c_Overlimit)
    wMDIClient.Refresh
    
    Set DocForm = wMDIClient.WaitVBObject("frmASDocForm", 2000)
    
    If DocForm.Exists Then
        'ISN-ի վերագրում փոփոխականին
        RcOptionOverlimit.Isn = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
        Call FillRcOverlimit(RcOptionOverlimit)
    Else
        Log.Error "Can Not Open Rc(OptionOverlimit/Գերածախս) Window",,,ErrorColor
    End If
    BuiltIn.Delay(1500)
    If DocForm.Exists Then
        Log.Error "Can Not Close Rc(OptionOverlimit/Գերածախս) Window",,,ErrorColor
    End If
End Sub

'---------------------------------------------------------------------------------------------
' Լրացնել "Գործողություններ/Տրամադրում/մարում/գերածախս" պատուհանի դաշտերը
'---------------------------------------------------------------------------------------------
Sub FillRcOverlimit(RcOverlimit)

    'Ստուգում "Պայմանագրի N" դաշտի խմբագրելիությունը և արժեքը
    Call Check_ReadOnly("Document",1,"General","CODE",True) 
    Call Compare_Two_Values("Պայմանագրի N",Get_Rekvizit_Value("Document",1,"Mask","CODE"),RcOverlimit.ExpectedAgreementN)
     
    'Լրացնել "Ամսաթիվ" դաշտը
    Call Rekvizit_Fill("Document",1,"General","DATE", RcOverlimit.Date)
    'Լրացնել "Գումար" դաշտը
    Call Rekvizit_Fill("Document",1,"General","SUMMA",RcOverlimit.Sum)
    'Լրացնել "Կանխիկ/Անկանխիկ" դաշտը
    Call Rekvizit_Fill("Document",1,"General","CASHORNO",RcOverlimit.CashOrNo)
    
    If Get_Rekvizit_Value("Document",1,"Mask","CASHORNO") = "2" Then
        'Լրացնել "Հաշիվ" դաշտը
        Call Rekvizit_Fill("Document",1,"General","ACCCORR", RcOverlimit.CalcAcc)
    End If
    
    'Լրացնել "Մեկանաբանություն" դաշտը
    Call Rekvizit_Fill("Document",1,"General","COMMENT", "![End][Del]" & RcOverlimit.Comment)
    'Լրացնել "Գրասենյակ/բաժին" դաշտերը
    Call Rekvizit_Fill("Document",1,"General","ACSBRANCH",RcOverlimit.Division & "[Tab]" & RcOverlimit.Department)
    
    Call ClickCmdButton(1, "Î³ï³ñ»É")
    BuiltIn.Delay(1000)
    
    Call MessageExists(2,"²í³ñï»±É ·áñÍáÕáõÃÛáõÝÁ ³ÝÙÇç³å»ëª ÃÕÃ³ÏóáõÃÛáõÝÁ" & vbNewLine & "Ï³ï³ñ»Éáí Ñ³ßíÇ Ñ»ï." & vbNewLine &""& vbNewLine &"      ² Ú à    -    ÃÕÃ³ÏóáõÃÛáõÝ Ñ³ßíÇ Ñ»ï" & vbNewLine & "      à â        -    ÷³ëï³ÃÕÃÇ áõÕ³ñÏáõÙ ÃÕÃ³å³Ý³ÏÝ»ñ")
    Call ClickCmdButton(5, "²Ûá")
End Sub


Function GetAccountIsnOverlimit()
    Dim Pttel
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_Folders & "|" & c_AgrFolder)
    wMDIClient.Refresh

    Set Pttel = wMDIClient.WaitVBObject("frmPttel_2", delay_middle)
    
    If Pttel.Exists Then
        GetAccountIsnOverlimit = GetIsn()
        BuiltIn.Delay(1000)
        wMDIClient.VBObject("frmPttel_2").Close
    Else
        Log.Error "Can Not Թղթապանակներ/Պայմանագրի թղթապանակ Window",,,ErrorColor      
    End If    
    BuiltIn.Delay(1500)
    If Pttel.Exists Then
        Log.Error "Can Not Close Թղթապանակներ/Պայմանագրի թղթապանակ Window",,,ErrorColor
    End If
End Function

'---------------------------------------------------------------------------------------------
' AccountsWithOverlimit - "գերածախս ունեցող հաշիվներ" պատուհանի Լրացման Class
'---------------------------------------------------------------------------------------------
Class AccountsWithOverlimit
    Public Curr
    Public Client
    Public AccountMask
    Public Clientame
    Public AccountNote
    Public AccountNote2
    Public AccountNote3
    Public ClientNote
    Public ClientNote2
    Public ClientNote3
    Public Division
    Public Department
    Public AccessType
    Public ShowClientsProperties
    Public ShowNotesOfClient
    Public ShowNotesOfAccount
    
    Private Sub Class_Initialize
        Curr = ""
        Client = ""
        AccountMask = ""
        Clientame = ""
        AccountNote = ""
        AccountNote2 = ""
        AccountNote3 = ""
        ClientNote = ""
        ClientNote2 = ""
        ClientNote3 = ""
        Division = ""
        Department = ""
        AccessType = ""
        ShowClientsProperties = 0
        ShowNotesOfClient = 0
        ShowNotesOfAccount = 0
    End Sub  
End Class

Function New_AccountsWithOverlimit()
    Set New_AccountsWithOverlimit = NEW AccountsWithOverlimit    
End Function

'--------------------------------------------------------------------------------------
'"աջ կլիկ - Գերածախսի բացում (խմբ.)" գործողության կատարում տրված ամսաթվով
'Ֆունկցիան վերադարձնում է "Գերածախս ունեցող հաշիվի" isn - ը
'--------------------------------------------------------------------------------------
Function OpenOverimitFromAccount(Date)
    Dim DocForm,Isn
    Set DocForm = wMDIClient.VBObject("frmPttel")
    If WaitForPttel("frmPttel") Then
        If DocForm.VBObject("tdbgView").ApproxCount <> 0 Then
            Isn = GetIsn()
            DocForm.Keys("[Ins]")
            Call wMainForm.MainMenu.Click(c_AllActions)
            Call wMainForm.PopupMenu.Click(c_OpenOverlimit)
            'Լրացնել "Ամսաթիվ" դաշտը
            Call Rekvizit_Fill("Dialog", 1, "General", "DATE", Date)
            Call ClickCmdButton(2, "Î³ï³ñ»É")
            BuiltIn.Delay(3000)
            OpenOverimitFromAccount = Isn
        Else 
            Log.Error "Տողը չի գտնվել Գերածախս ունեցող հաշիվներ-ում" ,,,ErrorColor
        End If  
        BuiltIn.Delay(2000)
        wMDIClient.WaitVBObject("frmPttel",delay_middle).Close
    Else
        Log.Error "Can Not Open Գերածախս ունեցող հաշիվներ Window",,,ErrorColor      
    End If  
    BuiltIn.Delay(delay_middle)
    If DocForm.Exists Then
        Log.Error "Can Not Close Գերածախս ունեցող հաշիվներ Window",,,ErrorColor
    End If
End Function

'------------------------------------------------------------------------------------
' Լրացնել (գերածախս ունեցող հաշիվներ) ֆիլտրը
'------------------------------------------------------------------------------------
Sub Fill_AccWithOverlimit(AccWithOverlimit)

    'Լրացնել "Արժույթ" երկու դաշտերը
    Call Rekvizit_Fill("Dialog", 1, "General", "CUR", AccWithOverlimit.Curr)
    'Լրացնել "Հաճախորդ" երկու դաշտերը
    Call Rekvizit_Fill("Dialog", 1, "General", "CLIENT", AccWithOverlimit.Client)
    'Լրացնել "Հաշվարկային հաշիվ" երկու դաշտերը
    Call Rekvizit_Fill("Dialog", 1, "General", "ACCMASK", AccWithOverlimit.AccountMask)
    'Լրացնել "Հաճախորդի անվանում" երկու դաշտերը
    Call Rekvizit_Fill("Dialog", 1, "General", "CLNAME", AccWithOverlimit.Clientame)
    'Լրացնել "Հաշվի նշում" երկու դաշտերը
    Call Rekvizit_Fill("Dialog", 1, "General", "ACCNOTE", AccWithOverlimit.AccountNote)
    'Լրացնել "Հաշվի նշում2" երկու դաշտերը
    Call Rekvizit_Fill("Dialog", 1, "General", "ACCNOTE2", AccWithOverlimit.AccountNote2)
    'Լրացնել "Հաշվի նշում3" երկու դաշտերը
    Call Rekvizit_Fill("Dialog", 1, "General", "ACCNOTE3", AccWithOverlimit.AccountNote3)
    'Լրացնել "Հաճախորդի նշում" երկու դաշտերը
    Call Rekvizit_Fill("Dialog", 1, "General", "CLINOTE", AccWithOverlimit.ClientNote)
    'Լրացնել "Հաճախորդի նշում2" երկու դաշտերը
    Call Rekvizit_Fill("Dialog", 1, "General", "CLINOTE2", AccWithOverlimit.ClientNote2)
    'Լրացնել "Հաճախորդի նշում3" երկու դաշտերը
    Call Rekvizit_Fill("Dialog", 1, "General", "CLINOTE3", AccWithOverlimit.ClientNote3)
    'Լրացնել "Գրասենյակ" երկու դաշտերը
    Call Rekvizit_Fill("Dialog", 1, "General", "ACSBRANCH", AccWithOverlimit.Division)
    'Լրացնել "Բաժին" երկու դաշտերը
    Call Rekvizit_Fill("Dialog", 1, "General", "ACSDEPART", AccWithOverlimit.Department)
    'Լրացնել "Հասան-ն տիպ" երկու դաշտերը
    Call Rekvizit_Fill("Dialog", 1, "General", "ACSTYPE", AccWithOverlimit.AccessType)
    'Լրացնել "Ցույց տալ հաճախորդնորի հատկանիշները" դաշտը
    Call Rekvizit_Fill("Dialog",1,"CheckBox","SHOWCLI",AccWithOverlimit.ShowClientsProperties)
    'Լրացնել "Ցույց տալ հաճախորդի նշումները" դաշտը
    Call Rekvizit_Fill("Dialog",1,"CheckBox","SHOWCLINOTES",AccWithOverlimit.ShowNotesOfClient)
    'Լրացնել "Ցույց տալ հաշիվների նշումները" դաշտը
    Call Rekvizit_Fill("Dialog",1,"CheckBox","SHOWACCNOTES",AccWithOverlimit.ShowNotesOfAccount)
    
    Call ClickCmdButton(2, "Î³ï³ñ»É")
End Sub

'------------------------------------------------------------------------------------
' Գերածախս ունեցող հաշիվներ թղթապանակում փաստատթղթի առկայության ստուգում
'------------------------------------------------------------------------------------
Function ExistsAccWithOverlimit_Filter_Fill(AccWithOverlimit, RowCount)
    Call wTreeView.DblClickItem("|¶»ñ³Í³Ëë|¶»ñ³Í³Ëë áõÝ»óáÕ Ñ³ßÇíÝ»ñ|")
    BuiltIn.Delay(delay_middle)
    Call Fill_AccWithOverlimit(AccWithOverlimit)
    Set DocForm = wMDIClient.VBObject("frmPttel")
    
    If WaitForPttel("frmPttel") Then
        wMDIClient.Refresh
        If DocForm.vbObject("tdbgView").ApproxCount = RowCount Then
            Log.Message "Row count of AccWithOverlimit is right",,,MessageColor
            ExistsAccWithOverlimit_Filter_Fill = True
        Else
            Log.Error "Row count of AccWithOverlimit is not right",,,ErrorColor
            ExistsAccWithOverlimit_Filter_Fill = False
        End If
    Else
        Log.Error "Can Not Open Գերածախս ունեցող հաշիվներ Window",,,ErrorColor      
    End If     
End Function 


'------------------------------------------------------------------------------------
' Գլխավոր հաշվապահ/Աշխատանքային փաստաթղթեր թղթապանակից կատարել հաշվառել գործողություն
' Վարադարձնում է Փաստաթղթի N - ը
'------------------------------------------------------------------------------------
Function ToCountPayment(ActionName,Date) 

    Dim PttelExists
    wTreeView.DblClickItem("|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ") 
    'Լրացնել "Ամսաթիվ" դաշտը
    Call Rekvizit_Fill("Dialog",1,"General","PERN", Date)
    Call Rekvizit_Fill("Dialog",1,"General","PERK", Date)
    Call ClickCmdButton(2, "Î³ï³ñ»É")
                   
    
    Set PttelExists = wMDIClient.VBObject("frmPttel")
    
    If WaitForPttel("frmPttel") Then
        If PttelExists.VBObject("tdbgView").ApproxCount <> 0 Then     
        
            BuiltIn.Delay(1000)
            Call wMainForm.MainMenu.Click(c_AllActions)
            Call wMainForm.PopupMenu.Click(c_View)
            
            Set DocForm = wMDIClient.vbObject("frmASDocForm")
            If DocForm.Exists Then
                'Վերցնեում է Փաստաթղթի N - ը
                ToCountPayment = Get_Rekvizit_Value("Document",1,"General","DOCNUM")
                BuiltIn.Delay(2000)
                DocForm.Close
            Else
                Log.Error "Can Not Open Window" ,,,ErrorColor    
            End If
            
            BuiltIn.Delay(1000)
            Call wMainForm.MainMenu.Click(c_AllActions)
            Call wMainForm.PopupMenu.Click(ActionName)
            
            Select Case ActionName
              Case c_ToCount
                  BuiltIn.Delay(2000)
                  Call MessageExists(2,"Ð³ßí³é»É")
                  BuiltIn.Delay(2000)
                  Call ClickCmdButton(5, "²Ûá")
              Case c_SendToVer
                  BuiltIn.Delay(2000)
                  Call ClickCmdButton(2, "Î³ï³ñ»É")
              Case c_ToConfirm
                  BuiltIn.Delay(2000)
                  Call ClickCmdButton(1, "Ð³ëï³ï»É")
             End Select   
        Else 
            Log.Error "Տողը չի գտնվել Աշխատանքային փաստաթղթեր թղթապանակում" ,,,ErrorColor
        End If  
        BuiltIn.Delay(1500)
        wMDIClient.WaitVBObject("frmPttel",delay_middle).Close
     Else
        Log.Error "Can Not Open Աշխատանքային փաստաթղթեր Window",,,ErrorColor      
     End If     
     If PttelExists.Exists Then
        Log.Error "Can Not Close Աշխատանքային փաստաթղթեր Window",,,ErrorColor
     End If
End Function

'------------------------------------------------------------------------------------
' Քարտային վճարումներ թղթապանակից "կատարել քարտային հաշվից" գործողությունը
'------------------------------------------------------------------------------------
Sub Card_Payment(Date) 

    Call wTreeView.DblClickItem(("|äÉ³ëïÇÏ ù³ñï»ñÇ ²Þî (SV)|ÂÕÃ³å³Ý³ÏÝ»ñ|ø³ñï³ÛÇÝ í×³ñáõÙÝ»ñ"))
    BuiltIn.Delay(1500)
    
    Set DocForm = wMDIClient.VBObject("frmPttel")
    
    If WaitForPttel("frmPttel") Then
        If DocForm.VBObject("tdbgView").ApproxCount <> 0 Then
            Call SearchInPttel("frmPttel",0, "20/11/20")
            BuiltIn.Delay(1000)
            Call wMainForm.MainMenu.Click(c_AllActions)
            Call wMainForm.PopupMenu.Click(c_MakeFromCardAccount)
            BuiltIn.Delay(2000)
            'Լրացնել "Ամսաթիվ" դաշտը
            Call Rekvizit_Fill("Document",1,"General","DATE", Date)
            Call ClickCmdButton(1, "Î³ï³ñ»É")
        Else 
            Log.Error " համարի պայմանագիրը չի գտնվել Քարտային վճարումներ թղթապանակում" ,,,ErrorColor
        End If  
        BuiltIn.Delay(1500)
        wMDIClient.WaitVBObject("frmPttel",delay_middle).Close
     Else
        Log.Error "Can Not Open Քարտային վճարումներ Window",,,ErrorColor      
     End If     
     If DocForm.Exists Then
        Log.Error "Can Not Close Քարտային վճարումներ Window",,,ErrorColor
     End If
End Sub

'------------------------------------------------------------------------------------
' Պլաստիկ քարտեր թղթապանակից "Քարտային վճարում" գործողության կատարում
'------------------------------------------------------------------------------------
Function Card_PaymentAction(Date,CardType,Amount,Payment)

    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_CardPayments)
    wMDIClient.Refresh
    
    Set DocForm = wMDIClient.WaitVBObject("frmASDocForm", 2000)
    
    If DocForm.Exists Then
        'ISN-ի վերագրում փոփոխականին
        Card_PaymentAction = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
        
        'Լրացնել "Ամսաթիվ" դաշտը
        Call Rekvizit_Fill("Document",1,"General","DATE", Date)
        'Լրացնել "Քարտային վճարման տիպ" դաշտը
        Call Rekvizit_Fill("Document",1,"General","CRDFEETP",CardType)
        'Լրացնել "Վճարման գումար" դաշտը
        Call Rekvizit_Fill("Document",1,"General","FEE",Amount)
        'Լրացնել "Վճարման եղանակ" դաշտը
        Call Rekvizit_Fill("Document",1,"General","MNTFEETP",Payment)
        
        Call ClickCmdButton(1, "Î³ï³ñ»É")
    Else
        Log.Error "Can Not Open Card_Payments Window",,,ErrorColor
    End If
    BuiltIn.Delay(2000)
    If DocForm.Exists Then
        Log.Error "Can Not Close Card_Payments Window",,,ErrorColor
    End If
End Function