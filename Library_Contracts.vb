'USEUNIT Library_Common 
'USEUNIT Library_Colour
'USEUNIT Constants
'USEUNIT Mortgage_Library
Option Explicit

Dim DocForm, securitiesRow, gridRows, clients_count, export_count

'---------------------------------------------------------------------------------------------
'Հաստատվող փաստաթղթեր ֆիլտր
'---------------------------------------------------------------------------------------------
Class VerifyContract
    Public ConFirmationGroup
				Public TermsStatesExists
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
							TermsStatesExists = false
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

Function New_VerifyContract()
    Set New_VerifyContract = NEW VerifyContract      
End Function

'------------------------------------------------------------------------------------
' Լրացնել (Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ) ֆիլտրը
'------------------------------------------------------------------------------------
Sub Fill_Verify(VerifyFilter1)
    'Լրացնել "Հաստատման խումբ" երկու դաշտերը
    Call Rekvizit_Fill("Dialog", 1, "General", "SSCONFTP", VerifyFilter1.ConFirmationGroup)
    If VerifyFilter1.TermsStatesExists Then
		    If Not Check_ReadOnly("Dialog",1,"General","SSAGRMEDR",True) Then
		        'Լրացնել "Պայմաններ-վիճակներ" երկու դաշտերը
		        Call Rekvizit_Fill("Dialog", 1, "General", "SSAGRMEDR", VerifyFilter1.TermsStates)
		    End If    
				End If
    If Not Check_ReadOnly("Dialog",1,"General","SSAGRMOPR",True) Then    
        'Լրացնել "Պայմանագրերի գործողություններ" երկու դաշտերը
        Call Rekvizit_Fill("Dialog", 1, "General", "SSAGRMOPR", VerifyFilter1.AgreementOperations)
    End If
    'Լրացնել "Պայմանագրի N" երկու դաշտերը
    Call Rekvizit_Fill("Dialog", 1, "General", "NUM", VerifyFilter1.AgreementN)
    'Լրացնել "Պայմ.պղպային N" երկու դաշտերը
    Call Rekvizit_Fill("Dialog", 1, "General", "PPRCODE", VerifyFilter1.AgreemPaperN)
    'Լրացնել "Արժույթ" երկու դաշտերը
    Call Rekvizit_Fill("Dialog", 1, "General", "CUR", VerifyFilter1.Curr)
    'Լրացնել "Հաճախորդ" երկու դաշտերը
    Call Rekvizit_Fill("Dialog", 1, "General", "CLIENT", VerifyFilter1.Client)
    'Լրացնել "Հաճախորդի անվանում" երկու դաշտերը
    Call Rekvizit_Fill("Dialog", 1, "General", "NAME", VerifyFilter1.ClientName)
    'Լրացնել "Նշում" երկու դաշտերը
    Call Rekvizit_Fill("Dialog", 1, "General", "NOTE", VerifyFilter1.Note)
    'Լրացնել "Նշում 2" երկու դաշտերը
    Call Rekvizit_Fill("Dialog", 1, "General", "NOTE2", VerifyFilter1.Note2)
    'Լրացնել "Նշում 3" երկու դաշտերը
    Call Rekvizit_Fill("Dialog", 1, "General", "NOTE3", VerifyFilter1.Note3)
    'Լրացնել "Կատարողներ" երկու դաշտերը
    Call Rekvizit_Fill("Dialog", 1, "General", "USER", VerifyFilter1.Executors)
    'Լրացնել "Գրասենյակ" երկու դաշտերը
    Call Rekvizit_Fill("Dialog", 1, "General", "ACSBRANCH", VerifyFilter1.Division)
    'Լրացնել "Բաժին" երկու դաշտերը
    Call Rekvizit_Fill("Dialog", 1, "General", "ACSDEPART", VerifyFilter1.Department)
    'Լրացնել "Հասան-ն տիպ" երկու դաշտերը
    Call Rekvizit_Fill("Dialog", 1, "General", "ACSTYPE", VerifyFilter1.AccessType)
    
    Call ClickCmdButton(2, "Î³ï³ñ»É")
End Sub
'------------------------------------------------------------------------------------
' Հաստատում է պայմանագիրը(Ð³ëï³ïíáÕ ÷³ëï³ÃÕÃ»ñ I)
'------------------------------------------------------------------------------------
Sub Verify_Contract(FolderPath,VerifyFilter1) 

    Call wTreeView.DblClickItem(FolderPath)
    BuiltIn.Delay(1500)
    Call Fill_Verify(VerifyFilter1)
    Set DocForm = wMDIClient.VBObject("frmPttel")
    
    If WaitForPttel("frmPttel") Then
        If DocForm.VBObject("tdbgView").ApproxCount <> 0 Then
            Call wMainForm.MainMenu.Click(c_AllActions)
            Call wMainForm.PopupMenu.Click(c_ToConfirm)
            Call ClickCmdButton(1, "Ð³ëï³ï»É")
        Else 
            Log.Error VerifyFilter1.AgreementN & " համարի պայմանագիրը չի գտնվել Հաստատվող փաստաթղթեր 1-ում" ,,,ErrorColor
        End If  
        BuiltIn.Delay(1500)
        wMDIClient.WaitVBObject("frmPttel",delay_middle).Close
     Else
        Log.Error "Can Not Open Հաստատվող փաստաթղթեր 1 Window",,,ErrorColor      
     End If     
     If DocForm.Exists Then
        Log.Error "Can Not Close Հաստատվող փաստաթղթեր 1 Window",,,ErrorColor
     End If
End Sub

'------------------------------------------------------------------------------------
' Պայմանագիրը ուղարկում է հաստատման
'------------------------------------------------------------------------------------
Sub SendToVerify_Contrct(messageType, winType, button)
    BuiltIn.Delay(2000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_SendToVer)
    
    Call MessageExists(messageType, "àõÕ³ñÏ»É Ñ³ëï³ïÙ³Ý")
    Call ClickCmdButton(winType, button)
    
    BuiltIn.Delay(2000)
    wMDIClient.VBObject("frmPttel").Close
End Sub

'---------------------------------------------------------------------------------------------
' Contracts Doc - "Պայմանագրեր" filter-ի Class
'---------------------------------------------------------------------------------------------
Class ContractsFilter
		Public LeaseAgree
		Public DateFill
		Public Date
		Public AgreementLevelExist
		Public AgreementLevel
		Public AgreementSpecies
		Public Mortgage
		Public MortgageType
		Public AgreementN
		Public AgreemPaperN
		Public CreditCodeExist
		Public CreditCode
		Public LRCodeNewExist
		Public LRCodeNew
		Public Curr
		Public Client
		Public ClientName
		Public GroupExists
		Public Group
		Public Note
		Public Note2
		Public Note3
		Public ShowAccounts
		Public ShowOnlyLinearExists
		Public ShowOnlyLinear
		Public ShowClosed
		Public ShowClientData
		Public NotFullClosedExist
		Public ShowNotFullClosedAgr
		Public Division
		Public Department
		Public AccessType
		Private Sub Class_Initialize
				LeaseAgree = false
				DateFill = false
				Date = ""
				AgreementLevelExist = true
		  AgreementLevel = "1"
		  AgreementSpecies = ""
				Mortgage = false
				MortgageType = ""
		  AgreementN = ""
		  AgreemPaperN = ""
				CreditCodeExist = false
		  CreditCode = ""
				LRCodeNewExist = false
		  LRCodeNew = ""
		  Curr = ""
		  Client = ""
		  ClientName = ""
				GroupExists = false
				Group = ""
		  Note = ""
		  Note2 = ""
		  Note3 = ""
		  ShowAccounts = 0
				ShowOnlyLinearExists = false
				ShowOnlyLinear = 0
		  ShowClosed = 0
				ShowClientData = 0
				NotFullClosedExist = false
		  ShowNotFullClosedAgr = 0
		  Division = ""
		  Department = ""
		  AccessType = ""
		End Sub  
End Class

Function New_ContractOverlimit()
    Set New_ContractOverlimit = NEW ContractsFilter      
End Function

Function New_ContractsFilter()
    Set New_ContractsFilter = NEW ContractsFilter      
End Function

'------------------------------------------------------------------------------------
'Լրացնել "Պայմանագրեր" Filter-ի արժեքները
'------------------------------------------------------------------------------------
Sub Fill_ContractsFilter(Contract)
		'Լրացնում է "Ամսաթիվ" դաշտը
		if Contract.DateFill then 
				Call Rekvizit_Fill("Dialog", 1, "General", "RDATE",  "![End]" & "[Del]" & Contract.Date)
		end if
		if Contract.AgreementLevelExist then
				'Լրացնում է "ä³ÛÙ³Ý³·ñÇ Ù³Ï³ñ¹³Ï" դաշտը
				Call Rekvizit_Fill("Dialog", 1, "General", "LEVEL",  "![End]" & "[Del]" & Contract.AgreementLevel)
		end if
		'Լրացնում է "ä³ÛÙ³Ý³·ñÇ ï»ë³Ï" դաշտը
		Call Rekvizit_Fill("Dialog", 1, "General", "AGRKIND", Contract.AgreementSpecies)
		if Contract.Mortgage then
				'Լրացնում է "ä³ÛÙ³Ý³·ñÇ ïÇå" դաշտը
				Call Rekvizit_Fill("Dialog", 1, "General", "AGRKIND", Contract.MortgageType)
		end if
		'Լրացնում է "ä³ÛÙ³Ý³·ñÇ N" դաշտը
		if Contract.LeaseAgree then
				Call Rekvizit_Fill("Dialog", 1, "General", "LCNUM", Contract.AgreementN)
		elseif Contract.Mortgage then
				Call Rekvizit_Fill("Dialog", 1, "General", "AGRNUM", Contract.AgreementN)
		else
				Call Rekvizit_Fill("Dialog", 1, "General", "NUM", Contract.AgreementN)
		end if
		'Լրացնում է "Պայմ.թղթային N" դաշտը
		Call Rekvizit_Fill("Dialog", 1, "General", "PPRCODE", Contract.AgreemPaperN)
		if Contract.CreditCodeExist then
				'Լրացնում է "Վարկային կոդ" դաշտը
				Call Rekvizit_Fill("Dialog", 1, "General", "CRDTCODE", Contract.CreditCode)
		end if
		if Contract.LRCodeNewExist then
				'Լրացնում է "ՎՌ կոդ(Նոր)" դաշտը
				Call Rekvizit_Fill("Dialog", 1, "General", "NEWLRCODE", Contract.LRCodeNew)
		end if
		'Լրացնում է "Արժույթ" դաշտը
		Call Rekvizit_Fill("Dialog", 1, "General", "CUR", Contract.Curr)
		'Լրացնում է "Հաճախորդ" դաշտը
		Call Rekvizit_Fill("Dialog", 1, "General", "CLIENT", Contract.Client)
		'Լրացնում է "Հաճախորդի անվանում" դաշտը
		Call Rekvizit_Fill("Dialog", 1, "General", "NAME", Contract.ClientName)
		'Լրացնում է "Խումբ/փուլ" դաշտը
		if Contract.GroupExists then
		  Call Rekvizit_Fill("Dialog", 1, "General", "CRDGROUPMASK", Contract.Group)
		end if
		'Լրացնում է "Նշում" դաշտը
		Call Rekvizit_Fill("Dialog", 1, "General", "NOTE", Contract.Note)
		'Լրացնում է "Նշում2" դաշտը
		Call Rekvizit_Fill("Dialog", 1, "General", "NOTE2", Contract.Note2)
		'Լրացնում է "Նշում3" դաշտը
		Call Rekvizit_Fill("Dialog", 1, "General", "NOTE3", Contract.Note3)
		'Լրացնում է "Ցույց տալ հաշիվները" դաշտը
		Call Rekvizit_Fill("Dialog", 1, "CheckBox", "SHOWACCS", Contract.ShowAccounts)
		if Contract.ShowOnlyLinearExists then 
				'Լրացնում է "Ցույց տալ միայն գծայինները" դաշտը
		  Call Rekvizit_Fill("Dialog", 1, "CheckBox", "SHOWLINE", Contract.ShowOnlyLinear)
		end if 
		'Լրացնում է "Ցույց տալ փակվածները" դաշտը
		if Contract.LeaseAgree then
				Call Rekvizit_Fill("Dialog", 1, "CheckBox", "SHOWCLOSED", Contract.ShowClosed)
				Call Rekvizit_Fill("Dialog", 1, "CheckBox", "SHOWCLIINFO", Contract.ShowClientData)
		else 
				Call Rekvizit_Fill("Dialog", 1, "CheckBox", "CLOSE", Contract.ShowClosed)
		end if 
		if Contract.NotFullClosedExist then
				'Լրացնում է "Ցույց տալ ոչ լրիվ փակվածները" դաշտը
				Call Rekvizit_Fill("Dialog", 1, "CheckBox", "NOTFULLCLOSE", Contract.ShowNotFullClosedAgr)
		end if
		'Լրացնում է "Գրասենյակ" դաշտը
		Call Rekvizit_Fill("Dialog", 1, "General", "ACSBRANCH", Contract.Division)
		'Լրացնում է "Բաժին" դաշտը
		Call Rekvizit_Fill("Dialog", 1, "General", "ACSDEPART", Contract.Department)
		'Լրացնում է "Հասան-ն տիպ" դաշտը
		Call Rekvizit_Fill("Dialog", 1, "General", "ACSTYPE", Contract.AccessType)
				
		'Սեղմել Կատարել կոճակը
		Call ClickCmdButton(2, "Î³ï³ñ»É")
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''GoTo_Contracts''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պայմանագրեր թղթապանակ մուտք գործելու պրոցեդուրա
'folderName - գտնբելու ճանապարհը
'Contract - պատուհանի լրացման կլաս
Sub GoTo_Contracts(folderName, Contract)
		wTreeView.DblClickItem(folderName & "ä³ÛÙ³Ý³·ñ»ñ")
		if p1.WaitVBObject("frmAsUstPar", 3000).Exists then
				Call Fill_ContractsFilter(Contract)
		else 
				Log.Error "Can't open frmAsUstPar widow.", "", pmNormal, ErrorColor
		end if
End	Sub

'---------------------------------------------------------------------------------------------
' RcStore  - "Պահուստավորում" պատուհանի Լրացման Class
'---------------------------------------------------------------------------------------------
Class RcStore
    Public Isn
    Public ExpectedAgreementN
    Public Date
    Public Provision
    Public UnProvision
    Public Comment
    Public Division
    Public Department
    
    Private Sub Class_Initialize
        Isn = ""
        ExpectedAgreementN = ""
        Date = ""
        Provision = ""
        UnProvision = ""
        Comment = ""
        Division = ""
        Department = ""
    End Sub  
End Class

Function New_RcStore()
    Set New_RcStore = NEW RcStore      
End Function

'---------------------------------------------------------------------------------------------
' CalculatePercents  - "Տոկոսների հաշվարկում" պատուհանի Լրացման Class
'---------------------------------------------------------------------------------------------
Class RcCalculatePercents
    Public Isn
    Public ExpectedAgreementN
    Public CalculationDate
    Public OperationDate
    Public FineOnPastDueSum
    Public FineOnPastDueSum2
    Public TotalPenalty
    Public TotalPenalty2
    Public Comment
    Public Division
    Public Department
    
    Private Sub Class_Initialize
        Isn = ""
        ExpectedAgreementN = ""
        CalculationDate = ""
        OperationDate = ""
        FineOnPastDueSum = ""
        FineOnPastDueSum2 = ""
        TotalPenalty = "0.00"
        TotalPenalty2 = "0.00"
        Comment = ""
        Division = ""
        Department = ""
    End Sub  
End Class

Function New_RcCalculatePercents()
    Set New_RcCalculatePercents = NEW RcCalculatePercents    
End Function

'---------------------------------------------------------------------------------------------
' OverlimitRepay - "Պարտքերի մարում" պատուհանի Լրացման Class
'---------------------------------------------------------------------------------------------
Class RcOverlimitRepay
    Public Isn
    Public ExpectedAgreementN
    Public ExpectedAgreementNComment
    Public Date
    Public RepaymentCurrency
    Public RepaymentCurrencyComment
    Public ExpectedBaseSum
    Public BaseSum
    Public AMD1
    Public ExpectedFineOnPastSum
    Public FineOnPastSum
    Public AMD2
    Public TotalAmount
    Public CashCashles
    Public ExchangeRate
    Public Per
    Public Account
    Public AccountComment
    Public AMDAccount
    Public Comment
    Public RemittanceInfo1
    Public RemittanceInfo2
    Public Division
    Public Department
    
    Private Sub Class_Initialize
        Isn = ""
        ExpectedAgreementN = ""
        ExpectedAgreementNComment = ""
        Date = ""
        RepaymentCurrency = ""
        RepaymentCurrencyComment = ""
        ExpectedBaseSum = ""
        BaseSum = ""
        AMD1 = "0.00"
        ExpectedFineOnPastSum = ""
        FineOnPastSum = ""
        AMD2 = "0.00"
        TotalAmount = ""
        CashCashles = ""
        ExchangeRate = "0"
        Per = "0"
        Account = ""
        AccountComment = ""
        AMDAccount = ""
        Comment = ""
        RemittanceInfo1 = ""
        RemittanceInfo2 = ""
        Division = ""
        Department = ""
    End Sub  
End Class

Function New_RcOverlimitRepay()
    Set New_RcOverlimitRepay = NEW RcOverlimitRepay    
End Function

'---------------------------------------------------------------------------------------------
' WriteOut - "¸áõñë ·ñáõÙ" պատուհանի Լրացման Class
'---------------------------------------------------------------------------------------------
Class RcWriteOut
    Public Isn
    Public ExpectedAgreementN
    Public Date
    Public ExpectedBaseSum
    Public BaseSum
    Public ExpectedFineOnPastSum
    Public FineOnPastSum
    Public TotalSum
    Public Comment
    Public Division
    Public Department
    
    Private Sub Class_Initialize
        Isn = ""
        ExpectedAgreementN = ""
        Date = ""
        ExpectedBaseSum = "0.00"
        BaseSum = ""
        ExpectedFineOnPastSum = "0.00"
        FineOnPastSum = ""
        TotalSum = ""
        Comment = ""
        Division = ""
        Department = ""
    End Sub  
End Class

Function New_RcWriteOut()
    Set New_RcWriteOut = NEW RcWriteOut   
End Function

'--------------------------------------------------------------------------------------
'"Պայմանագրի փակում" Պատուհանի Éñ³óáõÙ
'--------------------------------------------------------------------------------------
Sub CloseContract(Date)
    
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_AgrClose)
    wMDIClient.Refresh
    
    Set DocForm = AsBank.WaitVBObject("frmAsUstPar", delay_middle)
    
    If DocForm.Exists Then
        'Լրացնել "Ամսաթիվ" դաշտը
        Call Rekvizit_Fill("Dialog", 1, "General", "DATECLOSE", Date)
        Call ClickCmdButton(2, "Î³ï³ñ»É")
    Else
        Log.Error "Can Not Open Rc(CloseContract/Պայմանագրի փակում) Window",,,ErrorColor         
    End If    
    BuiltIn.Delay(1500)
    If DocForm.Exists Then
        Log.Error "Can Not Close Rc(CloseContract/Պայմանագրի փակում) Window",,,ErrorColor
    End If
End Sub

'--------------------------------------------------------------------------------------
'"Պայմանագրի Բացում" Պատուհանի Éñ³óáõÙ
'--------------------------------------------------------------------------------------
Sub OpenContract()
    
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_AgrOpen)
    wMDIClient.Refresh
    
    Set DocForm = AsBank.WaitVBObject("frmAsMsgBox", delay_middle)
    If DocForm.Exists Then
        Call MessageExists(2,"ä³ÛÙ³Ý³·ñÇ µ³óáõÙ")
        Call ClickCmdButton(5, "²Ûá")
    Else
        Log.Error "Can Not Open Rc(OpenContract/Պայմանագրի Բացում) Window",,,ErrorColor         
    End If    
    BuiltIn.Delay(1500)
    If DocForm.Exists Then
        Log.Error "Can Not Close Rc(OpenContract/Պայմանագրի Բացում) Window",,,ErrorColor
    End If
End Sub
'--------------------------------------------------------------------------------------
'"Դուրս գրածի վերականգնում" ÷³ëï³ÃÕÃÇ Éñ³óáõÙ :
'--------------------------------------------------------------------------------------
Sub WriteOut_Reconstruction(WriteOff, WarningWin)
    
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_Opers & "|" & c_WriteOff & "|" & c_WriteOffBack)
    wMDIClient.Refresh
    
    Set DocForm = wMDIClient.WaitVBObject("frmASDocForm", delay_middle)
    
    If DocForm.Exists Then
        'ISN-ի վերագրում
        WriteOff.isn = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
        Call Fill_WriteOut(WriteOff)
        If WarningWin Then
            Call ClickCmdButton(5, "Î³ï³ñ»É")
        End If
    Else
        Log.Error "Can Not Open Rc(WriteOffReconstruction/Դուրս գրածի վերականգնում) Window",,,ErrorColor     
    End If    
    BuiltIn.Delay(1500)
    If DocForm.Exists Then
        Log.Error "Can Not Close Rc(WriteOffReconstruction/Դուրս գրածի վերականգնում) Window",,,ErrorColor
    End If
End Sub

'--------------------------------------------------------------------------------------
'"Դուրս գրում" ÷³ëï³ÃÕÃÇ Éñ³óáõÙ :
'--------------------------------------------------------------------------------------
Sub Create_WriteOut(WriteOut)
    
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_Opers & "|" & c_WriteOff & "|" & c_WriteOff)
    wMDIClient.Refresh
    
    Set DocForm = wMDIClient.WaitVBObject("frmASDocForm", delay_middle)
    
    If DocForm.Exists Then
        'ISN-ի վերագրում
        WriteOut.isn = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
        Call Fill_WriteOut(WriteOut)
    Else
        Log.Error "Can Not Open Rc(WriteOut/Դուրս գրում) Window",,,ErrorColor      
    End If    
    BuiltIn.Delay(1500)
    If DocForm.Exists Then
        Log.Error "Can Not Close Rc(WriteOut/Դուրս գրում) Window",,,ErrorColor
    End If
    
End Sub

'---------------------------------------------------------------------------------------------
' Լրացնել "Գործողություններ/¹áõñë ·ñáõÙ" պատուհանի դաշտերը
'---------------------------------------------------------------------------------------------
Sub Fill_WriteOut(WriteOut)
    
    'Ստուգում "Պայմանագրի N" դաշտի խմբագրելիությունը և արժեքը
    Call Check_ReadOnly("Document",1,"General","CODE",True) 
    Call Compare_Two_Values("Պայմանագրի N",Get_Rekvizit_Value("Document",1,"Mask","CODE"),WriteOut.ExpectedAgreementN)
     
    'Լրացնել "Ամսաթիվ" դաշտը
    Call Rekvizit_Fill("Document", 1, "General", "DATE", WriteOut.Date)
    
    'Ստուգել և Լրացնել "Հիմանական գումար" դաշտը
    Call Compare_Two_Values("Հիմանական գումար",Get_Rekvizit_Value("Document",1,"General","SUMAGR"),WriteOut.ExpectedBaseSum)
    Call Rekvizit_Fill("Document", 1, "General", "SUMAGR", WriteOut.BaseSum)

    'Ստուգել և Լրացնել "Ժամկետանց գումարի տույժ" դաշտը
    Call Compare_Two_Values("Ժամկետանց գումարի տույժ",Get_Rekvizit_Value("Document",1,"General","SUMFINE"),WriteOut.ExpectedFineOnPastSum)
    Call Rekvizit_Fill("Document", 1, "General", "SUMFINE", WriteOut.FineOnPastSum)
    
    'Ստուգում "Ընդանուր գումար"  դաշտի արժեքը
    Call Compare_Two_Values("Ընդանուր գումար",Get_Rekvizit_Value("Document",1,"General","SUMMA"),WriteOut.TotalSum)
    
    'Լրացնել "Մեկանաբանություն" դաշտը
    Call Rekvizit_Fill("Document",1,"General","COMMENT","![End][Del]" & WriteOut.Comment)
    
    'Լրացնել "Գրասենյակ/բաժին" դաշտերը
    Call Rekvizit_Fill("Document",1,"General","ACSBRANCH",WriteOut.Division & "[Tab]" & WriteOut.Department)
    
    Call ClickCmdButton(1, "Î³ï³ñ»É")
End Sub

'--------------------------------------------------------------------------------------
'Պահուստավորում փաստաթղթի լրացում
'--------------------------------------------------------------------------------------
Sub Doc_Store(Store)
    
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_Opers & "|" & c_Store & "|" & c_Store)
    wMDIClient.Refresh

    Set DocForm = wMDIClient.WaitVBObject("frmASDocForm", delay_middle)
    
    If DocForm.Exists Then
        'ISN-ի վերագրում
        Store.isn = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
        Call Fill_StoreWin(Store)
    Else
        Log.Error "Can Not Open Rc(Store/Պահուստավորում) Window",,,ErrorColor      
    End If    
    BuiltIn.Delay(1500)
    If DocForm.Exists Then
        Log.Error "Can Not Close Rc(Store/Պահուստավորում) Window",,,ErrorColor
    End If
End Sub

'---------------------------------------------------------------------------------------------
' Լրացնել "Գործողություններ/Պահուստավորում" պատուհանի դաշտերը
'---------------------------------------------------------------------------------------------
Sub Fill_StoreWin(Store)
    
    'Ստուգում "Պայմանագրի N" դաշտի խմբագրելիությունը և արժեքը
    Call Check_ReadOnly("Document",1,"General","CODE",True) 
    Call Compare_Two_Values("Պայմանագրի N",Get_Rekvizit_Value("Document",1,"Mask","CODE"),Store.ExpectedAgreementN)
     
    'Լրացնել "Ամսաթիվ" դաշտը
    Call Rekvizit_Fill("Document", 1, "General", "DATE", Store.Date)
    
    'Լրացնել "ä³Ñáõëï³íáñում" դաշտը
    Call Rekvizit_Fill("Document", 1, "General", "SUMRES", Store.Provision)
    
    'Լրացնել "Ապապ³Ñáõëï³íáñում" դաշտը
    Call Rekvizit_Fill("Document", 1, "General", "SUMUNRES", Store.UnProvision)
    
    'Լրացնել "Մեկանաբանություն" դաշտը
    Call Rekvizit_Fill("Document",1,"General","COMMENT","![End][Del]" & Store.Comment)
    
    'Լրացնել "Գրասենյակ/բաժին" դաշտերը
    Call Rekvizit_Fill("Document",1,"General","ACSBRANCH",Store.Division & "[Tab]" & Store.Department)
    
    Call ClickCmdButton(1, "Î³ï³ñ»É")
End Sub

'--------------------------------------------------------------------------------------
'èÇëÏÇ ¹³ëÇã ¨ å³Ñáõëï³íáñÙ³Ý ïáÏáë ÷³ëïÃÕÃÇ Éñ³óáõÙ :
'--------------------------------------------------------------------------------------
'ExpectedAgreementN - Սպասվող Պայմանագրի համարը
'Date - Ամսաթիվ
'RiskLeve - Ռիսկի դասիչ
'Percent - Պահուստավորման տոկոս
'Comment - Մեկանաբանություն
Function Create_Risk_Classifier(ExpectedAgreementN, Date, RiskLeve, Percent,Comment)
    
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_TermsStates & "|" & c_Risking & "|" & c_RiskCatPerRes)
    wMDIClient.Refresh

    Set DocForm = wMDIClient.WaitVBObject("frmASDocForm", delay_middle)
    
    If DocForm.Exists Then
        'Ստուգում "Պայմանագրի N" դաշտի խմբագրելիությունը և արժեքը
        Call Check_ReadOnly("Document",1,"General","CODE",True) 
        Call Compare_Two_Values("Պայմանագրի N",Get_Rekvizit_Value("Document",1,"Mask","CODE"),ExpectedAgreementN)
     
        'Լրացնել "Ամսաթիվ" դաշտը
        Call Rekvizit_Fill("Document", 1, "General", "DATE", Date)
        'Լրացնել "Ռիսկի դասիչ" դաշտը
        Call Rekvizit_Fill("Document", 1, "General", "RISK", RiskLeve)
        'Լրացնել "ä³Ñáõëï³íáñÙ³Ý ïáÏáë" դաշտը
        Call Rekvizit_Fill("Document", 1, "General", "PERRES", Percent)
        'Լրացնել "Մեկանաբանություն" դաշտը
        Call Rekvizit_Fill("Document",1,"General","COMMENT","![End][Del]" & Comment)
        'ISN-ի վերագրում
        Create_Risk_Classifier = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
    
        Call ClickCmdButton(1, "Î³ï³ñ»É")
    Else
        Log.Error "Can Not Open Rc(Objective Risk/Օբյեկտիվ ռիսկի դասիչ) Window",,,ErrorColor      
    End If
    BuiltIn.Delay(1500)
    If DocForm.Exists Then
        Log.Error "Can Not Close Rc(Objective Risk/Օբյեկտիվ ռիսկի դասիչ) Window",,,ErrorColor
    End If
End Function

'------------------------------------------------------------------------------------
' Գերածախսի "Պարտքերի մարում" գործողության կատարում
'------------------------------------------------------------------------------------
Function Overlimit_Repay(OverlimitRepay)
    Dim DocForm
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_Opers & "|" & c_GiveAndBack & "|" & c_PayOffDebt)
    wMDIClient.Refresh
   
    Set DocForm = wMDIClient.WaitVBObject("frmASDocForm", delay_middle)
    
    If DocForm.Exists Then
        'ISN-ի վերագրում փոփոխականին
        OverlimitRepay.Isn = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
    
        Call Fill_OverlimitRepay(OverlimitRepay)
        
        If OverlimitRepay.CashCashles = "1" Then
            BuiltIn.Delay(3000)
            Call ClickCmdButton(1, "Î³ï³ñ»É")
        Else
            BuiltIn.Delay(1000)
            Call MessageExists(2,"²í³ñï»±É ·áñÍáÕáõÃÛáõÝÁ ³ÝÙÇç³å»ëª ÃÕÃ³ÏóáõÃÛáõÝÁ" & vbCrLf & "Ï³ï³ñ»Éáí Ñ³ßíÇ Ñ»ï." & vbCrLf & "" & vbCrLf & "      ² Ú à    -    ÃÕÃ³ÏóáõÃÛáõÝ Ñ³ßíÇ Ñ»ï" & vbCrLf & "      à â        -    ÷³ëï³ÃÕÃÇ áõÕ³ñÏáõÙ ÃÕÃ³å³Ý³ÏÝ»ñ")
            Call ClickCmdButton(5, "²Ûá")
        End If
    Else
        Log.Error "Can Not Open Rc(OverlimitRepay/Պարտքերի մարում) Window",,,ErrorColor  
    End If
    BuiltIn.Delay(3000)
    If DocForm.Exists Then
        Log.Error "Can Not Close Rc(OverlimitRepay/Պարտքերի մարում) Window",,,ErrorColor
    End If
End Function

'---------------------------------------------------------------------------------------------
' Լրացնել "Գործողություններ/պարտքերի մարում" պատուհանի դաշտերը
'---------------------------------------------------------------------------------------------
Sub Fill_OverlimitRepay(OverlimitRepay)
    'Ստուգում "Պայմանագրի N" դաշտի խմբագրելիությունը և արժեքը
    Call Check_ReadOnly("Document",1,"General","CODE",True) 
    Call Compare_Two_Values("Պայմանագրի N",Get_Rekvizit_Value("Document",1,"Mask","CODE"),OverlimitRepay.ExpectedAgreementN)
    'Լրացնել "ամսաթիվ" դաշտը
    Call Rekvizit_Fill("Document", 1, "General", "DATE", OverlimitRepay.Date)
    'Ստուգում "Մարման Արժույթ" դաշտի համար  
    If Not Check_ReadOnly("Document",1,"Mask","REPAYCURR",True) Then
        Call Rekvizit_Fill("Document", 1, "Mask", "REPAYCURR", OverlimitRepay.RepaymentCurrency) 
    End If
    
    'Ստուգել և Լրացնել "Հիմնական գումար" դաշտը
    Call Compare_Two_Values("Հիմնական գումար",Get_Rekvizit_Value("Document",1,"General","SUMAGR"),OverlimitRepay.ExpectedBaseSum)
    Call Rekvizit_Fill("Document", 1, "General", "SUMAGR", OverlimitRepay.BaseSum) 
    'Ստուգում "ՀՀ Դրամ" դաշտի խմբագրելիությունը և արժեքը
    Call Check_ReadOnly("Document",1,"General","AMDSUMAGR",True) 
    Call Compare_Two_Values("ՀՀ Դրամ",Get_Rekvizit_Value("Document",1,"General","AMDSUMAGR"),OverlimitRepay.AMD1)
    
    'Ստուգել և Լրացնել "Ժամկետանց գումարի տույժ" արժեքները
    Call Compare_Two_Values("Ժամկետանց գումարի տույժ",Get_Rekvizit_Value("Document",1,"General","SUMFINE"),OverlimitRepay.ExpectedFineOnPastSum)
    Call Rekvizit_Fill("Document", 1, "General", "SUMFINE", OverlimitRepay.FineOnPastSum) 
    'Ստուգել "ՀՀ Դրամ" դաշտերի խմբագրելիությունը և արժեքները
    Call Check_ReadOnly("Document",1,"General","AMDSUMFINE",True) 
    Call Compare_Two_Values("ՀՀ Դրամ",Get_Rekvizit_Value("Document",1,"General","AMDSUMFINE"),OverlimitRepay.AMD2)

    'Ստուգում "Ընդանուր գումար" դաշտի խմբագրելիությունը և արժեքը
    Call Check_ReadOnly("Document",1,"General","SUMMA",True) 
    Call Compare_Two_Values("Ընդանուր գումար",Get_Rekvizit_Value("Document",1,"General","SUMMA"),OverlimitRepay.TotalAmount)

    'Լրացնել "Կանխիկ/Անկանխիկ" դաշտը
    Call Rekvizit_Fill("Document",1,"General","CASHORNO",OverlimitRepay.CashCashles)
    
    'Ստուգում "Տոկոսի/տույժի փոխարժեք"  և "առ" դաշտերի խմբագրելիությունը և արժեքները
    Call Check_ReadOnly("Document",1,"Course1","CURS",True) 
    Call Check_ReadOnly("Document",1,"Course2","CURS",True) 
    Call Compare_Two_Values("Տոկոսի/տույժի փոխարժեք/առ",Get_Rekvizit_Value("Document",1,"Course","CURS"),OverlimitRepay.ExchangeRate & "/" & OverlimitRepay.Per)

    If Get_Rekvizit_Value("Document",1,"Mask","CASHORNO") = "2" Then
        'Լրացնել "Հաշիվ" դաշտը
        Call Rekvizit_Fill("Document",1,"General","ACCCORR", OverlimitRepay.Account)
    End If
    If Not Check_ReadOnly("Document",1,"Mask","AMDACCCORR",True) Then
        'Լրացնել "Դրամային Հաշիվ" դաշտը
        Call Rekvizit_Fill("Document",1,"General","AMDACCCORR", OverlimitRepay.AMDAccount)
    End If
    'Լրացնել "Մեկանաբանություն" դաշտը
    Call Rekvizit_Fill("Document",1,"General","COMMENT","![End][Del]" & OverlimitRepay.Comment)
    'Լրացնել "Ընթացիկ մնացորդներ" դաշտը
    If InStr(Get_Rekvizit_Value("Document",1,"General","REMINFO"), OverlimitRepay.RemittanceInfo1) = 0 Then
        Log.Error "Ընթացիկ մնացորդներ - "& OverlimitRepay.RemittanceInfo1 &" Does not exist This Text = "& Get_Rekvizit_Value("Document",1,"General","REMINFO"),,,ErrorColor
    End If
    If InStr(Get_Rekvizit_Value("Document",1,"General","REMINFO"), OverlimitRepay.RemittanceInfo2) = 0Then
        Log.Error "Ընթացիկ մնացորդներ - "& OverlimitRepay.RemittanceInfo2 &" Does not exist This Text = "& Get_Rekvizit_Value("Document",1,"General","REMINFO"),,,ErrorColor
    End If
    'Լրացնել "Գրասենյակ/բաժին" դաշտերը
    Call Rekvizit_Fill("Document",1,"General","ACSBRANCH",OverlimitRepay.Division & "[Tab]" & OverlimitRepay.Department)

    Call ClickCmdButton(1, "Î³ï³ñ»É")
End Sub

'------------------------------------------------------------------------------
'"Պայմաններ և վիճակներ/Օբյեկտիվ ռիսկի դասիչ" գործողության կատարում
'ExpectedAgreementN - Սպասվող Պայմանագրի համարը
'Date - Ամսաթիվ
'RiskLeve - Ռիսկի դասիչ
'Comment - Մեկանաբանություն
'------------------------------------------------------------------------------
Function Objective_Risk(ExpectedAgreementN,Date, RiskLeve, Comment, SecondWin)

    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_TermsStates & "|" & c_Risking & "|" & c_ObjRiskCat)
    wMDIClient.Refresh
    
    Set DocForm = wMDIClient.WaitVBObject("frmASDocForm", delay_middle)
    
    If DocForm.Exists Then
        'Ստուգում "Պայմանագրի N" դաշտի խմբագրելիությունը և արժեքը
        Call Check_ReadOnly("Document",1,"General","CODE",True) 
        Call Compare_Two_Values("Պայմանագրի N",Get_Rekvizit_Value("Document",1,"Mask","CODE"),ExpectedAgreementN)
     
        'Լրացնել "Ամսաթիվ" դաշտը
        Call Rekvizit_Fill("Document", 1, "General", "DATE", Date)
        'Լրացնել "Ռիսկի դասիչ" դաշտը
        Call Rekvizit_Fill("Document", 1, "General", "RISK", RiskLeve)
        'Լրացնել "Մեկանաբանություն" դաշտը
        Call Rekvizit_Fill("Document",1,"General","COMMENT","![End][Del]" & Comment)
        'ISN-ի վերագրում
        Objective_Risk = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
    
        Call ClickCmdButton(1, "Î³ï³ñ»É")
        BuiltIn.Delay(1000)
        If SecondWin Then
            Call MessageExists(2,"¶»ñ³Í³Ëë³ÛÇÝ å³ÛÙ³Ý³·Çñª  "&ExpectedAgreementN&"  /öáË³ÝóÙ³Ý ëïáõ·Ù³Ý "& vbCrLf &"Ñ³×³Ëáñ¹ 1/"& vbCrLf &"--------------------------------------------------------------------------------------------------------------"& vbCrLf &""& vbCrLf &"²Ûë ûµÛ»ÏïÇí éÇëÏÇ ¹³ëÇãÁ ³ñ¹»Ý Ýß³Ý³Ïí³Í ¿")
            Call ClickCmdButton(5, "Î³ï³ñ»É")
        End If  
    Else
        Log.Error "Can Not Open Rc(Objective Risk/Օբյեկտիվ ռիսկի դասիչ) Window",,,ErrorColor
    End If
    BuiltIn.Delay(1500)
    If DocForm.Exists Then
        Log.Error "Can Not Close Rc(Objective Risk/Օբյեկտիվ ռիսկի դասիչ) Window",,,ErrorColor
    End If
End Function

'------------------------------------------------------------------------------
'Տոկոսադրույքներ գործողության կատարում
'ExpectedAgreementN -Սպասվող Պայմանագրի համարը
'Date - Ամսաթիվ
'FineOnPastDueSum - Ժամկետանց գումարի տույժ
'Month - Ամսական
'Comment - Մեկանաբանություն
'------------------------------------------------------------------------------
Function ChangeOverlimitRete( ExpectedAgreementN, Date, FineOnPastDueSum, Month, Comment)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_TermsStates & "|" & c_Percentages & "|" & c_Percentages)
    wMDIClient.Refresh
    
    Set DocForm = wMDIClient.WaitVBObject("frmASDocForm", delay_middle)
    
    If DocForm.Exists Then
        'Ստուգում "Պայմանագրի N" դաշտի խմբագրելիությունը և արժեքը
        Call Check_ReadOnly("Document",1,"General","CODE",True) 
        Call Compare_Two_Values("Պայմանագրի N",Get_Rekvizit_Value("Document",1,"Mask","CODE"),ExpectedAgreementN)
        'Լրացնել "Ամսաթիվ" դաշտը
        Call Rekvizit_Fill("Document", 1, "General", "DATE", Date)
        'Լրացնել "Ժամկետանց գումարի տույժ" և "Ամսական" դաշտերը
        Call Rekvizit_Fill("Document", 1, "General", "PCPENAGR", FineOnPastDueSum &"[Tab]"& Month)
        'Լրացնել "Մեկանաբանություն" դաշտը
        Call Rekvizit_Fill("Document",1,"General","COMMENT","![End][Del]" & Comment)
        'ISN-ի վերագրում
        ChangeOverlimitRete = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
    
        Call ClickCmdButton(1, "Î³ï³ñ»É")
    Else
        Log.Error "Can Not Open Rc(Change Overlimit Rete/Տոկոսադրույքներ գործողության կատարում) Window",,,ErrorColor    
    End If
    BuiltIn.Delay(1500)
    If DocForm.Exists Then
        Log.Error "Can Not Close Rc(Change Overlimit Rete/Տոկոսադրույքներ գործողության կատարում) Window",,,ErrorColor
    End If
End Function

'------------------------------------------------------------------------------------
' Պայմանագրեր թղթապանակում փաստատթղթի առկայության ստուգում
'------------------------------------------------------------------------------------
Function ExistsContract_Filter_Fill(FolderName, ContractsFilter, RowCount)
    Call wTreeView.DblClickItem(FolderName & "ä³ÛÙ³Ý³·ñ»ñ")
    BuiltIn.Delay(1500)
    Call Fill_ContractsFilter(ContractsFilter)
    Set DocForm = wMDIClient.VBObject("frmPttel")
    
    If WaitForPttel("frmPttel") Then
        wMDIClient.Refresh
        If DocForm.vbObject("tdbgView").ApproxCount = RowCount Then
            Log.Message "Row count of Contracts is right",,,MessageColor
            ExistsContract_Filter_Fill = True
        Else
            Log.Error "Row count of Contracts is not right",,,ErrorColor
            ExistsContract_Filter_Fill = False
        End If
    Else
        Log.Error "Can Not Open Պայմանագրեր Window",,,ErrorColor      
    End If     
End Function 
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''AllocFundsOperations'''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Գործողությունների դիտում պատուհանի լրացման կլաս
Class AllocFundsOperations
		public agreeLevel
		public startDate
		public endDate
		public agreeType
		public agreeN
		public agreePaperN
		public creditCodeExists
		public creditCode
		public curr
		public clientExists
		public client 
		public clientName
		public operationType
		public performer
		public note
		public note2
		public note3
		public agreeOffice
		public agreeSection
		public accessType
		public toWatch
		public fill
		private Sub Class_Initialize()
				agreeLevel = "1"
				startDate = ""
				endDate = ""
				agreeType = ""
				agreeN = ""
				agreePaperN = ""
				creditCodeExists = false
				creditCode = ""
				curr = ""
				clientExists = false
				client = ""
				clientName = ""
				operationType = ""
				performer = ""
				note = ""
				note2 = ""
				note3 = ""
				agreeOffice = ""
				agreeSection = ""
				accessType = ""
				toWatch = "AGRDEALS"
				fill = "0"
		End Sub
End Class

Function New_AllocFundsOperations()
		Set New_AllocFundsOperations = new AllocFundsOperations
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''Fill_AllocFundsOperations'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Գործողությունների դիտում պատուհանի լրացման պրոցեդուրա
'AllocFundsOpers - պատուհանի լրացման կլաս
Sub Fill_AllocFundsOperations(AllocFundsOpers)
  ' Պայմանագրի մակարդակ դաշտի լրացում 
		Call Rekvizit_Fill("Dialog", 1, "General", "LEVEL", "![End]" & "[Del]" & AllocFundsOpers.agreeLevel)
		' Ժամանակահատված սկզբնական դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "START", "![End]" & "[Del]" & AllocFundsOpers.startDate)
		' Ժամանակահատված վերջնական դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "END", "![End]" & "[Del]" & AllocFundsOpers.endDate)
		' Պայմանագրի տեսակ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "AGRKIND", AllocFundsOpers.agreeType)
		' Պայմանագրի N դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "NUM", AllocFundsOpers.agreeN)
		' Պայմ. թղթային N դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "PPRCODE", AllocFundsOpers.agreePaperN)
		if AllocFundsOpers.creditCodeExists then
		  ' Վարկային կոդ դաշտի լրացում
				Call Rekvizit_Fill("Dialog", 1, "General", "CRDTCODE", AllocFundsOpers.creditCode)
		end if
		' Արժույթ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "CUR", AllocFundsOpers.curr)
		if AllocFundsOpers.clientExists then
		  ' Հաճախորդ դաշտի լրացում
				Call Rekvizit_Fill("Dialog", 1, "General", "CLIENT", AllocFundsOpers.client)
				' Հաճախորդի անվանում դաշտի լրացում
				Call Rekvizit_Fill("Dialog", 1, "General", "NAME", AllocFundsOpers.clientName)
		end if
		' Գործողության տեսակ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "DEALTYPE", AllocFundsOpers.operationType)
		' Կատարող դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "USER", AllocFundsOpers.performer)
		' Նշում դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "NOTE", AllocFundsOpers.note)
		' Նշում 2 դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "NOTE2", AllocFundsOpers.note2)
		' Նշում 3 դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "NOTE3", AllocFundsOpers.note3)
		' Պայմ. գրասենյակ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "ACSBRANCH", AllocFundsOpers.agreeOffice)
		' Պայմ. բաժին դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "ACSDEPART", AllocFundsOpers.agreeSection)
		' Հասան-ն տիպ բաժնի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "ACSTYPE", AllocFundsOpers.accessType)
		' Դիտելու ձև դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "SELECTED_VIEW", "![End]" & "[Del]" & AllocFundsOpers.toWatch)
		' Լրացնել դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "EXPORT_EXCEL", "![End]" & "[Del]" & AllocFundsOpers.fill)
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''GoTo_AllocFundsOperations''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Գործողությունների դիտում թղթապանակ մուտք գործելու պրոցեդուրա
'folderName - գտնբելու ճանապարհը
'AllocFundsOpers - պատուհանի լրացման կլաս
Sub GoTo_AllocFundsOperations(folderName, AllocFundsOpers)
		wTreeView.DblClickItem(folderName & "¶áñÍáÕáõÃÛáõÝÝ»ñ")
		if p1.WaitVBObject("frmAsUstPar", 3000).Exists then
				Call Fill_AllocFundsOperations(AllocFundsOpers)
				Call ClickCmdButton(2, "Î³ï³ñ»É")
		else 
				Log.Error "Can't open frmAsUstPar widow.", "", pmNormal, ErrorColor
		end if
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''columnSorting''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''
'Պրոցեդուրան սորտավորում է պատուհանի նշված սյունը
'colName - սորտավորվող սյան անունը (անունների զանգված)
'sortColCount - սորտավորվող սյուների քանակը
'frmWin - պատուհանի տեսակը
Sub columnSorting(colName, sortColCount, frmWin)
		Dim i, colNum
		for i = 0 to sortColCount - 1
				colNum =	wMDIClient.VBObject(frmWin).GetColumnIndex(colName(i))
				wMDIClient.VBObject(frmWin).Keys("[Hold]" & "^!" & (colNum + 1))
		next
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''SubsystemWorkingDocuments'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Աշխատանքային փաստաթղթեր պատուհանի լրացման կլաս
Class SubsystemWorkingDocuments
		public startDate
		public endDate
		public approvalGroup
		public stateFill
		public states
		public agreeOpers
		public agreeN
		public agreePaperN
		public curr
		public client
		public clientName
		public note
		public note2
		public note3
		public performers
		public office
		public section
		public accessType
		private Sub Class_Initialize()
				startDate = ""
				endDate = ""
				approvalGroup = ""
				stateFill = 0
				states = ""
				agreeOpers = ""
				agreeN = ""
				agreePaperN = ""
				curr = ""
				client = ""
				clientName = ""
				note = ""
				note2 = ""
				note3 = ""
				performers = ""
				office = ""
				section = ""
				accessType = ""
		End Sub
End Class

Function New_SubsystemWorkingDocuments()
		Set New_SubsystemWorkingDocuments = new SubsystemWorkingDocuments
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''Fill_SubsystemWorkingDocuments'''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Աշխատանքային փաստաթղթեր պատուհանի լրացման պրոցեդուրա
'WorkingDocs - պատուհանի լրացման կլաս
Sub Fill_SubsystemWorkingDocuments(WorkingDocs)
  ' Ժամանակահատված սկզբնական դաշտի լրացում 
		Call Rekvizit_Fill("Dialog", 1, "General", "PERN", "![End]" & "[Del]" & WorkingDocs.startDate)
		' Ժամանակահատված վերջնական դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "PERK", "![End]" & "[Del]" & WorkingDocs.endDate)
		' Հաստատման խումբ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "SSCONFTP", WorkingDocs.approvalGroup)
		if WorkingDocs.stateFill = 1 then 
		  ' Պայմաններ-վիճակներ դաշտի լրացում
				Call Rekvizit_Fill("Dialog", 1, "General", "SSAGRMEDR", WorkingDocs.states)
		end if
		' Պայմանագրերի գործողություններ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "SSAGRMOPR", WorkingDocs.agreeOpers)
		' Պայմանագրի N դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "NUM", WorkingDocs.agreeN)
		' Պայմ. թղթային N դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "PPRCODE", WorkingDocs.agreePaperN)
		' Արժույթ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "CUR", WorkingDocs.curr)
		' Հաճախորդ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "CLIENT", WorkingDocs.client)
		' Հաճախոդի անվանում դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "NAME", WorkingDocs.clientName)
		' Նշում դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "NOTE", WorkingDocs.note)
		' Նշում 2 դաշտի լարցում
		Call Rekvizit_Fill("Dialog", 1, "General", "NOTE2", WorkingDocs.note2)
		' Նշում 3 դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "NOTE3", WorkingDocs.note3)
		' Կատարողներ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "USER", "![End]" & "[Del]" & WorkingDocs.performers)
		' Գրասենյակ բաժնի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "ACSBRANCH", WorkingDocs.office)
		' Բաժին դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "ACSDEPART", WorkingDocs.section)
		' Հասան-ն տիպ դաշտի լրացումմ
		Call Rekvizit_Fill("Dialog", 1, "General", "ACSTYPE", WorkingDocs.accessType)
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''GoTo_SubsystemWorkingDocuments'''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Աշխատանքային փաստաթղթեր թղթապանակ մուտք գործելու պրոցեդուրա
'folderName - գտնբելու ճանապարհը
'WorkingDocs - պատուհանի լրացման կլաս
Sub GoTo_SubsystemWorkingDocuments(folderName, WorkingDocs)
		wTreeView.DblClickItem(folderName & "²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ")
		if p1.WaitVBObject("frmAsUstPar", 3000).Exists then
				Call Fill_SubsystemWorkingDocuments(WorkingDocs)
				Call ClickCmdButton(2, "Î³ï³ñ»É")
		else 
				Log.Error "Can't open frmAsUstPar widow.", "", pmNormal, ErrorColor
		end if
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''AgreementsCommomFilter''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պայմանագրերում միևնույն տիպի ֆիլտրերի լրացման կլաս
'Կլասը նախատեսված է հետևյալ դաշտերը պարունակող ֆիլտրերի լրացման համար՝
'Ժամանակահատված(սկիզբ, վերջ), Պայմանագրի N, Կատարող, Նշում(2,3)
'Երկարաձգում, Գրաֆիկի տեսակ, Ցույց տալ բացված տեսքով, Պայմ. գրասենյակ
'Պայմ. բաժին, Հասան-ն տիպ, Միայն փոփոխությունները
Class AgreementsCommomFilter
		public startDate																					
		public endDate				
		public leaseAgree
		public agreeType										
		public agreeN
		public performer
		public note
		public note2
		public note3
		public agreeOffice
		public agreeSection
		public accessType
		public onlyChangesExists
		public onlyChanges
		public showInOpFormExists
		public showInOpForm
		public scheduleTypeExists
		public scheduleType
		public extensionExists
		public extension
		private Sub Class_Initialize()
				startDate = ""
				endDate = ""
				leaseAgree = false
				agreeType = ""
				agreeN = ""
				performer = ""
				note = ""
				note2 = ""
				note3 = ""
				agreeOffice = ""
				agreeSection = ""
				accessType = ""
				onlyChangesExists = false
				onlyChanges = 0
				showInOpFormExists = false
				showInOpForm = 0
				scheduleTypeExists = false
				scheduleType = ""
				extensionExists = false
				extension = ""
		End Sub
End Class

Function New_AgreementsCommomFilter()
		Set New_AgreementsCommomFilter = new AgreementsCommomFilter
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''Fill_AgreementsCommomFilter''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պայմանագրերում միևնույն տիպի ֆիլտրերի պատուհանի լրացման պրոցեդուրա
'AgreeCommonFilter - պատուհանի լրացման կլաս
'Պրոցեդուրան նախատեսված է հետևյալ դաշտերը պարունակող ֆիլտրերի լրացման համար՝
'Ժամանակահատված(սկիզբ, վերջ), Պայմանագրի N, Կատարող, Նշում(2,3)
'Երկարաձգում, Գրաֆիկի տեսակ, Ցույց տալ բացված տեսքով, Պայմ. գրասենյակ
'Պայմ. բաժին, Հասան-ն տիպ, Միայն փոփոխությունները
Sub Fill_AgreementsCommomFilter(AgreeCommonFilter)
  ' Ժամանակահատված սկզբնական դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "START", "![End]" & "[Del]" & AgreeCommonFilter.startDate)
		' Ժամանակահատված վերջնական դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "END", "![End]" & "[Del]" & AgreeCommonFilter.endDate)
		if AgreeCommonFilter.leaseAgree then 
		  ' Պայմանագրի տեսակ դաշտի լրացում
				Call Rekvizit_Fill("Dialog", 1, "General", "AGRKIND", AgreeCommonFilter.agreeType)
				' Պայմանագրի N դաշտի լրացում
				Call Rekvizit_Fill("Dialog", 1, "General", "LCNUM", AgreeCommonFilter.agreeN)
		else 
		  ' Պայմանագրի N դաշտի լրացում
				Call Rekvizit_Fill("Dialog", 1, "General", "AGR", AgreeCommonFilter.agreeN)
		end if
		' Կատարող դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "USER", AgreeCommonFilter.performer)
		' Նշում դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "NOTE", AgreeCommonFilter.note)
		' Նշում 2 դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "NOTE2", AgreeCommonFilter.note2)
		' Նշում 3 դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "NOTE3", AgreeCommonFilter.note3)
		if AgreeCommonFilter.extensionExists then
		  ' Երկարաձգուն դաշտի լրացում 
				Call Rekvizit_Fill("Dialog", 1, "General", "PROLONGATION", AgreeCommonFilter.extension)
		end if
		if AgreeCommonFilter.scheduleTypeExists then
		  ' Գրաֆիկի տեսակ դաշտի լրացում
				Call Rekvizit_Fill("Dialog", 1, "General", "SCHTYPE", AgreeCommonFilter.scheduleType)
		end if
		if AgreeCommonFilter.showInOpFormExists then
		  ' Ցույց տալ բացված տեսքով դաշտի լրացում
				Call Rekvizit_Fill("Dialog", 1, "CheckBox", "SHOWOPENVIEW", AgreeCommonFilter.showInOpForm)
		end if
		' Պայմ. գրասենյակ դաշտի լրացում 
		Call Rekvizit_Fill("Dialog", 1, "General", "ACSBRANCH", AgreeCommonFilter.agreeOffice)
		' Պայմ. բաժին դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "ACSDEPART", AgreeCommonFilter.agreeSection)
		' Հասան-ն տիպ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "ACSTYPE", AgreeCommonFilter.accessType)
		if AgreeCommonFilter.onlyChangesExists then
		  'Միայն փոփոխությունները դաշտի լրացում
				Call Rekvizit_Fill("Dialog", 1, "CheckBox", "ONLYCH", AgreeCommonFilter.onlyChanges)
		end if
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''GoTo_AgreementsCommomFilter'''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պայմանագրերում միևնույն տիպի ֆիլտրերի թղթապանակ մուտք գործելու պրոցեդուրա
'folderName - գտնբելու ճանապարհը
'typeName - գործողություն անունը
'AgreeCommonFilter - պատուհանի լրացման կլաս
Sub GoTo_AgreementsCommomFilter(folderName, typeName, AgreeCommonFilter)
		wTreeView.DblClickItem(folderName & typeName)
		if p1.WaitVBObject("frmAsUstPar", 3000).Exists then
				Call Fill_AgreementsCommomFilter(AgreeCommonFilter)
				Call ClickCmdButton(2, "Î³ï³ñ»É")
		else 
				Log.Error "Can't open frmAsUstPar widow.", "", pmNormal, ErrorColor
		end if
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''RecordedNotes''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Հաշվառվող նշումներ պատուհանի լրացման կլաս
Class RecordedNotes
		public startDate
		public endDate
		public performer
		public onlyChanges
		private Sub Class_Initialize()
				startDate = ""
				endDate = ""
				performer = ""
				onlyChanges = 0
		End Sub
End Class

Function New_RecordedNotes()
		Set New_RecordedNotes = new RecordedNotes
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''Fill_RecordedNotes''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Հաշվառվող նշումներ պատուհանի լրացման պրոցեդուրա
'RecNotes - պատուհանի լրացման կլաս
Sub Fill_RecordedNotes(RecNotes)
  ' Ժամանակահատված սկզբնական դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "START", "![End]" & "[Del]" & RecNotes.startDate)
		' Ժամանակահատված վերջնական դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "END", "![End]" & "[Del]" & RecNotes.endDate)
		' Կատարող դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "USER", RecNotes.performer)
		' Միայն փոփոխությունները դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "CheckBox", "ONLYCH", RecNotes.onlyChanges)
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''GoTo_RecordedNotes''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Հաշվառվող նշումներ թղթապանակ մուտք գործելու պրոցեդուրա
'folderName - գտնբելու ճանապարհը
'typeName - գործողություն անունը
'RecNotes - պատուհանի լրացման կլաս
Sub GoTo_RecordedNotes(folderName, RecNotes)
		wTreeView.DblClickItem(folderName & "Ð³ßí³éíáÕ ÝßáõÙÝ»ñ")
		if p1.WaitVBObject("frmAsUstPar", 3000).Exists then
				Call Fill_RecordedNotes(RecNotes)
				Call ClickCmdButton(2, "Î³ï³ñ»É")
		else 
				Log.Error "Can't open frmAsUstPar widow.", "", pmNormal, ErrorColor
		end if
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''CreditCards''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Վարկային քարտեր պատուհանի լրացման կլաս
Class CreditCards
		public agreeN
		public curr
		public client
		public calcAcc
		public clientName
		public accNote
		public accNote2
		public accNote3
		public note 
		public note2
		public note3
		public office
		public section
		public accessType
		public showClientFeatures
		public showNotes
		public showAccNote
		private Sub Class_Initialize()
				agreeN = ""
				curr = ""
				client = ""
				calcAcc = ""
				clientName = ""
				accNote = ""
				accNote2 = ""
				accNote3 = ""
				note  = ""
				note2 = ""
				note3 = ""
				office = ""
				section = ""
				accessType = ""
				showClientFeatures = 0
				showNotes = 0
				showAccNote = 0
		End Sub
End Class

Function New_CreditCards()
		Set New_CreditCards = new CreditCards
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''Fill_CreditCards'''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Վարկային քարտեր պատուհանի լրացման պրոցեդուրա
'CrdCards - պատուհանի լրացման կլաս
Sub Fill_CreditCards(CrdCards)
  ' Պայմանագրի N դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "NUM", CrdCards.agreeN)
  ' Արժույթ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "CUR", CrdCards.curr)
		' Հաճախորդ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "CLIENT", CrdCards.client)
		' Հաշվարկային հաշիվ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "ACCMASK", CrdCards.calcAcc)
		' Հաճախորդի անվանում դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "CLNAME", CrdCards.clientName)
		' Հաշվի նշում դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "ACCNOTE", CrdCards.accNote)
		' Հաշվի նշում 2 դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "ACCNOTE2", CrdCards.accNote2)
		' Հաշվի նշում 3 դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "ACCNOTE3", CrdCards.accNote3)
		' Նշում դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "NOTE", CrdCards.note)
		' Նշում 2 դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "NOTE2", CrdCards.note2)
		' Նշում 3 դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "NOTE3", CrdCards.note3)
		' Գրասենյակ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "ACSBRANCH", CrdCards.office)
		' Բաժին դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "ACSDEPART", CrdCards.section)
		' Հասան-ն տիպ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "ACSTYPE", CrdCards.accessType)
		' Ցույց տալ հաճախորդների հատկանիշները դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "CheckBox", "SHOWCLI", CrdCards.showClientFeatures)
		' Ցույց տալ նշումները դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "CheckBox", "SHOWNOTES", CrdCards.showNotes)
		' Ցույց տալ հաշիվների նշումները դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "CheckBox", "SHOWACCNOTES", CrdCards.showAccNote)
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''GoTo_CreditCards'''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Վարկային քարտեր թղթապանակ մուտք գործելու պրոցեդուրա
'folderName - գտնբելու ճանապարհը
'CrdCards - պատուհանի լրացման կլաս
Sub GoTo_CreditCards(folderName, CrdCards)
		wTreeView.DblClickItem(folderName & "ì³ñÏ³ÛÇÝ ù³ñï»ñ")
		if p1.WaitVBObject("frmAsUstPar", 3000).Exists then
				Call Fill_CreditCards(CrdCards)
				Call ClickCmdButton(2, "Î³ï³ñ»É")
		else 
				Log.Error "Can't open frmAsUstPar widow.", "", pmNormal, ErrorColor
		end if
End	Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''AgreementAllOperations''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Կլաս "Պայամանագրի բոլոր գործողություններ" դիալոգային պատուհանի դաշտերը լրացնելու համար
Class AgreementAllOperations
  public startDate
  public endDate
  public agreementN
  public onlyChanges
  private Sub Class_Initialize()
    startDate = ""
    endDate = ""
    agreementN = ""
    onlyChanges = 0
  End Sub
End Class

Function New_AgreementAllOperations()
  Set New_AgreementAllOperations = new AgreementAllOperations
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''Fill_AgreementAllOperations'''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պրոցեդուրան լրացնում է "Պայամանագրի բոլոր գործողություններ" դիալոգային պատուհանի դաշտերը
'AgreeAllOperations - պատուհանում լրացվող տվյալների կլասն է -  New_AgreementAllOperations()
Sub Fill_AgreementAllOperations(AgreeAllOperations)
  ' Ժամանակահատված սկզբնական դաշտի լրացում
  Call Rekvizit_Fill("Dialog", 1, "General", "START", "![End]" & "[Del]" & AgreeAllOperations.startDate)
		' Ժամանակահատված վերջնական դաշտի լրացում
  Call Rekvizit_Fill("Dialog", 1, "General", "END", "![End]" & "[Del]" & AgreeAllOperations.endDate)
		' Պայմանագրի N դաշտի լրացում
  Call Rekvizit_Fill("Dialog", 1, "General", "NUM", AgreeAllOperations.agreementN)
		' Միայն փոփոխություները դաշտի լրացում
  Call Rekvizit_Fill("Dialog", 1, "CheckBox", "ONLYCH", AgreeAllOperations.onlyChanges)
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''GoTo_AgreementAllOperations''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Վարկային քարտեր թղթապանակ մուտք գործելու պրոցեդուրա
'folderName - գտնբելու ճանապարհը
'AgreeAllOperations - պատուհանի լրացման կլաս
Sub GoTo_AgreementAllOperations(folderName, AgreeAllOperations)
		wTreeView.DblClickItem(folderName & "ä³ÛÙ³Ý³·ñÇ µáÉáñ ·áñÍáÕáõÃÛáõÝÝ»ñ")
		if p1.WaitVBObject("frmAsUstPar", 3000).Exists then
				Call Fill_AgreementAllOperations(AgreeAllOperations)
				Call ClickCmdButton(2, "Î³ï³ñ»É")
		else 
				Log.Error "Can't open frmAsUstPar widow.", "", pmNormal, ErrorColor
		end if
End	Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''RegistryInputInformation''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Կլաս Ռեգիստրի մուտքային տեղեկատվություն դիալոգային պատուհանի դաշտերը լրացնելու համար
Class RegistryInputInformation
  public startDate
  public endDate
  public client
		public bankID
		public NumOfexpDaysExists
		public NumOfexpDays1
  public NumOfexpDays2
		public clientDataExists
		public clientData
  private Sub Class_Initialize()
    startDate = ""
    endDate = ""
    client = ""
				bankID = ""
				NumOfexpDaysExists = false
				NumOfexpDays1 = ""
		  NumOfexpDays2 = ""
				clientDataExists = false
				clientData = 0
  End Sub
End Class

Function New_RegistryInputInformation()
  Set New_RegistryInputInformation = new RegistryInputInformation
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''Fill_RegistryInputInformation''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պրոցեդուրան լրացնում է Ռեգիստրի մուտքային տեղեկատվություն և Գործողությունների դիտում դիալոգային պատուհաններիի դաշտերը
'RegistryInfo - պատուհանում լրացվող տվյալների կլասն է -  New_RegistryInputInformation()
Sub Fill_RegistryInputInformation(RegistryInfo)
  ' Ժամանակահատված սկզբնական դաշտի լարցում
  Call Rekvizit_Fill("Dialog", 1, "General", "START", "![End]" & "[Del]" & RegistryInfo.startDate)
		' Ժամանակահատված վերջնական դաշտի լրացում
  Call Rekvizit_Fill("Dialog", 1, "General", "END", "![End]" & "[Del]" & RegistryInfo.endDate)
		' Հաճախորդ դաշտի լրացում
  Call Rekvizit_Fill("Dialog", 1, "General", "CLIENT", RegistryInfo.client)
		' Նույնականացուցիչ(ID) դաշտի լրացում
  Call Rekvizit_Fill("Dialog", 1, "General", "BANKID", RegistryInfo.bankID)
		' Ժամկետանց օրերի քանակ դաշտերի լրացում
		if RegistryInfo.NumOfexpDaysExists then
				Call Rekvizit_Fill("Dialog", 1, "General", "FROM", RegistryInfo.NumOfexpDays1)
				Call Rekvizit_Fill("Dialog", 1, "General", "TO", RegistryInfo.NumOfexpDays2)
		end if
		if RegistryInfo.clientDataExists then 
		  ' Հաճախորդի տվյալներ դաշտի լրացում
				Call Rekvizit_Fill("Dialog", 1, "CheckBox", "CLIINF", RegistryInfo.clientData)
		end if
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''GoTo_RegistryInputInformation''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Ռեգիստրի մուտքային տեղեկատվություն և Գործողությունների դիտում թղթապանակների մուտք գործելու պրոցեդուրա
'folderName - գտնբելու ճանապարհը
'typeName - գործողություն անունը
'RegistryInfo - պատուհանի լրացման կլաս
Sub GoTo_RegistryInputInformation(folderName, typeName, RegistryInfo)
		wTreeView.DblClickItem(folderName & typeName)
		if p1.WaitVBObject("frmAsUstPar", 3000).Exists then
				Call Fill_RegistryInputInformation(RegistryInfo)
				Call ClickCmdButton(2, "Î³ï³ñ»É")
		else 
				Log.Error "Can't open frmAsUstPar widow.", "", pmNormal, ErrorColor
		end if
End	Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''RequestForChangeContractFields'''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պայմանագրերի դաշտերի փոփոխման հայտեր դիալոգային պատուհանի դաշտերը լրացման կլաս
Class RequestForChangeContractFields
		public state
  public startDate
  public endDate
  public performer
		public office
		public section
		private Sub Class_Initialize()
				state = ""
    startDate = ""
    endDate = ""
    performer = ""
				office = ""
				section = ""
  End Sub
End Class

Function New_RequestForChangeContractFields()
  Set New_RequestForChangeContractFields = new RequestForChangeContractFields
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''Fill_RequestForChangeContractFields'''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պրոցեդուրան լրացնում է Պայմանագրերի դաշտերի փոփոխման հայտեր դիալոգային պատուհաններիի դաշտերը
'RequestForContract - պատուհանում լրացվող տվյալների կլասն է -  New_RequestForChangeContractFields()
Sub Fill_RequestForChangeContractFields(RequestForContract)
		' Վիճակ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "DSTATE", "![End]" & "[Del]" & RequestForContract.state)
		' Ժամանակահատված սկզբնական դաշտի լրացում
  Call Rekvizit_Fill("Dialog", 1, "General", "DSDATE", "![End]" & "[Del]" & RequestForContract.startDate)
		' Ժամանակահատված վերջնական դաշտի լրացում
  Call Rekvizit_Fill("Dialog", 1, "General", "DEDATE", "![End]" & "[Del]" & RequestForContract.endDate)
		' Կատարող դաշտի լրացում
  Call Rekvizit_Fill("Dialog", 1, "General", "DUSER", RequestForContract.performer)
		' Գրասենյակ դաշտի լրացում
  Call Rekvizit_Fill("Dialog", 1, "General", "DACSBRANCH", RequestForContract.office)
		' Բաժին դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "DACSDEPART", RequestForContract.section)
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''GoTo_RequestForChangeContractFields'''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Ռեգիստրի մուտքային տեղեկատվություն և Գործողությունների դիտում թղթապանակների մուտք գործելու պրոցեդուրա
'folderName - գտնվելու ճանապարհը
'RequestForContract - պատուհանի լրացման կլաս
Sub GoTo_RequestForChangeContractFields(folderName, RequestForContract)
		wTreeView.DblClickItem(folderName & "ä³ÛÙ³Ý³·ñ»ñÇ ¹³ßï»ñÇ ÷á÷áËÙ³Ý Ñ³Ûï»ñ")
		if p1.WaitVBObject("frmAsUstPar", 3000).Exists then
				Call Fill_RequestForChangeContractFields(RequestForContract)
				Call ClickCmdButton(2, "Î³ï³ñ»É")
		else 
				Log.Error "Can't open frmAsUstPar widow.", "", pmNormal, ErrorColor
		end if
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''VerificationDocument'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Հաստատվող փաստաթղթեր (|) լրացման կլաս
Class VerificationDocument
		public DocType
    public User
    public Division
    public Department
    public View
    public FillInto

		private Sub Class_Initialize()
		     DocType = ""
         User = ""
         Division = ""
         Department = ""
         View = "VerPays"
         FillInto = "0"
		End Sub
End Class

Function New_VerificationDocument()
		Set New_VerificationDocument = new VerificationDocument
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''GoToVerificationDocument''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Հաստատվող փաստաթղթեր (|) Ֆիլտրի լրացմում
'FolderPath - գտնվելու ճանապարհը
'VerificationDoc - լրացվող արժեքների օբեկտ
Sub GoToVerificationDocument(FolderPath,VerificationDoc)
    BuiltIn.Delay(1000)
    wTreeView.DblClickItem(FolderPath)
    
	If p1.WaitVBObject("frmAsUstPar", 3000).Exists Then
        'Լրացնել "Փաստաթղթի տեսակ" դաշտը
        Call Rekvizit_Fill("Dialog",1,"General","DTYPE", VerificationDoc.DocType)
        'Լրացնել "Կատարող" դաշտը
        Call Rekvizit_Fill("Dialog",1,"General","USER", VerificationDoc.User)
        'Լրացնել "Գրասենյան" դաշտը
        Call Rekvizit_Fill("Dialog",1,"General","DIVISION", VerificationDoc.Division)
        'Լրացնել "Բաժին" դաշտը
        Call Rekvizit_Fill("Dialog",1,"General","DEPART", VerificationDoc.Department)
        'Լրացնել "Դիտելու ձև" դաշտը
        Call Rekvizit_Fill("Dialog",1,"General","SELECTED_VIEW", "^A[Del]" & VerificationDoc.View)
        'Լրացնել "Լրացնել" դաշտը
        Call Rekvizit_Fill("Dialog",1,"General","EXPORT_EXCEL", "^A[Del]" & VerificationDoc.FillInto)
            
        Call ClickCmdButton(2, "Î³ï³ñ»É")
    Else
        Log.Error "Can Not Open Verification Document Filter",,,ErrorColor      
    End If 
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''Common_SummaryOfContracts'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պայմանագրերի ամփոփում պատուհանի Ընդհանուր բաժնի լրացման կլաս
Class Common_SummaryOfContracts
		public date
		public closeDateExists
		public closeDateStart
		public closeDateEnd
		public agreeKind
		public agreeType
		public agreePaperN
		public LRCodeExist
		public LRCode
		public agreeN
		public curr
		public preferredCurr
		public client
		public clientName
		public clientInsurance
		public clientNameInsurance
		public groupExist
		public groupInsurance
		public isSignedStart
		public isSignedEnd
		public checkBoxExist
		public showWithoutExpiredPart
		public showNoWriteOffs
		public show
		public viewType
		public fill
		private sub Class_Initialize()
				date = ""
				closeDateExists = false
				closeDateStart = ""
				closeDateEnd = ""
				agreeKind = ""
				agreeType = ""
				agreePaperN = ""
				LRCodeExist = false
				LRCode = ""
				agreeN = ""
				curr = ""
				preferredCurr = ""
				client = ""
				clientName = ""
				clientInsurance = ""
				clientNameInsurance = ""
				groupExist = false
				groupInsurance = ""
				isSignedStart = ""
				isSignedEnd = ""
				checkBoxExist = false
				showWithoutExpiredPart = 0
				showNoWriteOffs = 0
				show = ""
				viewType = "MTOTAL"
				fill = "0"
		end sub
End Class

Function New_Common_SummaryOfContracts()
		Set New_Common_SummaryOfContracts = new Common_SummaryOfContracts
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''Fill_Common_SummaryOfContracts''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պայմանագրերի ամփոփում պատուհանի Ընդհանուր բաժնի լրացման պրոցեդուրա
'Common - պատուհանի լրացման կլաս
Sub Fill_Common_SummaryOfContracts(Common)
  ' Ամսաթիվ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "DATE", "![End]" & "[Del]" & Common.date)
		' Փակման ժամանակահատված դաշտերի լրացում
		if Common.closeDateExists then
				Call Rekvizit_Fill("Dialog", 1, "General", "CLOSEDATESTART", "![End]" & "[Del]" & Common.closeDateStart)
				Call Rekvizit_Fill("Dialog", 1, "General", "CLOSEDATEEND", "![End]" & "[Del]" & Common.closeDateEnd)
		end if
		' Պայմանագրի մակարդակ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "AGRKIND", Common.agreeKind)
		' Տեսակ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "TYPE", Common.agreeType)
		' Պայմ. թղթային N դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "PPRCODE", Common.agreePaperN)
		' ՎՌ կոդ(նոր) դաշտի լրացում
		if Common.LRCodeExist then
				Call Rekvizit_Fill("Dialog", 1, "General", "NEWLRCODE", Common.LRCode)
		end if
		' Պայմանագրի N դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "AGRNUM", Common.agreeN)
		' Արժույթ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "CUR", Common.curr)
		' Նախընտրելի արժույթ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "DEFCUR", Common.preferredCurr)
		' Հաճախորդ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "CLIENT", Common.client)
		' Հաճախորդի անվանում դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "NAME", Common.clientName)
		' Հաճախորդ (Ապահ. պայմ.) դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "CLISS", Common.clientInsurance)
		' Հաճախորդի անվանում (Ապահ. պայմ.) դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "CLISSNAME", Common.clientNameInsurance)
		' Խմբային (Ապահ. պայմ.) դաշտի լրացում
		if Common.groupExist then
				Call Rekvizit_Fill("Dialog", 1, "General", "CRDGROUPMASK", Common.groupInsurance)
		end if
		' Կնքված է սկզբնական դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "AGRDATESTART", Common.isSignedStart)
		' Կնքված է վերջնական դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "AGRDATEEND", Common.isSignedEnd)
		if Common.checkBoxExist then 
		  ' Գումարները առանց ժամկետանց մասի դաշտի լրացում
				Call Rekvizit_Fill("Dialog", 1, "CheckBox", "SHWITHOUTSUMJ", Common.showWithoutExpiredPart)
				' Գումարները առանց դուրսգրուների դաշտի լրացում
				Call Rekvizit_Fill("Dialog", 1, "CheckBox", "SHWITHOUTOUTSUM", Common.showNoWriteOffs)
		end if
		' Ցույց տալ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "MNCTRL1", "![End]" & "[Del]" & Common.show)
		' Դիտելու ձև դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "SELECTED_VIEW", "![End]" & "[Del]" & Common.viewType)
		' Լրացնել դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "EXPORT_EXCEL", "![End]" & "[Del]" & Common.fill)
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''Additional_SummaryOfContracts'''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պայմանագրերի ամփոփում պատուհանի Լրացուցիչ բաժնի լրացման կլաս
Class Additional_SummaryOfContracts
		public insuranceType
		public insuranceN
		public guaranteed_InsuranceType
		public guaranteed_InsuranceN
		public office
		public group 
		public accessType
		private sub Class_Initialize()
				insuranceType = ""
				insuranceN = ""
				guaranteed_InsuranceType = ""
				guaranteed_InsuranceN = ""
				office = ""
				group  = ""
				accessType = ""
		end sub
End Class

Function New_Additional_SummaryOfContracts()
		Set New_Additional_SummaryOfContracts = new Additional_SummaryOfContracts
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''Fill_Additional_SummaryOfContracts'''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պայմանագրերի ամփոփում պատուհանի Լրացուցիչ բաժնի լրացման պրոցեդուրա
'Additional - պատուհանի լրացման կլաս
Sub Fill_Additional_SummaryOfContracts(Additional)
  ' Ապահ. պայմ. տեսակ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 2, "General", "SSAGRKIND", Additional.insuranceType)
		' Ապահ. պայմ. N դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 2, "General", "CREDNUM", Additional.insuranceN)
		' Ապահ. պայմ. տիպ (Երաշխ.) դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 2, "General", "GUARTIP", Additional.guaranteed_InsuranceType)
		' Ապահ. պայմ. N (Երաշխ.) դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 2, "General", "GUARNUM", Additional.guaranteed_InsuranceN)
		' Գրասենյալ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 2, "General", "ACSBRANCH", Additional.office)
		' Բաժին դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 2, "General", "ACSDEPART", Additional.group)
		' Հասան-ն տիպ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 2, "General", "ACSTYPE", Additional.accessType)
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''SummaryOfContracts''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պայմանագրերի ամփոփում պատուհանի լրացման կլաս
Class SummaryOfContracts
		public common
		public additional
		private sub Class_Initialize()
				Set common = New_Common_SummaryOfContracts()
				Set additional = New_Additional_SummaryOfContracts()
		end sub
End Class

Function New_SummaryOfContracts()
		Set New_SummaryOfContracts = new SummaryOfContracts
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''Fill_SumaryOfContracts''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պայմանագրերի ամփոփում պատուհանի լրացման պրոցեդուրա
'SumOfContracts - պատուհանի լրացման կլաս
Sub Fill_SumaryOfContracts(SumOfContracts)
  ' Ընդհանուր բաժնի լրացում
		Call Fill_Common_SummaryOfContracts(SumOfContracts.common)
		' Լրացուցիչ բաժնի լրացում
		Call Fill_Additional_SummaryOfContracts(SumOfContracts.additional)
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''GoTo_SummaryOfContracts'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պայմանագրերի ամփոփում թղթապանակ մուտք գործելու պրոցեդուրա
'folderName - գտնբելու ճանապարհը
'typeName - փաստաթղթի անունը
'SumOfContracts - պատուհանի լրացման կլաս
Sub GoTo_SummaryOfContracts(folderName, typeName, SumOfContracts)
		wTreeView.DblClickItem(folderName & typeName)
		if p1.WaitVBObject("frmAsUstPar", 3000).Exists then
				Call Fill_SumaryOfContracts(SumOfContracts)
				Call ClickCmdButton(2, "Î³ï³ñ»É")
		else 
				Log.Error "Can't open frmAsUstPar widow.", "", pmNormal, ErrorColor
		end if
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''BankAllSecurities''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Բանկի բոլոր արժեթղթերը պատուհանի լրացման կլաս
Class BankAllSecurities
		public startDate
		public endDate
		public summaryByReleases
		private sub Class_Initialize()
				startDate = ""
				endDate = ""
				summaryByReleases = 0
		end sub
End Class

Function New_BankAllSecurities()
		Set New_BankAllSecurities = new BankAllSecurities
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''Fill_BankAllSecurities''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Բանկի բոլոր արժեթղթերը պատուհանի լրացման պրոցեդուրա
'bankSecurities - պատուհանի լրացման կլաս
Sub Fill_BankAllSecurities(bankSecurities)
  ' Ժամանակահատված սկզբնական դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "FDATE", bankSecurities.startDate)
		' Ժամանակահատված վերջնական դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "LDATE", bankSecurities.endDate)
		' Ամփոփ ըստ թողարկումների դաշտի լրացում 
		Call Rekvizit_Fill("Dialog", 1, "CheckBox", "GROUPBYNAME", bankSecurities.summaryByReleases)
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''GoTo_BankAllSecurities'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Բանկի բոլոր արժեթղթերը թղթապանակ մուտք գործելու պրոցեդուրա
'folderName - գտնբելու ճանապարհը
'bankSecurities - պատուհանի լրացման կլաս
Sub GoTo_BankAllSecurities(folderName, bankSecurities)
		wTreeView.DblClickItem(folderName & "´³ÝÏÇ µáÉáñ ³ñÅ»ÃÕÃ»ñÁ")
		if p1.WaitVBObject("frmAsUstPar", 3000).Exists then
				Call Fill_BankAllSecurities(bankSecurities)
				Call ClickCmdButton(2, "Î³ï³ñ»É")
		else 
				Log.Error "Can't open frmAsUstPar widow.", "", pmNormal, ErrorColor
		end if
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''BankOwnSecurities''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Բանկի սեփական արժեթղթերը պատուհանի լրացման կլաս
Class BankOwnSecurities
		public date
		public yieldCurveDate
		public agreeN
		public issue
		public showWithoutRepo
		public summaryByReleases
		private sub Class_Initialize()
				date = ""
				yieldCurveDate = ""
				agreeN = ""
				issue = ""
				showWithoutRepo = 0
				summaryByReleases = 0
		end sub
End Class

Function New_BankOwnSecurities()
		Set New_BankOwnSecurities = new BankOwnSecurities
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''Fill_BankOwnSecurities''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Բանկի սեփական արժեթղթերը պատուհանի լրացման պրոցեդուրա
'bankSecurities - պատուհանի լրացման կլաս
Sub Fill_BankOwnSecurities(bankSecurities)
  ' Ժամանակահատված դաշտի լրացում 
		Call Rekvizit_Fill("Dialog", 1, "General", "FDATE", bankSecurities.date)
		' Եկամտաբերության կորի ամսաթիվ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "FPCINCDATE", bankSecurities.yieldCurveDate)
		' Պայմանագրի N դաշտի լրացում 
		Call Rekvizit_Fill("Dialog", 1, "General", "CODE", bankSecurities.agreeN)
		' Թողարկում դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "SECNAME", bankSecurities.issue)
		' Ցույց տալ առանց հակադարձ ռեպոյով վաճ. մասի դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "CheckBox", "SHOWWITHOUTHAKREPSUM", bankSecurities.showWithoutRepo)
		' Ամփոփ ըստ թողարկուների դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "CheckBox", "GROUPBYNAME", bankSecurities.summaryByReleases)
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''GoTo_BankOwnSecurities'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Բանկի սեփական արժեթղթերը թղթապանակ մուտք գործելու պրոցեդուրա
'folderName - գտնբելու ճանապարհը
'typeName - փաստաթղթի անունը
'bankSecurities - պատուհանի լրացման կլաս
Sub GoTo_BankOwnSecurities(folderName, bankSecurities)
		wTreeView.DblClickItem(folderName & "´³ÝÏÇ ë»÷³Ï³Ý ³ñÅ»ÃÕÃ»ñÁ")
		if p1.WaitVBObject("frmAsUstPar", 3000).Exists then
				Call Fill_BankOwnSecurities(bankSecurities)
				Call ClickCmdButton(2, "Î³ï³ñ»É")
		else 
				Log.Error "Can't open frmAsUstPar widow.", "", pmNormal, ErrorColor
		end if
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''Genenral_DepositesSumOfCont'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պայմանագրերի ամփոփում պատուհանի Ընդհանուր բաժնի լրացման կլաս
Class Genenral_DepositesSumOfCont
		public date
		public agreeLevel
		public agreeType 
		public curr
		public agreeN
		public agreePaperN
		public LRCode
		public client
		public clientName
		public note
		public note2
		public note3
		public office 
		public department 
		public accessType
		public closeAgreeExists
		public showClosed
		public showNotAllClosed
		public riskIndicator
		public startSealed
		public endSealed
		public repayDateExists
		public repayDateStart
		public repayDateEnd
		public closeDateExists
		public closedStart
		public closedEnd
		public preferedCurr
		public circulatingInfo
		public started
		public amountsWithoutOverPart
		public amountsWithoutWriteOffs
		public showType
		public fill
		private sub Class_Initialize()
				date = ""
				agreeLevel = "1"
				agreeType = ""
				curr = ""
				agreeN = ""
				agreePaperN = ""
				LRCode = ""
				client = ""
				clientName = ""
				note = ""
				note2 = ""
				note3 = ""
				office = ""
				department = ""
				accessType = ""
				closeAgreeExists = false
				showClosed = 0
				showNotAllClosed = 0
				riskIndicator = ""
				startSealed = ""
				endSealed = ""
				repayDateExists = false
				repayDateStart = ""
				repayDateEnd = ""
				closeDateExists = False
				closedStart = ""
				closedEnd = ""
				preferedCurr = ""
				circulatingInfo = 0
				started = ""
				amountsWithoutOverPart = 0
				amountsWithoutWriteOffs = 0
				showType = "AGRTOTL"
				fill = "0"
		end sub
End Class

Function New_Genenral_DepositesSumOfCont()
		Set New_Genenral_DepositesSumOfCont = new Genenral_DepositesSumOfCont
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''Fill_Genenral_DepositesSumOfCont''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պայմանագրերի ամփոփում պատուհանի Ընդհանուր բաժնի լրացման պրոցեդուրա
'general - Ընդհանուր բաժնի լրացման կլաս
Sub Fill_Genenral_DepositesSumOfCont(general)
  ' Ամսաթիվ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "RDATE", "![End]" & "[Del]" & general.date) 
		' Պայմանագրի մակարդակ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "LEVEL", "![End]" & "[Del]" & general.agreeLevel) 
		' Տեսակ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "AGRKIND", general.agreeType) 
		' Արժույթ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "CUR", general.curr) 
		' Պայմանագրի N դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "NUM", general.agreeN) 
		' Թղթային N դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "PPRCODE", general.agreePaperN) 
		' ՎՌ կոդ (նոր) դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "NEWLRCODE", general.LRCode) 
		' Հաճախորդ դաշտիր լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "CLIENT", general.client) 
		' Անվանում դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "CLIENTNAME", general.clientName) 
		' Նշում դաշտի լարցում
		Call Rekvizit_Fill("Dialog", 1, "General", "NOTE", general.note) 
		' Նշում 2 դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "NOTE2", general.note2) 
		' Նշում 3 դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "NOTE3", general.note3) 
		' Գրասենյակ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "ACSBRANCH", general.office) 
		' Բաժին դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "ACSDEPART", general.department) 
		' Հ. տիպ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "ACSTYPE", general.accessType)
		if general.closeAgreeExists then 
		  ' Ցույց տալ փակվածները դաշտի լրացում
				Call Rekvizit_Fill("Dialog", 1, "CheckBox", "CLOSEAGRS", general.showClosed) 
				' Ոչ լրիվ փակ. դաշտի լրացում
				Call Rekvizit_Fill("Dialog", 1, "CheckBox", "NOTFULLCLOSEAGRS", general.showNotAllClosed) 
		end if
		'Ռիսկի դասիչ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "RISK", general.riskIndicator) 
		' Կնքված է սկզբնական դաշտի լարցում
		Call Rekvizit_Fill("Dialog", 1, "General", "AGRDATESTART", general.startSealed) 
		' Կնքված է վերջնայան դաշտի լրացում 
		Call Rekvizit_Fill("Dialog", 1, "General", "AGRDATEEND", general.endSealed) 
		' Մարման ժամկետ դաշտերի լարցում
		if general.repayDateExists then 
				Call Rekvizit_Fill("Dialog", 1, "General", "MARDATESTART", general.repayDateStart) 
				Call Rekvizit_Fill("Dialog", 1, "General", "MARDATEEND", general.repayDateEnd) 
		end if
		' Փակման ժամկետ դաշտերի լրացում
		if general.closeDateExists then 
				Call Rekvizit_Fill("Dialog", 1, "General", "CLOSEDATESTART", general.closedStart) 
				Call Rekvizit_Fill("Dialog", 1, "General", "CLOSEDATEEND", general.closedEnd)
		end if
		' Նախընտրելի արժույթ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "DEFAULTCUR", general.preferedCurr) 
		' Շրջանառու ինֆորմացիա դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "CheckBox", "CHKTURN", general.circulatingInfo)
		if  general.circulatingInfo = 1 then
		  ' Սկսած դաշտի լրացում
				Call Rekvizit_Fill("Dialog", 1, "General", "TURNDATESTART", general.started) 
		end if
		' Գումարները առանց ժամկետանց մասի դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "CheckBox", "SHWITHOUTSUMJ", general.amountsWithoutOverPart) 
		' Գումարները առանց դուրսգրումների դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "CheckBox", "SHWITHOUTOUTSUM", general.amountsWithoutWriteOffs) 
		' Դիտելու ձև դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "SELECTED_VIEW", "![End]" & "[Del]" & general.showType) 
		' Լրացնել դաշտի լարցում
		Call Rekvizit_Fill("Dialog", 1, "General", "EXPORT_EXCEL", "![End]" & "[Del]" & general.fill) 
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''Show_DepositesSumOfCont''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պայմանագրերի ամփոփում պատուհանի Ցույց տալ բաժնի լրացման կլաս
Class Show_DepositesSumOfCont
		public clientMainData
		public agreeMainData
		public mainAmounts
		public mainDate
		public accMainData
		public overlimitAmounts
		public notes
		public riskyInformation
		public clientOtherData
		public agreeOtherData
		public otherAmounts
		public otherDates
		public addData
		public writhdrawnAmounts
		public depositeData
		public addAmounts
		public addDates
		private sub Class_Initilize()
				clientMainData = 0
				agreeMainData = 0 
				mainAmounts = 0 
				mainDate = 0 
				accMainData = 0 
				overlimitAmounts = 0 
				notes = 0
				riskyInformation = 0
				clientOtherData = 0
				agreeOtherData = 0
				otherAmounts = 0 
				otherDates = 0
				addData = 0
				writhdrawnAmounts = 0
				depositeData = 0
				addAmounts = 0 
				addDates = 0
				end sub
End Class

Function New_Show_DepositesSumOfCont()
		Set New_Show_DepositesSumOfCont = new Show_DepositesSumOfCont
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''Fill_Show_DepositesSumOfCont'''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պայմանագրերի ամփոփում պատուհանի Ցույց տալ բաժնի լրացման պրոցեդուրա
'show - Ցույց տալ բաժնի լրացման կլաս
Sub Fill_Show_DepositesSumOfCont(show)
  ' Հաճախորդի հիմնական տվյալներ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 2, "CheckBox", "CL1", show.clientMainData) 
		' Հաճախորդի այլ տվյալներ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 2, "CheckBox", "CL2", show.clientOtherData)
		' Պայմանագրի հիմնական տվյալներ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 2, "CheckBox", "AGR1", show.agreeMainData) 
		' Պայմանագրի այլ տվյալներ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 2, "CheckBox", "AGR2", show.agreeOtherData)
		' Պայմանագրի գրավի տվյալներ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 2, "CheckBox", "AGR3", show.depositeData) 
		' Հիմնական գումրաներ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 2, "CheckBox", "SUM1", show.mainAmounts) 
		' Այլ գումարներ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 2, "CheckBox", "SUM2", show.otherAmounts) 
		' Լրացուցիչ գումրաներ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 2, "CheckBox", "SUM3", show.addAmounts) 
		' Հիմնական ամսաթվեր դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 2, "CheckBox", "DAT1", show.mainDate) 
		' Այլ ամսաթվեր դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 2, "CheckBox", "DAT2", show.otherDates)
		' Լրացուցիչ ամսաթվեր դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 2, "CheckBox", "DAT3", show.addDates) 
		' Հաշվապահական հիմնական տվյալներ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 2, "CheckBox", "ACC1", show.accMainData) 
		' Լրացուցիչ տվյալներ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 2, "CheckBox", "ACC2", show.addData)  
		' Ժամկետանց գումարներ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 2, "CheckBox", "SUMOS", show.overlimitAmounts)
		' Դուրսգրված գումարներ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 2, "CheckBox", "SUMJS", show.writhdrawnAmounts) 
		' Նշումներ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 2, "CheckBox", "NOTES", show.notes)
		' Ռիսկային ինֆորմացիա դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 2, "CheckBox", "RSK", show.riskyInformation)
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''Deposites_SumOfContracts'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պայմանագրերի ամփոփում պատուհանի լրացման կլաս
Class Deposites_SumOfContracts
		public general 
		public show
		private sub Class_Initialize()
				Set general = New_Genenral_DepositesSumOfCont()
				Set show = New_Show_DepositesSumOfCont()
		end sub
End Class

Function New_Deposites_SumOfContracts()
		Set New_Deposites_SumOfContracts = new Deposites_SumOfContracts
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''Fill_DepositesSumOfCont''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պայմանագրերի ամփոփում պատուհանի լրացման պրոցեդուրա
'depositesSumOfCont - պատուհանի լրացման կլաս
Sub Fill_DepositesSumOfCont(depositesSumOfCont)
		Call Fill_Genenral_DepositesSumOfCont(depositesSumOfCont.general)
		Call Fill_Show_DepositesSumOfCont(depositesSumOfCont.show)
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''GoTo_BankOwnSecurities'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Բանկի սեփական արժեթղթերը թղթապանակ մուտք գործելու պրոցեդուրա
'folderName - գտնբելու ճանապարհը
'typeName - փաստաթղթի անունը
'bankSecurities - պատուհանի լրացման կլաս
Sub GoTo_DepositesSumOfCont(folderName, typeName, depositesSumOfCont)
		wTreeView.DblClickItem(folderName & typeName)
		if p1.WaitVBObject("frmAsUstPar", 3000).Exists then
				Call Fill_DepositesSumOfCont(depositesSumOfCont)
				Call ClickCmdButton(2, "Î³ï³ñ»É")
		else 
				Log.Error "Can't open frmAsUstPar widow.", "", pmNormal, ErrorColor
		end if
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''DepositesSumOfCont_Cached'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պայմանագրերի ամփոփում (Քեշավորված) պատուհանի լրացման կլաս
Class DepositesSumOfCont_Cached
		public date
		public agreeLevel
		public curr
		public agreeType 
		public agreeN
		public agreePaperN
		public creditCode
		public LRCode
		public client
		public clientName
		public note
		public note2
		public note3
		public office 
		public department 
		public accessType
		public startSealed
		public endSealed
		public closedStart
		public closedEnd
		public show 
		public showOnlyOverlimAgree
		public accauntingData
		public clientData
		public notes
		public circulatingInfo
		public started
		public equivalentByCurr
		public showExportedData
		public amountWithoutOverPart
		public showType
		public fill
		private sub Class_Initialize()
						date = ""
						agreeLevel = "1"
						curr = ""
						agreeType = ""
						agreeN = ""
						agreePaperN = ""
						creditCode = ""
						LRCode = ""
						client = ""
						clientName = ""
						note = ""
						note2 = ""
						note3 = ""
						office = ""
						department = ""
						accessType = ""
						startSealed = ""
						endSealed = ""
						closedStart = ""
						closedEnd = ""
						show = "1"
						showOnlyOverlimAgree  = 0
						accauntingData = 0
						clientData = 0
						notes = 0
						circulatingInfo = 0
						started = ""
						equivalentByCurr = ""
						showExportedData = 0
						amountWithoutOverPart = 0
						showType = "AGRSINFO"
						fill = "0"
		end sub
End Class

Function New_DepositesSumOfCont_Cached()
		Set New_DepositesSumOfCont_Cached = new DepositesSumOfCont_Cached
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''Fill_DepositesSumOfCont_Cached''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պայմանագրերի ամփոփում (Քեշավորված) պատուհանի լրացման պրոցեդուրա
'cached - պատուհանի լրացման կլաս
Sub Fill_DepositesSumOfCont_Cached(cached)
  ' Ամսաթիվ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "RDATE", "![End]" & "[Del]" & cached.date) 
		' Պայմանագրի մակարդակ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "LEVEL", "![End]" & "[Del]" & cached.agreeLevel) 
		' Արժույթ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "CUR", cached.curr) 
		' Տեսակ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "AGRKIND", cached.agreeType) 
		' Պայմանագրի N դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "NUM", cached.agreeN) 
		' Թղթային N դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "PPRCODE", cached.agreePaperN) 
		' Վարկային կոդ դաշտի լրացում
'		Call Rekvizit_Fill("Dialog", 1, "General", "CRDTCODE", cached.creditCode) 
		' ՎՌ կոդ (նոր) դաշտի լրացում 
		Call Rekvizit_Fill("Dialog", 1, "General", "NEWLRCODE", cached.LRCode) 
		' Հաճախորդ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "CLIENT", cached.client) 
		' Անվանում դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "CLIENTNAME", cached.clientName) 
		' Նշում դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "NOTE", cached.note) 
		' Նշում 2 դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "NOTE2", cached.note2) 
		' Նշում 3 դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "NOTE3", cached.note3) 
		' Գրասենյակ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "ACSBRANCH", cached.office) 
		' Բաժին դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "ACSDEPART", cached.department) 
		' Հ. տիպ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "ACSTYPE", cached.accessType) 
		' Կնքված է սկզբնական դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "AGRDATESTART", cached.startSealed) 
		' Կնքված է վերջնական դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "AGRDATEEND", cached.endSealed) 
		' Փակված է սկզբնական դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "CLOSEDATESTART", cached.closedStart) 
		' Փակված է վերջնական դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "CLOSEDATEEND", cached.closedEnd) 
		' Ցույց տալ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "SHOWFOLOWING", "![End]" & "[Del]" & cached.show) 
		' Ցույց տալ միայն ժամկետանց պայմ.-երը դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "CheckBox", "SHOWONLYOVERDUES", cached.showOnlyOverlimAgree)
		' Հաշվապահական տվյալներ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "CheckBox", "ACC", cached.accauntingData)
		' Հաճախորդների տվյալներ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "CheckBox", "CL", cached.clientData)
		' Նշումներ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "CheckBox", "NOTES", cached.notes)
		' Շրջանառու ինֆորմացիա դաշտի լրացում 
		Call Rekvizit_Fill("Dialog", 1, "CheckBox", "CHKTURN", cached.circulatingInfo)
		if  cached.circulatingInfo = 1 then
		  ' Սկսած դաշտի լրացում
				Call Rekvizit_Fill("Dialog", 1, "General", "TURNDATESTART", cached.started) 
		end if
		' Համարժեք ըստ արժույթի դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "DEFAULTCUR", cached.equivalentByCurr) 
		' Ցույց տալ արտահանված տվյալներ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "CheckBox", "SHOWIMPDATA", cached.showExportedData) 
		' Գում. առանց ժամկետանց մասի դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "CheckBox", "SHWITHOUTSUMJ", cached.amountWithoutOverPart) 
		' Դիտելու ձև դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "SELECTED_VIEW", "![End]" & "[Del]" & cached.showType) 
		' Լրացնել դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "EXPORT_EXCEL", "![End]" & "[Del]" & cached.fill) 
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''GoTo_DepositesSumOfCont_Cached''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պայմանագրերի ամփոփում (Քեշավորված) թղթապանակ մուտք գործելու պրոցեդուրա
'folderName - գտնբելու ճանապարհը
'typeName - փաստաթղթի անունը
'cached - պատուհանի լրացման կլաս
Sub GoTo_DepositesSumOfCont_Cached(folderName, typeName, cached)
		wTreeView.DblClickItem(folderName & typeName)
		if p1.WaitVBObject("frmAsUstPar", 3000).Exists then
				Call Fill_DepositesSumOfCont_Cached(cached)
				Call ClickCmdButton(2, "Î³ï³ñ»É")
		else 
				Log.Error "Can't open frmAsUstPar widow.", "", pmNormal, ErrorColor
		end if
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''SummaryOfOperations'''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Գործողությունների ամփոփում պատուհանի լրացման կլաս
Class SummaryOfOperations
		public agreeLevel
		public startDate
		public endDate
		public agreeType
		public agreeN
		public agreePaperN
		public curr
		public pereferredCurr
		public client
		public clientName
		public summaryByAgree
		public summaryByDate
		public operationType
		public note
		public note2
		public note3
		public agreeOffice
		public agreeDepartment
		public accessType
		public showType
		public fill
		private sub Class_Initialize()
				agreeLevel = "1"
				startDate = ""
				endDate = ""
				agreeType = ""
				agreeN = ""
				agreePaperN = ""
				curr = ""
				pereferredCurr = ""
				client = ""
				clientName = ""
				summaryByAgree = 0
				summaryByDate = 0
				operationType = ""
				note = ""
				note2 = ""
				note3 = ""
				agreeOffice = ""
				agreeDepartment = ""
				accessType = ""
				showType = "DEALTOT"
				fill = "0"
		end sub
End Class

Function New_SummaryOfOperations()
		Set New_SummaryOfOperations = new SummaryOfOperations
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''Fill_SummaryOfOperations'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Գործողությունների ամփոփում պատուհանի լրացման պրոցեդուրա
'summOfOper - պատուհանի լրացման կլաս
Sub Fill_SummaryOfOperations(summOfOper)
  ' Պայմանագրի մակարդակ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "LEVEL", "![End]" & "[Del]" & summOfOper.agreeLevel)
		' Ժամանակահատված սկզբնական դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "FDATE", "![End]" & "[Del]" & summOfOper.startDate)
		' Ժամանակահատված վերջնական դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "LDATE", "![End]" & "[Del]" & summOfOper.endDate)
		' Պայմանագրի տեսակ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "AGRKIND", summOfOper.agreeType)
		' Պայմանագրի N դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "NUM", summOfOper.agreeN)
		' Պայմ. թղթային N դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "PPRCODE", summOfOper.agreePaperN)
		' Արժույթ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "CUR", summOfOper.curr)
		' Նախընտրելի արժույթ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "DEFAULTCUR", summOfOper.pereferredCurr)
		' Հաճախորդ դաշտի լրացում 
		Call Rekvizit_Fill("Dialog", 1, "General", "CLIENT", summOfOper.client)
		' Հաճախորդի անվանում դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "NAME", summOfOper.clientName)
		' Ամփոփ ըստ պայմանագրերի դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "CheckBox", "SHCONSOLFORAGR", summOfOper.summaryByAgree)
		' Ամփոփ ըստ ամսաթվերի դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "CheckBox", "SHCONSOLFORDATE", summOfOper.summaryByDate)
		' Գործողության տեսակ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "DEALTYPE", summOfOper.operationType)
		' Նշում դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "NOTE", summOfOper.note)
		' Նշում 2 դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "NOTE2", summOfOper.note2)
		' Նշում 3 դաշտի լրացում 
		Call Rekvizit_Fill("Dialog", 1, "General", "NOTE3", summOfOper.note3)
		' Պայմ. գրասենյակ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "ACSBRANCH", summOfOper.agreeOffice)
		' Պայմ. բաժին դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "ACSDEPART", summOfOper.agreeDepartment)
		' Հասան-ն տիպ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "ACSTYPE", summOfOper.accessType)
		' Դիտելու ձև դաշտի լրացում 
		Call Rekvizit_Fill("Dialog", 1, "General", "SELECTED_VIEW", "![End]" & "[Del]" & summOfOper.showType)
		' Լրացում դատշի լրացում 
		Call Rekvizit_Fill("Dialog", 1, "General", "EXPORT_EXCEL", "![End]" & "[Del]" & summOfOper.fill)
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''GoTo_SummaryOfOperations'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Գործողությունների ամփոփում թղթապանակ մուտք գործելու պրոցեդուրա
'folderName - գտնբելու ճանապարհը
'summOfOper - պատուհանի լրացման կլաս
Sub GoTo_SummaryOfOperations(folderName, summOfOper)
		wTreeView.DblClickItem(folderName & "¶áñÍáÕáõÃÛáõÝÝ»ñÇ ³Ù÷á÷áõÙ")
		if p1.WaitVBObject("frmAsUstPar", 3000).Exists then
				Call Fill_SummaryOfOperations(summOfOper)
				Call ClickCmdButton(2, "Î³ï³ñ»É")
		else 
				Log.Error "Can't open frmAsUstPar widow.", "", pmNormal, ErrorColor
		end if
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''RepaymentNewsletter'''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Մարումերի տեղեկագիր պատուհանի լրացման կլաս
Class RepaymentNewsletter
		public startDate
		public endDate
		public agreeType
		public agreeN
		public curr
		public client 
		public clientName
		public note
		public note2
		public note3
		public office
		public department
		public accessType
		public showType
		public fill
		private sub Class_Initialize()
				startDate = ""
				endDate = ""
				agreeType = ""
				agreeN = ""
				curr = ""
				client = ""
				clientName = ""
				note = ""
				note2 = ""
				note3 = ""
				office = ""
				department = ""
				accessType = ""
				showType  = "REPAYMNT"
				fill = "0"
		end sub
End Class

Function New_RepaymentNewsletter()
		Set New_RepaymentNewsletter = new RepaymentNewsletter
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''Fill_RepaymentNewsletter'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Մարումերի տեղեկագիր պատուհանի լրացման պրոցեդուրա
'repayNews - պատուհանի լրացման կլաս
Sub Fill_RepaymentNewsletter(repayNews)
  ' Ժամանակահատված սկզբնական դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "FDATE", "![End]" & "[Del]" & repayNews.startDate)
		' Ժամանակահատված վերջնական դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "LDATE", "![End]" & "[Del]" & repayNews.endDate)
		' Պայմանագրի տեսակ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "AGRKIND", repayNews.agreeType)
		' Պայմանագրի N դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "AGRNUM", repayNews.agreeN)
		' Արժույթ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "AGRCUR", repayNews.curr)
		' Հաճախորդ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "CLIENT", repayNews.client)
		' Հաճախորդի անվանում դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "NAME", repayNews.clientName)
		' Նշում դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "NOTE", repayNews.note)
		' Նշում 2 դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "NOTE2", repayNews.note2)
		' Նշում 3 դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "NOTE3", repayNews.note3)
		' Գասենյակ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "ACSBRANCH", repayNews.office)
		' Բաժին դաշտի լարցում
		Call Rekvizit_Fill("Dialog", 1, "General", "ACSDEPART", repayNews.department)
		' Հասան-ն տիպ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "ACSTYPE", repayNews.accessType)
		' Դիտելու ձև դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "SELECTED_VIEW", "![End]" & "[Del]" & repayNews.showType)
		' Լրացնել դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "EXPORT_EXCEL", "![End]" & "[Del]" & repayNews.fill)
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''GoTo_RepaymentNewsletter'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Մարումերի տեղեկագիր թղթապանակ մուտք գործելու պրոցեդուրա
'folderName - գտնբելու ճանապարհը
'repayNews - պատուհանի լրացման կլաս
Sub GoTo_RepaymentNewsletter(folderName, repayNews)
		wTreeView.DblClickItem(folderName & "Ø³ñáõÙÝ»ñÇ ï»Õ»Ï³·Çñ")
		if p1.WaitVBObject("frmAsUstPar", 3000).Exists then
				Call Fill_RepaymentNewsletter(repayNews)
				Call ClickCmdButton(2, "Î³ï³ñ»É")
		else 
				Log.Error "Can't open frmAsUstPar widow.", "", pmNormal, ErrorColor
		end if
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''ContractsFilterMini'''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պայմանագրեր (քիչ դաշտերով ֆիլտրի) պատուհանի լրացման կլաս
Class ContractsFilterMini
		public agreeLevel
		public agreeType
		public agreeN
		public agreePaperN
		public name
		public note
		public note2
		public note3
		public showAccounts
		public showClientsSecurities
		public showClosed
		public division
		public department
		public accessType
		private sub Class_Initialize
		  agreeLevel = "1"
				agreeType = ""
				agreeN = ""
				agreePaperN = ""
				name = ""
				note = ""
				note2 = ""
				note3 = ""
				showAccounts = 0
				showClientsSecurities = 0
				showClosed = 0
				division = ""
				department = ""
				accessType = ""
		end sub  
End Class

Function New_ContractsFilterMini()
    Set New_ContractsFilterMini = new ContractsFilterMini      
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''Fill_ContractsFilterMini'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պայմանագրեր (քիչ դաշտերով ֆիլտրի) պատուհանի լրացման պրոցեդուրա
'Contract - պատուհանի լրացման կլաս
Sub Fill_ContractsFilterMini(Contract)
    'Լրացնում է "ä³ÛÙ³Ý³·ñÇ Ù³Ï³ñ¹³Ï" դաշտը
    Call Rekvizit_Fill("Dialog", 1, "General", "LEVEL",  "![End]" & "[Del]" & Contract.agreeLevel)
    'Լրացնում է "ä³ÛÙ³Ý³·ñÇ ï»ë³Ï" դաշտը
    Call Rekvizit_Fill("Dialog", 1, "General", "AGRKIND", Contract.agreeType)
    'Լրացնում է "ä³ÛÙ³Ý³·ñÇ N" դաշտը
    Call Rekvizit_Fill("Dialog", 1, "General", "NUM", Contract.agreeN)
    'Լրացնում է "Պայմ.թղթային N" դաշտը
    Call Rekvizit_Fill("Dialog", 1, "General", "PPRCODE", Contract.agreePaperN)
    'Լրացնում է "Անվանում" դաշտը
    Call Rekvizit_Fill("Dialog", 1, "General", "AGRNAME", Contract.name)
    'Լրացնում է "Նշում" դաշտը
    Call Rekvizit_Fill("Dialog", 1, "General", "NOTE", Contract.note)
    'Լրացնում է "Նշում2" դաշտը
    Call Rekvizit_Fill("Dialog", 1, "General", "NOTE2", Contract.note2)
    'Լրացնում է "Նշում3" դաշտը
    Call Rekvizit_Fill("Dialog", 1, "General", "NOTE3", Contract.note3)
    'Լրացնում է "Ցույց տալ հաշիվները" դաշտը
    Call Rekvizit_Fill("Dialog", 1, "CheckBox", "SHOWACCS", Contract.showAccounts)
				'Լրացնում է "Ցույց տալ Հաճախորդների արժեթղթերը" դաշտը
    Call Rekvizit_Fill("Dialog", 1, "CheckBox", "SHOWCLISECS", Contract.showClientsSecurities)
    'Լրացնում է "Ցույց տալ փակվածները" դաշտը
    Call Rekvizit_Fill("Dialog", 1, "CheckBox", "CLOSE", Contract.showClosed)
    'Լրացնում է "Գրասենյակ" դաշտը
    Call Rekvizit_Fill("Dialog", 1, "General", "ACSBRANCH", Contract.division)
    'Լրացնում է "Բաժին" դաշտը
    Call Rekvizit_Fill("Dialog", 1, "General", "ACSDEPART", Contract.department)
    'Լրացնում է "Հասան-ն տիպ" դաշտը
    Call Rekvizit_Fill("Dialog", 1, "General", "ACSTYPE", Contract.accessType)
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''GoTo_ContractsFilterMini'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պայմանագրեր (քիչ դաշտերով ֆիլտրի) թղթապանակ մուտք գործելու պրոցեդուրա
'folderName - գտնբելու ճանապարհը
'Contract - պատուհանի լրացման կլաս
Sub GoTo_ContractsFilterMini(folderName, Contract)
		wTreeView.DblClickItem(folderName & "ä³ÛÙ³Ý³·ñ»ñ")
		if p1.WaitVBObject("frmAsUstPar", 3000).Exists then
				Call Fill_ContractsFilterMini(Contract)
    Call ClickCmdButton(2, "Î³ï³ñ»É")
		else 
				Log.Error "Can't open frmAsUstPar widow.", "", pmNormal, ErrorColor
		end if
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''OperationsViewMini'''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Գործողությունների դիտում պատուհանի լրացման կլաս
Class OperationsViewMini
		public agreeLevelExists
		public agreeLevel
		public startDate
		public endDate
		public agreeN
		public clientExists
		public client
		public reverseRepoAgreeExists
		public reverseRepoAgree
		public performer
		public note
		public note2
		public note3
		private sub Class_Initialize()
				agreeLevelExists = false
				agreeLevel = "1"				
				startDate = ""
				endDate = ""
				agreeN = ""
				clientExists = false
				client = ""
				reverseRepoAgreeExists = false
				reverseRepoAgree = ""
				performer = ""
				note = ""
				note2 = ""
				note3 = ""
		end sub
End Class

Function New_OperationsViewMini()
		Set New_OperationsViewMini = new OperationsViewMini
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''Fill_OperationsViewMini'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Գործողությունների դիտում պատուհանի լրացման պրոցեդուրա
'opertationsView - պատուհանի լրացման կլաս
Sub Fill_OperationsViewMini(operationsView)
		if operationsView.agreeLevelExists then
		  ' Պայմանագրի մակարդակ դաշտի լրացում
				Call Rekvizit_Fill("Dialog", 1, "General", "LEVEL", "![End]" & "[Del]" & operationsView.agreeLevel)
		end if
		' Ժամանակահատված սկզբնական դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "START", "![End]" & "[Del]" & operationsView.startDate)
		' Ժամանակահատված վերջնական դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "END", "![End]" & "[Del]" & operationsView.endDate)
		' Պայմանագրի N դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "AGR", operationsView.agreeN)
		if operationsView.clientExists then
		  ' Հաճախորդ դաշտի լրացում
				Call Rekvizit_Fill("Dialog", 1, "General", "CLIENT", operationsView.client)
		end if
		if operationsView.reverseRepoAgreeExists then
		  ' Հակադարձ ռեպո համաձայանգիր դաշտի լրացում
				Call Rekvizit_Fill("Dialog", 1, "General", "HAKREPAGR", operationsView.reverseRepoAgree)
		end if
		' Կատարող դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "USER", operationsView.performer)
		' Նշում դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "NOTE", operationsView.note)
		' Նշում 2 դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "NOTE2", operationsView.note2)
		' Նշում 3 դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "NOTE3", operationsView.note3)
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''GoTo_OperationsViewMini'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Գործողությունների դիտում թղթապանակ մուտք գործելու պրոցեդուրա
'folderName - գտնբելու ճանապարհը
'operationsView - պատուհանի լրացման կլաս
Sub GoTo_OperationsViewMini(folderName, typeName, operationsView)
		wTreeView.DblClickItem(folderName & typeName)
		if p1.WaitVBObject("frmAsUstPar", 3000).Exists then
				Call Fill_OperationsViewMini(operationsView)
    Call ClickCmdButton(2, "Î³ï³ñ»É")
		else 
				Log.Error "Can't open frmAsUstPar widow.", "", pmNormal, ErrorColor
		end if
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''GoTo_ChoosedTab''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'äñáó»¹áõñ³Ý ï»Õ³÷áËíáõÙ ¿ num Ñ³Ù³ñáí Ã³µ
'num - Ã³µÇ Ñ³Ù³ñÁ
Sub GoTo_ChoosedTab(num)
		Dim wTabStrip
		Set wTabStrip = wMDIClient.vbObject("frmASDocForm").vbObject("TabStrip")
  wTabStrip.SelectedItem = wTabStrip.Tabs(num)
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''InsuranceAgreeN'''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Ստացված գրավ՝ Արժեթղթերի "Պայմանագրի N" դաշտի լրացման կլաս
Class InsuranceAgreeN 
		public agreeLevel
		public subsystem 
		public agreeN
		public agreePaperN
		public client
		public clientName
		public searchedValue
		private sub Class_Initialize()
				agreeLevel = "1"
				subsystem = "C1"
				agreeN = ""
				agreePaperN = ""
				client = ""
				clientName = ""
				searchedValue = ""
		end sub 
End Class

Function New_InsuranceAgreeN()
		Set New_InsuranceAgreeN = new InsuranceAgreeN
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''Fill_RelatedAgreeWindow'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պրոցեդուրան լրացնում է "Կապակցվող պայմանագիր" պատուհանը
'InsurAgreeN - Ստացված գրավ՝ Արժեթղթերի "Պայմանագրի N" դաշտի լրացման կլաս
'formType - պատուհանի տեսակը
Sub Fill_RelatedAgreeWindow(InsurAgreeN, formType)
		If p1.VBObject("frmAsUstPar").Exists Then 
		  ' Պայմանագրի մակարդակ դաշտի լրացում
				Call Rekvizit_Fill(formType, 1, "General", "AGRLEVEL", InsurAgreeN.agreeLevel)
				' Ենթահամակարգ դաշտի լրացում
				Call Rekvizit_Fill(formType, 1, "General", "SSSUBSYS", InsurAgreeN.subsystem)
				' Պայմանագրի N դաշտի լրացում
				Call Rekvizit_Fill(formType, 1, "General", "AGRNUM", InsurAgreeN.agreeN)
				' Պայմանագրի թղթային N դաշտի լրացում
				Call Rekvizit_Fill(formType, 1, "General", "PPRCODE", InsurAgreeN.agreePaperN)
				' Հաճախորդ դաշտի լրացում
				Call Rekvizit_Fill(formType, 1, "General", "CLICODE", InsurAgreeN.client)
				' Հաճախորդի անվանում դաշտի լրացում 
				Call Rekvizit_Fill(formType, 1, "General", "CLINAME", InsurAgreeN.clientName)
		Else
				Log.Error "Can't open frmAsUstPar window.", "", pmNormal, ErrorColor
		End If
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''Fill_InsuranceAgreeN'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պրոցեդուրան լրացնում է "Պայմանագրի N" պատուհանը
'InsurAgreeN - Ստացված գրավ՝ Արժեթղթերի "Պայմանագրի N" դաշտի լրացման կլաս
Sub Fill_InsuranceAgreeN(InsurAgreeN)
		Dim count
		' Լրացնել կոճակի սեղմում
  wMDIClient.VBObject("frmASDocForm").VBObject("AS_LABELCODESSBTN").Keys("[Enter]")
  If Not p1.WaitVBObject("frmAsUstPar",1000).Exists Then 
        wMDIClient.VBObject("frmASDocForm").VBObject("AS_LABELCODESSBTN").Keys("[Enter]")
    End If
		' Լրացնել Կապակցվող պայմանագիր պատուհանը
  Call Fill_RelatedAgreeWindow(InsurAgreeN, "Dialog")
		Call ClickCmdButton(2, "Î³ï³ñ»É")
  'Պայմանագրի ընտրում ցուցակից
  count = p1.vbObject("frmModalBrowser").vbObject("tdbgView").ApproxCount
  If not Search_Row(count, InsurAgreeN.searchedValue) Then
    Log.Error InsurAgreeN.searchedValue & " isn't exist", "", pmNormal, ErrorColor
    p1.VBObject("frmModalBrowser").Close
  End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''Mortgage_General'''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Ստացված գրավ՝ Արժեթղթերի Ընդհանուր բաժնի լրացման կլաս
Class Mortgage_General
		public agreeType
		public agreeN
		public insuranceAgreeN
		public client
		public clientName
		public includeInterests
		public ratio
		public comment
		public signingDate
		public allocationDate
		public office
		public department
		public accessType
		private sub Class_Initialize()
				agreeType = ""
				agreeN = ""
				Set insuranceAgreeN = New_InsuranceAgreeN()
				client = ""
				clientName = ""
				includeInterests = 0
				ratio = ""
				comment = ""
				signingDate = ""
				allocationDate = ""
				office = ""
				department = ""
				accessType = ""
		end sub
End Class

Function New_Mortgage_General()
		Set New_Mortgage_General = new Mortgage_General
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''Fill_Mortgage_General''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պրոցեդուրան լրացնում է Ստացված գրավ՝ Արժեթղթերի Ընդհանուր բաժինը
'General - Ստացված գրավ՝ Արժեթղթերի Ընդհանուր բաժնի լրացման կլաս
Sub Fill_Mortgage_General(General)
  Call GoTo_ChoosedTab(1)
  ' Պայմանագրի տիպ դաշտի լրացում
  Call Rekvizit_Fill("Document", 1, "General", "SECTYPE", General.agreeType)

		' Ապահ. պայմ. N Գրիդի Լրացում
		BuiltIn.Delay(1000)
		Call Fill_InsuranceAgreeN(General.insuranceAgreeN)

		' Պայմանագրի համար դաշտի լրացում
    General.agreeN = Get_Rekvizit_Value("Document",1,"General","CODE")
  ' Հաճախորդ դաշտի լրացում
  Call Rekvizit_Fill("Document", 1, "General", "CLICOD", General.client)
		' Անվանում դաշտի լրացում
  Call Rekvizit_Fill("Document", 1, "General", "NAME", "![End]" & "[Del]" & General.clientName)
		' Ընդգրկել տոկոսները դաշտի լրացում
  Call Rekvizit_Fill("Document", 1, "CheckBox", "WITHPER", General.includeInterests)
		' Հարաբերակցություն դաշտի լրացում
  Call Rekvizit_Fill("Document", 1, "General", "RATIO", General.ratio)
		' Մեկնաբանություն դաշտի լրացում
  Call Rekvizit_Fill("Document", 1, "General", "COMMENT", General.comment)
		' Կնքման ամսաթիվ դաշտի լրացում
  Call Rekvizit_Fill("Document", 1, "General", "DATE", "![End]" & "[Del]" & General.signingDate)
		' Հատկացման ամսաթիվ դաշտի լրացում
  Call Rekvizit_Fill("Document", 1, "General", "DATEGIVE", "![End]" & "[Del]" & General.allocationDate)
		' Գրասենյակ դաշտի լրացում
  Call Rekvizit_Fill("Document", 1, "General", "ACSBRANCH", General.office)
		' Բաժին դաշտի լրացում
  Call Rekvizit_Fill("Document", 1, "General", "ACSDEPART", General.department)
		' Հասան-ն տիպ դաշտի լրացում
  Call Rekvizit_Fill("Document", 1, "General", "ACSTYPE", General.accessType)
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''Mortgage_Subject'''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Ստացված գրավ՝ Արժեթղթերի Գրավի առարկա բաժնի լրացման կլաս
'securitiesRow - ·Éáµ³É ÷á÷áË³Ï³Ý
Class Mortgage_Subject
		public thingsGrid()
		public agreeBalance
		public mortgageThing
		public mortgageLocation
		public additionalInfo
		private sub Class_Initialize()
				ReDim thingsGrid(securitiesRow, 5) 
				agreeBalance = 0
				mortgageThing = ""
				mortgageLocation = ""
				additionalInfo = ""
		end sub
End Class

'securitiesRow - ·Éáµ³É ÷á÷áË³Ï³Ý
Function New_Mortgage_Subject(rowCount)
		securitiesRow = rowCount
		Set New_Mortgage_Subject = new Mortgage_Subject
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''Fill_Mortgage_Subject'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պրոցեդուրան լրացնում է Ստացված գրավ՝ Արժեթղթերի Գրավի առարկա բաժինը
'Subject - Ստացված գրավ՝ Արժեթղթերի Գրավի առարկա բաժնի լրացման կլաս
Sub Fill_Mortgage_Subject(Subject)
		Dim i, j, DocGrid
		Set DocGrid = wMDIClient.vbObject("frmASDocForm").vbObject("TabFrame_2").vbObject("DocGrid")
		Call GoTo_ChoosedTab(2)
		for i = 0 to securitiesRow
				for j = 0 to 4
				  with DocGrid
				    .Col = j
				    .Row = i
				    .Keys(Subject.thingsGrid(i, j) & "[Right]")
						end with
				next
				DocGrid.Keys("[Home][Up]")
				BuiltIn.Delay(1000)
  next
		' Պայմանագրի մնացորդ դաշտի լրացում
  Call Rekvizit_Fill("Document", 2, "CheckBox", "FILLSEC", Subject.agreeBalance)
		' Գրավի առարկա(կրճատ) դաշտի լրացում
  Call Rekvizit_Fill("Document", 2, "General", "SHRTNAME", Subject.mortgageThing)
		' Գրավի գտնվելու վայր դաշտի լրացում
  Call Rekvizit_Fill("Document", 2, "General", "PLACE", Subject.mortgageLocation)
		' Լրացուցիչ ինֆորմացիա դաշտի լրացում
  Call Rekvizit_Fill("Document", 2, "General", "OTHER", Subject.additionalInfo)
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''Mortgage_Additional'''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Ստացված գրավ՝ Արժեթղթերի Լրացուցիչ բաժնի լրացման կլաս
Class Mortgage_Additional
		public restrictAvilability
		public riskWeight
		public CRD
		public fulctCoefficient
		public note
		public note2
		public note3
		public subjectACRA
		public subjectNewLR
		public agreePaperN
		public closeDate
		private sub Class_Initialize()
				restrictAvilability = 0
				riskWeight = ""
				CRD = 0
				fulctCoefficient = ""
				note = ""
				note2 = ""
				note3 = ""
				subjectACRA = ""
				subjectNewLR = ""
				agreePaperN = ""
				closeDate = ""
		end sub
End Class

Function New_Mortgage_Additional()
		Set New_Mortgage_Additional = new Mortgage_Additional
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''Fill_Mortgage_Additional'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պրոցեդուրան լրացնում է Ստացված գրավ՝ Արժեթղթերի Լրացուցիչ բաժինը
'Additional - Ստացված գրավ՝ Արժեթղթերի Լրացուցիչ բաժնի լրացման կլաս
Sub Fill_Mortgage_Additional(Additional)
		Call GoTo_ChoosedTab(3)
		' Սահամանափակումների առկայություն դաշտի լրացում
  Call Rekvizit_Fill("Document", 3, "CheckBox", "LIMITEDRISK", Additional.restrictAvilability)
		' Ռիսկի կշիռ դաշտի լրացում
  Call Rekvizit_Fill("Document", 3, "General", "RISKDEGREE", Additional.riskWeight)
		' ՎՌԶՄ դաշտի լրացում
  Call Rekvizit_Fill("Document", 3, "CheckBox", "VRZM", Additional.CRD)
		' Տատանման գործակից դաշտի լրացում
  Call Rekvizit_Fill("Document", 3, "General", "VARIATION", Additional.fulctCoefficient)
		' Նշում դաշտի լրացում
  Call Rekvizit_Fill("Document", 3, "General", "NOTE", Additional.note)
		' Նշում 2 դաշտի լրացում
  Call Rekvizit_Fill("Document", 3, "General", "NOTE2", Additional.note2)
		' Նշում 3 դաշտի լրացում
  Call Rekvizit_Fill("Document", 3, "General", "NOTE3", Additional.note3)
		' Գրավի առարկա ACRA դաշտի լրացում
  Call Rekvizit_Fill("Document", 3, "General", "ACRANOTE", Additional.subjectACRA)
		' Գրավի առարկա (Նոր ՎՌ) դաշտի լրացում
  Call Rekvizit_Fill("Document", 3, "General", "MORTSUBJECT", Additional.subjectNewLR)
		' Պայմ. թղթային N դաշտի լրացում
  Call Rekvizit_Fill("Document", 3, "General", "PPRCODE", Additional.agreePaperN)
		' Փակման ամսաթիվ դաշտի լրացում
  Call Rekvizit_Fill("Document", 3, "General", "DATECLOSE", Additional.closeDate)
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''Mortgage_Additional'''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Ստացված գրավ՝ Արժեթղթերի լրացման կլաս
'securitiesRow - ·Éáµ³É ÷á÷áË³Ï³Ý
Class Mortgage_Securities
		public fISN
		public general 
		public mortgageSubject
		public additional
		private sub Class_Initialize()
				fISN = ""
				Set general = New_Mortgage_General()
				Set mortgageSubject = New_Mortgage_Subject(securitiesRow)
				Set additional = New_Mortgage_Additional()
		end sub
End Class

'securitiesRow - ·Éáµ³É ÷á÷áË³Ï³Ý
Function New_Mortgage_Securities(rowCount)
		securitiesRow = rowCount
		Set New_Mortgage_Securities = new Mortgage_Securities
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''Fill_Mortgage_Additional'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պրոցեդուրան լրացնում է Ստացված գրավ՝ Արժեթղթերը
'Securities - Ստացված գրավ՝ Արժեթղթերի լրացման կլաս
Sub Fill_Mortgage_Securities(Securities)
		' Պայմանագրի ISN - ի վերագրում փոփոխականի
  Securities.fISN = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
		'	Լրացնել Ընդհանուր բաժինը
		Call Fill_Mortgage_General(Securities.general)
		'	Լրացնել Գրավի առարկա բաժինը
		Call Fill_Mortgage_Subject(Securities.mortgageSubject)
		'	Լրացնել Լրացուցիչ բաժինը
		Call Fill_Mortgage_Additional(Securities.additional)
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''ChooseAgreeRow'''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Ընտրել Պայմանագրի տեսակը բացված պատուհանում
Sub ChooseAgreeRow(FolderName, agrType)
		Dim rowCount
		Call wTreeView.DblClickItem(FolderName & "Üáñ å³ÛÙ³Ý³·ñÇ ëï»ÕÍáõÙ")
  rowCount = p1.vbObject("frmModalBrowser").vbObject("tdbgView").ApproxCount
  if not Search_Row(rowCount, agrType) then
		  Log.Error agrType & " isn't found", "", pmNormal, ErrorColor 
  End If
  wMDIClient.Refresh
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''Create_Mortgage_Securities'''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պրոցեդուրան ստեղծում է Ստացված գրավ՝ Արժեթղթեր պայմանագիր
'FolderName - Պայմանագրի անունը
'Securities - Ստացված գրավ՝ Արժեթղթերի լրացման կլաս
'agrType - Պայմանագրի տեսակ
Sub Create_Mortgage_Securities(FolderName, Securities, agrType)
		' Գրավի ընտրում "Նոր պայմանագրի ստեղծում" ցուցակից
  Call ChooseAgreeRow(FolderName, agrType)
		if wMDIClient.waitVbObject("frmASDocForm",2000).Exists then
				' Լրացնել Գրավի պայմանագիր՝ Արժեթղթեր պատուհանը
				Call Fill_Mortgage_Securities(Securities)	
				' Սեղմել Կատարել Կոճակը
				Call ClickCmdButton(1, "Î³ï³ñ»É")	
				' Փակել բացված պտտելը
				if wMDIClient.waitVbObject("frmPttel",2000).Exists then
						BuiltIn.Delay(3000)
						wMDIClient.VBObject("frmPttel").Close
				else 
						Log.Error "Can't open frmPttel window.", "", pmNormal, ErrorColor
				end if
		else 
				Log.Error "Can't open frmASDocForm window.", "", pmNormal, ErrorColor
		end if
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''Addition''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Լրացում գործողության լրացման կլաս
'gridRows - ·Éáµ³É ÷á÷áË³Ï³Ý
Class Addition
		public isn
		public date
		public agreems()
		public agreeBalance
		public comment
		private sub Class_Initialize()
				isn = ""
				date = ""
				ReDim agreems(gridRows, 5) 
				agreeBalance = 0
				comment = ""
		end sub
End Class

'gridRows - ·Éáµ³É ÷á÷áË³Ï³Ý
Function New_Addition(row)
		gridRows = row
		Set New_Addition = new Addition
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''Fill_Addition''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Լրացում գործողության լրացման պրոցեդուրա
'add - Լրացում գործողության լրացման կլաս
Sub Fill_Addition(add)
		Dim i, j, DocGrid
		Call GoTo_ChoosedTab(1)
		add.isn = wMDIClient.VBObject("frmASDocForm").DocFormCommon.Doc.isn
		' Ամսաթիվ դաշտի լրացում
  Call Rekvizit_Fill("Document", 1, "General", "CODE", add.date)
		' Պայմանագրեր գրիդի լրացում
		Set DocGrid = wMDIClient.vbObject("frmASDocForm").vbObject("TabFrame").vbObject("DocGrid")
		for i = 0 to gridRows
				for j = 0 to 5
				  with DocGrid
				    .Col = j
				    .Row = i
				    .Keys(add.agreems(i, j) & "[Right]")
						end with
				next
				DocGrid.Keys("[Home][Up]")
  next
		' Պայմանագրի մնացորդ դաշտի լրացում
  Call Rekvizit_Fill("Document", 1, "CheckBox", "FILLSEC", add.agreeBalance)
		' Մեկնաբանություն դաշտի լրացում
  Call Rekvizit_Fill("Document", 1, "General", "COMMENT", add.comment)
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''Addition_Operation'''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Լրացում գործողությում
'add - Լրացում գործողության լրացման կլաս
Sub Addition_Operation(add)
		Call wMainForm.MainMenu.Click(c_AllActions)
		Call wMainForm.PopupMenu.Click(c_Addition)
		if wMDIClient.waitVbObject("frmASDocForm",2000).Exists then
				Call Fill_Addition(add)
				' Սեղմել Կատարել Կոճակը
				Call ClickCmdButton(1, "Î³ï³ñ»É")
		else 
				Log.Error "Can't open frmASDocForm window.", "", pmNormal, ErrorColor
		end if
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''Reletion_Operation''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պայմանագրի կապաակցում գործողություն
'relet - Պայմանագրի կապակցում պատուհանի լրացման կլաս
Function Reletion_Operation(relet)
		Dim count, chBox, frmAsUstPar
		Call wMainForm.MainMenu.Click(c_AllActions)
		Call wMainForm.PopupMenu.Click(c_AgrBind)
		if p1.WaitVBObject("frmAsUstPar", delay_small).Exists then
				Set frmAsUstPar = p1.WaitVBObject("frmAsUstPar", delay_small)
				' Լրացնել կոճակի սեղմում
				chBox = GetVBObject_Dialog("FILLCODESS", frmAsUstPar)
				frmAsUstPar.vbObject("TabFrame").vbObject(chBox).Click
		'  Call Rekvizit_Fill("Dialog", 1, "CheckBox", "FILLCODESS", 1)
				' Լրացնել Կապակցվող պայմանագիր պատուհանը
		  Call Fill_RelatedAgreeWindow(relet, "Dialog_2")
				Call ClickCmdButton("2_2", "Î³ï³ñ»É")
		  'Պայմանագրի ընտրում ցուցակից
		  count = p1.vbObject("frmModalBrowser").vbObject("tdbgView").ApproxCount
		  If not Search_Row(count, relet.searchedValue) Then
		    Log.Error relet.searchedValue & " isn't exist", "", pmNormal, ErrorColor
		    Sys.Process("Asbank").VBObject("frmModalBrowser").Close
		  End If
				Call ClickCmdButton(2, "Î³ï³ñ»É")
		else 
				Log.Error "Can't find frmAsUstPar window.", "", pmNormal, ErrorColor
		end if
		if wMDIClient.VBObject("frmASDocForm").Exists then
				Reletion_Operation = wMDIClient.VBObject("frmASDocForm").DocFormCommon.Doc.isn
				Call ClickCmdButton(1, "Î³ï³ñ»É")
		else 
				Log.Error "Can't find frmASDocForm window.", "", pmNormal, ErrorColor
		end if
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''LeaseAgreements_General'''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Վարձակալության պայմանագրի Ընդհանուր բաժնի լրացման կլաս
Class LeaseAgreements_General
		public agreeType
		public agreeN
		public landlord
		public name
		public curr
		public leaseSumma
		public signingDate
		public repayDate
		public AAHTaxable
		public marketInterestRate
		public comment
		public office
		public department
		public accessType
		private sub Class_Initialize()
				agreeType = "1"
				agreeN = ""
				landlord = ""
				name = ""
				curr = ""
				leaseSumma = ""
				signingDate = ""
				repayDate = ""
				AAHTaxable = ""
				marketInterestRate = ""
				comment = ""
				office = ""
				department = ""
				accessType = ""
		end sub
End Class

Function New_LeaseAgreements_General()
		Set New_LeaseAgreements_General = new LeaseAgreements_General
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''Fill_LeaseAgreements_General'''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Վարձակալության պայմանագրի Ընդհանուր բաժնի լրացման պրոցեդուրա
'general - Վարձակալության պայմանագրի Ընդհանուր բաժնի լրացման կլաս
Sub Fill_LeaseAgreements_General(general)
		Call GoTo_ChoosedTab(1)
		' Պայմանագրի տեսակ դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "LCTYPE", general.agreeType)
		' Պայմանագրի N դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "CODE", general.agreeN)
		' Վարձատու դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "CLICOD", general.landlord)
		' Անվանում դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "NAME", general.name)
		' Արժույթ դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "CURRENCY", general.curr)
		' Վարձակալության գումար դշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "SUMMA", general.leaseSumma)
		' Կնքման ամսաթիվ դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "DATE", "![End]" & "[Del]" & general.signingDate)
		' Մարման ժամկետ դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "DATEAGR", general.repayDate)
		' ԱԱՀ-ով հարկվող դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "VATMETH", general.AAHTaxable)
		' Շուկայական տոկոսադրույք դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "PCMARKET", general.marketInterestRate)
		' Մեկնաբանություն դաշտի լրացում 
		Call Rekvizit_Fill("Document", 1, "General", "COMMENT", general.comment)
		' Գրասենյակ դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "ACSBRANCH", general.office)
		' Բաժին դաշտի լարցում 
		Call Rekvizit_Fill("Document", 1, "General", "ACSDEPART", general.department)
		' Հասան-ն տիպ դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "ACSTYPE", general.accessType)
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''LeaseAgreements_Additional'''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Վարձակալության պայմանագրի Լրացուցիչ բաժնի լրացման կլաս
Class LeaseAgreements_Additional
		public note 
		public note2
		public note3
		public agreePaperN
		public closeDate
		private sub Class_Initialize()
				note = ""
				note2 = ""
				note3 = ""
				agreePaperN = ""
				closeDate = ""
		end sub
End Class

Function New_LeaseAgreements_Additional()
		Set New_LeaseAgreements_Additional = new LeaseAgreements_Additional
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''Fill_LeaseAgreements_Additional''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Վարձակալության պայմանագրի Լրացուցիչ բաժնի լրացման պրոցեդուրա
'additional - Վարձակալության պայմանագրի Լրացուցիչ բաժնի լրացման կլաս
Sub Fill_LeaseAgreements_Additional(additional)
		Call GoTo_ChoosedTab(2)
		' Նշում դաշտի լրացում
		Call Rekvizit_Fill("Document", 2, "General", "NOTE", additional.note)
		' Նշում 2 դաշտի լրացում
		Call Rekvizit_Fill("Document", 2, "General", "NOTE2", additional.note2)
		' Նշում 3 դաշտի լրացում
		Call Rekvizit_Fill("Document", 2, "General", "NOTE3", additional.note3)
		' Պայմ. թղթային N դաշտի լրացում
		Call Rekvizit_Fill("Document", 2, "General", "PPRCODE", additional.agreePaperN)
		' Փակման ամսաթիվ դաշտի լրացում
		Call Rekvizit_Fill("Document", 2, "General", "DATECLOSE", additional.closeDate)
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''LeaseAgreements''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Վարձակալության պայմանագրի լրացման կլաս
Class LeaseAgreements
		public general 
		public additional
		public isn
		private sub Class_Initialize()
				Set general = New_LeaseAgreements_General()
				Set additional = New_LeaseAgreements_Additional()
				isn = ""
		end sub
End Class

Function New_LeaseAgreements()
		Set New_LeaseAgreements = new LeaseAgreements
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''Fill_LeaseAgreements'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Վարձակալության պայմանագրի լրացման պրոցեդուրա
'LeaseAgree - Վարձակալության պայմանագրի լրացման կլաս
Sub Fill_LeaseAgreements(LeaseAgree)
		' Ընդհանում բաժնի լրացում
		Call Fill_LeaseAgreements_General(LeaseAgree.general)
		' Լրացուցիչ դաշտի լրացում
		Call Fill_LeaseAgreements_Additional(LeaseAgree.additional)
		'Վերցնել "Պայմանագրի ISN-ը" դաշտի արժեքը
		LeaseAgree.isn = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''Create_LeaseAgreement'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Վարձակալության պայմանագրի ստեղծման պրոցեդուրա
'LeaseAgree - Վարձակալության պայմանագրի լրացման կլաս
Sub Create_LeaseAgreement(FolderName, LeaseAgree)
		wTreeView.DblClickItem(folderName & "Üáñ å³ÛÙ³Ý³·ñÇ ëï»ÕÍáõÙ")
		if wMDIClient.WaitVBObject("frmASDocForm", 3000).Exists then
				Call Fill_LeaseAgreements(LeaseAgree)
				Call ClickCmdButton(1, "Î³ï³ñ»É")
		else 
				Log.Error "Can't open Lease Agreement(frmASDocForm) widow.", "", pmNormal, ErrorColor
		end if
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''MainAccWorkingDocuments'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Գլխավոր հաշվապահի ԱՇՏ-ում Աշխատանքային փաստաթղթեր պատուհանի լրացման կլաս
Class MainAccWorkingDocuments
		Public startDate
		Public endDate
		Public curr
		Public performer
		Public docType
		Public inPaySys
		Public outPaySys
		Public note
		Public office
		Public section
		Public viewType
		Public fill
		Private Sub Class_Initialize()
				startDate = ""
				endDate = ""
				curr = ""
				performer = ""
				docType = ""
				inPaySys = ""
				outPaySys = ""
				note = ""
				office = ""
				section = ""
				viewType = "Oper"
				fill = "0"
		End Sub
End Class

Function New_MainAccWorkingDocuments()
		Set New_MainAccWorkingDocuments = New MainAccWorkingDocuments
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''Fill_MainAccWorkingDocuments'''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Գլխավոր հաշվապահի ԱՇՏ-ում Աշխատանքային փաստաթղթեր պատուհանի լրացման պրոցեդուրա
'WorkingDocs - պատուհանի լրացման կլաս
Sub Fill_MainAccWorkingDocuments(workingDocs)
  ' Ժամանակահատված սկզբնական դաշտի լրացում 
		Call Rekvizit_Fill("Dialog", 1, "General", "PERN", "![End]" & "[Del]" & workingDocs.startDate)
		' Ժամանակահատված վերջնական դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "PERK", "![End]" & "[Del]" & workingDocs.endDate)
		' Արժույթ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "CUR", workingDocs.curr)
		' Կատարող դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "USER", "![End]" & "[Del]" & workingDocs.performer)
		' Փաստաթղթի տեսակ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "DOCTYPE", workingDocs.docType)
		' Ընդ. վճ. համակարգ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "PAYSYSIN", workingDocs.inPaySys)
		' Ուղ. վճ. համակարգ դաշտի լարցում
		Call Rekvizit_Fill("Dialog", 1, "General", "PAYSYSOUT", workingDocs.outPaySys)
		' Նշում դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "PAYNOTES", workingDocs.note)
		' Գրասենյակ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "ACSBRANCH", workingDocs.office)
		' Բաժին դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "ACSDEPART", workingDocs.section)
		' Դիտելու ձև դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "SELECTED_VIEW", workingDocs.viewType)
		' Լրացնել դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "EXPORT_EXCEL", workingDocs.fill)
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''GoTo_MainAccWorkingDocuments'''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Գլխավոր հաշվապահի ԱՇՏ-ում Աշխատանքային փաստաթղթեր թղթապանակ մուտք գործելու պրոցեդուրա
'folderName - գտնբելու ճանապարհը
'WorkingDocs - պատուհանի լրացման կլաս
Sub GoTo_MainAccWorkingDocuments(folderName, workingDocs)
		wTreeView.DblClickItem(folderName & "²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ")
		if p1.WaitVBObject("frmAsUstPar", 3000).Exists then
				Call Fill_MainAccWorkingDocuments(workingDocs)
				Call ClickCmdButton(2, "Î³ï³ñ»É")
		else 
				Log.Error "Can't open frmAsUstPar widow.", "", pmNormal, ErrorColor
		end if
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''GoTo_Tasks'''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պայմանագրեր թղթապանակ մուտք գործելու պրոցեդուրա
'firstDate - Սկզբի ամսաթիվ
'lastDate - Վերջին ամսաթիվ
Sub GoTo_Tasks(firstDate, lastDate)
    wTreeView.DblClickItem("²é³ç³¹ñ³ÝùÝ»ñ|²é³ç³¹ñ³ÝùÝ»ñ|")
    if p1.WaitVBObject("frmAsUstPar", 3000).Exists then
        Call Rekvizit_Fill("Dialog", 1, "General", "SDATE", firstDate)
        Call Rekvizit_Fill("Dialog", 1, "General", "EDATE", lastDate)
        Call ClickCmdButton(2, "Î³ï³ñ»É")
    else 
        Log.Error "Can't open frmAsUstPar widow.", "", pmNormal, ErrorColor
    end if
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''Conjuction''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Կապակցված հաճախորդներ գրիդի լրացման կլաս
Class Conjuction
    Public tabN
    Public fIsn
    Public client()
    Public name()
    Public conjuctType()
    Public conjuctName()
    Public comment()
    Public clientsCount
    Private Sub Class_Initialize()
        tabN = 1
        fIsn = ""
        clientsCount = clients_count
        Redim client(clientsCount)
        Redim name(clientsCount)
        Redim conjuctType(clientsCount)
        Redim conjuctName(clientsCount)
        Redim comment(clientsCount)
    End Sub
End Class

Function New_Conjuction(cliCount)
    clients_count = cliCount
    Set New_Conjuction = New Conjuction 
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''Fill_ConjuctionGrid'''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Պրոցեդուրան լրացնում է Կապակցել հաճախորդներ պատուհանի գրիդային մասը
' cliConjuction - Կապակցված հաճախորդներ գրիդի լրացման կլաս
Sub Fill_ConjuctionGrid(cliConjuction)
    Dim i, DocGrid
    
    Set DocGrid = wMDIClient.vbObject("frmASDocForm").vbObject("TabFrame").vbObject("DocGrid")
    for i = 0 to cliConjuction.clientsCount - 1
        ' Լրացնել Հաճախորդ դաշտը
        Call Fill_Grid_Field(0, i, "Document", "General", cliConjuction.tabN, cliConjuction.client(i))
        ' Ստուգել Անվանում դաշտի արժեքը
        Call Check_Value_Grid (1, i, "Document", cliConjuction.tabN, cliConjuction.name(i))
        ' Լրացնել Կապի տեսակ դաշտը
        Call Fill_Grid_Field(2, i, "Document", "General", cliConjuction.tabN, cliConjuction.conjuctType(i))
        ' Ստուգել Կապի անվանում դաշտի արժեքը
        Call Check_Value_Grid (3, i, "Document", cliConjuction.tabN, cliConjuction.conjuctName(i))
        ' Լրացնել Մեկնաբանություն դաշտը
        Call Fill_Grid_Field(4, i, "Document", "General", cliConjuction.tabN, cliConjuction.comment(i))
        DocGrid.Keys("[Home][Up]")
        BuiltIn.Delay(1000)
    next
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''Cilent_Conjunction'''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Պրոցեդուրան կատարում է Կապակցել հաճախորդներ գործողություն
' cliConjuction - Կապակցված հաճախորդներ գրիդի լրացման կլաս
Sub Cilent_Conjunction(cliConjuction)
    BuiltIn.Delay(3000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_ConjuctPersons & "|" & c_ConjuctClients)
    If wMDIClient.WaitvbObject("frmASDocForm", 3000).Exists Then
        cliConjuction.fIsn = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
        Call Fill_ConjuctionGrid(cliConjuction)
        Call ClickCmdButton(1, "Î³ï³ñ»É")
    Else 
        Log.Error "Can't find frmASDocForm window.", "", pmNormal, ErrorColor    
    End If
End Sub