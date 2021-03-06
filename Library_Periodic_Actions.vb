'USEUNIT Library_Common
'USEUNIT Library_Colour
'USEUNIT Library_Contracts 
'USEUNIT Payment_Except_Library
'USEUNIT SWIFT_International_Payorder_Library
'USEUNIT Constants

Dim rowCount

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''Periodic_General_OpersGrid'''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պարբերական գործողությունների պայմանագրի Ընդհանուր բաժնի գրիդի լրացման կլաս
Class Periodic_General_OpersGrid
		public N_Edit
		public operType
		public calcMethod
		public debitAccount
		public debitCurr
		public depositAccount
		public depositCurr
		public percent
		public price
		public curr
		public transactionRate
		public rateChange
		public daysCount
		public minPrice
		public maxPrice
		public aim
		public addDocument
		public opersAddDoc
		public receivingBank
		public recipient
		public recipLegalStatus
		public docN
		public secID
		public debitAccountName
		public depositAccountName
		private sub Class_Initialze()
				N_Edit = ""
				operType = ""
				calcMethod = ""
				debitAccount = ""
				debitCurr = ""
				depositAccount  = ""
				depositCurr = ""
				percent = ""
				price = ""
				curr = ""
				transactionRate = ""
				rateChange = ""
				daysCount = ""
				minPrice = ""
				maxPrice = "" 
				aim = ""
				addDocument = ""
				opersAddDoc = 0
				receivingBank = ""
				recipient = ""
				recipLegalStatus = ""
				docN = ""
				secID = ""
				debitAccountName = ""
				depositAccountName = ""
		end sub
End Class

Function New_Periodic_Gen_OpersGrid()
		Set New_Periodic_Gen_OpersGrid = new Periodic_General_OpersGrid
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''Fill_General_OpersGrid'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պարբերական գործողությունների պայմանագրի Ընդհանուր բաժնի գրիդի լրացման պրոցեդուրա
'gridOpers - Պարբերական գործողությունների պայմանագրի Ընդհանուր բաժնի գրիդի լրացման կլաս
Sub Fill_General_OpersGrid(gridOpers, i)
		Dim DocGridDocGrid
		Set DocGrid = wMDIClient.vbObject("frmASDocForm").vbObject("TabFrame").vbObject("DocGrid")
  with DocGrid
				if gridOpers.operType = "01" or gridOpers.operType = "02" or gridOpers.operType = "06" then
						Call Grid_Fill_Inline(DocGrid, gridOpers, i)
				else 
						Call Grid_Fill_AttachedDoc(DocGrid, gridOpers, i)
				end if
  End With  
End Sub

Sub Grid_Fill_Inline(docGrid, gridOpers, i)
		with docGrid
				.Row = i
				' N/Խմբագրել դաշտի լրացում
				.Col = 0
    .Keys(gridOpers.N_Edit & "[Right]")
				' Գործ. տեսակ դաշտի լրացում
				.Col = 1
    .Keys(gridOpers.operType & "[Right]")
				' Հաշվարկման եղանակ դաշտի լրացում
				.Col = 2
    .Keys(gridOpers.calcMethod & "[Right]")
				' Հաշիվ դեբետ դաշտի լրացում
				.Col = 3
    .Keys(gridOpers.debitAccount & "[Right]")
				' Արժ. դբ. դաշտի լրացում
				.Col = 4
    .Keys(gridOpers.debitCurr & "[Right]")
				' Հաշիվ կրեդիտ/Ավանդային դաշտի լրացում
				.Col = 5
    .Keys(gridOpers.depositAccount & "[Right]")
				' Արժ. կր./ավանդ դաշտի լրացում
				.Col = 6
    .Keys(gridOpers.depositCurr & "[Right]")
				' Տոկոս դաշտի լրացում
				.Col = 7
    .Keys(gridOpers.percent & "[Right]")
				' Գումար/Բանաձև դաշտի լրացում
				.Col = 8
    .Keys(gridOpers.price & "[Right]")
				' Արժ. դաշտի լրացում
				.Col = 9
    .Keys(gridOpers.curr & "[Right]")
				' Գարծարքի փոխարժեք դաշտի լրացում
				.Col = 10
    .Keys(gridOpers.transactionRate & "[Right]")
				' Փոխարժեքի շեղում դաշտի լրացում
				.Col = 11
    .Keys(gridOpers.rateChange & "[Right]")
				' Օրերի քանակ դաշտի լրացում
				.Col = 12
    .Keys(gridOpers.daysCount & "[Right]")
				' Նվազագույն գումար դաշտի լրացում
				.Col = 13
    .Keys(gridOpers.minPrice & "[Right]")
				' Առավելագույն գումար դաշտի լրացում
				.Col = 14
    .Keys(gridOpers.maxPrice & "[Right]")
				' Նպատակ դաշտի լրացում 
				.Col = 15
    .Keys(gridOpers.aim & "[Right]")
				' Կից փաստաթուղթ դաշտի լրացում 
				.Col = 16
    .Keys(gridOpers.addDocument & "[Right]")
'				' Ստեղծել/խմբագրել/դիտել կից փաստաթուղթը դաշտի լրացում
'				.Col = 17
'    .Keys(gridOpers.opersAddDoc & "[Right]")
		end with
		DocGrid.Keys("[Home][Up]")
  BuiltIn.Delay(1000)
End Sub

Sub Grid_Fill_AttachedDoc(docGrid, gridOpers, i)
		with docGrid
				.Row = i
				' N/Խմբագրել դաշտի լրացում
				.Col = 0
		  .Keys(gridOpers.N_Edit & "[Right]")
				' Գործ. տեսակ դաշտի լրացում
				.Col = 1
		  .Keys(gridOpers.operType & "[Right]")
				.Keys("[End]")
				' Ստեղծել/խմբագրել/դիտել կից փաստաթուղթը դաշտի լրացում
				.Col = 17
				if gridOpers.opersAddDoc then
		    .Keys("y")
				end if
				if p1.WaitVBObject("frmASDocFormModal", 2000).Exists then
						select case gridOpers.operType
								case "03", "04", "08"
										Call PaymentOrder_Oper(gridOpers)
								case "05"
										Call InternationalPaymentOrder_Oper(gridOpers)
								case "07" 
										Call SecuritiesFreeShippingOrder_Oper(gridOpers)
								case "09"
								Call CountlessTransfers_Oper(gridOpers)
						end select
						
						' Սեղմել Կատարել կոճակը
						Call ClickCmdButton(4, "Î³ï³ñ»É")
				else 
				Log.Error "Can't open frmASDocFormModal widow.", "", pmNormal, ErrorColor
				end if
				' Հաշվարկման եղանակ դաշտի լրացում
				.Keys("[Home]")
				.Col = 2
    .Keys(gridOpers.calcMethod & "[Right]")
				' Գումար/Բանաձև դաշտի լրացում
				.Col = 8
    .Keys(gridOpers.price & "[Right]")
				' Արժ. դաշտի լրացում
				.Col = 9
    .Keys(gridOpers.curr & "[Right]")
				' Օրերի քանակ դաշտի լրացում
				.Col = 12
    .Keys(gridOpers.daysCount & "[Right]")
			end with
End Sub

Sub PaymentOrder_Oper(gridOpers)
		if p1.WaitVBObject("frmASDocFormModal", 1000).Exists then
				' Հաշիվ դեբետ դաշտի լրացում
				Call Rekvizit_Fill("DocumentModal", 1, "General", "ACCDB", gridOpers.debitAccount)
				' Հաշիվ կրեդիտ դաշտի լրացում
				Call Rekvizit_Fill("DocumentModal", 1, "General", "ACCCR", gridOpers.depositAccount)
				if gridOpers.operType = "08" then 
						' Ստացողի Իրավաբանական կարգավիճակ դաշտի լրացում
						Call Rekvizit_Fill("DocumentModal", 1, "General", "JURSTATR", gridOpers.recipLegalStatus)
				end if
				' Նպատակ դաշտի լրացում
				Call Rekvizit_Fill("DocumentModal", 1, "General", "AIM", gridOpers.aim)
		else
				Log.Error "Can't open frmASDocFormModal window.", "", pmNormal, ErrorColor
		end if
End	Sub

Sub InternationalPaymentOrder_Oper(gridOpers)
		if p1.WaitVBObject("frmASDocFormModal", 1000).Exists then
				' Վճարողի հաշիվ դաշտի լրացում
				Call Rekvizit_Fill("DocumentModal", 1, "General", "ACCDB", gridOpers.debitAccount)
				' Ստացող դաշտի լրացում
				Call Rekvizit_Fill("DocumentModal", 1, "General", "RECEIVER", gridOpers.recipient)
				' Նպատակ դաշտի լրացում
				Call Rekvizit_Fill("DocumentModal", 2, "General", "AIM", gridOpers.aim)	
		else
				Log.Error "Can't open frmASDocFormModal window.", "", pmNormal, ErrorColor
		end if
End	Sub

Sub SecuritiesFreeShippingOrder_Oper(gridOpers)
		if p1.WaitVBObject("frmASDocFormModal", 1000).Exists then
		' Փաստաթղթի համար դաշտի լրացում		
		Call Rekvizit_Fill("DocumentModal", 1, "General", "BMDOCNUM", gridOpers.docN)
		' Արժեթղթերի իդենտիֆիկատոր դաշտի լրացում
		Call Rekvizit_Fill("DocumentModal", 1, "General", "STOCKID", gridOpers.secID)
		' Արժեթղթեր առաքողի հաշիվ դաշտի լրացում
		Call Rekvizit_Fill("DocumentModal", 1, "General", "SSENDER", gridOpers.debitAccount)
		' Արժեթղթեր առաքող դաշտի լրացում
		Call Rekvizit_Fill("DocumentModal", 1, "General", "SSNAME", gridOpers.debitAccountName)
		' Արժեթղթեր ստացող հաշիվ դաշտի լրացում
		Call Rekvizit_Fill("DocumentModal", 1, "General", "SRECEIVER", gridOpers.depositAccount)
		' Արժեթղթեր ստացող դաշտի լրացում
		Call Rekvizit_Fill("DocumentModal", 1, "General", "SRNAME", gridOpers.depositAccountName)
		' Լրացուցիչ ինֆորմացիա դաշտի լրացում
		Call Rekvizit_Fill("DocumentModal", 1, "General", "ADDINFO", gridOpers.aim)
		else
				Log.Error "Can't open frmASDocFormModal window.", "", pmNormal, ErrorColor
		end if
End	Sub

Sub CountlessTransfers_Oper(gridOpers)
		if p1.WaitVBObject("frmASDocFormModal", 1000).Exists then
				' Հաշիվ դեբետ դաշտի լրացում
				Call Rekvizit_Fill("DocumentModal", 1, "General", "ACCDB", gridOpers.debitAccount)
		'		' Վճարող դաշտի լրացում
		'		Call Rekvizit_Fill("DocumentModal", 1, "General", "PAYER", gridOpers.debitAccount)
		'		' Վճարողի Իրավաբանական կարգավիճակ դաշտի լրացում
		'		Call Rekvizit_Fill("DocumentModal", 1, "General", "JURSTAT", gridOpers.debitAccount)	
		'		' Վճարող (անգլ.) դաշտի լրացում
		'		Call Rekvizit_Fill("DocumentModal", 1, "General", "EPAYER", gridOpers.debitAccount)	
				' Ստացող բանկ դաշտի լրացում
				Call Rekvizit_Fill("DocumentModal", 1, "General", "BANKCR", gridOpers.receivingBank)	
				' Ստացող դաշտի լրացում
				Call Rekvizit_Fill("DocumentModal", 1, "General", "RECEIVER", gridOpers.recipient)	
				' Ստացողի Իրավաբանական կարգավիճակ դաշտի լրացում
				Call Rekvizit_Fill("DocumentModal", 1, "General", "JURSTATR", gridOpers.recipLegalStatus)	
				' Նպատակ դաշտի լրացում
				Call Rekvizit_Fill("DocumentModal", 1, "General", "AIM", gridOpers.aim)	
		else
				Log.Error "Can't open frmASDocFormModal window.", "", pmNormal, ErrorColor
		end if
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''PeriodicActions_General'''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պարբերական գործողությունների պայմանագրի Ընդհանուր բաժնի լրացման կլաս
Class PeriodicActions_General
		public office
		public department
		public performer
		public agreeN
		public client
		public name 
		public englName
		public startDate
		public endDate
		public doInEveryCall
		public periodMounth
		public periodDay
		public mounthDays
		public implementDays_start
		public implementDays_end
		public bypassNonWorkDays
		public overlimit
		public opersGridRowCount
		public operations()
		private sub Class_Initialize()
				office = ""
				department = ""
				performer = ""
				agreeN = ""
				client = ""
				name = ""
				englName = ""
				startDate = aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%d%m%y")
				endDate = ""
				doInEveryCall = 0
				periodMounth = ""
				periodDay = ""
				mounthDays = ""
				implementDays_start = ""
				implementDays_end = ""
				bypassNonWorkDays = ""
				overlimit = 0
				ReDim operations(rowCount)
				for opersGridRowCount = 0 to rowCount - 1
						Set operations(opersGridRowCount) = New_Periodic_Gen_OpersGrid()
				next
				opersGridRowCount = rowCount
		end sub
End Class

Function New_PeriodicActions_General(row_Count)
		rowCount = row_Count
		Set New_PeriodicActions_General = new PeriodicActions_General
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''Fill_Periodic_General''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պարբերական գործողությունների պայմանագրի Ընդհանուր բաժնի լրացման պրոցեդուրա
'general - Պարբերական գործողությունների պայմանագրի Ընդհանուր բաժնի լրացման կլաս
Sub Fill_Periodic_General(general)
		Dim i
'		Call GoTo_ChoosedTab(1)
		' Գրասենյակ դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "ACSBRANCH", general.office)
		' Բաժին դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "ACSDEPART", general.department)
		' Կարտարող դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "USERID", general.performer)
		' Պայմանագրի N դաշտի ստացում
		general.agreeN = Get_Rekvizit_Value("Document", 1, "General", "CODE")
		' Հաճախորդ դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "CLICODE", general.client)
		' Անվանում դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "NAME", general.name)
		' Անգլերեն անվանում դաշտի լարցում
		Call Rekvizit_Fill("Document", 1, "General", "ENAME", general.englName)
		' Սկզբի ամսաթիվ դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "SDATE", general.startDate)
		' Վերջի ամսաթիվ դաշտի լարցում 
		Call Rekvizit_Fill("Document", 1, "General", "EDATE", general.endDate)
		' Կատարել ամեն կանչի ժամական դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "CheckBox", "CALCALWAYS", general.doInEveryCall)
		' Պարբերություն դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "PERIODICITY", general.periodMounth & "[Tab]" & general.periodDay)
  if not general.mounthDays = "" then 
				' Ամսվա օրեր դաշտի լրացում
				Call Rekvizit_Fill("Document", 1, "General", "DAYSOFMONTH", general.mounthDays)
		end if
		' Կատարման օրեր սկզբի դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "SDAY", general.implementDays_start)
		' Կատարման օրեր վերջի դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "LDAY", general.implementDays_end)
		' Ոչ աշխատանքային օրերի շրջանցում դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "NONWORKDAYS", general.bypassNonWorkDays)
		' Անբավարար միջոցների դեպքում բացել գերծախս
		Call Rekvizit_Fill("Document", 1, "CheckBox", "USEOVERLIMIT", general.overlimit)
		' Գործողություններ գրիդի լարցում
		for i = 0 to general.opersGridRowCount - 1
				Call Fill_General_OpersGrid(general.operations(i), i)
		next
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''PeriodicActions_Other'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պարբերական գործողությունների պայմանագրի Լրացուցիչ բաժնի լրացման կլաս
Class PeriodicActions_Other
		public informToClient
		public useClientEmail
		public clientEmail
		public otherEmail
		public note
		public note2
		public note3
		public addInfo
		public lastDate
		public closeDate
		private sub Class_Initialize()
				informToClient = 0
				useClientEmail = 0
				clientEmail = false 
				otherEmail = ""
				note = ""
				note2 = ""
				note3 = ""
				addInfo = ""
				lastDate = ""
				closeDate = ""
		end sub
End Class

Function New_PeriodicActions_Other()
		Set New_PeriodicActions_Other = new PeriodicActions_Other
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''Fill_Periodic_Additional'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պարբերական գործողությունների պայմանագրի Լրացուցիչ բաժնի լրացման պրոցեդուրա
'additional - Պարբերական գործողությունների պայմանագրի Լրացուցիչ բաժնի լրացման կլաս
Sub Fill_Periodic_Other(other)
		Call GoTo_ChoosedTab(2)
		' Տեղեկացնել հաճախորդին դաշտի լարցում
		Call Rekvizit_Fill("Document", 2, "CheckBox", "CLINOT", other.informToClient)
		if other.informToClient = 1 then
				if other.clientEmail then
						' Օգտ. հաճախորդի էլ. հասցեն դաշտի լարցում
						Call Rekvizit_Fill("Document", 2, "CheckBox", "USECLIEMAIL", other.useClientEmail)
				else 
				  ' Այլ էլ. հասցե դաշտի լրացում
						Call Rekvizit_Fill("Document", 2, "General", "EMAIL", other.otherEmail)
				end if
		end if
		' Նշում դաշտի լարցում
		Call Rekvizit_Fill("Document", 2, "General", "NOTE1", other.note)
		' Նշում 2 դաշտի լրացում
		Call Rekvizit_Fill("Document", 2, "General", "NOTE2", other.note2)
		' Նշում 3 դաշտի լրացում
		Call Rekvizit_Fill("Document", 2, "General", "NOTE3", other.note3)
		' Լրացուցիչ ինֆորմացի դաշտի լրացում
		Call Rekvizit_Fill("Document", 2, "General", "COMM", other.addInfo)
		' Վերջ. գործ-ն ամսաթիվ դաշտի լրացում
		Call Rekvizit_Fill("Document", 2, "General", "LASTOPDATE", other.lastDate)
		' Փակման ամսաթիվ դաշտի լրացում 
		Call Rekvizit_Fill("Document", 2, "General", "DATECLOSE", other.closeDate)
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''PeriodicActions''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պարբերական գործողությունների պայմանագրի լրացման կլաս
Class PeriodicActions
		public fisn
		public general 
		public other
		private sub Class_Initialize()
				fisn = ""
				Set general = New_PeriodicActions_General(rowCount)
				Set other = New_PeriodicActions_Other()
		end sub
End Class 

Function New_PeriodicActions(row_Count)
		rowCount = row_Count
		Set New_PeriodicActions = new PeriodicActions
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''Fill_PeriodicActions'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պարբերական գործողությունների պայմանագրի լրացման պրոցեդուրա
'periodActions - Պարբերական գործողությունների պայմանագրի լրացման կլաս
Sub Fill_PeriodicActions(periodActions)
		' Վերցնել "Պայմանագրի ISN-ը" դաշտի արժեքը
		periodActions.fisn = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
		' Ընդհանուր բաժնի լացում
		Call Fill_Periodic_General(periodActions.general)
		' Այլ բաժնի լրացում
		Call Fill_Periodic_Other(periodActions.other)
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''Create_PeriodicActions'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պարբերական գործողությունների պայմանագրի ստեղծման պրոցեդուրա
'periodActions - Պարբերական գործողությունների պայմանագրի լրացման կլաս
'folderName - պայմանագիր ճանապարհը
Sub Create_PeriodicActions(folderName, periodActions, state)
		Select Case state
				Case "create"
						wTreeView.DblClickItem(folderName & "Üáñ å³ñµ»ñ. å³ÛÙ³Ý³·Çñ")
				Case "add"
						Call wMainForm.MainMenu.Click(c_Opers)
						Call wMainForm.PopupMenu.Click(c_Add)
		End Select
		if wMDIClient.WaitVBObject("frmASDocForm", 3000).Exists then
				Call Fill_PeriodicActions(periodActions)
				Call ClickCmdButton(1, "Î³ï³ñ»É")
		else 
				Log.Error "Can't open Periodic Actions(frmASDocForm) widow.", "", pmNormal, ErrorColor
		end if
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''PerCommPay_General_ServicesGrid'''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պարբերական կոմունալ վճարումների պայմանագրի Ընդհանուր բաժնի գրիդի լրացման կլաս
Class PerCommPay_General_ServicesGrid
		public Num
		public service
		public place
		public clientN
		public minPrice
		public maxPrice
		public client
		public address
		public legalPerson
		private sub Class_Initialze()
				Num = ""
				service = ""
				place = ""
				clientN = ""
				minPrice = ""
				maxPrice = "" 
				client = ""
				address = ""
				legalPerson = 0
		end sub
End Class

Function New_PerCommPay_Gen_ServicesGrid()
		Set New_PerCommPay_Gen_ServicesGrid = new PerCommPay_General_ServicesGrid
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''Fill_General_ServicesGrid''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պարբերական կոմունալ վճարումների պայմանագրի Ընդհանուր բաժնի գրիդի լրացման պրոցեդուրա
'gridService - Պարբերական կոմունալ վճարումների պայմանագրի Ընդհանուր բաժնի գրիդի լրացման կլաս
Sub Fill_General_ServicesGrid(gridService, i)
		Dim DocGrid
		Set DocGrid = wMDIClient.vbObject("frmASDocForm").vbObject("TabFrame").vbObject("DocGrid")
  with DocGrid
				.Row = i
    ' N դաշտի լարցում
    .Col = 0
    .Keys(gridService.Num & "[Right]")
		  ' Ծառայություն դաշտի լրացում
				.Col = 1
    .Keys(gridService.service & "[Right]")
				' Վայր դաշտի լարցում 
				.Col = 2
    .Keys(gridService.place & "[Right]")
				' Բաժանորդի համար դաշտի լարցում
				.Col = 3
    .Keys(gridService.clientN & "[Right]")
				' Նվազագույն գումար դաշտի լրացում 
				.Col = 4
    .Keys(gridService.minPrice & "[Right]")
				' Առավելագույն գումար դաշտի լրացում
				.Col = 5
    .Keys(gridService.maxPrice & "[Right]")
				' Բաժանորդ դաշտի լրացում
				.Col = 6
    .Keys(gridService.client & "[Right]")
				' Հասցե դաշտի լրացում
				.Col = 7
    .Keys(gridService.address & "[Right]")
				' Իրավ. անձ դաշտի լրացում
				.Col = 8
    .Keys(gridService.legalPerson & "[Right]")
  end with
  DocGrid.Keys("[Home][Up]")
  BuiltIn.Delay(1000)
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''CommunalPayment_General'''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պարբերական կոմունալ վճարումների պայմանագրի Ընդհանուր բաժնի լրացման կլաս
Class CommunalPayment_General
		public office
		public department
		public performer
		public client
		public name 
		public englName
		public account
		public curr
		public maxPrice
		public servGridRowCount
		public services()
		private sub Class_Initialize()
				office = ""
				department = ""
				performer = ""
				client = ""
				name = ""
				englName = ""
				account = ""
				curr = ""
				maxPrice = ""
				ReDim services(rowCount)
				for servGridRowCount = 0 to rowCount - 1
						Set services(servGridRowCount) = New_PerCommPay_Gen_ServicesGrid()
				next
				servGridRowCount = rowCount
		end sub
End Class

Function New_CommunalPayment_General()
'		rowCount = row_Count
		Set New_CommunalPayment_General = new CommunalPayment_General
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''Fill_CommunalPayment_General''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պարբերական կոմունալ վճարումների պայմանագրի Ընդհանուր բաժնի լրացման պրոցեդուրա
'general - Պարբերական կոմունալ վճարումների պայմանագրի Ընդհանուր բաժնի լրացման կլաս
Sub Fill_CommunalPayment_General(general)
		Dim i
		Call GoTo_ChoosedTab(1)
		' Գրասենյակ դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "ACSBRANCH", general.office)
		' Բաժին դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "ACSDEPART", general.department)
		' Կատարող դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "USERID", general.performer)
		' Հաճախորդ դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "CLICODE", general.client)
		' Անվանում դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "NAME", general.name)
		' Անգլերեն անվանում դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "ENAME", general.englName)
		' Հաշիվ դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "FEEACC", general.account)
		' Արժույթ դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "FEECUR", general.curr)
		' Առավելագույն գումար դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "MAXSUM", general.maxPrice)
		' Ծառայություններ գրիդի լրացում
		for i = 0 to general.servGridRowCount - 1
				Call Fill_General_ServicesGrid(general.services(i), i)
		next
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''CommunalPayment_Additional'''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պարբերական կոմունալ վճարումների պայմանագրի Լրացուցիչ բաժնի լրացման կլաս
Class CommunalPayment_Additional
		public openDate
		public lastDate
		public payDays
		public payDays_to
		public informClient
		public useClientEmail
		public otherEmail
		public accsConnentScheme
		public useClientScheme
		public useCardAccs
		public addInfo
		public lastOpersDate
		public lastCompletedDate
		public closeDate
		private sub Class_Initialize()
				openDate = ""
				lastDate = ""
				payDays = "" 
				payDays_to = ""
				informClient = 0
				useClientEmail = 0
				otherEmail = ""
				accsConnentScheme = ""
				useClientScheme = 0
				useCardAccs = 0
				addInfo = ""
				lastOpersDate = ""
				lastCompletedDate = ""
				closeDate = ""
		end sub
End Class

Function New_CommunalPayment_Additional()
		Set New_CommunalPayment_Additional = new CommunalPayment_Additional
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''Fill_CommunalPayment_Other'''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պարբերական կոմունալ վճարումների պայմանագրի Լրացուցիչ բաժնի լրացման պրոցեդուրա
'other - Պարբերական կոմունալ վճարումների պայմանագրի Լրացուցիչ բաժնի լրացման կլաս
Sub Fill_CommunalPayment_Other(other)
		Call GoTo_ChoosedTab(2)
		' Բացման ամսաթիվ դաշտի լրացում 
		Call Rekvizit_Fill("Document", 2, "General", "SDATE", other.openDate)
		' Վերջին ամսաթիվ դաշտի լրացում
		Call Rekvizit_Fill("Document", 2, "General", "EDATE", other.lastDate)
		' Վճարման օրեր սկիզբ դաշտի լրացում
		Call Rekvizit_Fill("Document", 2, "General", "SDAY", other.payDays)
		' Վճարման օրեր վերջ դաշտի լրացում
		Call Rekvizit_Fill("Document", 2, "General", "LDAY", other.payDays_to)
		if other.informClient = 1 then 
		  ' Տեղեկացնել հաճախորդին դաշտի լրացում
				Call Rekvizit_Fill("Document", 2, "CheckBox", "CLINOT", other.informClient)
        ' Օգտ հաճախորդի Էլ. հասցե դաշտի լրացում
        Call Rekvizit_Fill("Document", 2, "CheckBox", "USECLIEMAIL", other.useClientEmail)
				if other.useClientEmail = 0 then
				    ' Այլ էլ. հասցե դաշտի լրացում
						Call Rekvizit_Fill("Document", 2, "General", "EMAIL", other.otherEmail)
				end if
		end if
		' Հաշիվների փոխկապակցման սխեմա դաշտի լրացում
		Call Rekvizit_Fill("Document", 2, "General", "ACCCONNECT", other.accsConnentScheme)
		' Օգտագործել հաճախորդի սխեման դաշտի լրացում
		Call Rekvizit_Fill("Document", 2, "CheckBox", "USECLISCH", other.useClientScheme)
		' Օգտագործել քարտային հաշիվները դաշտի լրացում
		Call Rekvizit_Fill("Document", 2, "CheckBox", "FEEFROMCARD", other.useCardAccs)
		' Լրացուցիչ ինֆորմացիա դաշտի լրացում
		Call Rekvizit_Fill("Document", 2, "General", "COMM", other.addInfo)
		' Վերջ. գործ-ն ամսաթիվ դաշտի լրացում
		Call Rekvizit_Fill("Document", 2, "General", "LASTOPDATE", other.lastOpersDate)
		' Վերջ. ավարտված գործ-ն ամսաթիվ դաշտի լրացում
		Call Rekvizit_Fill("Document", 2, "General", "LASTCOMPLETE", other.lastCompletedDate)
		' Փակման ամսաթիվ դաշտի լրացում
		Call Rekvizit_Fill("Document", 2, "General", "DATECLOSE", other.closeDate)
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''CommunalPayment''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պարբերական կոմունալ վճարումների պայմանագրի լրացման կլաս
Class CommunalPayment
		public fisn
		public general 
		public other
		public docNum
		private sub Class_Initialize()
				docNum = ""
				fisn = ""
				Set general = New_CommunalPayment_General()
				Set other = New_CommunalPayment_Additional()
		end sub
End Class 

Function New_CommunalPayment(gridCount)
		rowCount = gridCount
		Set New_CommunalPayment = new CommunalPayment
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''Fill_CommunalPayment'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պարբերական կոմունալ վճարումների պայմանագրի լրացման պրոցեդուրա
'communalPay - Պարբերական կոմունալ վճարումների պայմանագրի լրացման կլաս
Sub Fill_CommunalPayment(communalPay)
		' Վերցնել "Պայմանագրի ISN-ը" դաշտի արժեքը
		communalPay.fisn = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
		' Վարկային պայմանագրի համարի ստացում
		communalPay.docNum = Get_Rekvizit_Value("Document", 1, "General", "CODE")
		' Ընդհանուր բաժնի լրացում
		Call Fill_CommunalPayment_General(communalPay.general)
		' Այլ բաժնի լրացում
		Call Fill_CommunalPayment_Other(communalPay.other)
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''Create_CommunalPayment'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պարբերական կոմունալ վճարումների պայմանագրի ստեղծման պրոցեդուրա
'communalPay - Պարբերական կոմունալ վճարումների պայմանագրի լրացման կլաս
'folderName - պայմանագիր ճանապարհը
Sub Create_CommunalPayment(folderName, communalPay)
		wTreeView.DblClickItem(folderName & "Üáñ å³ñµ»ñ. ÏáÙáõÝ³É í×³ñ. å³ÛÙ³Ý³·Çñ")
		if wMDIClient.WaitVBObject("frmASDocForm", 3000).Exists then
				Call Fill_CommunalPayment(communalPay)
				Call ClickCmdButton(1, "Î³ï³ñ»É")
		else 
				Log.Error "Can't open Communal Payment(frmASDocForm) widow.", "", pmNormal, ErrorColor
		end if
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''CustomerService_General''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Հաճախորդի սպասարկման պայմանագրի Ընդհանուր բաժնի լրացման կլաս
Class CustomerService_General
		public num
		public AgreeType
		public standard
		public fill
		public client
		public name 
		public englName
		public account
		public curr
		public openDate
		public closeDate
		public firstPayDay
		public term
		public period_mounth
		public period_day
		public nonWorkigDays
		public calculateScheme
		public supportPrice
		public VATFired
		public office 
		public department
		public accessType
		private sub Class_Initialize()
				num = ""
				AgreeType = ""
				standard = ""
				fill = 0
				client = ""
				name = ""
				englName = ""
				account = ""
				curr = ""
				openDate = ""
				closeDate = ""
				firstPayDay = ""
				term = ""
				period_mounth = ""
				period_day = ""
				nonWorkingDays = ""
				calculateScheme = ""
				supportPrice = ""
				VATFired = ""
				office = ""
				department = ""
				accessType = ""
		end sub
End Class

Function New_CustomerService_General()
		Set New_CustomerService_General = new CustomerService_General
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''Fill_CustomerService_General''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Հաճախորդի սպասարկման պայմանագրի Ընդհանուր բաժնի լրացման պրոցեդուրա
'charge - Հաճախորդի սպասարկման պայմանագրի Ընդհանուր բաժնի լրացման կլաս
Sub Fill_CustomerService_General(general)
		Call GoTo_ChoosedTab(1)
		' N դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "CONTCODE", general.num)
		' Տիպ դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "TYPE", general.AgreeType)
		' Ստանդարտ դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "PCSTAND", general.standard)
		' Լրացնել դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "CheckBox", "FILLSTAND", general.fill)
		' Հաճախորդ դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "CLICODE", general.client)
		' Անվանում դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "NAME", general.name)
		' Անգլերեն անվանում դաշտի լարցում
		Call Rekvizit_Fill("Document", 1, "General", "ENAME", general.englName)
		' Հաշիվ դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "ACC", general.account)
		' Արժույթ դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "CUR", general.curr)
		' Բացման ամսաթիվ դաշտի լրացում 
		Call Rekvizit_Fill("Document", 1, "General", "OPENDATE", general.openDate)
		' Փակման ամսաթիվ դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "DATECLOSE", general.closeDate)
		' Առաջին վճարման օր դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "PAYDATE", general.firstPayDay)
		' Ժամկետ դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "ENDDATE", general.term)
		' Պարբերունթուն դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "PERIOD", general.period_mounth & "[Tab]" & general.period_day)
		' Ոչ աշխատանքային օրերի շրջանցում դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "NONWORKDAYS", general.nonWorkingDays)
		' Հաշվարկման սխեմա դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "SERVSCHEM", general.calculateScheme)
		' Սպասարկման վարձ (AMD) դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "SERVFEE", general.supportPrice)
		' ԱԱՀ-ով հարկվող
		Call Rekvizit_Fill("Document", 1, "General", "VATMETH", general.VATFired)
		' Գրասենյակ դաշտի լարցում
		Call Rekvizit_Fill("Document", 1, "General", "ACSBRANCH", general.office)
		' Բաժին դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "ACSDEPART", general.department)
		' Հասան-ն տիպ դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "ACSTYPE", general.accessType)
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''CustomerService_Charge'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Հաճախորդի սպասարկման պայմանագրի Գանձում բաժնի լրացման կլաս
Class CustomerService_Charge
		public chargeAcc
		public curr
		public gridRowCount
		public chargeGrid()
		public otherAccs
		public chargeCardAcc
		public chargeAviliableRes
		public debtMaxPrice
		public calculateManner
		public chargeType
		public addInfo
		private sub Class_Initialize()
				chargeAcc = ""
				curr = ""
				gridRowCount = rowCount
				Redim chargeGrid(gridRowCount)
				otherAccs = ""
				chargeCardAcc = ""
				chargeAviliableRes = ""
				debtMaxPrice = ""
				calculateManner = ""
				chargeType = ""
				addInfo = ""
		end sub
End Class

Function New_CustomerService_Charge(row_count)
		rowCount = row_count
		Set New_CustomerService_Charge = new CustomerSrvice_Charge
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''Fill_CustomerService_Charge'''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Հաճախորդի սպասարկման պայմանագրի Գանձում բաժնի լրացման պրոցեդուրա
'charge - Հաճախորդի սպասարկման պայմանագրի Գանձում բաժնի լրացման կլաս
Sub Fill_CustomerService_Charge(charge)
		Dim DocGrid, i
		Set DocGrid = wMDIClient.vbObject("frmASDocForm").vbObject("TabFrame").vbObject("DocGrid")
		Call GoTo_ChoosedTab(2)
		' Գանձման/համալրման հաշիվ դաշտի լրացում
		Call Rekvizit_Fill("Document", 2, "General", "FEEACC", charge.chargeAcc)
		' Արժույթ դաշտի լրացում
		Call Rekvizit_Fill("Document", 2, "General", "FEECUR", charge.curr)
		' Որպես գանձման/համալրման հաշիվ ըստ հերթականության ընտրել գրդի լրացում
		for i = 0 to gridRowCount - 1
				with DocGrid
						.Row = i
						' Տ. դաշտի լրացում
		    .Col = 0
		    .Keys(charge.chargeGrid(i) & "[Right]")
		  end with
		  DocGrid.Keys("[Home][Up]")
		  BuiltIn.Delay(1000)
		next
		' Այլ հաշիվներ (փոխկապ. սխեմա) դաշտի լրացում
		Call Rekvizit_Fill("Document", 2, "General", "ACCCONNECT", charge.otherAccs)
		' Գանձել նաև քարտային հաշիվներից դաշտի լրացում
		Call Rekvizit_Fill("Document", 2, "CheckBox", "FEEFROMCARD", charge.chargeCardAcc)
		' Գանձել առկա միջոցների չափով դաշտի լրացում
		Call Rekvizit_Fill("Document", 2, "CheckBox", "TAKEAVLB", charge.chargeAviliableRes)
		' Պարտքի առավելագույն գումար դաշտի լրացում
		Call Rekvizit_Fill("Document", 2, "General", "MAXDEBT", charge.debtMaxPrice)
		' Հաշվարկման եղանակ դաշտի լրացում
		Call Rekvizit_Fill("Document", 2, "General", "DEBTCALCMETH", charge.calculateManner)
		' Գանձման տիպ դաշտի լրացում
		Call Rekvizit_Fill("Document", 2, "General", "FEETYPE", charge.chargeType)
		' Լրացուցիչ ինֆորմացիա դաշտի լրացում
		Call Rekvizit_Fill("Document", 2, "General", "ADDINFO", charge.addInfo)
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''CustomerService''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Հաճախորդի սպասարկման պայմանագրի լրացման կլաս
Class CustomerService
		public fisn
		public genral 
		public charge
		private sub Class_Initialize()
				fisn = ""
				generel = New_CustomerService_General()
				charge = New_CustomerService_Charge()
		end sub
End Class 

Function New_CustomerService()
		Set New_CustomerService = new CustomerService
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''Fill_CustomerService'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Հաճախորդի սպասարկման պայմանագրի լրացման պրոցեդուրա
'customService - Հաճախորդի սպասարկման պայմանագրի լրացման կլաս
Sub Fill_CustomerService(customService)
		' Վերցնել "Պայմանագրի ISN-ը" դաշտի արժեքը
		customService.fisn = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
		' Ընդհանում բաժնի լրացում
		Call Fill_CustomerService_General(customService.general)
		' Գանձում բաժնի լրացում
		Call Fill_CustomerService_Charge(customService.charge)
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''Create_CustomerService'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Հաճախորդի սպասարկման պայմանագրի ստեղծման պրոցեդուրա
'customService - Հաճախորդի սպասարկման պայմանագրի լրացման կլաս
'folderName - պայմանագիր ճանապարհը
Sub Create_CustomerService(folderName, customService)
		wTreeView.DblClickItem(folderName & "Üáñ ëå³ë³ñÏÙ³Ý å³ÛÙ³Ý³·Çñ")
		if wMDIClient.WaitVBObject("frmASDocForm", 3000).Exists then
				Call Fill_CustomerService(customService)
				Call ClickCmdButton(1, "Î³ï³ñ»É")
		else 
				Log.Error "Can't open Customer Service(frmASDocForm) widow.", "", pmNormal, ErrorColor
		end if
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''ServiceAgreeType'''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Սպասարկման պայամանգրի տիպ պայմանագրի լրացման կլաս
Class ServiceAgreeType
		public fisn
		public code 
		public name
		public englName
		public onlyOneAgreeForCli
		public onlyOneAgreeForAcc
		public officeDepartNotInherit
		private sub Class_Initialize()
				fisn = ""
				code = ""
				name = ""
				englName = ""
				onlyOneAgreeForCli = 0
				onlyOneAgreeForAcc = 0
				officeDepartNotInherit = 0
		end sub
End Class 

Function New_ServiceAgreeType()
		Set New_ServiceAgreeType = new ServiceAgreeType
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''Fill_ServiceAgreeType'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Սպասարկման պայամանգրի տիպ պայմանագրի լրացման պրոցեդուրա
'serviceType - Սպասարկման պայամանգրի տիպ պայմանագրի լրացման կլաս
Sub Fill_ServiceAgreeType(serviceType)
		' Վերցնել "Պայմանագրի ISN-ը" դաշտի արժեքը
		serviceType.fisn = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
		' Կոդ դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "CODE", serviceType.code)
		' Անվանում դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "NAME", serviceType.name)
		' Անգլերեն անվանում դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "ENAME", serviceType.englName)
		' Միայն մեկ պայմանաիր հաճախորդի համար
		Call Rekvizit_Fill("Document", 1, "CheckBox", "ONLYONE", serviceType.onlyOneAgreeForCli)
		' Միայն մեկ պայմանագիր հաշվի համար
		Call Rekvizit_Fill("Document", 1, "CheckBox", "ONLYONEACC", serviceType.onlyOneAgreeForAcc)
		' Գրասենյակ/Բաժինը չժառանգել հաճախորդից
		Call Rekvizit_Fill("Document", 1, "CheckBox", "BRDEPONCLICHG", serviceType.officeDepartNotInherit)
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''Create_ServiceAgreeType'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Սպասարկման պայամանգրի տիպ պայմանագրի ստեղծման պրոցեդուրա
'serviceType - Սպասարկման պայամանգրի տիպ պայմանագրի լրացման կլաս
'folderName - պայմանագիր ճանապարհը
Sub Create_ServiceAgreeType(folderName, serviceType)
		wTreeView.DblClickItem(folderName & "Üáñ ëå³ë³ñÏÙ³Ý å³ÛÙ³Ý³·ñÇ ïÇå")
		if wMDIClient.WaitVBObject("frmASDocForm", 3000).Exists then
				Call Fill_ServiceAgreeType(serviceType)
				Call ClickCmdButton(1, "Î³ï³ñ»É")
		else 
				Log.Error "Can't open Service Agreement Type(frmASDocForm) widow.", "", pmNormal, ErrorColor
		end if 
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''CalculationScheme''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Հաշվարկման սխեմա պայմանագրի լրացման կլաս
Class CalculationScheme
		public fisn
		public code 
		public name
		public englName
		public supportPrice
		public period_mounth
		public period_day
		public doNotCharge
		public worksStarted
		private sub Class_Initialize()
				fisn = ""
				code = ""
				name = ""
				englName = ""
				supportPrice = ""
				period_mounth = ""
				period_day = ""
				doNotCharge = ""
				worksStarted = ""
		end sub
End Class 

Function New_CalculationScheme()
		Set New_CalculationScheme = new CalculationScheme
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''Fill_CalculationScheme''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Հաշվարկման սխեմա պայմանագրի լրացման պրոցեդուրա
'serviceType - Հաշվարկման սխեմա պայմանագրի լրացման կլաս
Sub Fill_CalculationScheme(serviceType)
		' Վերցնել "Պայմանագրի ISN-ը" դաշտի արժեքը
		serviceType.fisn = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
		' Կոդ դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "CODE", serviceType.code)
		' Անվանում դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "NAME", serviceType.name)
		' Անգլերեն անվանում դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "ENAME", serviceType.englName)
		' Սպասարկման վարձ (AMD) դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "SERVFEE", serviceType.supportPrice)
		' Պարբերություն դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "PERIOD", serviceType.period_mounth & "[Tab]" & serviceType.period_day)
		' Չգանձել Ենթ. պայմ. ունենալու դեպքում դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "NOCALCSUBSYS", serviceType.doNotCharge)
		' Գործում է սկսած դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "SDATE", serviceType.worksStarted)
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''Create_CalculationScheme'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Հաշվարկման սխեմա պայմանագրի ստեղծման պրոցեդուրա
'serviceType - Հաշվարկման սխեմա պայմանագրի լրացման կլաս
'folderName - պայմանագիր ճանապարհը
Sub Create_CalculationScheme(folderName, serviceType)
		wTreeView.DblClickItem(folderName & "Üáñ Ñ³ßí³ñÏÙ³Ý ëË»Ù³")
		if wMDIClient.WaitVBObject("frmASDocForm", 3000).Exists then
				Call Fill_CalculationScheme(serviceType)
				Call ClickCmdButton(1, "Î³ï³ñ»É")
		else 
				Log.Error "Can't open Calculation Scheme(frmASDocForm) widow.", "", pmNormal, ErrorColor
		end if 
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''PenaltyCalcScheme_Grid'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Տույժերի հաշվարկման սխեմա պայմանագրի գրիդի լրացման կլաս
Class PenaltyCalcScheme_Grid
		public mounth
		public midDiv
		public andOr
		public circulation
		public charge
		private sub Class_Initialze()
				mounth = ""
				midDiv = ""
				andOr = ""
				circulation = ""
				charge = "" 
		end sub
End Class

Function New_PenaltyCalcScheme_Grid()
		Set New_PenaltyCalcScheme_Grid = new PenaltyCalcScheme_Grid
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''Fill_PenaltyCalcScheme_Grid'''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Տույժերի հաշվարկման սխեմա պայմանագրի գրիդի լրացման պրոցեդուրա
'penaltyGrid - Տույժերի հաշվարկման սխեմա պայմանագրի գրիդի լրացման կլաս
Sub Fill_PenaltyCalcScheme_Grid(penaltyGrid, i)
		Dim DocGrid
		Set DocGrid = wMDIClient.vbObject("frmASDocForm").vbObject("TabFrame").vbObject("DocGrid")
  with DocGrid
				.Row = i
				' Ամիս դաշտի լարցում
    .Col = 0
    .Keys(gridService.mounth & "[Right]")
				' Միջին մնացորդ դաշտի լրացում
				.Col = 1
    .Keys(gridService.midDiv & "[Right]")
				' և/կամ դաշտի լրացում 
				.Col = 2
    .Keys(gridService.andOr & "[Right]")
				' Շրջանառություն դաշտի լրացում
				.Col = 3
    .Keys(gridService.circulation & "[Right]")
				' Գանձում դաշտի լրացում
				.Col = 4
    .Keys(gridService.charge & "[Right]")
  end with
  DocGrid.Keys("[Home][Up]")
  BuiltIn.Delay(1000)
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''PenaltyCalcScheme''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Տույժերի հաշվարկման սխեմա պայմանագրի լրացման կլաս
Class PenaltyCalcScheme
		public fisn
		public code 
		public name
		public englName
		public observedAccs
		public observedAccsTypes
		public doNotCharge
		public period_mounth
		public period_day
		public gridCount
		public pricesAMD()
		public worksStarted
		private sub Class_Initialize()
				fisn = ""
				code = ""
				name = ""
				englName = ""
				observedAccs = ""
				observedAccsTypes = ""
				period_mounth = ""
				period_day = ""
				doNotCharge = ""
				worksStarted = ""
				gridCount = rowCount
				Redim pricesAMD(gridCount)
		end sub
End Class 

Function New_PenaltyCalcScheme(row_Count)
		rowCount = row_Count
		Set New_PenaltyCalcScheme = new PenaltyCalcScheme
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''Fill_PenaltyCalcScheme''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Տույժերի հաշվարկման սխեմա պայմանագրի լրացման պրոցեդուրա
'penaltyScheme - Տույժերի հաշվարկման սխեմա պայմանագրի լրացման կլաս
Sub Fill_PenaltyCalcScheme(penaltyScheme)
		Dim i
		' Վերցնել "Պայմանագրի ISN-ը" դաշտի արժեքը
		penaltyScheme.fisn = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
		' Կոդ դաշտի լրացում 
		Call Rekvizit_Fill("Document", 1, "General", "CODE", penaltyScheme.code)
		' Անվանում դաշտի լարցում
		Call Rekvizit_Fill("Document", 1, "General", "NAME", penaltyScheme.name)
		' Անգլերեն անվանում դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "ENAME", penaltyScheme.englName)
		' Դիտարկվող հաշիվներ դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "INCACCOUNTS", penaltyScheme.observedAccs)
		' Դիտարկվող հաշիվների տիպեր դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "ACCSTYPE", penaltyScheme.observedAccsTypes)
		' Չգանձել Ենթ. պայմ. ունենեալու դեպքում դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "NOCALCSUBSYS", penaltyScheme.doNotCharge)
		' Գնաձման պարբերություն դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "PERIOD", penaltyScheme.period_mounth & "[Tab]" & penaltyScheme.period_day)
		' Գանձումներ (AMD) գրիդի լրացում
		for i = 0 to gridCount - 1
				Call Fill_PenaltyCalcScheme_Grid(penaltyScheme.pricesAMD(i), i)
		next
		' Գործում է սկսած դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "SDATE", penaltyScheme.worksStarted)
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''Create_PenaltyCalcScheme'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Տույժերի հաշվարկման սխեմա պայմանագրի ստեղծման պրոցեդուրա
'penaltyScheme - Տույժերի հաշվարկման սխեմա պայմանագրի լրացման կլաս
'folderName - պայմանագիր ճանապարհը
Sub Create_PenaltyCalcScheme(folderName, penaltyScheme)
		wTreeView.DblClickItem(folderName & "Üáñ ïáõÛÅ»ñÇ Ñ³ßí³ñÏÙ³Ý ëË»Ù³")
		if wMDIClient.WaitVBObject("frmASDocForm", 3000).Exists then
				Call Fill_PenaltyCalcScheme(penaltyScheme)
				Call ClickCmdButton(1, "Î³ï³ñ»É")
		else 
				Log.Error "Can't open Penalty Calculation Scheme(frmASDocForm) widow.", "", pmNormal, ErrorColor
		end if 
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''ServiceFeeAccounting''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Սպասարկման վարձի հաշվառում պայմանագրի լրացման կլաս
Class ServiceFeeAccounting
		public AgreeType
		public docN
		public agreement
		public client
		public agreeAcc
		public curr
		public calcScheme
		public supportPrice
		public period_mounth
		public period_day
		public calcDate
		public previousDebt
		public calcDebt
		public VATFired
		public calculate0
		public formulationDate
		public chargeAcc
		public chargeSumma
		public chargeSummaSec
		public incomeAcc
		public aim
		private sub Class_Initialize()
				AgreeType = ""
				docN = ""
				agreement = ""
				client = ""
				agreeAcc = ""
				curr = ""
				calcScheme = ""
				supportPrice = ""
				period_mounth = ""
				period_day = ""
				calcDate = ""
				previousDebt = ""
				calcDebt = ""
				VATFired = ""
				calculate0 = 0
				formulationDate = ""
				chargeAcc = ""
				chargeSumma = ""
				chargeSummaSec = ""
				incomeAcc = ""
				aim = ""
		end sub
End Class

Function New_ServiceFeeAccounting()
		Set New_ServiceFeeAccounting = new ServiceFeeAccounting
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''Fill_ServiceFeeAccounting''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Սպասարկման վարձի հաշվառում պայմանագրի լրացման պրոցեդուրա
'serviceFeeAcc - Սպասարկման վարձի հաշվառում պայմանագրի լրացման կլաս
Sub Fill_ServiceFeeAccounting(serviceFeeAcc)
  ' Տիպ դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "TYPE", serviceFeeAcc.AgreeType)
		' Փաստաթղթի N դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "DOCNUM", serviceFeeAcc.docN)
		' Պայմանագիր դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "CONTRACT", serviceFeeAcc.agreement)
		' Հաճախորդ դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "CLICODE", serviceFeeAcc.client)
		' Պայմանագրի հաշիվ դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "ACC", serviceFeeAcc.agreeAcc)
		' Արժ. դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "ACCCUR", serviceFeeAcc.curr)
		' Հաշվարկման սխեմա դաշտի լրացում 
		Call Rekvizit_Fill("Document", 1, "General", "SERVSCHEM", serviceFeeAcc.calcScheme)
		' Սպասարկման վարձ (AMD) դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "SERVFEE", serviceFeeAcc.supportPrice)
		' Պարբերություն դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "PERIOD", serviceFeeAcc.period_mounth & "[Tab]" & serviceFeeAcc.period_day)
		' Հաշվարկման ամսաթիվ դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "SERVDATE", serviceFeeAcc.calcDate) 
		' Նախորդ պարտք դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "PREVDEBT", serviceFeeAcc.previousDabt)
		' Հաշվարկված պարտք դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "DEBT", serviceFeeAcc.calcDebt)
		' ԱԱՀ-ով հարկվող դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "VATMETH", serviceFeeAcc.VATFired)
		' Հաշվարկել 0 դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "CheckBox", "CALCZERO", serviceFeeAcc.calculate0)
		' Ձևակերպել ամսաթիվ դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "DATE", serviceFeeAcc.forulationDate)
		' Գանձման հաշիվ դաշտի լրացում 
		Call Rekvizit_Fill("Document", 1, "General", "FEEACC", serviceFeeAcc.chargeAcc)
		' Գանձման գում. դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "FEESUMMA", serviceFeeAcc.chargeSumma)
		' Գանձման արժ. դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "FEECUR", serviceFeeAcc.chargeSummaSec)
		' Եկամտի հաշիվ դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "CHRGACC", serviceFeeAcc.incomeAcc)
		' Նպատակ դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "AIM", serviceFeeAcc.aim)
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''Create_ServiceFeeAccounting''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Սպասարկման վարձի հաշվառում պայմանագրի ստեղծման պրոցեդուրա
'serviceFeeAcc - Սպասարկման վարձի հաշվառում պայմանագրի լրացման կլաս
'folderName - պայմանագիր ճանապարհը
Sub Create_ServiceFeeAccounting(folderName, serviceFeeAcc)
		wTreeView.DblClickItem(folderName & "êå³ë³ñÏÙ³Ý í³ñÓÇ Ñ³ßí³éáõÙ")
		if wMDIClient.WaitVBObject("frmASDocForm", 3000).Exists then
				Call Fill_ServiceFeeAccounting(serviceFeeAcc)
				Call ClickCmdButton(1, "Î³ï³ñ»É")
		else 
				Log.Error "Can't open Service Fee Accounting(frmASDocForm) widow.", "", pmNormal, ErrorColor
		end if
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''DebtImportScheme'''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պարտքի ներմուծման սխեմա պայմանագրի լրացման կլաս
Class DebtImportScheme
		public code
		public name
		public englName
		public period_mounth
		public period_day
		public comment
		private sub Class_Initialize()
				code = ""
				name = ""
				englName = ""
				period_mounth = ""
				period_day = ""
				comment = ""
		end sub
End Class

Function New_DebtImportScheme()
		Set New_DebtImportScheme = new DebtImportScheme
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''Fill_DebtImportScheme'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պարտքի ներմուծման սխեմա պայմանագրի լրացման պրոցեդուրա
'debtImpScheme - Պարտքի ներմուծման սխեմա պայմանագրի լրացման կլաս
Sub Fill_DebtImportScheme(debtImpScheme)
  'Կոդ դաշտի լրացում 
		Call Rekvizit_Fill("Document", 1, "General", "TYPE", debtImpScheme.code)
		' Անվանում դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "DOCNUM", debtImpScheme.name)
		' Անգլերեն անվանում դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "CONTRACT", debtImpScheme.englName)
		' Պարբերություն դաշտի լրացում 
		Call Rekvizit_Fill("Document", 1, "General", "PERIOD", debtImpScheme.period_mounth & "[Tab]" & debtImpScheme.period_day)
		' Մեկնաբանություն դաշտի լրացում
		Call Rekvizit_Fill("Document", 1, "General", "SERVDATE", debtImpScheme.comment)
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''Create_DebtImportScheme'''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պարտքի ներմուծման սխեմա պայմանագրի ստեղծման պրոցեդուրա
'debtImpScheme - Պարտքի ներմուծման սխեմա պայմանագրի լրացման կլաս
'folderName - պայմանագիր ճանապարհը
Sub Create_DebtImportScheme(folderName, debtImpScheme)
		wTreeView.DblClickItem(folderName & "êå³ë³ñÏÙ³Ý í³ñÓÇ Ñ³ßí³éáõÙ")
		if wMDIClient.WaitVBObject("frmASDocForm", 3000).Exists then
				Call Fill_DebtImportScheme(debtImpScheme)
				Call ClickCmdButton(1, "Î³ï³ñ»É")
		else 
				Log.Error "Can't open Debt Import Scheme(frmASDocForm) widow.", "", pmNormal, ErrorColor
		end if
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''PeriodicWorkingDocuments'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Աշխատանքային փաստաթղթեր պատուհանի լրացման կլաս
Class PeriodicWorkingDocuments
		public startDate
		public endDate
		public curr
		public performers
		public docType
		public commonPaySys
		public addPaySys
		public note		
		public office
		public section
		public showType
		public fill
		private Sub Class_Initialize()
				startDate = aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%d%m%y")
				endDate = aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%d%m%y")
				curr = ""
				performers = ""
				docType = ""
				commonPaySys = ""
				addPaySys = ""
				note = ""			
				office = ""
				section = ""
				showType = "Oper"
				fill = 0
		End Sub
End Class

Function New_PeriodicWorkingDocuments()
		Set New_PeriodicWorkingDocuments = new PeriodicWorkingDocuments
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''Fill_PeriodicWorkingDocuments'''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Աշխատանքային փաստաթղթեր պատուհանի լրացման պրոցեդուրա
'WorkingDocs - պատուհանի լրացման կլաս
Sub Fill_PeriodicWorkingDocuments(WorkingDocs)
  ' Ժամանակահատված սկզբնական դաշտի լրացում 
		Call Rekvizit_Fill("Dialog", 1, "General", "PERN", "![End]" & "[Del]" & WorkingDocs.startDate)
		' Ժամանակահատված վերջնական դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "PERK", "![End]" & "[Del]" & WorkingDocs.endDate)
		' Արժույթ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "CUR", WorkingDocs.curr)
		 ' Կատարողներ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "USER", "![End]" & "[Del]" & WorkingDocs.performers)
		' Փաստաթղթի տեսակ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "DOCTYPE", WorkingDocs.docType)
		' Ընդ. վճ. համակարգ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "PAYSYSIN", WorkingDocs.commonPaySys)
		' Ուղ. վճ. համակարգ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "PAYSYSOUT", WorkingDocs.addPaySys)				
		' Նշում դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "PAYNOTES", WorkingDocs.note)
		' Գրասենյակ բաժնի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "ACSBRANCH", WorkingDocs.office)
		' Բաժին դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "ACSDEPART", WorkingDocs.section)
		' Դտելու ձև դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "SELECTED_VIEW", WorkingDocs.showType)
		' Լրացնել դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "EXPORT_EXCEL", WorkingDocs.fill)
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''GoTo_PeriodicWorkingDocuments'''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Աշխատանքային փաստաթղթեր թղթապանակ մուտք գործելու պրոցեդուրա
'folderName - գտնբելու ճանապարհը
'WorkingDocs - պատուհանի լրացման կլաս
Sub GoTo_PeriodicWorkingDocuments(folderName, WorkingDocs)
		wTreeView.DblClickItem(folderName & "²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ")
		if p1.WaitVBObject("frmAsUstPar", 3000).Exists then
				Call Fill_PeriodicWorkingDocuments(WorkingDocs)
				Call ClickCmdButton(2, "Î³ï³ñ»É")
		else 
				Log.Error "Can't open frmAsUstPar widow.", "", pmNormal, ErrorColor
		end if
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''PeriodicActionsAgree''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պարբերական գործ. պայմանագրեր պատուհանի լրացման կլաս
Class PeriodicActionsAgree
		public agreeN
		public client 
		public showClosedAgree
		public performer
		public note
		public note2 
		public note3
		public showOpened
		public operationType
		public calculateMethod
		public office
		public department
		private Sub Class_Initialize()
		  agreeN = ""
				client = ""
				showClosedAgree = 0
				performer = ""
				note = ""
				note2 = ""
				note3 = ""
				showOpened = 0
				operationType = ""
				calculateMethod = ""
				office = ""
				department = ""
		End Sub
End Class

Function New_PeriodicActionsAgree()
		Set New_PeriodicActionsAgree = new PeriodicActionsAgree
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''Fill_PeriodicWorkingDocuments'''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պարբերական գործ. պայմանագրեր պատուհանի լրացման պրոցեդուրա
'periodicAct - պատուհանի լրացման կլաս
Sub Fill_PeriodicActionsAgree(periodicAct)
  ' Պայմանագրի N դաշտի լրացում 
		Call Rekvizit_Fill("Dialog", 1, "General", "DAGRNUM", periodicAct.agreeN)
		' Հաճախորդ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "DCLICODE", periodicAct.client)
		' Ցույց տալ փակված պայմանագրերը նշիչի լրացում
		Call Rekvizit_Fill("Dialog", 1, "CheckBox", "DCLOSED", periodicAct.showClosedAgree)
		 ' Կատարողներ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "DUSER", periodicAct.performer)
		' Նշում դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "DNOTE1", periodicAct.note)
		' Նշում 2 դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "DNOTE2", periodicAct.note2)
		' Նշում 3 դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "DNOTE3", periodicAct.note3)
		' Ցույց տալ բացված տեսքով դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "CheckBox", "DADVANCED", periodicAct.showOpened)
		' Գործ. տեսակ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "DTYPE", periodicAct.operationType)
		' Հաշվարկման եղանակ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "DCALCTYPE", periodicAct.calculateMethod)
		' Գրասենյակ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "DACSBRANCH", periodicAct.office)
		' Բաժին դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "DACSDEPART", periodicAct.department)
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''GoTo_PeriodicWorkingDocuments'''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պարբերական գործ. պայմանագրեր թղթապանակ մուտք գործելու պրոցեդուրա
'folderName - գտնբելու ճանապարհը
'periodicAct - պատուհանի լրացման կլաս
Sub GoTo_PeriodicActionsAgree(folderName, periodicAct)
		wTreeView.DblClickItem(folderName & "ä³ñµ»ñ³Ï³Ý ·áñÍ. å³ÛÙ³Ý³·ñ»ñ")
		if p1.WaitVBObject("frmAsUstPar", 3000).Exists then
				Call Fill_PeriodicActionsAgree(periodicAct)
				Call ClickCmdButton(2, "Î³ï³ñ»É")
		else 
				Log.Error "Can't open frmAsUstPar widow.", "", pmNormal, ErrorColor
		end if
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''Check_PeriodicExisting'''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պարբերական գործ. պայմանագրեր թղթապանակում փաստատթղթի առկայության ստուգում 
'Ֆունկցիան վերադարձնում է true, եթե պայմանագիրը առկա է և false, եթե այն բացակայում է 
'agreement - Պարբերական գործ. պայմանագրեր թղթապանակ մուտք գործելու կլասի օբյեկտ
'folderName - Պայմանագրի ճանապարհը
'docNum - Պայմանագրի համարը
Function Check_PeriodicExisting(folderName, agreement, docNum)
  Dim isExist : isExist = false
		Dim grid
  Call GoTo_PeriodicActionsAgree(folderName, agreement)
  wMDIClient.Refresh
  if wMDIClient.WaitVBObject("frmPttel", 3000).Exists then
		  Set grid = wMDIClient.vbObject("frmPttel").vbObject("tdbgView")
		  grid.MoveFirst()
				for i = 0 to grid.ApproxCount - 1
						if Trim(Grid.Columns(0).Value) = Trim(docNum) then
								isExist = true
								exit for
						else
								grid.MoveNext()
						end if
				next
				if not isExist then
						Log.Error "There are no document with specified ID or there are more than one. There are " &_
						wMDIClient.vbObject("frmPttel").vbObject("tdbgView").ApproxCount & " rows.", "", pmNormal, ErrorColor
				end if
  end if
  Check_PeriodicExisting = isExist
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''MakePayment''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Կատարել վճարում գործողություն
'calcDate - Գործողության ամասաթիվ
'checkPeriodic - Ստուգել պարբ. և կատարման օրերը
'sendMail - Ուղարկել էլ. նամակ
Sub MakePayment(openDate, checkPeriodic, sendMail)
		Call wMainForm.MainMenu.Click(c_AllActions)
		Call wMainForm.PopupMenu.Click(c_MakePayment)
		
		if p1.WaitVBObject("frmAsUstPar", 3000).Exists then
				' Լրացնել Գործողության ամսաթիվ դաշտը 
				Call Rekvizit_Fill("Dialog", 1, "General", "OPDATE", openDate)
				' Լրացնել Ստուգել պարբ. և կատարման օրերը նշիչը 
				Call Rekvizit_Fill("Dialog", 1, "CheckBox", "CHECKDATE", checkPeriodic)
				' Լրացնել Ուղարկել էլ. նամակ նշիչը
				Call Rekvizit_Fill("Dialog", 1, "CheckBox", "SENDMAIL", sendMail)
				' Սեղմել Կատարել կոճակը
				Call ClickCmdButton(2, "Î³ï³ñ»É")
				Call MessageExists(2, "¶áñÍáÕáõÃÛáõÝÝ»ñÇ µ³ñ»Ñ³çáÕ ³í³ñï")
				Call ClickCmdButton(5, "Ok")
		else
				Log.Error "Can't open frmAsUstPar widow.", "", pmNormal, ErrorColor
		end if
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''PaymentShow''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Վճարումների դիտում գործողություն
'startDate - Ժամանակահատված սկզբնական
'endDate - Ժամանակահատված վերջնական
Function PaymentView(startDate, endDate, rowCount)
		Dim isExist : isExist = true
		
		Call wMainForm.MainMenu.Click(c_AllActions)
		Call wMainForm.PopupMenu.Click(c_PaymentView)
		
		if p1.WaitVBObject("frmAsUstPar", 3000).Exists then
				' Լրացնել Ժամանակահատված սկզբնական դաշտը 
				Call Rekvizit_Fill("Dialog", 1, "General", "DSDATE", startDate)
				' Լրացնել Ժամանակահատված վերջնական դաշտը 
				Call Rekvizit_Fill("Dialog", 1, "General", "DEDATE", endDate)
				' Սեղմսնլ Կատարել կոճակը
				Call ClickCmdButton(2, "Î³ï³ñ»É")
				
				wMDIClient.Refresh
		  If wMDIClient.vbObject("frmPttel_2").vbObject("tdbgView").ApproxCount <> rowCount Then
								Log.Error "There are no document with specified ID or there are more than one. There are " &_
								wMDIClient.vbObject("frmPttel_2").vbObject("tdbgView").ApproxCount & " rows.", "", pmNormal, ErrorColor
		      isExist = false
		  End If
		  PaymentView = isExist
		else
				Log.Error "Can't open frmAsUstPar widow.", "", pmNormal, ErrorColor
		end if
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''SelectRowByColumn'''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Ü³Ëáñáù µ³óí³Í ó³ÝÏÇó Ýß»É arrContractNumbers-Ç µáÉáñ ·ïÝí³Í å³ÛÙ³Ý³·ñ»ñÁ
'colIndex-Á µ³óí³Í ³ÕÛáõë³ÏáõÙ å³ÛÙÝ³·ñÇ Ñ³Ù³ñÁ å³ñáõÝ³ÏáÕ ëÛáõÝÝ ¿
Sub SelectRowByColumn(arrContractNumbers, colIndex)
		Dim grid, contractNumber, i
				
		if wMDIClient.WaitvbObject("frmPttel", 3000).Exists then
				Set grid = wMDIClient.vbObject("frmPttel").vbObject("tdbgView")
				for each contractNumber in arrContractNumbers
				  grid.MoveFirst()
						for i = 0 to grid.ApproxCount - 1
								if Trim(Grid.Columns(colIndex).Value) = Trim(contractNumber) then
										grid.Keys("[Ins]")
										' because after insert key cursor moves to the next row
										' if its not the last row in the grid
										grid.MovePrevious()
								end if
								grid.MoveNext()
						next
				next
		else 
				Log.Error "Can't open frmPttel widow.", "", pmNormal, ErrorColor
		end if
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''FieldsGroupEdit''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Դաշտերի խմբային խմբագրում պատուհանի լրացման կլաս
Class FieldsGroup_Edit
		public office
		public department
		public emptyPerformer
		public performer
		public startDate
		public endDate
		public emptyEndDate
		public doInEveryCall
		public period_mounth
		public period_day
		public editImplementDays
		public implementDays_start
		public implementDays_end
		public bypassNonWorkDays
		public informClient
		public emptyNote
		public note
		public emptyNote2
		public note2
		public emptyNote3
		public note3
		public emptyAddInfo
		public addInfo
		private sub Class_Initialize()
				office = ""
				department = ""
				emptyPerformer = 0
				performer = ""
				startDate = ""
				endDate = ""
				emptyEndDate = 0
				doInEveryCall = 0
				period_mounth = ""
				period_day = ""
				editImplementDays = 0
				implementDays_start = ""
				implementDays_end = ""
				bypassNonWorkDays = ""
				informClient = ""
				emptyNote = 0
				note = ""
				emptyNote2 = 0
				note2 = ""
				emptyNote3 = 0
				note3 = ""
				emptyAddInfo = 0
				addInfo = ""
		end sub
End Class

Function New_FieldsGroup_Edit()
		Set New_FieldsGroup_Edit = new FieldsGroup_Edit
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''Fill_FieldsGroupEdit''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Դաշտերի խմբային խմբագրում պատուհանի լրացման պրոցեդուրա
'groupEdit - Դաշտերի խմբային խմբագրում պատուհանի լրացման կլաս
Sub Fill_FieldsGroupEdit(groupEdit)
  ' Գարսենյակ դաշտի լրացում 
		Call Rekvizit_Fill("Dialog", 1, "General", "ACSBRANCH", groupEdit.office)
		' Բաժին դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "ACSDEPART", groupEdit.department)
		' Դատարկել կատարողին նշիչի լրացում
		Call Rekvizit_Fill("Dialog", 1, "CheckBox", "CLEARUSERID", groupEdit.emptyPerformer)
		if groupEdit.emptyPerformer <> 1 then
				 ' Կատարող դաշտի լրացում
				Call Rekvizit_Fill("Dialog", 1, "General", "USERID", groupEdit.performer)
		end if
		' Սկզբի ամսաթիվ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "SDATE", groupEdit.startDate)
		' Վերջի ամսաթիվ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "EDATE", groupEdit.endDate)
		' Դատարկել վերջին ամսաթիվը նշիչի լրացում
		Call Rekvizit_Fill("Dialog", 1, "CheckBox", "CLEAREDATE", groupEdit.emptyEndDate)
		' Կատարել ամեն կանչի ժամանակ նշիչի լրացում
		Call Rekvizit_Fill("Dialog", 1, "CheckBox", "CALCALWAYS", groupEdit.doInEveryCall)
		if groupEdit.doInEveryCall <> 1 then
				' Պարբերություն դաշտի լրացում
				Call Rekvizit_Fill("Dialog", 1, "General", "PERIODICITY", groupEdit.period_mounth & "[Tab]" & groupEdit.period_day)
		end if
		' Խմբագրել կատարման օրերը նշիչի լրացում
		Call Rekvizit_Fill("Dialog", 1, "CheckBox", "EDITDAYS", groupEdit.editImplementDays)
		if groupEdit.editImplementDays = 1 then 
				' Կատարման օրեր սկզբի դաշտի լրացում
				Call Rekvizit_Fill("Dialog", 1, "General", "SDAY", groupEdit.implementDays_start)
				' Կատրման օրեր վերջի դաշտի լրացում
				Call Rekvizit_Fill("Dialog", 1, "General", "LDAY", groupEdit.implementDays_end)
		end if
		' Ոչ աշխատանքային օրերի շրջանցում դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "NONWORKDAYS", groupEdit.bypassNonWorkDays)
		' Տեղեկացնել հաճախորդին 1-նշանակել, 2-հեռացնել դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "CLINOT", groupEdit.informClient)
		' Դատարկել Նշումը նշիչի լրացում
		Call Rekvizit_Fill("Dialog", 1, "CheckBox", "CLEARNOTE1", groupEdit.emptyNote)
		if groupEdit.emptyNote <> 1 then
				' Նշում դաշտի լրացում
				Call Rekvizit_Fill("Dialog", 1, "General", "NOTE1", groupEdit.note)
		end if
		' Դատարկել Նշում 2-ը նշիչի լրացում
		Call Rekvizit_Fill("Dialog", 1, "CheckBox", "CLEARNOTE2", groupEdit.emptyNote2)
		if groupEdit.emptyNote2 <> 1 then
				' Նշում 2 դաշտի լրացում
				Call Rekvizit_Fill("Dialog", 1, "General", "NOTE2", groupEdit.note2)
		end if
		' Դատարկել Նշում 3-ը նշիչի լրացում
		Call Rekvizit_Fill("Dialog", 1, "CheckBox", "CLEARNOTE3", groupEdit.emptyNote3)
		if groupEdit.emptyNote3 <> 1 then 
				' Նշում 3 դաշտի լրացում
				Call Rekvizit_Fill("Dialog", 1, "General", "NOTE3", groupEdit.note3)
		end if
		' Դատարկել լրացուցիչ ինֆորմացիան նշիչի լրացում
		Call Rekvizit_Fill("Dialog", 1, "CheckBox", "CLEARCOMM", groupEdit.emptyAddInfo)
		if groupEdit.emptyAddInfo <> 1 then 
				' Լրացուցիչ ինֆորմացիա դաշտի լրացում
				Call Rekvizit_Fill("Dialog", 1, "General", "COMM", groupEdit.addInfo)
		end if
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''FieldsGroupEdit'''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Դաշտերի խմբային խմբագրում գործողություն
'groupEdit - Դաշտերի խմբային խմբագրում պատուհանի լրացման կլաս
Sub FieldsGroupEdit(groupEdit)
 	Call wMainForm.MainMenu.Click(c_AllActions)
		Call wMainForm.PopupMenu.Click(c_FieldsGroupEdit)
		
		if p1.WaitVBObject("frmAsUstPar", 3000).Exists then
				Call Fill_FieldsGroupEdit(groupEdit)
				Call ClickCmdButton(2, "Î³ï³ñ»É")
				Call MessageExists(2, "ö³ëï³ÃÕÃ»ñÇ ÷á÷áËÙ³Ý Ñ³Ûï»ñÝ áõÕ³ñÏí³Í »Ý Ñ³ëï³ïÙ³Ý")
				Call ClickCmdButton(5, "Ok")
		else
				Log.Error "Can't open frmAsUstPar widow.", "", pmNormal, ErrorColor
		end if
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''ChangeRequests'''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Փոփոխման հայտեր պատուհանի լրացման կլաս
Class ChangeRequests
		public state
		public startDate
		public endDate
		public performer 
		public office
		public department
		private Sub Class_Initialize()
		  state = ""
				startDate = ""
				endDate = ""
				performer = ""
				office = ""
				department = ""
		End Sub
End Class

Function New_ChangeRequests()
		Set New_ChangeRequests = new ChangeRequests
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''Fill_ChangeRequests''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Փոփոխման հայտեր պատուհանի լրացման պրոցեդուրա
'chgRequests - պատուհանի լրացման կլաս
Sub Fill_ChangeRequests(chgRequests)
  ' Վիճակ դաշտի լրացում 
		Call Rekvizit_Fill("Dialog", 1, "General", "DSTATE", chgRequests.state)
		' Ժամանակահատված սկզբնական դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "DSDATE", chgRequests.startDate)
		' Ժամանակահատված վերջնական դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "DEDATE", chgRequests.endDate)
		 ' Կատարողներ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "DUSER", chgRequests.performer)
		' Գրասենյակ դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "DACSBRANCH", chgRequests.office)
		' Բաժին դաշտի լրացում
		Call Rekvizit_Fill("Dialog", 1, "General", "DACSDEPART", chgRequests.department)
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''GoTo_ChangeRequests'''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Փոփոխման հայտեր թղթապանակ մուտք գործելու պրոցեդուրա
'folderName - գտնբելու ճանապարհը
'chgRequests - պատուհանի լրացման կլաս
Sub GoTo_ChangeRequests(folderName, chgRequests)
		wTreeView.DblClickItem(folderName & "öá÷áËÙ³Ý Ñ³Ûï»ñ")
		if p1.WaitVBObject("frmAsUstPar", 3000).Exists then
				Call Fill_ChangeRequests(chgRequests)
				Call ClickCmdButton(2, "Î³ï³ñ»É")
		else 
				Log.Error "Can't open frmAsUstPar widow.", "", pmNormal, ErrorColor
		end if
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''GoTo_ChangeRequests'''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub Confirm_ChangeRequest(columnValue, colNum)
		BuiltIn.Delay(3000)
		if SearchInPttel("frmPttel", colNum, columnValue) then
				Call wMainForm.MainMenu.Click(c_AllActions)
				Call wMainForm.PopupMenu.Click(c_ToVerify)
				Call ClickCmdButton(2, "Î³ï³ñ»É")
		else 
				Log.Error "Can't find searched row.", "", pmNormal, ErrorColor
		end if
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''AgreeGroupClose''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պայմանագրերի խմբային փակում 
'date - ամսաթիվ դաշտի լրացման արժեք
Sub AgreeGroupClose(date)
		Call wMainForm.MainMenu.Click(c_AllActions)
		Call wMainForm.PopupMenu.Click(c_GroupClose)
		if p1.WaitVBObject("frmAsUstPar", 3000).Exists then
				Call Rekvizit_Fill(Dialog, 1, "General", "SELECTEDROWCOUNT", date)
				Call ClickCmdButton(2, "Î³ï³ñ»É")
				Call MessageExists(2, "ö³ëï³ÃÕÃ»ñÇ ÷á÷áËÙ³Ý Ñ³Ûï»ñÝ áõÕ³ñÏí³Í »Ý Ñ³ëï³ïÙ³Ý")
				Call ClickCmdButton(5, "Ok")
		else 
				Log.Error "Can't open frmAsUstPar widow.", "", pmNormal, ErrorColor
		end if
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''AgreeGroupOpen'''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Պայմանագրերի խմբային բացում 
'date - ամսաթիվ դաշտի լրացման արժեք
Sub AgreeGroupOpen()
		Call wMainForm.MainMenu.Click(c_AllActions)
		Call wMainForm.PopupMenu.Click(c_GroupOpen)
		if p1.WaitVBObject("frmAsMsgBox", 3000).Exists then
				Call MessageExists(2, "ä³ÛÙ³Ý³·ñ»ñÇ µ³óáõÙ: Üßí³Í ïáÕ»ñÇ ù³Ý³Ï - 2")
				Call ClickCmdButton(5, "Î³ï³ñ»É")
				Call MessageExists(2, "¶áñÍáÕáõÃÛáõÝÝ»ñÇ µ³ñ»Ñ³çáÕ ³í³ñï")
				Call ClickCmdButton(5, "Ok")
		else 
				Log.Error "Can't open frmAsMsgBox widow.", "", pmNormal, ErrorColor
		end if
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''Verify_Periodic_Actions'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Հաստատել պարբերական գործողությունների պայմանագիրը
Sub Verify_Periodic_Actions()
		Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_ToConfirm)
  Call ClickCmdButton(1, "Ð³ëï³ï»É")
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''Add_TaskTemplate''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Ավելացնել առաջադրանքի ձևանմուշ
'code - Կոդ դաշտի արժեք
'name - Անվանում դաշտի արժեք
'englName - Անգլերեն անվանում դաշտի արժեք
Sub Add_TaskTemplate(code, name, englName)
		Call ChangeWorkspace(c_Admin40)
		wTreeView.DblClickItem("²é³ç³¹ñ³ÝùÝ»ñ|²é³ç³¹ñ³ÝùÇ Ó¨³ÝÙáõßÝ»ñ")
		Call wMainForm.MainMenu.Click(c_Opers)
		Call wMainForm.PopupMenu.Click(c_Add)
		if p1.WaitVBObject("frmTreeNode", 3000).Exists then 
				Call Rekvizit_Fill("TreeNode", 1, "General", "lblCode", code)
				Call Rekvizit_Fill("TreeNode", 1, "General", "lblName", name)
				Call Rekvizit_Fill("TreeNode", 1, "General", "lblEName", englName)
				Call ClickCmdButton(8, "Î³ï³ñ»É")
		else 
				Log.Error "Can't open frmTreeNode widow.", "", pmNormal, ErrorColor
		end if
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''Add_SampleTemplate''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Տողի առկա լինելը ստուգող ֆունկցիա
'editTree - համապատասխան EditTree-ի անունը (օր.՝"frmEditTree", "frmEditTree_2")
'searchValue - փնտրվող արժեքը
Function Search_In_EditTree(editTree, searchValue)
		Dim treeView, itemExists, itemValue
		Dim status : status = false
		itemExists = true
		
  Set treeView = wMDIClient.VBObject(editTree).VBObject("TreeView")
		
		treeView.Keys("[Home]")
		itemValue = ""
		do while itemExists 
				if Trim(treeView.SelectedItem) = Trim(itemValue)  then
						itemExists = false
				  exit do
				end if
		  if Trim(treeView.SelectedItem) = Trim(searchValue)  then
		      status = true  
								exit do
		  else
				itemValue = Trim(treeView.SelectedItem)
				 treeView.Keys("[Down]")
		  end if 
		loop
		if not status then
				Log.Error "Can't find searched item.", "", pmNormal, ErrorColor
		end if
		Search_In_EditTree = status
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''Add_TemplateElement''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Ավելացնել ձևանմուշի տարր
'priority - Առաջնություն դաշտի արժեք
'taskElem - Առաջադրանքի տարր դաշտի արժեք
'comment - Մեկնաբանություն անվանում դաշտի արժեք
'editTree - համապատասխան EditTree-ի անունը (օր.՝"frmEditTree", "frmEditTree_2")
'searchValue - փնտրվող արժեքը
Sub Add_TemplateElement(priority, taskElem, comment, editTree, searchValue)
		Dim TreeView

  if Search_In_EditTree(editTree, searchValue) then		
		  Set TreeView = wMDIClient.VBObject(editTree).VBObject("TreeView")
				TreeView.Keys("[Enter]")
				if wMDIClient.WaitvbObject("frmPttel", 3000).Exists then
						Call wMainForm.MainMenu.Click(c_AllActions)
						Call wMainForm.PopupMenu.Click(c_Add)
						if p1.WaitvbObject("frmAsUstPar", 1000).Exists then 
								Call Rekvizit_Fill("Dialog", 1, "General", "PRIORITY", priority)
								Call Rekvizit_Fill("Dialog", 1, "General", "JOB", taskElem)
								Call Rekvizit_Fill("Dialog", 1, "General", "COMMENT", comment)
								Call ClickCmdButton(2, "Î³ï³ñ»É")
						else 
								Log.Error "Can't open frmTreeNode widow.", "", pmNormal, ErrorColor
						end if
				else
						Log.Error "Can't open frmPttel widow.", "", pmNormal, ErrorColor
				end if
		else 
				Log.Error "Can't find edit tree elenent", "", pmNormal, ErrorColor
		end if
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''Delete_TemplateElement'''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Ջնջել ձևանմուշի տարրը
'editTree - համապատասխան EditTree-ի անունը (օր.՝"frmEditTree", "frmEditTree_2")
'searchValue - փնտրվող արժեքը
Sub Delete_TemplateElement(editTree, searchValue)
		Call ChangeWorkspace(c_Admin40)
		wTreeView.DblClickItem("²é³ç³¹ñ³ÝùÝ»ñ|²é³ç³¹ñ³ÝùÇ Ó¨³ÝÙáõßÝ»ñ|")
  if Search_In_EditTree(editTree, searchValue) then		
				wMDIClient.VBObject(editTree).VBObject("TreeView").Keys("[Enter]")
				if wMDIClient.WaitvbObject("frmPttel", 3000).Exists then
						Call SearchAndDelete("frmPttel", 1, "PerAgrOp", "Ð³ëï³ï»ù ÷³ëï³ÃÕÃÇ çÝç»ÉÁ")
						BuiltIn.Delay(3000)
						wMDIClient.VBObject("frmPttel").Close
				else
						Log.Error "Can't open frmPttel widow.", "", pmNormal, ErrorColor
				end if
		Call wMainForm.MainMenu.Click(c_AllActions)
		Call wMainForm.PopupMenu.Click(c_Delete)
  BuiltIn.Delay(delay_small) 
		Call MessageExists(2, "Ð³ëï³ï»ù Ñ³Ý·áõÛóÇ çÝç»ÉÁ")
  Call ClickCmdButton(5, "²Ûá")
		wMDIClient.VBObject(editTree).Close
		else 
				Log.Error "Can't find edit tree elenent", "", pmNormal, ErrorColor
		end if
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''Fill_Task_Window''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Առաջադրանքի լրացման կլաս
Class Task 
		public executeStart
		public executeHour
		public taskDate
		public taskGroup
		public comment
		public continuePeriodicly
		public operDate
		public note
		public note2
		public note3
		public sendEMail
		public task_Num
		private sub Class_Initialize()
				executeStart = ""
				executeHour = ""
				taskDate = ""
				taskGroup = ""
				comment = ""
				continuePeriodicly = 0
				operDate = ""
				note = ""
				note2 = ""
				note3 = ""
				sendEmail = ""
				task_Num = ""
		end sub
End Class

Function New_Task()
		Set New_Task = new Task
End	Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''Fill_Task_Window''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Առաջադրանքի լրացման պրոցեդուրա
'task - Առաջադրանքի լրացման կլաս
Sub Fill_Task_Window(task)
		Dim TDBGridParams
		
		if p1.WaitVBObject("frmEditJob", 3000).Exists then 
				Call Rekvizit_Fill("EditJob", 1, "General", "LblStartTime", task.executeStart)
				Call Rekvizit_Fill("EditJob", 1, "General", "TDBStartTime", task.executeHour)
				Call Rekvizit_Fill("EditJob", 1, "General", "LblJobDate", task.taskDate)
				Call Rekvizit_Fill("EditJob", 1, "General", "LblJobGroup", task.taskGroup)
				Call Rekvizit_Fill("EditJob", 1, "General", "LblComment", task.comment)
				Call Rekvizit_Fill("EditJob", 1, "CheckBox", "lblRepeatPeriodically", task.continuePeriodicly)
				Set TDBGridParams = p1.VBObject("frmEditJob").VBObject("TDBGridParams")
				with TDBGridParams
						.Col = 1
		    ' Գործողության ամսաթիվ դաշտի լարցում
'		    .Row = 0
'		    .Keys(task.operDate & "[Down]")
						' Նշում դաշտի լարցում
		    .Row = 1
		    .Keys(task.note & "[Down]")
						' Նշում 2 դաշտի լարցում
		    .Row = 2
		    .Keys(task.note2 & "[Down]")
						' Նշում 3 դաշտի լարցում
		    .Row = 3
		    .Keys(task.note3 & "[Down]")
						' Ուղարկել էլ. նամակ դաշտի լարցում
		    .Row = 4
		    .Keys(task.sendEmail & "[Down]")
				end with
		else 
						Log.Error "Can't open frmEditJob widow.", "", pmNormal, ErrorColor
				end if
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''Add_Task''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Ավելացնել Առաջադրանք
'task - Առաջադրանքի լրացման կլաս
Sub Add_Task(task, PttelName)
		Dim tdbgView 
		
  Call ChangeWorkspace(c_Admin40)
		wTreeView.DblClickItem("²é³ç³¹ñ³ÝùÝ»ñ|²é³ç³¹ñ³ÝùÝ»ñ|")
		Call Rekvizit_Fill("Dialog", 1, "General", "SDATE", aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%d%m%y"))
		Call Rekvizit_Fill("Dialog", 1, "General", "EDATE", aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%d%m%y"))
		Call ClickCmdButton(2, "Î³ï³ñ»É")
		if wMDIClient.WaitvbObject("frmPttel", 3000).Exists then
				Set tdbgView = wMDIClient.VBObject(PttelName).VBObject("tdbgView")
				Call wMainForm.MainMenu.Click(c_Opers)
				Call wMainForm.PopupMenu.Click(c_Add)
				Call Fill_Task_Window(task)
				Call ClickCmdButton(9, "Î³ï³ñ»É ³ÝÙÇç³å»ë")
				BuiltIn.Delay(6000) 
				task.task_Num = Trim(tdbgView.Columns.Item(0).Value)
		else 
				Log.Error "Can't open frmPttel widow.", "", pmNormal, ErrorColor
		end if
End	Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''Delete_Task''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Ջնջել Առաջադրանքը
'task - Առաջադրանքի լրացման կլաս
Sub Delete_Task(task)
		Call ChangeWorkspace(c_Admin40)
		wTreeView.DblClickItem("²é³ç³¹ñ³ÝùÝ»ñ|²é³ç³¹ñ³ÝùÝ»ñ|")
		Call Rekvizit_Fill("Dialog", 1, "General", "SDATE", aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%d%m%y"))
		Call Rekvizit_Fill("Dialog", 1, "General", "EDATE", aqConvert.DateTimeToFormatStr(aqDateTime.Now(), "%d%m%y"))
		Call ClickCmdButton(2, "Î³ï³ñ»É")
		if wMDIClient.WaitvbObject("frmPttel", 3000).Exists then
				If SearchInPttel("frmPttel", 0, task) Then
        Call wMainForm.MainMenu.Click(c_Opers & "|" & c_Delete)
        BuiltIn.Delay(1000) 
        Call MessageExists(2, "Ð³ëï³ï»ù ³é³ç³¹ñ³ÝùÇ çÝç»ÉÁ")
        Call ClickCmdButton(5, "²Ûá") 
    Else
        Log.Error "Can Not find this row!",,,ErrorColor
    End If 
		else 
				Log.Error "Can't open frmPttel widow.", "", pmNormal, ErrorColor
		end if
End	Sub


' Մուտք"Պարբերական գործողությունների պայմանագրեր/Կոմունալ վճ. պայմանագրեր" թղթապանակ
Class CommunalPayDoc
        Public folderName
        Public dagrN
        Public wClient
        Public showClose
        Public wCompleted
        Public opendType
        Public wService
        Public wLocation
        Public wBranch
        Public  wDepart
        Private Sub Class_Initialize
              folderName = ""
              dagrN = ""
              wClient = ""
              showClose = False
              wCompleted = False
              opendType = False
              wService = ""
              wLocation = ""
              wBranch = ""
              wDepart = ""
        End Sub
End Class

Function New_CommunalPayDoc()
    Set New_CommunalPayDoc = NEW CommunalPayDoc      
End Function

' Մուտք"Կոմունալ վճ. պայմանագրեր" թղթապանակ - Fill
Sub Fill_CommunalPayDoc(CommunalPayDoc)
      
      wTreeView.DblClickItem(CommunalPayDoc.folderName)
      If Sys.Process("Asbank").WaitVBObject("frmAsUstPar", 2000).Exists Then
            
            ' "Պայմանագրի N" դաշտի լրացում
            Call Rekvizit_Fill("Dialog", 1, "General", "DAGRNUM", CommunalPayDoc.dagrN)
            ' "Հաճախորդ" դաշտի լրացում
            Call Rekvizit_Fill("Dialog", 1, "General", "DCLICODE", CommunalPayDoc.wClient)
            ' "Ցույց տալ փակված պայմանագրերը" դաշտի լրացում
            Call Rekvizit_Fill("Dialog", 1, "CheckBox", "DCLOSED", CommunalPayDoc.showClose)
            ' "Ցույց տալ միայն անավարտները ընթացիկ ամսում" դաշտի լրացում
            Call Rekvizit_Fill("Dialog", 1, "CheckBox", "DCOMPLETED", CommunalPayDoc.wCompleted)
            ' "Ցույց տալ բացված տեսքով" դաշտի լրացում
            Call Rekvizit_Fill("Dialog", 1, "CheckBox", "DADVANCED", CommunalPayDoc.opendType)
            ' "Ծառայություն" դաշտի լրացում
            Call Rekvizit_Fill("Dialog", 1, "General", "DSYS", CommunalPayDoc.wService)
            ' "Վայր" դաշտի լրացում
            Call Rekvizit_Fill("Dialog", 1, "General", "DLOCATION", CommunalPayDoc.wLocation)
            ' "Գրասենյակ" դաշտի լրացում
            Call Rekvizit_Fill("Dialog", 1, "General", "DACSBRANCH", CommunalPayDoc.wBranch)
            ' "Բաժին" դաշտի լրացում
            Call Rekvizit_Fill("Dialog", 1, "General", "DACSDEPART", CommunalPayDoc.wDepart)
            
            ' Սեղմել "Կատարել" կոճակը
            Call ClickCmdButton(2, "Î³ï³ñ»É")
            BuiltIn.Delay(2000)
            
      Else 
            Log.Error"Պարբերական կոմունալ վճարումների պայմանագրեր դիալոգը չի բացվել" ,,,ErrorColor
      End If
      
End Sub


' Պարբերական կոմունալ վճարումների դիտում 
Class ViewPayment
        Public stDate
        Public eDate
        Public wBranch
        Public wDepart
        Private Sub Class_Initialize
              stDate = ""
              eDate = ""
              wBranch = ""
              wDepart = ""
        End Sub
End Class

Function New_ViewPayment()
    Set New_ViewPayment = NEW ViewPayment      
End Function

' Պարբերական կոմունալ վճարումների ֆիլտրի լրացում
Sub Fill_ViewPayment(ViewPayment)
      
      ' Վճարումների դիտում
      Call wMainForm.MainMenu.Click(c_AllActions)    
      Call wMainForm.PopupMenu.Click(c_PaymentView)
      
      If Sys.Process("Asbank").WaitVBObject("frmAsUstPar", 2000).Exists Then
            
            ' "Ժամանակահատվածի սկիզբ" դաշտի լրացում
            Call Rekvizit_Fill("Dialog", 1, "General", "DSDATE", "^A[Del]"  & ViewPayment.stDate)
            ' "Ժամանակահատվածի ավարտ" դաշտի լրացում
            Call Rekvizit_Fill("Dialog", 1, "General", "DEDATE", "^A[Del]"  & ViewPayment.eDate)
            ' "Գրասենյակ" դաշտի լրացում
            Call Rekvizit_Fill("Dialog", 1, "General", "DACSBRANCH", ViewPayment.wBranch)
            ' "Բաժին" դաշտի լրացում
            Call Rekvizit_Fill("Dialog", 1, "General", "DACSDEPART", ViewPayment.wDepart)
            
            ' Սեղմել "Կատարել" կոճակը
            Call ClickCmdButton(2, "Î³ï³ñ»É")
            BuiltIn.Delay(2000)
            
      Else 
            Log.Error"Պարբերական կոմունալ վճարումներ դիալոգը չի բացվել" ,,,ErrorColor
      End If
      
End Sub


' Մուտք կոմունալ վճարումներ թղթապանակ
Class CommPaymentFolder
        Public folderName
        Public stDate
        Public eDate
        Public wType
        Public wLocation
        Public wCode
        Public wName
        Public wAddress
        Public wISN
        Public wBranch
        Public wDepart
        Private Sub Class_Initialize
              folderName = ""
              stDate = ""
              eDate = ""
              wType = ""
              wLocation = ""
              wCode = ""
              wName = ""
              wAddress = ""
              wISN = ""
              wBranch = ""
              wDepart = ""
        End Sub
End Class

Function New_CommPaymentFolder()
    Set New_CommPaymentFolder = NEW CommPaymentFolder      
End Function

' Կոմունալ վճարումներ ֆիլտրի լրացում
Sub Fill_CommPaymentFolder(CommPaymentFolder)
      
      wTreeView.DblClickItem(CommPaymentFolder.folderName)
      If Sys.Process("Asbank").WaitVBObject("frmAsUstPar", 2000).Exists Then
            
            ' "Ժամանակահատվածի սկիզբ" դաշտի լրացում
            Call Rekvizit_Fill("Dialog", 1, "General", "DSDATE", "^A[Del]"  & CommPaymentFolder.stDate)
            ' "Ժամանակահատվածի ավարտ" դաշտի լրացում
            Call Rekvizit_Fill("Dialog", 1, "General", "DEDATE", "^A[Del]"  & CommPaymentFolder.eDate)
            ' "Ծառայություն" դաշտի լրացում
            Call Rekvizit_Fill("Dialog", 1, "General", "TYPE", CommPaymentFolder.wType)
            ' "Վայրեր" դաշտի լրացում
            Call Rekvizit_Fill("Dialog", 1, "General", "LOC", CommPaymentFolder.wLocation)
            ' "Կոդ" դաշտի լրացում
            Call Rekvizit_Fill("Dialog", 1, "General", "DCODE", CommPaymentFolder.wCode)
            ' "Անվանում" դաշտի լրացում
            Call Rekvizit_Fill("Dialog", 1, "General", "DNAME", CommPaymentFolder.wName)
            ' "Հասցե" դաշտի լրացում
            Call Rekvizit_Fill("Dialog", 1, "General", "DADDRESS", CommPaymentFolder.wAddress)
            ' "Փաստաթղթի ISN" դաշտի լրացում
            Call Rekvizit_Fill("Dialog", 1, "General", "DISN", CommPaymentFolder.wISN)            
            ' "Գրասենյակ" դաշտի լրացում
            Call Rekvizit_Fill("Dialog", 1, "General", "DBRANCH", CommPaymentFolder.wBranch)
            ' "Բաժին" դաշտի լրացում
            Call Rekvizit_Fill("Dialog", 1, "General", "DDEPART", CommPaymentFolder.wDepart)
            
            ' Սեղմել "Կատարել" կոճակը
            Call ClickCmdButton(2, "Î³ï³ñ»É")
            BuiltIn.Delay(2000)
            
      Else 
            Log.Error"Կոմունալ վճարումներ դիալոգը չի բացվել" ,,,ErrorColor
      End If
      
End Sub



 ' Վճարման կատարում սխալի հաղորդագրությամբ
Sub PaymentWithError(opDate, checkDate, sendMail, savePath, folderName, fileName2, fileName1, param)

      ' Կատարել վճարում
      Call wMainForm.MainMenu.Click(c_AllActions)
		  Call wMainForm.PopupMenu.Click(c_MakePayment)
		
  		If p1.WaitVBObject("frmAsUstPar", 3000).Exists then
  				' Լրացնել Գործողության ամսաթիվ դաշտը 
  				Call Rekvizit_Fill("Dialog", 1, "General", "OPDATE", opDate)
  				' Լրացնել Ստուգել պարբ. և կատարման օրերը նշիչը 
  				Call Rekvizit_Fill("Dialog", 1, "CheckBox", "CHECKDATE", checkDate)
  				' Լրացնել Ուղարկել էլ. նամակ նշիչը
  				Call Rekvizit_Fill("Dialog", 1, "CheckBox", "SENDMAIL", sendMail)
  				' Սեղմել Կատարել կոճակը
  				Call ClickCmdButton(2, "Î³ï³ñ»É")

          BuiltIn.Delay(1000)
          If  wMDIClient.WaitVBObject("FrmSpr", 3000).Exists Then
      
           ' Հիշել քաղվածքը
           Call SaveDoc(savePath, folderName)

           ' Համեմատել ֆայլերը
           Call Compare_Files(fileName2, fileName1, param)
            
           BuiltIn.Delay(1000)
           wMDIClient.VBObject("FrmSpr").Close
            
           Else
                   Log.Error("Պարբերական կոմունալ վճարումների պայմանագիր վճարման համար սխալի պատուհանը չի բացվել"),,,ErrorColor
           End If    
      
  		Else
  				Log.Error "Կատարել վճարում գործողությունը չի իրականացել",,, ErrorColor
  		End If

End Sub

' Ավելացնել առաջադրանք Պարբերական կոմունալ վճարումների համար - Class
Class AddJobForPetComm
      Public folderName
      Public sDate
      Public eDate
      Public startTime
      Public wHour
      Public JobDate
      Public jobGroup
      Public wComment
      Public repeatPeriodicly
      Public operDate
      Public sendEmail
      Private Sub Class_Initialize()
            folderName = ""
            sDate = ""
            eDate = ""
            startTime = ""
            wHour = ""
            JobDate = ""
            jobGroup = ""
            wComment = ""
            repeatPeriodicly = False
            operDate = ""
            sendEmail = ""
      End Sub 
End Class

Function New_AddJobForPetComm()
    Set New_AddJobForPetComm = NEW AddJobForPetComm      
End Function

' Ավելացնել առաջադրանք Պարբերական կոմունալ վճարումների համար- fill
Sub Fill_AddJobForPetComm(CommJob)

      Dim tDBGridParams
      ' Թղթապանակի ուղղությունը
      wTreeView.DblClickItem(CommJob.folderName)
    
      If p1.VBObject("frmAsUstPar").Exists Then
          ' Լրացնել "Ժամանակահատվածի սկիզբ" դաշտը
  				Call Rekvizit_Fill("Dialog", 1, "General", "SDATE", CommJob.sDate)
          ' Լրացնել "Ժամանակահատվածի ավարտ" դաշտը
  				Call Rekvizit_Fill("Dialog", 1, "General", "EDATE", CommJob.eDate)
          ' Կատարել կոճակի սեղմում
          Call ClickCmdButton(2, "Î³ï³ñ»É")
      End If
      
      ' Կատարել Գործողություններ/Ավելացնել
      Call wMainForm.MainMenu.Click(c_Opers)
  		Call wMainForm.PopupMenu.Click(c_Add)
		
  		if p1.WaitVBObject("frmEditJob", 3000).Exists then 
          ' Լրացնել "Կատարման սկիզբ" դաշտը
  				Call Rekvizit_Fill("EditJob", 1, "General", "LblStartTime", CommJob.startTime)
          ' Լրացնել "Կատարման ժամ" դաշտը
  				Call Rekvizit_Fill("EditJob", 1, "General", "TDBStartTime", CommJob.wHour)
          ' Լրացնել "Առաջադրանքի ամսաթիվ" դաշտը
  				Call Rekvizit_Fill("EditJob", 1, "General", "LblJobDate", CommJob.JobDate)
          ' Լրացնել "Առաջադրանքների խումբ" դաշտը
  				Call Rekvizit_Fill("EditJob", 1, "General", "LblJobGroup", CommJob.jobGroup)
          ' Լրացնել "Մենկնաբանություն" դաշտը
  				Call Rekvizit_Fill("EditJob", 1, "General", "LblComment", CommJob.wComment)
          ' Լրացնել "Կրկնել Պարբերաբար" դաշտը
  				Call Rekvizit_Fill("EditJob", 1, "CheckBox", "lblRepeatPeriodically", CommJob.repeatPeriodicly)
        
  				Set tDBGridParams = p1.VBObject("frmEditJob").VBObject("TDBGridParams")
  				with tDBGridParams
           ' Գործողության ամսաթիվ դաշտի լարցում
  				.Col = 1
  		    .Row = 0
  		    .Keys(CommJob.operDate & "[Down]")
          ' Ուղարկել Էլ. նամակ դաշտի լրացում
  		    .Row = 1
  		    .Keys(CommJob.sendEmail & "[Down]")
  				end with
          
          ' Կատարել անմիջապես կոճակի սեղմում
          p1.VBObject("frmEditJob").VBObject("RunButtom").Click
          BuiltIn.Delay(5000)
  		else 
  						Log.Error("Նոր առաջադրանք դիալոգը չի բացվել"),,,ErrorColor
			end if
      
End	Sub