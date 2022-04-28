'USEUNIT SWIFT_International_Payorder_Library
'USEUNIT Library_Common
'USEUNIT OLAP_Library
'USEUNIT Constants
'USEUNIT Library_Colour
Option Explicit

Sub Report_16_Test()
    Dim DateStart, DateEnd, file1, file2, param
    
    DateStart = "20120101"
    DateEnd = "20240101"

'Մուտք գործել ՀԾ- Բանկ համակարգ ARMSOFT օգտագործողով
		Call Initialize_AsBankQA(DateStart, DateEnd)
		
		'Անցում կատարել "Ենթահամակարգեր" ԱՇՏ
  Call ChangeWorkspace(c_Subsystems)
  Call wTreeView.DblClickItem("|ºÝÃ³Ñ³Ù³Ï³ñ·»ñ (§ÐÌ¦)|Ð³ßí»ïíáõÃÛáõÝÝ»ñ,  Ù³ïÛ³ÝÝ»ñ|Î´ Ñ³ßí»ïíáõÃÛáõÝÝ»ñ|16 ØÇçµ³ÝÏ³ÛÇÝ å³Ñ³ÝçÝ»ñ ¨ å³ñï³íáñáõÃÛáõÝÝ»ñ")
    
		If p1.WaitVBObject("frmAsUstPar", 3000).Exists Then
				'Լրացնում է  "Ամսաթիվ" դաշտը
		  Call Rekvizit_Fill("Dialog", 1, "General", "Date", "151121")
				Call ClickCmdButton(2, "Î³ï³ñ»É")
		Else
				Log.Error "Can't open frmAsUstPar window", "", pmNormal, ErrorColor
				Exit Sub
		End If
		
		If wMDIClient.WaitVBObject("FrmSpr", 3000).Exists Then
				'Սեղմել "Հիշել որպես"
		  Call wMainForm.MainMenu.Click(c_SaveAs)
		  p1.Window("#32770", "ÐÇß»É áñå»ë", 1).Window("DUIViewWndClassName", "", 1).Window("DirectUIHWND", "", 1).Window("FloatNotifySink", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1).Keys(Project.Path & "Stores\CB\Actual\16.txt")
		  p1.Window("#32770", "ÐÇß»É áñå»ë", 1).Window("Button", "&Save", 1).Click()
    
		  'Համեմատել ֆայլերը 
		  file1 = Project.Path & "Stores\CB\Actual\16.txt"
		  file2 = Project.Path & "Stores\CB\Expected\Expected 16.txt"
		  Call Compare_Files(file1, file2, param)
				
				BuiltIn.Delay(1000)
				wMDIClient.VBObject("FrmSpr").Close
		Else 
				Log.Error "Can't open FrmSpr window", "", pmNormal, ErrorColor
		End If
		
		'Փակել ՀԾ-Բանկ համակարգը
  Call Close_AsBank()
End	Sub