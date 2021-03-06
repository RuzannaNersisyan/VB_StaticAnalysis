'USEUNIT Library_Common
'USEUNIT Library_Colour
'USEUNIT OLAP_Library
'USEUNIT Mortgage_Library
'USEUNIT SWIFT_International_Payorder_Library
'USEUNIT Constants
'USEUNIT Deposit_Contract_Library
Option Explicit

'Test Case Id - 162417

Sub Check_Filter_Window_2()
  
    Dim sDATE,fDATE
    Dim Path1, Path2, toolbar, Button, Exists
    Dim ConditionField,wTabStrip,Deposit_Attract,FilterWin,i
    Dim SaveAsWin, SortArr(1), ViewWin, ActualValue, ArmColumnNameArr,EngColumnNameArr
    SortArr(0) = "fKEY"

    'Համակարգ մուտք գործել ARMSOFT օգտագործողով
    sDATE = "20030101"
    fDATE = "20260101"
    Call Initialize_AsBank("bank_Report", sDATE, fDATE)
    Login("ARMSOFT")

    Call ChangeWorkspace(c_Deposits)
    
    Set Deposit_Attract = New_Deposit_Attracted()
        Deposit_Attract.ShowAccounts = 1

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''-- Ստուգել ֆիլտրել Պատուհանի "Կիրառել" կոճակը --''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Ստուգել ֆիլտրել Պատուհանի Կիրառել կոճակը --" ,,, DivideColor  
    
    Call GoToDeposit_Attracted(Deposit_Attract)
    
    Call ColumnSorting(SortArr, 1, "frmPttel")
    
    BuiltIn.Delay(2000)
    'Բացել Ֆիլտրել պատուհանը
    Set toolbar = wMainForm.VBObject("tbToolBar").Window("ToolbarWindow32")
    Call toolbar.ClickItem(125, False)
  
    BuiltIn.Delay(1000)
    Set FilterWin = p1.WaitVBObject("frmPttelFilter", 2000)
    'Ստուգել Ֆիլտրել պատուհանը բացվել է թե ոչ
    If FilterWin.Exists Then
    Set ConditionField = FilterWin.VBObject("FilterControl").VBObject("TDBGridFilter")
    
      Set Button = FilterWin.VBObject("FilterControl").VBObject("ToolbarFilterActions").Window("msvb_lib_toolbar")
      Call Button.ClickItem(101, False)
      Call Button.ClickItem(101, False)
      Call Button.ClickItem(102, False)
        
        ConditionField.Row = 0
        ConditionField.Col = 2
        ConditionField.Window("Edit", "", 1).Keys("A[Down][Down][Down]")
        ConditionField.Col = 3
        ConditionField.Window("Edit", "", 1).Keys("A[Down][Down][Down][Down][Down]")
        ConditionField.Col = 4
        ConditionField.Keys("50000000")
        
        'Անցում 2րդ տաբ
        Set wTabStrip = FilterWin.VBObject("TabStrip1")
			  wTabStrip.SelectedItem = wTabStrip.Tabs(2)
    
        FilterWin.VBObject("Frame3").VBObject("List4").Keys("[Down][Down][Down]")
        FilterWin.VBObject("Frame3").VBObject("Command9").Click
        FilterWin.VBObject("Frame3").VBObject("Command7").Click  
        FilterWin.VBObject("Frame3").VBObject("Command7").Click  
        FilterWin.VBObject("Frame3").VBObject("Command6").Click  
        FilterWin.VBObject("Frame3").VBObject("List4").Keys("[Down][Down][Down]")
        FilterWin.VBObject("Command4").Click   
        FilterWin.VBObject("Command4").Click   
        FilterWin.VBObject("Command4").Click   
        FilterWin.VBObject("Command2").Click   
        
        'Սեղմել "Կիրառել" կոճակը
        FilterWin.VBObject("Command1").Click   
        BuiltIn.Delay(1000)
        
        Log.Message "-- Ստուգել ֆիլտրել Պատուհանի Դադարեցնել կոճակը --" ,,, DivideColor  
        FilterWin.VBObject("CancelButton").Click   
    Else
        Log.Error "Ֆիլտրել պատուհանը չի բացվել",,,ErrorColor   
    End if
    
    Call CheckPttel_RowCount("frmPttel", 5)
    Path1 = Project.Path & "Stores\FilterCheck\Actual\Check_FilterActual_5.txt"
    Path2 = Project.Path & "Stores\FilterCheck\Expected\Check_FilterExpected_5.txt"
    
    'Արտահանել և Ð³Ù»Ù³ï»É »ñÏáõ TXT ý³ÛÉ»ñ
    Call ExportToTXTFromPttel("frmPttel",Path1)
    Call Compare_Files(Path2, Path1, "")
    Call Close_Pttel("frmPttel")
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''-- Ստուգել ֆիլտրել պատուհանի Սյուներ <Tab>-ի գործողությունները  --''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Ստուգել ֆիլտրել պատուհանի Սյուներ <Tab>-ի գործողությունները --" ,,, DivideColor  

    Call GoToDeposit_Attracted(Deposit_Attract) 
    
    BuiltIn.Delay(2000)
    'Բացել Ֆիլտրել պատուհանը
    Call wMainForm.MainMenu.Click(c_Opers)
    Call wMainForm.PopupMenu.Click( c_Folder & "|" & c_Filter)
           
    BuiltIn.Delay(1000)
    Set FilterWin = p1.WaitVBObject("frmPttelFilter", 2000)
    'Ստուգել Ֆիլտրել պատուհանը բացվել է թե ոչ
    If FilterWin.Exists Then
        
    Set ConditionField = FilterWin.VBObject("FilterControl").VBObject("TDBGridFilter")
        ConditionField.Row = 0
        ConditionField.Col = 2
        ConditionField.Window("Edit", "", 1).Keys("A[Down][Down][Down]")
        ConditionField.Col = 4
        ConditionField.Keys("80000")
                
        'Սեղմել "Կատարել" կոճակը
        FilterWin.VBObject("Command5").Click       
    Else
        Log.Error "Ֆիլտրել պատուհանը չի բացվել",,,ErrorColor   
    End if
    
    Call ColumnSorting(SortArr, 1, "frmPttel")
    Call CheckPttel_RowCount("frmPttel", 3)
    Path1 = Project.Path & "Stores\FilterCheck\Actual\Check_FilterActual_6.txt"
    Path2 = Project.Path & "Stores\FilterCheck\Expected\Check_FilterExpected_6.txt"
    
    'Արտահանել և Ð³Ù»Ù³ï»É »ñÏáõ TXT ý³ÛÉ»ñ
    Call ExportToTXTFromPttel("frmPttel",Path1)
    Call Compare_Files(Path2, Path1, "")
    Call Close_Pttel("frmPttel")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''-- Ստուգել ֆիլտրել պատուհանի Արժեք դաշտում բացված սյուն և սյան անվանում սյուների արժեքները --''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Ստուգել ֆիլտրել պատուհանի Արժեք դաշտում բացված սյուն և սյան անվանում սյուների արժեքները --" ,,, DivideColor  
    
    Call GoToDeposit_Attracted(Deposit_Attract)
    
    BuiltIn.Delay(2000)
    'Բացել Ֆիլտրել պատուհանը
    Call wMainForm.MainMenu.Click(c_Opers)
    Call wMainForm.PopupMenu.Click( c_Folder & "|" & c_Filter)  
    
    EngColumnNameArr = Array("FKEY","FCOM","FCURRENCY","FAGRREM","FPERREM","FDATE","FAGRDATE","FDATECLOSE","FCLICODE","AGRACC","PERACC","FNOTE","FNOTE2","FNOTE3","ACSBRANCH","ACSDEPART","ACSTYPE","FPPRCODE")  
    ArmColumnNameArr = Array("ä³ÛÙ³Ý³·ñÇ N","²Ýí³ÝáõÙ","²ñÅ.","ØÝ³óáñ¹","Ð³ßí.% ÙÝ³óáñ¹","²Ùë³ÃÇí","Ø³ñÙ³Ý Å³ÙÏ»ï","ö³ÏÙ³Ý ³Ùë³ÃÇí","Ð³×³Ëáñ¹","ä³ÛÙ³Ý³·ñÇ Ñ³ßÇí","Îáõï³ÏÙ³Ý Ñ³ßÇí","ÜßáõÙ","ÜßáõÙ 2","ÜßáõÙ 3","¶ñ³ë»ÝÛ³Ï","´³ÅÇÝ","Ð³ë³Ý-Ý ïÇå","ä³ÛÙ.ÃÕÃ³ÛÇÝ N")  
    
    Set FilterWin = p1.WaitVBObject("frmPttelFilter", 2000)
    Set ConditionField = FilterWin.VBObject("FilterControl").VBObject("TDBGridFilter")
    
        ConditionField.Row = 0
        ConditionField.Col = 3
        ConditionField.Window("Edit", "", 1).Keys("[PageDown][Down][Down][Down][Down][Down][Down][Down][Down][Down][Down][Down][Enter]")
        ConditionField.Col = 4
        ConditionField.Window("Edit", "", 1).Keys("[PageDown]")
        
        For i = 0 to p1.VBObject("frmModalBrowser").VBObject("tdbgView").ApproxCount - 1
        
            'Ստուգել "Սյուն" սյան արժեքները
            ActualValue = p1.VBObject("frmModalBrowser").VBObject("tdbgView").Columns.Item(1)
            Call Compare_Two_Values(EngColumnNameArr(i),ActualValue,EngColumnNameArr(i))
            
            'Ստուգել "Սյան անվանում" սյան արժեքները
            ActualValue = p1.VBObject("frmModalBrowser").VBObject("tdbgView").Columns.Item(2)
            Call Compare_Two_Values(ArmColumnNameArr(i),ActualValue,ArmColumnNameArr(i))
            
            ConditionField.Window("Edit", "", 1).Keys("[Down]")    
        Next
        ConditionField.Window("Edit", "", 1).Keys("[Enter]")  
        FilterWin.VBObject("CancelButton").Click

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''-- Ստուգել ֆիլտրել պատուհանի Սյուներ <Tab>-ում "=[Սյուն]" նշանով  --''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Log.Message "-- Ստուգել ֆիլտրել պատուհանը Սյուներ <Tab>-ում =[Սյուն] նշանով--" ,,, DivideColor         

    BuiltIn.Delay(2000)
    'Բացել Ֆիլտրել պատուհանը
    Call wMainForm.MainMenu.Click(c_Opers)
    Call wMainForm.PopupMenu.Click( c_Folder & "|" & c_Filter)       
    
    Set FilterWin = p1.WaitVBObject("frmPttelFilter", 2000)
    Set ConditionField = FilterWin.VBObject("FilterControl").VBObject("TDBGridFilter")
    
    If FilterWin.Exists Then
        ConditionField.Row = 0
        ConditionField.Col = 2
        ConditionField.Window("Edit", "", 1).Keys("[PageDown][Down][Down][Down][Down][Down][Enter]")
        ConditionField.Col = 3
        ConditionField.Window("Edit", "", 1).Keys("[PageDown][Down][Down][Down][Down][Down][Down][Down][Down][Down][Down][Down][Enter]")
        ConditionField.Col = 4
        ConditionField.Window("Edit", "", 1).Keys("[PageDown][Down][Down][Down][Down][Down][Down][Down][Enter]")
    
        'Սեղմել "Կատարել" կոճակը
        FilterWin.VBObject("Command5").Click       
    Else
        Log.Error "Ֆիլտրել պատուհանը չի բացվել",,,ErrorColor   
    End if
        
    Call ColumnSorting(SortArr, 1, "frmPttel")
    Call CheckPttel_RowCount("frmPttel", 3)
    Path1 = Project.Path & "Stores\FilterCheck\Actual\Check_FilterActual_7.txt"
    Path2 = Project.Path & "Stores\FilterCheck\Expected\Check_FilterExpected_7.txt"
    
    'Արտահանել և Ð³Ù»Ù³ï»É »ñÏáõ TXT ý³ÛÉ»ñ
    Call ExportToTXTFromPttel("frmPttel",Path1)
    Call Compare_Files(Path2, Path1, "")
    Call Close_Pttel("frmPttel")
    
    Call Close_AsBank()   
End Sub    






