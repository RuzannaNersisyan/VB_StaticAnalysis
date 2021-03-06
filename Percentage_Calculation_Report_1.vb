Option Explicit
'USEUNIT Library_Common
'USEUNIT Constants
'USEUNIT Subsystems_SQL_Library
'USEUNIT  Percentage_Calculation_Filter_Library
'USEUNIT SWIFT_International_Payorder_Library
'USEUNIT Library_Colour
'USEUNIT OLAP_Library
'USEUNIT Mortgage_Library
'USEUNIT Debit_Dept_Filter_Library

' Test Case ID 161128

' Տոկոսների Հաշվարկման ԱՇՏ-ում  Ֆիլտրերի ստուգում (1)
Sub Percentage_Calculation_Report_1_Test()

      Dim fDATE, sDATE
      Dim folderDirect, stDate, eDate, wCur, wUser, docType, paySysin, paySysOut, payNotes, acsBranch, acsDepart, selectView, exportExcel
      Dim PttelName, status, Path1, Path2, resultWorksheet, passTax, wName, wState, exists
      Dim SortArr(3), i, wFrame3, FilterWin, wTabStrip
      
      fDATE = "20250101"
      sDATE = "20030101"
      Call Initialize_AsBank("bank_Report", sDATE, fDATE)
      
      ' Մուտք գործել համակարգ ARMSOFT օգտագործողով 
      Call Create_Connection()
      Login("ARMSOFT")
      
      ' Մուտք Տոկոսների հաշվարկման ԱՇՏ
      Call ChangeWorkspace(c_PercentCalc)
      
      ' Դրույթներից, Տնտեսել հիշողոթյունը սկսած (Տողերի քանակ) դաշտի արժեքի փոփոխում
      Call  SaveRAM_RowsLimit("10")
      
      '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
      Log.Message "--- Աշխատանքային փաստաթղթեր թղթապանակ ---" ,,, DivideColor 
      ' Մուտք աշխատանքային փաստաթղթեր դիալոգ և արժեքների լրացում
      folderDirect = "|îáÏáëÝ»ñÇ Ñ³ßí³ñÏÙ³Ý ²Þî|²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ"
      stDate = "010113"
      eDate = "010122"
      wCur = ""
      wUser = ""
      docType = ""
      paySysin = ""
      paySysOut = ""
      payNotes = ""
      acsBranch = ""
      acsDepart = "" 
      selectView = "Oper"
      exportExcel = "0"
      Call WorkingDocsFilter(folderDirect, stDate, eDate, wCur, wUser, docType, paySysin, paySysOut, payNotes, acsBranch, acsDepart, selectView, exportExcel)
      
      ' Ստուգում է Աշխատանքային փաստաթղթեր թղթապանակը բացվել է թե ոչ
      PttelName = "frmPttel"
      status =  WaitForPttel(PttelName) 
      
      If  status Then
      
          ' Համեմատել պտտելի տողերի քանակը
          Call CheckPttel_RowCount("frmPttel", 5)
      
          ' Դասավորել ըստ պայմանագրերի համարի
          SortArr(0) = "DOCNUM"
          Call FastColumnSorting(SortArr, 1, "frmPttel")
      
          ' Արտահանել թղթապանակի տողերը բացված Pttel-ից
          Path1 = Project.Path & "Stores\Reports\Percentage Calculation\Report_1\Actual\WorkingDocsFolderAct_1.xlsx"
      
          ' Արտահանել excel ֆայլ
          Call ExportToExcel("frmPttel",Path1)

          ' Համեմատել երկու EXCEL ֆայլերի բոլոր sheet-երը
          Path2 = Project.Path &  "Stores\Reports\Percentage Calculation\Report_1\Expected\WorkingDocsFolderExp_1.xlsx"
          resultWorksheet = Project.Path &  "Stores\Reports\Percentage Calculation\Report_1\Result\WorkingDocsFolderRes_1.xlsx"
          Call CompareTwoExcelFiles(Path1, Path2, resultWorksheet)
      
          ' Փակել Աշխատանքային փաստաթղթեր թղթապանակը
          Call Close_Pttel(PttelName)
      
          ' Փակել բոլոր excel ֆայլերը
          Call CloseAllExcelFiles()
      
      Else
            Log.Error "Աշխատանքային փաստաթղթեր թղթապանակը չի բացվել",,,ErrorColor
      End If
      
       '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
       Log.Message "--- Աշխատանքային փաստաթղթեր թղթապանակ ---" ,,, DivideColor 
      ' Մուտք աշխատանքային փաստաթղթեր թղթապանակ և արժեքների լրացում
      wCur = "000"
      wUser = "253"
      docType = "CrPayOrd"
      paySysin = "Ð"
      paySysOut = "1"
      acsBranch = "P00"
      acsDepart = "08" 
      exportExcel = "1"
      Call WorkingDocsFilter(folderDirect, stDate, eDate, wCur, wUser, docType, paySysin, paySysOut, payNotes, acsBranch, acsDepart, selectView, exportExcel)

      ' Արտահանել թղթապանակի տողերը բացված Pttel-ից
      Path1 = Project.Path & "Stores\Reports\Percentage Calculation\Report_1\Actual\WorkingDocsFolderAct_2.xlsx"
      Path2 = Project.Path &  "Stores\Reports\Percentage Calculation\Report_1\Expected\WorkingDocsFolderExp_2.xlsx"
      resultWorksheet = Project.Path &  "Stores\Reports\Percentage Calculation\Report_1\Result\WorkingDocsFolderRes_2.xlsx"
            
      exists = aqFile.Exists(Path1)
      If exists Then
          aqFileSystem.DeleteFile(Path1)
      End If
      
      'Արտահանել և Ð³Ù»Ù³ï»É »ñÏáõ EXCEL ý³ÛÉ»ñ
      BuiltIn.Delay(3000)
      
      If Sys.Process("EXCEL").Exists Then
      
          Sys.Process("EXCEL").Window("XLMAIN", "* - Excel", 1).Window("XLDESK", "", 1).Window("EXCEL7", "*", 1).Keys("[F12]")
          Sys.Process("EXCEL").Window("#32770", "Save As", 1).Keys(Path1 & "[Enter]")
          
      Else 
          Log.Error "Excel does not Open!" ,,,ErrorColor
      End If 
      
      ' Համեմատել երկու EXCEL ֆայլերի բոլոր sheet-երը
      Call CompareTwoExcelFiles(Path1, Path2, resultWorksheet)  
      'Փակել բոլոր excel ֆայլերը
      Call CloseAllExcelFiles()
      
       '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
       Log.Message "--- Նոր ստեղծված քարտ փաստաթղթեր ---" ,,, DivideColor  
      ' Մուտք նոր ստեղծված քարտ փաստաթղթեր դիալոգ և արժեքների լրացում
      folderDirect = "|îáÏáëÝ»ñÇ Ñ³ßí³ñÏÙ³Ý ²Þî|Üáñ ëï»ÕÍí³Í ù³ñï ÷³ëï³ÃÕÃ»ñ"
      docType = "Cli"
      acsBranch = "P06"
      acsDepart = "05"
      passTax = "AN0298941"
      wName = "Ð³ÏáµÛ³Ý Ð³Ïáµ ê»ñÅÇÏÇ"
      wUser = "251"
      Call NewCreatedCardDoc(folderDirect, docType, acsBranch, acsDepart, passTax, wName, wUser)
      
      ' Ստուգում է նոր ստեղծված քարտ փաստաթղթեր թղթապանակը բացվել է թե ոչ
      status =  WaitForPttel(PttelName) 
      
      If  status Then
      
          ' Համեմատել պտտելի տողերի քանակը
          Call CheckPttel_RowCount("frmPttel", 1)
      
          ' Արտահանել թղթապանակի տողերը բացված Pttel-ից
          Path1 = Project.Path & "Stores\Reports\Percentage Calculation\Report_1\Actual\New_Card_DocAct.txt"
      
          ' Արտահանել txt ֆայլ
          Call ExportToTXTFromPttel(PttelName,Path1)

          ' Համեմատել երկու txt ֆայլերը
          Path2 = Project.Path &  "Stores\Reports\Percentage Calculation\Report_1\Expected\New_Card_DocExp.txt"
          Call Compare_Files(Path2, Path1, "")
      
          Call Close_Pttel("frmPttel")
      
      Else
           Log.Error "Նոր ստեղծված քարտ փաստաթղթեր թղթապանակը չի բացվել",,,ErrorColor
      End If
      
      '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
      
      Log.Message "--- Քարտ փաստաթղթերի փոփոխման պատմություն - ստեղծել ֆիլտր ---" ,,, DivideColor  
      ' "Քարտ փաստաթղթերի փոփոխման պատմություն" դիալոգի բացում և տվյալների լրացում
      folderDirect = "|îáÏáëÝ»ñÇ Ñ³ßí³ñÏÙ³Ý ²Þî|ø³ñï ÷³ëï³ÃÕÃ»ñÇ ÷á÷áËÙ³Ý Ñ³Ûï»ñÇ å³ïÙáõÃÛáõÝ"
      docType = ""
      wState = ""
      stDate = "010113"
      eDate = "010122"
      wUser = ""
      Call HistoryOfChangeRequest(folderDirect, docType, wState, stDate, eDate, wUser)
      
      ' Սպասում է այնքան մինչև "կատարման ընթացքը" վերջանա 
      Call  WaitForExecutionProgress()
      
      ' Ստուգում է "Քարտ փաստաթղթերի փոփոխման պատմություն" թղթապանակը բացվել է թե ոչ
      PttelName = "frmPttel"
      status =  WaitForPttel(PttelName) 
      
      If  status Then
      
         ' Համեմատել պտտելի տողերի քանակը
         Call CheckPttel_RowCount("frmPttel", 3371)
         
         ' Բացել Ֆիլտրել պատուհանը
         Call wMainForm.MainMenu.Click(c_Opers)
         Call wMainForm.PopupMenu.Click( c_Folder & "|" & c_Filter)
         
         Set FilterWin = p1.WaitVBObject("frmPttelFilter", 2000)
         ' Ստուգել Ֆիլտրել պատուհանը բացվել է թե ոչ
         If FilterWin.Exists Then
    
              ' Անցում 2րդ թաբ
              Set wTabStrip = FilterWin.VBObject("TabStrip1")
     			    wTabStrip.SelectedItem = wTabStrip.Tabs(2)
                  
              Set wFrame3= Sys.Process("Asbank").VBObject("frmPttelFilter").VBObject("Frame3")
              ' "Ժամանակ" սյունը տեղափոխել սկիզբ  
              i = 0
              Call wFrame3.VBObject("List4").FocusItem("ö³ëï³ÃÕÃÇ ï»ë³Ï")
              For i = 0 To 20
                   wFrame3.VBObject("List4").Keys("[Down]")
              Next
                      
              i = 0
              For i = 0 To 20
                   wFrame3.VBObject("Command6").Click
              Next
               
              ' Սեղմել "Կատարել" կոճակը
              FilterWin.VBObject("Command5").Click  
               
         Else
             Log.Error "Ֆիլտրել պատուհանը չի բացվել",,,ErrorColor   
         End if
    
         BuiltIn.Delay(2000)
         ' Ֆիլտրել ըստ "Հաճախորդ" , "ժամանակ", "Հաշվի", "Անձնագիր/ՀՎՀՀ"  սյուների
         wMDIClient.VBObject("frmPttel").Keys("[Hold]" & "^!" & (5))
         BuiltIn.Delay(500)
				 wMDIClient.VBObject("frmPttel").Keys("[Hold]" & "^!" & (1))
         BuiltIn.Delay(500)
         wMDIClient.VBObject("frmPttel").Keys("[Hold]" & "^!" & (6))
         BuiltIn.Delay(500)
         wMDIClient.VBObject("frmPttel").Keys("[Hold]" & "^!" & (8))

         ' Արտահանել excel ֆայլ  բացված Pttel-ից
         Path1 = Project.Path & "Stores\Reports\Percentage Calculation\Report_1\Actual\HistoryChangeReqAct.xlsx"
         Call ExportToExcel("frmPttel",Path1)
          
         ' Համեմատել երկու EXCEL ֆայլերի բոլոր sheet-երը
         Path2 = Project.Path &  "Stores\Reports\Percentage Calculation\Report_1\Expected\HistoryChangeReqExp.xlsx"
         resultWorksheet = Project.Path &  "Stores\Reports\Percentage Calculation\Report_1\Result\HistoryChangeReqRes.xlsx"
         Call CompareTwoExcelFiles(Path1, Path2, resultWorksheet)
      
         ' Փակել Քարտ փաստաթղթերի փոփոխման պատմություն թղթապանակը
         Call Close_Pttel(PttelName)
      
         ' Փակել բոլոր excel ֆայլերը
         Call CloseAllExcelFiles()
      
      Else
            Log.Error "Քարտ փաստաթղթերի փոփոխման պատմություն թղթապանակը չի բացվել",,,ErrorColor
      End If
      
      '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
      Log.Message "--- Քարտ փաստաթղթերի փոփոխման պատմություն -2 ---" ,,, DivideColor  
      ' "Քարտ փաստաթղթերի փոփոխման պատմություն" դիալոգի բացում և տվյալների լրացում
      docType = "NBAcc"
      wState = "10"
      wUser = "110"
      Call HistoryOfChangeRequest(folderDirect, docType, wState, stDate, eDate, wUser)
      
      ' Ստուգում է "Քարտ փաստաթղթերի փոփոխման պատմություն" թղթապանակը բացվել է թե ոչ
      status =  WaitForPttel(PttelName) 
      
      If  status Then
      
          ' Համեմատել պտտելի տողերի քանակը
          Call CheckPttel_RowCount("frmPttel", 0)
      
          BuiltIn.Delay(1000)
          ' Փակել Պտտելը
          Call Close_Pttel("frmPttel")
          
      Else
            Log.Error "Քարտ փաստաթղթերի փոփոխման պատմություն թղթապանակը չի բացվել",,,ErrorColor
      End If
      
      ' Փակել ՀԾ-Բանկ ծրագիրը
      Call Close_AsBank()
      
End Sub