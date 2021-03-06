Option Explicit
'USEUNIT Library_Common
'USEUNIT Constants
'USEUNIT Subsystems_SQL_Library
'USEUNIT  Percentage_Calculation_Filter_Library
'USEUNIT SWIFT_International_Payorder_Library
'USEUNIT Library_Colour
'USEUNIT OLAP_Library
'USEUNIT Mortgage_Library

' Test Case ID 161139

' Տոկոսների Հաշվարկման ԱՇՏ-ում Ֆիլտրերի ստուգում (2)
Sub Percentage_Calculation_Report_2_Test()

      Dim fDATE, sDATE
      Dim folderDirect, balAcc, accMask, wCur, accType, acsBranch, acsDepart, acsType, accNote, accNote2
      Dim accNote3, wScale, showClosed, oldAccMask, newAccMask, showBal, showRem, wDate, selectView, exportExcel
      Dim PttelName, status, Path1, Path2, resultWorksheet, stDate, eDate, wUser
      Dim wState, docType, selectedView, expExcel
      Dim SortArr(3)
      
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
      Log.Message "--- Պայմանագրեր - 1 ---" ,,, DivideColor 
      ' "Պայմանագրեր" դիալոգի բացում և տվյալների լրացում
      folderDirect = "|îáÏáëÝ»ñÇ Ñ³ßí³ñÏÙ³Ý ²Þî|ä³ÛÙ³Ý³·ñ»ñ"
      showClosed = 1
      selectView = "PercView"
      exportExcel = "0"
      Call OpenContractsFolder(folderDirect, balAcc, accMask, wCur, accType, acsBranch, acsDepart, acsType, accNote, accNote2, _ 
                                                        accNote3, wScale, showClosed, oldAccMask, newAccMask, showBal, showRem, wDate, selectView, exportExcel)
      
      ' Սպասում է այնքան մինչև "կատարման ընթացքը" վերջանա 
      Call  WaitForExecutionProgress()
      
      ' Ստուգում է "Պայմանագրեր" թղթապանակը բացվել է թե ոչ
      PttelName = "frmPttel"
      status =  WaitForPttel(PttelName) 
      
      If  status Then

          ' Համեմատել պտտելի տողերի քանակը
          Call CheckPttel_RowCount("frmPttel", 18817)
      
          ' Արտահանել excel ֆայլ
          Path1 = Project.Path & "Stores\Reports\Percentage Calculation\Report_2\Actual\ContractsAct_1.xlsx"
          Call ExportToExcel("frmPttel",Path1)

          ' Համեմատել երկու EXCEL ֆայլերի բոլոր sheet-երը
          Path2 = Project.Path &  "Stores\Reports\Percentage Calculation\Report_2\Expected\ContractsExp_1.xlsx"
          resultWorksheet = Project.Path &  "Stores\Reports\Percentage Calculation\Report_2\Result\ContractsRes_1.xlsx"
          Call CompareTwoExcelFiles(Path1, Path2, resultWorksheet)
      
          ' Փակել Պայմանագրեր թղթապանակը
          Call Close_Pttel(PttelName)
      
          ' Փակել բոլոր excel ֆայլերը
          Call CloseAllExcelFiles()
      
      Else
            Log.Error "Պայմանագրեր թղթապանակը չի բացվել",,,ErrorColor
      End If
      
       '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
       Log.Message "--- Պայմանագրեր - 2 ---" ,,, DivideColor 
      ' "Պայմանագրեր" դիալոգի բացում և տվյալների լրացում
      balAcc = "3030201"
      accMask = "4670013"
      wCur = "000"
      accType = "01"
      acsBranch = "P00"
      acsDepart = "02"
      acsType = "01"
      accNote = ""
      accNote2 = "006"
      accNote3 = "001"
      wScale = "000011"
      oldAccMask = "46700"
      newAccMask = ""
      showBal = 1
      showRem = 1
      wDate = "101014"
      Call OpenContractsFolder(folderDirect, balAcc, accMask, wCur, accType, acsBranch, acsDepart, acsType, accNote, accNote2, _ 
                                                        accNote3, wScale, showClosed, oldAccMask, newAccMask, showBal, showRem, wDate, selectView, exportExcel)
      
      ' Ստուգում է "Պայմանագրեր" թղթապանակը բացվել է թե ոչ
      status =  WaitForPttel(PttelName) 
      
     If  status Then
      
          ' Համեմատել պտտելի տողերի քանակը
          Call CheckPttel_RowCount("frmPttel", 1)
      
          ' Արտահանել txt ֆայլ
          Path1 = Project.Path & "Stores\Reports\Percentage Calculation\Report_2\Actual\ContractsAct_2.txt"
          Call ExportToTXTFromPttel(PttelName,Path1)

          ' Համեմատել երկու txt ֆայլերը
          Path2 = Project.Path &  "Stores\Reports\Percentage Calculation\Report_2\Expected\ContractsExp_2.txt"
          Call Compare_Files(Path2, Path1, "")
      
          Call Close_Pttel("frmPttel")
      
      Else
           Log.Error "Պայմանագրեր թղթապանակը չի բացվել",,,ErrorColor
      End If
      
      '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
      Log.Message "--- Հաշվառված փաստաթղթեր ---" ,,, DivideColor 
      ' Մուտք Հաշվառված փաստաթղթեր դիալոգ և արժեքների լրացում
      folderDirect = "|îáÏáëÝ»ñÇ Ñ³ßí³ñÏÙ³Ý ²Þî|Ð³ßí³éí³Í ÷³ëï³ÃÕÃ»ñ"
      stDate = "010113"
      eDate = "010122"
      wUser = ""
      Call RegisteredDocuments(folderDirect, stDate, eDate, wUser)
      
      ' Սպասում է այնքան մինչև "կատարման ընթացքը" վերջանա 
      Call  WaitForExecutionProgress()
      
      ' Ստուգում է Հաշվառված փաստաթղթեր թղթապանակը բացվել է թե ոչ
      PttelName = "frmPttel"
      status =  WaitForPttel(PttelName) 
      
      If  status Then
      
          ' Դասավորել ըստ պայմանագրերի համարի
          SortArr(0) = "DOCNUM"
          SortArr(1) = "fCOM"
          SortArr(2)  = "USERID"
          Call columnSorting(SortArr, 3, "frmPttel")
      
          ' Համեմատել պտտելի տողերի քանակը
          Call CheckPttel_RowCount("frmPttel", 734)
      
          ' Արտահանել թղթապանակի տողերը բացված Pttel-ից
          Path1 = Project.Path & "Stores\Reports\Percentage Calculation\Report_2\Actual\Registered_DocsAct.xlsx"
      
          ' Արտահանել excel ֆայլ
          Call ExportToExcel("frmPttel",Path1)

          ' Համեմատել երկու EXCEL ֆայլերի բոլոր sheet-երը
          Path2 = Project.Path &  "Stores\Reports\Percentage Calculation\Report_2\Expected\Registered_DocsExp.xlsx"
          resultWorksheet = Project.Path &  "Stores\Reports\Percentage Calculation\Report_2\Result\Registered_DocsRes.xlsx"
          Call CompareTwoExcelFiles(Path1, Path2, resultWorksheet)
      
          ' Փակել Հաշվառված փաստաթղթեր թղթապանակը
          Call Close_Pttel(PttelName)
      
          ' Փակել բոլոր excel ֆայլերը
          Call CloseAllExcelFiles()
      
      Else
            Log.Error "Հաշվառված փաստաթղթեր թղթապանակը չի բացվել",,,ErrorColor
      End If
      
      '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
      Log.Message "--- Ստեղծված խմբային խմբագրումներ - 1 ---" ,,, DivideColor 
      ' Ստեղծված խմբային խմբագրումներ ֆիլտրի բացում և տվյալների լրացում
      folderDirect = "|îáÏáëÝ»ñÇ Ñ³ßí³ñÏÙ³Ý ²Þî|êï»ÕÍí³Í ËÙµ³ÛÇÝ ËÙµ³·ñáõÙÝ»ñ"
      stDate = "010113"
      eDate = "010122"
      selectedView = "CDGrpEds"
      expExcel = "0"
      Call CreatedGroupEdits(folderDirect, wState, stDate, eDate, docType, wUser, selectedView, expExcel)
      
      ' Ստուգում է Ստեղծված խմբային խմբագրումներ թղթապանակը բացվել է թե ոչ
      PttelName = "frmPttel"
      status =  WaitForPttel(PttelName) 
      
      If  status Then

          ' Համեմատել պտտելի տողերի քանակը
          Call CheckPttel_RowCount("frmPttel", 12)

          ' Կատարել բոլոր գործողությունները
          Call wMainForm.MainMenu.Click(c_Views)
          ' Ֆիլտրել ըստ ամսաթվի
          Call wMainForm.PopupMenu.Click(c_SortTimeColmn)
      
          ' Արտահանել թղթապանակի տողերը բացված Pttel-ից
          Path1 = Project.Path & "Stores\Reports\Percentage Calculation\Report_2\Actual\CreatedGroupEditsAct_1.xlsx"
      
          ' Արտահանել excel ֆայլ
          Call ExportToExcel("frmPttel",Path1)

          ' Համեմատել երկու EXCEL ֆայլերի բոլոր sheet-երը
          Path2 = Project.Path &  "Stores\Reports\Percentage Calculation\Report_2\Expected\CreatedGroupEditsExp_1.xlsx"
          resultWorksheet = Project.Path &  "Stores\Reports\Percentage Calculation\Report_2\CreatedGroupEditsRes_1.xlsx"
          Call CompareTwoExcelFiles(Path1, Path2, resultWorksheet)
      
          ' Փակել Ստեղծված խմբային խմբագրումներ  թղթապանակը
          Call Close_Pttel(PttelName)
      
          ' Փակել բոլոր excel ֆայլերը
          Call CloseAllExcelFiles()
      
      Else
            Log.Error "Ստեղծված խմբային խմբագրումներ թղթապանակը չի բացվել",,,ErrorColor
      End If
      
      '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
      Log.Message "--- Ստեղծված խմբային խմբագրումներ - 2 ---" ,,, DivideColor 
      ' Ստեղծված խմբային խմբագրումներ ֆիլտրի բացում և տվյալների լրացում
      wState = "2"
      stDate = "310316"
      eDate = "310316"
      docType = "GrpEdCli"
      wUser = "253"
      selectedView = "CDGrpEds\1"
      Call CreatedGroupEdits(folderDirect, wState, stDate, eDate, docType, wUser, selectedView, expExcel)
      
      ' Ստուգում է Ստեղծված խմբային խմբագրումներ թղթապանակը բացվել է թե ոչ
      status =  WaitForPttel(PttelName) 
      
      If  status Then
      
          ' Համեմատել պտտելի տողերի քանակը
          Call CheckPttel_RowCount("frmPttel", 10)
      
          ' Արտահանել թղթապանակի տողերը բացված Pttel-ից
          Path1 = Project.Path & "Stores\Reports\Percentage Calculation\Report_2\Actual\CreatedGroupEditsAct_2.txt"
      
          ' Արտահանել txt ֆայլ
          Call ExportToTXTFromPttel(PttelName,Path1)

          ' Համեմատել երկու txt ֆայլերը
          Path2 = Project.Path &  "Stores\Reports\Percentage Calculation\Report_2\Expected\CreatedGroupEditsExp_2.txt"
          Call Compare_Files(Path2, Path1, "")
      
          Call Close_Pttel("frmPttel")
      
      Else
           Log.Error "Ստեղծված խմբային խմբագրումներ թղթապանակը չի բացվել",,,ErrorColor
      End If
      
      ' Փակել ՀԾ-Բանկ ծրագիրը
      Call Close_AsBank()
End Sub