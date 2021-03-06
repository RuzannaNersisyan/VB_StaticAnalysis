Option Explicit
'USEUNIT Library_Common
'USEUNIT Constants
'USEUNIT Subsystems_SQL_Library
'USEUNIT  Main_Accountant_Filter_Library
'USEUNIT SWIFT_International_Payorder_Library
'USEUNIT Library_Colour
'USEUNIT OLAP_Library
'USEUNIT Mortgage_Library
' Test Case ID 161729

' Գլխավոր հաշվապահի ԱՇՏ-ում ֆիլտրերի ստուգում (3)
Sub Main_Accountant_Report_3_Test()

      Dim fDATE, sDATE
      Dim folderDirect, stDate, eDate, wUser, wCue, notCur, docType, paySysIn, paySysOut, _
              showLongNames, acsBranch, acsDepart, selectedView, exportExcel, status, state
      Dim PttelName, Path1, Path2, resultWorksheet
      Dim coaNum, balAcc, accMask, wCur, operType, showPrc, showRel, showRst, _
              showAccNames, showPayres, showCrDate
      Dim SortArr(2)
      
      fDATE = "20220101"
      sDATE = "20030101"
      Call Initialize_AsBank("bank_Report", sDATE, fDATE)
      
      ' Մուտք գործել համակարգ ARMSOFT օգտագործողով 
      Call Create_Connection()
      Login("ARMSOFT")
      
      ' Մուտք Գլխավոր հաշվապահի ԱՇՏ
      Call ChangeWorkspace(c_ChiefAcc)
      
      ' Դրույթներից, Տնտեսել հիշողոթյունը սկսած (Տողերի քանակ) դաշտի արժեքի փոփոխում
      Call  SaveRAM_RowsLimit("10")
      
      '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------       
      Log.Message "--- Հաշվառման ենթակա - 1 ---" ,,, DivideColor     
      ' Մուտք Հաշվառման ենթակա թղթապանակ
      folderDirect = "|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|êï³óí³Í Ñ³ÝÓÝ³ñ³ñ³·ñ»ñ|Ð³ßí³éÙ³Ý »ÝÃ³Ï³"
      stDate = "010103"
      eDate = "010122"
      wUser = ""
      wCue= ""
      notCur = 0
      docType = ""
      paySysIn = ""
      paySysOut = ""
      showLongNames = 1
      acsBranch = ""
      acsDepart = ""
      selectedView = "Payins"
      exportExcel = "0"
      status = True
      Call OpenSubjectToRegistrationFolder(folderDirect, stDate, eDate, wUser, wCue, notCur, docType, paySysIn, paySysOut, _
                                                             showLongNames, acsBranch, acsDepart, selectedView, exportExcel, status)
        
      ' Ստուգում է Ստացված հանձնարարագրեր թղթապանակը բացվել է թե ոչ
      PttelName = "frmPttel"
      status =  WaitForPttel(PttelName)
      
      If  status Then
      
            ' Համեմատել պտտելի տողերի քանակը
            Call CheckPttel_RowCount("frmPttel", 28)

            ' Դասավորել ըստ N , Գումար սուների
      			Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").Keys("[Hold]" & "^" & (3))

            SortArr(0) = "SUMMA"
            Call FastColumnSorting(SortArr, 1, "frmPttel")
      
            ' Արտահանել թղթապանակի տողերը բացված Pttel-ից
            Path1 = Project.Path & "Stores\Reports\Main Accountant\Report_3\Actual\SubjectToRegAct_1.xlsx"
      
            ' Արտահանել excel ֆայլ
            Call ExportToExcel("frmPttel",Path1)

            ' Համեմատել երկու EXCEL ֆայլերի բոլոր sheet-երը
            Path2 = Project.Path &  "Stores\Reports\Main Accountant\Report_3\Expected\SubjectToRegExp_1.xlsx"
            resultWorksheet = Project.Path &  "Stores\Reports\Main Accountant\Report_3\Result\SubjectToRegRes_1.xlsx"
            Call CompareTwoExcelFiles(Path1, Path2, resultWorksheet)
      
            ' Փակել Ստացված հանձնարարագրեր թղթապանակը
            Call Close_Pttel(PttelName)
      
      ' Փակել բոլոր excel ֆայլերը
      Call CloseAllExcelFiles()
      
      Else
            Log.Error("Ստացված հանձնարարագրեր թղթապանակը չի բացվել")
      End If
      
      '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------   
      Log.Message "--- Հաշվառման ենթակա - 2 ---" ,,, DivideColor       
      ' Մուտք Հաշվառման ենթակա թղթապանակ
      stDate = "010104"
      eDate = "010115"
      wUser = "19"
      wCue= "000"
      notCur = 0
      docType = "TransPay"
      paySysIn = "5"
      paySysOut = ""
      showLongNames = 1
      acsBranch = "P00"
      acsDepart = "08"
      Call OpenSubjectToRegistrationFolder(folderDirect, stDate, eDate, wUser, wCue, notCur, docType, paySysIn, paySysOut, _
                                                              showLongNames, acsBranch, acsDepart, selectedView, exportExcel, status)
        
      ' Ստուգում է Ստացված հանձնարարագրեր թղթապանակը բացվել է թե ոչ
      status =  WaitForPttel(PttelName)
      
      If  status Then
      
            ' Համեմատել պտտելի տողերի քանակը
            Call CheckPttel_RowCount("frmPttel", 8)
      
            ' Դասավորել ըստ N , Գումար սուների
      			Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").Keys("[Hold]" & "^" & (3))

            SortArr(0) = "SUMMA"
            Call FastColumnSorting(SortArr, 1, "frmPttel")
      
            ' Արտահանել թղթապանակի տողերը բացված Pttel-ից
            Path1 = Project.Path & "Stores\Reports\Main Accountant\Report_3\Actual\SubjectToRegAct_2.xlsx"
      
            ' Արտահանել excel ֆայլ
            Call ExportToExcel("frmPttel",Path1)

            ' Համեմատել երկու EXCEL ֆայլերի բոլոր sheet-երը
            Path2 = Project.Path &  "Stores\Reports\Main Accountant\Report_3\Expected\SubjectToRegExp_2.xlsx"
            resultWorksheet = Project.Path &  "Stores\Reports\Main Accountant\Report_3\Result\SubjectToRegRes_2.xlsx"
            Call CompareTwoExcelFiles(Path1, Path2, resultWorksheet)
      
            ' Փակել Ստացված հանձնարարագրեր թղթապանակը
            Call Close_Pttel(PttelName)
      
            ' Փակել բոլոր excel ֆայլերը
            Call CloseAllExcelFiles()
      
      Else
            Log.Error("Ստացված հանձնարարագրեր թղթապանակը չի բացվել")
      End If
      
      '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------       
      Log.Message "--- Տարանցիկ - 1 ---" ,,, DivideColor   
      ' Մուտք Տարանցիկ թղթապանակ
      folderDirect = "|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|êï³óí³Í Ñ³ÝÓÝ³ñ³ñ³·ñ»ñ|î³ñ³ÝóÇÏ"
      stDate = "010103"
      eDate = "010122"
      wUser = ""
      wCue= ""
      paySysIn = ""
      paySysOut = ""
      docType = ""
      acsBranch = ""
      acsDepart = ""
      selectedView = "PayinsV"
      exportExcel = "0"
      state = False
      Call OpenSubjectToRegistrationFolder(folderDirect, stDate, eDate, wUser, wCue, notCur, docType, paySysIn, paySysOut, _
                                                             showLongNames, acsBranch, acsDepart, selectedView, exportExcel, state)
        
      ' Ստուգում է Տարանցիկ թղթապանակը բացվել է թե ոչ
      PttelName = "frmPttel"
      status =  WaitForPttel(PttelName)
      
      If  status Then
      
            ' Համեմատել պտտելի տողերի քանակը
            Call CheckPttel_RowCount("frmPttel", 6)

            ' Դասավորել ըստ Փաստաթղթի N  սյան
      			Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").Keys("[Hold]" & "^" & (2))
      
            ' Արտահանել թղթապանակի տողերը բացված Pttel-ից
            Path1 = Project.Path & "Stores\Reports\Main Accountant\Report_3\Actual\TransitAct_1.xlsx"
      
            ' Արտահանել excel ֆայլ
            Call ExportToExcel("frmPttel",Path1)

            ' Համեմատել երկու EXCEL ֆայլերի բոլոր sheet-երը
            Path2 = Project.Path &  "Stores\Reports\Main Accountant\Report_3\Expected\TransitExp_1.xlsx"
            resultWorksheet = Project.Path &  "Stores\Reports\Main Accountant\Report_3\Result\TransitRes_1.xlsx"
            Call CompareTwoExcelFiles(Path1, Path2, resultWorksheet)
      
            ' Փակել Տարանցիկ թղթապանակը
            Call Close_Pttel(PttelName)
      
            ' Փակել բոլոր excel ֆայլերը
            Call CloseAllExcelFiles()
      
      Else
            Log.Error("Տարանցիկ թղթապանակը չի բացվել")
      End If
      
      '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ 
      Log.Message "--- Տարանցիկ - 2 ---" ,,, DivideColor           
      ' Մուտք Տարանցիկ թղթապանակ
      stDate = "110113"
      eDate = "120314"
      wUser = "59"
      wCue= "001"
      paySysIn = ""
      paySysOut = "A"
      docType = "1"
      acsBranch = "P00"
      acsDepart = "062"
      Call OpenSubjectToRegistrationFolder(folderDirect, stDate, eDate, wUser, wCue, notCur, docType, paySysIn, paySysOut, _
                                                                                showLongNames, acsBranch, acsDepart, selectedView, exportExcel, state)
        
      ' Ստուգում է Տարանցիկ թղթապանակը բացվել է թե ոչ
      status =  WaitForPttel(PttelName)
      
      If  status Then
      
            ' Համեմատել պտտելի տողերի քանակը
            Call CheckPttel_RowCount("frmPttel", 2)
      
            ' Դասավորել ըստ Փաստաթղթի N  սյան
      			Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").Keys("[Hold]" & "^" & (2))
      
            ' Արտահանել թղթապանակի տողերը բացված Pttel-ից
            Path1 = Project.Path & "Stores\Reports\Main Accountant\Report_3\Actual\TransitAct_2.xlsx"
      
            ' Արտահանել excel ֆայլ
            Call ExportToExcel("frmPttel",Path1)

            ' Համեմատել երկու EXCEL ֆայլերի բոլոր sheet-երը
            Path2 = Project.Path &  "Stores\Reports\Main Accountant\Report_3\Expected\TransitExp_2.xlsx"
            resultWorksheet = Project.Path &  "Stores\Reports\Main Accountant\Report_3\Result\TransitRes_2.xlsx"
            Call CompareTwoExcelFiles(Path1, Path2, resultWorksheet)
      
            ' Փակել Տարանցիկ թղթապանակը
            Call Close_Pttel(PttelName)
      
            ' Փակել բոլոր excel ֆայլերը
            Call CloseAllExcelFiles()
      
      Else
            Log.Error("Տարանցիկ թղթապանակը չի բացվել")
      End If
      
      '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ 
      Log.Message "--- Ընդհանուր դիտում - 1 ---" ,,, DivideColor        
      ' Մուտք Ընդհանուր դիտում թղթապանակ
      folderDirect = "|¶ÉË³íáñ Ñ³ßí³å³ÑÇ ²Þî|Ð³ßí»ïíáõÃÛáõÝÝ»ñ,  Ù³ïÛ³ÝÝ»ñ|ÀÝ¹Ñ³Ýáõñ ¹ÇïáõÙ"
      stDate = "010118"
      eDate = "010122"
      coaNum = "1"
      balAcc = "1601100"
      accMask = "16??0544001"
      wCur = "001"
      operType = "PRC"
      wUser = "10"
      showPrc = 1
      showRel = 1
      showRst = 1
      showAccNames = 1
      showPayres = 1
      showCrDate = 1
      acsBranch = "P00"
      acsDepart = "05"
      selectedView = "CommView"
      exportExcel = "0"
      Call OpenOverallviewFolder(folderDirect, stDate, eDate, coaNum, balAcc, accMask, wCur, operType, wUser, showPrc, showRel, showRst, _
                                                    showAccNames, showPayres, showCrDate, acsBranch, acsDepart, selectedView, exportExcel)
                                                  
      ' Ստուգում է Ընդհանուր դիտում թղթապանակը բացվել է թե ոչ
      status =  WaitForExecutionProgress() 
      
      If  status Then
      
            ' Համեմատել պտտելի տողերի քանակը
            Call CheckPttel_RowCount("frmPttel", 60)
      
            SortArr(0) = "fSUM"
            Call columnSorting(SortArr, 1, "frmPttel")
      
            ' Արտահանել թղթապանակի տողերը բացված Pttel-ից
            Path1 = Project.Path & "Stores\Reports\Main Accountant\Report_3\Actual\AccGeneralViewAct_1.xlsx"
      
            ' Արտահանել excel ֆայլ
            Call ExportToExcel("frmPttel",Path1)

            ' Համեմատել երկու EXCEL ֆայլերի բոլոր sheet-երը
            Path2 = Project.Path &  "Stores\Reports\Main Accountant\Report_3\Expected\AccGeneralViewExp_1.xlsx"
            resultWorksheet = Project.Path &  "Stores\Reports\Main Accountant\Report_3\Result\AccGeneralViewRes_1.xlsx"
            Call CompareTwoExcelFiles(Path1, Path2, resultWorksheet)
      
            ' Փակել Ընդհանուր դիտում թղթապանակը
            Call Close_Pttel(PttelName)
      
            ' Փակել բոլոր excel ֆայլերը
            Call CloseAllExcelFiles()
      
      Else
            Log.Error("Ընդհանուր դիտում թղթապանակը չի բացվել")
      End If
      
       '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------      
       Log.Message "--- Ընդհանուր դիտում - 2 ---" ,,, DivideColor        
      ' Մուտք Ընդհանուր դիտում թղթապանակ
      stDate = "010118"
      eDate = "010122"
      coaNum = "3"
      balAcc = "999999"
      accMask = "16283949300"
      wCur = "000"
      operType = "PRC"
      wUser = "10"
      showPrc = 1
      showRel = 1
      showRst = 1
      showAccNames = 0
      showPayres = 0
      showCrDate = 1
      acsBranch = "P00"
      acsDepart = "05"
      Call OpenOverallviewFolder(folderDirect, stDate, eDate, coaNum, balAcc, accMask, wCur, operType, wUser, showPrc, showRel, showRst, _
                                                    showAccNames, showPayres, showCrDate, acsBranch, acsDepart, selectedView, exportExcel)
                                                  
      ' Ստուգում է Ընդհանուր դիտում թղթապանակը բացվել է թե ոչ
      status =  WaitForPttel(PttelName) 
      
      If  status Then
      
            ' Համեմատել պտտելի տողերի քանակը
            Call CheckPttel_RowCount("frmPttel", 73)
      
            SortArr(0) = "fSUM"
            Call columnSorting(SortArr, 1, "frmPttel")
      
            ' Արտահանել թղթապանակի տողերը բացված Pttel-ից
            Path1 = Project.Path & "Stores\Reports\Main Accountant\Report_3\Actual\AccGeneralViewAct_2.xlsx"
      
            ' Արտահանել excel ֆայլ
            Call ExportToExcel("frmPttel",Path1)

            ' Համեմատել երկու EXCEL ֆայլերի բոլոր sheet-երը
            Path2 = Project.Path &  "Stores\Reports\Main Accountant\Report_3\Expected\AccGeneralViewExp_2.xlsx"
            resultWorksheet = Project.Path &  "Stores\Reports\Main Accountant\Report_3\Result\AccGeneralViewRes_2.xlsx"
            Call CompareTwoExcelFiles(Path1, Path2, resultWorksheet)
      
            ' Փակել Ընդհանուր դիտում թղթապանակը
            Call Close_Pttel(PttelName)
      
            ' Փակել բոլոր excel ֆայլերը
            Call CloseAllExcelFiles()
      
      Else
            Log.Error("Ընդհանուր դիտում թղթապանակը չի բացվել")
      End If
      
      Call Close_AsBank()  
      
End Sub