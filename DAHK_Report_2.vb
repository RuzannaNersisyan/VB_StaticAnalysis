Option Explicit
'USEUNIT Library_Common
'USEUNIT Constants
'USEUNIT Subsystems_SQL_Library
'USEUNIT  Debit_Dept_Filter_Library
'USEUNIT SWIFT_International_Payorder_Library
'USEUNIT Library_Colour
'USEUNIT OLAP_Library
'USEUNIT Mortgage_Library
'USEUNIT  Library_Contracts
'USEUNIT  DAHK_Library_Filter
'Test Case ID 161575

' ԴԱՀԿ հաղ. մշակման ԱՇՏ-ում Հաշիվներ ֆիլտրի ստուգում - 2
Sub DAHK_Report_2_Test()
      
      Dim fDATE, sDATE
      Dim folderDirect, folderName, stDate, eDate, messType, inqNumber, inquestId, sentMess, passTax
      Dim PttelName, status, Path1, Path2, resultWorksheet, clCode, clName, blockId, wSource, showClosed, wUser
      Dim SortArr(1)
      
      fDATE = "20220101"
      sDATE = "20030101"
      Call Initialize_AsBank("bank_Report", sDATE, fDATE)
      
      ' Մուտք գործել համակարգ ARMSOFT օգտագործողով 
      Call Create_Connection()
      Login("ARMSOFT")
      
      ' Մուտք ԴԱՀԿ հաղ. մշակման ԱՇՏ
      Call ChangeWorkspace(c_DAHK)
      
      ' Դրույթներից, Տնտեսել հիշողոթյունը սկսած (Տողերի քանակ) դաշտի արժեքի փոփոխում
      Call  SaveRAM_RowsLimit("10")
      
      '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
      Log.Message "--- Ուղարկված հաղորդագրություններ - 1 ---" ,,, DivideColor 
      ' Բացել Ուղարկված հաղորդագրություններ թղթապանակը 
      folderDirect = "|¸²ÐÎ Ñ³Õ. Ùß³ÏÙ³Ý ²Þî|àõÕ³ñÏí³Í"
      folderName = "Ուղարկված հաղորդագրություններ"
      stDate = "010113"
      eDate = "070413"
      messType = ""
      inqNumber = ""
      inquestId = ""
      passTax = ""
      sentMess = True 
      Call OpenEditableMessFolder(folderDirect, folderName, stDate, eDate, messType, inqNumber, inquestId, sentMess, passTax )
                                  
      ' Սպասում է այնքան մինչև "կատարման ընթացքը" վերջանա 
      Call  WaitForExecutionProgress()
      
      ' Ստուգում է Ուղարկված հաղորդագրություններ թղթապանակը բացվել է թե ոչ
      PttelName = "frmPttel"
      status =  WaitForPttel(PttelName) 
      
      If  status Then
      
          ' Համեմատել պտտելի տողերի քանակը
          Call CheckPttel_RowCount("frmPttel", 19622)
      
          ' Դասավորել ըստ Հղում սյան
          SortArr(0) = "REFERENCE"
          Call columnSorting(SortArr, 1, "frmPttel")
      
          BuiltIn.Delay(6000)
      
          ' Արտահանել թղթապանակի տողերը բացված Pttel-ից
          Path1 = Project.Path & "Stores\Reports\DAHK\Report_2\Actual\SentMessAct_1.xlsx"
      
          ' Արտահանել excel ֆայլ
          Call ExportToExcel("frmPttel",Path1)

          ' Համեմատել երկու excel ֆայլերը
          Path2 = Project.Path &  "Stores\Reports\DAHK\Report_2\Expected\SentMessExp_1.xlsx"
          resultWorksheet = Project.Path &  "Stores\Reports\DAHK\Report_2\Result\SentMessRes_1.xlsx"
          Call CompareTwoExcelFiles(Path1, Path2, resultWorksheet)
      
          ' Փակել Ուղարկված հաղորդագրություններ թղթապանակը
          Call Close_Pttel(PttelName)
      
          ' Փակել բոլոր excel ֆայլերը
          Call CloseAllExcelFiles()
      
      Else
            Log.Error("Ուղարկված հաղորդագրություններ թղթապանակը չի բացվել")
      End If
      
      '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
      Log.Message "--- Ուղարկված հաղորդագրություններ - 2 ---" ,,, DivideColor 
      ' Բացել Ուղարկված հաղորդագրություններ թղթապանակը 
      stDate = "301012"
      messType = "06"
      inqNumber = "07-00381/10"
      inquestId = ""
      passTax = "AG??38212"
      Call OpenEditableMessFolder(folderDirect, folderName, stDate, stDate, messType, inqNumber, inquestId, sentMess, passTax )
                                  
      ' Ստուգում է Ուղարկված հաղորդագրություններ թղթապանակը բացվել է թե ոչ
      status =  WaitForPttel(PttelName) 
      
      If  status Then
      
            ' Համեմատել պտտելի տողերի քանակը
            Call CheckPttel_RowCount("frmPttel", 1)
      
            ' Արտահանել թղթապանակի տողերը բացված Pttel-ից
            Path1 = Project.Path & "Stores\Reports\DAHK\Report_2\Actual\SentMessAct_2.txt"
      
            ' Արտահանել txt ֆայլ
            Call ExportToTXTFromPttel(PttelName,Path1)

            ' Համեմատել երկու txt ֆայլերը
            Path2 = Project.Path &  "Stores\Reports\DAHK\Report_2\Expected\SentMessExp_2.txt"
            Call Compare_Files(Path2, Path1, "")
      
            ' Փակել Ուղարկված հաղորդագրություններ թղթապանակը
            Call Close_Pttel(PttelName)
      
      Else
            Log.Error("Ուղարկված հաղորդագրություններ թղթապանակը չի բացվել")
      End If
      
      '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
      Log.Message "--- Ուղարկված հաղորդագրություններ - 3 ---" ,,, DivideColor 
      ' Բացել Ուղարկված հաղորդագրություններ թղթապանակը 
      stDate = "010413"
      eDate = "010414"
      messType = "04"
      inqNumber = ""
      inquestId = ""
      passTax = ""
      sentMess = True 
      Call OpenEditableMessFolder(folderDirect, folderName, stDate, eDate, messType, inqNumber, inquestId, sentMess, passTax )
                                  
      ' Ստուգում է Ուղարկված հաղորդագրություններ թղթապանակը բացվել է թե ոչ
      PttelName = "frmPttel"
      status =  WaitForPttel(PttelName) 
      
      If  status Then
      
            ' Համեմատել պտտելի տողերի քանակը
            Call CheckPttel_RowCount("frmPttel", 65896)
      
            ' Դասավորել ըստ Հղում սյան
            Call FastColumnSorting(SortArr, 1, "frmPttel")
      
            ' Արտահանել թղթապանակի տողերը բացված Pttel-ից
            Path1 = Project.Path & "Stores\Reports\DAHK\Report_2\Actual\SentMessAct_3.xlsx"
      
            ' Արտահանել excel ֆայլ
            Call ExportToExcel("frmPttel",Path1)

            ' Համեմատել երկու excel ֆայլերը
            Path2 = Project.Path &  "Stores\Reports\DAHK\Report_2\Expected\SentMessExp_3.xlsx"
            resultWorksheet = Project.Path &  "Stores\Reports\DAHK\Report_2\Result\SentMessRes_3.xlsx"
            Call CompareTwoExcelFiles(Path1, Path2, resultWorksheet)
      
            ' Փակել Ուղարկված հաղորդագրություններ թղթապանակը
            Call Close_Pttel(PttelName)
      
      ' Փակել բոլոր excel ֆայլերը
      Call CloseAllExcelFiles()
      
      Else
            Log.Error("Ուղարկված հաղորդագրություններ թղթապանակը չի բացվել")
      End If
      
      '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
      Log.Message "--- Գումարների արգելադրումներ - 1 ---" ,,, DivideColor 
      ' Բացել Գումարների արգելադրումներ թղթապանակը 
      folderDirect = "|¸²ÐÎ Ñ³Õ. Ùß³ÏÙ³Ý ²Þî|¶áõÙ³ñÝ»ñÇ ³ñ·»É³¹ñáõÙÝ»ñ"
      folderName = "Գումարների արգելադրումներ "
      stDate = "010103"
      eDate = "010122"
      clCode = "00026327"
      clName = "Ð³×³Ëáñ¹ 00026327"
      blockId = "º01000249973"
      wSource = ""
      showClosed = 1
      Call OpenMoneyBarriersFolder(folderDirect, folderName, stDate, eDate, clCode, clName, blockId, wSource, showClosed)
                                  
      ' Ստուգում է Գումարների արգելադրումներ թղթապանակը բացվել է թե ոչ
      status =  WaitForPttel(PttelName) 
      
      If  status Then
      
            ' Համեմատել պտտելի տողերի քանակը
            Call CheckPttel_RowCount("frmPttel", 1)
      
            ' Արտահանել թղթապանակի տողերը բացված Pttel-ից
            Path1 = Project.Path & "Stores\Reports\DAHK\Report_2\Actual\MoneyBarriersAct_1.txt"
      
            ' Արտահանել txt ֆայլ
            Call ExportToTXTFromPttel(PttelName,Path1)

            ' Համեմատել երկու excel ֆայլերը
            Path2 = Project.Path &  "Stores\Reports\DAHK\Report_2\Expected\MoneyBarriersExp_1.txt"
            Call Compare_Files(Path2, Path1, "")
      
            ' Փակել Գումարների արգելադրումներ թղթապանակը
            Call Close_Pttel(PttelName) 
      
      Else
            Log.Error("Գումարների արգելադրումներ թղթապանակը չի բացվել")
      End If
      
      '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
      Log.Message "--- Գումարների արգելադրումներ - 2 ---" ,,, DivideColor 
      ' Բացել Գումարների արգելադրումներ թղթապանակը 
      stDate = "050618"
      eDate = "010122"
      clCode = "00026335"
      clName = "Ð³×³Ëáñ¹ 00026335"
      blockId = ""
      wSource = ""
      showClosed = 0
      Call OpenMoneyBarriersFolder(folderDirect, folderName, stDate, eDate, clCode, clName, blockId, wSource, showClosed)
                                  
      ' Ստուգում է Գումարների արգելադրումներ թղթապանակը բացվել է թե ոչ
      status =  WaitForPttel(PttelName) 
      
      If  status Then
          ' Համեմատել պտտելի տողերի քանակը
          Call CheckPttel_RowCount("frmPttel", 0)
          BuiltIn.Delay(1000)
          ' Փակել Պտտելը  
          Call Close_Pttel("frmPttel")
      Else
            Log.Error("Գումարների արգելադրումներ թղթապանակը չի բացվել")
      End If
      
      
      '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
      Log.Message "--- Օգտագործողի սևագրեր ---" ,,, DivideColor 
      ' Բացել Օգտագործողի սևագրեր թղթապանակը 
      folderDirect = "|¸²ÐÎ Ñ³Õ. Ùß³ÏÙ³Ý ²Þî|ú·ï³·áñÍáÕÇ ë¨³·ñ»ñ"
      folderName = "Օգտագործողի սևագրեր "
      wUser = "253"
      Call OpenUserDraftFolder(folderDirect, folderName, wUser)
                       
      ' Ստուգում է Օգտագործողի սևագրեր թղթապանակը բացվել է թե ոչ
      status =  WaitForPttel(PttelName) 
      
      If  status Then
      
            ' Համեմատել պտտելի տողերի քանակը
            Call CheckPttel_RowCount("frmPttel", 1)
      
            ' Արտահանել թղթապանակի տողերը բացված Pttel-ից
            Path1 = Project.Path & "Stores\Reports\DAHK\Report_2\Actual\UserDraftAct_1.txt"
      
            ' Արտահանել txt ֆայլ
            Call ExportToTXTFromPttel(PttelName,Path1)

            ' Համեմատել երկու excel ֆայլերը
            Path2 = Project.Path &  "Stores\Reports\DAHK\Report_2\Expected\UserDraftExp_1.txt"
            Call Compare_Files(Path2, Path1, "")
      
            ' Փակել Օգտագործողի սևագրեր թղթապանակը
            Call Close_Pttel(PttelName)
      
      Else
            Log.Error("Օգտագործողի սևագրեր թղթապանակը չի բացվել")
      End If
      
      Call Close_AsBank()
End Sub