Option Explicit
'USEUNIT Library_Common
'USEUNIT Constants
'USEUNIT  Library_Contracts
'USEUNIT Subsystems_SQL_Library
'USEUNIT Library_Colour
'USEUNIT OLAP_Library
'USEUNIT Debit_Dept_Filter_Library
'USEUNIT Mortgage_Library

'Test Case ID 162431

' Փնտրել պատուհանի աշխատանքի ստուգում
Sub Check_Search_Filter_Test()

      Dim fDATE, sDATE
      Dim WorkingDocuments, folderName, arr, lastRowNum, count, i, state   
      Dim Path1, Path2, resultWorksheet, pttelName, status
      Dim folderDirect, wLevel, eDate, accBal, wAcc, pprCode, agrAccType, wCur, defaultCur, wClient, wName, wNote, wNote2, _
              wNote3, acsBranch, acsDepart, ascType, clientInfo, showInfo, showOutSum, showNotes, showAcc, wClose, notFullClose
      Dim stDate, agrNum, dealType, wUser, asType, accMaskOld, accMaskNew, tdbgFind
      Dim SortArr(5)
      
      fDATE = "20220101"
      sDATE = "20030101"
      Call Initialize_AsBank("bank_Report", sDATE, fDATE)
      
      ' Մուտք գործել համակարգ ARMSOFT օգտագործողով 
      Call Create_Connection()
      Login("ARMSOFT")
      
      ' Մուտք Տոկոսների հաշվարկման ԱՇՏ
      Call ChangeWorkspace(c_BillReceivables)
      
       Log.Message "--- Աշխատանքային փաստաթղթեր թղթապանակ -1---" ,,, DivideColor 
      Set WorkingDocuments = New_SubsystemWorkingDocuments()
      
      ' Աշխատանքային փաստաթղթեր ֆիլտրի բացում և տվյալների լրացում
      folderName = "|¸»µÇïáñ³Ï³Ý å³ñïù»ñ|"
      Call GoTo_SubsystemWorkingDocuments(folderName, WorkingDocuments)
      
      ' Ստուգում է Աշխատանքային փաստաթղթեր թղթապանակը բացվել է թե ոչ
      pttelName = "frmPttel"
      status =  WaitForPttel(pttelName) 
      
      If  status Then

            ' Համեմատել պտտելի տողերի քանակը
            Call CheckPttel_RowCount("frmPttel", 5)

            ' Բացել փնտրել պատուհանը
            count = 19
            arr = Array("27/03/17","¸»µÇïáñ³Ï³Ý å³ñïù","10200060400","000","0.00","0.00","00020031","Ð³×³Ëáñ¹ 00020031","AA00020031","0002003100","","","","253","Üáñ å³ÛÙ³Ý³·Çñ","10200060400","P00","03","BR1")
            Call FilterBySearchWindow(count, arr)

            ' Արտահանել excel ֆայլ
            Path1 = Project.Path & "Stores\Search\Actual\ActWorkingDoc.xlsx"
      
            Call ExportToExcel("frmPttel",Path1)
      
            ' Համեմատել երկու EXCEL ֆայլերի բոլոր sheet-երը
            Path2 = Project.Path &  "Stores\Search\Expected\ExpWorkingDoc.xlsx"
            resultWorksheet = Project.Path &  "Stores\Search\Result\ResWorkingDoc.xlsx"
            Call CompareTwoExcelFiles(Path1, Path2, resultWorksheet)
      
            ' Փակել Աշխատանքային փաստաթղթեր թղթապանակը
            Call Close_Pttel(pttelName)
      
            ' Փակել բոլոր excel ֆայլերը
            Call CloseAllExcelFiles()
      
      Else
            Log.Error("Աշխատանքային փաստաթղթեր թղթապանակը չի բացվել")
      End If
      
       '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
       Log.Message "--- Պայմանագրեր թղթապանակ -1 ---" ,,, DivideColor 
       ' Բացել և լրացնել Պայմանագրեր ֆիլտրը
       folderDirect = "|¸»µÇïáñ³Ï³Ý å³ñïù»ñ|ä³ÛÙ³Ý³·ñ»ñ"
       folderName = "Պայմանագրեր "
       wLevel = "1"
       clientInfo = 0
       showInfo = 1
       showOutSum = 1
       showNotes = 1
       showAcc = 1
       wClose = 1
       notFullClose = 1
       Call DebitDeptContractsFilter(folderDirect, folderName, wLevel, eDate, accBal, wAcc, pprCode, agrAccType, wCur, defaultCur, wClient, wName, wNote, wNote2, _
                                                     wNote3, acsBranch, acsDepart, ascType, clientInfo, showInfo, showOutSum, showNotes, showAcc, wClose, notFullClose )
                         
      ' Սպասում է այնքան մինչև "կատարման ընթացքը" վերջանա 
      Call  WaitForExecutionProgress()
                                  
      ' Ստուգում է Պայմանագրեր  թղթապանակը բացվել է թե ոչ
      PttelName = "frmPttel"
      status =  WaitForPttel(PttelName) 
      
      If  status Then
      
            ' Համեմատել պտտելի տողերի քանակը
            Call CheckPttel_RowCount("frmPttel", 870)

            ' Դասավորել ըստ պայմանագրերի համարի
            SortArr(0) = "fCODE"
            Call columnSorting(SortArr, 1, "frmPttel")
      
            ' Բացել փնտրել պատուհանը
            count = 40
            arr = Array("","","","","045","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","P00","02","BR1")
            Call FilterBySearchWindow(count, arr)
            
            ' Արտահանել թղթապանակի տողերը բացված Pttel-ից
            Path1 = Project.Path & "Stores\Search\Actual\ActContracts.xlsx"
      
            ' Արտահանել excel ֆայլ
            Call ExportToExcel("frmPttel",Path1)

            ' Համեմատել երկու EXCEL ֆայլերի բոլոր sheet-երը
            Path2 = Project.Path &  "Stores\Search\Expected\ExpContracts.xlsx"
            resultWorksheet = Project.Path &  "Stores\Search\Result\ResContracts.xlsx"
            Call CompareTwoExcelFiles(Path1, Path2, resultWorksheet)
      
            ' Փակել Պայմանագրեր թղթապանակը
            Call Close_Pttel(PttelName)
      
            ' Փակել բոլոր excel ֆայլերը
            Call CloseAllExcelFiles()
      
      Else
            Log.Error("Պայմանագրեր թղթապանակը չի բացվել")
      End If
      
       '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
       Log.Message "--- Հաշիվների խմբագրումներ թղթապանակ ---" ,,, DivideColor 
       ' Բացել և լրացնել Հաշիվների խմբագրումներ ֆիլտրը
       folderDirect = "|¸»µÇïáñ³Ï³Ý å³ñïù»ñ|¶áñÍáÕáõÃÛáõÝÝ»ñ, ÷á÷áËáõÃÛáõÝÝ»ñ|Ð³ßÇíÝ»ñÇ ËÙµ³·ñáõÙÝ»ñ"
       Call EditAccFromDebitDebt(folderDirect, stDate, eDate, agrNum, accMaskOld, accMaskNew, wUser)

      ' Սպասում է այնքան մինչև "կատարման ընթացքը" վերջանա 
      Call  WaitForExecutionProgress()
                                  
      ' Ստուգում է Հաշիվների խմբագրումներ թղթապանակը բացվել է թե ոչ
      PttelName = "frmPttel"
      status =  WaitForPttel(PttelName) 
      
      If  status Then
      
            ' Համեմատել պտտելի տողերի քանակը
            Call CheckPttel_RowCount("frmPttel", 226)
      
            ' Դասավորել ըստ ամսաթիվ, Պայմանագրի N, Անվանում սյուների
            SortArr(0) = "fAGRNUM"
            SortArr(1) = "fDATE"
            SortArr(2) = "fCOM"
            SortArr(3) = "ACCNEW"
            Call FastColumnSorting(SortArr, 4, "frmPttel")
            
            ' Բացել փնտրել պատուհանը
            count = 7
            arr = Array("17/02/12","10310030301","Ð³×³Ëáñ¹ 00002859","9450726000","19400050500","01 é.¹ - ä³Ñáõëï³íáñÙ³Ý Ñ³ßíÇ ËÙµ³·ñáõÙ","11")
            Call FilterBySearchWindow(count, arr)
      
            wMDIClient.Keys("~!")
            
            ' Արտահանել թղթապանակի տողերը բացված Pttel-ից
            Path1 = Project.Path & "Stores\Search\Actual\ActEditAcc.xlsx"
      
            ' Արտահանել excel ֆայլ
            Call ExportToExcel("frmPttel",Path1)

            ' Համեմատել երկու EXCEL ֆայլերի բոլոր sheet-երը
            Path2 = Project.Path &  "Stores\Search\Expected\ExpEditAcc.xlsx"
            resultWorksheet = Project.Path &  "Stores\Search\Result\ResEditAcc.xlsx"
            Call CompareTwoExcelFiles(Path1, Path2, resultWorksheet)
            
            wMDIClient.VBObject("frmPttel").Keys("~[F7]")
            BuiltIn.Delay(1000)
            
            Set tdbgFind = Sys.Process("Asbank").VBObject("FrmHorFind").VBObject("tdbgFind")
            Sys.Process("Asbank").VBObject("FrmHorFind").VBObject("Command3").Click 
            BuiltIn.Delay(1000)
            
            If Not (tdbgFind.Columns.Item(0) = "" and tdbgFind.Columns.Item(1) = "" and tdbgFind.Columns.Item(2) = "" and  tdbgFind.Columns.Item(3) = "" ) Then
                  Log.Error("Մաքրել նշիչը ճիշտ չի աշխատել")
            End If
            
            If Not ( tdbgFind.Columns.Item(4) = "" and tdbgFind.Columns.Item(5) = "" and tdbgFind.Columns.Item(6) = "" ) Then
                  Log.Error("Մաքրել նշիչը ճիշտ չի աշխատել")
            End If
            ' Սեղմել դադարեցնել կոճակը
            Sys.Process("Asbank").VBObject("FrmHorFind").VBObject("CancelButton").Click
      
            ' Փակել Հաշիվների խմբագրումներ թղթապանակը
            Call Close_Pttel(PttelName)
      
            ' Փակել բոլոր excel ֆայլերը
            Call CloseAllExcelFiles()
      
      Else
            Log.Error("Հաշիվների խմբագրումներ թղթապանակը չի բացվել")
      End If
      
       '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
      Log.Message "--- Հաշվեկշռային ձևակերպումներ ---" ,,, DivideColor 
       ' Բացել և լրացնել Հաշվեկշռային ձևակերպումներ ֆիլտրը
       folderDirect = "|¸»µÇïáñ³Ï³Ý å³ñïù»ñ|¶áñÍáÕáõÃÛáõÝÝ»ñ, ÷á÷áËáõÃÛáõÝÝ»ñ|Ð³ßí»Ïßé³ÛÇÝ Ó¨³Ï»ñåáõÙÝ»ñ"
       stDate = "010111"
       eDate = "010116"
       agrNum = ""
       wCur = "000"
       dealType = "21"
       wUser = "11"
       acsBranch = "P00" 
       acsDepart = "02"
       asType = "BR1"
       Call BalanceSheetFormulation(folderDirect, stDate, eDate, agrNum, wCur, dealType, wUser, wNote, wNote2, wNote3, acsBranch, acsDepart, asType)

      ' Սպասում է այնքան մինչև "կատարման ընթացքը" վերջանա 
      Call  WaitForExecutionProgress()
                                  
      ' Ստուգում է Հաշվեկշռային ձևակերպումներ թղթապանակը բացվել է թե ոչ
      status =  WaitForPttel(PttelName) 
      
      If  status Then
      
            ' Համեմատել պտտելի տողերի քանակը
            Call CheckPttel_RowCount("frmPttel", 40)
      
            ' Դասավորել ըստ ամսաթիվ, Անվանում, Հ/Պ դեբետ, Գումար, Հ/Պ կրեդիտ սյուների
            SortArr(0) = "fDATE"
            SortArr(1) = "ACCCR"
            SortArr(2) = "fCOM"
            SortArr(3) = "BALDB"
            SortArr(4) = "BALCR"
            Call FastColumnSorting(SortArr, 5, "frmPttel")
      
            ' Բացել փնտրել պատուհանը
            count = 12
            arr = Array("","","","","19520000100","","","","","","","11")
            Call FilterBySearchWindow(count, arr)
            
            ' Արտահանել թղթապանակի տողերը բացված Pttel-ից
            Path1 = Project.Path & "Stores\Search\Actual\ActBalanceSheet.xlsx"
      
            ' Արտահանել excel ֆայլ
            Call ExportToExcel("frmPttel",Path1)

            ' Համեմատել երկու EXCEL ֆայլերի բոլոր sheet-երը
            Path2 = Project.Path &  "Stores\Search\Expected\ExpBalanceSheet.xlsx"
            resultWorksheet = Project.Path &  "Stores\Search\Result\ResBalanceSheet.xlsx"
            Call CompareTwoExcelFiles(Path1, Path2, resultWorksheet)
      
            ' Փակել Հաշվեկշռային ձևակերպումներ թղթապանակը
            Call Close_Pttel(PttelName)
      
            ' Փակել բոլոր excel ֆայլերը
            Call CloseAllExcelFiles()
      
      Else
            Log.Error("Հաշվեկշռային ձևակերպումներ թղթապանակը չի բացվել")
      End If
      
      Call Close_AsBank()
      
End Sub
      

 ' Բացել փնտրել պատուհանը, լրացնել տվյալները, որոնել
 ' count - փնտրել պատուհանում որոնման համար լրացվող դաշտերի քանակ
 ' arr - զանգված, ստանում է լրացվող արժեքները
Sub FilterBySearchWindow(count, arr)
      
      Dim  i, state, frmHorFind
      i = 0
      state = False
      wMDIClient.vbObject("frmPttel").vbObject("tdbgView").MoveFirst
      
      wMDIClient.VBObject("frmPttel").Keys("~[F7]")
      BuiltIn.Delay(1000)
      Set frmHorFind = Sys.Process("Asbank").VBObject("FrmHorFind")
      If Sys.Process("Asbank").WaitVBObject("FrmHorFind", 2000).Exists Then
      
            For  i= 0 To count - 1
                  frmHorFind.VBObject("tdbgFind").Window("Edit", "", 1).Keys(arr(i) & "[Tab]")
            Next
            
                Do Until state
                      ' Սեղմել "Հաջորդը" կոճակը
                      frmHorFind.VBObject("Command1").Click
                      if  Sys.Process("Asbank").WaitVBObject("frmAsMsgBox", 1000).Exists Then
                            If  MessageExists(2, "²Û¹åÇëÇ ·ñ³éáõÙ ãÏ³") Then
                                  Call ClickCmdButton(5, "OK")  
                                  state = True
                                  Exit Do
                            End If
                      End If 
                                                
                      ' Սեղմել "Նշել" կոճակը
                      frmHorFind.VBObject("cmdSelect").Click
                Loop
            
            ' Սեղմել դադարեցնել կոճակը
            frmHorFind.VBObject("CancelButton").Click
      End If
End Sub