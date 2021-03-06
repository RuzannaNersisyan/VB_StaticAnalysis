Option Explicit
'USEUNIT  Library_Common
'USEUNIT Constants
'USEUNIT Subsystems_SQL_Library
'USEUNIT Debit_Dept_Library
'USEUNIT Clients_Library
'USEUNIT Payment_Order_ConfirmPhases_Library
'USEUNIT Overdraft_NewCases_Library
'USEUNIT Loan_Agreemnts_With_Schedule_Library
'USEUNIT Subsystems_Special_Library
'USEUNIT Mortgage_Library

' Test Case ID 132304
 
Sub Bill_Receivables_Test()

      Dim fDATE, sDATE
      Dim direction, wColNum, contType, fISN, cliCode, comment, accAcc, wDate, acsBranch, branchSect, acsType,_
               autoDebt, useAccBalanc, accConnect, headNum, autoDateChild, typeAutoDate, fixedDays, agrPeriod,_
               agrPeryodDay, passDirrect, passType, dateAgr, clsDays, state, brType, notClass, subjRisk,_
               sector, wAim, riksDegree, repCode, wNote, wNote2, wNote3, pprCode, dateClose, cenceled, n16AccType, _
               fillAccs, complRef, status, storageAcc, cost, income, accOutAgr
              
      Dim folderDirect, folderName, rekvName, writeOffISN, writeOffBackISN, debtLetISN, storeISN
      
      Dim frmPttel, confPath, confInput, risk, perc, queryString, sqlValue, sql_isEqual, docName
      
      Dim Prov, Removal, sumRes, sumUnres, wComment, acsSect, sumAgr, action, wSumma, param
      
      Dim wFrmPttel, colN, docTypeName, wDbt, wRes, wOut, wInc, wCls, wRsk, colNum
      
      Dim paramN, OutISN , acsDepart, fillDefault , colN2, tdbgView, wCommentP, wCommentD
      
      Dim  fillAcsBranch, fillAcsDepart, fillAcsType, riskClassfISN, riskClassfISN2, dbtISN, reservISN
      
      Dim  sDatePar, dateGive, eDatePar, dateCl,wCenceled, wBranch, wDepart, wAcsType
      
      fDATE = "20250101"
      sDATE = "20120101"
      Call Initialize_AsBank("bank", sDATE, fDATE)
      
      ' Մուտք գործել համակարգ ARMSOFT օգտագործողով 
      Call Create_Connection()
      Login("ARMSOFT")
      
      ' Մուտք Դեբիտորական պարտքեր ԱՇՏ
      Call ChangeWorkspace(c_BillReceivables)

      ' Դեբիտորական պարտք փաստաթղթի ստեղծում
      direction = "|¸»µÇïáñ³Ï³Ý å³ñïù»ñ|Üáñ å³ÛÙ³Ý³·ñÇ ëï»ÕÍáõÙ"
      contType = "¸»µÇïáñ³Ï³Ý å³ñïù"
      status = True
      wColNum = 1 
      state = True
      cliCode = "000329601"  
      wDate = "010120"
      acsBranch = "00"
      branchSect = "1"
      acsType = "BR0"
      autoDateChild = 1
      autoDebt = 1
      useAccBalanc = "3"
      accConnect = "001"
      typeAutoDate = "1"
      fixedDays = "15"
      passDirrect = "3"
      brType = "9" 
      sector = "C"
      wAim = "00"
      riksDegree = "0.00"
      repCode = "10102"
      fillAccs = 1
      complRef = 1
      subjRisk = 1
      storageAcc = "19510066600"
      cost = "001344400"
      income = "000441900"  
      accOutAgr = "8000100/000008"
      Call SelectContracType(direction, wColNum, contType, fISN, cliCode, cliCode, comment, wDate, acsBranch, branchSect, acsType,_
                                            autoDebt, useAccBalanc, accConnect, headNum, autoDateChild, typeAutoDate, fixedDays, agrPeriod,_
                                            agrPeryodDay, passDirrect, passType, dateAgr, clsDays, state, brType, notClass, subjRisk,_
                                            sector, wAim, riksDegree, repCode, wNote, wNote2, wNote3, pprCode, dateClose, cenceled, n16AccType, _
                                            fillAccs, complRef, status, storageAcc, cost, income, accOutAgr)
                                            
      Log.Message("Դեբիտորական պարտք փաստաթղթի ISN` " & fISN)
   
      ' Մուտք աշխատանքային փաստաթղթեր թղթապանակ և  պայմանագրի առկայության ստուգում
      folderDirect = "|¸»µÇïáñ³Ï³Ý å³ñïù»ñ|²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ"
      folderName = "Աշխատանքային փաստաթղթեր"
      rekvName = "NUM"
      state =  OpenFolder(folderDirect, folderName, rekvName, cliCode)
                           
      If Not state Then
            Log.Error("Դեբիտորական պարտք փաստաթուղթն առկա չէ Աշխատանքային փաստաթղթեր թղթապանակում")
      End If
                  'CONTRACTS
                  queryString = " SELECT COUNT(*) FROM CONTRACTS WHERE fDGISN = " & fISN & _
                                            " and fDGSUMMA = '0.00' and fDGALLSUMMA = '0.00' " & _
                                            " and fDGRISKDEGNB = '0.00' and fDGRISKDEGREE = '0.00' " & _
                                            " and fDGCUR = '001' and fDGMPERCENTAGE = '0.00' "
                  sqlValue = 1
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                        
      ' Փաստաթուղթն ուղարկել հաստատման
      Call PaySys_Send_To_Verify()
      Set frmPttel = wMDIClient.VBObject("frmPttel")
      Call Close_Pttel("frmPttel")
      
                  'HIF
                  queryString = " SELECT COUNT(*) FROM HIF WHERE fBASE = " & fISN & _
                                            " and ((fSUM = '0.00' and fCURSUM = '0.00' and fOP = 'ORC' and fTRANS = '1') " & _
                                            " or (fSUM = '1.00' and fCURSUM = '0.00' and fOP = 'PRS' and fTRANS = '1') " & _
                                            " or (fSUM = '0.00' and fCURSUM = '0.00' and fOP = 'RSK' and fTRANS = '1')) "
                  sqlValue = 3
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
      ' Մուտք Պայմանագրեր թղթապանակ և  պայմանագրի առկայության ստուգում
      folderDirect = "|¸»µÇïáñ³Ï³Ý å³ñïù»ñ|ä³ÛÙ³Ý³·ñ»ñ"
      folderName = "Պայմանագրեր"
      rekvName = "ACC"
      state =  OpenFolder(folderDirect, folderName, rekvName, cliCode)
                           
      If Not state Then
            Log.Error("Դեբիտորական պարտք փաստաթուղթն առկա չէ Պայմանագրեր թղթապանակում")
      End If
                  
      ' Ռիսկի դասիչ և պահուստավորման տոկոս գործողության կատարում
      wDate = "010120"
      risk = "05"
      perc = "100"
      Call FillDoc_Risk_Classifier(wDate, risk, perc)

      wFrmPttel = "frmPttel_2"
      paramN = c_ViewEdit & "|" & c_Risking & "|" & c_RisksPersRes
      colN = 0
      Call GetfISNFromActionsView(paramN, wDate, wDate, wFrmPttel, colN, cliCode, riskClassfISN )
      Log.Message(" Ռիսկի դասիչ և պահուստավորման տոկոս փաստաթղթի ISN` " & riskClassfISN)

                 'HI
                  queryString = " SELECT COUNT(*) FROM HIF WHERE fBASE = " & riskClassfISN & _
                                            " and fCURSUM = '0.00' and fTYPE = 'N0' and fCURSUM = '0.00'  " & _
                                            " and ((fSUM = '100.00' and fOP = 'PRS')" & _
                                            " or (fSUM = '0.00' and fOP = 'RSK' and fSPEC = '05')) " 
                  sqlValue = 2
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
                  'HI
                  queryString = " Select COUNT(*) from DOCS where fISN = " & riskClassfISN & " and fNAME = 'BRTSRsPr' and fBODY = '" & vbCRLF _
                                           & "CODE:000329601" & vbCRLF _
                                           & "DATE:20200101" & vbCRLF _
                                           & "RISK:05" & vbCRLF _
                                           & "PERRES:100" & vbCRLF _
                                           & "COMMENT:èÇëÏÇ ¹³ëÇã ¨ å³Ñáõëï³íáñÙ³Ý ïáÏáë" & vbCRLF _
                                           & "USERID:  77" & vbCRLF _
                                           & "'" 
                  sqlValue = 1
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
      
      ' Պահուստավորում գործողության կատարում
      wDate = "020120"
      action = c_Store
      state = False
      sumRes = "1000000"
      sumUnres = "0.00"
      wCommentP = "ä³Ñáõëï³íáñáõÙ"
      Call ProvisionAction(action, storeISN, wDate, state, wSumma, sumAgr, sumRes, sumUnres, wCommentP, acsBranch, acsSect)
      Log.Message("Պահուստավորում փաստաթղթի ISN` " & storeISN)
      
                  'HI
                  queryString = " SELECT COUNT(*) FROM HI WHERE fBASE = " & storeISN & _
                                            " and ((fSUM = '1000000.00' and fCUR = '000' and fCURSUM = '1000000.00' and fDBCR = 'D') " & _
                                            " or (fSUM = '1000000.00' and fCUR = '000' and fCURSUM = '1000000.00' and fDBCR = 'C')) " 
                  sqlValue = 2
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
                  'HIR
                  queryString = " SELECT fCURSUM FROM HIR WHERE fBASE = " & storeISN & " and fCUR = '000' "
                  sqlValue = 1000000.00
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
                  'HIRREST
                  queryString = " SELECT fLASTREM FROM HIRREST WHERE fOBJECT = " & fISN & _
                                            " and fPENULTREM = '0.00' and fSTARTREM = '0.00'"
                  sqlValue = 1000000.00
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
      ' Դուրս գրում գործողության կատարում
      action = c_WriteOff
      sumAgr = "1000"
      Call WriteOffOnAction(action, writeOffISN, wDate, sumAgr, wComment, acsBranch, acsSect)
      Log.Message("Դուրս գրում փաստաթղթի ISN` " & writeOffISN)
      
                  'HI
                  queryString = " SELECT COUNT(*) FROM HI WHERE fBASE = " & writeOffISN & _
                                            " and ((fSUM = '400000.00' and fCUR = '001' and fCURSUM = '1000.00' and fDBCR = 'C') " & _
                                            " or (fSUM = '400000.00' and fCUR = '000' and fCURSUM = '400000.00' and fDBCR = 'D')" & _
                                            " or (fSUM = '400000.00' and fCUR = '001' and fCURSUM = '1000.00' and fDBCR = 'D')) " 
                  sqlValue = 3
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If

                  'HIR
                  queryString = " SELECT COUNT(*) FROM HIR WHERE fBASE = " & writeOffISN &_
                                           " and ((fCUR = '000' and fCURSUM = '400000.00')" &_
                                           " or (fCUR = '001' and fCURSUM = '1000.00'))"
                  sqlValue = 2
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
                  'HIRREST
                  queryString = " SELECT COUNT(*) FROM HIRREST WHERE fOBJECT = " & fISN & _
                                            " and ((fLASTREM = '600000.00' and fPENULTREM = '0.00' and fSTARTREM = '0.00')" &_
                                            " or (fLASTREM = '1000.00' and fPENULTREM = '0.00' and fSTARTREM = '0.00'))"
                  sqlValue = 2
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
 
      ' Պարտքերի խմբային մարում
      wDate = "150220"
      wDbt = 1
      Call GroupCalculation(wDate, wDate, wDbt, wRes, wOut, wInc, wCls, wRsk)
      
      ' մարում փաստաթղթի ISN-ի ստացում գործողությունների դիտում թղթապանակից 
      paramN = c_OpersView
      colN = 5
      docTypeName = "¸»µÇïáñ³Ï³Ý å³ñïùÇ Ù³ñáõÙ"
      Call GetfISNFromActionsView(paramN, wDate, wDate, wFrmPttel, colN, docTypeName, dbtISN )
      log.Message("Մարում փաստաթղթի ISN` " & dbtISN)
      
                 'HI
                  queryString = " SELECT COUNT(*) FROM HI WHERE fBASE = " & dbtISN & "and fOP = 'MSC' " & _
                                            " and ((fTYPE = '01' and fSUM = '400000.00' and fCUR = '001' and fCURSUM = '1000.00' and fDBCR = 'D') " & _
                                            " or (fTYPE = '01' and fSUM = '3615016.00' and fCUR = '001' and fCURSUM = '9037.54' and fDBCR = 'D')" & _
                                            " or (fTYPE = '01' and fSUM = '3615016.00' and fCUR = '001' and fCURSUM = '9037.54' and fDBCR = 'C')" & _
                                            " or (fTYPE = '01' and fSUM = '400000.00' and fCUR = '000' and fCURSUM = '400000.00' and fDBCR = 'C') " & _
                                            " or (fTYPE = '02' and fSUM = '400000.00' and fCUR = '001' and fCURSUM = '1000.00' and fDBCR = 'C'))"
                  sqlValue = 5
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If

                  'HIR
                  queryString = " SELECT COUNT(*) FROM HIR WHERE fBASE = " & dbtISN & " and fADB = '0' " &_
                                           " and (( fTYPE = 'R4' and fCUR = '000' and fCURSUM = '400000.00' and fOP = 'INC' and fDBCR = 'D')" &_
                                           " or ( fTYPE = 'R5' and fCUR = '001' and fCURSUM = '1000.00' and fOP = 'INC' and fDBCR = 'C')" &_
                                           " or ( fTYPE = 'RI' and fCUR = '001' and fCURSUM = '9037.54' and fOP = 'DBT' and fDBCR = 'D'))"
                  sqlValue = 3
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
                  'HIRREST
                  queryString = " SELECT COUNT(*) FROM HIRREST WHERE fOBJECT = " & fISN & "and fSTARTREM = '0.00' " &_
                                            " and ((fLASTREM = '1000000.00' and fPENULTREM = '600000.00' and fTYPE = 'R4' )" &_
                                            " or (fLASTREM = '0.00' and fPENULTREM = '1000.00' and fTYPE = 'R5')" &_
                                            " or (fLASTREM = '9037.54' and fPENULTREM = '0.00' and fTYPE = 'RI'))"
                  sqlValue = 3 
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If

                   'DOCS
                  queryString = " SELECT COUNT(*) from DOCS where fISN = " & dbtISN & "and fNAME = 'BRGrOp' "
                  
                  sqlValue = 1
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                        
      ' Դուրս գրում գործողության կատարում
      wDate = "160220"
      action = c_WriteOff
      sumAgr = "1000"
      Call WriteOffOnAction(action, writeOffISN, wDate, sumAgr, wComment, acsBranch, acsSect)
      Log.Message("Դուրս գրում փաստաթղթի ISN` " & writeOffISN)
      
                  'HI
                  queryString = " SELECT COUNT(*) FROM HI WHERE fBASE = " & writeOffISN & " and fSUM = '400000.00' and fOP = 'MSC'" &_
                                            " and ((fTYPE = '01'  and fCUR = '001' and fCURSUM = '1000.00' and fDBCR = 'C') " & _
                                            " or (fTYPE = '01'  and fCUR = '000' and fCURSUM = '400000.00' and fDBCR = 'D')" & _
                                            " or (fTYPE = '02' and fCUR = '001' and fCURSUM = '1000.00' and fDBCR = 'D')) " 
                  sqlValue = 3
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If

                  'HIR
                  queryString = " SELECT COUNT(*) FROM HIR WHERE fBASE = " & writeOffISN &" and fOP = 'OUT' " & _
                                           " and ((fTYPE = 'R4'  and fCUR = '000' and fCURSUM = '400000.00' and fDBCR = 'C')" &_
                                           " or (fTYPE = 'R5'  and fCUR = '001' and fCURSUM = '1000.00' and fDBCR = 'D'))"
                  sqlValue = 2
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
                  'HIRREST
                  queryString = " SELECT COUNT(*) FROM HIRREST WHERE fOBJECT = " & fISN &  " and fSTARTREM = '0.00' " & _
                                            " and ((fTYPE = 'R4'  and fLASTREM = '600000.00' and fPENULTREM = '1000000.00')" &_
                                            "or (fTYPE = 'R5'  and fLASTREM = '1000.00' and fPENULTREM = '0.00' ) " &_
                                            " or (fTYPE = 'RI'  and fLASTREM = '9037.54' and fPENULTREM = '0.00'))"
                  sqlValue = 3
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
      
      ' Դուրս գրվածի վերականգնում գործողության կատարում
      wDate = "200220"
      action = c_WriteOffBack
      sumAgr = "500"
      Call WriteOffOnAction(action, writeOffBackISN, wDate, sumAgr, wComment, acsBranch, acsSect)
      Log.Message("Դուրս գրվածի վերականգնում փաստաթղթի ISN` " & writeOffBackISN)
            
                  'HI
                  queryString = " SELECT COUNT(*) FROM HI WHERE fBASE = " & writeOffBackISN & " and fSUM = '200000.00' and fOP = 'MSC'" & _
                                            " and ((fTYPE = '01'  and fCUR = '001' and fCURSUM = '500.00' and fDBCR = 'D')" & _
                                            " or (fTYPE = '01'  and fCUR = '000' and fCURSUM = '200000.00' and fDBCR = 'C')" & _
                                            " or (fTYPE = '02' and fCUR = '001' and fCURSUM = '500.00' and fDBCR = 'C')) " 
                  sqlValue = 3
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If

                  'HIR
                  queryString = " SELECT COUNT(*) FROM HIR WHERE fBASE = " & writeOffBackISN &_
                                           " and ((fTYPE = 'R4'  and fCUR = '000' and fCURSUM = '200000.00' and fDBCR = 'D')" &_
                                           " or (fCUR = '001' and fCURSUM = '500.00'))"
                  sqlValue = 2
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
                  'HIRREST
                  queryString = " SELECT COUNT(*) FROM HIRREST WHERE fOBJECT = " & fISN & _
                                            " and ((fTYPE = 'R4'  and fLASTREM = '800000.00' and fPENULTREM = '600000.00')" &_
                                            " or (fTYPE = 'R5'  and fLASTREM = '500.00' and fPENULTREM = '1000.00' )" &_
                                            " or (fTYPE = 'RI'  and fLASTREM = '9037.54' and fPENULTREM = '0.00'))"
                  sqlValue = 3
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
      ' Պարտքերի զիջում գործողության կատարում
      wDate = "010320"
      action = c_DebtLet
      sumRes = ""
      sumUnres = "500"
      wCommentD = "ä³ñïù»ñÇ ½ÇçáõÙ"
      state = True
      Call ProvisionAction(action, debtLetISN, wDate, state, wSumma, sumAgr, sumRes, sumUnres, wCommentD, acsBranch, acsSect)
      Log.Message(" Պարտքերի զիջում փաստաթղթի ISN` " & debtLetISN)
 
                 'HI
                  queryString = " SELECT COUNT(*) FROM HI WHERE fBASE = " & debtLetISN & _
                                            " and fTYPE = '02' and fSUM = '200000.00' and fCURSUM = '500.00' and fCUR = '001' and fOP = 'MSC' and fDBCR = 'C' "
                  sqlValue = 1
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If

                  'HIR
                  queryString = " SELECT COUNT(*) FROM HIR WHERE fBASE = " & debtLetISN &_
                                           "and fTYPE = 'R5' and fCURSUM = '500.00' and fCUR = '001' and fOP = 'LET' and fDBCR = 'C'"
                  sqlValue = 1
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
                  'HIRREST
                  queryString = " SELECT COUNT(*) FROM HIRREST WHERE fOBJECT = " & fISN & _
                                            " and ((fTYPE = 'R4'  and fLASTREM = '800000.00' and fPENULTREM = '600000.00')" &_
                                            " or (fTYPE = 'R5'  and fLASTREM = '0.00' and fPENULTREM = '500.00' )" &_
                                            " or (fTYPE = 'RI'  and fLASTREM = '9037.54' and fPENULTREM = '0.00'))"
                  sqlValue = 3
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
    
      ' Փոխել Գրասենյակը|Բաժինը|Տիպը գործողության կատարում
      acsBranch = "01"
      acsDepart = "2"
      acsType = "BR0"
      fillAcsBranch = 1
      fillAcsDepart = 1
      fillAcsType = 1
      status = True
      Call ChangeBranchDepartType(status, fillAcsBranch, fillAcsDepart, fillAcsType, acsBranch, acsDepart, acsType, fillDefault )
      
                   'HIF
                  queryString = " select COUNT(*) from HIF where fBASE  = " & fISN &_
                                           " and fTYPE = 'N0' and fCURSUM = '0.00' and fADB = '0'" &_
                                           " and ((fSUM = '0.00' and fOP = 'ORC')" &_
                                           " or (fSUM = '1.00' and fOP = 'PRS')" &_
                                           " or (fSUM = '0.00' and fOP = 'RSK'))"
                  sqlValue = 3
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
                  'HIRREST
                  queryString = " Select COUNT(*) from DOCS where fISN = " & fISN & " and fNAME = 'BRBill' and fBODY = '"& vbCRLF _
                                            & "CODE:000329601"& vbCRLF _
                                            & "CLICOD:00000104"& vbCRLF _
                                            & "ACCTYPE:1"& vbCRLF _
                                            & "ACCAGR:000329601"& vbCRLF _
                                            & "NAME:Î³ó³Ï Ø³ó³ÏÛ³Ý"& vbCRLF _
                                            & "CURRENCY:001"& vbCRLF _
                                            & "ACCACC:000329601"& vbCRLF _
                                            & "DATE:20200101"& vbCRLF _
                                            & "ACSBRANCH:01"& vbCRLF _
                                            & "ACSDEPART:2"& vbCRLF _
                                            & "ACSTYPE:BR0"& vbCRLF _
                                            & "JURSTAT:21"& vbCRLF _
                                            & "VOLORT:7"& vbCRLF _
                                            & "PETBUJ:2"& vbCRLF _
                                            & "REZ:1"& vbCRLF _
                                            & "RELBANK:0"& vbCRLF _
                                            & "RABBANK:0"& vbCRLF _
                                            & "AUTODEBT:1"& vbCRLF _
                                            & "DEBTJPART1:2"& vbCRLF _
                                            & "DEBTJPART:0"& vbCRLF _
                                            & "ACCCONNMODE:3"& vbCRLF _
                                            & "ACCCONNSCH:001"& vbCRLF _
                                            & "AUTODATECHILD:1"& vbCRLF _
                                            & "TYPEAUTODATE:1"& vbCRLF _
                                            & "FIXEDDAYS:15"& vbCRLF _
                                            & "PASSOVDIRECTION:3"& vbCRLF _
                                            & "PASSOVTYPE:0"& vbCRLF _
                                            & "SHOWINSTAT:0"& vbCRLF _
                                            & "BRTYPE:9"& vbCRLF _
                                            & "SECTOR:C"& vbCRLF _
                                            & "AIM:00"& vbCRLF _
                                            & "PERRES:1"& vbCRLF _
                                            & "REPCODE:10102"& vbCRLF _
                                            & "NOTCLASS:0"& vbCRLF _
                                            & "SUBJRISK:1"& vbCRLF _
                                            & "ISNBOUT:0"& vbCRLF _
                                            & "FILLACCS:0"& vbCRLF _
                                            & "OPENACCS:0"& vbCRLF _
                                            & "ACCOUTAGR:8000100/000008"& vbCRLF _
                                            & "'"
                  sqlValue = 1
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
      
      
      ' Ստուգել Գրասենյակ և Բաժին դաշտերի արժեքները
      wBranch = wMDIClient.VBObject("frmPttel").GetColumnIndex("ACSBRANCH")
      wDepart = wMDIClient.VBObject("frmPttel").GetColumnIndex("ACSDEPART")
      wAcsType = wMDIClient.VBObject("frmPttel").GetColumnIndex("ACSTYPE")
       Set tdbgView = frmPttel.VBObject("tdbgView")
      If  Not (Trim( tdbgView.Columns.Item(wBranch).Value) = Trim(acsBranch)  and   Trim( tdbgView.Columns.Item(wDepart).Value) = Trim(acsDepart)  and  Trim( tdbgView.Columns.Item(wAcsType).Value) = Trim(acsType))   Then
            Log.Error("Փոխել Գրասենյակը|Բաժինը|Տիպը գործողության կատարումը հաջողությամբ չի իրականացել")
      End If
                  
      ' Պահուստավորում գործողության կատարում
      wDate = "010420"
      action = c_Store
      state = False
      sumUnres = "0.00"
      docTypeName = "ä³Ñáõëï³íáñáõÙ"
      Call ProvisionAction(action, storeISN, wDate, state, wSumma, sumAgr, sumRes, sumUnres, wCommentP, acsBranch, acsSect)
      Log.Message("Պահուստավորում փաստաթղթի ISN` " & storeISN)

                  'HI
                  queryString = " SELECT COUNT(*) FROM HI WHERE fBASE = " & storeISN &_
                                            " and fSUM = '2615016.00' and fOP = 'RST' and fTYPE = '01'  " &_
                                            " and fCURSUM = '2615016.00' and  fCUR = '000' " &_
                                            " and ( fDBCR = 'D' or fDBCR = 'C') "
                  sqlValue = 2
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If

                  'HIR
                  queryString = " SELECT COUNT(*) FROM HIR WHERE fBASE = " & storeISN & " and fOP = 'RES' " &_ 
                                            "and fTYPE = 'R4'  and fCUR = '000' and fCURSUM = '2615016.00' and fDBCR = 'D'  " 
                  sqlValue = 1
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
                  'HIRREST
                  queryString = " SELECT COUNT(*) FROM HIRREST WHERE fOBJECT = " & fISN & _
                                            " and ((fTYPE = 'R4'  and fLASTREM = '3415016.00' and fPENULTREM = '800000.00')" &_
                                            " or (fTYPE = 'R5'  and fLASTREM = '0.00' and fPENULTREM = '500.00' )" &_
                                            " or (fTYPE = 'RI'  and fLASTREM = '9037.54' and fPENULTREM = '0.00'))"
                  sqlValue = 3
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
      ' Խմբային Դուրս գրում գործողության կատարում
      wOut = "1"
      wDbt = "0"
      Call GroupCalculation(wDate, wDate, wDbt, wRes, wOut, wInc, wCls, wRsk)
                 
      ' Դուրս գրում փաստաթղթի ISN-ի ստացում գործողությունների դիտում թղթապանակից 
      paramN = c_OpersView
      colN = 5
      docTypeName = "¸»µÇïáñ³Ï³Ý å³ñïùÇ ¹áõñë·ñáõÙ"
      Call GetfISNFromActionsView(paramN, wDate, wDate, wFrmPttel, colN, docTypeName, OutISN )
      log.Message("Խմբային Դուրս գրում փաստաթղթի ISN` " & OutISN)
      
                 'HI
                  queryString = " SELECT COUNT(*) FROM HI WHERE fBASE = " & OutISN &  " and fOP = 'MSC' "  &_
                                            " and ((fTYPE = '01' and fSUM = '3415016.00' and fCUR = '001' and fCURSUM = '8537.54' and fDBCR = 'C') " &_
                                            " or (fTYPE = '01'and  fSUM = '3415016.00' and fCUR = '000' and fCURSUM = '3415016.00' and fDBCR = 'D') " &_
                                            " or (fTYPE = '02' and fSUM = '3415016.00' and fCUR = '001' and fCURSUM = '8537.54' and fDBCR = 'D'))"
                  sqlValue = 3
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If

                  'HIR
                  queryString = " SELECT COUNT(*)  FROM HIR WHERE fBASE = " & OutISN & "and fOP = 'OUT' " &_
                                           " and((fTYPE = 'R4' and fCUR = '000' and fCURSUM = '3415016.00' and fDBCR = 'C' and fSPEC = '²å³å³Ñáõëï³íáñáõÙ ¹áõñë ·ñáõÙÇó') " &_
                                           " or(fTYPE = 'R5' and fCUR = '001' and fCURSUM = '8537.54'  and fDBCR = 'D' and fSPEC = '¸»µÇïáñ³Ï³Ý å³ñïùÇ ¹áõñë·ñáõÙ')) "
                  sqlValue = 2
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
                  'HIRREST
                  queryString = " SELECT COUNT(*) FROM HIRREST WHERE fOBJECT = " & fISN & _
                                            " and ((fLASTREM = '0.00' and fPENULTREM = '800000.00' and fSTARTREM = '0.00' and fTYPE = 'R4' )" &_
                                            " or (fLASTREM = '8537.54' and fPENULTREM = '0.00' and fSTARTREM = '0.00' and fTYPE = 'R5' ))"
                  sqlValue = 2
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
       ' Ռիսկի դասիչ և պահուստավորման տոկոս գործողության կատարում
      wDate = "150420"
      risk = "01"
      perc = "1"
      Call FillDoc_Risk_Classifier(wDate, risk, perc)

      paramN = c_ViewEdit & "|" & c_Risking & "|" & c_RisksPersRes
      colN = 0
      Call GetfISNFromActionsView(paramN, wDate, wDate, wFrmPttel, colN, cliCode, riskClassfISN2 )
      Log.Message("Ռիսկի դասիչ և պահուստավորման տոկոս փաստաթղթի ISN` " & riskClassfISN2)

                 'HIF
                  queryString = " SELECT COUNT(*) FROM HIF WHERE fBASE = " & riskClassfISN2 & _
                                            " and fCURSUM = '0.00' and fTYPE = 'N0' and fCURSUM = '0.00' " & _
                                            " and ((fSUM = '1.00' and fOP = 'PRS')" & _
                                            " or (fSUM = '0.00' and fOP = 'RSK' and fSPEC = '01')) " 
                  sqlValue = 2
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
                  'DOCS
                  queryString = " Select COUNT(*) from DOCS where fISN = " & fISN & " and fNAME = 'BRBill' and fBODY ='" & vbCRLF _
                                        & "CODE:000329601" & vbCRLF _
                                        & "CLICOD:00000104" & vbCRLF _
                                        & "ACCTYPE:1" & vbCRLF _
                                        & "ACCAGR:000329601" & vbCRLF _
                                        & "NAME:Î³ó³Ï Ø³ó³ÏÛ³Ý" & vbCRLF _
                                        & "CURRENCY:001" & vbCRLF _
                                        & "ACCACC:000329601" & vbCRLF _
                                        & "DATE:20200101" & vbCRLF _
                                        & "ACSBRANCH:01" & vbCRLF _
                                        & "ACSDEPART:2" & vbCRLF _
                                        & "ACSTYPE:BR0" & vbCRLF _
                                        & "JURSTAT:21" & vbCRLF _
                                        & "VOLORT:7" & vbCRLF _
                                        & "PETBUJ:2" & vbCRLF _
                                        & "REZ:1" & vbCRLF _
                                        & "RELBANK:0" & vbCRLF _
                                        & "RABBANK:0" & vbCRLF _
                                        & "AUTODEBT:1" & vbCRLF _
                                        & "DEBTJPART1:2" & vbCRLF _
                                        & "DEBTJPART:0" & vbCRLF _
                                        & "ACCCONNMODE:3" & vbCRLF _
                                        & "ACCCONNSCH:001" & vbCRLF _
                                        & "AUTODATECHILD:1" & vbCRLF _
                                        & "TYPEAUTODATE:1" & vbCRLF _
                                        & "FIXEDDAYS:15" & vbCRLF _
                                        & "PASSOVDIRECTION:3" & vbCRLF _
                                        & "PASSOVTYPE:0" & vbCRLF _
                                        & "SHOWINSTAT:0" & vbCRLF _
                                        & "BRTYPE:9" & vbCRLF _
                                        & "SECTOR:C" & vbCRLF _
                                        & "AIM:00" & vbCRLF _
                                        & "PERRES:1" & vbCRLF _
                                        & "REPCODE:10102" & vbCRLF _
                                        & "NOTCLASS:0" & vbCRLF _
                                        & "SUBJRISK:1" & vbCRLF _
                                        & "ISNBOUT:0"& vbCRLF _
                                        & "FILLACCS:0" & vbCRLF _
                                        & "OPENACCS:0" & vbCRLF _
                                        & "ACCOUTAGR:8000100/000008" & vbCRLF _
                                        & "'"

                  sqlValue = 1
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
      
      ' Խմբային Դուրս գրածի վերականգնում գործողության կատարում
      wDate = "010520"
      wOut = 0
      wInc = 1
      Call GroupCalculation(wDate, wDate, wDbt, wRes, wOut, wInc, wCls, wRsk)
      
      ' Դուրս գրածի վերականգնում փաստաթղթի ISN-ի ստացում գործողությունների դիտում թղթապանակից 
      paramN = c_OpersView
      colN = 5
      docTypeName = "¸»µÇïáñ³Ï³Ý å³ñïùÇ í»ñ³Ï³Ý·ÝáõÙ"
      Call GetfISNFromActionsView(paramN, wDate, wDate, wFrmPttel, colN, docTypeName, OutISN )
      log.Message("Խմբային Դուրս գրում փաստաթղթի ISN` " & OutISN)
      
                  'HI
                  queryString = " SELECT COUNT(*) FROM HI WHERE fBASE = " & OutISN &  " and fOP = 'MSC' and fSUM = '3415016.00'"  &_
                                            " and ((fTYPE = '01' and fCUR = '001' and fCURSUM = '8537.54' and fDBCR = 'D') " &_
                                            " or (fTYPE = '01' and fCUR = '000' and fCURSUM = '3415016.00' and fDBCR = 'C') " &_
                                            " or (fTYPE = '02' and fCUR = '001' and fCURSUM = '8537.54'  and fDBCR = 'C'))"
                  sqlValue = 3
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If

                  'HIR
                  queryString = " SELECT COUNT(*)  FROM HIR WHERE fBASE = " & OutISN & "and fOP = 'INC' " &_
                                           " and ((fTYPE = 'R4' and fCUR = '000' and fCURSUM = '3415016.00' and fDBCR = 'D' and fSPEC = 'ä³Ñáõëï³íáñáõÙ í»ñ³Ï³Ý·ÝáõÙÇó')" &_
                                           " or (fTYPE = 'R5' and fCUR = '001' and fCURSUM = '8537.54' and fDBCR = 'C' and fSPEC = '¸»µÇïáñ³Ï³Ý å³ñïùÇ í»ñ³Ï³Ý·ÝáõÙ')) "
                  sqlValue = 2
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
                  'HIRREST
                  queryString = " SELECT COUNT(*) FROM HIRREST WHERE fOBJECT = " & fISN & _
                                            " and ((fLASTREM = '3415016.00' and fPENULTREM = '0.00' and fSTARTREM = '0.00' and fTYPE = 'R4' )" &_
                                            " or (fLASTREM = '0.00' and fPENULTREM = '8537.54' and fSTARTREM = '0.00' and fTYPE = 'R5' ))"
                  sqlValue = 2
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If

      ' Ապապահուստավորում գործողության կատարում
      wDate = "150520"
      wInc = 0
      wRes = 1
      Call GroupCalculation(wDate, wDate, wDbt, wRes, wOut, wInc, wCls, wRsk)
      
      ' Ապապահուստավորում փաստաթղթի ISN-ի ստացում գործողությունների դիտում թղթապանակից 
      paramN = c_OpersView
      colN = 5
      docTypeName = "²å³å³Ñáõëï³íáñáõÙ"
      Call GetfISNFromActionsView(paramN, wDate, wDate, wFrmPttel, colN, docTypeName, reservISN )
      log.Message("Մարում փաստաթղթի ISN` " & reservISN)
      
                  'HI
                  queryString = " SELECT COUNT(*) FROM HI WHERE fBASE = " & reservISN & _
                                            " and fSUM = '3380865.80'  and fOP = 'RST'  and fCUR = '000' " & _
                                            " and fTYPE = '01' and fCURSUM = '3380865.80' and (fDBCR = 'D' or  fDBCR = 'C') " 
                  sqlValue = 2
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
                  'HIR
                  queryString = " SELECT COUNT(*)  FROM HIR WHERE fBASE  = " & reservISN & " and fOP = 'UNR' "& _
                                            " and fTYPE = 'R4'  and fCUR = '000' and fCURSUM = '3380865.80' and fDBCR = 'C' "
                  sqlValue = 1
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
                  'HIRREST
                  queryString = " SELECT COUNT(*) FROM HIRREST WHERE fOBJECT  = " & fISN &  " and fSTARTREM = '0.00' "& _
                                            " and ((fTYPE = 'R4'  and fLASTREM = '34150.20' and fPENULTREM = '3415016.00') "& _
                                            " or (fTYPE = 'R5'  and fLASTREM = '0.00' and fPENULTREM = '8537.54' ) "& _
                                            " or (fTYPE = 'RI'  and fLASTREM = '9037.54' and fPENULTREM = '0.00')) "
                  sqlValue = 3
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If

      ' Պահուստավորում գործողության կատարում
      wDate = "010620"
      action = c_Store
      state = False
      sumRes = "500000"
      sumUnres = "0.00"
      wCommentP = "ä³Ñáõëï³íáñáõÙ"
      Call ProvisionAction(action, storeISN, wDate, state, wSumma, sumAgr, sumRes, sumUnres, wCommentP, acsBranch, acsSect)
      Log.Message("Պահուստավորում փաստաթղթի ISN` " & storeISN)
      
                  'HI
                  queryString = " SELECT COUNT(*) FROM HI WHERE fBASE = " & storeISN & _
                                            " and fSUM = '500000.00' and fOP = 'RST'  and fCUR = '000' " & _
                                            " and fTYPE = '01' and fCURSUM = '500000.00' and (fDBCR = 'D' or  fDBCR = 'C') " 
                  sqlValue = 2
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
                  'HIR
                  queryString = " SELECT COUNT(*) FROM HIR WHERE fBASE = " & storeISN & " and fOP = 'RES' "& _
                                            " and fTYPE = 'R4'  and fCUR = '000' and fCURSUM = '500000.00' and fDBCR = 'D'"
                  sqlValue = 1
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
                  'HIRREST
                  queryString = " SELECT COUNT(*) FROM HIRREST WHERE fOBJECT = " & fISN & " and fSTARTREM = '0.00'"& _
                                            " and ((fTYPE = 'R4'  and fLASTREM = '534150.20' and fPENULTREM = '34150.20') "& _
                                            " or (fTYPE = 'R5'  and fLASTREM = '0.00' and fPENULTREM = '8537.54' ) "& _
                                            " or (fTYPE = 'RI'  and fLASTREM = '9037.54' and fPENULTREM = '0.00')) "
                  sqlValue = 3
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
      ' Դուրս գրում գործողության կատարում
      action = c_WriteOff
      sumAgr = "1000"
      Call WriteOffOnAction(action, writeOffISN, wDate, sumAgr, wComment, acsBranch, acsSect)
      Log.Message("Դուրս գրում փաստաթղթի ISN` " & writeOffISN)
      
                  'HI
                  queryString = " SELECT COUNT(*) FROM HI WHERE fBASE = " & writeOffISN & " and fSUM = '400000.00' and fOP = 'MSC'  "& _
                                            " and ((fDBCR = 'C' and fCUR = '001' and fTYPE = '01' and fCURSUM = '1000.00') " & _
                                            " or (fDBCR = 'D'and fCUR = '000' and fTYPE = '01' and fCURSUM = '400000.00')" & _
                                            " or (fDBCR = 'D' and fCUR = '001' and fTYPE = '02' and fCURSUM = '1000.00')) " 
                  sqlValue = 3
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If

                  'HIR
                  queryString = " SELECT COUNT(*) FROM HIR WHERE fBASE = " & writeOffISN & " and fOP = 'OUT' " &_
                                           " and ((fTYPE = 'R4'  and fCUR = '000' and fCURSUM = '400000.00' and fDBCR = 'C')" &_
                                           " or (fTYPE = 'R5'  and fCUR = '001' and fCURSUM = '1000.00' and fDBCR = 'D'))"
                  sqlValue = 2
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
                  'HIRREST
                  queryString = " SELECT COUNT(*) FROM HIRREST WHERE fOBJECT = " & fISN & " and fSTARTREM = '0.00'"  &_
                                            " and ((fTYPE = 'R4'  and fLASTREM = '134150.20' and fPENULTREM = '34150.20')" &_
                                            " or (fTYPE = 'R5'  and fLASTREM = '1000.00' and fPENULTREM = '0.00' ) " &_
                                            " or (fTYPE = 'RI'  and fLASTREM = '9037.54' and fPENULTREM = '0.00'))"
                  sqlValue = 3
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
      
      ' Պարտքերի խմբային մարում
      wDate = "150620"
      wDbt = 1
      wRes = 0
      Call GroupCalculation(wDate, wDate, wDbt, wRes, wOut, wInc, wCls, wRsk)
      
      ' Մարում փաստաթղթի ISN-ի ստացում գործողությունների դիտում թղթապանակից 
      paramN = c_OpersView
      colN = 5
      docTypeName = "¸»µÇïáñ³Ï³Ý å³ñïùÇ Ù³ñáõÙ"
      Call GetfISNFromActionsView(paramN, wDate, wDate, wFrmPttel, colN, docTypeName, dbtISN )
      log.Message("Մարում փաստաթղթի ISN` " & dbtISN)
      
                 'HI
                  queryString = " SELECT COUNT(*) FROM HI WHERE fBASE = " & dbtISN & " and fOP = 'MSC'  "& _
                                            " and (( fTYPE = '01' and fSUM = '400000.00' and fDBCR = 'D' and fCUR = '001'  and fCURSUM = '1000.00') " & _
                                            " or ( fTYPE = '01' and fSUM = '3815016.00' and fDBCR = 'D' and fCUR = '001'  and fCURSUM = '9537.54') " & _
                                            " or ( fTYPE = '01' and fSUM = '3815016.00' and fDBCR = 'C' and fCUR = '001'  and fCURSUM = '9537.54') " & _
                                            " or ( fTYPE = '01' and fSUM = '400000.00' and fDBCR = 'C' and fCUR = '000'  and fCURSUM = '400000.00') " & _
                                            " or ( fTYPE = '02' and fSUM = '400000.00' and fDBCR = 'C' and fCUR = '001'  and fCURSUM = '1000.00')) "
                  sqlValue = 5
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If

                  'HIR
                  queryString = " SELECT COUNT(*) FROM HIR WHERE fBASE = " & dbtISN &_
                                           " and ((fTYPE = 'R4'  and fCUR = '000' and fCURSUM = '400000.00' and fDBCR = 'D' and fOP = 'INC')" &_
                                           " or (fTYPE = 'R5'  and fCUR = '001' and fCURSUM = '1000.00' and fDBCR = 'C' and fOP = 'INC')" &_
                                           " or (fTYPE = 'RI'  and fCUR = '001' and fCURSUM = '9537.54' and fDBCR = 'D' and fOP = 'DBT'))"
                  sqlValue = 3
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
                  'HIRREST
                  queryString = " SELECT COUNT(*) FROM HIRREST WHERE fOBJECT = " & fISN & " and fSTARTREM = '0.00'"  &_
                                            " and ((fTYPE = 'R4'  and fLASTREM = '534150.20' and fPENULTREM = '134150.20')" &_
                                            " or (fTYPE = 'R5'  and fLASTREM = '0.00' and fPENULTREM = '1000.00' ) " &_
                                            " or (fTYPE = 'RI'  and fLASTREM = '18575.08' and fPENULTREM = '9037.54'))"
                  sqlValue = 3
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If

      ' Ապապահուստավորում գործողության կատարում
      wDate = "010720"
      action = c_Store
      sumRes = "0.00"
      sumUnres = "534,150.20"
      wComment = "²å³å³Ñáõëï³íáñáõÙ"
      Call ProvisionAction(action, storeISN, wDate, state, wSumma, sumAgr, sumRes, sumUnres, wComment, acsBranch, acsSect)
      Log.Message("Պահուստավորում փաստաթղթի ISN` " & storeISN)

                  'HI
                  queryString = " SELECT COUNT(*) FROM HI WHERE fBASE = " & storeISN & " and fSUM = '534150.20' and fOP = 'RST' "&_
                                            " and fTYPE = '01' and fCURSUM = '534150.20' and (fDBCR = 'D' or  fDBCR = 'C') " 
                  sqlValue = 2
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If

                  'HIR
                  queryString = " SELECT COUNT(*) FROM HIR WHERE fBASE = " & storeISN & " and fOP = 'UNR' " &_
                                           " and fTYPE = 'R4'  and fCUR = '000' and fCURSUM = '534150.20' and fDBCR = 'C' " 

                  sqlValue = 1
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
                  'HIRREST
                  queryString = " SELECT COUNT(*) FROM HIRREST WHERE fOBJECT = " & fISN & " and fSTARTREM = '0.00' " &_
                                            " and ((fTYPE = 'R4'  and fLASTREM = '0.00' and fPENULTREM = '534150.20')" &_
                                            " or (fTYPE = 'R5'  and fLASTREM = '0.00' and fPENULTREM = '1000.00' )" &_
                                            " or (fTYPE = 'RI'  and fLASTREM = '18575.08' and fPENULTREM = '9037.54'))"
                  sqlValue = 3
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
      ' Պայմանագրի փակում
      wDate = "020720"
      wDbt = 0
      wCls = 1
      Call GroupCalculation(wDate, wDate, wDbt, wRes, wOut, wInc, wCls, wRsk)
      
      BuiltIn.Delay(1000)
      Call Close_Pttel("frmPttel")
      BuiltIn.Delay(4000)

      ' Մուտք Հաճախորդներ թղթապանակ 
      Call wTreeView.DblClickItem("|¸»µÇïáñ³Ï³Ý å³ñïù»ñ|ä³ÛÙ³Ý³·ñ»ñ")
      BuiltIn.Delay(1000)
      
      If Not Sys.Process("Asbank").VBObject("frmAsUstPar").Exists Then
            Log.Error(folderName & " դիալոգը չի բացվել")
      End If
      
      ' Հաճախորդ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "ACC", cliCode)
      ' Ցույց տալ փակվածները
      Call Rekvizit_Fill("Dialog", 2, "CheckBox", "CLOSE", 1)
      
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
      BuiltIn.Delay(5000)
      
       If  frmPttel.VBObject("TDBGView").ApproxCount <> 1 Then
             Log.Error(" Դեբիտորական պարտք պայմանագրիը առկա չէ Պայմանագրեր թղթապանակում")
       End If

      dateCl = wMDIClient.VBObject("frmPttel").GetColumnIndex("fDATECLOSE")
      wCenceled = wMDIClient.VBObject("frmPttel").GetColumnIndex("fCANCELED")

      If  Not (Trim( tdbgView.Columns.Item(dateCl).Value) = Trim("02/07/20")  and  Trim( tdbgView.Columns.Item(wCenceled).Value) = Trim("02/07/20")) Then
            Log.Error("Պայմանագիրը չի փակվել")
      End If
      
      ' Պայմանագրի բացում
      Call wMainForm.MainMenu.Click(c_AllActions)
      Call wMainForm.PopupMenu.Click(c_AgrOpen)
      Call ClickCmdButton(5, "²Ûá")
      BuiltIn.Delay(2000)
            
      If  Not (Trim( tdbgView.Columns.Item(dateCl).Value) = Trim("")  and  Trim( tdbgView.Columns.Item(wCenceled).Value) = Trim("")) Then
            Log.Error("Պայմանագիրը չի բացվել")
      End If
      
      ' Կատարած գործողությունների ջնջում
      sDatePar = "START"
      eDatePar = "END"
      action = c_Delete
      dateGive = "010120"
      param = c_OpersView
      state = True
      Call DeleteFromAllActionDoc(param, sDatePar, dateGive, eDatePar, wDate, state, action)
      
      Call Close_Pttel("frmPttel_2")
      BuiltIn.Delay(2000)
      
      ' Ռիսկի դասիչ և պահուստավորման փաստաթղթերի ջնջում
      param = c_ViewEdit & "|" & c_Risking & "|" & c_RisksPersRes
      Call DeleteFromAllActionDoc(param, sDatePar, dateGive, eDatePar, wDate, state, action)
      
      Call Close_Pttel("frmPttel_2")
      BuiltIn.Delay(2000)

      ' Ջնջել պայմանագիրը                
      Call DelDoc()
      Call Close_Pttel("frmPttel")
      
      ' Փակել ՀԾ-Բանկ ծրագիրը
      Call Close_AsBank()
                 
End Sub