Option Explicit
'USEUNIT Subsystems_SQL_Library
'USEUNIT  Library_Common
'USEUNIT Constants
'USEUNIT External_Transfers_Library
'USEUNIT Subsystems_SQL_Library
'USEUNIT Percentage_Calculation_Filter_Library
'USEUNIT BankMail_Library
'USEUNIT Akreditiv_Library

'Test Case 159958

' "Վճարման պահանջագիր" փաստաթղթի ստեղծում և գործողությունների կատարում
Sub Payment_Request_Test()

      Dim fDATE, sDATE
      Dim PaymentRequestDoc, SentToClearingDoc, PartiallyEditableAssign, AccForTransfers
      Dim folderDirect, wCur, wUser, docType, paySysin, paySysOut, payNotes, acsBranch, acsDepart, selectView, exportExcel
      Dim  workEnv, wStatus, isnRekName, docISN, status, colNum, todayDMY, state, frmPttel
      Dim queryString, sqlValue, sql_isEqual, todayTime
      
      fDATE = "20250101"
      sDATE = "20120101"
      Call Initialize_AsBank("bank", sDATE, fDATE)
      
      ' Մուտք գործել համակարգ ARMSOFT օգտագործողով 
      Call Create_Connection()
      Login("ARMSOFT")
      
      ' Մուտք Արտաքին փոխանցումների ԱՇՏ
      Call ChangeWorkspace(c_ExternalTransfers)

      todayDMY = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"%d/%m/%y")
      
      ' Ստեղծել Վճարման պահանջագիր
      Set PaymentRequestDoc = New_PaymentRequestDoc()
      With PaymentRequestDoc
            .folderDirect = "|²ñï³ùÇÝ ÷áË³ÝóáõÙÝ»ñÇ ²Þî|Üáñ ÷³ëï³ÃÕÃ»ñ|ì×³ñÙ³Ý å³Ñ³Ýç³·Çñ"
            .acsBranch = "00"
            .acsDepart = "1"
            .wDate = todayDMY
            .accDb = "10300/4200012"
            .wPayer = "¾ÏáÝáÙ.Ý³Ë³ñ.Ï»Ýïñ.Ñ³ß."
            .accCr = "10300/4200046"
            .wReceiver = "Î»ÝïñáÝ.·³ÝÓ³å.·áñÍ³éÝ³Ï.µ³Å³Ý."
            .wSumma = "10,000"
            .wCur = "001"
            .wAim = "Ð³ñÏ»ñÇ Ù³ñáõÙ"
            .wPack = "123"
      End With
      
      Call Fill_PaymentRequestDoc(PaymentRequestDoc)
      
      Log.Message(PaymentRequestDoc.fISN)
      Log.Message(PaymentRequestDoc.cliCode)
      
      todayTime = aqConvert.DateTimeToFormatStr(aqDateTime.Today,"20%y%m%d")
                  'DOCS
                  queryString = " SELECT COUNT(*) FROM DOCS WHERE fISN = " & PaymentRequestDoc.fISN & _
                                            " and fNAME = 'PayReq  ' and fSTATE = '1' and fBODY = '"& vbCRLF _
                                        & "USERID:  77"& vbCRLF _
                                        & "ACSBRANCH:00"& vbCRLF _
                                        & "ACSDEPART:1"& vbCRLF _
                                        & "DOCNUM:"& PaymentRequestDoc.cliCode & vbCRLF _
                                        & "DATE:"& todayTime & vbCRLF _
                                        & "ACCDB:103004200012"& vbCRLF _
                                        & "PAYER:¾ÏáÝáÙ.Ý³Ë³ñ.Ï»Ýïñ.Ñ³ß."& vbCRLF _
                                        & "ACCCR:103004200046"& vbCRLF _
                                        & "RECEIVER:Î»ÝïñáÝ.·³ÝÓ³å.·áñÍ³éÝ³Ï.µ³Å³Ý."& vbCRLF _
                                        & "SUMMA:10000"& vbCRLF _
                                        & "CUR:001"& vbCRLF _
                                        & "AIM:Ð³ñÏ»ñÇ Ù³ñáõÙ"& vbCRLF _
                                        & "PACK:123"& vbCRLF _
                                        & "PAYSYSIN:Ð"& vbCRLF _
                                        & "TYPECODE1:-10 20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28"& vbCRLF _
                                        & "TYPECODE2:-10 20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28"& vbCRLF _
                                        & "'"
                  sqlValue = 1
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
                  'FOLDERS
                  queryString = " SELECT COUNT(*) FROM FOLDERS WHERE fISN = " & PaymentRequestDoc.fISN & _
                                            " and fNAME = 'PayReq  ' and fCOM = 'ì×³ñÙ³Ý å³Ñ³Ýç³·Çñ' and fECOM = 'Payment on Request'" & _
                                            " and fDCBRANCH = '00' and fDCDEPART = '1' and fSPEC ='"& PaymentRequestDoc.cliCode &"103004200012    103004200046            10000.00001Üáñ                                                   77                                                                                       Ð        Ð³ñÏ»ñÇ Ù³ñáõÙ'" 
                  sqlValue = 1
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                
                  
                  'DOCLOG
                  queryString = " SELECT COUNT(*) FROM DOCLOG WHERE fISN = " & PaymentRequestDoc.fISN & _
                                            " and fSTATE = '1' and fOP = 'N' " 
                  sqlValue = 1
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
      
      ' Մուտք աշխատանքային փաստաթղթեր դիալոգ և արժեքների լրացում
      folderDirect = "|²ñï³ùÇÝ ÷áË³ÝóáõÙÝ»ñÇ ²Þî|²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ"
      wCur = "001"
      wUser = "77"
      docType = "PayReq"
      paySysin = "Ð"
      acsBranch = "00"
      acsDepart = "1" 
      selectView = "Oper"
      exportExcel = "0"
      Call WorkingDocsFilter(folderDirect, todayDMY, todayDMY, wCur, wUser, docType, paySysin, paySysOut, payNotes, acsBranch, acsDepart, selectView, exportExcel)
      
      ' Փաստաթուղթն ուղարկել արտաքին բաժին
      state =  ConfirmContractDoc(2, PaymentRequestDoc.cliCode, c_SendToExternalSec, 5, "²Ûá")
      If Not state Then
            Log.Error("Փաստաթուղթը չի ուղարկվել արտաքին բաժին")
      End If
      BuiltIn.Delay(1000)
      
                  'DOCS
                  queryString = " SELECT COUNT(*) FROM DOCS WHERE fISN = " & PaymentRequestDoc.fISN & _
                                            " and fNAME = 'PayReq  ' and fSTATE = '2' "
                  sqlValue = 1
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
                  'FOLDERS
                  queryString = " SELECT COUNT(*) FROM FOLDERS WHERE fISN = " & PaymentRequestDoc.fISN & _
                                            " and fNAME = 'PayReq  ' and fCOM = 'ì×³ñÙ³Ý å³Ñ³Ýç³·Çñ' and fECOM = 'Payment on Request'" & _
                                            " and fDCBRANCH = '00' and fDCDEPART = '1' and fSPEC ='"& PaymentRequestDoc.cliCode &"103004200012    103004200046            10000.00001Ð³ëï³ïí³Í             77Ð³ñÏ»ñÇ Ù³ñáõÙ                  ¾ÏáÝáÙ.Ý³Ë³ñ.Ï»Ýïñ.Ñ³ß.         Î»ÝïñáÝ.·³ÝÓ³å.·áñÍ³éÝ³Ï.µ³Å³Ý. 123                                               Ð '" 
                  sqlValue = 1
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
                  'DOCLOG
                  queryString = " SELECT COUNT(*) FROM DOCLOG WHERE fISN = " & PaymentRequestDoc.fISN & _
                                            " and ((fSTATE = '1' and fOP = 'N') or (fSTATE = '2' and fOP = 'E'))" 
                  sqlValue = 2
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
                  'PAYMENTS
                  queryString = " SELECT COUNT(*) FROM PAYMENTS WHERE fISN = " & PaymentRequestDoc.fISN & _
                                            " and fDOCTYPE = 'PayReq  ' and fCUR = '001' and fSUMMA = '10000.00'" & _
                                            " and fSUMMAAMD = '4000000.00' and fSUMMAUSD = '10000.00' and fCOM = 'Ð³ñÏ»ñÇ Ù³ñáõÙ'" & _
                                            " and fCHARGESUM = '0.00' and fCHARGESUMAMD = '0.00' and fCHARGESUM2 = '0.00' and fCHARGESUMAMD2 = '0.00'" 
                  sqlValue = 1
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
      Set frmPttel = Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel")
      frmPttel.Close
      
      ' Մուտք Ուղարկվող հանձնարարագրեր թղթապանակ
      folderDirect = "|²ñï³ùÇÝ ÷áË³ÝóáõÙÝ»ñÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|àõÕ³ñÏíáÕ Ñ³ÝÓÝ³ñ³ñ³·ñ»ñ|àõÕ³ñÏíáÕ Ñ³ÝÓÝ³ñ³ñ³·ñ»ñ"
      workEnv = "Ուղարկվող հանձնարարագրեր թղթապանակ"
      wStatus = False
      
      state = AccessFolder(folderDirect, workEnv, "PERN", todayDMY, "PERK", todayDMY, wStatus, isnRekName, docISN)
      
      If Not state Then
            Log.Error("Մուտք Ուղարկվող հանձնարարագրեր թղթապանակ ձախողվել է")
            Exit Sub
      End If
      
      ' Կատարել "Ուղարկել թղթային" գործողությունը
      state =  ConfirmContractDoc(2, PaymentRequestDoc.cliCode, c_Send2Cl, 5, "²Ûá")
      If Not state Then
            Log.Error("'Ուղարկել թղթային' գործողությունը չի կատարվել")
      End If
      BuiltIn.Delay(1500)
      
                  'DOCS
                  queryString = " SELECT COUNT(*) FROM DOCS WHERE fISN = " & PaymentRequestDoc.fISN & _
                                            " and fNAME = 'PayReq  ' and fSTATE = '5' "
                  sqlValue = 1
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
                  'FOLDERS
                  queryString = " SELECT COUNT(*) FROM FOLDERS WHERE fISN = " & PaymentRequestDoc.fISN & _
                                            " and fNAME = 'PayReq  ' and fCOM = 'ì×³ñÙ³Ý å³Ñ³Ýç³·Çñ' and fDCBRANCH = '00' and fDCDEPART = '1'" & _
                                            " and fSPEC ='"& PaymentRequestDoc.cliCode &"103004200012    103004200046            10000.00001àõÕ³ñÏí³Í             77Ð³ñÏ»ñÇ Ù³ñáõÙ                  ¾ÏáÝáÙ.Ý³Ë³ñ.Ï»Ýïñ.Ñ³ß.         Î»ÝïñáÝ.·³ÝÓ³å.·áñÍ³éÝ³Ï.µ³Å³Ý. 123                    00000000Ð'" 
                  sqlValue = 1
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
                  'DOCLOG
                  queryString = " SELECT COUNT(*) FROM DOCLOG WHERE fISN = " & PaymentRequestDoc.fISN & _
                                            " and ((fSTATE = '1' and fOP = 'N' ) "& _
                                            " or (fSTATE = '2' and fOP = 'E' )"& _
                                            " or (fSTATE = '5' and fOP = 'M' and fCOM = 'àõÕ³ñÏí»É ¿ ÃÕÃ³ÛÇÝ'))" 
                  sqlValue = 3
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
                  'PAYMENTS
                  queryString = " SELECT COUNT(*) FROM PAYMENTS WHERE fISN = " & PaymentRequestDoc.fISN & _
                                            " and fDOCTYPE = 'PayReq  ' and fCUR = '001' and fSUMMA = '10000.00' and fSTATE = '5'" & _
                                            " and fSUMMAAMD = '4000000.00' and fSUMMAUSD = '10000.00' and fCOM = 'Ð³ñÏ»ñÇ Ù³ñáõÙ'" & _
                                            " and fCHARGESUM = '0.00' and fCHARGESUMAMD = '0.00' and fCHARGESUM2 = '0.00' and fCHARGESUMAMD2 = '0.00'" & _
                                            " and fPAYER = '¾ÏáÝáÙ.Ý³Ë³ñ.Ï»Ýïñ.Ñ³ß.' and fRECEIVER = 'Î»ÝïñáÝ.·³ÝÓ³å.·áñÍ³éÝ³Ï.µ³Å³Ý.'"
                  sqlValue = 1
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
      frmPttel.Close
      
      ' Մուտք Ուղարկված թղթային թղթապանակ
      Set SentToClearingDoc = New_SentToClearingDoc()
      With SentToClearingDoc
            .folderDirect = "|²ñï³ùÇÝ ÷áË³ÝóáõÙÝ»ñÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|àõÕ³ñÏí³Í  Ñ³ÝÓÝ³ñ³ñ³·ñ»ñ|àõÕ³ñÏí³Í ÃÕÃ."
            .stDate = todayDMY
            .eDate = todayDMY
            .pysysIn = "Ð"
            .wCur = "001"
            .acsBranch = "00"
            .acsDepart = "1"
      End With
      
      Call Fill_SentToClearingDoc(SentToClearingDoc)
      
      ' Ուղարկել մասնակի խմբագրման
      state = ConfirmContractDoc(1, PaymentRequestDoc.cliCode, c_SendToPartEd, 2, "Î³ï³ñ»É")
      If Not state Then
            Log.Error("Փաստաթուղթը չի ուղարկվել մասնակի խմբագրման")
      End If
      BuiltIn.Delay(1500)
      
                  'DOCS
                  queryString = " SELECT COUNT(*) FROM DOCS WHERE fISN = " & PaymentRequestDoc.fISN & _
                                            " and fNAME = 'PayReq  ' and fSTATE = '12' "
                  sqlValue = 1
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
                  'FOLDERS
                  queryString = " SELECT COUNT(*) FROM FOLDERS WHERE fISN = " & PaymentRequestDoc.fISN & _
                                            " and fNAME = 'PayReq  ' and fCOM = 'ì×³ñÙ³Ý å³Ñ³Ýç³·Çñ' and fDCBRANCH = '00' and fDCDEPART = '1' and fECOM = 'Payment on Request'" & _
                                            " and fSPEC ='"& PaymentRequestDoc.cliCode &"103004200012    103004200046            10000.00001ÎñÏÝ³ÏÇ áõÕ³ñÏíáÕ     77Ð³ñÏ»ñÇ Ù³ñáõÙ                  ¾ÏáÝáÙ.Ý³Ë³ñ.Ï»Ýïñ.Ñ³ß.         Î»ÝïñáÝ.·³ÝÓ³å.·áñÍ³éÝ³Ï.µ³Å³Ý. 123Ð1'" 
                  sqlValue = 1
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If

                  
                  'DOCLOG
                  queryString = " SELECT COUNT(*) FROM DOCLOG WHERE fISN = " & PaymentRequestDoc.fISN & _
                                            " and ((fSTATE = '1' and fOP = 'N' ) " & _
                                            " or (fSTATE = '2' and fOP = 'E' )" & _
                                            " or (fSTATE = '5' and fOP = 'M' and fCOM = 'àõÕ³ñÏí»É ¿ ÃÕÃ³ÛÇÝ')" & _
                                            " or (fSTATE = '12' and fOP = 'M' and fCOM = 'àõÕ³ñÏí»É ¿ Ù³ëÝ³ÏÇ ËÙµ³·ñÙ³Ý'))" 
                  sqlValue = 4
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If

      frmPttel.Close
      
      ' Մուտք Մասնակի խմբագրվող հանձնարարագրեր
      Set PartiallyEditableAssign = New_PartiallyEditableAssign()
      With PartiallyEditableAssign
              .folderDirect = "|²ñï³ùÇÝ ÷áË³ÝóáõÙÝ»ñÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|àõÕ³ñÏíáÕ Ñ³ÝÓÝ³ñ³ñ³·ñ»ñ|Ø³ëÝ³ÏÇ ËÙµ³·ñíáÕ Ñ³ÝÓÝ³ñ³ñ³·ñ»ñ"
              .stDate = todayDMY
              .eDate = todayDMY
              .paySysin =  "Ð"
              .acsBranch = "00"
              .acsDepart = "1"
              .selectView = "RePay"
              .exportExcel = "0"
      End With
      
      Call Fill_PartiallyEditableAssign(PartiallyEditableAssign)
      
      ' Ստուգել փաստաթղթի առկայությունը "Մասնակի խմբագրվող հանձնարարագրեր" թղթապանակում
      status = CheckContractDoc(1, PaymentRequestDoc.cliCode)
      
      If Not status Then
            Log.Error()
      End If

      ' Կատարել"Խմբագրել" գործողությունը
      Call ContractAction(c_ToEdit)
      
      ' Խմբագրել "Վճարող" դաշտը
      Call Rekvizit_Fill("Document", 1, "General", "PAYER", "^A[Del]" & "Ð³Û³·ñáµ³ÝÏ")
      Call ClickCmdButton(1, "Î³ï³ñ»É")
      BuiltIn.Delay(1500)
      
      colNum =	wMDIClient.VBObject("frmPttel").GetColumnIndex("PAYER")
      If Not Trim(frmPttel.VBObject("TDBGView").Columns.Item(colNum)) = "Ð³Û³·ñáµ³ÝÏ" Then
            Log.Error("Վճարող դաշտը չի խմբագրվել")
      End If
      
                  'DOCS
                  queryString = " SELECT COUNT(*) FROM DOCS WHERE fISN = " & PaymentRequestDoc.fISN & _
                                            " and fNAME = 'PayReq  ' and fSTATE = '12' and fBODY ='"& vbCRLF _
                                        & "USERID:  77"& vbCRLF _
                                        & "ACSBRANCH:00"& vbCRLF _
                                        & "ACSDEPART:1"& vbCRLF _
                                        & "DOCNUM:"&PaymentRequestDoc.cliCode & vbCRLF _
                                        & "DATE:"& todayTime & vbCRLF _
                                        & "ACCDB:103004200012"& vbCRLF _
                                        & "PAYER:Ð³Û³·ñáµ³ÝÏ"& vbCRLF _
                                        & "ACCCR:103004200046"& vbCRLF _
                                        & "RECEIVER:Î»ÝïñáÝ.·³ÝÓ³å.·áñÍ³éÝ³Ï.µ³Å³Ý."& vbCRLF _
                                        & "SUMMA:10000"& vbCRLF _
                                        & "CUR:001"& vbCRLF _
                                        & "AIM:Ð³ñÏ»ñÇ Ù³ñáõÙ"& vbCRLF _
                                        & "PACK:123"& vbCRLF _
                                        & "PAYSYSIN:Ð"& vbCRLF _
                                        & "PAYSYSOUT:1"& vbCRLF _
                                        & "TYPECODE1:-10 20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28"& vbCRLF _
                                        & "TYPECODE2:-10 20 21 22 23 24 30 31 32 25 26 92 93 11 27 33 28"& vbCRLF _
                                        & "'"
                  sqlValue = 1
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
                  'FOLDERS
                  queryString = " SELECT COUNT(*) FROM FOLDERS WHERE fISN = " & PaymentRequestDoc.fISN & _
                                            " and fNAME = 'PayReq  ' and fCOM = 'ì×³ñÙ³Ý å³Ñ³Ýç³·Çñ' and fECOM = 'Payment on Request'" & _
                                            " and fDCBRANCH = '00' and fDCDEPART = '1' and fSPEC ='"& PaymentRequestDoc.cliCode &"103004200012    103004200046            10000.00001ÎñÏÝ³ÏÇ áõÕ³ñÏíáÕ     77Ð³ñÏ»ñÇ Ù³ñáõÙ                  Ð³Û³·ñáµ³ÝÏ                     Î»ÝïñáÝ.·³ÝÓ³å.·áñÍ³éÝ³Ï.µ³Å³Ý. 123Ð1'" 
                  sqlValue = 1
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If

                  
                  'DOCLOG
                  queryString = " SELECT COUNT(*) FROM DOCLOG WHERE fISN = " & PaymentRequestDoc.fISN & _
                                            " and ((fSTATE = '1' and fOP = 'N' )" & _
                                            " or (fSTATE = '2' and fOP = 'E' )" & _
                                            " or (fSTATE = '5' and fOP = 'M' and fCOM = 'àõÕ³ñÏí»É ¿ ÃÕÃ³ÛÇÝ')" & _
                                            " or (fSTATE = '12' and fOP = 'M' and fCOM = 'àõÕ³ñÏí»É ¿ Ù³ëÝ³ÏÇ ËÙµ³·ñÙ³Ý')" & _
                                            " or (fSTATE = '12' and fOP = 'E' ))"
                  sqlValue = 5
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
      
      ' Փաստաթուղթն ուղարկել արտաքին բաժին
      state = ConfirmContractDoc(1, PaymentRequestDoc.cliCode, c_SendToExternalSec, 5, "²Ûá")
      If Not state Then
            Log.Error("Փաստաթուղթը չի ուղարկվել արտաքին բաժին")
      End If
      BuiltIn.Delay(1500)
      
                  'DOCS
                  queryString = " SELECT COUNT(*) FROM DOCS WHERE fISN = " & PaymentRequestDoc.fISN & _
                                            " and fNAME = 'PayReq  ' and fSTATE = '2' "
                  sqlValue = 1
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
                  'FOLDERS
                  queryString = " SELECT COUNT(*) FROM FOLDERS WHERE fISN = " & PaymentRequestDoc.fISN & _
                                            " and fNAME = 'PayReq  ' and fCOM = 'ì×³ñÙ³Ý å³Ñ³Ýç³·Çñ' and fECOM = 'Payment on Request'" & _
                                            " and fDCBRANCH = '00' and fDCDEPART = '1' and fSPEC ='"& PaymentRequestDoc.cliCode &"103004200012    103004200046            10000.00001Ð³ëï³ïí³Í             77Ð³ñÏ»ñÇ Ù³ñáõÙ                  Ð³Û³·ñáµ³ÝÏ                     Î»ÝïñáÝ.·³ÝÓ³å.·áñÍ³éÝ³Ï.µ³Å³Ý. 123                                               Ð1'" 
                  sqlValue = 1
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If

                  
                  'DOCLOG
                  queryString = " SELECT COUNT(*) FROM DOCLOG WHERE fISN = " & PaymentRequestDoc.fISN 
                  sqlValue = 6
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
                  'PAYMENTS
                  queryString = " SELECT COUNT(*) FROM PAYMENTS WHERE fISN = " & PaymentRequestDoc.fISN & _
                                            " and fDOCTYPE = 'PayReq  ' and fCUR = '001' and fSUMMA = '10000.00'" & _
                                            " and fSUMMAAMD = '4000000.00' and fSUMMAUSD = '10000.00' and fCOM = 'Ð³ñÏ»ñÇ Ù³ñáõÙ'" & _
                                            " and fCHARGESUM = '0.00' and fCHARGESUMAMD = '0.00' and fCHARGESUM2 = '0.00' and fCHARGESUMAMD2 = '0.00'" &_
                                            "and fPAYER = 'Ð³Û³·ñáµ³ÝÏ' and fRECEIVER = 'Î»ÝïñáÝ.·³ÝÓ³å.·áñÍ³éÝ³Ï.µ³Å³Ý.'"
                  sqlValue = 1
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
      frmPttel.Close
      
      ' Մուտք Ուղարկվող հանձնարարագրեր թղթապանակ
      folderDirect = "|²ñï³ùÇÝ ÷áË³ÝóáõÙÝ»ñÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|àõÕ³ñÏíáÕ Ñ³ÝÓÝ³ñ³ñ³·ñ»ñ|àõÕ³ñÏíáÕ Ñ³ÝÓÝ³ñ³ñ³·ñ»ñ"
      workEnv = "Ուղարկվող հանձնարարագրեր թղթապանակ"
      wStatus = False
      
      state = AccessFolder(folderDirect, workEnv, "PERN", todayDMY, "PERK", todayDMY, wStatus, isnRekName, docISN)
      
      If Not state Then
            Log.Error("Մուտք Ուղարկվող հանձնարարագրեր թղթապանակ ձախողվել է")
            Exit Sub
      End If
      
      ' Կատարել "Ուղարկել թղթային" գործողությունը
      state =  ConfirmContractDoc(2, PaymentRequestDoc.cliCode, c_Send2Cl, 5, "²Ûá")
      If Not state Then
            Log.Error("'Ուղարկել թղթային' գործողությունը չի կատարվել")
      End If
      
      BuiltIn.Delay(1500)
      
                  'DOCS
                  queryString = " SELECT COUNT(*) FROM DOCS WHERE fISN = " & PaymentRequestDoc.fISN & _
                                            " and fNAME = 'PayReq  ' and fSTATE = '5' "
                  sqlValue = 1
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
                  'FOLDERS
                  queryString = " SELECT COUNT(*) FROM FOLDERS WHERE fISN = " & PaymentRequestDoc.fISN & _
                                            " and fNAME = 'PayReq  ' and fCOM = 'ì×³ñÙ³Ý å³Ñ³Ýç³·Çñ' and fDCBRANCH = '00' and fDCDEPART = '1'" & _
                                            " and fSPEC ='"& PaymentRequestDoc.cliCode &"103004200012    103004200046            10000.00001àõÕ³ñÏí³Í             77Ð³ñÏ»ñÇ Ù³ñáõÙ                  Ð³Û³·ñáµ³ÝÏ                     Î»ÝïñáÝ.·³ÝÓ³å.·áñÍ³éÝ³Ï.µ³Å³Ý. 123                    00000000Ð' " 
                  sqlValue = 1
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
                  'DOCLOG
                  queryString = " SELECT COUNT(*) FROM DOCLOG WHERE fISN = " & PaymentRequestDoc.fISN  
                  sqlValue = 7
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
                  'PAYMENTS
                  queryString = " SELECT COUNT(*) FROM PAYMENTS WHERE fISN = " & PaymentRequestDoc.fISN & _
                                            " and fDOCTYPE = 'PayReq  ' and fCUR = '001' and fSUMMA = '10000.00' and fSTATE = '5'" & _
                                            " and fSUMMAAMD = '4000000.00' and fSUMMAUSD = '10000.00' and fCOM = 'Ð³ñÏ»ñÇ Ù³ñáõÙ'" & _
                                            " and fCHARGESUM = '0.00' and fCHARGESUMAMD = '0.00' and fCHARGESUM2 = '0.00' and fCHARGESUMAMD2 = '0.00'" & _
                                            " and fPAYER = 'Ð³Û³·ñáµ³ÝÏ' and fRECEIVER = 'Î»ÝïñáÝ.·³ÝÓ³å.·áñÍ³éÝ³Ï.µ³Å³Ý.'"
                  sqlValue = 1
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
      frmPttel.Close
      
      ' Մուտք Ուղարկված թղթային թղթապանակ
      Set SentToClearingDoc = New_SentToClearingDoc()
      With SentToClearingDoc
            .folderDirect = "|²ñï³ùÇÝ ÷áË³ÝóáõÙÝ»ñÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|àõÕ³ñÏí³Í  Ñ³ÝÓÝ³ñ³ñ³·ñ»ñ|àõÕ³ñÏí³Í ÃÕÃ."
            .stDate = todayDMY
            .eDate = todayDMY
            .pysysIn = "Ð"
            .wCur = "001"
            .acsBranch = "00"
            .acsDepart = "1"
      End With
      
      Call Fill_SentToClearingDoc(SentToClearingDoc)
      
      ' Մարել վճարման հանձնարարագիրը
      state = ConfirmContractDoc(1, PaymentRequestDoc.cliCode, c_ToFade, 5, "²Ûá")
      
      If Not state Then
            Log.Error("Մարել գործողությունը չի կատարվել")
            ' Փակել ՀԾ-Բանկ ծրագիրը
            Call Close_AsBank()  
            Exit Sub
      End If 
      
      BuiltIn.Delay(2000)
      frmPttel.Close
      
                  'DOCS
                  queryString = " SELECT COUNT(*) FROM DOCS WHERE fISN = " & PaymentRequestDoc.fISN & _
                                            " and fNAME = 'PayReq' and fSTATE = '6'"
                  sqlValue = 1
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
                  'FOLDERS
                  queryString = " SELECT COUNT(*) FROM FOLDERS WHERE fISN = " & PaymentRequestDoc.fISN 
                  sqlValue = 0
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If

                  'DOCLOG
                  queryString = " SELECT COUNT(*) FROM DOCLOG WHERE fISN = " & PaymentRequestDoc.fISN 
                  sqlValue = 8
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
                  'PAYMENTS
                  queryString = " SELECT COUNT(*) FROM PAYMENTS WHERE fISN = " & PaymentRequestDoc.fISN & _
                                            " and fDOCTYPE = 'PayReq' and fCUR = '001' and fSUMMA = '10000.00'" & _
                                            " and fSUMMAAMD = '4000000.00' and fSUMMAUSD = '10000.00' and fCOM = 'Ð³ñÏ»ñÇ Ù³ñáõÙ'" & _
                                            " and fCHARGESUM = '0.00' and fCHARGESUMAMD = '0.00' and fCHARGESUM2 = '0.00'"
                  sqlValue = 1
                  colNum = 0
                  sql_isEqual = CheckDB_Value(queryString, sqlValue, colNum)
                  If Not sql_isEqual Then
                    Log.Error("Querystring = " & queryString & ":  Expected result = " & sqlValue)
                  End If
                  
      ' Մուտք "Հաշվառված ուղարկված փոխանցումներ" թղթապանակ
      Set AccForTransfers = New_AccForTransfers()
      With AccForTransfers
              .folderDirect = "|²ñï³ùÇÝ ÷áË³ÝóáõÙÝ»ñÇ ²Þî|ÂÕÃ³å³Ý³ÏÝ»ñ|Ð³ßí³éí³Í áõÕ³ñÏí³Í ÷áË³ÝóáõÙÝ»ñ"
              .stDate = todayDMY
              .eDate = todayDMY
              .wUser = "77"
              .docType = "PayReq"
              .acsBranch = "00"
              .acsDepart = "1"
              .selectedView = "SentPay"
      End With
      
      Call Fill_AccForTransfers(AccForTransfers)
      
      state = CheckContractDoc(2, PaymentRequestDoc.cliCode)
      
      If Not state Then
            Log.Error("Փաստաթուղթն առկա չէ 'Ուղարկված փոխանցումներ' թղթապանակում")
      End If
      
      frmPttel.Close
      
      ' Փակել ՀԾ-Բանկ ծրագիրը
      Call Close_AsBank()  
      
End Sub