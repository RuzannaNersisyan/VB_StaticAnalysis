'USEUNIT Constants
'USEUNIT Library_Common
'USEUNIT Mortgage_Library

Dim rowCount_1, rowCount_2, rowCount_3, rowCount_4, rowCount_5

' Գրաֆիկով վարկային պայմանագրի ստեղծում
' clientCode - հաճախորդ
' mAccacc - Հաշվարկային հաշիվ
' summ - վարկի գումար
' mDate - Պայմանագրի/Վարկի կնքման ամսաթիվ
' dateGive - Վարկի հատկացման ամսաթիվ
' dateAgr - Մարման ժամկետ
' valCheck - Պարտքերի ավտոմատ մարում
' datesFilltype - Ամսաթվերի լրացման ձև
' fixDays - Մարման օրը
' passDirection - Շջանցման ուղղություն
' summDateSelect - Գումարների ամսաթվերի ընտրություն
' summFillType - Գումարների բաշխման ձև
' loanRatesSect - Վարկի Տոկոսադրույք
' unusedPortRate - Չոգտագործված մասի տոկոսադրույք
' unusedPortRateSec - Չոգտագործված մասի տոկոսադրույքի բաժին
' subsRate - Սուբսիդավորման տոկոսադրույք
' subsRateSect - Սուբսիդավորման տոկոսադրույք բաժին
' penOverMoney - Ժամկետանց գումարի տույժ
' penOverMoneySect - Ժամկետանց գումարի տույժ բաժին
' penOverLoan - Ժամկետանց տոկոսի տույժ
' penOverLoanSect - Ժամկետանց տոկոսի տույժ բաժին
' sect - Ճուղայնություն
' purpose - Նպատակ
' mShedule - ծրագիր
' mGuarantee - Երաշխավորություն
' mCountry - Երկիր
' lRegion - Մարզ
' mRegion - Մարզ(Նոր ՎՌ)
' mNote - Նշում 2 
' paperCode - Պայմ. թղթային N
Sub CreatingLoanAgrWithSchedule(contType, fISN, docNum, clientCode, wCurr, mAccacc,summ, mDate,dateGive, dateLngEnd, dateAgr, valCheck,_
                                                                 mixedSum, datesFilltype, fixDays, agrPeriod, agrPeriodPer, passDirection, summDateSelect, summFillType, loanRates,_
                                                                 loanRatesSect, unusedPortRate, unusedPortRateSec, subsRate, subsRateSect,_
                                                                 penOverMoney, penOverMoneySect, penOverLoan, penOverLoanSect, sect, purpose,_
                                                                 mShedule, mGuarantee,mCountry, lRegion, mRegion, mNote, paperCode)
        
       Dim wtdbgView, tdbgViewn
       Set wtdbgView  =  p1.VBObject("frmModalBrowser").VBObject("tdbgView")
       BuiltIn.Delay(1000)   
       
       ' Ստուգում որ պայմանագիր պատուհանը բացվել է
       If Not p1.WaitVBObject("frmModalBrowser",2000).Exists  Then
                Log.Error("Պայմանագիր պատուհանը չի բացվել")
                Exit Sub 
       End If
       
       ' Ընտրել անհրաժեշտ պայմանագրի տեսակը
       Do until  wtdbgView.EOF
             If Trim(wtdbgView.Columns.Item(1).value) = Trim(contType) then
                    wtdbgView.Keys("[Enter]")
                    Exit Do
             Else
                    wtdbgView.MoveNext
             End If
       Loop  
        
       BuiltIn.Delay(1000)
       ' Վարկային պայմանագրի ISN - ի ստացում
       fISN = wMDIClient.VBObject("frmASDocForm").DocFormCommon.Doc.ISN
       BuiltIn.Delay(1000)   
       
       ' Ստուգում որ Վարկային պայմանագիր պատուհանը բացվել է
       If Not wMDIClient.WaitVBObject("frmASDocForm",2000).Exists Then
              Log.Error("Վարկային պայմանագիր պատուհանը չի բացվել")
              Exit Sub
       End If
       
       ' Վարկային պայմանագրի համարի ստացում
       
       docNum = Get_Rekvizit_Value("Document",1,"General","CODE")
      
       ' Հաճախորդ դաշտի լրացում
            Call Rekvizit_Fill("Document", 1, "General", "CLICOD", clientCode)
       ' Արժույթ դաշտի լրացում
            Call Rekvizit_Fill("Document", 1, "General", "CURRENCY", wCurr)
       ' Հաշվարկային հաշիվ դաշտի լրացում
            Call Rekvizit_Fill("Document", 1, "General", "ACCACC", mAccacc)
       ' Վարկի գումար դաշտի լրացում
            Call Rekvizit_Fill("Document", 1, "General", "SUMMA", summ)
       ' Կնքման ամսաթիվ դաշտի լրացում
            Call Rekvizit_Fill("Document", 1, "General", "DATE","^!A[Del]" &  mDate)
       ' Հատկացման ամսաթիվ դաշտի լրացում
            Call Rekvizit_Fill("Document", 1, "General", "DATEGIVE","^!A[Del]" &  dateGive)
       ' Մարման ժամկետ դաշտի լրացում
            Call Rekvizit_Fill("Document", 1, "General", "DATEAGR","^!A[Del]" &  dateAgr)   
            If contType = "¶ñ³ýÇÏáí í³ñÏ³ÛÇÝ å³ÛÙ³Ý³·Çñ (·Í³ÛÇÝ)" Then
                ' Վարկային գծի գործելու ժամկետ դաշտի լրացում
                Call Rekvizit_Fill("Document", 1, "General", "DATELNGEND", dateLngEnd)  
                ' Գումարները վերաբաշխել չեկբոքսի լրացում
                Call Rekvizit_Fill("Document", 4, "CheckBox", "MIXEDSUMSINSCH", mixedSum)  
            End If
       ' Պարտքերի ավտոմատ մարում դաշտի լրացում    
            Call Rekvizit_Fill("Document", 3, "CheckBox", "AUTODEBT", valCheck)        
       ' Ամսաթվերի լրացման ձև դաշտի լրացում
            Call Rekvizit_Fill("Document", 4, "General", "DATESFILLTYPE", datesFilltype)
       ' Մարման օրը դաշտի լրացում
            Call Rekvizit_Fill("Document", 4, "General", "FIXEDDAYS", fixDays)
       ' Պարբերություն դաշտի լրացում
            Call Rekvizit_Fill("Document", 4, "General", "AGRPERIOD", agrPeriod & "[Tab]" & agrPeriodPer)
       ' Շջանցման ուղղություն դաշտի լրացում
            Call Rekvizit_Fill("Document", 4, "General", "PASSOVDIRECTION", passDirection)
       ' Գումարների ամսաթվերի ընտրություն դաշտի լրացում
            Call Rekvizit_Fill("Document", 4, "General", "SUMSDATESFILLTYPE", summDateSelect)
       ' Գումարների բաշխման ձև դաշտի լրացում
            Call Rekvizit_Fill("Document", 4, "General", "SUMSFILLTYPE", summFillType) 
       ' Վարկի Տոկոսադրույք  դաշտի լրացում 
            Call Rekvizit_Fill("Document", 6, "General", "PCAGR", loanRates & "[Tab]" & loanRatesSect ) 
       ' Չոգտագործված մասի տոկոսադրույք դաշտի լրացում
            Call Rekvizit_Fill("Document", 6, "General", "PCNOCHOOSE", unusedPortRate & "[Tab]" & unusedPortRateSec )
       ' Սուբսիդավորման տոկոսադրույք դաշտի լրացում
            Call Rekvizit_Fill("Document", 6, "General", "PCGRANT", subsRate & "[Tab]" & subsRateSect )       
       ' Ժամկետանց գումարի տույժ դաշտի լրացում
            Call Rekvizit_Fill("Document", 7, "General", "PCPENAGR", penOverMoney & "[Tab]" & penOverMoneySect )
       ' Ժամկետանց տոկոսի տույժ դաշտի լրացում
            Call Rekvizit_Fill("Document", 7, "General", "PCPENPER", penOverLoan & "[Tab]" & penOverLoanSect )
       ' Ճուղայնություն դաշտի լրացում
            Call Rekvizit_Fill("Document", 8, "General", "SECTOR", sect)
       ' Նպատակ դաշտի լրացում
            Call Rekvizit_Fill("Document", 8, "General", "AIM", purpose)
       ' Ծրագիր դաշտի լրացում
            Call Rekvizit_Fill("Document", 8, "General", "SCHEDULE", mShedule)
       ' Երաշխավորություն դաշտի լրացում
            Call Rekvizit_Fill("Document", 8, "General", "GUARANTEE", mGuarantee)
       ' Երկիր դաշտի լրացում
            Call Rekvizit_Fill("Document", 8, "General", "COUNTRY", mCountry)
       ' Մարզ դաշտի լրացում
            Call Rekvizit_Fill("Document", 8, "General", "LRDISTR", lRegion)
       ' Մարզ(Նոր ՎՌ) դաշտի լրացում
            Call Rekvizit_Fill("Document", 8, "General", "REGION", mRegion)
       ' Նշում 2 դաշտի լրացում
            Call Rekvizit_Fill("Document", 8, "General", "NOTE2", mNote)
       ' Պայմ. թղթային N դաշտի լրացում
            Call Rekvizit_Fill("Document", 8, "General", "PPRCODE", paperCode)
       ' Կատարել կոճակի սեղմում
            Call ClickCmdButton(1, "Î³ï³ñ»É")
          
       BuiltIn.Delay(2000) 
       Set tdbgViewn  =  wMDIClient.VBObject("frmPttel").VBObject("tdbgView")
       BuiltIn.Delay(1000)   
       
       ' Ստուգում որ Վարկային պայմանագիրը ստեղծվել է
       If  tdbgViewn.ApproxCount <> 1 Then
             Log.Error("Վարկային պայմանագրիը չի ստեղծվել")
             Exit Sub
       End If
          
End Sub
    
' Վճարումների գրաֆիկի փաստաթղթերի ստեղծում
' contractName - Պայմանագրի անուն
' Nominal - Ամսաթիվ
' Price - Գումար
Sub PaymentScheduleDocumentCreation(contractName, nominal, price)

        Dim  tdbgView
        Set tdbgView  =  wMDIClient.VBObject("frmPttel").VBObject("tdbgView")
    
        ' Կատարել բոլոր գործողությունները
        Call wMainForm.MainMenu.Click(c_AllActions)
        ' Մարման գրաֆիկի նշանակում
        Call wMainForm.PopupMenu.Click(c_RepaySchedule)

        Do Until  tdbgView.EOF
                 If Trim( tdbgView.Columns.Item(0).Value) = Trim(contractName) Then
                        ' Կատարել բոլոր գործողությունները
                        Call wMainForm.MainMenu.Click(c_AllActions)
                        ' Այլ վճարումների գրաֆիկի նշանակում
                        Call wMainForm.PopupMenu.Click(c_OtherPaySchedule)
                                With  wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame").VBObject("DocGrid")
                                      ' Ամսաթիվ դաշտի լրացում
                                      .Row = 0
                                      .Col = 0
                                      .Keys(nominal & "[Enter]")
                                      ' Գումար դաշտի լրացում
                                      .Row = 0
                                      .Col = 2
                                      .Keys(price & "[Enter]" )
                                End With 
                        ' Կատարել կոճակի սեղմում
                        Call ClickCmdButton(1, "Î³ï³ñ»É")
                        Exit Do                   
                 Else
                          tdbgView.MoveNext
                 End If
        Loop 
     
End Sub

' Պայմանագիրն ուղարկել հաստատման
' contractName - Պայմանագրի անուն
Sub  SendContractForApproval(contractName)

       Dim  tdbgView
       
       Set tdbgView = wMDIClient.VBObject("frmPttel").VBObject("tdbgView")
       tdbgView.MoveFirst
       Do until  tdbgView.EOF
               If  Trim(tdbgView.Columns.Item(0).Value) = Trim(contractName) Then
                      BuiltIn.Delay(2000)
                      ' Կատարել բոլոր գործողությունները
                      Call wMainForm.MainMenu.Click(c_AllActions)
                      ' Ուղարկել հաստատման գործողության կատարում
                      Call wMainForm.PopupMenu.Click(c_SendToVer)
                      ' Այո կոճակի սեղմում
                      Call ClickCmdButton(5, "²Ûá")
                      BuiltIn.Delay(2000)
                      Call Close_Pttel("frmPttel")
                      Exit Do
               Else
                      ' Անցնել հաջորդ տողի վրա
                      tdbgView.MoveNext
               End If     
      Loop

End Sub


' Փաստաթղթի վավերացում
Sub DocValidate(docNum)

        BuiltIn.Delay(2000)
        Dim tdbgView
        Set tdbgView= wMDIClient.VBObject("frmPttel").VBObject("tdbgView")

        Do until  tdbgView.EOF
              If  Trim(tdbgView.Columns.Item(2).Value) = Trim(docNum)  Then
                  ' Կատարել բոլոր գործողությունները
                  Call wMainForm.MainMenu.Click(c_AllActions)
                  ' Վավերացնել փաստաթուղթը
                  Call wMainForm.PopupMenu.Click(c_ToConfirm)
                  BuiltIn.Delay(1800)
                  ' Հաստատել կոճակի սեղմում
                  Call ClickCmdButton(1, "Ð³ëï³ï»É")
                  BuiltIn.Delay(1800)
                   Exit do
              Else
                      tdbgView.MoveNext
              End If
      Loop   
      
End Sub
      
  
' Գրավ Պայմանագրի ստեղծում           
' pledgeType - Գործարքի տեսակ
' cliPledge - Գրավատու
' datePledge - Կնքման ամսաթիվ
' curPledge -  Արժույթ
' commPledge - Մեկնաբանություն
' docNumb - Դակումենտի համարը
' customer - Հաճախորդ
' thPledge - Գրավի առարկա
' newthPledge - Գրավի առարկա(նոր ՌՎ)
' plfISN - Գրավ Պայմանագրի ISN
Sub CreationOfPledgeContract(pledgeType, plfISN, pladgeDocNum,  cliPledge, datePledge, curPledge, commPledge,_
                                                         docNumb, customer, thPledge, newthPledge)

      Dim  frmASDocForm
      
      Call wMainForm.MainMenu.Click(c_AllActions)
      Call wMainForm.PopupMenu.Click(c_Safety  & "|" & c_AgrOpen & "|" & c_Mortgage)
      
'      wTdbgView = Sys.Process("Asbank").VBObject("frmModalBrowser").VBObject("tdbgView")
      
       Do Until  Sys.Process("Asbank").VBObject("frmModalBrowser").VBObject("tdbgView").EOF
               If   Trim(Sys.Process("Asbank").VBObject("frmModalBrowser").VBObject("tdbgView").Columns.Item(1).Value) = Trim(pledgeType)  Then
                      Sys.Process("Asbank").VBObject("frmModalBrowser").VBObject("tdbgView").Keys("[Enter]")
                      
                      BuiltIn.Delay(1000)   
                     ' Ստուգում որ Գրավի պայմանագիր պատուհանը բացվել է
                      If Not wMDIClient.WaitVBObject("frmASDocForm",2000).Exists Then
                            Log.Error("Գրավի պայմանագիր պատուհանը չի բացվել")
                            Exit Sub
                     End If
                     
                     ' ISN - ի ստացում
                     plfISN = wMDIClient.VBObject("frmASDocForm").DocFormCommon.Doc.ISN
                     
                     ' Պայմանագրի համարի ստացում
                     pladgeDocNum = Get_Rekvizit_Value("Document",1,"General","CODE")

                    ' Գրավատու դաշտի լրացում
                    Call Rekvizit_Fill("Document", 1, "General", "CLICOD", cliPledge)
                    ' Կնքման ամսաթիվ դաշտի լրացում
                    Call Rekvizit_Fill("Document", 1, "General", "DATE", datePledge)
                    ' Արժույթ դաշտի լրացում
                    Call Rekvizit_Fill("Document", 1, "General", "CURRENCY", curPledge)
                    ' Մեկնաբանություն դաշտի լրացում
                    Call Rekvizit_Fill("Document", 1, "General", "COMMENT", commPledge)     
                    
                    With   wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame").VBObject("DocGrid")
                               wMDIClient.Refresh
                               
                               ' Սեղմել enter  գրիդի երկրորդ սյան առաջին տողի վրա 
                              .Col = 2
                              .Keys("[Enter]" & " ")
                              
                              ' Ստուգում որ Գրիդի կապակցվող պայմանագիր պատուհանը բացվել է
                              If  Not p1.WaitVBObject("frmAsUstPar",3000).Exists Then
                                      Log.Error("Գրիդի կապակցվող պայմանագիր պատուհանը չի բացվել")
                                      Exit Sub
                              End If
                                  
                              ' Պայմանագրի N դաշտի լրացում
                              Call Rekvizit_Fill("Dialog", 1, "General", "AGRNUM", docNumb)
                              ' Հաճախորդ դաշտի լրացում
                              Call Rekvizit_Fill("Dialog", 1, "General", "CLICODE", customer)
                              ' Կատարել կոճակի սեղմում
                              Call ClickCmdButton(2, "Î³ï³ñ»É")
                              Sys.Process("Asbank").VBObject("frmModalBrowser").VBObject("tdbgView").Keys("[Enter]")
                                                  
                              ' Գրավի առարկա դաշտի լրացում
                              Call Rekvizit_Fill("Document", 3, "General", "SHRTNAME", thPledge)
                              ' Գրավի առարկա(նոր ՌՎ) դաշտի լրացում
                              Call Rekvizit_Fill("Document", 3, "General", "MORTSUBJECT", newthPledge)
                              ' Կատարել կոճակի սեղմում
                              Call ClickCmdButton(1, "Î³ï³ñ»É")
      
                              ' Ստուգում որ գրավի պայմանագիրը ստեղծվել է
                              If  Not wMDIClient.WaitVBObject("frmPttel_2",2000).Exists Then
                                  Log.Error("Գրավի պայմանագիրը չի ստեղծվել")
                                    Exit Sub
                              End If
                                      
                              Call Close_Pttel("frmPttel_2")
                   End With        
                   Exit Do
               Else
                       Sys.Process("Asbank").VBObject("frmModalBrowser").VBObject("tdbgView").MoveNext
               End If     
       Loop        
                   
End Sub

' Գանձում տրամադրումից
' datePay - Ամսաթիվ
' cashOrNo - Կանխիկ/Անկանխիկ
' docNumIn - Կանխիկ մուտք
Sub Charging(fISNChar, docNumChar, datePay, cashOrNo, fISNInput, docNumInput, kassNish, accCorr, applayConn)

        ' Գործողություններ /  Բոլոր գործողությունները
        Call wMainForm.MainMenu.Click(c_AllActions)
        ' Գործողություններ /  Տրամադրում/Մարում  /  Գանձում տրամադրումից
        Call wMainForm.PopupMenu.Click(c_Opers  & "|" & c_GiveAndBack & "|" & c_GiveCharge)
        
        ' Ստուգում որ Գանձում տրամադրումից պատուհանը բացվել է
        If  Not wMDIClient.WaitVBObject("frmASDocForm",2000).Exists Then
                  Log.Error("Գանձում տրամադրումից պատուհանը չի բացվել")
                  Exit Sub
        End If
        
        ' Գանձում տրամադրումից փաստաթղթի ISN - ի ստացում
        fISNChar = wMDIClient.VBObject("frmASDocForm").DocFormCommon.Doc.ISN
        
        ' Գանձում տրամադրումից փաստաթղթի համարի ստացում
        docNumChar = Get_Rekvizit_Value("Document",1,"Mask","CODE")
        
        ' Ամսաթիվ դաշտի լրացում
        Call Rekvizit_Fill("Document", 1, "General", "DATE", datePay)
        ' Կանխիկ/Անկանխիկ դաշտի լրացում
        Call Rekvizit_Fill("Document", 1, "General", "CASHORNO", cashOrNo)
        
        If  cashOrNo = "1" Then
              ' Կատարել կոճակի սեղմում
              Call ClickCmdButton(1, "Î³ï³ñ»É")  
              
              ' Կանխիկ մուտք փաստաթղթի ISN - ի ստացում
              fISNInput = wMDIClient.VBObject("frmASDocForm").DocFormCommon.Doc.ISN
              
             ' Կանխիկ մուտք փաստաթղթի համարի ստացում
              docNumInput = Get_Rekvizit_Value("Document",1,"General","DOCNUM")
              
              ' Դրամարկղի նիշ դաշտի ստացում
              kassNish = Get_Rekvizit_Value("Document",1,"Mask","KASSIMV")
              ' Կատարել կոճակի սեղմում
              Call ClickCmdButton(1, "Î³ï³ñ»É")
        
              BuiltIn.Delay(1000)
              wMDIClient.VBObject("FrmSpr").Close
        Else
              ' Հաշիվ դաշտի լրացում
               Call Rekvizit_Fill("Document", 1, "General", "ACCCORR", accCorr) 
               ' Կիրառել փախկապակցման սխեման չեքբոքսի լրացում
               Call Rekvizit_Fill("Document", 1, "CheckBox", "APPLYCONNSCH", applayConn) 
               
               ' Կատարել կոճակի սեղմում
               Call ClickCmdButton(1, "Î³ï³ñ»É")
        End If
        
End Sub

' Վարկի տրամադրում ֆունկցիա
' mDate - Ամսաթիվ
' cashOrNo - Կանխիկ/Անկանխիկ
' docNumOut  - Վարկի տրամադրում փաստաթղթի համար
Sub SupplyCredit(docNumLoan, fISNLoan, mDate, cashOrNo, docNumOut, fISNOut, accCor)

         BuiltIn.Delay(1000) 
         ' Գործողություններ /  Բոլոր գործողությունները
         Call wMainForm.MainMenu.Click(c_AllActions)
         ' Գործողություններ /  Տրամադրում/Մարում  /  Վարկի տրամադրում
         Call wMainForm.PopupMenu.Click(c_Opers  & "|" & c_GiveAndBack & "|" & c_CredGrant)
         BuiltIn.Delay(1000)
          
         ' Վարկի տրամադրում պատուհանի բացման ստուգում
         If Not wMDIClient.WaitVBObject("frmASDocForm",2000).Exists Then   
                   Log.Error("Կանխիկ ելք պատուհանը չի բացվել")
                   Exit Sub
         End If
         
         ' Վարկի տրամադրում փաստաթղթի համարի ստացում
         docNumLoan = Get_Rekvizit_Value("Document",1,"Mask","CODE")
           
         ' Վարկի տրամադրում փաստաթղթի ISN -ի ստացում
         fISNLoan = wMDIClient.VBObject("frmASDocForm").DocFormCommon.Doc.ISN
         
          ' Ամսաթիվ դաշտի լրացում
         Call Rekvizit_Fill("Document", 1, "General", "DATE","^A[Del]" & mDate)
         ' Կանխիկ/Անկանխիկ դաշտի լրացում
         Call Rekvizit_Fill("Document", 1, "General", "CASHORNO","^A[Del]" & cashOrNo)
         
         If cashOrNo = "1" Then
             ' Կատարել կոճակի սեղմում
             Call ClickCmdButton(1, "Î³ï³ñ»É")
          
             ' Վարկի տրամադրում փաստաթղթի համարի ստացում
             docNumOut = Get_Rekvizit_Value("Document",1,"General","DOCNUM")
           
             ' Վարկի տրամադրում փաստաթղթի ISN -ի ստացում
             fISNOut = wMDIClient.VBObject("frmASDocForm").DocFormCommon.Doc.ISN
         
             ' Կատարել կոճակի սեղմում
             Call ClickCmdButton(1, "Î³ï³ñ»É")
             wMDIClient.VBObject("FrmSpr").Close
         Else 
              ' Հաշիվ դաշտի լրացում
              Call Rekvizit_Fill("Document", 1, "General", "ACCCORR", accCor)
              ' Կատարել կոճակի սեղմում
             Call ClickCmdButton(1, "Î³ï³ñ»É")
         End If
         
         Call Close_Pttel("frmPttel")
         
End Sub

' Կանխիկ մուտք փաստաթուղթը ուղարկել հաստատման
' stDate - Ժամանակահատվածի սկիզբ
' enDate - Ժամանակահատվածի վերջ
' docNumIn - Կանխիկ Մուտք փաստաթղթի համարը
Sub SendConfirmCashAccessDoc(stDate, enDate, docNumIn)

        Dim tdbgView, wMainForm
        
        ' Մուտք Աշխատանքային փաստաթղթեր թղթապանակ
        Call wTreeView.DblClickItem("|Ð³×³Ëáñ¹Ç ëå³ë³ñÏáõÙ ¨ ¹ñ³Ù³ñÏÕ |²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ")
        Set wMainForm =  Sys.Process("Asbank").VBObject("MainForm")
        BuiltIn.Delay(1000)   
        
        ' Ստուգում որ Աշխատանքային Փաստաթղթեր թղթապանակ բացվել է
        If Not Sys.Process("Asbank").VBObject("frmAsUstPar").Exists Then
            Log.Error("Աշխատանքային Փաստաթղթեր պատուհանը չի բացվել")
            Exit Sub
        End If
        
        ' Ժամանակահատվածի սկիզբ դաշտի լրացում
        Call Rekvizit_Fill("Dialog", 1, "General", "PERN", stDate )
        ' Ժամանակահատվածի ավարտ դաշտի լրացում
        Call Rekvizit_Fill("Dialog", 1, "General", "PERK", enDate )
        ' Կատարել կոճակի սեղմում
        Call ClickCmdButton(2, "Î³ï³ñ»É")
        BuiltIn.Delay(1000)   
        
        Set  tdbgView = wMainForm.Window("MDIClient", "", 1).VBObject("frmPttel").VBObject("tdbgView")
        Do until  tdbgView.EOF
                ' Կանխիկ մուտք փաստաթղթի փնտրում 
                If  Trim(tdbgView.Columns.Item(2).Value) = Trim(docNumIn)  Then
                      ' Կատարել բոլոր գործողությունները
                      Call wMainForm.MainMenu.Click(c_AllActions)
                      ' Ուղարկել հաստատման գործողության կատարում
                      Call wMainForm.PopupMenu.Click(c_SendToVer )
                      ' Հաստատել կոճակի սեղմում
                      Call ClickCmdButton(2, "Î³ï³ñ»É")
                      Exit Do
                Else
                      tdbgView.MoveNext
                End If      
        Loop
        
End Sub

' Կանխիկ մուտք փաստաթղթի հաստատում
' docNumIn - Կանխիկ Մուտքի համարը
Function ConfirmCashAccessDoc(docNumIn)

        Dim tdbgView, state
        state = False
         
        ' Ստուգում որ Աշխատանքային Փաստաթղթեր թղթապանակը բացվել է  
        If  Not wMDIClient.WaitVBObject("frmPttel",3000).Exists Then
              Log.Error("Աշխատանքային Փաստաթղթեր թղթապանակը չի բացվել")
              Exit Function
        End If
           
        Set  tdbgView = wMDIClient.VBObject("frmPttel").VBObject("tdbgView")
      
        Do until  tdbgView.EOF
             ' Ստուգում որ կանխիկ մուտք փաստաթուղթը գոյություն ունի
             If  Trim(tdbgView.Columns.Item(3).Value) = Trim(docNumIn) Then
                    ' Կատարել բոլոր գործողությունները
                    Call wMainForm.MainMenu.Click(c_AllActions)
                    ' Վավերացնել գործողության կատարում
                    Call wMainForm.PopupMenu.Click(c_ToConfirm)
                    ' Հաստատել կոճակի սեղմում
                    Call ClickCmdButton(1, "Ð³ëï³ï»É")
                    Call Close_Pttel("frmPttel")
                    state = True
                    Exit Do
             Else
                tdbgView.MoveNext
             End If     
        Loop
      ConfirmCashAccessDoc = state
      
End Function 
 
' Տոկոսների հաշվարկ
' dateCharge - Հաշվարկման ամսաթիվ
' dateAction - Գործողության ամսաթիվ
' calcfISN - Տոկոսների հաշվարկի ISN
Sub PercentCalculation(dateCharge, dateAction, percentMoney, calcfISN )

        Dim mDIClient
        
        ' Կատարել բոլոր գործողությունները
        Call wMainForm.MainMenu.Click(c_AllActions)
        ' Կատարել Գործողություններ/Տոկոսներ/Տոկոսների հաշվարկ
        Call wMainForm.PopupMenu.Click(c_Opers  & "|" & c_Interests & "|" & c_PrcAccruing)
        
        ' Տոկոսների հաշվարկ փաստաթղթի ISN - ի ստացում
        Set mDIClient = Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1)
        calcfISN = mDIClient.VBObject("frmASDocForm").DocFormCommon.Doc.ISN
        
        ' Հաշվարկման ամսաթիվ դաշտի լրացում
        Call Rekvizit_Fill("Document", 1, "General", "DATECHARGE","^!A[Del]" &  dateCharge )
        ' Գործողության ամսաթիվ դաշտի լրացում
        Call Rekvizit_Fill("Document", 1, "General", "DATE","^!A[Del]" &  dateAction )
        ' Տոկոսագումար դաշտի լրացում
        Call Rekvizit_Fill("Document", 1, "General", "SUMPER", percentMoney )
        ' Կատարել կոճակի սեղմում
        Call ClickCmdButton(1, "Î³ï³ñ»É")

End Sub


' Վարկի պարտքի մարում փաստաթղթի ստեղծում
' debtDate - Ամսաթիվ
' debtSum - Հիմնական գումար 
' debtSumPer - Տոկոսագումար
Sub LoanDebtRepaymentDocCreate(loanReptISN, debtDate, debtSum, debtSumPer)

        ' Կատարել բոլոր գործողությունները
        Call wMainForm.MainMenu.Click(c_AllActions)
        ' Կատարել Գործողություններ/Տրամադրում\Մարում/ Պարտքերի մարում
        Call wMainForm.PopupMenu.Click(c_Opers  & "|" & c_GiveAndBack & "|" & c_PayOffDebt)
        
        BuiltIn.Delay(1000)   
        ' Ստուգում, որ Վարկի պարտքի մարում պատուհանը բացվել է
        If  Not wMDIClient.WaitVBObject("frmASDocForm",2000).Exists Then
                Log.Error("Վարկի պարտքի մարում փաստաթուղթը չի բացվել")
                Exit Sub
        End If         
        
        ' Վարկի մարում փաստաթղթի ISN - ի ստացում
        loanReptISN = wMDIClient.VBObject("frmASDocForm").DocFormCommon.Doc.ISN
        
        ' Ամսաթիվ դաշտի լրացում
        Call Rekvizit_Fill("Document", 1, "General", "DATE", debtDate  )
        ' Հիմնական գումար դաշտի լրացում
        Call Rekvizit_Fill("Document", 1, "General", "SUMAGR", debtSum)
        ' Տոկոսագումար դաշտի լրացում
        Call Rekvizit_Fill("Document", 1, "General", "SUMPER", debtSumPer )
        ' Կատարել կոճակի սեղմում
        Call ClickCmdButton(1, "Î³ï³ñ»É")
        BuiltIn.Delay(1500)
        ' Այո կոճակի սեղմում
        Call ClickCmdButton(5, "²Ûá")       

End Sub

' խմբային տոկոսների հաշվարկ
' calcDate - Հաշվարկման ամսաթիվ
' regDate - Ձևակերպման ամսաթիվ
Sub InterestGroupCalculation (calcDate, regDate, checkCount)
    BuiltIn.Delay(3000)
    wMDIClient.VBObject("frmPttel").Keys("[Ins]")
    ' Կատարել բոլոր գործողությունները
    Call wMainForm.MainMenu.Click(c_AllActions)
    ' Կատարել Խմբային հաշվարկ գործողությունը
    Call wMainForm.PopupMenu.Click(c_GroupCalc )
        
    ' Հաշվարկման ամսաթիվ դաշտի լրացում
    Call Rekvizit_Fill("Dialog", 1, "General", "CLOSEDATE", calcDate)
    ' Ձևակերպման ամսաթիվ դաշտի լրացում
    Call Rekvizit_Fill("Dialog", 1, "General", "SETDATE", regDate)
    ' Տոկոսների հաշվարկ դաշտի լրացում
    Call Rekvizit_Fill("Dialog", 1, "CheckBox", "CHG", checkCount)
    ' Կատարել կոճակի սեղմում
    Call ClickCmdButton(2, "Î³ï³ñ»É")  
    ' Այո կոճակի սեղմում
    Call ClickCmdButton(5, "²Ûá")  
End Sub

' խմբային պարտքերի մարում
' calcDate - Հաշվարկման ամսաթիվ
' regDate - Ձևակերպման ամսաթիվ
' checkCount - Պարտքի մարում նշիչ
Sub GroupDebt (calcDate, regDate, checkCount)

        wMDIClient.VBObject("frmPttel").Keys("[Ins]")
        ' Կատարել բոլոր գործողությունները
        Call wMainForm.MainMenu.Click(c_AllActions)
        ' Կատարել Խմբային հաշվարկ գործողությունը
        Call wMainForm.PopupMenu.Click(c_GroupCalc )
        ' Հաշվարկման Ամսաթիվ դաշտի լրացում
        Call Rekvizit_Fill("Dialog", 1, "General", "CLOSEDATE", calcDate)
        ' Ձևակերպման ամսաթիվ դաշտի լրացում
        Call Rekvizit_Fill("Dialog", 1, "General", "SETDATE", regDate)
        ' Պարտքի մարում դաշտի լրացում
        Call Rekvizit_Fill("Dialog", 1, "CheckBox", "DBT", checkCount)
        ' Կատարել կոճակի սեղմում
        Call ClickCmdButton(2, "Î³ï³ñ»É")  
        ' Այո կոճակի սեղմում
        Call ClickCmdButton(5, "²Ûá")  
        
End Sub


' Մարումներ գրաֆիկի վերանայում
' reDate - Ամսաթիվ
' reDateAgr - Մարման ժամկետ
' sumTotal - Գումարի քանակ
Sub RepaymentScheduleReview(repShedISN, reDate, reDateAgr, sumTotal)

        ' Կատարել բոլոր գործողությունները
        Call wMainForm.MainMenu.Click(c_AllActions)
        ' Կատարել Գործողություններ/Ժամկետներ/Գրաֆիկի վերանայում
        Call wMainForm.PopupMenu.Click(c_TermsStates & "|" & c_Dates & "|" & c_ReviewSchedule)
        
        ' Մարումների գրաֆիկ պայմանագրի ISN - ի ստացում 
        repShedISN = wMDIClient.VBObject("frmASDocForm").DocFormCommon.Doc.ISN
           
        ' Ամսաթիվ դաշտի լրացում
        Call Rekvizit_Fill("Document", 1, "General", "DATE", reDate )
        ' Մարման ժամկետ դաշտի լրացում
        Call Rekvizit_Fill("Document", 1, "General", "DATEAGR", reDateAgr )
        ' Նշել Ընթացիկ գրաֆիկի պատճ. դաշտը
        wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame").VBObject("CheckBox").Click
        
        BuiltIn.Delay(1000)   
        ' Ստուգում որ Գրաֆիկի լրացման ձև դիալոգը բացվել է
        If Not p1.WaitVBObject("frmAsUstPar",2000).Exists Then
                Log.Error("Գրաֆիկի լրացման ձև դիալոգը չի բացվել")
                Exit Sub
        End If
        
        Call Rekvizit_Fill("Dialog", 1, "General", "SUMTOT", sumTotal )
        ' Կատարել կոճակի սեղմում
        Call ClickCmdButton(2, "Î³ï³ñ»É")  
        ' Կատարել կոճակի սեղմում
        Call ClickCmdButton(1, "Î³ï³ñ»É")     
        
        ' Ստուգում որ Զգուշացում պատուհանը բացվել է
        If Not  p1.WaitVBObject("frmAsMsgBox", 2000).Exists Then
               Log.Message("Զգուշացում պատուհանը չի բացվել")
               Exit Sub
        End If      
         
        ' Կատարել կոճակի սեղմում
        Call ClickCmdButton(5, "Î³ï³ñ»É")  

End Sub

' Վարկի տրամադրում փաստաթղթի ստեղծում
' mDate - Ժամանակահատվածի սկիզբ
' dateAgr - Ժամանակահատվածի ավարտ
' createDate - Վարկի տրամադրում փաստաթղթի ստեղծման ամսաթիվ
' giveLoan - Փաստաթղթի անունը
Function CheckCreatedLoanDocOrNo(mDate, dateAgr, createDate, giveLoan )
      
      Dim state, tdbgView
      
      ' Կատարել բոլոր գործողությունները
      Call wMainForm.MainMenu.Click(c_AllActions)
      ' Գործողությունների դիտում
      Call wMainForm.PopupMenu.Click(c_OpersView)
      
      Call Rekvizit_Fill("Dialog", 1, "General", "START", mDate)
      
      Call Rekvizit_Fill("Dialog", 1, "General", "END", dateAgr)
      ' Հաստատել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")  
      
      BuiltIn.Delay(2000) 
      ' Ստուգում որ Գործողությունների դիտում պատուհանը բացվել է
      If  Not wMDIClient.WaitVBObject("frmPttel_2",2000).Exists Then
             Log.Error("Գործողությունների դիտում պատուհանը չի բացվել")
             Exit Function
      End if
       
      BuiltIn.Delay(2000) 
      Set tdbgView = wMDIClient.VBObject("frmPttel_2").VBObject("tdbgView")
      
      state = False
      Do Until tdbgView.EOF
                If Trim(tdbgView.Columns.Item(0).Value) = createDate And Trim(tdbgView.Columns.Item(5).Value) = giveLoan  Then
                      Log.Message("Վարկի տրամադրում փաստաթուղթը ստեղծվել է")
                      Call Close_Pttel("frmPttel_2")
                      state = True
                      Exit Do 
                Else
                      tdbgView.MoveNext
                End If
      Loop
      CheckCreatedLoanDocOrNo = state
End Function 

' Մարումներ գրաֆիկի ջնջում
Sub DelRepSched()

        ' Գործողություններ / Բոլոր գործողությունները
        Call wMainForm.MainMenu.Click(c_AllActions)
        ' Մուտք Գրաֆիկներ թղթապանակ
        Call wMainForm.PopupMenu.Click(c_Folders & "|" & c_SchFolder )
        BuiltIn.Delay(1500) 
        
        wMDIClient.VBObject("frmPttel_2").VBObject("tdbgView").MoveLast  
        
        ' Ջնջել գործողության կատարում
        Call wMainForm.MainMenu.Click(c_Opers & "|" & c_Delete)
        ' Այո կոճակի սեղմում
        Call ClickCmdButton(3, "²Ûá")  
        Call Close_Pttel("frmPttel_2")
        
End Sub


' Կանխիկ ելք/Կանխիկ մուտք փաստաթղթերի ջնջում
' mDate - Ժամանակահատված
' docType - Մարման օրը
' docNumber - Փաստաթղթի տեսակ
Sub  DelPaymentdoc(mDate, docType, docNumber)

        Dim tdbgView
        ' Մուտք հաշվառված վճարային փաստաթղթեր
        Call wTreeView.DblClickItem("|Ð³×³Ëáñ¹Ç ëå³ë³ñÏáõÙ ¨ ¹ñ³Ù³ñÏÕ |Ð³ßí³éí³Í í×³ñ³ÛÇÝ ÷³ëï³ÃÕÃ»ñ")
        
        ' Ժամանակահատվածի սկիզբ դաշտի լրացում
        Call Rekvizit_Fill("Dialog", 1, "General", "PERN", mDate)
        ' Ժամանակահատվածի ավարտ դաշտի լրացում
        Call Rekvizit_Fill("Dialog", 1, "General", "PERK", mDate )
        ' Մարման օրը դաշտի լրացում
        Call Rekvizit_Fill("Dialog", 1, "General", "DOCTYPE", docType)
        ' Կատարել կոճակի սեղմում
        Call ClickCmdButton(2, "Î³ï³ñ»É")
        
        ' Ստուգում որ Վճարային փաստաթղթեր դիալոգը բացվել է
        If  Not wMDIClient.WaitVBObject("frmPttel",10000).Exists Then
               Log.Error("Վճարային փաստաթղթեր դիալոգը չի բացվել")
               Exit Sub 
        End If
        Set  tdbgView = wMDIClient.VBObject("frmPttel").VBObject("tdbgView")
               
        Do until  tdbgView.EOF
                ' Ստուգում որ  թղթապանակում փաստաթուղթը գոյություն ունի
                If  Trim(tdbgView.Columns.Item(2).Value) = Trim(docNumber)  Then
                      ' Կատարել բոլոր գործողությունները
                      Call wMainForm.MainMenu.Click(c_AllActions)
                      ' Ուղարկել հաստատման գործողության կատարում
                      Call wMainForm.PopupMenu.Click(c_Delete )
                      ' Այո կոճակի սեղմում
                      Call ClickCmdButton(3, "²Ûá")
                      ' Այո կոճակի սեղմում
                      Call ClickCmdButton(5, "²Ûá")
                      BuiltIn.Delay(1000)
                      BuiltIn.Delay(2000)
                      Call wMainForm.MainMenu.Click(c_Windows)
                      Call wMainForm.PopupMenu.Click(c_ClCurrWindow)
                      Exit Do
                Else
                      tdbgView.MoveNext
                End If      
        Loop
    
 End Sub 
 
 
' Ջնջել Պայամանգրեր թղթապանակից
Sub  DelTermsDoc(pladgeNumber)

        Dim  tdbgView
        
        ' Գործողություններ / Բոլոր գործողությունները
        Call wMainForm.MainMenu.Click(c_AllActions)
        ' Մուտք Պայմանագրի թղթապանակ
        Call wMainForm.PopupMenu.Click(c_Folders & "|" & c_AgrFolder )
        BuiltIn.Delay(2000)
        Set tdbgView = wMDIClient.VBObject("frmPttel_2").VBObject("tdbgView")
        ' Անցնել տողերի վրայով
        Do until  tdbgView.EOF
               ' Ստուգում որ պայմանագրեր թղթապանակում փաստաթուղթը գոյություն ունի
               If  Trim(tdbgView.Columns.Item(0).Value) = Trim(pladgeNumber)  Then
                      ' Կատարել բոլոր գործողությունները
                      Call wMainForm.MainMenu.Click(c_AllActions)
                      ' Ջնջել գործողության կատարում
                      Call wMainForm.PopupMenu.Click(c_Delete )
                      ' Այո կոճակի սեղմում
                      Call ClickCmdButton(3, "²Ûá")
                      Exit Do
               Else
                      ' Անցնել հաջորդ տող
                      tdbgView.MoveNext
               End If      
        Loop
        
        Call Close_Pttel("frmPttel_2")
 
End Sub


' Ջնջել պայմանագիրը
Sub DelDoc()

          BuiltIn.Delay(2000)
        ' Կատարել բոլոր գործողությունները
        Call wMainForm.MainMenu.Click(c_AllActions)
        ' Ջնջել գործողության կատարում
        Call wMainForm.PopupMenu.Click(c_Delete )
        ' Այո կոճակի սեղմում
        Call ClickCmdButton(3, "²Ûá")
        
End Sub


' Օվերդրաֆտ պայմանագրի ստեղծում
' contType - պայմանագրի տեսակը
' fISN - Օվերդրաֆտ պայմանագրի ISN
' docNum - Օվերդրաֆտ պայմանագրի համարը
' clientCode - հաճախորդ
' agType - Ձևանմուշի N
' curr - Արժույթ
' mAccacc - Հաշվարկային հաշիվ
' limitSumm - Սահմանաչափ
' isrGenerativ - Վերականգնվող նշիչ
' allLim - Սահմանաչափերով բաշխվող
' autoCap - Կապիտալացվող արժեք
' mDate - Կնքման ամսաթիվ
' dateGive - հատկացման ամսաթիվ
' dateAgr - Մարման ժամկետ
' valCheck - Պարտքերի ավտոմատ մարում
' cardDebtType - Գումարի մարում
' datesFilltype - Ամսաթվերի լրացման ձև
' agrBeg - Մարումների սկիզբ
' agrFin - Մարումների վերջ 
' fixDays - Մարման օրերը
' agrPeriod - Պարբերություն
' passDirection - Շջանցման ուղղություն
' summDateSelect - Գումարների ամսաթվերի ընտրություն
' summFillType - Գումարների բաշխման ձև
' overRates - Օվերդրաֆտի Տոկոսադրույք
' overRatesSect - Օվերդրաֆտի Տոկոսադրույք բաժին
' unusedPortRate - Չոգտագործված մասի տոկոսադրույք
' unusedPortRateSec - Չոգտագործված մասի տոկոսադրույք բաժին
' sect - Ճուղայնություն
' sectNew - Ճուղայնություն(Նոր ՎՌ) 
' purpose - Նպատակ
' mShedule - ծրագիր
' mGuarantee - Երաշխավորություն
' mCountry - Երկիր
' lRegion - Մարզ
' mRegion - Մարզ(Նոր ՎՌ)
' paperCode - Պայմ. թղթային N 

Sub CreatingOverdraftWithSchedule(contType, fISN, docNum, clientCode, agType, curr, mAccacc, limitSumm,_
                                                                   isrGenerativ, allLim, autoCap, mDate, dateGive, dateAgr, valCheck,debtJPart,_
                                                                   cardDebtType, datesFilltype, agrBeg, agrFin, fixDays, agrPeriod, passDirection,_
                                                                   summDateSelect, summFillType, overRates, overRatesSect, unusedPortRate,_
                                                                   unusedPortRateSec, sect, sectNew, purpose, mShedule, mGuarantee, mCountry,_
                                                                   lRegion, mRegion, paperCode)
        
         Dim tdbgView, tdbgViewn
         Set tdbgView  =  p1.VBObject("frmModalBrowser").VBObject("tdbgView")
        
         ' Ստուգում որ Նոր պայմանագրի ստեղծում պատուհանը բացվել է
         If Not p1.WaitVBObject("frmModalBrowser", 10000).Exists  Then
                  Log.Error("Նոր պայմանագիր ստեղծում պատուհանը չի բացվել")
                  Exit Sub 
         End If
         
         ' Ընտրել պայմանագրի տեսակը
         Do until  tdbgView.EOF
                If Trim(tdbgView.Columns.Item(1).value) = Trim(contType) then
                       tdbgView.Keys("[Enter]")
                       Exit Do
                Else
                       tdbgView.MoveNext
                End If
         Loop  
      
         ' Ստուգում որ Գրաֆիկով Օվերդրաֆտի պայմանագիր պատուհանը բացվել է
         If Not wMDIClient.WaitVBObject("frmASDocForm",20000).Exists Then
              Log.Error("Գրաֆիկով Օվերդրաֆտի պայմանագիր պատուհանը չի բացվել")
              Exit Sub
         End If
          
         ' Պայմանագրի ISN - ի ստացում
         fISN = wMDIClient.VBObject("frmASDocForm").DocFormCommon.Doc.ISN
         
         ' Պայմանագրի համարի ստացում
         docNum = Get_Rekvizit_Value("Document",1,"General","CODE")
     
      ' Հաճախորդ դաշտի լրացում
            Call Rekvizit_Fill("Document", 1, "General", "CLICOD", clientCode)
       ' Ձևանմուշի N դաշտի լրացում
            Call Rekvizit_Fill("Document", 1, "General", "AGRTYPE", agType)
       ' Արժույթ դաշտի լրացում
            Call Rekvizit_Fill("Document", 1, "General", "CURRENCY", curr)                     
       ' Հաշվարկային հաշիվ դաշտի լրացում
            Call Rekvizit_Fill("Document", 1, "General", "ACCACC", mAccacc)
       ' Սահմանաչափ դաշտի լրացում
            Call Rekvizit_Fill("Document", 1, "General", "SUMMA", limitSumm)
       ' Վերականգնվող նշիչ դաշտի լրացում
            Call Rekvizit_Fill("Document", 1, "CheckBox", "ISREGENERATIVE", isrGenerativ)
       ' Սահմանաչափերով բաշխվող դաշտի լրացում
            Call Rekvizit_Fill("Document", 1, "CheckBox", "ALLOCATEWITHLIM", allLim)
       ' Կապիտալացվող դաշտի լրացում
            Call Rekvizit_Fill("Document", 1, "CheckBox", "AUTOCAP", autoCap)            
       ' Կնքման ամսաթիվ դաշտի լրացում
            Call Rekvizit_Fill("Document", 1, "General", "DATE", mDate)
       ' Հատկացման ամսաթիվ դաշտի լրացում
            Call Rekvizit_Fill("Document", 1, "General", "DATEGIVE", dateGive)
       ' Մարման ժամկետ դաշտի լրացում
            Call Rekvizit_Fill("Document", 1, "General", "DATEAGR", dateAgr)   
       ' Պարտքերի ավտոմատ մարում դաշտի լրացում    
            Call Rekvizit_Fill("Document", 3, "CheckBox", "AUTODEBT", valCheck)  
       ' Պարտքերի ավտոմատ մարում դաշտի լրացում    
            Call Rekvizit_Fill("Document", 3, "General", "DEBTJPART1", debtJPart)  
       ' Գումարի մարում ըստ հաշվարկային ամսի դաշտի լրացում   
            Call Rekvizit_Fill("Document", 3, "General", "CARDDEBTTYPE", cardDebtType)    
       ' Ամսաթվերի լրացման ձև դաշտի լրացում
            Call Rekvizit_Fill("Document", 4, "General", "DATESFILLTYPE", datesFilltype)
       ' Մարումների սկիզբ դաշտի լրացում
            Call Rekvizit_Fill("Document", 4, "General", "AGRMARBEG", agrBeg)            
       ' Մարումների վերջ դաշտի լրացում
            Call Rekvizit_Fill("Document", 4, "General", "AGRMARFIN", agrFin)            
       ' Մարման օրերը դաշտի լրացում
            Call Rekvizit_Fill("Document", 4, "General", "FIXEDDAYS", fixDays)
       ' Պարբերություն դաշտի լրացում
            Call Rekvizit_Fill("Document", 4, "General", "AGRPERIOD", agrPeriod & "[Tab]")     
       ' Շջանցման ուղղություն դաշտի լրացում
            Call Rekvizit_Fill("Document", 4 ,"General", "PASSOVDIRECTION", passDirection)
       ' Գումարների ամսաթվերի ընտրություն դաշտի լրացում
            Call Rekvizit_Fill("Document", 4 ,"General", "SUMSDATESFILLTYPE", summDateSelect)
       ' Գումարների բաշխման ձև դաշտի լրացում
            Call Rekvizit_Fill("Document", 4, "General", "SUMSFILLTYPE", summFillType) 
       ' Օվերդրաֆտի Տոկոսադրույք  դաշտի լրացում 
            Call Rekvizit_Fill("Document", 6, "General", "PCAGR", overRates & "[Tab]" & overRatesSect ) 
       ' Չոգտագործված մասի տոկոսադրույք դաշտի լրացում
            Call Rekvizit_Fill("Document", 6, "General", "PCNOCHOOSE", unusedPortRate & "[Tab]" & unusedPortRateSec )
       ' Ճուղայնություն դաշտի լրացում
            Call Rekvizit_Fill("Document", 8, "General", "SECTOR", sect)
       ' Ճուղայնություն(Նոր ՎՌ)  դաշտի լրացում
            Call Rekvizit_Fill("Document", 8, "General", "USAGEFIELD", sectNew)    
       ' Նպատակ դաշտի լրացում
            Call Rekvizit_Fill("Document", 8, "General", "AIM", purpose)
       ' Ծրագիր դաշտի լրացում
            Call Rekvizit_Fill("Document", 8, "General", "SCHEDULE", mShedule)
       ' Երաշխավորություն դաշտի լրացում
            Call Rekvizit_Fill("Document", 8, "General", "GUARANTEE", mGuarantee)
       ' Երկիր դաշտի լրացում
            Call Rekvizit_Fill("Document", 8, "General", "COUNTRY", mCountry)
       ' Մարզ դաշտի լրացում
            Call Rekvizit_Fill("Document", 8, "General", "LRDISTR", lRegion)
       ' Մարզ(Նոր ՎՌ) դաշտի լրացում
            Call Rekvizit_Fill("Document", 8, "General", "REGION", mRegion)
       ' Պայմ. թղթային N դաշտի լրացում
            Call Rekvizit_Fill("Document", 8, "General", "PPRCODE", paperCode)
       ' Կատարել կոճակի սեղմում
            Call ClickCmdButton(1, "Î³ï³ñ»É")
          
        Set tdbgViewn  =  wMDIClient.VBObject("frmPttel").VBObject("tdbgView")
        BuiltIn.Delay(1000)
        
        ' Ստուգում որ Վարկային պայմանագիրը ստեղծվել է
        If  tdbgViewn.ApproxCount <> 1 Then
             Log.Error("Վարկային պայմանագիրը չի ստեղծվել")
             Exit Sub
        End If
          
End Sub

' Գրաֆիկով Օվերդրաֆտի պայմանագիրն ուղարկել հաստատման
Function SendToApprove(contractName)

      Dim tdbgView, status
      status = False
      Set tdbgView = wMDIClient.VBObject("frmPttel").VBObject("tdbgView")
       
      Do until  tdbgView.EOF
           If  Trim(tdbgView.Columns.Item(0).Value) = Trim(contractName) Then
                  ' Կատարել բոլոր գործողությունները
                  Call wMainForm.MainMenu.Click(c_AllActions)
                  ' Ուղարկել հաստատման գործողության կատարում
                  Call wMainForm.PopupMenu.Click(c_SendToVer)
                  ' Այո կոճակի սեղմում
                  Call ClickCmdButton(5, "²Ûá")
                  status = True 
                  BuiltIn.Delay(2000)
                  Call Close_Pttel("frmPttel")
                  Exit Do
           Else
                  tdbgView.MoveNext   
           End If     
      Loop
  
      SendToApprove = status

End Function


' Օվերդրաֆտի տրամադրում
' giveMoneyISN - Օվերդրաֆտի տրամադրում փաստաթղթի ISN
' docNumOut - Օվերդրաֆտի տրամադրում փաստաթղթի ISN
' ovDate - Ամսաթիվ
' overSumm - Գումար
' cashOrNo - Կանխիկ/Անկանխիկ
' accCorr - Հաշիվ
Sub GiveOverdraft(giveMoneyISN, docNumOut, ovDate, overSumm, cashOrNo, accCorr)
      
         ' Կատարել Գործողություններ/Բոլոր Գործողություններ
         Call wMainForm.MainMenu.Click(c_AllActions)
         ' Կատարել Օվերդրաֆտի տրամադրում
         Call wMainForm.PopupMenu.Click(c_Opers  & "|" & c_GiveAndBack & "|" & c_GiveOverdraft)  
               
         ' Գումարի տրամադրում պատուհանի բացման ստուգում
         If Not wMDIClient.WaitVBObject("frmASDocForm",2000).Exists Then   
                    Log.Error("Գումարի տրամադրում պատուհանը չի բացվել")
                    Exit Sub
         End If
          
         ' Գումարի տրամադրում փաստաթղթի ISN - ի ստացում
         giveMoneyISN = wMDIClient.VBObject("frmASDocForm").DocFormCommon.Doc.ISN
           
         ' Ամսաթիվ դաշտի լրացում
         Call Rekvizit_Fill("Document", 1, "General", "DATE", ovDate)
         ' Գումար դաշտի լրացում
         Call Rekvizit_Fill("Document", 1, "General", "SUMMA", overSumm)
         ' Կանխիկ/Անկանխիկ դաշտի լրացում
         Call Rekvizit_Fill("Document", 1, "General", "CASHORNO", cashOrNo)
         ' Հաշիվ դաշտի լրացում
         Call Rekvizit_Fill("Document", 1, "General", "ACCCORR", accCorr)
         ' Կատարել կոճակի սեղմում
         Call ClickCmdButton(1, "Î³ï³ñ»É")
          ' Այո կոճակի սեղմում
         Call ClickCmdButton(5, "²Ûá")

End Sub

' Հիշարար օրդերի ստեղծում
' mDate - Ժամանակահատվածի սկիզբ
' dateAgr - Ժամանակահատվածի ավարտ
' orderISN - Հիշարար օրդերի ISN
' orderNum - Հիշարար օրդերի փաստաթղթի համարը
' ordDate - Ամսաթիվ
' ordD - Հաշիվ դեբետ
' ordC - Հաշիվ կրեդիտ 
' ordMoney - Գումար
' aim - Նպատակ
Sub CreateMemOrders(orderISN, orderNum, ordDate, AccDb, AccCr, ordMoney, aim)
        
         BuiltIn.Delay(3000)
         
         ' Գործողություններ / Բոլոր Գործողությունները
         Call wMainForm.MainMenu.Click(c_AllActions)
         ' Հիշարար օրդեր գործողության կատարում
         Call wMainForm.PopupMenu.Click(c_MemOrds  & "|" & c_MemOrd) 
         
         ' Ստուգում Հիշարար օրդեր պատուհանը բացվել է թէ ոչ
         If Not wMDIClient.WaitVBObject("frmASDocForm",10000).Exists Then
               Log.Error("Հիշարար օրդեր պատուհանը չի բացվել")
               Exit Sub
         End If
        
         BuiltIn.Delay(700)
         
         ' ISN-ի ստացում        
         orderISN = wMDIClient.VBObject("frmASDocForm").DocFormCommon.Doc.ISN
         
          ' Պայմանագրի համարի ստացում
         orderNum = Get_Rekvizit_Value("Document",1,"General","DOCNUM")
             
         ' Ամսաթիվ դաշտի լրացում
         Call Rekvizit_Fill("Document", 1, "General", "DATE", ordDate)
         ' Հաշիվ դեբետ դաշտի լրացում 
         Call Rekvizit_Fill("Document", 1, "General", "ACCDB", AccDb)
         ' Հաշիվ կրեդիտ դաշտի լրացում
         Call Rekvizit_Fill("Document", 1, "General", "ACCCR", AccCr)
         ' Գումար դաշտի լրացում
         Call Rekvizit_Fill("Document", 1, "General", "SUMMA", ordMoney)
         ' Նպատակ դաշտի լրացում
         Call Rekvizit_Fill("Document", 1, "General", "AIM", aim)
         ' Կատարել կոճակի սեղմում
         Call ClickCmdButton(1, "Î³ï³ñ»É")
         
         wMDIClient.VBObject("FrmSpr").Close
         ' Գործողություններ / Բոլոր գործողություններ       
         Call wMainForm.MainMenu.Click(c_AllActions)
         ' Հաշվառել գործողության կատարում
         Call wMainForm.PopupMenu.Click(c_DoTrans)
         ' Այո կոճակի սեղմում
         Call ClickCmdButton(5, "²Ûá")
         BuiltIn.Delay(700)
        
         Call Close_Pttel("frmPttel")
        
End Sub


' խմբային մարում
' calcDate - Հաշվարկման ամսաթիվ
' regDate - Ձևակերպման ամսաթիվ
Sub OverdraftGroupRepayment(startDate, endDate)

        If  Not wMDIClient.WaitVBObject("frmPttel",2000).Exists Then
               Log.Error("Օվերդրաֆտներ թղթապանակը չի բացվել")
               Exit Sub
        End If
        
        wMDIClient.VBObject("frmPttel").Keys("[Ins]")
        ' Կատարել բոլոր գործողությունները
        Call wMainForm.MainMenu.Click(c_AllActions)
        ' Կատարել Խմբային մարում գործողությունը
        Call wMainForm.PopupMenu.Click(c_GroupDebt )
        
        ' Ստուգում որ Խմբային մարում պատուհանը բացվել է
        If Not p1.WaitVBObject("frmAsUstPar",2000).Exists Then
               Log.Error("Խմբային մարում պատուհանը չի բացվել")
               Exit Sub
        End If
        
        ' Հաշվարկման ամսաթիվ դաշտի լրացում
        Call Rekvizit_Fill("Dialog", 1, "General", "START", startDate)
        ' Ձևակերպման ամսաթիվ դաշտի լրացում
        Call Rekvizit_Fill("Dialog", 1, "General", "END", endDate)
        ' Կատարել կոճակի սեղմում
        Call ClickCmdButton(2, "Î³ï³ñ»É")  
        BuiltIn.Delay(2000)
        ' Այո կոճակի սեղմում
        Call ClickCmdButton(5, "²Ûá")  

End Sub


' Օվերդրաֆտի խմբային տրամադրում
' startD - Օվերդրաֆտի տրամադրման սկիզբ
' endD - Օվերդրաֆտի տրամադրման ավարտ
 Sub  GiveOverdraftGroup(startD, endD)
 
        ' Ստուգում որ Օվերդրաֆտներ թղթապանակը բացվել է
        If  Not wMDIClient.WaitVBObject("frmPttel",2000).Exists Then
              Log.Error("Օվերդրաֆտներ թղթապանակը չի բացվել")
              Exit Sub
        End If
                
        wMDIClient.VBObject("frmPttel").Keys("[Ins]")
        ' Կատարել բոլոր գործողությունները
        Call wMainForm.MainMenu.Click(c_AllActions)
        ' Կատարել խմբային տրամադրում
        Call wMainForm.PopupMenu.Click(c_GroupGive )
        
        ' Ստուգում որ Խմբային տրամադրում պատուհանը բացվել է
        If Not p1.VBObject("frmAsUstPar").Exists Then
              Log.Error("Խմբային տրամադրում պատուհանը չի բացվել")
              Exit Sub
        End If
        
        ' Օվերդրաֆտի տրամադրման սկիզբ
        Call Rekvizit_Fill("Dialog", 1, "General", "START", startD)
        ' Օվերդրաֆտի տրամադրման ավարտ
        Call Rekvizit_Fill("Dialog", 1, "General", "END", endD)
        ' Կատարել կոճակի սեղմում
        Call ClickCmdButton(2, "Î³ï³ñ»É")
        ' Կատարել կոճակի սեղմում
        Call ClickCmdButton(5, "²Ûá")
        
 End Sub
 
 
' խմբային տոկոսների հաշվարկ
' calcDate - Հաշվարկման ամսաթիվ
' regDate - Ձևակերպման ամսաթիվ
' checkCount - Տոկոսների հաշվարկ
Sub InterestGroupCalculationOverdraft (calcDate,regDate, checkCount)

        wMDIClient.VBObject("frmPttel").Keys("[Ins]")
        
        ' Կատարել բոլոր գործողությունները
        Call wMainForm.MainMenu.Click(c_AllActions)
        ' Կատարել Խմբային հաշվարկ գործողությունը
        Call wMainForm.PopupMenu.Click(c_GroupCalc )
        
        ' Հաշվարկման ամսաթիվ դաշտի լրացում
        Call Rekvizit_Fill("Dialog", 1, "General", "CLOSEDATE", calcDate)
        ' Ձևակերպման ամսաթիվ դաշտի լրացում
        Call Rekvizit_Fill("Dialog", 1, "General", "SETDATE", regDate)
        ' Տոկոսների հաշվարկ դաշտի լրացում
        Call Rekvizit_Fill("Dialog", 1, "CheckBox", "CHG", checkCount)
        ' Կատարել կոճակի սեղմում
        Call ClickCmdButton(2, "Î³ï³ñ»É")  
        ' Այո կոճակի սեղմում
        Call ClickCmdButton(5, "²Ûá")  
                
End Sub

' ISN-ի ստացում
' calcDate - Ժամանակահատվածի սկիզբ/ավարտ
' dateType - Օգտագործող/ փաստաթղթի տեսակ
' insGrISN - Պայմանագրի ISN
Sub  GetDocISN(paramN, calcDate, status, dateType, insGrISN)

        BuiltIn.Delay(2000) 
        Call wMainForm.MainMenu.Click(c_AllActions)
        Call wMainForm.PopupMenu.Click(paramN)
         
        ' Ժամանակահատվածի սկիզբ դաշտի լրացում
        Call Rekvizit_Fill("Dialog", 1, "General", "START", calcDate )
        ' Ժամանակահատվածի ավարտ դաշտի լրացում
        Call Rekvizit_Fill("Dialog", 1, "General", "END", calcDate)
         
        If  status Then
              ' Կատարող դաշտի լրացում
              Call Rekvizit_Fill("Dialog", 1, "General", "USER", dateType)
        Else
              ' Տիպ դաշտի լրացում
              Call Rekvizit_Fill("Dialog", 1, "General", "DEALTYPE", dateType)
        End If
        
        ' Կատարել կոճակի սեղմում
        Call ClickCmdButton(2, "Î³ï³ñ»É")  
        BuiltIn.Delay(2000) 
        
        ' Գործողություններ / Բոլոր գործողություններ
        Call wMainForm.MainMenu.Click(c_AllActions)
        ' Խմբագրել գործողության կատարում
        Call wMainForm.PopupMenu.Click(c_ToEdit)
        BuiltIn.Delay(1000) 
        
        ' Փաստաթղթի ISN - ի ստացում
        insGrISN = wMDIClient.VBObject("frmASDocForm").DocFormCommon.Doc.ISN
        
        wMDIClient.VBObject("frmASDocForm").Close
        Call Close_Pttel("frmPttel_2")

End Sub


' խմբային տոկոսների հաշվարկի ջնջում
' dateGive - ժամանակահատվածի սկիզբ
' dateAgr - Ժամանակահատվածի ավարտ
' userName - Ում կողմից
' dateType - Փաստաթղթի տեսակ
Sub DeleteActionOverdraft(param, dateGive, dateAgr, status, dateType ) 

        ' Գործողություններ / Բոլոր գործողություններ
        Call wMainForm.MainMenu.Click(c_AllActions)
        Call wMainForm.PopupMenu.Click(param)
          
        ' Ժամանակահատվածի սկիզբ դաշտի լրացում
        Call Rekvizit_Fill("Dialog",1, "General", "START", dateGive )
        ' Ժամանակահատվածի ավարտ դաշտի լրացում
        Call Rekvizit_Fill("Dialog",1, "General", "END", dateAgr)
        
        If  status Then
             Call Rekvizit_Fill("Dialog",1, "General", "DEALTYPE", "^A[Del]" & dateType)
        End If
        
        ' Կատարել կոճակի սեղմում
        Call ClickCmdButton(2, "Î³ï³ñ»É")  
          
        BuiltIn.Delay(2000) 
        ' Քանի դեռ թղթապանակում տողերի քանակը հավասար չէ զրոյի
        Do While wMDIClient.VBObject("frmPttel_2").VBObject("tdbgView").VisibleRows <> 0  
               ' Ջնջել գործողության կատարում
               Call wMainForm.MainMenu.Click(c_Opers & "|" & c_Delete)
               BuiltIn.Delay(1000) 
               If  p1.WaitVBObject("frmAsMsgBox", delay_small).Exists Then
                    ' Այո կոճակի սեղմում  
                    Call ClickCmdButton(5, "²Ûá") 
                    ' Այո կոճակի սեղմում  
                    Call ClickCmdButton(3, "²Ûá") 
                    BuiltIn.Delay(6000) 
               Else 
                    ' Այո կոճակի սեղմում  
                    Call ClickCmdButton(3, "²Ûá") 
                    BuiltIn.Delay(6000) 
               End If
               
        Loop
         
        Call Close_Pttel("frmPttel_2")
        BuiltIn.Delay(2000) 
         
End Sub


' Ջնջում գործողությունների կատարում (Ջնջել հիշարար օրդերը)
' creatDate - ժամանակահատվածի սկիզբ/ ավարտ
Sub DeleteMemOrderFromRegPayment(creatDate)  

        ' Մուտք Հաշվառված վճարային փաստաթղթեր թղթապանակ
        Call wTreeView.DblClickItem("|Ð³×³Ëáñ¹Ç ëå³ë³ñÏáõÙ ¨ ¹ñ³Ù³ñÏÕ |Ð³ßí³éí³Í í×³ñ³ÛÇÝ ÷³ëï³ÃÕÃ»ñ") 
        
        Dim  tdbgView

        ' Ստուգում որ Վճարային փաստաթղթեր դիալոգը բացվել է
        If Not p1.WaitVBObject("frmAsUstPar",2000).Exists Then
                Log.Message("Վճարային փաստաթղթեր դիալոգը չի բացվել")
                Exit Sub
        End If
       
        ' Ժամանակահատվածի սկիզբ դաշտի լրացում
        Call Rekvizit_Fill("Dialog", 1, "General", "PERN", creatDate )
        ' Ժամանակահատվածի ավարտ դաշտի լրացում
        Call Rekvizit_Fill("Dialog", 1, "General", "PERK", creatDate)
        ' Կատարել կոճակի սեղմում
        Call ClickCmdButton(2, "Î³ï³ñ»É")
        
        BuiltIn.Delay(1200)   
        ' Ստուգում որ Վճարային փաստաթղթեր թղթապանակը բացվել է
        If Not wMDIClient.WaitVBObject("frmPttel",2000).Exists Then
                Log.Message("Վճարային փաստաթղթեր թղթապանակը չի բացվել")
                Exit Sub       
        End If
         
        Set tdbgView = wMDIClient.VBObject("frmPttel").VBObject("tdbgView")
         
        ' Քանի դեռ թղթապանակում տողերի քանակը հավասար չէ զրոյի
        Do While tdbgView.VisibleRows <> 0  
        tdbgView.MoveLast 
               ' Ջնջել գործողության կատարում
               Call wMainForm.MainMenu.Click(c_Opers & "|" & c_Delete)
               If  Sys.Process("Asbank").WaitVBObject("frmAsMsgBox", delay_small).Exists Then
                    ' Այո կոճակի սեղմում  
                    Call ClickCmdButton(5, "²Ûá") 
                    ' Այո կոճակի սեղմում  
                    Call ClickCmdButton(3, "²Ûá") 
                    BuiltIn.Delay(1000) 
               Else 
                    ' Այո կոճակի սեղմում  
                    Call ClickCmdButton(3, "²Ûá") 
                    BuiltIn.Delay(1000) 
               End If
        Loop
         
        BuiltIn.Delay(1000) 
        wMainForm.Window("MDIClient", "", 1).VBObject("frmPttel").Close
         
End Sub


' Վճարող կազմակերպությունների կարգավորումներ - կլաս
Class SettingsForPayerCompanies
        Public payerCompISN
        Public partCode
        Public cliCode
        Public partName
        Public eName
        Public partKey
        Public ourKey
        Public paySys
        Public dateClose
        Public wCurr()
        Public lowerLimit()
        Public upperLimit()
        Public getLoan
        Public getLoanP
        Public getByPas
        Public showBaseSum
        Public payLoan
        Public lPayLoan
        Public payloanCnt
        Public accDebt()
        Public getCredit
        Public showPhone
        Public showBalanceCr
        Public contrProvision
        Public dailyAction
        Public accDebt2()
        Public getAcc
        Public getAccByCard
        Public getAccP
        Public showRestr
        Public showPhones
        Public showBalance
        Public fillAcc
        Public fillAccP
        Public addCashAcc
        Public sillAccCnt
        Public wTransfer
        Public transferEx
        Public transferCnt
        Public accDebt4
        Public getCard
        Public getExtCard
        Public cardholdersName
        Public tf2Card
        Public tf2ExtCard
        Public tf2CardCnt
        Public accDebt5
        Public cashOut
        Public cashAcs
        Public expiredSecond
        Public chashOutCnt
        Public cashOutByPhn
        Public cashAscByPhn
        Public expiredMinute
        Public dailyAmountLimit
        Public accCredit
        Public accDebt6
        Public termID
        Public gridRow1
        public gridRow2
        public gridRow5
        
        Private Sub Class_Initialize
              payerCompISN = ""
              partCode = ""
              cliCode = ""
              partName = ""
              eName = ""
              partKey = ""
              ourKey = ""
              paySys = "ê"
              dateClose = ""
              gridRow1 = rowCount_1
              gridRow2 = rowCount_2
              gridRow5 = rowCount_5
              ReDim wCurr(gridRow1)
              ReDim lowerLimit(gridRow1)
              ReDim upperLimit(gridRow1)
              ReDim accDebt2(gridRow2)
              ReDim accDebt(gridRow5)
              getLoan = False
              getLoanP = False
              getByPas = False
              showBaseSum = False
              payLoan = False
              lPayLoan = False
              payloanCnt = ""
              getCredit = False
              showPhone = False
              showBalanceCr = False
              contrProvision = False
              dailyAction = ""
              getAcc = False
              getAccByCard = False
              getAccP = False
              showRestr = False
              showPhones = False
              showBalance = False
              fillAcc = False
              fillAccP = False
              addCashAcc = False
              sillAccCnt = ""
              wTransfer = False
              transferEx = False
              transferCnt = ""
              accDebt4 = ""
              getCard = False
              getExtCard = False
              cardholdersName = False
              tf2Card = False
              tf2ExtCard = False
              tf2CardCnt = ""
              accDebt5 = ""
              cashOut = False
              cashAcs = False
              expiredSecond = ""
              chashOutCnt = ""
              cashOutByPhn = False
              cashAscByPhn = False
              expiredMinute = ""
              dailyAmountLimit = ""
              accCredit = ""
              accDebt6 = ""
              termID = ""
        End Sub
End Class

Function New_SettingsForPayerCompanies(rowCount1, rowCount2, rowCount5)
    rowCount_1 = rowCount1
    rowCount_2 = rowCount2
    rowCount_5 = rowCount5
    Set New_SettingsForPayerCompanies = NEW SettingsForPayerCompanies      
End Function


' Վճարող կազմակերպությունների կարգավորումներ - sub
Sub Fill_SettingsForPayerCompanies(PayerCompanies)
      Dim tabN
      Call wTreeView.DblClickItem("|ÐÌ-ì×³ñáõÙÝ»ñ ²Þî|Üáñ ÷³ëï³ÃÕÃ»ñ|ì×³ñáÕ Ï³½Ù³Ï»ñåáõÃÛáõÝÝ»ñÇ Ï³ñ·³íáñáõÙÝ»ñ") 
      
      If wMDIClient.WaitVBObject("frmASDocForm", 2000).Exists Then
            ' Վճարող կազմակերպությունների կարգավորումներ փաստաթղթի ISN - ի ստացում
             PayerCompanies.payerCompISN = wMDIClient.VBObject("frmASDocForm").DocFormCommon.Doc.ISN
             tabN = 1
            ' Գործընկեր դաշտի լրացում
            Call Rekvizit_Fill("Document", tabN, "General", "PARTCODE", PayerCompanies.partCode)
            ' Հաճախորդ դաշտի լրացում
            Call Rekvizit_Fill("Document", tabN, "General", "CLICODE", PayerCompanies.cliCode)
            ' Անվանում դաշտի լրացում
            Call Rekvizit_Fill("Document", tabN, "General", "PARTNAME", PayerCompanies.partName)
            ' Անգլերեն անվանում դաշտի լրացում
            Call Rekvizit_Fill("Document", tabN, "General", "PARTENAME", PayerCompanies.eName)
            ' Գործընկերոջ բանալին դաշտի լրացում
            Call Rekvizit_Fill("Document", tabN, "General", "PARTKEY", PayerCompanies.partKey)
            ' Մեր բանալին դաշտի լրացում
            Call Rekvizit_Fill("Document", tabN, "General", "OURKEY", PayerCompanies.ourKey)
            ' Օգտագործվող վճարային համակարգ դաշտի լրացում
            Call Rekvizit_Fill("Document", tabN, "General", "PAYSYS", PayerCompanies.paySys)
            ' Փակման ամսաթիվ դաշտի լրացում
            Call Rekvizit_Fill("Document", tabN, "General", "DATECLOSE","^!A[Del]" &  PayerCompanies.dateClose)

            ' Սահմանաչափեր գրիդի լրացում   
            For i = 0 To PayerCompanies.gridRow1 - 1
              With wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame").VBObject("DocGrid")
                ' Արժույթ դաշտի լրացում
                .Row = i
                .Col = 0
                .Keys(PayerCompanies.wCurr(i) & "[Enter]")
                ' Ստորին սահման դաշտի լրացում
                .Col = 1
                .Keys(PayerCompanies.lowerLimit(i) & "[Enter]")
                ' Վերին սահման դաշտի լրացում
                .Col = 2
                .Keys(PayerCompanies.upperLimit(i) & "[Enter]")
              End With 
            Next
            
            ' Անցնել 2. Պայմ. մարում բաժին
            tabN = tabN + 1
            ' Թույլատրել պայմանագրի փնտրումը դաշտի լրացում
            Call Rekvizit_Fill("Document", tabN, "CheckBox", "GETLOAN", PayerCompanies.getLoan)
            If PayerCompanies.getLoan Then
                  ' Անձը հաստատող փաստ. պարտադիր է դաշտի լրացում
                  Call Rekvizit_Fill("Document", tabN, "CheckBox", "GETLOANP", PayerCompanies.getLoanP)
                  ' Թույլատրել պայմանագրերի փնտրումը անձը հասատտող փաստ.-ով դաշտի լրացում
                  Call Rekvizit_Fill("Document", tabN, "CheckBox", "GETBYPAS", PayerCompanies.getByPas)
                  ' Ցույց տալ պայմանագրի գումարը դաշտի լրացում
                  Call Rekvizit_Fill("Document", tabN, "CheckBox", "SHOWBASESUMS", PayerCompanies.showBaseSum)
            End If
            ' Թույլատրել պարտքերի մարումը դաշտի լրացում
            Call Rekvizit_Fill("Document", tabN, "CheckBox", "PAYLOAN", PayerCompanies.payLoan)
            If PayerCompanies.payLoan Then
                  ' Մարելիս անձը հաստատող փաստաթուղթը պարտադիր է դաշտի լրացում
                  Call Rekvizit_Fill("Document", tabN, "CheckBox", "PAYLOANP", PayerCompanies.lPayLoan)
                  ' Օրական գործողությունների քանակ (նույն պայմանագրի համար) դաշտի լրացում
                  Call Rekvizit_Fill("Document", tabN, "General", "PAYLOANCNT", PayerCompanies.payloanCnt)
                  For i = 0 To PayerCompanies.gridRow5 - 1
                        With wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame_2").VBObject("DocGrid_2")
                          ' Լրացնել հաշիվ դեբետ դաշտը
                          .Row = i
                          .Col = 0
                          .Keys(PayerCompanies.accDebt(i) & "[Enter]")
                        End With 
                  Next
            End If
            
            ' Անցնել 3. Պայմ. տրամադրում բաժին
            tabN = tabN + 1
            ' Թույլատրել պայմանագրերի փնտրումը դաշտի լրացում
            Call Rekvizit_Fill("Document", tabN, "CheckBox", "GETCREDIT", PayerCompanies.getCredit)
            If PayerCompanies.getCredit Then
                    ' Ցույց տալ հեռախոսահամարները դաշտի լրացում
                    Call Rekvizit_Fill("Document", tabN, "CheckBox", "SHOWPHONESCR", PayerCompanies.showPhone)
                    ' Ցույց տալ մնացորդը դաշտի լրացում
                    Call Rekvizit_Fill("Document", tabN, "CheckBox", "SHOWBALANCECR", PayerCompanies.showBalanceCr)
            End If
            ' Թույլատրել պայմանգրի տրամադրումը դաշտի լրացում
            Call Rekvizit_Fill("Document", tabN, "CheckBox", "CRDISB", PayerCompanies.contrProvision)
            If PayerCompanies.contrProvision Then
                  ' Օրական գործողությունների քանակ(նույն պայմանագրի համար) դաշտի լրացում
                  Call Rekvizit_Fill("Document", tabN, "General", "CRDISBCNT", PayerCompanies.dailyAction)
                  For  i = 0 To PayerCompanies.gridRow2 -1 
                      With wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame_2").VBObject("DocGrid_3")
                         ' Լրացնել հաշիվ դեբետ դաշտը
                        .Row = i
                        .Col = 0
                        .Keys(PayerCompanies.accDebt2(i) & "[Enter]")
                      End With 
                  Next
           End If
            
             ' Անցնել 4. Հաշիվների գործարքներ բաժին
            tabN = tabN + 1
            ' Թույլատրել հաշվի փնտրումը դաշտի լրացում
            Call Rekvizit_Fill("Document", tabN, "CheckBox", "GETACC", PayerCompanies.getAcc)
            If PayerCompanies.getAcc Then
                  ' Թույլատրել հաշվի փնտրումը քարտի համարով դաշտի լրացում
                  Call Rekvizit_Fill("Document", tabN, "CheckBox", "GETACCBYCARD", PayerCompanies.getAccByCard)
                  ' Անձ հաստատող փաստ. պարտադիր է դաշտի լրացում
                  Call Rekvizit_Fill("Document", tabN, "CheckBox", "GETACCP", PayerCompanies.getAccP)
                  ' Ցույց տալ սահմանափակումները դաշտի լրացում
                  Call Rekvizit_Fill("Document", tabN, "CheckBox", "SHOWRESTR", PayerCompanies.showRestr)
                  ' Ցույց տալ հեռախոսահամարները դաշտի լրացում
                  Call Rekvizit_Fill("Document", tabN, "CheckBox", "SHOWPHONES", PayerCompanies.showPhones)
                  ' Ցույց տալ մնացորդը դաշտի լրացում
                  Call Rekvizit_Fill("Document", tabN, "CheckBox", "SHOWBALANCE", PayerCompanies.showBalance)
            End If
            ' Թույլատրել հաշվի համալրումը դաշտի լրացում
            Call Rekvizit_Fill("Document", tabN, "CheckBox", "FILLACC", PayerCompanies.fillAcc)
            If PayerCompanies.fillAcc Then
                  ' Անձը հաստատող փաստ. պարտադիր է դաշտի լրացում
                  Call Rekvizit_Fill("Document", tabN, "CheckBox", "FILLACCP", PayerCompanies.fillAccP)
                  ' Ավելացնել կանխիկի հաշվառումը դաշտի լրացում
                  Call Rekvizit_Fill("Document", tabN, "CheckBox", "CASHAC", PayerCompanies.addCashAcc)
                  ' Օրական գործողությունների քանակ(նույն հաշվի համար) դաշտի լրացում
                  Call Rekvizit_Fill("Document", tabN, "General", "FILLACCCNT", PayerCompanies.sillAccCnt)
            End If
            If PayerCompanies.fillAcc Then
                With wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame_4").VBObject("DocGrid_4")
                  ' Հաշիվ դեբետ դաշտի լրացում
                  .Row = 0
                  .Col = 0
                  .Keys(PayerCompanies.accDebt4 & "[Enter]")
                End With 
            End If
            
            ' Անցնել 5. Փոխանցման գործարքներ բաժին
            tabN = tabN + 1
            ' Թույլատրել հաշվից փոխանցումը դաշտի լրացում
            Call Rekvizit_Fill("Document", tabN, "CheckBox", "TRANSFER", PayerCompanies.wTransfer)
            ' Ստեղծել Վճ. Հանձնարարագրի/Արտարժ. փոխանակում դաշտի լրացում
            Call Rekvizit_Fill("Document", tabN, "CheckBox", "TRANSFEREX", PayerCompanies.transferEx)
            If PayerCompanies.wTransfer Then
                  ' Օրական գործողությունների քանակ (նույն հաշիվների համար)  դաշտի լրացում
                  Call Rekvizit_Fill("Document", tabN, "General", "TRANSFERCNT", PayerCompanies.transferCnt)
            End If
            
            ' Անցնել 6. Քարտային փոխանցման գործարքներ բաժին
            tabN = tabN + 1
            ' Թույլատրել քարտի փնտրումը դաշտի լրացում
            Call Rekvizit_Fill("Document", tabN, "CheckBox", "GETCARD", PayerCompanies.getCard)
            If PayerCompanies.getCard Then
                  ' Թույլատրել արտաքին քարտի փնտրումը դաշտի լրացում
                  Call Rekvizit_Fill("Document", tabN, "CheckBox", "GETEXTCARD", PayerCompanies.getExtCard)
                  ' Քարտապանի անունը պարտադիր է դաշտի լրացում
                  Call Rekvizit_Fill("Document", tabN, "CheckBox", "FILLEMBNAME", PayerCompanies.cardholdersName)
            End If
            ' Թույլատրել քարտին փոխանցումը դաշտի լրացում
            Call Rekvizit_Fill("Document", tabN, "CheckBox", "TF2CARD", PayerCompanies.tf2Card)
            If PayerCompanies.tf2Card Then
                  ' Թույլատրել արտաքին քարտին փոխանցումը դաշտի լրացում
                  Call Rekvizit_Fill("Document", tabN, "CheckBox", "TF2EXTCARD", PayerCompanies.tf2ExtCard)
                  ' Օրական գործողությունների քանակ (նույն քարտի համար) դաշտի լրացում
                  Call Rekvizit_Fill("Document", tabN, "General", "TF2CARDCNT", PayerCompanies.tf2CardCnt)
                  With wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame_5").VBObject("DocGrid_5")
                  ' Հաշիվ դեբետ դաշտի լրացում
                    .Row = 0
                    .Col = 0
                    .Keys(PayerCompanies.accDebt5 & "[Enter]")
                  End With 
            End If
            
            ' Անցնել 7. Կանխիկացման գործարքներ բաժին
            tabN = tabN + 1
            ' Թույլատրել կանխիկացումը դաշտի լրացում
            Call Rekvizit_Fill("Document", tabN, "CheckBox", "CASHOUT", PayerCompanies.cashOut)
            If PayerCompanies.cashOut Then
                  ' Նվազեցնել կանխիկի մնացորդը դաշտի լրացում
                  Call Rekvizit_Fill("Document", tabN, "CheckBox", "CASHACS", PayerCompanies.cashAcs)
                  ' Կանխիկացման հայտի վավերականության ժամանակահատված (վայրկյան) դաշտի լրացում
                  Call Rekvizit_Fill("Document", tabN, "General", "EXPIREDSCCO", PayerCompanies.expiredSecond)
                  ' Օրական գործողությունների քանակ(նույն հաշվի համար) դաշտի լրացում
                  Call Rekvizit_Fill("Document", tabN, "General", "CASHOUTCNT", PayerCompanies.chashOutCnt)
            End If
            ' Թույլատրել կանխիկացում դաշտի լրացում
            Call Rekvizit_Fill("Document", tabN, "CheckBox", "CASHOUTBYPHN", PayerCompanies.cashOutByPhn)
            If PayerCompanies.cashOutByPhn Then
                   ' Նվազեցնել կանխիկի մնացորդը դաշտի լրացում
                  Call Rekvizit_Fill("Document", tabN, "CheckBox", "CASHACSBYPHN", PayerCompanies.cashAscByPhn)
                  ' Կանխիկացման հայտի օգտագործման ժամանակահատված (րոպե) դաշտի լրացում
                  Call Rekvizit_Fill("Document", tabN, "General", "EXPSCCOBYPHN", PayerCompanies.expiredMinute)
                  ' Գումարի օրական սահմանափակում(Նույն հեռախոսահամարի համար) դաշտի լրացում
                  Call Rekvizit_Fill("Document", tabN, "General", "COBPHNRECEIVERSUMLIMIT", PayerCompanies.dailyAmountLimit)
            End If
            
            If (PayerCompanies.cashOutByPhn or PayerCompanies.cashOutByPhn) Then
                With wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame_6").VBObject("DocGrid_6")
                  ' Հաշիվ կրեդիտ դաշտի լրացում
                  .Row = 0
                  .Col = 0
                  .Keys(PayerCompanies.accCredit & "[Enter]")
                  ' Հաշիվ դեբետ դաշտի լրացում
                  .Col = 2
                  .Keys(PayerCompanies.accDebt6 & "[Enter]")
                  ' Տերմինալ ID դաշտի լրացում
                  .Col = 3
                  .Keys(PayerCompanies.termID & "[Enter]")
                End With 
            End If
            
            Call ClickCmdButton(1, "Î³ï³ñ»É")
      Else 
            Log.Error"Վճարող կազմակերպությունների կարգավորումներ փաստաթուղթը չի բացվել" ,,,ErrorColor
      End If
      
End Sub


' Ստացող կազմակերպությունների կարգավորումներ - կլաս
Class SettingsForRecipientCompanies
        Public recCompISN
        Public partCode
        Public partName
        Public partEName
        Public partKey
        Public endPoint
        Public ourCode
        Public ourKey
        Public paySys
        Public dateClose 
        Public getLoan
        Public getLoanP
        Public getByPas
        Public payLoan
        Public payLoanP
        Public getAcc
        Public getAccByCard
        Public getAccP
        Public fillAcc
        Public fillAccP
        Public actType()
        Public accDbt()
        Public wCur()
        Public chargeType()
        Public actionType()
        Public wCurrency()
        Public wActor()
        Public  incomeAcc()
        Public wOffice()
        Public wPart()
        Public gridRowCount1
        Public gridRowCount2

        Private Sub Class_Initialize
              recCompISN = ""
              partCode = ""
              partName = ""
              partEName = ""
              partKey = ""
              endPoint = ""
              ourCode = ""
              ourKey = ""
              paySys = ""
              dateClose = ""
              getLoan = False
              getLoanP = False
              getByPas = False
              payLoan = False
              payLoanP = False
              getAcc = False
              getAccByCard = False
              getAccP = False
              fillAcc = False
              fillAccP = False
              gridRowCount1 = rowCount_3
              gridRowCount2 = rowCount_4
              ReDim actType(gridRowCount1)
              ReDim accDbt(gridRowCount1)
              ReDim wCur(gridRowCount1)
              ReDim chargeType(gridRowCount1)
              ReDim actionType(gridRowCount2)
              ReDim wCurrency(gridRowCount2)
              ReDim wActor(gridRowCount2)
              ReDim incomeAcc(gridRowCount2)
              ReDim wOffice(gridRowCount2)
              ReDim wPart(gridRowCount2)
        End Sub
End Class

Function New_SettingsForRecipientCompanies(rowCount3, rowCount4)
    rowCount_3 = rowCount3
    rowCount_4 = rowCount4
    Set New_SettingsForRecipientCompanies = NEW SettingsForRecipientCompanies      
End Function


Sub Fill_SettingsForRecipientCompanies(RecipientCompanies)

      Dim tabN
      Call wTreeView.DblClickItem("|ÐÌ-ì×³ñáõÙÝ»ñ ²Þî|Üáñ ÷³ëï³ÃÕÃ»ñ|êï³óáÕ Ï³½Ù³Ï»ñåáõÃÛáõÝÝ»ñÇ Ï³ñ·³íáñáõÙÝ»ñ") 
      
      If wMDIClient.WaitVBObject("frmASDocForm", 2000).Exists Then
            ' Վճարող կազմակերպությունների կարգավորումներ փաստաթղթի ISN - ի ստացում
             RecipientCompanies.recCompISN = wMDIClient.VBObject("frmASDocForm").DocFormCommon.Doc.ISN
             tabN = 1
            ' Գործընկեր դաշտի լրացում
            Call Rekvizit_Fill("Document", tabN, "General", "PARTCODE", RecipientCompanies.partCode)
            ' Անվանում դաշտի լրացում
            Call Rekvizit_Fill("Document", tabN, "General", "PARTNAME", RecipientCompanies.partName)
            ' Անգլերեն անվանում դաշտի լրացում
            Call Rekvizit_Fill("Document", tabN, "General", "PARTENAME", RecipientCompanies.partEName)
            ' Գործընկերոջ բանալին դաշտի լրացում
            Call Rekvizit_Fill("Document", tabN, "General", "PARTKEY", RecipientCompanies.partKey)
            ' Հասցե (EndPoint) դաշտի լրացում
            Call Rekvizit_Fill("Document", tabN, "General", "ENDPOINT", RecipientCompanies.endPoint)
            ' Մեր կոդը գործընկերոջ մոտ դաշտի լրացում
            Call Rekvizit_Fill("Document", tabN, "General", "OURCODE", RecipientCompanies.ourCode)
            ' Մեր բանալին դաշտի լրացում
            Call Rekvizit_Fill("Document", tabN, "General", "OURKEY", RecipientCompanies.ourKey)
            ' Օգտագործվող վճարային համակարգ դաշտի լրացում
            Call Rekvizit_Fill("Document", tabN, "General", "PAYSYS", RecipientCompanies.paySys)
            ' Փակման ամսաթիվ դաշտի լրացում
            Call Rekvizit_Fill("Document", tabN, "General", "DATECLOSE","^!A[Del]" &  RecipientCompanies.dateClose)
            
            ' Անցնել 2. Պայմ. մարում բաժին
            tabN = tabN + 1
            ' Թույլատրել պայմանագրի փնտրումը դաշտի լրացում
            Call Rekvizit_Fill("Document", tabN, "CheckBox", "GETLOAN", RecipientCompanies.getLoan)
             If RecipientCompanies.getLoan Then
                   ' Անձը հաստատող փաստ. պարտադիր է դաշտի լրացում
                    Call Rekvizit_Fill("Document", tabN, "CheckBox", "GETLOANP", RecipientCompanies.getLoanP)
                    ' Թույլատրել պայմանագրերի փնտրումը անձը հաստատող փաստ.-ով դաշտի լրացում
                    Call Rekvizit_Fill("Document", tabN, "CheckBox", "GETBYPAS", RecipientCompanies.getByPas)
             End If
            ' Թույլատրել պարտքերի մարումը դաշտի լրացում
            Call Rekvizit_Fill("Document", tabN, "CheckBox", "PAYLOAN", RecipientCompanies.payLoan)
            If RecipientCompanies.payLoan Then
                    ' Մարելիս անձը հաստատող փաստ. պարտադիր է դաշտի լրացում
                    Call Rekvizit_Fill("Document", tabN, "CheckBox", "PAYLOANP", RecipientCompanies.payLoanP)
            End If
            
            ' Անցնել 3. Հաշիվների գործարքներ բաժին
            tabN = tabN + 1
            ' Թույլատրել հաշվի փնտրումը դաշտի լրացում
            Call Rekvizit_Fill("Document", tabN, "CheckBox", "GETACC", RecipientCompanies.getAcc)
            If RecipientCompanies.getAcc Then
                  ' Թույլատրել հաշվի փնտրումը քարտի համարով դաշտի լրացում
                  Call Rekvizit_Fill("Document", tabN, "CheckBox", "GETACCBYCARD", RecipientCompanies.getAccByCard)
                  ' Անձը հաստատող փաստ. պարտադիր է դաշտի լրացում
                  Call Rekvizit_Fill("Document", tabN, "CheckBox", "GETACCP", RecipientCompanies.getAccP)
            End If
            ' Թույլատրել հաշվի համալրումը դաշտի լրացում
            Call Rekvizit_Fill("Document", tabN, "CheckBox", "FILLACC", RecipientCompanies.fillAcc)
            If RecipientCompanies.fillAcc Then
                  ' Համալրելիս անձը հաստատող փաստ. պարտադիր է դաշտի լրացում
                  Call Rekvizit_Fill("Document", tabN, "CheckBox", "FILLACCP", RecipientCompanies.fillAccP)
            End If
            
            ' Անցնել Թղթակցային հաշիվներ բաժին
            tabN = tabN + 1
            wMDIClient.VBObject("frmASDocForm").VBObject("TabStrip").SelectedItem = wMDIClient.VBObject("frmASDocForm").VBObject("TabStrip").Tabs(tabN)
            For  i = 0 To RecipientCompanies.gridRowCount1 - 1 
                   With wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame_4").VBObject("DocGrid")
                       ' Գործ. տեսակ դաշտի լրացում
                      .Row = i
                      .Col = 0
                      .Keys(RecipientCompanies.actType(i) & "[Enter]")
                      ' Հաշիվ կրեդիտ դաշտի լրացում
                      .Col = 1
                      .Keys(RecipientCompanies.accDbt(i) & "[Enter]")
                      ' Արժույթ դաշտի լրացում
                      .Col = 2
                      .Keys(RecipientCompanies.wCur(i) & "[Enter]")
                      ' Գանձման տեսակ դաշտի լրացում
                      .Col = 3
                      .Keys(RecipientCompanies.chargeType(i))
                   End With 
            Next
 
            ' Անցնել Եկամտի հաշիվներ բաժին
            tabN = tabN + 1
            wMDIClient.VBObject("frmASDocForm").VBObject("TabStrip").SelectedItem = wMDIClient.VBObject("frmASDocForm").VBObject("TabStrip").Tabs(tabN)
            For i = 0 To RecipientCompanies.gridRowCount2 - 1
                     With wMDIClient.VBObject("frmASDocForm").VBObject("TabFrame_5").VBObject("DocGrid_2")
                        ' Գործ. տեսակ դաշտի լրացում
                        .Row = i
                        .Col = 0
                        .Keys(RecipientCompanies.actionType(i) & "[Enter]")
                        ' Արժ. դաշտի լրացում
                        .Col = 1
                        .Keys(RecipientCompanies.wCurrency(i) & "[Enter]")
                        ' Եկամտի հաշիվ դաշտի լրացում
                        .Col = 2
                        .Keys(RecipientCompanies.wActor(i) & "[Enter]")
                        ' Կատարող դաշտի լրացում
                        .Col = 3
                        .Keys(RecipientCompanies.incomeAcc(i) & "[Enter]")
                        ' Գրասենյակ դաշտի լրացում
                        .Col = 4
                        .Keys(RecipientCompanies.wOffice(i) & "[Enter]")
                        ' Բաժին դաշտի լրացում
                        .Col = 5
                        .Keys(RecipientCompanies.wPart(i))
                  End With 
            Next

            ' Սեղմել կատարել կոճակը
            Call ClickCmdButton(1, "Î³ï³ñ»É")
      Else 
            Log.Error"Ստացող կազմակերպությունների կարգավորումներ փաստաթուղթը չի բացվել",,,ErrorColor
      End If
            
End Sub


' ՀԾ-Վճարումներ Պայմանագրի պարտքերի մարում գործողության կատարում
Class RepaymentOfContractDebts
        Public repayContrISN
        Public wPartner
        Public wPassport
        Public creditCode
        Public wContract
        Public docNum
        Public agrCurr
        Public wCurr
        Public wAmount
        Public repayType
        Public wComment
        Public wName
        Public debtType
        Public repaySum
        Public printDetails
        Public payDocNum
        Public payDate
        Public accDb
        Public curDb
        Public wPayer
        Public passNum
        Public pasBy
        Public datePress
        Public dateExpire
        Public wReceiver
        Public wSumma
        Public wAim
        Public accType
        Public chrgAcc
        Public accAMD
        Public payScale
        Public chrgCur
        Public chrgSum
        Public chrgSumAmd
        Public wPercent
        Public incAcc
        Public bankSum
        Public bankSumDv
        Public paySysIn
        Public sencCustPay
        Public wKassa
        Public wBase
        Public inSum
        Public wOffice
        Public dateBirth
        Public birthPlace
        Public wAddress
        Public regCert
        Public taxCode

        Private Sub Class_Initialize
              repayContrISN = ""
               wPartner = ""
               wPassport = ""
               creditCode = ""
               wContract = ""
               docNum = ""
               agrCurr = ""
               wCurr = ""
               wAmount = ""
               repayType = ""
               wComment = ""
               wName = ""
               debtType = ""
               repaySum = ""
               printDetails = False
               payDocNum = ""
               payDate = ""
               accDb = ""
               curDb = ""
               wPayer = ""
               passNum = ""
               pasBy = ""
               datePress = ""
               dateExpire = ""
               wReceiver = ""
               wSumma = ""
               wAim = ""
               accType = ""
               chrgAcc = ""
               accAMD = ""
               payScale = ""
               chrgCur = ""
               chrgSum = ""
               chrgSumAmd = ""
               wPercent = ""
               incAcc = ""
               bankSum = ""
               bankSumDv = ""
               paySysIn = ""
               sencCustPay = ""
               wKassa = ""
               wBase = ""
               inSum = ""
               wOffice = ""
               dateBirth = ""
               birthPlace = ""
               wAddress = ""
               regCert = ""
               taxCode = ""
        End Sub
End Class

Function New_RepaymentOfContractDebts()
    Set New_RepaymentOfContractDebts = NEW RepaymentOfContractDebts      
End Function


Sub Fill_RepaymentOfContractDebts(RepayDebts)

      ' ՀԾ-Վճարումներ -> Պայմանագրի պարտքերի մարում
      Call wMainForm.MainMenu.Click(c_AllActions)
      Call wMainForm.PopupMenu.Click(c_ASPayAction & "|" & c_ASContractRepay)
      
      If Sys.Process("Asbank").WaitVBObject("frmAsUstPar", 2000).Exists Then
            ' "Կազմակերպության կոդ" դաշտի լրացում
            Call Rekvizit_Fill("Dialog", 1, "General", "PARTNER", RepayDebts.wPartner)
            BuiltIn.Delay(4000)
            ' "Պայմանագրում անձը հաստատող փաստ." դաշտի լրացում
            Call Rekvizit_Fill("Dialog", 1, "General", "PASSPORTNO", RepayDebts.wPassport)
            ' "Վարկային կոդ" դաշտի լրացում
            Call Rekvizit_Fill("Dialog", 1, "General", "CNTRCREDITCODE", RepayDebts.creditCode)
            ' "Պայմանագրի N" դաշտի լրացում
            Call Rekvizit_Fill("Dialog", 1, "General", "CONTRACTNO", RepayDebts.wContract)
            Sys.Process("Asbank").VBObject("frmAsUstPar").VBObject("TabFrame").VBObject("cmdButton").Click
            If Sys.Process("Asbank").WaitVBObject("frmModalBrowser",3000).Exists Then
                  Do Until p1.VBObject("frmModalBrowser").VBObject("tdbgView").EOF
                        If  Trim(p1.VBObject("frmModalBrowser").VBObject("tdbgView").Columns.Item(2).Value) = RepayDebts.docNum  Then
                              p1.VBObject("frmModalBrowser").VBObject("tdbgView").Keys("[Enter]")
                              Exit Do
                        Else
                              Log.Error" "& RepayDebts.docNum &" համարով պայմանագիր չկա",,,ErrorColor
                              p1.VBObject("frmModalBrowser").VBObject("tdbgView").MoveNext
                        End If
                  Loop
            End If
            
            ' "Պայմանագրի արժույթ" դաշտի լրացում
            Call Rekvizit_Fill("Dialog", 1, "General", "AGRCURRENCY", RepayDebts.agrCurr)
            ' "Պարտքի տեսակ" դաշտի լրացում
            Call Rekvizit_Fill("Dialog", 1, "General", "DEBTTYPE", RepayDebts.debtType)
            ' "Գործարքի արժույթ" դաշտի լրացում
            Call Rekvizit_Fill("Dialog", 1, "General", "CURRENCY", RepayDebts.wCurr)
            ' "Գումար" դաշտի լրացում
            Call Rekvizit_Fill("Dialog", 1, "General", "AMOUNT", RepayDebts.wAmount)
            ' "Մարման տեսակ" դաշտի լրացում
            Call Rekvizit_Fill("Dialog", 1, "General", "REPAYTYPE", RepayDebts.repayType)
            ' "Մեկնաբանություն" դաշտի լրացում
            Call Rekvizit_Fill("Dialog", 1, "General", "COMMENT", RepayDebts.wComment)
            ' "Անվանում" դաշտի լրացում
            Call Rekvizit_Fill("Dialog", 1, "General", "NAME", RepayDebts.wName)
            ' "Մարման ենթակա գումար" դաշտի լրացում
            Call Rekvizit_Fill("Dialog", 1, "General", "REPAYSUM", RepayDebts.repaySum)
            ' "Տպել լրացուցիչ տվյալներ" դաշտի լրացում
            Call Rekvizit_Fill("Dialog", 1, "CheckBox", "PRINTDETAILS", RepayDebts.printDetails)
            
            ' Սեղմել կատարել կոճակը
            Call ClickCmdButton(2, "Î³ï³ñ»É")
            
            If wMDIClient.WaitVBObject("frmASDocForm",10000).Exists Then
                  ' Պարտքերի մարում փաստաթղթի ISN - ի ստացում
                  RepayDebts.repayContrISN = wMDIClient.VBObject("frmASDocForm").DocFormCommon.Doc.ISN
                  ' Ստանալ փաստաթղթի N դաշտի արժեքը
                  RepayDebts.payDocNum = Get_Rekvizit_Value("Document",1,"General","DOCNUM")
                  ' "Ամսաթիվ" դաշտի լրացում
                  Call Rekvizit_Fill("Document", 1, "General", "DATE", RepayDebts.payDate)
                  ' "Հաշիվ դեբետ" դաշտի լրացում
                  Call Rekvizit_Fill("Document", 1, "General", "ACCDB", RepayDebts.accDb)
                  ' "Վճարման արժույթ" դաշտի լրացում
                  Call Rekvizit_Fill("Document", 1, "General", "CURDB", RepayDebts.curDb)
                  ' "Վճարող" դաշտի լրացում
                  Call Rekvizit_Fill("Document", 1, "General", "PAYER", RepayDebts.wPayer)
                  ' "Անձը հաստատող փաստաթուղթ" դաշտի լրացում
                  Call Rekvizit_Fill("Document", 1, "General", "PASSNUM", RepayDebts.passNum)
                  ' "Տրված" դաշտի լրացում
                  Call Rekvizit_Fill("Document", 1, "General", "PASBY", RepayDebts.pasBy)
                  ' "Վավեր է սկսած" դաշտի լրացում
                  Call Rekvizit_Fill("Document", 1, "General", "DATEPASS", RepayDebts.datePress)
                  ' "Վավեր է մինչև" դաշտի լրացում
                  Call Rekvizit_Fill("Document", 1, "General", "DATEEXPIRE", RepayDebts.dateExpire)
                  ' "Ստացող" դաշտի լրացում
                  Call Rekvizit_Fill("Document", 1, "General", "RECEIVER", RepayDebts.wReceiver)
                  ' "Գումար" դաշտի լրացում
                  Call Rekvizit_Fill("Document", 1, "General", "SUMMA", RepayDebts.wSumma)
                  ' "Նպատակ" դաշտի լրացում
                  Call Rekvizit_Fill("Document", 1, "General", "AIM", RepayDebts.wAim)
                  ' "Հաշվի տիպ" դաշտի լրացում
                  Call Rekvizit_Fill("Document", 1, "General", "ACCTYPE", RepayDebts.accType)
                  
                  ' "Գանձման հաշիվ" դաշտի լրացում, Repay
                  Call Rekvizit_Fill("Document", 2, "General", "CHRGACC", RepayDebts.chrgAcc)
                  ' "Տարանցիկ հաշիվ AMD" դաշտի լրացում
                  Call Rekvizit_Fill("Document", 2, "General", "TCORRACCAMD", RepayDebts.accAMD)
                  ' "Գանձման տեսակ" դաշտի լրացում
                  Call Rekvizit_Fill("Document", 2, "General", "PAYSCALE", RepayDebts.payScale)
                  ' "Արժույթ" դաշտի լրացում
                  Call Rekvizit_Fill("Document", 2, "General", "CHRGCUR", RepayDebts.chrgCur)
                  ' "Գանձման գումար" դաշտի լրացում
                  Call Rekvizit_Fill("Document", 2, "General", "CHRGSUM", RepayDebts.chrgSum)
                  ' "Գանձման գումար AMD" դաշտի լրացում
                  Call Rekvizit_Fill("Document", 2, "General", "CHRGSUMAMD", RepayDebts.chrgSumAmd)
                  ' "Տոկոս" դաշտի լրացում
                  Call Rekvizit_Fill("Document", 2, "General", "PRSNT", RepayDebts.wPercent)
                  ' "Եկամտի հաշիվ" դաշտի լրացում
                  Call Rekvizit_Fill("Document", 2, "General", "INCACC", RepayDebts.incAcc)
                  ' "Միջնորդավճար" դաշտի լրացում
                  Call Rekvizit_Fill("Document", 2, "General", "BANKSUM", RepayDebts.bankSum)
                  ' "Միջնորդավճար AMD" դաշտի լրացում
                  Call Rekvizit_Fill("Document", 2, "General", "BNKSUMAMDV", RepayDebts.bankSumDv)
                  ' "Ընդ. վճ. համակարգ" դաշտի լրացում
                  Call Rekvizit_Fill("Document", 2, "General", "PAYSYSIN", RepayDebts.paySysIn)
                  ' "Ուղարկող մասնակցի վճար" դաշտի լրացում
                  Call Rekvizit_Fill("Document", 2, "General", "ARUSFEESA", RepayDebts.sencCustPay)
                  
                  ' "Դրամարկղ" դաշտի լրացում
                  Call Rekvizit_Fill("Document", 3, "General", "KASSA", RepayDebts.wKassa)
                  ' "Հիմք" դաշտի լրացում
                  Call Rekvizit_Fill("Document", 3, "General", "BASE", RepayDebts.wBase)
                  ' "Փոխանցման նպատակ" դաշտի լրացում
                  Call Rekvizit_Fill("Document", 3, "General", "INSUM", RepayDebts.inSum)
                  ' "Գրասենյակ" դաշտի լրացում
                  Call Rekvizit_Fill("Document", 3, "General", "OFFICE", RepayDebts.wOffice)
                  ' "Ծննդյան ամսաթիվ" դաշտի լրացում
                  Call Rekvizit_Fill("Document", 3, "General", "DATEBIRTH", RepayDebts.dateBirth)
                  ' "Ծննդավայր" դաշտի լրացում
                  Call Rekvizit_Fill("Document", 3, "General", "BIRTHPLACE", RepayDebts.birthPlace)
                  ' "Հասցե" դաշտի լրացում
                  Call Rekvizit_Fill("Document", 3, "General", "ADDRESS", RepayDebts.wAddress)
                  ' "Պետ. գրանցման վկայականի համար" դաշտի լրացում
                  Call Rekvizit_Fill("Document", 3, "General", "REGCERT", RepayDebts.regCert)
                  ' "ՀՎՀՀ(վճարող)" դաշտի լրացում
                  Call Rekvizit_Fill("Document", 3, "General", "TAXCODSD", RepayDebts.taxCode)
                  
                  ' Սեղմել կատարել կոճակը
                  Call ClickCmdButton(1, "Î³ï³ñ»É")
                  If wMDIClient.WaitVBObject("FrmSpr", 2000).Exists Then
                       wMDIClient.VBObject("FrmSpr").Close
                  End If
                  
            Else
                  Log.Error"Վճարման հանձնարարագիր փաստաթուղթը չի բացվել",,,ErrorColor
            End If
      Else
            Log.Error"Պայմանագրի պարտքերի մարում դիալոգը չի բացվել",,,ErrorColor
      End If
      
End Sub



' Ավելացնել §ՀԾ-Վճարումներ¦-ի մասկանից
Class ASPartOfPayments

          Public wCode
          Public wName
          Public eName
          Public wParent
            
          Sub Class_Initialize
                wCode = ""
                wName = ""
                eName = ""
                wParent = ""
          End Sub
            
End Class

Function New_ASPartOfPayments()
        Set New_ASPartOfPayments = New ASPartOfPayments
End Function

Sub Fill_ASPartOfPayments(ASPartOfPayments)

      Call wTreeView.DblClickItem( "|ÐÌ-ì×³ñáõÙÝ»ñ ²Þî|î»Õ»Ï³ïáõÝ»ñ|§ÐÌ-ì×³ñáõÙÝ»ñ¦-Ç Ù³ëÝ³ÏÇóÝ»ñ")
      If wMDIClient.WaitVBObject("frmEditTree", 2000).Exists Then
            ' Կատարել Ավելացնել գործողությունը
            Call wMainForm.MainMenu.Click(c_AllActions)
            Call wMainForm.PopupMenu.Click(c_Add)
      
            If p1.WaitVBObject("frmTreeNode", 2000).Exists Then
                 ' "Կոդ"  դաշտի լրացում
                  Call Rekvizit_Fill("TreeNode", 1, "General", "lblCode", ASPartOfPayments.wCode)
                  ' "Անվանում"  դաշտի լրացում
                  Call Rekvizit_Fill("TreeNode", 1, "General", "lblName", ASPartOfPayments.wName)
                  ' "Անգլերեն Անվանում"  դաշտի լրացում
                  Call Rekvizit_Fill("TreeNode", 1, "General", "lblEName", ASPartOfPayments.eName)
                  ' "Կուտակիչ"  դաշտի լրացում
                  Call Rekvizit_Fill("TreeNode", 1, "General", "lblParent", ASPartOfPayments.wParent)
                  ' Սեղմել կատարել կոճակը
                  Call ClickCmdButton(8, "Î³ï³ñ»É")
                  wMDIClient.VBObject("frmEditTree").Close
            Else
                  Log.Error"Ավելացնել նոր հանգույց պատուհանը չի բացվել",,,ErrorColor
            End If
      Else
                Log.Error"§ՀԾ-Վճարումներ¦-ի մասկանից պատուհանը չի բացվել",,,ErrorColor
      End If
  
End Sub


' Մուտք Մշակման ենթանա մուտքային հաղորդագրություններ թղթապանակ
Class IncMessToBeProcessed

          Public direction
          Public sDate
          Public eDate
          Public sysTem
          Public msgType
          Public customType
          Public cliMask
          Public cliBranch
          Public cliDepart
          Public shorProbMess
          Public showsgDate
          Public showAttach
          Public selectedView
          Public expExcel
            
          Sub Class_Initialize
                direction = ""
                sDate = ""
                eDate = ""
                sysTem = ""
                msgType = ""
                customType = ""
                cliMask = ""
                cliBranch = ""
                cliDepart = ""
                shorProbMess = False
                showsgDate = False
                showAttach = False
                selectedView = "RemMsg"
                expExcel = "0"
          End Sub
            
End Class

Function New_IncMessToBeProcessed()
        Set New_IncMessToBeProcessed = New IncMessToBeProcessed
End Function


Sub Fill_IncMessToBeProcessed(IncMessProcessed)

      Call wTreeView.DblClickItem(IncMessProcessed.direction)
      If p1.WaitVBObject("frmAsUstPar", 2000).Exists Then
                  ' "Ժամանակահատվածի սկիզբ"  դաշտի լրացում
                  Call Rekvizit_Fill("Dialog", 1, "General", "SDATE","^!A[Del]" &  IncMessProcessed.sDate)
                  ' "Ժամանակահատվածի ավարտ"  դաշտի լրացում
                  Call Rekvizit_Fill("Dialog", 1, "General", "EDATE","^!A[Del]" &  IncMessProcessed.eDate)
                  ' "Համակարգ"  դաշտի լրացում
                  Call Rekvizit_Fill("Dialog", 1, "General", "SYSTEM", IncMessProcessed.sysTem)
                  ' "Հաղորդագրության տեսակ"  դաշտի լրացում
                  Call Rekvizit_Fill("Dialog", 1, "General", "MSGTYPE", IncMessProcessed.msgType)
                  ' "Հայտի տեսակ"  դաշտի լրացում
                  Call Rekvizit_Fill("Dialog", 1, "General", "CUSTOMTYPE", IncMessProcessed.customType)
                  ' "Հաճախորդի կոդ"  դաշտի լրացում
                  Call Rekvizit_Fill("Dialog", 1, "General", "CLIMASK", IncMessProcessed.cliMask)
                  ' "Գրասենյակ"  դաշտի լրացում
                  Call Rekvizit_Fill("Dialog", 1, "General", "CLIBRANCH", IncMessProcessed.cliBranch)
                  ' "Բաժին"  դաշտի լրացում
                  Call Rekvizit_Fill("Dialog", 1, "General", "CLIDEPART", IncMessProcessed.cliDepart)
                  ' "Ցույց տալ միայն խնդրահարույց հաղորդագրությունները"  դաշտի լրացում
                  Call Rekvizit_Fill("Dialog", 1, "CheckBox", "SHOWPROBLEMATICS", IncMessProcessed.shorProbMess)
                  ' "Ցույց տալ հաղորդագրության ամսաթիվը"  դաշտի լրացում
                  Call Rekvizit_Fill("Dialog", 1, "CheckBox", "SHOWMSGDATE", IncMessProcessed.showsgDate)
                  ' "Ցույց տալ կցված ֆայլերը"  դաշտի լրացում
                  Call Rekvizit_Fill("Dialog", 1, "CheckBox", "SHOWATTACHEDFILES", IncMessProcessed.showAttach)
                  ' "Դիտելու ձև"  դաշտի լրացում
                  Call Rekvizit_Fill("Dialog", 1, "General", "SELECTED_VIEW", "^A[Del]"  &  IncMessProcessed.selectedView)
                  ' "Լրացնել"  դաշտի լրացում
                  Call Rekvizit_Fill("Dialog", 1, "General", "EXPORT_EXCEL", "^A[Del]"  &  IncMessProcessed.expExcel)
                  
                  ' Սեղմել կատարել կոճակը
                  Call ClickCmdButton(2, "Î³ï³ñ»É")
      Else
                Log.Error"Մշակման ենթանա մուտքային հաղորդագրություններ դիալոգը չի բացվել",,,ErrorColor
      End If
  
End Sub