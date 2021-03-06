'USEUNIT Library_Common
'USEUNIT Library_Contracts
'USEUNIT Online_PaySys_Library
'USEUNIT Mortgage_Library
'USEUNIT Constants

Dim loanAgrClient()
Dim loanAgrNum()
Dim partnerCode()
Dim loanAgrType()
Dim attr
'--------------------------------------------------------------------------------------
'Գրավի պայմանագրի լրացում
'--------------------------------------------------------------------------------------
' agrType - Ստեղծվող պայմանագրի տեսակը (Անվանումը, օրինակ ` " Գրավ(Այլ)" )

' pType - "Պայմանագրի տիպ" դաշտի արժեք
' pNumber - "Պայմանագրի N" դաշտի արժեք
' cliCode - "Գրավատու" դաշտի արժեք
' mortName - "Անվանում" դաշտի արժեք
' fillGrid - True արժեքի դեպքում "Ապահ.պայմ.N" գրիդը լրացվում է , False -ի դեպքում ` ոչ

' loanAgrClient - "Ապահովվով պայմանագիր"  ֆիլտրի "Հաճախորդ " դաշտի արժեք
' loanAgrNum -  "Պայմանագրեր"  ֆիլրտի "Պայմանագրի N "  դաշտի արժեք
' mortCurr - "Արժույթ" դաշտի արժեք
' mortSumma - "Սկզբնական արժեք" դաշտի արժեք
' mortCount - "Սկզբնական քանակ" դաշտի արժեք
' mortComment - "Մեկնաբանություն" դաշտի արժեք
' startDate - "Կնքման ամսաթիվ " դաշտի արժեք
' partnerCode - "Համագրավատուներ" գրիդի "Հաճախորդ " դաշտի արժեք
' partnerName - "Համագրավատուներ" գրիդի "Անվանում " դաշտի արժեք
' transInNB - True արծեքի դեպքում "Հաշվառել ետհաշվեկշռում " նշիչը դրվում է
' fBASE - Գրավի պայմանգրի ISN
' docNumber - Գրավի պայմանագրի համար
' MortSubject - Գրավի առարկա
Sub Mortgage_Doc_Fill (agrType , pType, pNumber, cliCode, mortName, fillGrid, _
                       loanAgrType , loanAgrClient , loanAgrNum, mortCurr , mortSumma, mortCount , mortComment, _
                       startDate, fBASE, docNumber,mortageItemNew, MortSubject)
    
    Dim Count, textMsg, Count1, rowCount
    
    'Գրավի ընտրում "Նոր պայմանագրեր ցուցակից" ցուցակից
    Call wTreeView.DblClickItem("|êï³óí³Í ·ñ³í|Üáñ å³ÛÙ³Ý³·ñÇ ëï»ÕÍáõÙ")
    rowCount = Sys.Process("Asbank").vbObject("frmModalBrowser").vbObject("tdbgView").ApproxCount
    Is_Found = Search_Row(RowCount, agrType)
    If Is_Found Then
        TextMSG = agrType & " is found"
        Call Log_Print_My()
        Log.Message TextMSG, "", pmNormal, attr
    Else
        TextMSG = agrType & " is'n found"
        Call Log_Error_My()
        Log.Error TextMSG, "", pmNormal, attr      
    End If

    'Պայմանագրի ISN - ի վերագրում փոփոխականի
    fBASE = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
    ' Պայմանագրի տիպ դաշտի լրացում
    Call Rekvizit_Fill("Document", 1, "General", "SECTYPE", pType)
    'Պայմանագրի համար դաշտի լրացում
    Call Rekvizit_Fill("Document", 1, "General", "CODE", pNumber)
    'Գրավատու դաշտի լրացում
    Call Rekvizit_Fill("Document", 1, "General", "CLICOD", cliCode)
    BuiltIn.Delay(3000)
    '"Ապահ. պայմ. N " Գրիդի Լրացում
    If fillGrid Then
        
        For Count = 1 To UBound(loanAgrNum)
            'Լրացնել կոճակի սեղմում
            With wMDIClient.vbObject("frmASDocForm").vbObject("TabFrame").vbObject("DocGrid")
              .Col = 3
              .Row = Count -1
              .Keys(" ")
            End With
            'Պայմանագրի մակարդակ դաշտի լրացում "Ապահովվող պայմանագրեր" ֆիլտրում
            Call Rekvizit_Fill("Dialog", 1, "General", "AGRLEVEL", loanAgrType(Count))
            'Հաճախորդ դաշտի լրացում "Ապահովվող պայմանագրեր" ֆիլտրում
            Call Rekvizit_Fill("Dialog", 1, "General", "CLICODE", loanAgrClient(Count))
            'Կատարել Կոճակ
            Call ClickCmdButton(2, "Î³ï³ñ»É")
            'Պայմանագրի ընտրում ցուցակից
            Count1 = Sys.Process("Asbank").vbObject("frmModalBrowser").vbObject("tdbgView").ApproxCount
            Is_Found = Search_Row(Count1, loanAgrNum(Count))
            
            If Is_Found Then
                textMsg = loanAgrNum(Count) & " is exist"
                Call Log_Print_My()
                Log.Message textMsg, "", pmNormal, attr
                
            Else
                textMsg = loanAgrNum(Count) & " isn't exist"
                Call Log_Error_My()
                Log.Error textMsg, "", pmNormal, attr
                p1.VBObject("frmModalBrowser").Close
            End If
        Next
        
    End If

    ' Արժույթ դաշտի լրացում
    Call Rekvizit_Fill("Document", 1, "General", "CURRENCY", mortCurr)
    'Սկզբնական արժեք դաշտի լրացում
    Call Rekvizit_Fill("Document", 1, "General", "SUMMA", mortSumma)
    'Ակզբնական քանակ դաշտի լրացում
    Call Rekvizit_Fill("Document", 1, "General", "COUNT", mortCountmortComment)
    'ՄԵկնաբանություն դաշտի լրացում
    Call Rekvizit_Fill("Document", 1, "General", "COMMENT", mortComment)
    'Կնքման ամսաթիվ դաշտի լրացում
    Call Rekvizit_Fill("Document", 1, "General", "DATE", date_arg)
    'Պայմանագրի համար դաշտի արժեքի վեռագռրում փոփոխականի
    Str = GetVBObject ("CODE", wMDIClient.vbObject("frmASDocForm"))
    docNumber = wMDIClient.vbObject("frmASDocForm").vbObject("TabFrame").vbObject(Str).Text
    
    'Անցում 2.Համագրավատուներ էջին
    Call GoTo_ChoosedTab(2)
    
    'Համագրավատուներ գրիդի լրացում
    For Count = 1 To UBound(partnerCode)
        'Լրացնել կոճակի սեղմում
        With wMDIClient.vbObject("frmASDocForm").vbObject("TabFrame_2").vbObject("DocGrid_2")
          .Col = 0
          .Row = Count -1
          .Keys(partnerCode(Count) & "[Tab]")
        End With
    Next
    
    'Անցում 3.Լրացուցիչ
    Call GoTo_ChoosedTab(3)
    
    'Գրավի առարկա(նոր ՎՌ) դաշտի լրացում
    Call Rekvizit_Fill("Document", 3, "General", "MORTSUBJECT", MortSubject)
    
    'Կատարել կոճակի սեղմում
    Call ClickCmdButton(1, "Î³ï³ñ»É")
End Sub

Sub Initialize_Arrays(count1, count2, count3, count4)
    ReDim loanAgrClient(count1)
    ReDim loanAgrNum(count2)
    ReDim partnerCode(count3)
    ReDim loanAgrType(count4)  
End Sub

'--------------------------------------------------------------------------------------
'¶ñ³íÇ(³ÛÉ) ï»ë³ÏÇ å³ÛÙ³Ý³·ñÇ "Üáñ ³é³ñÏ³"-ÛÇ µ³óáõÙ
'-------------------------------------------------------------------------------------
Sub Create_New_Object_Other(objCount , objSum)
    BuiltIn.Delay(3000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_NewPledge)
    BuiltIn.Delay(1000)
    
    'ø³Ý³Ï ¹³ßïÇ Éñ³óáõÙ
    Call Rekvizit_Fill("Document", 1, "General", "COUNT", objCount)
    '¶áõÙ³ñ ¹³ßïÇ Éñ³óáõÙ
    Call Rekvizit_Fill("Document", 1, "General", "SUMMA", objSum)
    'Î³ï³ñ»É
    Call ClickCmdButton(1, "Î³ï³ñ»É")
End Sub

'--------------------------------------------------------------------------------------
'¶ñ³íÇ ³éÏ³ÛáõÃÛ³Ý ëïáõ·áõÙ "êï³óí³Í ·ñ³í " ²Þî-Ç "²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ïÃÕÃ»ñ" ÃÕÃ³å³Ý³ÏáõÙ :
'üáõÝÏóÇ³Ý í»ñ³¹³ñÓÝáõÙ ¿ True, »Ã» å³ÛÙ³Ý³·ÇñÁ ³éÏ³ ¿, Ñ³Ï³é³Ï ¹»åùáõÙ` false :
'-------------------------------------------------------------------------------------
Function Search_Mortgage_In_WorkPapers(mortNumber)
    Dim isExists : isExists = False
    
    BuiltIn.Delay(3000)
    Call wTreeView.DblClickItem("|êï³óí³Í ·ñ³í|²ßË³ï³Ýù³ÛÇÝ ÷³ëï³ÃÕÃ»ñ")
    Call ClickCmdButton(2, "Î³ï³ñ»É")
    If wMDIClient.vbObject("frmPttel").vbObject("tdbgView").VisibleRows = 0 Then
        Call Log_Warning_My()
        Log.Message "There are no document with specified ID" , "" , pmNormal, attr
    Else
        wMDIClient.vbObject("frmPttel").vbObject("tdbgView").MoveFirst
        Do Until wMDIClient.vbObject("frmPttel").vbObject("tdbgView").EOF
            If Trim(wMDIClient.vbObject("frmPttel").vbObject("tdbgView").Columns.Item(2).Text) = mortNumber Then
                isExists = True
                Exit Do
            Else
                Call wMDIClient.vbObject("frmPttel").vbObject("tdbgView").MoveNext
            End If
        Loop
    End If
    
    Search_Mortgage_In_WorkPapers = isExists    
End Function

'--------------------------------------------------------------------------------------
'"Ð³×³Ëáñ¹Ç ÃÕÃ³å³Ý³Ï" ÃÕÃ³å³Ý³ÏáõÙ å³ÛÙ³Ý³·ñÇ ³éÏ³ÛáõÃÛ³Ý ëïáõ·áõÙ :
'--------------------------------------------------------------------------------------
Function Search_Morgage(mortNum)
    Dim isExists : isExists = False
    
    BuiltIn.Delay(3000)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_Folders & "|" & c_ClFolder)
    BuiltIn.Delay(2000)
    ' ¶ñ³íÇ å³ÛÙ³Ý³·ñÇ ÷ÝïñáõÙ óáõó³ÏáõÙ
    Do Until wMDIClient.vbObject("frmPttel_2").vbObject("tdbgView").EOF
        If Left((wMDIClient.vbObject("frmPttel_2").vbObject("tdbgView").Columns.Item(1).Text), 6) = mortNum Then
            isExists = True
            Exit Do
        Else
            Call wMDIClient.vbObject("frmPttel_2").vbObject("tdbgView").MoveNext
        End If
    Loop
    BuiltIn.Delay(1000)
    Search_Morgage = isExists
    BuiltIn.Delay(1000)
    wMDIClient.vbObject("frmPttel_2").Close()
End Function

'''''''''''''''''''''''''''''''''''''''
'-------------LOG---------------------'
'''''''''''''''''''''''''''''''''''''''
Sub Log_Print_My()
    Set attr = Log.CreateNewAttributes()
    attr.BackColor = BuiltIn.clMoneyGreen
    attr.FontColor = BuiltIn.clWindowText
    attr.Bold = True
    'Log.Message tEXT_MSG, "", pmNormal, attr
End Sub

Sub Log_Error_My()
    Set attr = Log.CreateNewAttributes()
    attr.BackColor = BuiltIn.clRed
    attr.FontColor = BuiltIn.clWindowText
    attr.Bold = True
    'Log.Error tEXT_MSG, "", pmNormal, attr
End Sub

Sub Log_Warning_My()
    Set attr = Log.CreateNewAttributes()
    attr.BackColor = BuiltIn.clYellow
    attr.FontColor = BuiltIn.clWindowText
    attr.Bold = True
    'Log.Warning tEXT_MSG, "", pmNormal, attr
End Sub