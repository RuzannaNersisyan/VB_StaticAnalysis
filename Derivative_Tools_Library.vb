'USEUNIT Library_Common
'USEUNIT Constants

'----------------------------------------------------------------------------------------------------------------------------------
' Պայմանագրի "Ձևակերպում" գործողության կատարում 
'----------------------------------------------------------------------------------------------------------------------------------
'date - Ամսաթիվ 
Sub Entry(date)
  
    BuiltIn.Delay(2000)
    Call Sys.Process("Asbank").VBObject("MainForm").MainMenu.Click(c_AllActions)
    Call Sys.Process("Asbank").VBObject("MainForm").PopupMenu.Click(c_Opers & "|" &  c_EntAndRep & "|" & c_Entry)
    'Ամսաթիվ դաշտի լրացում
    Call  Rekvizit_Fill("Document",1,"General","DATE", date)
    'Սեղմել Կտարել կոճակը
    Call ClickCmdButton(1,"Î³ï³ñ»É")
    'Սեղմել Այո կոճակը
    Call ClickCmdButton(5,"²Ûá")
    
End Sub

'-------------------------------------------------------------------------------------------------------------------------------------
' " Ժամկետների վերանայում" գործողության կատարում 
'-------------------------------------------------------------------------------------------------------------------------------------
'date - Անսաթիվ
'repDate - Մարման ամսաթիվ
'extention - Երկարաձգում
Sub ReviewTerms(date,repDate,extention)
  
    Call Sys.Process("Asbank").VBObject("MainForm").MainMenu.Click(c_AllActions)
    Call Sys.Process("Asbank").VBObject("MainForm").PopupMenu.Click(c_TermsStates & "|" & c_Dates & "|" & c_ReviewTerms)
    'Լրացնել Ամսաթիվ դաշտը
    Call Rekvizit_Fill("Document",1,"General","DATE",date)
    'Լրացնել Մարման ամսաթիվ դաշտը
    Call Rekvizit_Fill("Document",1,"General","DATEAGR",repDate)
    'Լրացնել Երկարաձգում դաշտը
    Call Rekvizit_Fill("Document",1,"General","ISPROLONG",extention)
    'Սեղմել Կատարել կոճակը
    Call ClickCmdButton(1,"Î³ï³ñ»É")
  
End Sub


'-----------------------------------------------------------------------------------------------------------------
' " Տոկոսադորւյք" գործողության կատարում
'-----------------------------------------------------------------------------------------------------------------
'date - Անսաթիվ
'per - Ստացվելիք տոկոսադրոյք
'part -  բաժ`
'sPer - Վճարվելիք տոկոսադրույք
'sPart - բաժ`
Sub Set_Persentage(fBase,date,per,part, sPer, sPart)

    Call Sys.Process("Asbank").VBObject("MainForm").MainMenu.Click(c_AllActions)
    Call Sys.Process("Asbank").VBObject("MainForm").PopupMenu.Click(c_TermsStates & "|" & c_Percentages & "|" & c_Percentages)
   
    fBase = wMDIClient.vbObject("frmASDocForm").DocFormCommon.Doc.isn
    'Լրացնել Ամսաթիվ դաշտը
    Call Rekvizit_Fill("Document",1,"General","DATE",date)
    'Լրացնել Ստացվելիք տոկոսադրոյք դաշտը
    Call Rekvizit_Fill("Document",1,"General","PCAGR",per)
    'Լրացնել բաժ` դաշտը
    Call Rekvizit_Fill("Document",1,"General","PCAGR",part)
    'Լրացնել Վճարվելիք տոկոսադրույք դաշտը
    Call Rekvizit_Fill("Document",1,"General","PCAGRCR",sPer)
    'Լրացնել բաժ` դաշտը
    Call Rekvizit_Fill("Document",1,"General","PCAGRCR",sPart)
    'Սեղմել Կատարել կոճակը
    Call ClickCmdButton(1,"Î³ï³ñ»É")
  
End Sub

'-------------------------------------------------------------------------------------------------------
' " Մարում " գործողության կատարում 
'-------------------------------------------------------------------------------------------------------
'date - Ամսաթիվ
'repType - Մարման տեսակ
'time - Ժամանակ
'sold - Առք/Վաճառք
'place - Գործողության վայր
Sub Repayments(date, repType, time, sold, place,docType,contr)
  
    If Not contr = "îáÏáë³¹ñáõÛù³ÛÇÝ ëíá÷" Then
      Call Sys.Process("Asbank").VBObject("MainForm").MainMenu.Click(c_AllActions)
      Call Sys.Process("Asbank").VBObject("MainForm").PopupMenu.Click(c_Opers & "|" &  c_EntAndRep & "|" & c_Repayment)
    Else 
      Call Sys.Process("Asbank").VBObject("MainForm").MainMenu.Click(c_AllActions)
      Call Sys.Process("Asbank").VBObject("MainForm").PopupMenu.Click(c_Opers & "|" &  c_Repayment & "|" & c_Repayment)
    End If
    'Լրացնել Ամսաթիվ դաշտը
    Call Rekvizit_Fill("Document",1,"General","DATE",date)
    If Not  docType = 1 then
      'Լրացնել Մարման տեսակ դաշտը
      Call Rekvizit_Fill("Document",1,"General","DEBTTYPE",repType)
    End If
    'Լրացնել Ժամանակ դաշտը
    Call Rekvizit_Fill("Document",1,"General","TIME",time)
    'Լրացնել Առք/Վաճառք դաշտը
    Call Rekvizit_Fill("Document",1,"General","CUPUSA",sold)
    'Լրացնել Գործողության վայր դաշտը
    Call Rekvizit_Fill("Document",1,"General","CURVAIR",place)
    'Սեղմել Կատարել կոճակը
    Call ClickCmdButton(1,"Î³ï³ñ»É")
    'Սեղմել Այե կոճակը
    Call ClickCmdButton(5,"²Ûá")
  
End Sub


'----------------------------------------------------------------------------
'Հաշվարկների ճշգրտում
'------------------------------------------------------------------------------------------
'dateStart - Ամսաթիվ
'summperc - Գումար
Sub Correc_Calculation(dateStart, summperc, fBase)
    
    Dim Str
    Dim wMainForm, wTabStrip
    Set wMainForm = Sys.Process("Asbank").VBObject("MainForm") 
    Set wMDIClient = wMainForm.Window("MDIClient", "", 1)
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_Opers & "|" & c_Interests & "|" & c_AccAdjust)
    fBase = Sys.Process("Asbank").vbObject("MainForm").Window("MDIClient", "", 1).vbObject("frmASDocForm").DocFormCommon.Doc.isn
    'Ամսաթիվ դաշտի լրացում
    Call Rekvizit_Fill("Document",1,"General","DATE", "![End]" & "[Del]" & dateStart)
  
    'Տոկոսադրույք դաշտի լրացում
    Call Rekvizit_Fill("Document",1,"General","SUMPERDB","![End]" & "[Del]" & summperc)
  
    'Կատարել կոճակի սեղմում
    Sys.Process("Asbank").vbObject("MainForm").Window("MDIClient", "", 1).vbObject("frmASDocForm").vbObject("CmdOk_2").Click()
 
End Sub


'------------------------------------------------------------------------------------------------------
'Ռեկվիզիտի խմբային խմբագրում(Վարկային ռեգիստր)
'------------------------------------------------------------------------------------------------------
'registr - Հաշվառել վարկային ռեգիստրում նշիչը
'info - Լրացուցիչ ինֆորմացիա նշիչը
'addInfo - լրացուցիչ ինֆորմացիա դաշտը
'review - Պայմանագրի վերանայման պատճառ նշիչը
'reason - Պայմանագրի վերանայման պատճառ դաշտը
'rep - Մարման աղբյուր նշիչը
'repSource - Մարման աղբյուր դաշտը
'mortage - Գրավի առարկա (նոր ՎՌ) նշիչը
'morSub  - Գրավի առարկա (նոր ՎՌ) դաշտը
'insure - Ապահովված է այլ ապահովվությամբ նշիչը
'otherInsure - Ապահովված է այլ ապահովվությամբ դաշտը
'chageClosed - Փոխել փակվաշները նշիչը
Sub Rekvizit_Group_Fill(registr,info,addInfo,review,reason,rep,repSource,_
                          mortage,morSub,insure,otherInsure,chageClosed)
    
    BuiltIn.Delay(1000)
    Call wMainForm.MainMenu.Click(c_ToRefresh)  
    Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").Keys("[Ins]")
    BuiltIn.Delay(2000) 
    Call wMainForm.MainMenu.Click(c_AllActions)    
    Call wMainForm.PopupMenu.Click(c_RekvGroupFill & "|" & c_CreditReg )
  
    If info = 1 Then
      Call Rekvizit_Fill("Dialog",1,"CheckBox","FILLUPPUTINLR",credReg)
      Call Rekvizit_Fill("Dialog",1,"CheckBox","PUTINLR",regisrt)
    End If
    If info = 1 Then
      Call Rekvizit_Fill("Dialog",1,"CheckBox","FILLUPOTHER",info)
      Call Rekvizit_Fill("Dialog",1,"General","OTHER",addInfo)
    End If
    If review = 1 Then
      Call Rekvizit_Fill("Dialog",1,"CheckBox","FILLUPREVISIONREASON",review)
      Call Rekvizit_Fill("Dialog",1,"General","REVISIONREASON",reason)
    End If
    If rep = 1 Then
      Call Rekvizit_Fill("Dialog",1,"CheckBox","FILLUPREPSOURCE",rep)
      Call Rekvizit_Fill("Dialog",1,"General","REPSOURCE",repSource)
    End If
    If mortage = 1 Then
      Call Rekvizit_Fill("Dialog",1,"CheckBox","FILLUPMORTSUBJECT",mortage)
      Call Rekvizit_Fill("Dialog",1,"General","MORTSUBJECT",morSub)
    End If
    If insure = 1 Then
      Call Rekvizit_Fill("Dialog",1,"CheckBox","FILLUPOTHERCOLLATERAL",insure)
      Call Rekvizit_Fill("Dialog",1,"CheckBox","OTHERCOLLATERAL",otherInsure)
    End If
  
    Call Rekvizit_Fill("Dialog",1,"CheckBox","CLOSED",chageClosed)
    'Սեղմել Կատարել կոճակը
    Call ClickCmdButton(2,"Î³ï³ñ»É")  
    'Սեղմել " - "
    Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").VBObject("tdbgView").Keys("NumMinus")   

End Sub

'--------------------------------------------------------------------------------------------
'"Պայմանագրի փակում" գործողության կատարում 
'--------------------------------------------------------------------------------------------
'CloseDate - փակման ամսաթիվ
Sub Close_Contract(CloseDate)

    BuiltIn.Delay(4000) 
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_AgrClose)
    'Լրացնում է փակման ամսաթիվ դաշտը
    Call Rekvizit_Fill("Dialog", 1, "General", "DATECLOSE", CloseDate)
    'Սեղմել Կատարել կոճակը
    BuiltIn.Delay(3000) 
    Call ClickCmdButton(2,"Î³ï³ñ»É")  
    
End Sub

'--------------------------------------------------------------------------------------
'Օվերդրաֆտի համար Խմբային "Տոկոսների հաշվարկում" գործողության կատարում :
'--------------------------------------------------------------------------------------
'calcDate - Հաշվարկման ամսաթիվ դատի արժեք
'givenDate - Ջևակերօման ամսաթիվ դաշտի արժեք
Sub Group_Percent_Calculate_Overdraft(calcDate , givenDate)
    wMDIClient.vbObject("frmPttel").vbObject("tdbgView").Keys("[Ins]")
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click(c_GroupCalc)
    'Հաշվարկման ամսաթիվ դաշտի լրացում
    Call Rekvizit_Fill("Dialog", 1, "General", "CloseDate", calcDate) 
    'Հատկացման ամսաթիվ դաշտի լրացում
    Call Rekvizit_Fill("Dialog", 1, "General", "SetDate", givenDate) 
'    Sys.Process("Asbank").vbObject("frmAsUstPar").vbObject("TabFrame").vbObject("Checkbox_3").Click()
    '"Տոկոսների հաշվարկում"  նշիչի նշում
    Call Rekvizit_Fill("Dialog", 1, "CheckBox", "CHG", 1) 
    'Կատարել կոճակի սեղմում
    Sys.Process("Asbank").vbObject("frmAsUstPar").vbObject("CmdOK").Click()
    Sys.Process("Asbank").vbObject("frmAsMsgBox").vbObject("cmdButton").Click()   
End Sub