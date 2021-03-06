'USEUNIT Library_Common
'USEUNIT Constants
'USEUNIT Subsystems_SQL_Library


' Աշխատանքային փաստաթղթեր Ֆիլտրի լրացում
' folderDirect -  թղթապանակի անվանումը
' stDate - Ժամանակահատվածի սկիզբ
' eDate - Ժամանակահատվածի ավարտ
' wCur - Արժույթ
' wUser - Կատարողներ
' docType - Փաստաթղթի տեսակ
' paySysin - Ընդ. վճ. համակարգ
' paySysOut - Ուղ. վճ. համակարգ
' payNotes - Նշում
' acsBranch - Գրասենյակ
' acsDepart - Բաժին
' selectView - Դիտելու ձև
' exportExcel - Լրացնել
Sub WorkingDocsFilter(folderDirect, stDate, eDate, wCur, wUser, docType, paySysin, paySysOut, payNotes, acsBranch, acsDepart, selectView, exportExcel)

      Call wTreeView.DblClickItem(folderDirect)
      BuiltIn.Delay(1000)
      
      If Not Sys.Process("Asbank").VBObject("frmAsUstPar").Exists Then
            Log.Error( "Ֆիլտրման դիալոգը չի բացվել")
            Exit Sub
      End If
      
      ' Ժամանակահատվածի սկիզբ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "PERN", stDate )
      ' Ժամանակահատվածի ավարտ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "PERK", eDate )
      ' Արժույթ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "CUR",  wCur)
      ' Կատարողներ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "USER",   "^A[Del]"  & wUser)
      ' Փաստաթղթի տեսակ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "DOCTYPE", docType )
      ' Ընդ. վճ. համակարգ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "PAYSYSIN", paySysin )
      ' Ուղ. վճ. համակարգ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "PAYSYSOUT", paySysOut )
      ' Նշում դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "PAYNOTES", payNotes )
      ' Գրասենյակ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "ACSBRANCH", acsBranch )
      ' Բաժին դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "ACSDEPART", acsDepart )
      ' Դիտելու ձև դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "SELECTED_VIEW", "^A[Del]" & selectView )
      ' Լրացնել դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "EXPORT_EXCEL",  "^A[Del]" & exportExcel )

      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
      BuiltIn.Delay(1000)

End Sub

' նոր ստեղծված քարտ փաստաթղթեր ֆիլտրի լրացում
' folderDirect - թղթապանակի անվանումը
' docType - Փասատթղթի տեսակ 
' acsBranch - Գրասենյակ
' acsDepart - Բաժին
' passTax - Անձնագիր/ՀվՀՀ
' wName -Անվանում
' wUser - Կատարող
Sub NewCreatedCardDoc(folderDirect, docType, acsBranch, acsDepart, passTax, wName, wUser)

      Call wTreeView.DblClickItem(folderDirect)
      BuiltIn.Delay(1000)
      
      If Not Sys.Process("Asbank").VBObject("frmAsUstPar").Exists Then
            Log.Error("Ֆիլտրման դիալոգը չի բացվել")
            Exit Sub
      End If
      
      ' Փասատթղթի տեսակ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "DOCTYPE", docType )
      ' Գրասենյակ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "ACSBRANCH",  acsBranch)
      ' Բաժին դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "ACSDEPART", acsDepart )
      ' Անձնագիր/ՀվՀՀ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "PASSTAX", passTax )
      ' Անվանում դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "NAME", wName )
      ' Կատարող դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "USER", wUser )
      
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
      BuiltIn.Delay(1000)
      
End Sub


' "Քարտ փաստաթղթերի փոփոխման պատմություն" դիալոգի բացում և տվյալների լրացում
' folderDirect - թղթապանակի անվանումը
' docType - Փասատթղթի տեսակ 
' wState - Վիճակ
' sDate - ժամանակահատվածի սկիզբ
' eDate - ժամանակահատվածի ավարտ
' wUser - Օգտագործող
Sub HistoryOfChangeRequest(folderDirect, docType, wState, stDate, eDate, wUser)

      Call wTreeView.DblClickItem(folderDirect)
      BuiltIn.Delay(1000)
      
      If Not Sys.Process("Asbank").VBObject("frmAsUstPar").Exists Then
            Log.Error("Ֆիլտրման դիալոգը չի բացվել")
            Exit Sub
      End If

      ' Փասատթղթի տեսակ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "DOCTYPE", docType )
      ' Վիճակ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "STATE", "^A[Del]"  & wState )
      ' ժամանակահատվածի սկիզբ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "SDATE", "^A[Del]"  & stDate )
      ' Ժամանակահատվածի ավարտ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "EDATE", "^A[Del]"  & eDate )
      ' Օգտագործող դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "USER", wUser )
      
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
      BuiltIn.Delay(1000)
      
End Sub

' Պայմանագրեր ֆիլտրի բացում և տվյալների լրացում
' folderDirect -  թղթապանակի անվանումը
' balAcc - Հ/Պ հաշվեկշռային հաշիվ
' accMask - Հաշվի շաբլոն
' wCur - Արժույթ
' accType - Հաշիվների տիպեր
' acsBranch - Գրասնեյակ
' acsDepart - Բաժին
' acsType - Հասան-ն տիպ
' accNote - Նշում
' accNote2 - Նշում 2
' accNote3 - Նշում 3
' wScale - Սանդղակ
' showClosed - Ցույց տալ փակված պայմանագրերը
' oldAccMask - Հին հաշիվ
' newAccMask - Նոր հաշիվ
' showBal - Ցույց Տալ Կկուտ. և թղթ. Հ/Հ-ները
' showRem - Ցույց տալ հաշիվների մնացորդը
' wDate - Ամսաթիվ (_/_/_- ընթացիկ)
' selectView - Դիտելու ձև
' exportExcel - Լրացնել
Sub OpenContractsFolder(folderDirect, balAcc, accMask, wCur, accType, acsBranch, acsDepart, acsType, accNote, accNote2, _ 
                                                      accNote3, wScale, showClosed, oldAccMask, newAccMask, showBal, showRem, wDate, selectView, exportExcel)

      Call wTreeView.DblClickItem(folderDirect)
      BuiltIn.Delay(1000)
      
      If Not Sys.Process("Asbank").VBObject("frmAsUstPar").Exists Then
            Log.Error( "Ֆիլտրման դիալոգը չի բացվել")
            Exit Sub
      End If

      ' Հ/Պ հաշվեկշռային հաշիվ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "BALACC", balAcc )
      ' Հաշվի շաբլոն դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "ACCMASK", accMask )
      ' Արժույթ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "CUR", wCur )
      ' Հաշիվների տիպեր դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "ACCTYPE", accType )
      ' Գրասնեյակ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "ACSBRANCH", acsBranch )
      ' Բաժին աշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "ACSDEPART", acsDepart )
      ' Հասան-ն տիպ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "ACSTYPE", acsType )
      ' Նշում դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "ACCNOTE", accNote )
      ' Նշում 2  դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "ACCNOTE2", accNote2 )
      ' Նշում 3 դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "ACCNOTE3", accNote3 )
      ' Սանդղակ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "SCALE", wScale )
      ' Ցույց տալ փակված պայմանագրերը դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "CheckBox", "SHOWCLOSED", showClosed )
      ' Հին հաշիվ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "OLDACCMASK", oldAccMask )
      ' Նոր հաշիվ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "NEWACCMASK", newAccMask )
      ' Ցույց Տալ Կկուտ. և թղթ. Հ/Հ-ները դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "CheckBox", "SHOWBAL", showBal )
      ' Ցույց տալ հաշիվների մնացորդը դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "CheckBox", "SHOWREM", showRem )
      ' Ամսաթիվ (_/_/_- ընթացիկ) դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "DATE", wDate )
      ' Դիտելու ձև դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "SELECTED_VIEW", "^A[Del]"  &  selectView )
      ' Լրացնել դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "EXPORT_EXCEL", "^A[Del]"  & exportExcel )
      
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
      BuiltIn.Delay(1000)
      
End Sub


' Մուտք հաշվառված փասատթղթեր թղթապանակ
' folderDirect - թղթապանակի անվանումը
' stDate - Ժամանակահատվածի սկիզբ
' eDate - Ժամանակահատվածի ավարտ
' wUser - Կատարող
Sub RegisteredDocuments(folderDirect, stDate, eDate, wUser)
      
      Call wTreeView.DblClickItem(folderDirect)
      BuiltIn.Delay(1000)
      
      If Not Sys.Process("Asbank").VBObject("frmAsUstPar").Exists Then
            Log.Error("Ֆիլտրման դիալոգը չի բացվել")
            Exit Sub
      End If
      
      ' Ժամանակահատվածի սկիզբ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "PERN", "^A[Del]"  &  stDate)
       ' Ժամանակահատվածի ավարտ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "PERK", "^A[Del]"  & eDate )
       ' Կատարող դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "USER", wUser )
      
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
      BuiltIn.Delay(1000)

End Sub


' Ստեղծված խմբային խմբագրումներ ֆիլտրի բացում և տվյալների լրացում
' folderDirect - թղթապանակի անվանումը
' wState - Վիճակ
' stDate - Ժամանակահատվածի սկիզբ
' eDate - Ժամանակահատվածի ավարտ
' docType - Փաստաթղթի տեսակ 
' wUser - Կատարող
Sub CreatedGroupEdits(folderDirect, wState, stDate, eDate, docType, wUser, selectedView, expExcel)

      Call wTreeView.DblClickItem(folderDirect)
      BuiltIn.Delay(1000)
      
      If Not Sys.Process("Asbank").VBObject("frmAsUstPar").Exists Then
            Log.Error( "Ֆիլտրման դիալոգը չի բացվել")
            Exit Sub
      End If
      
      ' Վիճակ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "STATE",  wState)
      ' Ժամանակահատվածի սկիզբ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "SDATE", "^A[Del]"  &  stDate)
       ' Ժամանակահատվածի ավարտ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "EDATE", "^A[Del]"  & eDate )
      ' Փաստաթղթի տեսակ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "DOCTYPE",  docType )
       ' Կատարող դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "USER", wUser )
      ' Դիտելու ձև դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "SELECTED_VIEW", "^A[Del]" & selectedView )
      ' Լրացնել դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "EXPORT_EXCEL", "^A[Del]" & expExcel )
      
      
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
      BuiltIn.Delay(1000)
End Sub



' Ջնջված փաստաթղթեր ֆիլտրի բացում և տվյալների լրացում
' folderDirect - թղթապանակի անվանումը
' stDate - Ժամանակահատվածի սկիզբ
' eDate - Ժամանակահատվածի ավարտ
' docsP - Փաստաթղթի տեսակ
' wISN - Փաստաթղթի ISN 
' accRow - Հաշվեկշռ. գործ. ունեցող
' wUser - Կատարող
' acsBranch - Գրասենյակ
' acsDepart - Բաժին
Sub DeletedDocFilter(folderDirect, stDate, eDate, docsP, wISN, accRow, wUser, acsBranch, acsDepart)

      Call wTreeView.DblClickItem(folderDirect)
      BuiltIn.Delay(1000)
      
      If Not Sys.Process("Asbank").VBObject("frmAsUstPar").Exists Then
            Log.Error( "Ֆիլտրման դիալոգը չի բացվել")
            Exit Sub
      End If

      ' Ժամանակահատվածի սկիզբ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "SDATE", "^A[Del]"  &  stDate)
       ' Ժամանակահատվածի ավարտ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "EDATE", "^A[Del]"  & eDate )
      ' Փաստաթղթի տեսակ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "DOCTP", docsP )
      ' Փաստաթղթի ISN դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "ISN", " ^A[Del]"  & wISN )
      ' Հաշվեկշռ. գործ. ունեցող դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "CheckBox", "HAD01ACCROW",  accRow)
      ' Կատարող դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "USER", "^A[Del]"  & wUser )
      ' Գրասենյակ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "ACSBRANCH", acsBranch )
      ' Բաժին դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "ACSDEPART", acsDepart )
      
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
      BuiltIn.Delay(1000)
      
End Sub


' Գործողությունների դիտում ֆիլտրի բացում և տվյալների լրացում
' folderDirect - թղթապանակի անվանումը
' stDate - Ժամանակահատվածի սկիզբ
' eDate - Ժամանակահատվածի ավարտ
' balAcc - Հ/Պ հաշվեկշռային հաշիվ
' accMask - Հաշվի շաբլոն
' wCur - Արժույթ
' acsBranch - Գրասենյակ
' acsDepart - Բաժին
' ascType - Հասան-ն տիպ
Sub ViewActionFilter(folderDirect, stDate, eDate, balAcc, accMask, wCur, acsBranch, acsDepart, ascType)

      Call wTreeView.DblClickItem(folderDirect)
      BuiltIn.Delay(1000)
      
      If Not Sys.Process("Asbank").VBObject("frmAsUstPar").Exists Then
            Log.Error( "Ֆիլտրման դիալոգը չի բացվել")
            Exit Sub
      End If
      
      ' Ժամանակահատվածի սկիզբ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "SDATE", "^A[Del]"  &  stDate)
       ' Ժամանակահատվածի ավարտ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "EDATE", "^A[Del]"  & eDate )
      ' Հ/Պ հաշվեկշռային հաշիվ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "BALACC", balAcc )
      ' Հաշվի շաբլոն դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "ACCMASK", accMask )
      ' Արժույթ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "CUR", wCur )
      ' Գրասենյակ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "ACSBRANCH", acsBranch )
      ' Բաժին դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "ACSDEPART", acsDepart )
      ' Հասան-ն տիպ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "ACSTYPE", ascType )
      
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
      BuiltIn.Delay(1000)
      
End Sub


' Տոկոսների հաշվարկման ԱՇՏ-ում "Պայմանագրերի սանդղակների տեղեկատու" ֆիլտրի ստուգում:
' folderDirect  - թղթապանակի անվանումը
' showAccs - Ցույց տալ կուտ. և թղթ. հաշիվները
' showInt - Ցույց տալ կապիտալացումները
' showPer - Ցույց տալ հաշվետվության մասը
Sub ContractScaleInformation(folderDirect, showAccs, showInt, showPer)

      Call wTreeView.DblClickItem(folderDirect)
      BuiltIn.Delay(1000)
      
      If Not Sys.Process("Asbank").VBObject("frmAsUstPar").Exists Then
            Log.Error("Ֆիլտրման դիալոգը չի բացվել")
            Exit Sub
      End If

      ' Ցույց տալ կուտ. և թղթ. հաշիվները չեքբոքսի լրացում
      Call Rekvizit_Fill("Dialog", 1, "CheckBox", "SHOWACCS", showAccs)
      ' Ցույց տալ կապիտալացումները չեքբոքսի լրացում
      Call Rekvizit_Fill("Dialog", 1, "CheckBox", "SHOWINT", showInt)
      ' Ցույց տալ հաշվետվության մասը չեքբոքսի լրացում
      Call Rekvizit_Fill("Dialog", 1, "CheckBox", "SHOWREP", showPer)
      
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
      BuiltIn.Delay(1000)
      
End Sub