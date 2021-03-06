'USEUNIT Library_Common
'USEUNIT Constants
'USEUNIT Subsystems_SQL_Library


' Պայմանագրեր ֆիլտրի բացում և լրացում
' folderDirect - Թղթապանակաի ճանապարհը
' folderName - Բացվող դիալոգի անվանումը
' wLevel - Պայմանագրի մակարդակ
' eDate - Ամսաթիվ
' accBal - Հ/Պ հաշիվ 
' wAcc - Հաշիվ
' pprCode - Պայմ. թղթային N
' agrAccType - Հաշվի տիպ 
' wCur - Արժույթ
' defaultCur - Նախընտրելի արժույթ
' wClient - Հաճախորդ
' wName -  Հաճախորդի անվանում
' wNote - Նշում
' wNote2 - Նշում 2
' wNote3 - Նշում 3 
' acsBranch - Գրասենյակ
' acsDepart - Բաժին
' ascType - Հասան-ն տիպ
' clientInfo - Ցույց տալ հաճախորդի տվյալները
' showInfo - Ցույց տալ պայմանագրի պայմանները
' showOutSum - Ցույց տալ դուրս գրված  գումարները
' showNotes -  Ցույց տալ նշումները
' showAcc - Ցույց տալ հաշիվները
' wClose - Ցույց տալ փակվածները
' notFullClose - Ցույց տա ոչ լչիվ փակվածները
Sub DebitDeptContractsFilter(folderDirect, folderName, wLevel, eDate, accBal, wAcc, pprCode, agrAccType, wCur, defaultCur, wClient, wName, wNote, wNote2, _
                                                     wNote3, acsBranch, acsDepart, ascType, clientInfo, showInfo, showOutSum, showNotes, showAcc, wClose, notFullClose )

      Call wTreeView.DblClickItem(folderDirect)
      BuiltIn.Delay(1000)
      
      If Not Sys.Process("Asbank").VBObject("frmAsUstPar").Exists Then
            Log.Error(folderName & " դիալոգը չի բացվել")
            Exit Sub
      End If
      
      ' Պայմանագրի մակարդակ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "LEVEL", "^A[Del]" & wLevel )
      ' Ամսաթիվ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "EDATE",  "^A[Del]" & eDate )
      ' Հ/Պ հաշիվ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "ACCBAL", accBal )
      ' Հաշիվ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "ACC", wAcc )
      ' Պայմ. թղթային N  դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "PPRCODE", pprCode )
      ' Հաշվի տիպ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "AGRACCTYPE", agrAccType )
      ' Արժույթ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "CUR", wCur )
      ' Նախընտրելի արժույթ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "DEFAULTCUR", defaultCur )
      ' Հաճախորդ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "CLIENT", wClient )
      ' Հաճախորդի անվանում  դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "NAME", wName )
      ' Նշում դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "NOTE", wNote )
      ' Նշում 2 դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "NOTE2", wNote2 )
      ' Նշում 3  դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "NOTE3", wNote3 )
      ' Գրասենյակ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "ACSBRANCH", acsBranch )
      ' Բաժին դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "ACSDEPART", acsDepart )
      ' Հասան-ն տիպ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "ACSTYPE", ascType )
      
      ' Ցույց տալ հաճախորդի տվյալները չեքբոքսի լրացում
      Call Rekvizit_Fill("Dialog", 2, "CheckBox", "CLIENTINFO", clientInfo )
      ' Ցույց տալ պայմանագրի պայմանները չեքբոքսի լրացում
      Call Rekvizit_Fill("Dialog", 2, "CheckBox", "SHOWOTHINFO", showInfo )
      ' Ցույց տալ դուրս գրված  գումարները չեքբոքսի լրացում
      Call Rekvizit_Fill("Dialog", 2, "CheckBox", "SHOWOUTSUM", showOutSum )
      ' Ցույց տալ նշումները չեքբոքսի լրացում
      Call Rekvizit_Fill("Dialog", 2, "CheckBox", "SHOWNOTES", showNotes )
      ' Ցույց տալ հաշիվները չեքբոքսի լրացում
      Call Rekvizit_Fill("Dialog", 2, "CheckBox", "SHOWACCS", showAcc )
      ' Ցույց տալ փակվածները չեքբոքսի լրացում
      Call Rekvizit_Fill("Dialog", 2, "CheckBox", "CLOSE", wClose )
      ' Ցույց տա ոչ լրիվ փակվածները չեքբոքսի լրացում
      Call Rekvizit_Fill("Dialog", 2, "CheckBox", "NOTFULLCLOSE", notFullClose )

      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
      BuiltIn.Delay(1000)
      
End Sub

' Գործողությունների դիտում ֆիլտրի բացում և տվյալների լրացում
' folderDirect - թղթապանակի անվանումը
' stDate - Ժամանակահատվածի սկիզբ
' eDate - Ժամանակահատվածի ավարտ
' wAgr - Պայմանագրի N
' pprCode - Պայմանագրի թղթային N
' wClient - Հաճախորդ
' wName - Հաճախորդի անվանում
' acsBranch - Գրասենյակ
' acsDepart - Բաժին
' ascType - Հասան-ն տիպ
Sub OpenViewActionFilterFromDebitDept(folderDirect, stDate, eDate, wAgr, pprCode, wClient, wName, acsBranch, acsDepart, ascType)

      Call wTreeView.DblClickItem(folderDirect)
      BuiltIn.Delay(1000)
      
      If Not Sys.Process("Asbank").VBObject("frmAsUstPar").Exists Then
            Log.Error( "Ֆիլտրման դիալոգը չի բացվել")
            Exit Sub
      End If
      
      ' Ժամանակահատվածի սկիզբ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "START", "^A[Del]"  &  stDate)
       ' Ժամանակահատվածի ավարտ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "END", "^A[Del]"  & eDate )
      ' Պայմանագրի N դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "AGR", wAgr )
      ' Պայմանագրի թղթային N դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "PPRCODE", pprCode )
      ' Հաճախորդ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "CLIENT", wClient )
      ' Հաճախորդի անվանում դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "NAME", wName )
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



' Հաշիվների խմբագրում թղթապանակի բացում
' folderDirect - թղթապանակի անվանումը
' stDate - Ժամանակահատվածի սկիզբ
' eDate - Ժամանակահատվածի ավարտ
' agrNum - Պայմանագրի N
' accMaskOld - Հին հաշիվ
' accMaskNew - Նոր հաշիվ
' wUser - Կատարող
Sub EditAccFromDebitDebt(folderDirect, stDate, eDate, agrNum, accMaskOld, accMaskNew, wUser)

      Call wTreeView.DblClickItem(folderDirect)
      BuiltIn.Delay(1000)
      
      If Not Sys.Process("Asbank").VBObject("frmAsUstPar").Exists Then
            Log.Error( "Ֆիլտրման դիալոգը չի բացվել")
            Exit Sub
      End If
      
      ' Ժամանակահատվածի սկիզբ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "START", "^A[Del]"  &  stDate)
       ' Ժամանակահատվածի ավարտ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "END", "^A[Del]"  & eDate )
      ' Պայմանագրի N  դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "AGRNUM", agrNum )
      ' Հին հաշիվ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "ACCMASKOLD", accMaskOld )
      ' Նոր հաշիվ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "ACCMASKNEW",  accMaskNew)
      ' Կատարող դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "USER", wUser )
     
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
      BuiltIn.Delay(1000) 

End Sub



' Մուտք Հաշվեկշռային/Ետհաշվեկշռային ձևակերպումներ թղթապանակ
' folderDirect - թղթապանակի անվանումը
' stDate - Ժամանակահատվածի սկիզբ
' eDate - Ժամանակահատվածի ավարտ
' agrNum - Պայմանագրի N
' wCur - Արժույթ
' dealType - Գործողության տեսակ
' wUser - Կատարող
' wNote - Նշում
' wNote2 - Նշում 2
' wNote3 - Նշում 3
' acsBranch - Գրասնեյակ
' acsDepart - Բաժին
' asType - Հասան-ն տիպ
Sub BalanceSheetFormulation(folderDirect, stDate, eDate, agrNum, wCur, dealType, wUser, wNote, wNote2, wNote3, acsBranch, acsDepart, asType)

      Call wTreeView.DblClickItem(folderDirect)
      BuiltIn.Delay(1000)
      
      If Not Sys.Process("Asbank").VBObject("frmAsUstPar").Exists Then
            Log.Error( "Ֆիլտրման դիալոգը չի բացվել")
            Exit Sub
      End If
      
      ' Ժամանակահատվածի սկիզբ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "START", "^A[Del]"  &  stDate)
       ' Ժամանակահատվածի ավարտ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "END", "^A[Del]"  & eDate )
      ' Պայմանագրի N  դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "AGR", agrNum )
      ' Արժույթ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "CUR", wCur )
      ' Գործողության տեսակ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "DEALTYPE", dealType )
      ' Կատարող դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "USER", wUser )
      ' Նշում դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "NOTE", wNote )
      ' Նշում 2 դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "NOTE2", wNote2 )
      ' Նշում 3 դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "NOTE3", wNote3 )
      ' Գրասնեյակ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "ACSBRANCH", acsBranch )
      ' Բաժին դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "ACSDEPART", acsDepart )
      ' Հասան-ն տիպ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "ACSTYPE", asType )
     
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
      BuiltIn.Delay(1000) 

End Sub


' Մուտք "Պայմանագրերի դաշտերի փոփոխման հայտեր" թղթապանակ
' folderDirect - թղթապանակի անվանումը
' dState - Վիճակ
' stDate - Ժամանակահատվածի սկիզբ
' eDate - Ժամանակահատվածի ավարտ
' dUser - Կատարող
' dAcsBranch - Գրասենյակ
' dAcsDepart - Բաժին
Sub OpenChangeRequestContractFieldsDoc(folderDirect, dState, stDate, eDate, dUser, dAcsBranch, dAcsDepart)

      Call wTreeView.DblClickItem(folderDirect)
      BuiltIn.Delay(1000)
      
      If Not Sys.Process("Asbank").VBObject("frmAsUstPar").Exists Then
            Log.Error( "Ֆիլտրման դիալոգը չի բացվել")
            Exit Sub
      End If
      
      ' Վիճակ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "STATE", "^A[Del]"  & dState )
      ' Ժամանակահատվածի սկիզբ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "SDATE", "^A[Del]"  &  stDate)
       ' Ժամանակահատվածի ավարտ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "EDATE", "^A[Del]"  & eDate )
      ' Կատարող դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "USERS", dUser )
      ' Գրասենյակ դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "ACSBRANCH", dAcsBranch )
      ' Բաժին դաշտի լրացում
      Call Rekvizit_Fill("Dialog", 1, "General", "ACSDEPART", dAcsDepart )
      
      ' Կատարել կոճակի սեղմում
      Call ClickCmdButton(2, "Î³ï³ñ»É")
      BuiltIn.Delay(1000) 
      
End Sub



' Ֆիլտրեր գործողության կատարում, Ֆիլտրի ստեղծում
' filterName - ֆիլտրի անվանում
Sub  CreateFilterForSort(filterName)

        BuiltIn.Delay(2000)
        ' Բացել Ֆիլտրել պատուհանը
       Call wMainForm.MainMenu.Click(c_Opers)
       Call wMainForm.PopupMenu.Click( c_Folder & "|" & c_Filter)
           
       ' Ստուգել Ֆիլտրել պատուհանը բացվել է թե ոչ
       If  Sys.Process("Asbank").WaitVBObject("frmPttelFilter", 3000).Exists Then
           
            ' Սեղմել "Հիշել որպես" կոճակը
            Sys.Process("Asbank").VBObject("frmPttelFilter").VBObject("Command11").Click
            If  Sys.Process("Asbank").VBObject("frmPttelFilterSaveAs").Exists Then
                
                  ' Լրացնել Ֆիլտրի անվանում դաշտը
                  Sys.Process("Asbank").VBObject("frmPttelFilterSaveAs").VBObject("Frame1").VBObject("Combo1").Window("Edit", "", 1).Keys(filterName)
                  ' Սեղմել "Հիշել" կոճակը
                  Sys.Process("Asbank").VBObject("frmPttelFilterSaveAs").VBObject("Command12").Click
                  If p1.WaitVBObject("frmAsMsgBox",2000).Exists Then
                    If  MessageExists(2, "ÀÝÃ³óÇÏ üÇÉïñÁ ÏÑÇßíÇ §"& filterName &"¦ ³Ýí³Ý ï³Ï") Then
                        Call ClickCmdButton(5, "²Ûá") 
                    End If 
                  End If
                  ' Սեղմել "Կատարել" կոճակը
                  Sys.Process("Asbank").VBObject("frmPttelFilter").VBObject("Command5").Click
                
            Else  
                  Log.Error("Հիշել որպես պատուհանը չի բացվել")
            End If
                
       Else
            Log.Error("Ֆիլտրել պատուհանը չի բացվել")
       End if
           
End Sub