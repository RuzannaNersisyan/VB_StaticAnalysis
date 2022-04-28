Option Explicit
'USEUNIT Library_Common
'USEUNIT Constants
'USEUNIT Library_Contracts
'USEUNIT Library_Colour

'-----------------------------------------------------------------------------------------
'����� � ������� �߳������� ������û� ��ó�ݳ� г׳���� �������� � ��ٳ��� ���-��
'-----------------------------------------------------------------------------------------
'startDate - ��˳������� ������û� ������� �Ͻ��ݳϳ� ������
'endDate - ��˳������� ������û� ������� ��絳ϳ� ������
Sub Online_PaySys_Go_To_Agr_WorkPapers(workSpace, startDate, endDate)
  BuiltIn.Delay(3000)
  Call wTreeView.DblClickItem(workSpace)
  if p1.WaitVBObject("frmAsUstPar", 1000).Exists then
    Call Rekvizit_Fill("Dialog", 1, "General", "PERN", startDate)
    Call Rekvizit_Fill("Dialog", 1, "General", "PERK", endDate)
    Call ClickCmdButton(2, "γ���")
  else 
    Log.Error "Can't open frmAsUstPar window", "", pmNormal, ErrorColor
  end if
End Sub

'-----------------------------------------------------------------------------------------
'г����� �׳���� ������û���� �������� ��ϳ����۳� ��������
'-----------------------------------------------------------------------------------------
'docNum - �������� N
'startDate - �������� ������ ������(��ǽ�)
'endDate - �������� ������ ������(���)
Function Online_PaySys_Check_Doc_In_Registered_Payment_Documents(docNum, startDate, endDate)
    Dim exists, my_vbobj,Count
    Dim wMainForm,wTreeView
    
    Set wMainForm = Sys.Process("Asbank").VBObject("MainForm") 
    Set wMDIClient = wMainForm.Window("MDIClient", "", 1) 
    Set wTreeView = wMDIClient.VBObject("frmExplorer").VBObject("tvTreeView")
    Call wTreeView.DblClickItem("|г׳���� �������� � ��ٳ��� |г����� �׳���� ������û�")
    Sys.Process("Asbank").vbObject("frmAsUstPar").vbObject("TabFrame").vbObject("TDBDate").Keys(startDate & "[Tab]")
    Sys.Process("Asbank").vbObject("frmAsUstPar").vbObject("TabFrame").vbObject("TDBDate_2").Keys(endDate & "[Tab]")
    Sys.Process("Asbank").vbObject("frmAsUstPar").vbObject("TabFrame").vbObject("ASTypeTree").vbObject("TDBMask").Keys("[End]" & "[BS]" & "[BS]" & "[Tab]")
    Sys.Process("Asbank").vbObject("frmAsUstPar").vbObject("CmdOK").Click()
    Count = 0
    exists = False 
    BuiltIn.Delay(5000)
    Set my_vbobj = wMDIClient.WaitVBObject("frmPttel", delay_middle)
    If my_vbobj.Exists Then
        Do Until Sys.Process("Asbank").vbObject("MainForm").Window("MDIClient", "", 1).vbObject("frmPttel").vbObject("tdbgView").ApproxCount < Count  Or exists = True
            If Trim(wMDIClient.vbObject("frmPttel").vbObject("tdbgView").Columns.Item(2).text) = docNum Then
                exists = True
            Else
            Count = Count + 1 
                Call wMDIClient.vbObject("frmPttel").vbObject("tdbgView").MoveNext
            End If
        Loop
    Else
        Log.Error("Rgistered payment view doesn't exists")
    End If
    
    Online_PaySys_Check_Doc_In_Registered_Payment_Documents = exists
End Function

'-----------------------------------------------------------------------------------------
'�������� �������� ѳ���ٳ�(��ٳ��� ϳ� г����� 1 ��� ϳ� � �����)
'-----------------------------------------------------------------------------------------
'sendTo - 1��Ż�� ������� ���������� ��ճ������ � ѳ���ٳ� ��ٳ���(�׳�ٳ� ѳ��ݳ���� (Online ��� ���.)), 2 - � �������` г������ 1���(�׳�ٳ� ѳ��ݳ���� (Online ��� ��.)),3 -� ������� � �����
Sub Online_PaySys_Send_To_Verify(sendTo)
  BuiltIn.Delay(3000)
  Call wMainForm.MainMenu.Click(c_AllActions)
  Select Case sendTo
    Case 1
      Call wMainForm.PopupMenu.Click(c_SendToCash)
      Call ClickCmdButton(2, "γ���")
    Case 2
      Call wMainForm.PopupMenu.Click(c_SendToVer)
      Call ClickCmdButton(2, "γ���")
    Case 3
      Call wMainForm.PopupMenu.Click(c_SendtoVerBL)
      Call ClickCmdButton(5, "���")
  End Select
  if p1.WaitVBObject("frmAsMsgBox", 8000).Exists Then
    Call ClickCmdButton(5, "OK")
  End if
  BuiltIn.Delay(1000)
  wMDIClient.vbObject("frmPttel").Close()
End Sub

'------------------------------------------------------------------------------
'�������� ѻ����� г���� �׳���� ������û� ��ó�ݳ���
'-------------------------------------------------------------------------------
Sub Online_PaySys_Delete_Agr()
  BuiltIn.Delay(3000)
  Call wMainForm.MainMenu.Click(c_AllActions)
  Call wMainForm.PopupMenu.Click(c_Delete)
  If p1.WaitVBObject("frmDeleteDoc", 3000).Exists Then 
    Call ClickCmdButton(3, "���")
  Else 
    Log.Error "Can't find frmDeleteDoc window", "", pmNormal, ErrorColor
  End If
End Sub

'------------------------------------------------------------------------------
'�������� ��ϳ����۳� �������� � ��������
'-------------------------------------------------------------------------------
Function Online_PaySys_Check_Doc_In_Black_List(docNum)
    Dim exists : exists = False

    Call wTreeView.DblClickItem("|�� ����Ϧ ѳ������ ���|г������ �׳���� ������û�")
    If wMDIClient.WaitVBObject("frmPttel", 2000).Exists Then
        Do Until wMDIClient.vbObject("frmPttel").vbObject("tdbgView").EOF Or exists = True
            If Trim(wMDIClient.vbObject("frmPttel").vbObject("tdbgView").Columns.Item(1).text) = docNum Then
                exists = True
                Online_PaySys_Check_Doc_In_Black_List = exists
            Else
                Call wMDIClient.vbObject("frmPttel").vbObject("tdbgView").MoveNext
            End If
        Loop
    Else
        Online_PaySys_Check_Doc_In_Black_List = exists
        Log.Error("Workpapers folder view doesn't exists")
    End If
End Function

'------------------------------------------------------------------------------------------------------
'� �������� ������ �������� ѳٳ� ���� �� �������٦ ѳ�������ݻ�� ٳ��ٳ�� ��ջ�� ��ݳ�� ��������
'----------------------------------------------------------------------------------------------------
Function Online_PaySys_Check_Assertion_In_Black_List()
    Dim rcount : rcount = False  
    
    BuiltIn.Delay(3000)             
    Call wMainForm.MainMenu.Click(c_AllActions)
    Call wMainForm.PopupMenu.Click("????? �?? ?????????� ?????????????? ?????????")
    BuiltIn.Delay(1000)
    Call ClickCmdButton(2, "γ���")
    rcount = wMDIClient.vbObject("frmPttel_2").vbObject("tdbgView").VisibleRows
    BuiltIn.Delay(1000)
    wMDIClient.vbObject("frmPttel_2").Close()
    
    Online_PaySys_Check_Assertion_In_Black_List = rcount
End Function

'---------------------------------------------------------------------
' �������� ��ϳ����۳� �������� ��˳������� ������û����
'----------------------------------------------------------------------
'docNum - ��������
'startDate - �������� ������ ������(��ǽ�)
'endDate - �������� ������ ������(���)
Function Online_PaySys_Check_Doc_In_Workpapers(docNum, startDate, endDate)
  Dim is_exists : is_exists = false
  Dim colN
  
  BuiltIn.Delay(3000)
  Call wTreeView.DblClickItem("|г׳���� �������� � ��ٳ��� |��˳������� ������û�")
  If p1.WaitVBObject("frmAsUstPar", 6000).Exists Then
    Call Rekvizit_Fill("Dialog", 1, "General", "PERN", startDate)
    Call Rekvizit_Fill("Dialog", 1, "General", "PERK", endDate)
    Call Rekvizit_Fill("Dialog", 1, "General", "USER", "^A" & "[Del]")
    Call ClickCmdButton(2, "γ���")
  Else 
    Log.Error "Can't find frmAsUstPar window", "", pmNormal, ErrorColor
  End If

  If wMDIClient.WaitVBObject("frmPttel", 6000).Exists Then
    colN = wMDIClient.vbObject("frmPttel").GetColumnIndex("DOCNUM")
    BuiltIn.Delay(3000)
    If SearchInPttel("frmPttel", colN, docNum) Then
      is_exists = true
    End If
  Else
    Log.Message "The sending documnet frmPttel doesn't exist", "", pmNormal, ErrorColor
  End If
  
  Online_PaySys_Check_Doc_In_Workpapers = is_exists
End Function

'----------------------------------------------------------------------
' �������� ��ϳ����۳� �������� 1-�� г������ ���
'----------------------------------------------------------------------
'docNum - ��������
'startDate - �������� ������ ������(��ǽ�)
'endDate - �������� ������ ������(���)
Function Online_PaySys_Check_Doc_In_Verifier(docNum, startDate, endDate)  
  Dim exists, verifyDocuments, colN
  exists = False
  
  BuiltIn.Delay(1000)
  Set verifyDocuments = New_VerificationDocument()
  verifyDocuments.User = "^A[Del]"
  Call GoToVerificationDocument("|г����� I ���|г������ �׳���� ������û�",verifyDocuments)
  BuiltIn.Delay(3000)
  If wMDIClient.WaitVBObject("frmPttel", delay_middle).Exists Then
    colN = wMDIClient.vbObject("frmPttel").GetColumnIndex("DOCNUM")
    BuiltIn.Delay(2000)
    If SearchInPttel("frmPttel", colN, docNum) Then
      exists = True
    End If
  Else
    Log.Error "Verifiers folder view doesn't exists", "", pmNormal, ErrorColor
  End If
  
  Online_PaySys_Check_Doc_In_Verifier = exists
End Function

'---------------------------------------------------------------------------------------------------------
' �׳�ٳ� ѳ��ݳ���� (Online ���. ��.) �������� ��ϳ����۳� ��������  ��ﳷ������ 먳��� ��ó�ݳ����
'---------------------------------------------------------------------------------------------------------
Function Online_PaySys_Check_Doc_In_Drafts(doc_isn)
  Dim exists : exists = False
  Dim colN
 
  'ꨳ��� ��ó�ݳ��� �������� ��ճ����� ��˳������� ������û� ��ó�ݳ�
  Call wTreeView.DblClickItem("|г׳���� �������� � ��ٳ��� |��ó�ݳ�ݻ�|��ﳷ������ 먳���")
  If p1.WaitVBObject("frmAsUstPar", 2000).Exists Then 
    Call ClickCmdButton(2, "γ���")
  Else 
    Log.Error "Can't find frmAsUstPar window", "", pmNormal, ErrorColor
  End If
    
  If wMDIClient.WaitVBObject("frmPttel", 3000).Exists Then
    colN = wMDIClient.vbObject("frmPttel").GetColumnIndex("FISN")
    If SearchInPttel("frmPttel", colN, doc_isn) Then
        exists = true
    End If
  Else
    Log.Message "The sending documnet frmPttel doesn't exist", "", pmNormal, ErrorColor
  End If
    
  Online_PaySys_Check_Doc_In_Drafts = exists
End Function