<%@ LANGUAGE=VBSCript%>
<%Option Explicit%> 
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<%

Call LoadBasisGlobalInf()

    On Error Resume Next
    Err.Clear

    Call HideStatusWnd                                                               '��: Hide Processing message
	Dim lgOpModeCRUD

    lgOpModeCRUD      = Request("txtMode")                                           '��: Read Operation Mode (CRUD)

    Select Case lgOpModeCRUD 
        Case CStr(UID_M0001)                                                         '��: Query
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '��: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '��: Delete
             Call SubBizDelete()
    End Select


'========================================================================================
Sub SubBizQuery()
    On Error Resume Next
    Err.Clear
End Sub
'========================================================================================
Sub SubBizSave()
    On Error Resume Next
    Err.Clear
End Sub
'========================================================================================
Sub SubBizDelete()
    On Error Resume Next
    Err.Clear
End Sub

'========================================================================================
Sub SubBizQueryMulti()

    Dim iPABG010
    Dim iStrData
    Dim exportData
    Dim exportData1
    Dim iLngRow,iLngCol
    Dim iStrPrevKey
    Dim iStrCtrlCd
    Dim iIntMaxRows
    Dim iIntQueryCount
    Dim importArray
    Dim iIntLoopCount
	Dim E1_a_ctrl_item
    Dim txtCtrlCd

    Const C_SHEETMAXROWS_D   = 100

	Const i_a_ctrl_item_ctrl_cd = 0
	Const i_a_ctrl_item_ctrl_nm	= 1

    Const C_QueryConut	   = 0
    Const C_MaxQueryReCord = 1
    Const C_Ctrl_Cd   = 2

    On Error Resume Next
    Err.Clear

	txtCtrlCd		= Request("txtCtrlCd")
	iStrPrevKey		= Request("lgStrPrevKey")
    iIntMaxRows		= Request("txtMaxRows")
    iIntQueryCount	= Request("lgPageNo")
    If Trim(iStrPrevKey) = "" Then
		iStrCtrlCd	= txtCtrlCd
	Else
		iStrCtrlCd	= iStrPrevKey
    End If

    If Len(Trim(iIntQueryCount))  Then
       If Isnumeric(iIntQueryCount) Then
          iIntQueryCount = CInt(iIntQueryCount)
       End If
    Else
       iIntQueryCount = 0
    End If

    ReDim importArray(2)
    importArray(C_QueryConut)		= iIntQueryCount
    importArray(C_MaxQueryReCord)	= C_SHEETMAXROWS_D
    importArray(C_Ctrl_Cd)			= iStrCtrlCd

	Set iPABG010 = Server.CreateObject("PABG010.cAListCtlItmSvr")

    If CheckSYSTEMError(Err, True) = True Then
       Set iPABG010 = Nothing
       Exit Sub
    End If

	Call iPABG010.A_List_Ctrl_Item_Svr (gStrGlobalCollection, importArray, txtCtrlCd, E1_a_ctrl_item, exportData)

    If CheckSYSTEMError(Err, True) = True Then
		Set iPABG010 = Nothing
		Exit Sub
    End If

    Set iPABG010 = Nothing

	iStrData = ""
	iIntLoopCount = 0

	For iLngRow = 0 To UBound(exportData, 1)
		iIntLoopCount = iIntLoopCount + 1

	    If  iIntLoopCount < (C_SHEETMAXROWS_D + 1) Then

			iStrData = iStrData & Chr(11) & ConvSPChars(exportData(iLngRow, 0))			'�������ڵ� 
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData(iLngRow, 1)))	'�����׸� 
			iStrData = iStrData & Chr(11) & ConvSPChars(exportData(iLngRow, 2))			'�ý��۱��� 
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData(iLngRow, 3)))	'�����׸񿵹��� 
			iStrData = iStrData & Chr(11) & exportData(iLngRow, 4)						'�ڷ����� 
			istrData = iStrData & Chr(11) & ""											'�ڷ����� 
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData(iLngRow, 5)))	'�ڷ���� 
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData(iLngRow, 6)))	'��ǥ�����׸� 
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData(iLngRow, 7)))	'��ǥ�����׸�� 
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData(iLngRow, 8)))	'���̺�ID
			istrData = iStrData & Chr(11) & ""											'���̺�ID �˾� 
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData(iLngRow, 9)))	'�÷�ID
			istrData = iStrData & Chr(11) & ""											'�÷� ID �˾� 
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData(iLngRow, 10)))	'�÷��� ID
			istrData = iStrData & Chr(11) & ""											'�÷��� ID �˾� 
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData(iLngRow, 11)))	'�����ڵ� 
			iStrData = iStrData & Chr(11) & ""											'�����ڵ� �˾� 
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData(iLngRow, 12)))	'�����ڵ�� 
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData(iLngRow, 13)))	'kEY �÷�1
			istrData = iStrData & Chr(11) & ""											'kEY �÷��˾�1
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData(iLngRow, 14)))	'�ڷ�����1
			istrData = iStrData & Chr(11) & "" 											'�ڷ�����1
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData(iLngRow, 15)))	'kEY �÷�2
			istrData = iStrData & Chr(11) & "" 											'kEY �÷��˾�2
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData(iLngRow, 16)))	'�ڷ�����2
			istrData = iStrData & Chr(11) & "" 											'�ڷ�����2
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData(iLngRow, 17)))	'kEY �÷�3
			istrData = iStrData & Chr(11) & "" 											'kEY �÷��˾�3
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData(iLngRow, 18)))	'�ڷ�����3
			istrData = iStrData & Chr(11) & "" 											'�ڷ�����3
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData(iLngRow, 19)))	'kEY �÷�4
			istrData = iStrData & Chr(11) & "" 											'kEY �÷��˾�4
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData(iLngRow, 20)))	'�ڷ�����4
			istrData = iStrData & Chr(11) & ""											'�ڷ�����4
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData(iLngRow, 21)))	'kEY �÷�5
			istrData = iStrData & Chr(11) & "" 											'kEY �÷��˾�5
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData(iLngRow, 22)))	'�ڷ�����5
			istrData = iStrData & Chr(11) & "" 											'�ڷ�����5
			iStrData = iStrData & Chr(11) & iIntMaxRows + iLngRow + 1
			iStrData = iStrData & Chr(11) & Chr(12)

	    Else
			iStrPrevKey = exportData(UBound(exportData, 1), 0)
			iIntQueryCount = iIntQueryCount + 1
			Exit For
		End If
	Next

	If  iIntLoopCount < (C_SHEETMAXROWS_D + 1) Then
		iStrPrevKey = ""
	    iIntQueryCount = ""
	End If

	Response.Write " <Script Language=vbscript>	                        " & vbCr
	Response.Write " With parent                                        " & vbCr
    Response.Write "	.ggoSpread.Source = .frm1.vspdData              " & vbCr
    Response.Write "	.ggoSpread.SSShowData """ & iStrData       & """" & vbCr
    Response.Write "	.lgPageNo = """ & iIntQueryCount           & """" & vbCr
    Response.Write "	.lgStrPrevKey = """ & ConvSPChars(iStrPrevKey)          & """" & vbCr
    Response.Write "	.frm1.hCtrlCd.value =	""" & ConvSPChars(E1_a_ctrl_item(i_a_ctrl_item_ctrl_cd))          & """" & vbCr
    Response.Write "	.frm1.txtCtrlNM.value = """ & ConvSPChars(E1_a_ctrl_item(i_a_ctrl_item_ctrl_nm))          & """" & vbCr
	
    Response.Write "	.DbQueryOk                                      " & vbCr
    Response.Write " End With                                           " & vbCr
    Response.Write " </Script>                                          " & vbCr

End Sub

'========================================================================================
Sub SubBizSaveMulti()
    Dim iPABG010
    Dim iErrorPosition


    On Error Resume Next
    Err.Clear

	Set iPABG010 = Server.CreateObject ("PABG010.cAMngCtlItmSvr")

    If CheckSYSTEMError(Err, True) = True Then
       Set iPABG010 = Nothing
       Exit Sub
    End If

    Call iPABG010.A_Manage_Ctrl_Item_Svr(gStrGlobalCollection, Trim(request("txtSpread")), iErrorPosition)

  	If CheckSYSTEMError2(Err, True,iErrorPosition & "��","","","","") = True Then
       Set iPABG010 = Nothing
       Exit Sub
    End If

    Set iPABG010 = Nothing

	Response.Write " <Script Language=vbscript> " & vbCr
	Response.Write " parent.DbSaveOk            " & vbCr
    Response.Write "</Script>                   " & vbCr

End Sub

'========================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    On Error Resume Next
    Err.Clear
End Sub

'========================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
    On Error Resume Next
    Err.Clear
End Sub

'========================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
    On Error Resume Next
    Err.Clear
End Sub

'========================================================================================
Sub SubMakeSQLStatements(pDataType,arrColVal)
    Dim iSelCount
    On Error Resume Next
End Sub

'========================================================================================
Sub CommonOnTransactionCommit()
End Sub

'========================================================================================
Sub CommonOnTransactionAbort()
End Sub

'========================================================================================
Sub SetErrorStatus()
End Sub
'========================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next
    Err.Clear
End Sub

%>
