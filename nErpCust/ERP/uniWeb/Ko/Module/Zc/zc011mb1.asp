<%@ LANGUAGE=VBSCript%>
<%Option Explicit%> 
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<%

Call LoadBasisGlobalInf()

    On Error Resume Next
    Err.Clear

    Call HideStatusWnd                                                               '☜: Hide Processing message
	Dim lgOpModeCRUD

    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)

    Select Case lgOpModeCRUD 
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
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

    Dim PZCG050
    Dim iStrData
    Dim exportData
    Dim exportData1
    Dim iLngRow,iLngCol
    Dim iStrPrevKey
    Dim iStrMItemCd
    Dim iIntMaxRows
    Dim iIntQueryCount
    Dim importArray
    Dim iIntLoopCount
	Dim E1_Z_MOVE_ITEM
    Dim txtMoveItem
    Dim txtLANG

    Const C_SHEETMAXROWS_D   = 100

	Const i_Z_MOVE_ITEM_PGM_cd = 0
	Const i_Z_MOVE_ITEM_PGM_NM	= 1

    Const C_QueryConut	   = 0
    Const C_MaxQueryReCord = 1
    Const C_MItem_Cd   = 2

    On Error Resume Next
    Err.Clear

	txtMoveItem		= Request("txtMoveItem")
	txtLANG			= Request("txtLANG")
	iStrPrevKey		= Request("lgStrPrevKey")
    iIntMaxRows		= Request("txtMaxRows")
    iIntQueryCount	= Request("lgPageNo")
    If Trim(iStrPrevKey) = "" Then
		iStrMItemCd	= txtMoveItem
	Else
		iStrMItemCd	= iStrPrevKey
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
    importArray(C_MItem_Cd)			= iStrMItemCd

	Set PZCG050 = Server.CreateObject("PZCG050.cAListMvpgmSvr")

    If CheckSYSTEMError(Err, True) = True Then
       Set IZCCG040 = Nothing
       Exit Sub
    End If

	Call PZCG050.A_List_Move_PGM_Svr (gStrGlobalCollection, importArray, txtMoveItem,txtLANG, E1_Z_MOVE_ITEM, exportData)

    If CheckSYSTEMError(Err, True) = True Then
		Set PZCG050 = Nothing
		Exit Sub
    End If

    Set PZCG050 = Nothing

	iStrData = ""
	iIntLoopCount = 0

	For iLngRow = 0 To UBound(exportData, 1)
		iIntLoopCount = iIntLoopCount + 1

	    If  iIntLoopCount < (C_SHEETMAXROWS_D + 1) Then

			iStrData = iStrData & Chr(11) & ConvSPChars(exportData(iLngRow, 0))			'관리항코드 
			istrData = iStrData & Chr(11) & ""
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData(iLngRow, 1)))	'관리항명 
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData(iLngRow, 2)))	'종합코드 
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData(iLngRow, 3)))	'VALUE
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData(iLngRow, 4)))	'방향 
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData(iLngRow, 5)))	'방향 
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData(iLngRow, 6)))	'다음프로그램 
			istrData = iStrData & Chr(11) & ""											'
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData(iLngRow, 7)))	'다음프로그램명 
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData(iLngRow, 8)))	'비고 
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData(iLngRow, 3)))	'VALUE
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData(iLngRow, 4)))	'방향 
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData(iLngRow, 6)))	'다음프로그램 
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
    Response.Write "	.frm1.txtMoveItem.value =	""" & ConvSPChars(E1_Z_MOVE_ITEM(i_Z_MOVE_ITEM_PGM_cd))          & """" & vbCr
    Response.Write "	.frm1.txtMoveItemNM.value = """ & ConvSPChars(E1_Z_MOVE_ITEM(i_Z_MOVE_ITEM_PGM_NM))          & """" & vbCr
	
    Response.Write "	.DbQueryOk                                      " & vbCr
    Response.Write " End With                                           " & vbCr
    Response.Write " </Script>                                          " & vbCr

End Sub

'========================================================================================
Sub SubBizSaveMulti()
    Dim PZCG050
    Dim iErrorPosition


    On Error Resume Next
    Err.Clear

	Set PZCG050 = Server.CreateObject ("PZCG050.cAMngMvPGMSvr")

    If CheckSYSTEMError(Err, True) = True Then
       Set PZCG050 = Nothing
       Exit Sub
    End If

    Call PZCG050.A_MANAGE_Move_PGM_SVR(gStrGlobalCollection, Trim(request("txtSpread")), iErrorPosition)

  	If CheckSYSTEMError2(Err, True,iErrorPosition & "행","","","","") = True Then
       Set PZCG050 = Nothing
       Exit Sub
    End If

    Set PZCG050 = Nothing

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
