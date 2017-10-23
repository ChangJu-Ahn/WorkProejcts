<%@ LANGUAGE=VBSCript%>
<%Option Explicit%> 
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<%

Call LoadBasisGlobalInf()

    On Error Resume Next
    Err.Clear

    Call HideStatusWnd                                                               '☜: Hide Processing message
	Dim lgOpModeCRUD
	Dim StrQry

    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
	Dim StrID
	StrID		= Request("txtPGMID2")
	
    Select Case lgOpModeCRUD 
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQueryMulti()
			 CALL SubBizQueryMulti2()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
             Call SubBizSaveMulti2()
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

    Dim PZCG043
    Dim iStrData
    Dim exportData
    Dim iLngRow,iLngCol
    Dim iStrPrevKey
    Dim iStrMPGMCd
    Dim iIntMaxRows
    Dim iIntQueryCount
    Dim importArray
    Dim iIntLoopCount
	Dim E1_Z_MOVE_PGM
    Dim txtPGMID
    Dim txtLANG

    Const C_SHEETMAXROWS_D   = 100

	Const i_Z_MOVE_ITEM_PGM_cd = 0
	Const i_Z_MOVE_ITEM_PGM_NM	= 1

    Const C_QueryConut	   = 0
    Const C_MaxQueryReCord = 1
    Const C_MPGM_Cd   = 2

    On Error Resume Next
    Err.Clear

	txtPGMID		= Request("txtPGMID")
	txtLANG			= Request("txtLANG")
	iStrPrevKey		= Request("lgStrPrevKey")
    iIntMaxRows		= Request("txtMaxRows")
    iIntQueryCount	= Request("lgPageNo")
    
    
    If Trim(iStrPrevKey) = "" Then
		iStrMPGMCd	= txtPGMID
	Else
		iStrMPGMCd	= iStrPrevKey
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
    importArray(C_MPGM_Cd)			= iStrMPGMCd

	Set PZCG043 = Server.CreateObject("PZCG043.cAListMvpgmkeySvr")

    If CheckSYSTEMError(Err, True) = True Then
       Set IZCCG040 = Nothing
       Exit Sub
    End If

	Call PZCG043.A_List_Key_pgm_Svr (gStrGlobalCollection, importArray, txtPGMID,txtLANG, E1_Z_MOVE_PGM, exportData)

    If CheckSYSTEMError(Err, True) = True Then
		Set PZCG043 = Nothing
		Exit Sub
    End If

    Set PZCG043 = Nothing

	iStrData = ""
	iIntLoopCount = 0
	
	For iLngRow = 0 To UBound(exportData, 1)
		iIntLoopCount = iIntLoopCount + 1

	    If  iIntLoopCount < (C_SHEETMAXROWS_D + 1) Then

			iStrData = iStrData & Chr(11) & ConvSPChars(exportData(iLngRow, 0))			'키 값 
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData(iLngRow, 1)))	'키 필드명 
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData(iLngRow, 2)))	'표시명칭 
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData(iLngRow, 3)))	'히든 PGM_ID
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
    Response.Write "	.frm1.txtPGMID.value =	""" & ConvSPChars(E1_Z_MOVE_PGM(i_Z_MOVE_ITEM_PGM_cd))          & """" & vbCr
    Response.Write "	.frm1.txtPGMID2.value =	""" & ConvSPChars(E1_Z_MOVE_PGM(i_Z_MOVE_ITEM_PGM_cd))          & """" & vbCr
    Response.Write "	.frm1.txtPGMNM.value = """ & ConvSPChars(E1_Z_MOVE_PGM(i_Z_MOVE_ITEM_PGM_NM))          & """" & vbCr
    Response.Write "	.frm1.txtPGMNM2.value = """ & ConvSPChars(E1_Z_MOVE_PGM(i_Z_MOVE_ITEM_PGM_NM))          & """" & vbCr
	Response.Write " End With                                           " & vbCr
    Response.Write " </Script>											" & vbCr
	
	

End Sub

Sub SubBizQueryMulti2()
	Dim PZCG044
    Dim iStrData1
    Dim exportData1
    Dim iLngRow,iLngCol
    Dim iStrPrevKey
    Dim iStrMPGMCd
    Dim iIntMaxRows2
    Dim iIntQueryCount
    Dim importArray
    Dim iIntLoopCount
	Dim E1_Z_MOVE_PGM
    Dim txtPGMID
    Dim txtLANG

	Const C_SHEETMAXROWS_D   = 100

	Const i_Z_MOVE_ITEM_PGM_cd = 0
	Const i_Z_MOVE_ITEM_PGM_NM	= 1

    Const C_QueryConut	   = 0
    Const C_MaxQueryReCord = 1
    Const C_MPGM_Cd   = 2

    On Error Resume Next
    Err.Clear

	txtPGMID		= Request("txtPGMID")
	txtLANG			= Request("txtLANG")
	iStrPrevKey		= Request("lgStrPrevKey")
    iIntMaxRows2	= Request("txtMaxRows2")
    iIntQueryCount	= Request("lgPageNo")
    
    If Trim(iStrPrevKey) = "" Then
		iStrMPGMCd	= txtPGMID
	Else
		iStrMPGMCd	= iStrPrevKey
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
    importArray(C_MPGM_Cd)			= iStrMPGMCd

 Set PZCG044 = Server.CreateObject("PZCG044.cAListMvpgmqrySvr")

    If CheckSYSTEMError(Err, True) = True Then
       Set PZCG044 = Nothing
       Exit Sub
    End If

	Call PZCG044.A_List_qry_pgm_Svr (gStrGlobalCollection, importArray, txtPGMID,txtLANG, E1_Z_MOVE_PGM, exportData1)

    If CheckSYSTEMError(Err, True) = True Then
		Set PZCG044 = Nothing
		Exit Sub
    End If

    Set PZCG044 = Nothing

	iStrData1 = ""
	iIntLoopCount = 0
	StrQry = ""
	For iLngRow = 0 To UBound(exportData1, 1)
		iIntLoopCount = iIntLoopCount + 1

	    If  iIntLoopCount < (C_SHEETMAXROWS_D + 1) Then

			iStrData1 = iStrData1 & Chr(11) & ConvSPChars(exportData1(iLngRow, 0))			'항목 값 
			iStrData1 = iStrData1 & Chr(11) & ""
			iStrData1 = iStrData1 & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, 1)))	'항목 명 
			iStrData1 = iStrData1 & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, 2)))	'쿼리문1 
			iStrData1 = iStrData1 & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, 2)))	'쿼리문1 
			iStrData1 = iStrData1 & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, 3)))	'히든 PGM_ID
			iStrData1 = iStrData1 & Chr(11) & ConvSPChars(exportData1(iLngRow, 0))			'히든항목 값 
			iStrData1 = iStrData1 & Chr(11) & iIntMaxRows2 + iLngRow + 1
			iStrData1 = iStrData1 & Chr(11) & Chr(12)
			StrQry = StrQry & ConvSPChars(Trim(exportData1(iLngRow, 2))) & chr(11)
			
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
    Response.Write "	.ggoSpread.Source = .frm1.vspdData2              " & vbCr
    Response.Write "	.ggoSpread.SSShowData """ & iStrData1       & """" & vbCr
    Response.Write "	.lgPageNo = """ & iIntQueryCount           & """" & vbCr
    Response.Write "	.lgStrPrevKey = """ & ConvSPChars(iStrPrevKey)          & """" & vbCr
    Response.Write "	.frm1.txtPGMID.value =	""" & ConvSPChars(E1_Z_MOVE_PGM(i_Z_MOVE_ITEM_PGM_cd))          & """" & vbCr
    Response.Write "	.frm1.txtPGMID2.value =	""" & ConvSPChars(E1_Z_MOVE_PGM(i_Z_MOVE_ITEM_PGM_cd))          & """" & vbCr
    Response.Write "	.frm1.txtPGMNM.value = """ & ConvSPChars(E1_Z_MOVE_PGM(i_Z_MOVE_ITEM_PGM_NM))          & """" & vbCr
    Response.Write "	.frm1.txtPGMNM2.value = """ & ConvSPChars(E1_Z_MOVE_PGM(i_Z_MOVE_ITEM_PGM_NM))          & """" & vbCr
    Response.Write "	.frm1.hquery.value =	""" &   StrQry        & """" & vbCr
    Response.Write "	.DbQueryOk                                      " & vbCr
    Response.Write " End With                                           " & vbCr
    Response.Write " </Script>											" & vbCr
End Sub

'========================================================================================
Sub SubBizSaveMulti()
    Dim PZCG043
    Dim iErrorPosition


    On Error Resume Next
    Err.Clear

	if request("txtSpread") = "" then

		Exit sub
	End if
	
	Set PZCG043 = Server.CreateObject ("PZCG043.cAMngMvPGMKEYSvr")

    If CheckSYSTEMError(Err, True) = True Then
       Set PZCG043 = Nothing
       Exit Sub
    End If

    Call PZCG043.A_MANAGE_KEY_PGM_SVR(gStrGlobalCollection, Trim(request("txtSpread")), iErrorPosition)

  	If CheckSYSTEMError2(Err, True,iErrorPosition & "행","","","","") = True Then
       Set PZCG043 = Nothing
       Exit Sub
    End If

   
	Response.Write " <Script Language=vbscript> " & vbCr
	Response.Write "	parent.frm1.txtPGMID.value =	""" & StrID & """" & vbCr
	'Response.Write " parent.DbSaveOk            " & vbCr
    Response.Write "</Script>                   " & vbCr

End Sub

Sub SubBizSaveMulti2()
    
    Dim PZCG044
    Dim iErrorPosition


    On Error Resume Next
    Err.Clear
    

  
    Set PZCG044 = Server.CreateObject ("PZCG044.cAMngMvPGMqrySvr")

    If CheckSYSTEMError(Err, True) = True Then
       Set PZCG044 = Nothing
       Exit Sub
    End If

    Call PZCG044.A_MANAGE_QRY_PGM_SVR(gStrGlobalCollection, Trim(request("txtSpread2")), iErrorPosition)

  	If CheckSYSTEMError2(Err, True,iErrorPosition & "행","","","","") = True Then
       Set PZCG044 = Nothing
       Exit Sub
    End If

    Set PZCG044 = Nothing

	Response.Write " <Script Language=vbscript> " & vbCr
	Response.Write "	parent.frm1.txtPGMID.value =	""" & StrID          & """" & vbCr
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
