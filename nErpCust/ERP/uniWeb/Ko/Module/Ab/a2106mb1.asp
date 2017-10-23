<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<%Option Explicit%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<%


Call LoadBasisGlobalInf()

    'Dim lgOpModeCRUD
    On Error Resume Next
    Err.Clear


    Call HideStatusWnd                                                               '¢Ð: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '¢Ð: Query
          '  Call SubBizQuery()
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '¢Ð: Save,Update
          '  Call SubBizSave()
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '¢Ð: Delete
             Call SubBizDelete()
    End Select


'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next
    Err.Clear
End Sub

'============================================================================================================
Sub SubBizSave()
    On Error Resume Next
    Err.Clear
End Sub

'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next
    Err.Clear
End Sub

'============================================================================================================
Sub SubBizQueryMulti()

    Dim iPABG030
    Dim I1_a_jnl_item_jnl_cd
    Dim E1_a_jnl_item_jnl_nm 
    Dim Export_group
    Dim iStrPrevKey
	Dim txtJnlCd
    Dim iIntMaxRows    
    Dim iIntQueryCount
	Dim importArray

    ReDim importArray(2)
    ReDim E1_a_jnl_item_jnl_nm(1)

    Dim iStrData
    Dim iIntLoopCount
    Dim iLngRow,iLngCol

    Const C_SHEETMAXROWS_D  = 100 

	Const C_QueryConut	   = 0
    Const C_MaxQueryReCord = 1
    Const C_Jnl_Cd   = 2

	Const l2_a_jnl_item_jnl_cd = 0
	Const l2_a_jnl_item_jnl_nm = 1


	Const EG_JNL_CD = 0
	Const EG_JNL_NM = 1
	Const EG_JNL_ENG_NM = 2
	Const EG_JNL_TYPE = 3
	Const EG_SYS_FG = 4
	Const EG_TRNS_TBL_NM = 5
	Const EG_TRNS_COLM_NM = 6

    On Error Resume Next
    Err.Clear


	iStrPrevKey		= Request("lgStrPrevKey")
	txtJnlCd		= Request("txtJnlCd")
    iIntMaxRows		= Request("txtMaxRows")
    iIntQueryCount	= Request("lgPageNo")
    
    If Trim(iStrPrevKey) = "" Then
		I1_a_jnl_item_jnl_cd	= txtJnlCd
	Else
		I1_a_jnl_item_jnl_cd	= iStrPrevKey
    End If

	importArray(C_QueryConut)		= iIntQueryCount
    importArray(C_MaxQueryReCord)	= C_SHEETMAXROWS_D
    importArray(C_Jnl_Cd)			= I1_a_jnl_item_jnl_cd

	Set iPABG030 = Server.CreateObject("PABG030.cAListJnlItemSvr")
    If CheckSYSTEMError(Err, True) = True Then
       Set iPABG030 = Nothing
       Exit Sub
    End If

	Call iPABG030.A_LIST_JNL_ITEM_SVR(gStrGlobalCollection, importArray, txtJnlCd, E1_a_jnl_item_jnl_nm, Export_group)

    If CheckSYSTEMError(Err, True) = True Then
       Set iPABG030 = Nothing
       Exit Sub
    End If

    Set iPABG030 = Nothing

	iStrData = ""
	iIntLoopCount = 0

	For iLngRow = 0 To UBound(Export_group, 1)
		iIntLoopCount = iIntLoopCount + 1

	    If  iIntLoopCount < (C_SHEETMAXROWS_D + 1) Then

				istrData = istrData & Chr(11) & ConvSPChars(Trim(Export_group(iLngRow, EG_JNL_CD)))
				istrData = istrData & Chr(11) & ConvSPChars(Trim(Export_group(iLngRow, EG_JNL_NM)))
				istrData = istrData & Chr(11) & ConvSPChars(Trim(Export_group(iLngRow, EG_JNL_ENG_NM)))
				istrData = istrData & Chr(11) & Trim(Export_group(iLngRow, EG_JNL_TYPE))
				istrData = istrData & Chr(11) & " "

				Select Case Trim(Export_group(iLngRow, EG_SYS_FG))
					Case "Y"
						istrData = istrData & Chr(11) & 1
					Case "N"
						istrData = istrData & Chr(11) & 0
					case else
						istrData = istrData & Chr(11) & 0
				End Select

				istrData = istrData & Chr(11) & ConvSPChars(Trim(Export_group(iLngRow, EG_TRNS_TBL_NM)))
				istrData = istrData & Chr(11) & ConvSPChars(Trim(Export_group(iLngRow, EG_TRNS_COLM_NM)))

				istrData = istrData & Chr(11) & iIntMaxRows + iIntLoopCount
				istrData = istrData & Chr(11) & Chr(12)

	    Else
			iStrPrevKey = Export_group(UBound(Export_group, 1), 0)
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
		Response.Write "	.frm1.txtJnlNM.value = """ & ConvSPChars(E1_a_jnl_item_jnl_nm(l2_a_jnl_item_jnl_nm))    & """" & vbCr
		Response.Write "	.frm1.hJnlCd.value = """ & ConvSPChars(E1_a_jnl_item_jnl_nm(l2_a_jnl_item_jnl_cd))    & """" & vbCr	
		Response.Write "	.lgPageNo = """ & iIntQueryCount    & """" & vbCr
		Response.Write "	.lgStrPrevKey = """ & ConvSPChars(iStrPrevKey)    & """" & vbCr
		Response.Write "	.DbQueryOk " & vbCr
		Response.Write " End With   " & vbCr
		Response.Write " </Script>  " & vbCr

End Sub    

'============================================================================================================
Sub SubBizSaveMulti()
    On Error Resume Next
    Err.Clear
    Dim iPABG030
    Dim iErrorPosition

    Set iPABG030 = Server.CreateObject("PABG030.cAMngJnlItemSvr")

    If CheckSYSTEMError(Err, True) = True Then
       Set iPABG030 = Nothing
       Exit Sub
    End If

'	 Response.Write Trim(Request("txtSpread"))
	 'Response.Write "I1_a_jnl_item_jnl_cd=" 
	'Response.End

    Call iPABG030.A_MANAGE_JNL_ITEM_SVR(gStrGloBalCollection, Request("txtSpread"), iErrorPosition)
  	If CheckSYSTEMError2(Err, True,iErrorPosition & "Çà","","","","") = True Then
       Set iPABG030 = Nothing
       Exit Sub
    End If

    Set iPABG030 = Nothing
	
	Response.Write " <Script Language=vbscript> " & vbCr
	Response.Write " parent.DbSaveOk            " & vbCr
    Response.Write "</Script>                   " & vbCr
End Sub


'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    On Error Resume Next
    Err.Clear
End Sub


'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
    On Error Resume Next
    Err.Clear
End Sub


'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
    On Error Resume Next
    Err.Clear
End Sub


'============================================================================================================
Sub SubMakeSQLStatements(pDataType,arrColVal)
    Dim iSelCount
    On Error Resume Next
End Sub


'============================================================================================================
Sub CommonOnTransactionCommit()
End Sub


'============================================================================================================
Sub CommonOnTransactionAbort()
End Sub


'============================================================================================================
Sub SetErrorStatus()
End Sub

'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next
    Err.Clear
End Sub
%>
