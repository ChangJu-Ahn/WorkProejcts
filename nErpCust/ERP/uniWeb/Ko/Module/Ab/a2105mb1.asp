<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<%


Call LoadBasisGlobalInf()


Dim lgOpModeCRUD

On Error Resume Next
Err.Clear

Call HideStatusWnd

lgOpModeCRUD = Request("txtMode")	'☜ : 현재 상태를 받음 


Select Case lgOpModeCRUD
	Case CStr(UID_M0001)                                                         '☜: Query
		 Call SubBizQueryMulti()
	Case CStr(UID_M0002)                                                         '☜: Save,Update
		 Call SubBizSaveMulti()
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

	Dim PABG025LIST
	Dim iBizUnitCd
	Dim iStrData
	Dim iTransTypeNm
	Dim iTransTypeGp
	Dim iLngRow,iLngCol
	Dim txtTransType

	'==============신규================
	Dim iIntQueryCount
	Dim iIntLoopCount
	Dim iStrTransType
	Dim iStrPrevKey
	Dim iIntMaxRows
	'==================================

    Const A122_EG2_E1_a_acct_trans_type_trans_cd = 0
    Const A122_EG2_E1_a_acct_trans_type_trans_nm = 1
    Const A122_EG2_E1_a_acct_trans_type_trans_eng_nm = 2
    Const A122_EG2_E1_a_acct_trans_type_batch_fg = 3
    Const A122_EG2_E1_a_acct_trans_type_batch_fg_nm = 4
    Const A122_EG2_E1_a_acct_trans_type_gl_posting_fg = 5
    Const A122_EG2_E1_a_acct_trans_type_gl_posting_fg_nm = 6 
    Const A122_EG2_E1_a_acct_trans_type_mo_cd = 7
	Const A122_EG2_E1_a_acct_trans_type_mo_nm = 8    
    Const A122_EG2_E1_a_acct_trans_type_sys_fg = 9
    Const A122_EG2_E1_a_acct_trans_type_reverse_fg = 10
    Const A122_EG2_E1_a_acct_trans_type_reverse_fg_nm = 11    
    Const A122_EG2_E1_a_acct_trans_type_acct_sum_fg = 12
    Const A122_EG2_E1_a_acct_trans_type_acct_arrayal_fg = 13
    Const A122_EG2_E1_b_company_inv_post_fg = 14

	Const C_SHEETMAXROWS_D = 100

    On Error Resume Next
    Err.Clear

	txtTransType = Request("txtTransType")
	iStrPrevKey  = Trim(Request("lgStrPrevKey"))

	If iStrPrevKey = "" Then
		iStrPrevKey = txtTransType
	Else
		iStrPrevKey	= iStrPrevKey
	End IF
	
    iIntMaxRows		= Request("txtMaxRows")
    iIntQueryCount	= Request("lgPageNo")

    If Len(Trim(iIntQueryCount))  Then                                        '☜ : Chnage Nextkey str into int value
       If Isnumeric(iIntQueryCount) Then
          iIntQueryCount = CInt(iIntQueryCount)
       End If
    Else
       iIntQueryCount = 0
    End If
    
	If CheckSYSTEMError(Err, True) = True Then
      Exit Sub
    End If

    Set PABG025LIST = Server.CreateObject("PABG025.cAListAcctTrTpSvr")

	If CheckSYSTEMError(Err, True) = True Then
       Exit Sub
    End If

	Redim iTransTypeNm(1) 
	Const C_TRANS_TYPE = 0
	Const C_TRANS_TYPE_NM = 1

    Call PABG025LIST.C_LISTL_ACCT_TRANS_TYPE_SVR(gStrGlobalCollection,C_SHEETMAXROWS_D,iStrPrevKey,txtTransType,iTransTypeNm,iTransTypeGp)

	If CheckSYSTEMError(Err, True) = True Then
		Set PABG025LIST = Nothing
		Exit Sub
    End If

    Set PABG025LIST = Nothing

    iStrData = ""
    iStrPrevKey = ""
    iIntLoopCount = 0

	For iLngRow = 0 To UBound(iTransTypeGp, 1)
		iIntLoopCount = iIntLoopCount + 1

		If  iIntLoopCount < (C_SHEETMAXROWS_D + 1) Then
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(iTransTypeGp(iLngRow, A122_EG2_E1_a_acct_trans_type_trans_cd)))
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(iTransTypeGp(iLngRow, A122_EG2_E1_a_acct_trans_type_trans_nm)))
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(iTransTypeGp(iLngRow, A122_EG2_E1_a_acct_trans_type_trans_eng_nm)))
			iStrData = iStrData & Chr(11) & Trim(iTransTypeGp(iLngRow, A122_EG2_E1_a_acct_trans_type_batch_fg))
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(iTransTypeGp(iLngRow, A122_EG2_E1_a_acct_trans_type_batch_fg_nm)))
			iStrData = iStrData & Chr(11) & Trim(iTransTypeGp(iLngRow, A122_EG2_E1_a_acct_trans_type_gl_posting_fg))
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(iTransTypeGp(iLngRow, A122_EG2_E1_a_acct_trans_type_gl_posting_fg_nm)))
			iStrData = iStrData & Chr(11) & Trim(iTransTypeGp(iLngRow, A122_EG2_E1_a_acct_trans_type_mo_cd))
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(iTransTypeGp(iLngRow, A122_EG2_E1_a_acct_trans_type_mo_nm)))
			iStrData = iStrData & Chr(11) & Trim(iTransTypeGp(iLngRow, A122_EG2_E1_a_acct_trans_type_sys_fg))
			iStrData = iStrData & Chr(11) & Trim(iTransTypeGp(iLngRow, A122_EG2_E1_a_acct_trans_type_reverse_fg))
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(iTransTypeGp(iLngRow, A122_EG2_E1_a_acct_trans_type_reverse_fg_nm)))
			iStrData = iStrData & Chr(11) & Trim(iTransTypeGp(iLngRow, A122_EG2_E1_a_acct_trans_type_acct_sum_fg))
			iStrData = iStrData & Chr(11) & Trim(iTransTypeGp(iLngRow, A122_EG2_E1_a_acct_trans_type_acct_arrayal_fg))
			iStrData = iStrData & Chr(11) & Trim(iTransTypeGp(iLngRow, A122_EG2_E1_b_company_inv_post_fg))
			iStrData = iStrData & Chr(11) & Cstr(iIntMaxRows + iLngRow + 1)
			iStrData = iStrData & Chr(11) & Chr(12)
		Else
			iStrPrevKey = iTransTypeGp(UBound(iTransTypeGp, 1), 0)
			iIntQueryCount = iIntQueryCount + 1
			Exit For
		End If
	Next

	If  iIntLoopCount < (C_SHEETMAXROWS_D + 1) Then
		iStrPrevKey = ""
	    iIntQueryCount = ""
	End If

	Response.Write " <Script Language=vbscript>															 " & vbCr
	Response.Write " With parent																		 " & vbCr
    Response.Write "	.ggoSpread.Source		= .frm1.vspdData										 " & vbCr
    Response.Write "	.ggoSpread.SSShowData	  """ & iStrData									& """" & vbCr
    Response.Write "	.frm1.txtTransNM.value	= """ & ConvSPChars(iTransTypeNm(C_TRANS_TYPE_NM))	& """" & vbCr
    Response.Write "	.frm1.hTransType.value	= """ & ConvSPChars(iTransTypeNm(C_TRANS_TYPE))		& """" & vbCr
    Response.Write "	.lgPageNo				= """ & iIntQueryCount								& """" & vbCr
    Response.Write "	.lgStrPrevKey			= """ & iStrPrevKey									& """" & vbCr
    Response.Write "	.DbQueryOk																		 " & vbCr
    Response.Write "End With																			 " & vbCr
    Response.Write "</Script>																			 " & vbCr

End Sub

'============================================================================================================
Sub SubBizSaveMulti()
	Dim PABG025CUD
'	Dim arrRowVal,arrColVal
    Dim iErrorPosition

    Redim iErrorPosition(10)

    On Error Resume Next
    Err.Clear

    Set PABG025CUD = Server.CreateObject("PABG025.cAMngAcctTrTpSvr")

    If CheckSYSTEMError(Err, True) = True Then
       Exit Sub
    End If

    Call PABG025CUD.C_MANAGE_ACCT_TRANS_TYPE_SVR(gStrGlobalCollection,Request("txtSpread"),iErrorPosition)

  	If CheckSYSTEMError2(Err, True,iErrorPosition & "행","","","","") = True Then
       Set PABG025CUD = Nothing
       Exit Sub
    End If

    Set PABG025CUD = nothing

	Response.Write " <Script Language=vbscript> " & vbCr
	Response.Write " parent.DbSaveOk            " & vbCr
    Response.Write " </Script>                  " & vbCr
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
