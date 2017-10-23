<%@ LANGUAGE=VBSCript%>
<% Option Explicit %>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<%'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Loan
'*  3. Program ID           : f4240mb1
'*  4. Program Name         : �������ڿ���� 
'*  5. Program Desc         : 
'*  6. Comproxy List        : PAFG421
'*  7. Modified date(First) : 2002-01-02
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Ahn, do hyun
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'=======================================================================================================

'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call HideStatusWnd

On Error Resume Next

Call LoadBasisGlobalInf() 
Call LoadInfTB19029B("I","*", "NOCOOKIE", "MB")

Dim txtLoanNo

lgErrorStatus	= "NO"
lgErrorPos		= ""                                                           '��: Set to space
lgOpModeCRUD	= Request("txtMode")					'�� : ���� ���¸� ���� 

Const C_SHEETMAXROWS_D   = 30

Select Case lgOpModeCRUD
     Case CStr(UID_M0001)                                                         '��: Query
          Call SubBizQuery()

     Case CStr(UID_M0002)       
          Call SubBizSave()

     Case CStr(UID_M0003)                                                         '��: Delete
          Call SubBizDelete()
End Select
   
Response.End 


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                                 '��: Protect system from crashing
    Err.Clear                                                                            '��: Clear Error status

	Dim PAFG421LOOKUP	
	Dim I1_f_ln_info
	Dim E1_f_ln_info, EG1_export_group	
	Dim iLngRow,iLngCol
	Dim iIntLoopCount
	Dim iStrData
	Dim iStrPrevKey
	Dim iIntMaxRows
	Dim iIntQueryCount
	Dim lDecClsAmt(3)

    Const EA_f_ln_info_loan_cd = 0
    Const EA_f_ln_info_loan_nm = 1
    Const EA_f_ln_info_acct_cd = 2
    Const EA_f_ln_info_acct_nm = 3
    Const EA_f_ln_info_biz_area_nm_from = 4				'///�������� �߰� by JYK
    Const EA_f_ln_info_biz_area_nm_to = 5				'///�������� �߰� by JYK

    Const EA_f_ln_mon_dfr_choice_fg = 0
    Const EA_f_ln_mon_dfr_cls_fg = 1
    Const EA_f_ln_mon_dfr_seq = 2
    Const EA_f_ln_mon_dfr_loan_no = 3
    Const EA_f_ln_mon_dfr_loan_nm = 4
    Const EA_f_ln_mon_dfr_exp_acct_cd = 5
    Const EA_f_ln_mon_dfr_exp_acct_nm = 6
    Const EA_f_ln_mon_dfr_adv_int_acct_cd = 7
    Const EA_f_ln_mon_dfr_adv_int_acct_nm = 8
    Const EA_f_ln_mon_dfr_doc_cur = 9
    Const EA_f_ln_mon_dfr_xch_rate = 10
    Const EA_f_ln_mon_dfr_int_cls_plan_dt = 11
    Const EA_f_ln_mon_dfr_int_cls_amt = 12
    Const EA_f_ln_mon_dfr_int_cls_loc_amt = 13
    Const EA_f_ln_mon_dfr_int_cls_plan_amt = 14
    Const EA_f_ln_mon_dfr_int_cls_plan_loc_amt = 15
    Const EA_f_ln_mon_dfr_pay_dt = 16
    Const EA_f_ln_mon_dfr_pay_amt = 17
    Const EA_f_ln_mon_dfr_pay_loc_amt = 18
    Const EA_f_ln_mon_dfr_loan_int_rate = 19
    Const EA_f_ln_mon_dfr_loan_dt = 20
    Const EA_f_ln_mon_dfr_due_dt = 21
    Const EA_f_ln_mon_dfr_temp_gl_no = 22
    Const EA_f_ln_mon_dfr_gl_no = 23
    Const EA_f_ln_mon_dfr_conf_fg = 24
    Const EA_f_ln_mon_dfr_pay_no = 25
    Const EA_f_ln_mon_dfr_minor_nm = 26

    iStrPrevKey		= Trim(Request("lgStrPrevKey"))        
    iIntMaxRows		= Request("txtMaxRows")
    iIntQueryCount	= Request("lgPageNo")


    I1_f_ln_info	= Split(Request("txtKeyStream"), gColSep)

'    If Len(Trim(iIntQueryCount))  Then                                        '�� : Chnage Nextkey str into int value
'       If Isnumeric(iIntQueryCount) Then
'          iIntQueryCount = CInt(iIntQueryCount)
'       End If
'    Else   
'       iIntQueryCount = 0
'    End If

	Set PAFG421LOOKUP = Server.CreateObject("PAFG421.bFLkUpMonDfrSvr")

	If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
    End If
    

	Call PAFG421LOOKUP.F_LOOKUP_MON_DFR_SVR(gStrGloBalCollection, iStrPrevKey, C_SHEETMAXROWS_D, I1_f_ln_info, E1_f_ln_info, EG1_export_group)

    If CheckSYSTEMError(Err,True) = True Then
		Set PAFG421LOOKUP = nothing		
		%><Script Language=vbscript>Parent.frm1.txtBaseDt.focus</Script><%
		Response.Write "<Script Language=vbscript>		" & vbCr
		Response.Write " parent.frm1.txtBizAreaNm.value  = """ & ConvSPChars(E1_f_ln_info(EA_f_ln_info_biz_area_nm_from)) & """" & vbCr
		Response.Write " parent.frm1.txtBizAreaNm1.value = """ & ConvSPChars(E1_f_ln_info(EA_f_ln_info_biz_area_nm_to))   & """" & vbCr
		Response.Write " parent.DbQueryOk				" & vbCr
		Response.Write "</Script>						" & vbCr  
		Exit Sub
    End If
    Set PAFG421LOOKUP	 = nothing

	Response.Write "<Script Language=vbscript>	" & vbCr
	Response.Write " parent.frm1.txtLOAN_NM.value      = """ & ConvSPChars(E1_f_ln_info(EA_f_ln_info_loan_nm))          & """" & vbCr
	Response.Write " parent.frm1.txtIntExpAcctNm.value = """ & ConvSPChars(E1_f_ln_info(EA_f_ln_info_acct_nm))          & """" & vbCr
	Response.Write " parent.frm1.txtBizAreaNm.value    = """ & ConvSPChars(E1_f_ln_info(EA_f_ln_info_biz_area_nm_from)) & """" & vbCr
	Response.Write " parent.frm1.txtBizAreaNm1.value   = """ & ConvSPChars(E1_f_ln_info(EA_f_ln_info_biz_area_nm_to))   & """" & vbCr
    Response.Write "</Script>               " & vbCr

	iStrData = ""

	lDecClsAmt(0) = 0
	lDecClsAmt(1) = 0
	lDecClsAmt(2) = 0
	lDecClsAmt(3) = 0

	If IsEmpty(EG1_export_group) = False Then
		For iLngRow = 0 To UBound(EG1_export_group, 1)
			iStrData = iStrData & Chr(11) & "0"
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow,EA_f_ln_mon_dfr_minor_nm)))
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow,EA_f_ln_mon_dfr_seq)))
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow,EA_f_ln_mon_dfr_loan_no)))
			iStrData = iStrData & Chr(11) & ""
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow,EA_f_ln_mon_dfr_loan_nm)))
			If Trim(EG1_export_group(iLngRow,EA_f_ln_mon_dfr_cls_fg)) = "Y" then
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow,EA_f_ln_mon_dfr_exp_acct_cd)))
	 			iStrData = iStrData & Chr(11) & ""
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow,EA_f_ln_mon_dfr_exp_acct_nm)))
			Else
				iStrData = iStrData & Chr(11) & ""
	 			iStrData = iStrData & Chr(11) & ""
				iStrData = iStrData & Chr(11) & ""
			End If
			If Trim(EG1_export_group(iLngRow,EA_f_ln_mon_dfr_cls_fg)) = "Y" then
				iStrData = iStrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow,EA_f_ln_mon_dfr_int_cls_amt),ggAmtOfMoney.DecPoint	,0)
				lDecClsAmt(0) = lDecClsAmt(0) + UniCDbl(EG1_export_group(iLngRow,EA_f_ln_mon_dfr_int_cls_amt),0)
			Else
				If UniCDbl(EG1_export_group(iLngRow,EA_f_ln_mon_dfr_int_cls_amt), 0) = 0 Then
					iStrData = iStrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow,EA_f_ln_mon_dfr_int_cls_plan_amt),ggAmtOfMoney.DecPoint	,0)
					lDecClsAmt(0) = lDecClsAmt(0) + UniCDbl(EG1_export_group(iLngRow,EA_f_ln_mon_dfr_int_cls_plan_amt),0)
				Else
					iStrData = iStrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow,EA_f_ln_mon_dfr_int_cls_amt),ggAmtOfMoney.DecPoint	,0)
					lDecClsAmt(0) = lDecClsAmt(0) + UniCDbl(EG1_export_group(iLngRow,EA_f_ln_mon_dfr_int_cls_amt),0)
				End If
			End If
			iStrData = iStrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow,EA_f_ln_mon_dfr_int_cls_loc_amt),ggAmtOfMoney.DecPoint	,0)
			iStrData = iStrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow,EA_f_ln_mon_dfr_int_cls_plan_amt),ggAmtOfMoney.DecPoint	,0)
			iStrData = iStrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow,EA_f_ln_mon_dfr_int_cls_plan_loc_amt),ggAmtOfMoney.DecPoint	,0)
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow,EA_f_ln_mon_dfr_doc_cur)))
			iStrData = iStrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow,EA_f_ln_mon_dfr_xch_rate),ggExchRate.DecPoint	,0)
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow,EA_f_ln_mon_dfr_adv_int_acct_cd)))
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow,EA_f_ln_mon_dfr_adv_int_acct_nm)))
			iStrData = iStrData & Chr(11) & UNIDateClientFormat(Trim(EG1_export_group(iLngRow,EA_f_ln_mon_dfr_int_cls_plan_dt)))
			iStrData = iStrData & Chr(11) & UNIDateClientFormat(Trim(EG1_export_group(iLngRow,EA_f_ln_mon_dfr_pay_dt)))
			iStrData = iStrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow,EA_f_ln_mon_dfr_pay_amt),ggAmtOfMoney.DecPoint	,0)
			iStrData = iStrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow,EA_f_ln_mon_dfr_pay_loc_amt),ggAmtOfMoney.DecPoint	,0)
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow,EA_f_ln_mon_dfr_loan_int_rate)))
			iStrData = iStrData & Chr(11) & UNIDateClientFormat(Trim(EG1_export_group(iLngRow,EA_f_ln_mon_dfr_loan_dt)))
			iStrData = iStrData & Chr(11) & UNIDateClientFormat(Trim(EG1_export_group(iLngRow,EA_f_ln_mon_dfr_due_dt)))
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow,EA_f_ln_mon_dfr_temp_gl_no)))
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow,EA_f_ln_mon_dfr_gl_no)))
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow,EA_f_ln_mon_dfr_cls_fg)))
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow,EA_f_ln_mon_dfr_conf_fg)))
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow,EA_f_ln_mon_dfr_pay_no)))
		    iStrData = iStrData & Chr(11) & iIntMaxRows + iLngRow + 1
			iStrData = iStrData & Chr(11) & Chr(12)

			lDecClsAmt(1) = lDecClsAmt(1) + UniCDbl(EG1_export_group(iLngRow,EA_f_ln_mon_dfr_int_cls_loc_amt),0)
			lDecClsAmt(2) = lDecClsAmt(2) + UniCDbl(EG1_export_group(iLngRow,EA_f_ln_mon_dfr_int_cls_plan_amt),0)
			lDecClsAmt(3) = lDecClsAmt(3) + UniCDbl(EG1_export_group(iLngRow,EA_f_ln_mon_dfr_int_cls_plan_loc_amt),0)
		Next
	End If

	Response.Write "<Script Language=vbscript>	" & vbCr
	Response.Write " parent.frm1.txtAlcSum.value  = """ & UNINumClientFormat(lDecClsAmt(0),ggAmtOfMoney.DecPoint	,0) & """			" & vbCr
	Response.Write " parent.frm1.txtAlcLocSum.value  = """ & UNINumClientFormat(lDecClsAmt(1),ggAmtOfMoney.DecPoint	,0) & """			" & vbCr
	Response.Write " parent.frm1.txtPlanSum.value  = """ & UNINumClientFormat(lDecClsAmt(2),ggAmtOfMoney.DecPoint	,0) & """			" & vbCr
	Response.Write " parent.frm1.txtPlanLocSum.value  = """ & UNINumClientFormat(lDecClsAmt(3),ggAmtOfMoney.DecPoint	,0) & """			" & vbCr
    Response.Write "</Script>               " & vbCr

	Response.Write " <Script Language=vbscript> "							& vbCr
	Response.Write " With parent "											& vbCr
	Response.Write "	.ggoSpread.Source = .frm1.vspdData "				& vbcr
	Response.Write "    .frm1.vspdData.Redraw = False "						& vbCr
	Response.Write "	.ggoSpread.SSShowData  """ & iStrData & """ ,""F""" & vbCr
	Response.Write "	.lgPageNo			= """ & iIntQueryCount	& """"	& vbCr
	Response.Write "	.lgStrPrevKey		= """ & iStrPrevKey		& """"	& vbCr
	Response.Write "	Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData," & iIntMaxRows + 1 & "," & iIntMaxRows + iLngRow & ",.C_DOC_CUR, .C_INT_CLS_AMT,   ""A"" ,""I"",""X"",""X"")" & vbCr
	Response.Write "	Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData," & iIntMaxRows + 1 & "," & iIntMaxRows + iLngRow & ",.C_DOC_CUR, .C_INT_CLS_PLAN_AMT,   ""A"" ,""I"",""X"",""X"")" & vbCr
	Response.Write "	Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData," & iIntMaxRows + 1 & "," & iIntMaxRows + iLngRow & ",.C_DOC_CUR, .C_INT_PAY_AMT,   ""A"" ,""I"",""X"",""X"")" & vbCr
	Response.Write "	.DbQueryOk "										& vbcr
	Response.Write "    .frm1.vspdData.Redraw = True "						& vbCr
    Response.Write " End With "												& vbCr
    Response.Write " </Script> "											& vbCr

End Sub
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                                 '��: Protect system from crashing
    Err.Clear                                                                            '��: Clear Error status
    
	Dim PAFG421CU
	Dim iarrData
	Dim I1_f_ln_info
	Dim txtSpread
	
	iarrData = Split(Request("txtSpread"), gRowSep)
	
    Set PAFG421CU = server.CreateObject("PAFG421.cMngMonDfrSvr")   

    If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
    End If
     
	Call PAFG421CU.F_MANAGE_MON_DFR_SVR(gStrGlobalCollection, iarrData)

    If CheckSYSTEMError(Err,True) = True Then
		Set PAFG421CU = nothing
		Exit Sub	
    End If
	 
    Set PAFG421CU = nothing
    
	Response.Write "<Script Language=vbscript>		" & vbCr
	Response.Write " parent.DbSaveOk()				" & vbCr
	Response.Write "</Script>						" & vbCr  


End Sub

'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status    
End Sub    

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,arrColVal)
    Dim iSelCount
    On Error Resume Next

	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
End Sub


%>
