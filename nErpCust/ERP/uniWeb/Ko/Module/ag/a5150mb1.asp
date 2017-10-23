<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<%
Response.Expires = -1								'☜ : ASP가 캐쉬되지 않도록 한다.
Response.Buffer = True								'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.
%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../ag/incAcctMBFunc.asp"  -->

<% 
	Call LoadBasisGlobalInf() 
    Call LoadInfTB19029B("I", "*","NOCOOKIE","MB")
    Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")
%>
<%
    On Error Resume Next                                                            '☜: Protect system from crashing
    Err.Clear                                                                       '☜: Clear Error status

' 권한관리 추가
Dim lgBizAreaAuthYn, lgAuthBizAreaCd, lgAuthBizAreaNm								' 사업장
Dim lgInternalAuthYn, lgInternalCd													' 내부부서
Dim lgSubInternalAuthYn, lgSubInternalCd											' 내부부서(하위포함)
Dim lgAuthUsrIDAuthYn, lgAuthUsrID													' 개인

    Call HideStatusWnd                                                              '☜: Hide Processing message

    lgOpModeCRUD = Request("txtMode")												'☜: Read Operation Mode (CRUD)
	lgIntFlgMode = CInt(Request("txtFlgMode"))										'☜: 저장시 Create/Update 판별     

	' 권한관리 추가
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID			= Trim(Request("lgAuthUsrID"))
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
'            Call SubBizQuery()
            Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
			Call SubBizSave()
'            Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
            Call SubBizDelete()
'            Call SubBizDeleteMulti()
    End Select

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
'    On Error Resume Next                                                             '☜: Protect system from crashing
'    Err.Clear                                                                        '☜: Clear Error status

End Sub    
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDeleteMulti()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	On Error Resume Next
    Err.Clear    
	
    Const A320_11_allc_no = 0

    Const A320_E1_allc_no = 0
    Const A320_E1_allc_dt = 1
    Const A320_E1_org_change_id = 2
    Const A320_E1_dept_cd = 3
    Const A320_E1_dept_nm = 4
    Const A320_E1_bp_cd = 5
    Const A320_E1_bp_nm = 6
    Const A320_E1_doc_cur = 7
    Const A320_E1_temp_gl_no = 8
    Const A320_E1_gl_no = 9
    Const A320_E1_allc_desc = 10
    Const A320_E1_dr_loc_amt = 11
    Const A320_E1_cr_loc_amt = 12
    Const A320_E1_diff_loc_amt = 13
    
    Const A320_EG1_open_type = 0
    Const A320_EG1_open_type_nm = 1
    Const A320_EG1_gl_no = 2
    Const A320_EG1_open_dt = 3
    Const A320_EG1_bp_cd = 4
    Const A320_EG1_bp_nm = 5
    Const A320_EG1_open_acct_cd = 6
    Const A320_EG1_open_acct_nm = 7
    Const A320_EG1_dr_cr_fg = 8
    Const A320_EG1_dr_cr_nm = 9
    Const A320_EG1_doc_cur = 10
    Const A320_EG1_open_amt = 11
    Const A320_EG1_bal_amt = 12
    Const A320_EG1_cls_amt = 13
    Const A320_EG1_cls_loc_amt = 14
    Const A320_EG1_dc_amt = 15
    Const A320_EG1_dc_loc_amt = 16
    Const A320_EG1_item_desc = 17
    Const A320_EG1_due_dt = 18
    Const A320_EG1_open_no = 19
    Const A320_EG1_open_gl_seq = 20
    Const A320_EG1_org_change_id = 21
    Const A320_EG1_dept_cd = 22
    Const A320_EG1_dept_nm = 23
    Const A320_EG1_biz_area_cd = 24
    Const A320_EG1_biz_area_nm = 25
    Const A320_EG1_xch_rate = 26
    
    Const A320_EG2_Item_seq = 0
    Const A320_EG2_dept_cd = 1
    Const A320_EG2_dept_nm = 2
    Const A320_EG2_acct_cd = 3
    Const A320_EG2_acct_nm = 4
    Const A320_EG2_dr_cr_fg = 5
    Const A320_EG2_dr_cr_nm = 6
    Const A320_EG2_doc_cur = 7
    Const A320_EG2_exch_rate = 8
    Const A320_EG2_item_amt = 9
    Const A320_EG2_item_loc_amt = 10
    Const A320_EG2_item_desc = 11

	Dim iPAGG030
	Dim iStrData
	Dim iStrData2
    Dim I1_a_allc_info
    Dim E1_a_allc_hdr
    Dim EG1_export_allc_info
    Dim EG2_export_open_item

    Dim iLngRow
    Dim iLngRow2
    Dim iLngCol
    Dim iStrCurrency

	On Error Resume Next
    Err.Clear    

    ReDim I1_a_allc_info(4)
    I1_a_allc_info(A320_11_allc_no) = Trim(Request("txtAllcNo"))
    I1_a_allc_info(A320_11_allc_no+1) = lgAuthBizAreaCd
    I1_a_allc_info(A320_11_allc_no+2) = lgInternalCd
    I1_a_allc_info(A320_11_allc_no+3) = lgSubInternalCd
    I1_a_allc_info(A320_11_allc_no+4) = lgAuthUsrID

    Set iPAGG030 = Server.CreateObject("PAGG030.cALkUpMultiClsSvr")

    If CheckSYSTEMError(Err, True) = True Then
		Response.Write "<Script Language=vbscript>                     " & vbcr
		Response.Write " With Parent                                   " & vbCr				
		Response.Write " 	Call .ggoOper.ClearField(.Document, ""2"") " & vbCr
		Response.Write " End With                                      " & vbCr
		Response.Write "</Script>                                      " & vbCr
		Exit Sub
    End If

	Call iPAGG030.A_LOOKUP_MULTI_CLS_SVR(gStrGlobalCollection, I1_a_allc_info, E1_a_allc_hdr, EG1_export_allc_info,EG2_export_open_item)

	If CheckSYSTEMError(Err, True) = True Then
		Response.Write "<Script Language=vbscript>                     " & vbcr
		Response.Write " With Parent                                   " & vbCr
		Response.Write " 	Call .ggoOper.ClearField(.Document, ""2"") " & vbCr
		Response.Write " End With                                      " & vbCr
		Response.Write "</Script>                                      " & vbCr
		Set iPAGG030 = Nothing
		Exit Sub
    End If    

    Set iPAGG030 = Nothing

    iStrData = ""	

	For iLngRow = 0 To UBound(EG1_export_allc_info, 1) 
		iStrCurrency = ConvSPChars(EG1_export_allc_info(iLngRow, A320_EG1_doc_cur))

		iStrData = iStrData & Chr(11) & ConvSPChars(EG1_export_allc_info(iLngRow, A320_EG1_open_type))
		iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_allc_info(iLngRow, A320_EG1_open_type_nm)))
		iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_allc_info(iLngRow, A320_EG1_gl_no)))
		iStrData = iStrData & Chr(11) & UNIDateClientFormat(EG1_export_allc_info(iLngRow, A320_EG1_open_dt))
		iStrData = iStrData & Chr(11) & ConvSPChars(EG1_export_allc_info(iLngRow, A320_EG1_bp_cd))
		iStrData = iStrData & Chr(11) & ConvSPChars(EG1_export_allc_info(iLngRow, A320_EG1_bp_nm))
		iStrData = iStrData & Chr(11) & ConvSPChars(EG1_export_allc_info(iLngRow, A320_EG1_open_acct_cd))
		iStrData = iStrData & Chr(11) & ConvSPChars(EG1_export_allc_info(iLngRow, A320_EG1_open_acct_nm))
		iStrData = iStrData & Chr(11) & ConvSPChars(EG1_export_allc_info(iLngRow, A320_EG1_dr_cr_fg))
		iStrData = iStrData & Chr(11) & ConvSPChars(EG1_export_allc_info(iLngRow, A320_EG1_dr_cr_nm))		
		iStrData = iStrData & Chr(11) & ConvSPChars(UCase(EG1_export_allc_info(iLngRow, A320_EG1_doc_cur)))
		iStrData = iStrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_allc_info(iLngRow, A320_EG1_open_amt),iStrCurrency,ggAmtOfMoneyNo, "X" , "X")
		iStrData = iStrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_allc_info(iLngRow, A320_EG1_bal_amt),iStrCurrency,ggAmtOfMoneyNo, "X" , "X")
		iStrData = iStrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_allc_info(iLngRow, A320_EG1_cls_amt),iStrCurrency,ggAmtOfMoneyNo, "X" , "X")
		iStrData = iStrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_allc_info(iLngRow, A320_EG1_cls_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo , "X")
		iStrData = iStrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_allc_info(iLngRow, A320_EG1_dc_amt),iStrCurrency,ggAmtOfMoneyNo, "X" , "X")
		iStrData = iStrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_allc_info(iLngRow, A320_EG1_dc_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo , "X")			
		iStrData = iStrData & Chr(11) & ConvSPChars(EG1_export_allc_info(iLngRow, A320_EG1_item_desc))
		iStrData = iStrData & Chr(11) & UNIDateClientFormat(EG1_export_allc_info(iLngRow, A320_EG1_due_dt))
		iStrData = iStrData & Chr(11) & ConvSPChars(EG1_export_allc_info(iLngRow, A320_EG1_open_no))
		iStrData = iStrData & Chr(11) & EG1_export_allc_info(iLngRow, A320_EG1_open_gl_seq)
		iStrData = iStrData & Chr(11) & ConvSPChars(EG1_export_allc_info(iLngRow, A320_EG1_org_change_id))
		iStrData = iStrData & Chr(11) & ConvSPChars(UCase(EG1_export_allc_info(iLngRow, A320_EG1_dept_cd)))
		iStrData = iStrData & Chr(11) & ConvSPChars(EG1_export_allc_info(iLngRow, A320_EG1_dept_nm))
		iStrData = iStrData & Chr(11) & ConvSPChars(UCase(EG1_export_allc_info(iLngRow, A320_EG1_biz_area_cd)))
		iStrData = iStrData & Chr(11) & ConvSPChars(EG1_export_allc_info(iLngRow, A320_EG1_biz_area_nm))										
		iStrData = iStrData & Chr(11) & UNINumClientFormat(EG1_export_allc_info(iLngRow, A320_EG1_xch_rate),ggExchRate.DecPoint, 0)
		iStrData = iStrData & Chr(11) & iLngRow + 1
		iStrData = iStrData & Chr(11) & Chr(12)		
	Next

    iStrData2 = ""	
	If Not IsEmpty(EG2_export_open_item) Then
		For iLngRow2 = 0 To UBound(EG2_export_open_item, 1) 
			iStrCurrency = ConvSPChars(EG2_export_open_item(iLngRow2, A320_EG2_doc_cur))

			iStrData2 = iStrData2 & Chr(11) & EG2_export_open_item(iLngRow2, A320_EG2_Item_seq)
			iStrData2 = iStrData2 & Chr(11) & ConvSPChars(UCase(EG2_export_open_item(iLngRow2, A320_EG2_dept_cd)))
			iStrData2 = iStrData2 & Chr(11) & ""
			iStrData2 = iStrData2 & Chr(11) & ConvSPChars(EG2_export_open_item(iLngRow2, A320_EG2_dept_nm))
			iStrData2 = iStrData2 & Chr(11) & ConvSPChars(UCase(EG2_export_open_item(iLngRow2, A320_EG2_acct_cd)))
			iStrData2 = iStrData2 & Chr(11) & ""
			iStrData2 = iStrData2 & Chr(11) & ConvSPChars(EG2_export_open_item(iLngRow2, A320_EG2_acct_nm))
			iStrData2 = iStrData2 & Chr(11) & ConvSPChars(EG2_export_open_item(iLngRow2, A320_EG2_dr_cr_fg))
			iStrData2 = iStrData2 & Chr(11) & ConvSPChars(EG2_export_open_item(iLngRow2, A320_EG2_dr_cr_nm))
			iStrData2 = iStrData2 & Chr(11) & ConvSPChars(UCase(EG2_export_open_item(iLngRow2, A320_EG2_doc_cur)))
			iStrData2 = iStrData2 & Chr(11) & ""
			iStrData2 = iStrData2 & Chr(11) & UNINumClientFormat(EG2_export_open_item(iLngRow2, A320_EG2_exch_rate),ggExchRate.DecPoint, 0)
			iStrData2 = iStrData2 & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG2_export_open_item(iLngRow2, A320_EG2_item_amt),iStrCurrency,ggAmtOfMoneyNo, "X" , "X")
			iStrData2 = iStrData2 & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG2_export_open_item(iLngRow2, A320_EG2_item_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo , "X")
			iStrData2 = iStrData2 & Chr(11) & ConvSPChars(EG2_export_open_item(iLngRow2, A320_EG2_item_desc))
			iStrData2 = iStrData2 & Chr(11) & iLngRow2 + 1
			iStrData2 = iStrData2 & Chr(11) & Chr(12)
		Next
	End If
	    
	Response.Write "<Script Language=vbscript>" & vbcr
	Response.Write " With Parent  																																					 " & vbCr		
	Response.Write " 	.frm1.txtAllcNo.value		= """ & ConvSPChars(UCase(Trim(E1_a_allc_hdr(A320_E1_allc_no))))															& """" & vbCr
	Response.Write "	.frm1.txtAllcDt.Text		= """ & UNIDateClientFormat(E1_a_allc_hdr(A320_E1_allc_dt))																	& """" & vbCr
	Response.Write " 	.frm1.hOrgChangeId.value	= """ & ConvSPChars(E1_a_allc_hdr(A320_E1_org_change_id))																	& """" & vbCr
	Response.Write " 	.frm1.txtDeptCd.value		= """ & ConvSPChars(UCase(E1_a_allc_hdr(A320_E1_dept_cd)))																	& """" & vbCr
	Response.Write " 	.frm1.txtDeptNm.Value		= """ & ConvSPChars(E1_a_allc_hdr(A320_E1_dept_nm))																			& """" & vbCr
	Response.Write " 	.frm1.txtBpCd.value			= """ & ConvSPChars(UCase(Trim(E1_a_allc_hdr(A320_E1_bp_cd))))																& """" & vbCr
	Response.Write " 	.frm1.txtBpNm.value			= """ & ConvSPChars(E1_a_allc_hdr(A320_E1_bp_nm))																			& """" & vbCr
	Response.Write " 	.frm1.txtDocCur.value		= """ & ConvSPChars(E1_a_allc_hdr(A320_E1_doc_cur))																			& """" & vbCr
	Response.Write " 	.frm1.hDocCur.value		    = """ & ConvSPChars(E1_a_allc_hdr(A320_E1_doc_cur))																			& """" & vbCr
	Response.Write " 	.frm1.txtTempGlNo.value		= """ & ConvSPChars(E1_a_allc_hdr(A320_E1_temp_gl_no))																		& """" & vbCr
	Response.Write " 	.frm1.txtGlNo.value			= """ & ConvSPChars(E1_a_allc_hdr(A320_E1_gl_no))																			& """" & vbCr	
	Response.Write " 	.frm1.txtDesc.Value		    = """ & ConvSPChars(E1_a_allc_hdr(A320_E1_allc_desc))																		& """" & vbCr
	Response.Write " 	.frm1.txtDrLocAmt.Text		= """ & UNIConvNumDBToCompanyByCurrency(E1_a_allc_hdr(A320_E1_dr_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")	& """" & vbCr							
	Response.Write " 	.frm1.txtCrLocAmt.Text		= """ & UNIConvNumDBToCompanyByCurrency(E1_a_allc_hdr(A320_E1_cr_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")	& """" & vbCr
	Response.Write " 	.frm1.txtDiffLocAmt.Text	= """ & UNIConvNumDBToCompanyByCurrency(E1_a_allc_hdr(A320_E1_diff_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")	& """" & vbCr							
	Response.Write " 	.frm1.htxtAllcNo.value		= """ & ConvSPChars(UCase(Trim(E1_a_allc_hdr(A320_E1_allc_no))))															& """" & vbCr
	Response.Write " 	.ggoSpread.Source = .frm1.vspdData4																															 " & vbCr
	Response.Write "	.ggoSpread.SSShowData """ & iStrData   & """ ,""F""																											 " & vbCr
	Response.Write  "    Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData4," & 1 & "," & iLngRow  & ",.C_DOC_CUR,.C_OPEN_AMT, ""A"" ,""I"",""X"",""X"")						 " & vbCr
	Response.Write  "    Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData4," & 1 & "," & iLngRow  & ",.C_DOC_CUR,.C_BAL_AMT,  ""A"" ,""I"",""X"",""X"")						 " & vbCr
	Response.Write  "    Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData4," & 1 & "," & iLngRow  & ",.C_DOC_CUR,.C_CLS_AMT,  ""A"" ,""I"",""X"",""X"")						 " & vbCr
	Response.Write  "    Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData4," & 1 & "," & iLngRow  & ",.C_DOC_CUR,.C_DC_AMT,   ""A"" ,""I"",""X"",""X"")						 " & vbCr
	Response.Write  "    Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData4," & 1 & "," & iLngRow  & ",.C_DOC_CUR,.C_XCH_RATE, ""D"" ,""I"",""X"",""X"")						 " & vbCr
	Response.Write " 	.ggoSpread.Source = .frm1.vspdData																															 " & vbCr
	Response.Write "	.ggoSpread.SSShowData """ & iStrData2   & """ ,""F""																										 " & vbCr
	Response.Write  "    Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData," & 1 & "," & iLngRow2  & ",.C_DocCur,.C_ItemAmt, ""A"" ,""I"",""X"",""X"")							 " & vbCr
	Response.Write  "    Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData," & 1 & "," & iLngRow2  & ",.C_DocCur,.C_ExchRate, ""D"" ,""I"",""X"",""X"")						 " & vbCr
	Response.Write " 	.DbQueryOk																																					 " & vbCr
	Response.Write " End With																																						 " & vbCr
	Response.Write "</Script>																																						 " & vbCr 
End Sub

'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
	On Error Resume Next	
	Err.Clear
	
    Const A380_I2_org_change_id = 0
    Const A380_I2_dept_cd = 1

    Const A380_I3_allc_no = 0
    Const A380_I3_allc_dt = 1
    Const A380_I3_bp_cd = 2
    Const A380_I3_doc_cur = 3
    Const A380_I3_allc_desc = 4
    Const A380_I3_rcpt_fg = 5

	Const A380_I4_a_data_auth_data_BizAreaCd = 0
	Const A380_I4_a_data_auth_data_internal_cd = 1
	Const A380_I4_a_data_auth_data_sub_internal_cd = 2
	Const A380_I4_a_data_auth_data_auth_usr_id = 3

	Dim iCommandSent
    Dim I1_a_acct_trans_type			' 통합반제거래유형    
    Dim I2_b_acct_dept					' 통합반제부서정보
    Dim I3_a_allc_hdr					' 통합반제헤더정보
	Dim I4_a_data_auth					' 권한관리용
    Dim txtSpread						' 미결반제
    Dim txtSpread1						' 추가계정
    Dim txtSpread3						' 추가계정관리항목
	
	Dim iPAGG030
	Dim iStrRetAllcNo

	If lgIntFlgMode = OPMD_CMODE Then
		iCommandSent = "CREATE"
	Elseif lgIntFlgMode = OPMD_UMODE Then
		iCommandSent = "UPDATE"	
	End If	

	I1_a_acct_trans_type = "AG001"
	
	Redim I2_b_acct_dept(1)
	I2_b_acct_dept(A380_I2_org_change_id) = Trim(Request("hOrgChangeId"))
	I2_b_acct_dept(A380_I2_dept_cd)       = Trim(Request("txtDeptCd")) 

	Redim I3_a_allc_hdr(5)
	I3_a_allc_hdr(A380_I3_allc_no)   = Trim(Request("txtAllcNo"))
	I3_a_allc_hdr(A380_I3_allc_dt)   = Trim(Request("txtAllcDt")) 
	I3_a_allc_hdr(A380_I3_bp_cd)     = Trim(Request("txtBpCd")) 
	I3_a_allc_hdr(A380_I3_doc_cur)   = Trim(Request("txtDocCur"))
	I3_a_allc_hdr(A380_I3_allc_desc) = Trim(Request("txtDesc"))
	I3_a_allc_hdr(A380_I3_rcpt_fg)   = "U"

	Redim I4_a_data_auth(3)
	I4_a_data_auth(A380_I4_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I4_a_data_auth(A380_I4_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I4_a_data_auth(A380_I4_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I4_a_data_auth(A380_I4_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))

	txtSpread  = Trim(Request("txtSpread"))
	txtSpread1 = Trim(Request("txtSpread1"))
	txtSpread3 = Trim(Request("txtSpread3"))

	'--------------------------------------------------------------------
	'실행하기.
	'--------------------------------------------------------------------	
	Set iPAGG030 = Server.CreateObject("PAGG030.cAMngMultiClsSvr")

	If CheckSYSTEMError(Err, True) = True Then
		Exit Sub
    End If

	iStrRetAllcNo = iPAGG030.A_MANAGE_MULTI_CLS_SVR(gStrGlobalCollection, iCommandSent, I1_a_acct_trans_type, I2_b_acct_dept, I3_a_allc_hdr,txtSpread,txtSpread1, txtSpread3,I4_a_data_auth) 	

	If CheckSYSTEMError(Err, True) = True Then		
		Set iPAGG030 = Nothing
		Exit Sub
    End If

    Set iPAGG030  = Nothing

	Response.Write " <Script Language=vbscript>											" & vbCr
	Response.Write " With parent														" & vbCr
    Response.Write "	.DbSaveOk """ & ConvSPChars(iStrRetAllcNo)	&			     """" & vbCr    
    Response.Write " End With															" & vbCr
    Response.Write " </Script>															" & vbCr
End Sub    

'============================================================================================================
' Name : SubBizDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizDelete()
	On Error Resume Next
	Err.Clear

    Const A380_I3_allc_no = 0
    Const A380_I3_allc_dt = 1
    Const A380_I3_bp_cd = 2
    Const A380_I3_doc_cur = 3
    Const A380_I3_allc_desc = 4
    Const A380_I3_rcpt_fg = 5

	Const A380_I4_a_data_auth_data_BizAreaCd = 0
	Const A380_I4_a_data_auth_data_internal_cd = 1
	Const A380_I4_a_data_auth_data_sub_internal_cd = 2
	Const A380_I4_a_data_auth_data_auth_usr_id = 3

	Dim iPAGG030					
	Dim iCommandSent				' 구분자
	Dim I3_a_allc_hdr				' 반제헤더정보
	Dim I4_a_data_auth				' 권한관리용
	
	Dim iStrRetAllcNo		

	iCommandSent = "DELETE"

	Redim I3_a_allc_hdr(5)
	I3_a_allc_hdr(A380_I3_allc_no)   = Trim(Request("txtAllcNo"))
	I3_a_allc_hdr(A380_I3_allc_dt)   = Trim(Request("txtAllcDt")) 
	I3_a_allc_hdr(A380_I3_bp_cd)     = ""
	I3_a_allc_hdr(A380_I3_doc_cur)   = ""
	I3_a_allc_hdr(A380_I3_allc_desc) = ""
	I3_a_allc_hdr(A380_I3_rcpt_fg)   = ""
	
	Redim I4_a_data_auth(3)
	I4_a_data_auth(A380_I4_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I4_a_data_auth(A380_I4_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I4_a_data_auth(A380_I4_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I4_a_data_auth(A380_I4_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))

	Set iPAGG030 = Server.CreateObject("PAGG030.cAMngMultiClsSvr")

	If CheckSYSTEMError(Err, True) = True Then
		Exit Sub
    End If

	iStrRetAllcNo = iPAGG030.A_MANAGE_MULTI_CLS_SVR(gStrGlobalCollection, iCommandSent, , , I3_a_allc_hdr)

	If CheckSYSTEMError(Err, True) = True Then
		Set iPAGG030 = Nothing
		Exit Sub
    End If

    Set iPAGG030  = Nothing

	Response.Write " <Script Language=vbscript>	                        " & vbCr
	Response.Write " With parent                                        " & vbCr
    Response.Write "	.DbDeleteOK										" & vbCr    
    Response.Write " End With											" & vbCr
    Response.Write " </Script>											" & vbCr
End Sub

%>

