
<%
'**********************************************************************************************
'*  1. Module Name          : Account
'*  2. Function Name        : 채권관리 
'*  3. Program ID           : A3105mb2
'*  4. Program Name         : 입금등록및 채권반제 
'*  5. Program Desc         : 
'*  6. Comproxy List        : +B21011 (Manager)
'                             +B21019 (조회용)
'*  7. Modified date(First) : 2001/02/22
'*  8. Modified date(Last)  : 2001/02/22
'*  9. Modifier (First)     : Chang Sung Hee
'* 10. Modifier (Last)      : Chang Sung Hee
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/03/22 : ..........
'**********************************************************************************************




'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
' 아래 함수는 비지니스 로직 시작되는 시점에서 호출해 주세요..
Call HideStatusWnd		
On Error Resume Next														'☜: 

Call LoadBasisGlobalInf()

Dim pAr0041d																	'☆ : 조회용 ComProxy Dll 사용 변수	
Dim IntRows
Dim IntCols
Dim vbIntRet
Dim lEndRow
Dim boolCheck
Dim lgIntFlgMode
Dim LngMaxRow
Dim LngMaxRow1
Dim LngMaxRow3

' Com+ Conv. 변수 선언 
    
Dim strGroup
Dim arrCount

Dim iCommandSent 
Dim I1_a_acct_trans_type 
Dim I2_a_acct_trans_type 
Dim I3_b_biz_partner 
Dim I4_b_acct_dept 
Dim I5_a_rcpt 
Dim I6_b_bank 
Dim I7_b_bank_acct 
Dim I8_a_allc_rcpt 
Dim I9_a_rcpt_item 
Dim I10_b_currency 
Dim I11_a_rcpt_a_acct
Dim importArray																' Dim IG1_import_group_ar 
Dim importArray1															' Dim IG2_import_group_dc 
Dim importArray3															' Dim IG3_import_group_dc_dtl 
Dim E1_a_rcpt

'[CONVERSION INFORMATION]  View Name : import b_acct_dept
Const A379_I4_org_change_id = 0
Const A379_I4_dept_cd = 1

'[CONVERSION INFORMATION]  View Name : import a_rcpt
Const A379_I5_rcpt_no = 0
Const A379_I5_rcpt_dt = 1
Const A379_I5_doc_cur = 2
Const A379_I5_xch_rate = 3
Const A379_I5_bnk_chg_amt = 4
Const A379_I5_bnk_chg_loc_amt = 5
Const A379_I5_rcpt_amt = 6
Const A379_I5_rcpt_loc_amt = 7
Const A379_I5_rcpt_type = 8
Const A379_I5_rcpt_desc = 9
Const A379_I5_insrt_user_id = 10
Const A379_I5_updt_user_id = 11
Const A379_I5_ref_no = 12
Const A379_I5_rcpt_sts = 13
Const A379_I5_allc_amt = 14
Const A379_I5_allc_loc_amt = 15
Const A162_I2_insrt_dt = 16
Const A162_I2_updt_dt = 17
Const A379_I5_rcpt_fg = 18

'[CONVERSION INFORMATION]  Group Name : import_group_ar
 Const A379_IG1_I1_ief_supplied_select_char = 0
 '[CONVERSION INFORMATION]  View Name : import_cls_ar a_acct
 Const A379_IG1_I2_a_acct_acct_cd = 1
 '[CONVERSION INFORMATION]  View Name : import a_open_ar
 Const A379_IG1_I3_a_open_ar_ar_no = 2
 Const A379_IG1_I3_a_open_ar_ar_dt = 3
 '[CONVERSION INFORMATION]  View Name : import a_cls_ar
 Const A379_IG1_I4_a_cls_ar_cls_dt = 4
 Const A379_IG1_I4_a_cls_ar_ar_due_dt = 5
 Const A379_IG1_I4_a_cls_ar_doc_cur = 6
 Const A379_IG1_I4_a_cls_ar_diff_kind_cur = 7
 Const A379_IG1_I4_a_cls_ar_xch_rate = 8
 Const A379_IG1_I4_a_cls_ar_cls_amt = 9
 Const A379_IG1_I4_a_cls_ar_cls_loc_amt = 10
 Const A379_IG1_I4_a_cls_ar_dc_amt = 11
 Const A379_IG1_I4_a_cls_ar_dc_loc_amt = 12
 Const A379_IG1_I4_a_cls_ar_cls_ar_desc = 13

 '[CONVERSION INFORMATION]  Group Name : improt_group_dc
 '[CONVERSION INFORMATION]  View Name : import_dc ief_supplied
 Const A379_IG2_I1_ief_supplied_select_char = 0
 '[CONVERSION INFORMATION]  View Name : import a_rcpt_dc
 Const A379_IG2_I2_a_rcpt_dc_seq = 1
 Const A379_IG2_I2_a_rcpt_dc_dc_amt = 2
 Const A379_IG2_I2_a_rcpt_dc_dc_loc_amt = 3
 Const A379_IG2_I2_a_rcpt_dc_dc_desc = 4
 '[CONVERSION INFORMATION]  View Name : import_dc a_acct
 Const A379_IG2_I3_a_acct_acct_cd = 5

'[CONVERSION INFORMATION]  Group Name : import_group_dc_dtl
'[CONVERSION INFORMATION]  View Name : import_dc_dtl ief_supplied
Const A379_IG3_I1_ief_supplied_select_char = 0
'[CONVERSION INFORMATION]  View Name : import_dc_dtl a_rcpt_dc
Const A379_IG3_I2_a_rcpt_dc_seq = 1
'[CONVERSION INFORMATION]  View Name : import_dc_dtl a_ctrl_item
Const A379_IG3_I3_a_ctrl_item_ctrl_cd = 2
'[CONVERSION INFORMATION]  View Name : import a_rcpt_dc_dtl
Const A379_IG3_I4_a_rcpt_dc_dtl_dtl_seq = 3
Const A379_IG3_I4_a_rcpt_dc_dtl_ctrl_val = 4

'[CONVERSION INFORMATION]  View Name : import a_allc_rcpt
Const A379_I8_dc_amt = 0
Const A379_I8_dc_loc_amt = 1
Const A379_I8_allc_no = 2
Const A379_I8_insrt_user_id = 3
Const A379_I8_insrt_dt = 4
Const A379_I8_allc_dt = 5
Const A379_I8_allc_type = 6
Const A379_I8_ref_no = 7
Const A379_I8_allc_amt = 8
Const A379_I8_allc_loc_amt = 9
Const A379_I8_updt_user_id = 10
Const A379_I8_updt_dt = 11
Const A379_I8_prrcpt_no = 12
Const A379_I8_allc_rcpt_desc = 13

'[CONVERSION INFORMATION]  View Name : import_item a_rcpt_item
Const A379_I9_seq = 0
Const A379_I9_rcpt_type = 1
Const A379_I9_note_no = 2


LngMaxRow = CInt(Request("txtMaxRows"))											'☜: 최대 업데이트된 갯수 
LngMaxRow1 = CInt(Request("txtMaxRows1"))											'☜: 최대 업데이트된 갯수 
LngMaxRow3 = CInt(Request("txtMaxRows3"))											'☜: 최대 업데이트된 갯수 
lgIntFlgMode = CInt(Request("txtMode"))									'☜: 저장시 Create/Update 판별 

'-----------------------
'Data manipulate area
'-----------------------
ReDim I5_a_rcpt(A379_I5_rcpt_fg)
ReDim I4_b_acct_dept(A379_I4_dept_cd)
ReDim I8_a_allc_rcpt(A379_I8_allc_rcpt_desc)
ReDim I9_a_rcpt_item(A379_I9_note_no)

Dim I12_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 
Const A379_I12_a_data_auth_data_BizAreaCd = 0
Const A379_I12_a_data_auth_data_internal_cd = 1
Const A379_I12_a_data_auth_data_sub_internal_cd = 2
Const A379_I12_a_data_auth_data_auth_usr_id = 3 
 
  	Redim I12_a_data_auth(3)
	I12_a_data_auth(A379_I12_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I12_a_data_auth(A379_I12_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I12_a_data_auth(A379_I12_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I12_a_data_auth(A379_I12_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))

I2_a_acct_trans_type = "AR002"
I10_b_currency	= gCurrency

If lgIntFlgMode = OPMD_CMODE Then
	I5_a_rcpt(A379_I5_rcpt_no)			= Trim(Request("txtRcptNo"))
ElseIf lgIntFlgMode = OPMD_UMODE Then
	I5_a_rcpt(A379_I5_rcpt_no)			= Trim(Request("txtAllcNo"))
End If	
I5_a_rcpt(A379_I5_rcpt_dt)			= UNIConvDate(Request("txtRcptDt"))
I5_a_rcpt(A379_I5_doc_cur)			= UCase(TRim(Request("txtDocCur")))
I5_a_rcpt(A379_I5_xch_rate)			= UNIConvNum(Request("txtXchRate"),0)
I5_a_rcpt(A379_I5_insrt_user_id)	= ""
I5_a_rcpt(A379_I5_updt_user_id)		= ""
I5_a_rcpt(A379_I5_rcpt_fg)			= "D"
I5_a_rcpt(A379_I5_rcpt_amt)			= UNIConvNum(Request("txtRcptAmt"),0)
I5_a_rcpt(A379_I5_rcpt_loc_amt)		= 0
I11_a_rcpt_a_acct					= Trim(Request("txtAcctCd"))
I5_a_rcpt(A379_I5_rcpt_desc)		= Request("txtRcptDesc")
I9_a_rcpt_item(A379_I9_rcpt_type)	= Request("txtInputType")
I9_a_rcpt_item(A379_I9_note_no)		= Trim(Request("txtCheckCd"))
I6_b_bank							= Trim(Request("txtBankCD"))
I7_b_bank_acct						= Trim(Request("txtBankAcct"))
I4_b_acct_dept(A379_I4_org_change_id)	= UCase(Request("hOrgChangeId"))
I4_b_acct_dept(A379_I4_dept_cd)		= UCase(Trim(Request("txtDeptCd")))
I3_b_biz_partner					= UCase(Trim(Request("txtBpCd")))
I8_a_allc_rcpt(A379_I8_allc_dt)		= UNIConvDate(Request("txtRcptDt"))
I8_a_allc_rcpt(A379_I8_dc_amt)		= 0
I8_a_allc_rcpt(A379_I8_dc_loc_amt)	= 0
I8_a_allc_rcpt(A379_I8_allc_rcpt_desc) = Request("txtRcptDesc")

If lgIntFlgMode = OPMD_CMODE Then
	iCommandSent = "CREATE"
ElseIf lgIntFlgMode = OPMD_UMODE Then
	iCommandSent = "UPDATE"
End If

If Request("txtSpread") = "" Then
	Call DisplayMsgBox("112100", vbOKOnly, "", "", I_MKSCRIPT)
	Response.End
End If

If Request("txtSpread") <> "" Then

	importArray = Request("txtSpread")
	
	If Request("txtSpread1") <> "" Then
		importArray1 = Request("txtSpread1")
		
		If Request("txtSpread3") <> "" Then
			importArray3 = Request("txtSpread3")
		Else
			importArray3 = ""
		End If
	Else
		importArray1 = ""
		
		If Request("txtSpread3") <> "" Then
			importArray3 = Request("txtSpread3")
		Else
			importArray3 = ""
		End If
	End If	
	
	Set pAr0041d = Server.CreateObject("PARG025.cAMngDirectRcSvr")

	If CheckSYSTEMError(Err,True) = True Then
		Response.End																			'☜: 비지니스 로직 처리를 종료함 
	End If		

	E1_a_rcpt = pAr0041d.A_MANAGE_DIRECT_RCPT_SVR(gStrGlobalCollection,iCommandSent,I1_a_acct_trans_type,I2_a_acct_trans_type,I3_b_biz_partner, _ 
			I4_b_acct_dept,I5_a_rcpt,I6_b_bank,I7_b_bank_acct,I8_a_allc_rcpt,I9_a_rcpt_item,I10_b_currency,I11_a_rcpt_a_acct,Request("txtSpread"), _ 
			importArray1,importArray3,I12_a_data_auth)

	If CheckSYSTEMError(Err,True) = True Then
		Set pAr0041d = Nothing																	'☜: ComProxy Unload
		Response.End																			'☜: 비지니스 로직 처리를 종료함 
	End If		

End If

    Response.Write "<Script Language=VBScript> " & vbCr         
    Response.Write "With parent "				 & vbCr	
    If Not IsEmpty(E1_a_rcpt) And Trim(E1_a_rcpt) <> "" Then
		Response.Write " .frm1.txtRcptNo.value = """ & ConvSPChars(E1_a_rcpt) & """" & vbCr
		Response.Write " .frm1.txtAllcNo.value = """ & ConvSPChars(E1_a_rcpt) & """" & vbCr
	End If
	
	Set pAr0041d = Nothing                                                   '☜: Unload Comproxy
	
	Response.Write " .DBSaveOK " & vbCr
    Response.Write "End With "				 & vbCr	  
    Response.Write "</Script>"           	
	
%>
