<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<%
'**********************************************************************************************
'*  1. Module Name          : Paymment ���� ���� ó�� ASP
'*  2. Function Name        : 
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : +B19029LookupNumericFormat
'                            
'*  7. Modified date(First) : 1999/09/10
'*  8. Modified date(Last)  : 1999/09/10
'*  9. Modifier (First)     : Mr  Kim
'* 10. Modifier (Last)      : Mrs Kim
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            -1999/09/12 : ..........
'**********************************************************************************************

								'�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
								'�� : ASP�� ���ۿ� ����Ǿ� �������� �ٷ� Client�� ��������.


'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
' �Ʒ� �Լ��� �����Ͻ� ���� ���۵Ǵ� �������� ȣ���� �ּ���..
Call HideStatusWnd		
On Error Resume Next														'��: 

Call LoadBasisGlobalInf()

Dim pAr004m																	'�� : ��ȸ�� ComProxy Dll ��� ���� 
	
Dim IntRows
Dim IntCols
Dim vbIntRet
Dim lEndRow
Dim boolCheck
Dim lgIntFlgMode
Dim LngMaxRow

Dim strGroup
Dim arrCount

Dim iCommandSent 

Dim I1_a_acct_trans_type
Dim I2_a_acct
Dim I3_a_allc_rcpt_assn
Dim IG1_import_group_rcpt
Dim I4_b_acct_dept
Dim I5_a_allc_rcpt
Dim I6_b_currency
Dim I7_b_biz_partner
Dim importArray

    '[CONVERSION INFORMATION]  Group Name : import_group_rcpt
Const A358_IG1_I1_select_char = 0    '[CONVERSION INFORMATION]  View Name : import_rcpt ief_supplied
Const A358_IG1_I2_allc_dt = 1    '[CONVERSION INFORMATION]  View Name : import a_allc_rcpt_assn
Const A358_IG1_I2_doc_cur = 2
Const A358_IG1_I2_xch_rate = 3
Const A358_IG1_I2_allc_amt = 4
Const A358_IG1_I2_allc_loc_amt = 5
Const A358_IG1_I2_insrt_user_id = 6
Const A358_IG1_I2_updt_user_id = 7
Const A358_IG1_I3_rcpt_no = 8    '[CONVERSION INFORMATION]  View Name : import a_rcpt
Const A358_IG1_I3_rcpt_dt = 9
Const A358_IG1_I3_diff_kind_cur_amt = 10
Const A358_IG1_I3_diff_kind_cur_loc_amt = 11
Const A358_IG1_I3_diff_kind_cur = 12
Const A358_IG1_I3_insrt_dt = 13
Const A358_IG1_I3_updt_dt = 14

Const A358_I4_org_change_id = 0    '[CONVERSION INFORMATION]  View Name : import b_acct_dept
Const A358_I4_dept_cd = 1

Const A358_I5_allc_no = 0    '[CONVERSION INFORMATION]  View Name : import a_allc_rcpt
Const A358_I5_allc_dt = 1
Const A358_I5_allc_type = 2
Const A358_I5_ref_no = 3
Const A358_I5_allc_amt = 4
Const A358_I5_allc_loc_amt = 5
Const A358_I5_dc_amt = 6
Const A358_I5_dc_loc_amt = 7
Const A358_I5_allc_rcpt_desc = 8
Const A358_I5_insrt_user_id = 9
Const A358_I5_updt_user_id = 10

	ReDim IG1_import_group_rcpt(0,A358_IG1_I3_updt_dt)
	ReDim I4_b_acct_dept(A358_I4_dept_cd)
	ReDim I5_a_allc_rcpt(A358_I5_updt_user_id)

Dim I8_a_data_auth  '--> �Ķ������ ������ ���� ���̹� ���� 
Const A358_I8_a_data_auth_data_BizAreaCd = 0
Const A358_I8_a_data_auth_data_internal_cd = 1
Const A358_I8_a_data_auth_data_sub_internal_cd = 2
Const A358_I8_a_data_auth_data_auth_usr_id = 3 
 
  	Redim I8_a_data_auth(3)
	I8_a_data_auth(A358_I8_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I8_a_data_auth(A358_I8_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I8_a_data_auth(A358_I8_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I8_a_data_auth(A358_I8_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))


	LngMaxRow = CInt(Request("txtMaxRows"))											'��: �ִ� ������Ʈ�� ���� 
	lgIntFlgMode = CInt(Request("txtFlgMode"))										'��: ����� Create/Update �Ǻ� 

	'-----------------------
	'Data manipulate area
	'-----------------------														'��: Single ����Ÿ ���� 
	I6_b_currency										= gCurrency
	I1_a_acct_trans_type								= "AR006"
	I5_a_allc_rcpt(A358_I5_allc_no)						= Trim(Request("txtAllcNo"))
	I5_a_allc_rcpt(A358_I5_allc_dt)						= UNIConvDate(Request("txtAllcDt"))
	IG1_import_group_rcpt(0,A358_IG1_I2_allc_dt)		= UNIConvDate(Request("txtAllcDt"))
	I5_a_allc_rcpt(A358_I5_allc_type)					= "M"
	IG1_import_group_rcpt(0,A358_IG1_I3_rcpt_no)		= Trim(Request("txtRcptNo"))
	IG1_import_group_rcpt(0,A358_IG1_I3_rcpt_dt)		= UNIConvDate(Request("txtRcptDt"))
	I3_a_allc_rcpt_assn									= Request("txtDocCur")
	I5_a_allc_rcpt(A358_I5_allc_amt)					= UNIConvNum(Request("txtClsAmt"),0)
	I5_a_allc_rcpt(A358_I5_allc_loc_amt)				= UNIConvNum(Request("txtClsLocAmt"),0)
	I5_a_allc_rcpt(A358_I5_allc_rcpt_desc)				= Trim(Request("txtDesc"))
	I5_a_allc_rcpt(A358_I5_insrt_user_id)				= Request("txtUpdtUserId")
	I5_a_allc_rcpt(A358_I5_updt_user_id)				= Request("txtUpdtUserId")
	IG1_import_group_rcpt(0,A358_IG1_I2_insrt_user_id)	= Request("txtUpdtUserId")
	IG1_import_group_rcpt(0,A358_IG1_I2_updt_user_id)	= Request("txtUpdtUserId")
	I4_b_acct_dept(A358_I4_org_change_id)				= GetGlobalInf("gChangeOrgId")
	I4_b_acct_dept(A358_I4_dept_cd)						= Trim(Request("txtDeptCd"))
	I7_b_biz_partner									= Trim(Request("txtBpCd"))

	If lgIntFlgMode = OPMD_CMODE Then
		iCommandSent = "CREATE"
	ElseIf lgIntFlgMode = OPMD_UMODE Then
		iCommandSent = "UPDATE"
	End If

	Set pAr004m = Server.CreateObject("PARG080.cAMntAllcRcByApSvr")
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If CheckSYSTEMError(Err,True) = True Then
		Response.End																'��: �����Ͻ� ���� ó���� ������ 
	End If		

	If Trim(Request("txtRcptNo")) = "" Then
		Call DisplayMsgBox("112500", vbOKOnly, "", "", I_MKSCRIPT)
		Response.End
	End IF

	If Request("txtSpread") <> "" Then
		importArray = Request("txtSpread")
	Else
		importArray = ""
		Call DisplayMsgBox("111100", vbOKOnly, "", "", I_MKSCRIPT)
		Response.End	
	End If

	E2_b_auto_numbering = pAr004m.A_MAINT_ALLC_RCPT_BY_AP_SVR(gStrGlobalCollection,iCommandSent,I1_a_acct_trans_type, I2_a_acct, _
		I3_a_allc_rcpt_assn, IG1_import_group_rcpt, I4_b_acct_dept, I5_a_allc_rcpt, I6_b_currency, I7_b_biz_partner, importArray)

	If CheckSYSTEMError(Err,True) = True Then
		Set pAr004m = Nothing														'��: ComProxy Unload
		Response.End																'��: �����Ͻ� ���� ó���� ������ 
	End If		

	Set pAr004m = Nothing															'��: Unload Comproxy	

    Response.Write "<Script Language=VBScript> "									& vbCr         
    Response.Write "With parent "													& vbCr	

	IF Trim(E2_b_auto_numbering) <> "" Then
		Response.Write " .frm1.txtAllcNo.value = (""" & ConvSPChars(E2_b_auto_numbering) & """)"	& vbCr
	END IF		    
    
    Response.Write " .DBSaveOk(""" & ConvSPChars(E2_b_auto_numbering)  & """)"		& vbCr
    Response.Write "End With "														& vbCr	  
    Response.Write "</Script>"                                                             
%>
