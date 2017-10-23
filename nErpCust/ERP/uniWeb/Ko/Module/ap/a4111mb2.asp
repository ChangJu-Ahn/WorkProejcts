<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name          : Account
'*  2. Function Name        : 
'*  3. Program ID           : a4111mb2
'*  4. Program Name         : ä��/ä�� ��� ���� Logic
'*  5. Program Desc         :
'*  6. Complus List         : 
'*  7. Modified date(First) : 1999/09/10
'*  8. Modified date(Last)  : 1999/09/10
'*  9. Modifier (First)     : Mr  Kim
'* 10. Modifier (Last)      : Mrs Kim
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
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
<%																			'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide

On Error Resume Next														'��: 
Err.Clear 

Call LoadBasisGlobalInf()

Dim txtClearNo
Dim pAp0081																	'��ȸ�� ComProxy Dll ��� ���� 
Dim lgIntFlgMode
Dim LngMaxRow
Dim LngMaxRow1

' Com+ Conv. ���� ���� 
    
Dim iCommandSent 

Dim I1_a_acct_trans_type
Dim I2_b_acct_dept
Dim I3_b_currency
Dim I4_a_clear_ap_ar
Dim importArrayAr
Dim importArrayAp
Dim E3_b_auto_numbering

Const A360_I2_org_change_id = 0    
Const A360_I2_dept_cd = 1

Const A360_I4_clear_no = 0    
Const A360_I4_clear_dt = 1
Const A360_I4_ref_no = 2
Const A360_I4_clear_amt = 3
Const A360_I4_clear_loc_amt = 4
Const A360_I4_clear_desc = 5
Const A360_I4_internal_cd = 6
Const A360_I4_insrt_user_id = 7
Const A360_I4_insrt_dt = 8
Const A360_I4_updt_user_id = 9
Const A360_I4_updt_dt = 10
Const A360_I4_doc_cur = 11

	Dim I5_a_data_auth  '--> �Ķ������ ������ ���� ���̹� ���� 
	Const A360_I5_a_data_auth_data_BizAreaCd = 0
	Const A360_I5_a_data_auth_data_internal_cd = 1
	Const A360_I5_a_data_auth_data_sub_internal_cd = 2
	Const A360_I5_a_data_auth_data_auth_usr_id = 3 
 
  	Redim I5_a_data_auth(3)
	I5_a_data_auth(A360_I5_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I5_a_data_auth(A360_I5_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I5_a_data_auth(A360_I5_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I5_a_data_auth(A360_I5_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))

	LngMaxRow = CInt(Request("txtMaxRows"))										'��: �ִ� ������Ʈ�� ���� 
	LngMaxRow1 = CInt(Request("txtMaxRows1"))									'��: �ִ� ������Ʈ�� ���� 
	lgIntFlgMode = CInt(Request("txtFlgMode"))									'��: ����� Create/Update �Ǻ� 

	ReDIm I2_b_acct_dept(A360_I2_dept_cd)
	ReDIm I4_a_clear_ap_ar(A360_I4_doc_cur)

	If lgIntFlgMode = OPMD_CMODE Then
		iCommandSent = "CREATE"
	ElseIf lgIntFlgMode = OPMD_UMODE Then
		iCommandSent = "UPDATE"	
	End If

	'-----------------------
	'Data manipulate area
	'-----------------------												    'Single ����Ÿ ���� 
	I1_a_acct_trans_type					= "AP004"
	I3_b_currency							= gCurrency
	I2_b_acct_dept(A360_I2_org_change_id)	= Trim(Request("hOrgChangeId"))
	I2_b_acct_dept(A360_I2_dept_cd)			= Trim(Request("txtDeptCd"))
	I4_a_clear_ap_ar(A360_I4_clear_no)      = Trim(Request("txtClearNo"))
	txtClearNo								= Request("txtClearNo")
	I4_a_clear_ap_ar(A360_I4_clear_dt)		= UNIConvDate(Request("txtAllcDt"))
	I4_a_clear_ap_ar(A360_I4_clear_amt)		= UNIConvNum(Request("txtAllcAmt"),0)
	I4_a_clear_ap_ar(A360_I4_clear_loc_amt)	= UNIConvNum(Request("txtAllcLocAmt"),0)
	I4_a_clear_ap_ar(A360_I4_clear_desc)	= Trim(Request("txtDesc"))
	I4_a_clear_ap_ar(A360_I4_insrt_user_id)	= ""
	I4_a_clear_ap_ar(A360_I4_updt_user_id)	= ""
	I4_a_clear_ap_ar(A360_I4_doc_cur)		= Request("txtDocCur")

	If Request("txtSpread") <> "" Then
		importArrayAp = Request("txtSpread")
	Else
		importArrayAp = ""
		Call DisplayMsgBox("111100", vbOKOnly, "", "", I_MKSCRIPT)
		Response.End	
	End If

	If Request("txtSpread1") <> "" Then
		importArrayAr = Request("txtSpread1")
	Else
		importArrayAr = ""
		Call DisplayMsgBox("112100", vbOKOnly, "", "", I_MKSCRIPT)
		Response.End	
	End If

	Set pAp0081 = Server.CreateObject("PAPG055.cAMntClearApArSvr")

	If CheckSYSTEMError(Err,True) = True Then
		Response.End																			'��: �����Ͻ� ���� ó���� ������ 
	End If		

	E3_b_auto_numbering = pAp0081.A_MAINT_CLEAR_AP_AR_SVR(gStrGlobalCollection,iCommandSent,I1_a_acct_trans_type, _
		I2_b_acct_dept,I3_b_currency,I4_a_clear_ap_ar,importArrayAr,importArrayAp,I5_a_data_auth)
		
	If CheckSYSTEMError(Err,True) = True Then
		Set pAp0081 = Nothing																	'��: ComProxy Unload
		Response.End																			'��: �����Ͻ� ���� ó���� ������ 
	End If		

	Set pAp0081 = Nothing

    Response.Write "<Script Language=VBScript> " & vbCr         
    Response.Write "With parent "				 & vbCr	
    
	IF Trim(txtClearNo) = "" Then
		Response.Write " .DBSaveOk(""" & ConvSPChars(E3_b_auto_numbering)  & """)"	& vbCr
	Else
		Response.Write " .DBSaveOk(""" & txtClearNo  & """)"						& vbCr
	END IF		    
    
    Response.Write "End With "														& vbCr	  
    Response.Write "</Script>"                                                             

%>
