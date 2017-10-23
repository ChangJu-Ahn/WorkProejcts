<%
'**********************************************************************************************
'*  1. Module Name          : Account
'*  2. Function Name        : Open Ap
'*  3. Program ID           : a4101mb3
'*  4. Program Name         : Open Ap �����ϴ� Logic
'*  5. Program Desc         :
'*  6. Comproxy List        : +AP001M
'*  7. Modified date(First) : 2000/04/10
'*  8. Modified date(Last)  : 2000/04/10
'*  9. Modifier (First)     : You So Eun
'* 10. Modifier (Last)      : You So Eun
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
<%																			'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide

On Error Resume Next														'��: 

Call LoadBasisGlobalInf()

Dim pAr004m																	'�� : ��ȸ�� ComProxy Dll ��� ���� 

Dim strMode																	'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

Dim iCommandSent 

Dim I1_a_acct_trans_type
Dim I2_a_acct
Dim I3_a_allc_rcpt_assn
Dim importArray
Dim I4_b_acct_dept
Dim importArray1
Dim importArray2
Dim importArray3
Dim I5_a_allc_rcpt
Dim I6_b_currency
Dim I7_b_biz_partner

Const A366_I5_allc_no = 0    
Const A366_I5_allc_dt = 1
Const A366_I5_allc_type = 2
Const A366_I5_ref_no = 3
Const A366_I5_allc_amt = 4
Const A366_I5_allc_loc_amt = 5
Const A366_I5_dc_amt = 6
Const A366_I5_dc_loc_amt = 7
Const A366_I5_insrt_user_id = 8
Const A366_I5_updt_user_id = 9

	strMode = Request("txtMode")														'�� : ���� ���¸� ���� 

	If strMode = "" Then
		Response.End 
		Call HideStatusWnd		
	ElseIf strMode <> CStr(UID_M0003) Then												'�� : ��ȸ ���� Biz �̹Ƿ� �ٸ����� �׳� ������ 
		Response.End 
		Call HideStatusWnd		
	ElseIf Request("txtAllcNo") = "" Then												'��: ��ȸ�� ���� ���� ���Դ��� üũ 
		Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)						'��ȸ ���ǰ��� ����ֽ��ϴ�!
		Response.End
		Call HideStatusWnd		 
	End If

Dim I8_a_data_auth  '--> �Ķ������ ������ ���� ���̹� ���� 
Const A366_I8_a_data_auth_data_BizAreaCd = 0
Const A366_I8_a_data_auth_data_internal_cd = 1
Const A366_I8_a_data_auth_data_sub_internal_cd = 2
Const A366_I8_a_data_auth_data_auth_usr_id = 3 
 
  	Redim I8_a_data_auth(3)
	I8_a_data_auth(A366_I8_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I8_a_data_auth(A366_I8_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I8_a_data_auth(A366_I8_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I8_a_data_auth(A366_I8_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))

	ReDim I5_a_allc_rcpt(A366_I5_updt_user_id)

	iCommandSent = "DELETE"
	I1_a_acct_trans_type = "AR003"
	I5_a_allc_rcpt(A366_I5_allc_no) = Trim(Request("txtAllcNo"))

	importArray  = ""
	importArray1 = ""
	importArray2 = ""
	importArray3 = ""

	Set pAr004m = Server.CreateObject("PARG055.cAMntRcAllcSvr")

	If CheckSYSTEMError(Err,True) = True Then
		Response.End																	'��: �����Ͻ� ���� ó���� ������ 
	End If	
		
	Call pAr004m.A_MAINT_RCPT_ALLC_SVR(gStrGlobalCollection,iCommandSent,I1_a_acct_trans_type,,,importArray, ,importArray1,importArray2,importArray3,I5_a_allc_rcpt,,I7_b_biz_partner,I8_a_data_auth)	
		
	If CheckSYSTEMError(Err,True) = True Then
		Set pAr004m = Nothing															'��: ComProxy Unload
		Response.End																	'��: �����Ͻ� ���� ó���� ������ 
	End If
	
	Set pAr004m = Nothing																'��: Unload Comproxy	
	                                                
	Response.Write " <Script Language=vbscript> " & vbCr
   	Response.Write " Call parent.DbDeleteOk()   " & vbCr
	Response.Write " </Script>                  " & vbCr
%>
