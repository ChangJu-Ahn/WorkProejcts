<%
'**********************************************************************************************
'*  1. Module Name          : �����ݹ��� 
'*  2. Function Name        : 
'*  3. Program ID           : a3108mb3.aps
'*  4. Program Name         :	
'*  5. Program Desc         :
'*  6. Comproxy List        : +Ar0041pr
'*  7. Modified date(First) : 2000/03/30
'*  8. Modified date(Last)  : 2000/06/17
'*  9. Modifier (First)     : YOU SO EUN
'* 10. Modifier (Last)      : Chang Sung Hee
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
' �Ʒ� �Լ��� �����Ͻ� ���� ���۵Ǵ� �������� ȣ���� �ּ���..
Call HideStatusWnd		
On Error Resume Next														'��: 

Call LoadBasisGlobalInf()

Dim pAr004m																	'�� : ��ȸ�� ComProxy Dll ��� ���� 

' Com+ Conv. ���� ���� 
    
Dim strGroup
Dim arrCount

Dim strMode																	'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

Dim iCommandSent
Dim I1_a_acct_trans_type
Dim I2_b_acct_dept
Dim I3_f_prrcpt
Dim I4_a_allc_rcpt
Dim I5_b_currency
Dim I6_b_biz_partner
Dim importArray
Dim importArray1
Dim importArray2
Dim E1_a_allc_rcpt

Const A365_I4_allc_no = 0    '[CONVERSION INFORMATION]  View Name : import a_allc_rcpt
Const A365_I4_allc_dt = 1
Const A365_I4_allc_type = 2
Const A365_I4_ref_no = 3
Const A365_I4_allc_amt = 4
Const A365_I4_allc_loc_amt = 5
Const A365_I4_dc_amt = 6
Const A365_I4_dc_loc_amt = 7
Const A365_I4_insrt_user_id = 8
Const A365_I4_updt_user_id = 9

	strMode = Request("txtMode")											'�� : ���� ���¸� ���� 

	If strMode = "" Then
		Response.End 
	ElseIf strMode <> CStr(UID_M0003) Then									'�� : ��ȸ ���� Biz �̹Ƿ� �ٸ����� �׳� ������ 
		Response.End 
	ElseIf Request("txtAllcNo") = "" Then									'��: ��ȸ�� ���� ���� ���Դ��� üũ 
		Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)			'��ȸ ���ǰ��� ����ֽ��ϴ�!
		Response.End 
	End If

Dim I7_a_data_auth  '--> �Ķ������ ������ ���� ���̹� ���� 
Const A365_I7_a_data_auth_data_BizAreaCd = 0
Const A365_I7_a_data_auth_data_internal_cd = 1
Const A365_I7_a_data_auth_data_sub_internal_cd = 2
Const A365_I7_a_data_auth_data_auth_usr_id = 3 
 
  	Redim I7_a_data_auth(3)
	I7_a_data_auth(A365_I7_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I7_a_data_auth(A365_I7_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I7_a_data_auth(A365_I7_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I7_a_data_auth(A365_I7_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))

	'-----------------------
	'Data manipulate  area(import view match)
	'-----------------------
	ReDim I4_a_allc_rcpt(A365_I4_updt_user_id)
	I4_a_allc_rcpt(A365_I4_allc_no) = Trim(Request("txtAllcNo"))
	I1_a_acct_trans_type	= "AR004"
	iCommandSent = "DELETE"

	importArray = ""
	importArray1 = ""
	importArray2 = ""
	'-----------------------
	'Com Action Area
	'-----------------------
	Set pAr004m = Server.CreateObject("PARG040.cAMntPrAllcSvr")
	
	If CheckSYSTEMError(Err,True) = True Then
		Response.End														'��: �����Ͻ� ���� ó���� ������ 
	End If	

	E1_a_allc_rcpt = pAr004m.A_MAINT_PRERCPT_ALLC_SVR(gStrGlobalCollection, iCommandSent, I1_a_acct_trans_type, , , I4_a_allc_rcpt, , , importArray, importArray1, importArray2,I7_a_data_auth)

	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If CheckSYSTEMError(Err,True) = True Then
		Set pAr004m = Nothing												'��: ComProxy Unload
		Response.End														'��: �����Ͻ� ���� ó���� ������ 
	End If

	Set pAr004m = Nothing				

	Response.Write " <Script Language=vbscript> " & vbCr
	Response.write " Call parent.DbDeleteOk()   " & vbCr
	Response.Write " </Script>                  " & vbCr
%>
