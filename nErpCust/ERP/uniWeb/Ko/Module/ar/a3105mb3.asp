<%
'**********************************************************************************************
'*  1. Module Name          : Account
'*  2. Function Name        : ä�ǰ��� 
'*  3. Program ID           : A3105mb3
'*  4. Program Name         : �Աݵ�Ϲ� ä�ǹ��� 
'*  5. Program Desc         : 
'*  6. Comproxy List        : +B21011 (Manager)
'                             +B21019 (��ȸ��)
'*  7. Modified date(First) : 2001/02/22
'*  8. Modified date(Last)  : 2001/02/22
'*  9. Modifier (First)     : Chang Sung Hee
'* 10. Modifier (Last)      : Chang Sung Hee
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            -2000/03/22 : ..........
'**********************************************************************************************




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
	
Dim pAr0041d																'�� : ��ȸ�� ComProxy Dll ��� ���� 
Dim strMode																	'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

' Com+ Conv. ���� ���� 
    
Dim importArray
Dim importArray1
Dim importArray3
Dim AcctTransTypeTransType
Dim I_a_rcpt 
Dim E1_a_rcpt
Dim iCommandSent

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

ReDim I_a_rcpt(A379_I5_rcpt_fg)

Dim I12_a_data_auth  '--> �Ķ������ ������ ���� ���̹� ���� 
Const A379_I12_a_data_auth_data_BizAreaCd = 0
Const A379_I12_a_data_auth_data_internal_cd = 1
Const A379_I12_a_data_auth_data_sub_internal_cd = 2
Const A379_I12_a_data_auth_data_auth_usr_id = 3 
 
  	Redim I12_a_data_auth(3)
	I12_a_data_auth(A379_I12_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I12_a_data_auth(A379_I12_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I12_a_data_auth(A379_I12_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I12_a_data_auth(A379_I12_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))


strMode = Request("txtMode")													'�� : ���� ���¸� ���� 

importArray = ""
importArray1 = ""
importArray3 = ""

If strMode = "" Then
	Response.End 
ElseIf strMode <> CStr(UID_M0003) Then											'�� : ��ȸ ���� Biz �̹Ƿ� �ٸ����� �׳� ������ 
	Response.End 
End If

If Request("txtRcptNo") = "" Then												'��: ��ȸ�� ���� ���� ���Դ��� üũ 
	Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)					'��ȸ ���ǰ��� ����ֽ��ϴ�!
	Response.End 
End If

I_a_rcpt(A379_I5_rcpt_no) = Trim(Request("txtRcptNo"))
AcctTransTypeTransType = "AR002"

iCommandSent = "DELETE"

Set pAr0041d = Server.CreateObject("PARG025.cAMngDirectRcSvr")

If CheckSYSTEMError(Err,True) = True Then
	Response.End																'��: �����Ͻ� ���� ó���� ������ 
End If

E1_a_rcpt = pAr0041d.A_MANAGE_DIRECT_RCPT_SVR(gStrGlobalCollection,iCommandSent,,AcctTransTypeTransType,,,I_a_rcpt,,,,,,,importArray,importArray1,importArray3,I12_a_data_auth)

If CheckSYSTEMError(Err,True) = True Then
	Set pAr0041d = Nothing														'��: ComProxy Unload
	Response.End																'��: �����Ͻ� ���� ó���� ������ 
End If

Set pAr0041d = Nothing															'��: Unload Comproxy

Response.Write " <Script Language=vbscript> " & vbCr
Response.Write " Call parent.DbDeleteOk()   " & vbCr
Response.Write " </Script>                  " & vbCr

%>
