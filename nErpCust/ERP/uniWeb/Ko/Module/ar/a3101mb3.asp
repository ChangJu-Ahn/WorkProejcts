<%
Option Explicit
'**********************************************************************************************
'*  1. Module Name          : Account
'*  2. Function Name        : Open Ap
'*  3. Program ID           : a3101mb3
'*  4. Program Name         : Open Ap �����ϴ� Logic
'*  5. Program Desc         :
'*  6. Complus List         : 
'*  7. Modified date(First) : 2000/04/10
'*  8. Modified date(Last)  : 2002/11/12
'*  9. Modifier (First)     : You So Eun
'* 10. Modifier (Last)      : Jeong Yong Kyun
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************




'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<%																			'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

Call HideStatusWnd														'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide

On Error Resume Next													'��: 

Call LoadBasisGlobalInf()

Dim iPARG005															'�� : ��ȸ�� ComProxy Dll ��� ���� 

Dim strMode																'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim iArrData
Dim iArrSpread
Dim iArrSpread3															'������� ���� �迭 ���� 
Dim iArrDept

' -- ���Ѱ����߰� 
Const A114_I11_a_data_auth_data_BizAreaCd = 0
Const A114_I11_a_data_auth_data_internal_cd = 1
Const A114_I11_a_data_auth_data_sub_internal_cd = 2
Const A114_I11_a_data_auth_data_auth_usr_id = 3

Dim I11_a_data_auth  '--> �Ķ������ ������ ���� ���̹� ���� 

Redim I11_a_data_auth(3)
I11_a_data_auth(A114_I11_a_data_auth_data_BizAreaCd)       = Trim(Request("lgAuthBizAreaCd"))
I11_a_data_auth(A114_I11_a_data_auth_data_internal_cd)     = Trim(Request("lgInternalCd"))
I11_a_data_auth(A114_I11_a_data_auth_data_sub_internal_cd) = Trim(Request("lgSubInternalCd"))
I11_a_data_auth(A114_I11_a_data_auth_data_auth_usr_id)     = Trim(Request("lgAuthUsrID"))

	strMode = Request("txtMode")											'�� : ���� ���¸� ���� 

	Redim iArrData(28)
	iArrData(0) = Trim(Request("txtArNo"))

	Redim iArrDept(1)

	If strMode = "" Then
		Response.End 
		Call HideStatusWnd		
	ElseIf strMode <> CStr(UID_M0003) Then									'�� : ��ȸ ���� Biz �̹Ƿ� �ٸ����� �׳� ������ 
		Response.End 
		Call HideStatusWnd		
	End If

	Set iPARG005 = Server.CreateObject("PARG005.cAMngOpenArSvr")

	If CheckSYSTEMError(Err,True) = True Then
		Response.End 
	End If

	Call iPARG005.A_MANAGE_OPEN_AR_SVR (gStrGlobalCollection, "DELETE", , , , , , , , , , iArrData,,, I11_a_data_auth)

	If CheckSYSTEMError(Err,True) = True Then
		Set iPARG005 = Nothing		
		Response.End 
	End If
	    
	Set iPARG005 = Nothing

	Response.Write "<Script Language=vbscript> " & vbcr
	Response.Write " parent.DbDeleteOk()       " & vbcr
	Response.Write "</Script>                  " & vbcr

%>
