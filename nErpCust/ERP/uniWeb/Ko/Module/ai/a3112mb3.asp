
<%
'**********************************************************************************************
'*  1. Module Name          : Account
'*  2. Function Name        : Open Ap
'*  3. Program ID           : a4101mb3
'*  4. Program Name         : ���� Open Ap �����ϴ� Logic
'*  5. Program Desc         :
'*  6. Complus List         : 
'*  7. Modified date(First) : 2000/04/10
'*  8. Modified date(Last)  : 2002/11/13
'*  9. Modifier (First)     : You So Eun
'* 10. Modifier (Last)      : Jeong Yong Kyun
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************

Response.Expires = -1															'�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
Response.Buffer = True															'�� : ASP�� ���ۿ� ����Ǿ� �������� �ٷ� Client�� ��������.

'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<%																				'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

	Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide

	On Error Resume Next														'��: 

	Call LoadBasisGlobalInf()

	Dim strMode																	'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

	Dim iPARG005																'��ȸ�� ComPlus Dll ��� ���� 
	Dim iArrData
	Dim ImportTypeTransType

	Const OpenArNo = 0

	strMode = Request("txtMode")												'�� : ���� ���¸� ���� 

	If strMode = "" Then
		Response.End 
		Call HideStatusWnd		
	ElseIf strMode <> CStr(UID_M0003) Then										'�� : ��ȸ ���� Biz �̹Ƿ� �ٸ����� �׳� ������ 
		Response.End
		Call HideStatusWnd		
	End If

	ImportTypeTransType = "AR005"

	Redim iarrdata(28)    
	iArrData(OpenArNo)  = Trim(Request("txtArNo"))

	Set iPARG005 = Server.CreateObject("PARG005.cAMngOpenArSvr")

	If CheckSYSTEMError(Err,True) = True Then
		Response.End 
	End If
	   
	Call iPARG005.A_MANAGE_OPEN_AR_SVR (gStrGlobalCollection, "DELETE", ImportTypeTransType, _ 
	                                           , , , , , , , , iArrData)
	If CheckSYSTEMError(Err,True) = True Then
		Set iPARG005 = Nothing		
		Response.End 
	End If
	    
	Set iPARG005 = Nothing

	Response.Write "<Script Language=vbscript>" & vbcr
	Response.Write " parent.dbDeleteOk        " & vbcr
	Response.Write "</Script>" & vbcr
%>

