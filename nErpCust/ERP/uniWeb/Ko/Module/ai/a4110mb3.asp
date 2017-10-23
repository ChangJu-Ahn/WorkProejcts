<%
Option Explicit
'**********************************************************************************************
'*  1. Module Name          : Account
'*  2. Function Name        : Open Ap
'*  3. Program ID           : a4110mb3
'*  4. Program Name         : ���� Open Ap �����ϴ� Logic
'*  5. Program Desc         :
'*  6. Complus List         : 
'*  7. Modified date(First) : 2000/04/10
'*  8. Modified date(Last)  : 2002/11/14
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

	Dim iPAPG005																'��ȸ�� ComProxy Dll ��� ���� 
	Dim lgIntFlgMode
	Dim ImportData1 

	Dim strMode																	'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

	Const ApNo			= 0
	Const ApType        = 20
	Const ApLocAmt      = 35

	strMode = Request("txtMode")												'�� : ���� ���¸� ���� 
												
	Redim ImportData1(ApLocAmt)
	ImportData1(ApNo)            = Trim(Request("txtApNo"))
	ImportData1(ApType)          = "NT"

	If strMode = "" Then
		Response.End 
		Call HideStatusWnd		
	ElseIf strMode <> CStr(UID_M0003) Then										'�� : ��ȸ ���� Biz �̹Ƿ� �ٸ����� �׳� ������ 
		Response.End 
		Call HideStatusWnd		
	End If

	Set iPAPG005 = Server.CreateObject("PAPG005.cAMngOpenApSvr")

	If CheckSYSTEMError(Err,True) = True Then
		Response.End 
	End If

	Call iPAPG005.A_MANAGE_OPEN_AP_SVR (gStrGlobalCollection, "DELETE", , ImportData1)

	If CheckSYSTEMError(Err,True) = True Then
		Set iPAPG005 = Nothing		
		Response.End 
	End If
	    
	Set iPAPG005 = Nothing 

	Response.Write "<Script Language=vbscript>" & vbcr
	Response.Write " parent.dbDeleteOk" & vbcr
	Response.Write "</Script>" & vbcr

%>



	
