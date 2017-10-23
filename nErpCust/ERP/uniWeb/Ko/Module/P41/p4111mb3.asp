<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: p4111mb3.asp
'*  4. Program Name			: Release Order
'*  5. Program Desc			: Release By Production Order
'*  6. Comproxy List		: PP4G166.cPReleaseProdOrder
'*  7. Modified date(First) : 2000/03/21
'*  8. Modified date(Last)  : 2002/09/17
'*  9. Modifier (First)     : Park, BumSoo
'* 10. Modifier (Last)      : Chen, Jae Hyun
'* 11. Comment              :
'**********************************************************************************************
'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%																														'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call LoadBasisGlobalInf
Call HideStatusWnd

On Error Resume Next

Dim oPP4G166
Dim iErrorPosition																			'�� : �Է�/������ ComProxy Dll ��� ���� 
Dim strTxtSpread

Err.Clear																		'��: Protect system from crashing

strTxtSpread = Request("txtSpread")

Set oPP4G166 = Server.CreateObject("PP4G166.cPReleaseProdOrder")
    
'-----------------------
'Com action result check area(OS,internal)
'-----------------------
If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If
	
Call oPP4G166.P_RELEASE_BY_ORDER_SVR(gStrGlobalCollection, _
									Request("txtSpread"), _
									 , _
									 , _
									iErrorPosition)

If CheckSYSTEMError(Err,True) = True Then
	Set oPP4G166 = Nothing
	Response.End
End If

Set oPP4G166 = Nothing								'��: Unload Comproxy	

        
Response.Write "<Script Language=VBScript>" & vbCrLF
Response.Write "parent.DbSaveOk(True)" & vbCrLF
Response.Write "</Script>" & vbCrLF
Response.End
%>