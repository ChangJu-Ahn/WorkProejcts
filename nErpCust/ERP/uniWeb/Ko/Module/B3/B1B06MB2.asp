<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: B1B06MB2.asp
'*  4. Program Name			: Basic Data
'*  5. Program Desc			: Save B_ITEM_ACCT
'*  6. Comproxy List		: PB3S115.cBSetItemAcct
'*  7. Modified date(First)	: 2004/12/01
'*  8. Modified date(Last) 	: 2004/12/01
'*  9. Modifier (First)		: Chen, Jae Hyun
'* 10. Modifier (Last)		: Chen, Jae Hyun
'* 11. Comment		:
'**********************************************************************************************
'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%														'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call LoadBasisGlobalInf

Call HideStatusWnd

On Error Resume Next

Dim oPB3S115
Dim iErrorPosition																			'�� : �Է�/������ ComProxy Dll ��� ���� 
Dim itxtSpread

Err.Clear																		'��: Protect system from crashing

itxtSpread = ""

itxtSpread = Request("txtSpread")

Set oPB3S115 = Server.CreateObject("PB3S115.cBSetItemAcct")
    
'-----------------------
'Com action result check area(OS,internal)
'-----------------------
If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If
	
Call oPB3S115.B_MANAGE_ITEM_ACCT(gStrGlobalCollection, _
								itxtSpread, _
								iErrorPosition)

If CheckSYSTEMError(Err,True) = True Then
	Set oPB3S115 = Nothing
	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "Call parent.SheetFocus(" & iErrorPosition & ",1)" & vbCrLF
	Response.Write "</Script>" & vbCrLF
	Response.End
End If

If Not (oPB3S115 Is nothing) Then
	Set oPB3S115 = Nothing								'��: Unload Comproxy	
End If
        
Response.Write "<Script Language=VBScript>" & vbCrLF
Response.Write "Call parent.DbSaveOk" & vbCrLF
Response.Write "</Script>" & vbCrLF
Response.End
%>