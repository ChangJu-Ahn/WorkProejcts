<%
'**********************************************************************************************
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Organization(���κμ��ڵ����)
'*  3. Program ID           : B2405mb2.asp
'*  4. Program Name         : B2405mb2.asp
'*  5. Program Desc         :
'*  6. Comproxy List        : +B24052MakeInternalCd
'*  7. Modified date(First) : 2000/10/30
'*  8. Modified date(Last)  : 2002/12/03
'*  9. Modifier (First)     : Hwnag Jeong-won
'* 10. Modifier (Last)      : Sim Hae Young
'* 11. Comment              :
'**********************************************************************************************

%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%													                       '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call HideStatusWnd
On Error Resume Next														'��: 
Err.Clear   

Dim PB6G062																	'�� : �Է�/������ ComProxy Dll ��� ���� 
Dim strCase

Call LoadBasisGlobalInf()

strCase = Request("txtOrgId")

If strCase <> "" Then

    Set PB6G062 = Server.CreateObject("PB6G062.bBMakeInternalCd")
    call PB6G062.B_MAKE_INTERNAL_CD(gStrGlobalCollection,strCase)
    Set PB6G062 = Nothing

	If CheckSYSTEMError(Err,True) = True Then
	    ''Response.End 
	End If	
    On error goto 0
%>
<Script Language=vbscript>	
	With parent																		'��: ȭ�� ó�� ASP �� ��Ī�� 
		'window.status = "���� ����"
		.Batch_OK
	End With
</Script>
<%
End If 
    
%>
