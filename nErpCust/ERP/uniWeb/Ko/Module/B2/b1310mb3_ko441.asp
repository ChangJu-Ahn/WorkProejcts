<%@ LANGUAGE=VBSCript%>
<%Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->

<%
'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

Call LoadBasisGlobalInf() 
Call HideStatusWnd

On Error Resume Next

Dim pPB2SA05																	'�� : ��ȸ�� ComProxy Dll ��� ���� 
Dim strMode																	'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 


' Com+ Conv. ���� ���� 
    
Dim importArray
Dim pvCommandSent



' ÷�� ���� 
Const C_import_b_bank_bank_cd = 0

strMode = Request("txtMode")												'�� : ���� ���¸� ���� 

If Request("txtBankCd") = "" Then												'��: ��ȸ�� ���� ���� ���Դ��� üũ 
	Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)	'��ȸ ���ǰ��� ����ֽ��ϴ�!
	Response.End 
End If

Set pPB2SA05 = Server.CreateObject("PB2SA05_KO441.cBMngBankSvr")	    	    

'-----------------------
'Com action result check area(OS,internal)
'-----------------------
If Err.Number <> 0 Then
	Set pPB2SA05 = Nothing												'��: ComProxy Unload
	Call ServerMesgBox(Err.description, vbInformation, I_MKSCRIPT)						'��:
	Response.End															'��: �����Ͻ� ���� ó���� ������ 
End If
	
'-----------------------
'Data manipulate  area(import view match)
'-----------------------
ReDim importArray(C_import_b_bank_bank_cd)
importArray(C_import_b_bank_bank_cd) = Request("txtBankCd")
pvCommandSent = "DELETE"


Call pPB2SA05.B_MANAGE_BANK_SVR(gStrGlobalCollection, CStr(pvCommandSent), importArray)
'------------------------
'Com action result check area(OS,internal)
'-----------------------

If CheckSYSTEMError(Err,True) = True Then
	Set pPB2SA05 = Nothing
	Response.End 
End If

Set pPB2SA05 = Nothing                                                   '��: Unload Comproxy

%>
<Script Language=vbscript>
	Call parent.DbDeleteOk()
</Script>
<%
Response.End
%>
