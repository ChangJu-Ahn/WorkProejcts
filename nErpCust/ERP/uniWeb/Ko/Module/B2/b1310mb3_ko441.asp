<%@ LANGUAGE=VBSCript%>
<%Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->

<%
'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

Call LoadBasisGlobalInf() 
Call HideStatusWnd

On Error Resume Next

Dim pPB2SA05																	'☆ : 조회용 ComProxy Dll 사용 변수 
Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 


' Com+ Conv. 변수 선언 
    
Dim importArray
Dim pvCommandSent



' 첨자 선언 
Const C_import_b_bank_bank_cd = 0

strMode = Request("txtMode")												'☜ : 현재 상태를 받음 

If Request("txtBankCd") = "" Then												'⊙: 조회를 위한 값이 들어왔는지 체크 
	Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)	'조회 조건값이 비어있습니다!
	Response.End 
End If

Set pPB2SA05 = Server.CreateObject("PB2SA05_KO441.cBMngBankSvr")	    	    

'-----------------------
'Com action result check area(OS,internal)
'-----------------------
If Err.Number <> 0 Then
	Set pPB2SA05 = Nothing												'☜: ComProxy Unload
	Call ServerMesgBox(Err.description, vbInformation, I_MKSCRIPT)						'⊙:
	Response.End															'☜: 비지니스 로직 처리를 종료함 
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

Set pPB2SA05 = Nothing                                                   '☜: Unload Comproxy

%>
<Script Language=vbscript>
	Call parent.DbDeleteOk()
</Script>
<%
Response.End
%>
