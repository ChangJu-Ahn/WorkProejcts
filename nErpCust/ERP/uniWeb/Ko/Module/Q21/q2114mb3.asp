<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029B("I", "*", "NOCOOKIE","MB") %>
<%Call LoadBasisGlobalInf%>
<%
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q2114MB3
'*  4. Program Name         : 판정 
'*  5. Program Desc         : Quality Configuration
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/05/15
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Park Hyun Soo
'* 11. Comment
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
							
On Error Resume Next
Call HideStatusWnd 

Dim strinsp_class_cd
strinsp_class_cd = "R"	'@@@주의 
	
Dim PQIG100																	'☆ : 조회용 ComProxy Dll 사용 변수 
Dim strInspReqNo
Dim IG1_import_Group
Dim iCommand

Dim I3_q_inspection_result
Redim I3_q_inspection_result(6)
Const Q236_I3_insp_result_no = 0	
	
Set PQIG100 = Server.CreateObject("PQIG100.cQMtDecisionSvr")

If CheckSystemError(Err,True) Then											'☜: ComProxy Unload
	Response.End														'☜: 비지니스 로직 처리를 종료함 
End If
		
icommand = "CANCEL"
	
strInspReqNo = Trim(Request("txtInspReqNo"))
I3_q_inspection_result(Q236_I3_insp_result_no) = 1
	
Call PQIG100.Q_MAINT_DECISION_SVR(gstrglobalcollection, _
	                                  icommand, strInspReqNo, _
	                                  I3_q_inspection_result, _
	                                  IG1_import_Group)


If CheckSYSTEMError(Err,True) Then
	Set PQIG100 = Nothing 
	Response.End
End If

Set PQIG100 = Nothing                                                   '☜: Unload Comproxy
%>
<Script Language=vbscript>
Call parent.DbDeleteOk()
</Script>