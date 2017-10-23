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
'*  3. Program ID           : Q2512MB3
'*  4. Program Name         : 검사의뢰 삭제 
'*  5. Program Desc         : 
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

Call HideStatusWnd														'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

'공통 
Const I1_common_insp_req_no = 0
Const I1_common_plant_cd = 1
    
Dim objPQIG250																	
Dim I1_q_inspection_request
	
ReDim I1_q_inspection_request(1)
	
I1_q_inspection_request(I1_common_insp_req_no) = Trim(Request("txtInspReqNo"))
I1_q_inspection_request(I1_common_plant_cd) = Trim(Request("txtPlantCd"))
	
Set objPQIG250 = Server.CreateObject("PQIG250.cQMaintInspRequestSvr")

If CheckSYSTEMError(Err,True) = True Then
   Response.End
End if

Call objPQIG250.Q_MAINT_INSP_REQUEST_SVR(gStrGlobalCollection,  "DELETE", I1_q_inspection_request, "N")

If CheckSYSTEMError(Err,True) = true Then
   Set objPQIG250 = Nothing
   Response.End
End if

Set objPQIG250 = Nothing
%>
<Script Language=vbscript>
	Parent.DbDeleteOk
</Script>