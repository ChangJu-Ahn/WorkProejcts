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
'*  3. Program ID           : Q2211MB3
'*  4. Program Name         : 검사등록 
'*  5. Program Desc         : 검사결과 삭제 
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
	
Dim PQIG010																	'☆ : 조회용 ComProxy Dll 사용 변수 
Dim strInspReqNo
Dim strPlantCd
	
Dim iCommand
Dim I3_q_inspection_result
	
Dim IG1_import_group
	
strInspReqNo = Request("txtInspReqNo")	
strPlantCd = Request("txtPlantCd")	

Set PQIG010 = Server.CreateObject("PQIG010.cQMtInspResultSvr")
'-----------------------
'Com action result check area(OS,internal)
'-----------------------
If CheckSystemError(Err,True) Then
	Response.End														'☜: 비지니스 로직 처리를 종료함 
End If
	
Redim I3_q_inspection_result(3)
I3_q_inspection_result(0) = 1
I3_q_inspection_result(1) = UNIConvNum(Request("txtLotSize"), 0)
I3_q_inspection_result(2) = "P"
I3_q_inspection_result(3) = Trim(Request("txtPlantCd"))
iCommand = "D"
    	
Call PQIG010.Q_MAINT_INSP_RESULT_SVR(gStrGlobalCollection,iCommand,strInspReqNo,I3_q_inspection_result,IG1_import_group)
		
If CheckSYSTEMError(Err,True) Then
	Set PQIG010 = Nothing 
	Response.End
End if
	
Set PQIG010 = Nothing                                                  '☜: Unload Comproxy
%>
<Script Language=vbscript>
Call parent.DbDeleteOk()
</Script>