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
'*  3. Program ID           : Q2314MB4
'*  4. Program Name         : 불량원인 전체 삭제 
'*  5. Program Desc         : Quality Configuration
'*  6. Component List       : 
'*  7. Modified date(First) : 2004/08/02
'*  8. Modified date(Last)  : 2004/08/02
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Koh Jae Woo
'* 11. Comment
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************

On Error Resume Next
Call HideStatusWnd

Dim strinsp_class_cd
strinsp_class_cd = "F"	'@@@주의 
Dim PQIG170																	'☆ : 조회용 ComProxy Dll 사용 변수 

Dim strInspReqNo
Dim strPlantCd


strInspReqNo = UCase(Trim(Request("txtInspReqNo")))
strPlantCd = UCase(Trim(Request("txtPlantCd")))
	
Dim I2_q_inspection_result
Redim I2_q_inspection_result(2)			
Const Q270_I2_insp_result_no = 0
Const Q270_I2_plant_cd = 1
Const Q270_I2_insp_class_cd = 2

I2_q_inspection_result(Q270_I2_insp_result_no) = 1	
I2_q_inspection_result(Q270_I2_plant_cd) = strPlantCd
I2_q_inspection_result(Q270_I2_insp_class_cd) = strinsp_class_cd	'@@@주의 

Set PQIG170 = Server.CreateObject("PQIG170.cQmtInspDefCauseSvr")

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call PQIG170.Q_MAINT_INSP_DEFT_CAUSE_SVR(gStrGlobalCollection, UCase(strInspReqNo), _
											I2_q_inspection_result, "Y")

If CheckSYSTEMError(Err,True) = True Then
	Set PQIG170 = Nothing	
	Response.End
End If

Set PQIG170 = Nothing	
%>
<Script Language=vbscript>
With parent	
	.DbDeleteOk															
End With
</Script>
