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
'*  3. Program ID           : Q2212MB4
'*  4. Program Name         : 내역 전체 삭제 
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
strinsp_class_cd = "P"	'@@@주의 
Dim PQIG060																	'☆ : 조회용 ComProxy Dll 사용 변수 

Dim strInspReqNo
Dim strPlantCd


strInspReqNo = UCase(Trim(Request("txtInspReqNo")))
strPlantCd = UCase(Trim(Request("txtPlantCd")))
	
Dim I2_q_inspection_result
ReDim I2_q_inspection_result(2)
Const Q221_I2_insp_result_no = 0    '[CONVERSION INFORMATION]  View Name : import q_inspection_result
Const Q221_I2_plant_cd = 1
Const Q221_I2_insp_class_cd = 2

I2_q_inspection_result(Q221_I2_insp_result_no) = 1	
I2_q_inspection_result(Q221_I2_plant_cd) = strPlantCd
I2_q_inspection_result(Q221_I2_insp_class_cd) = strinsp_class_cd	'@@@주의 

Set PQIG060 = Server.CreateObject("PQIG060.cQMtInspMeaValSvr")

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call PQIG060.Q_MAINT_INSP_MEAS_VALUE_SVR(gStrGlobalCollection, _
										 strInspReqNo, _
										 I2_q_inspection_result, _
										 "Y")

If CheckSYSTEMError(Err,True) = True Then
	Set PQIG060 = Nothing	
	Response.End
End If

Set PQIG060 = Nothing	
%>
<Script Language=vbscript>
With parent	
	.DbDeleteOk															
End With
</Script>