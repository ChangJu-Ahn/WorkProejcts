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
'*  3. Program ID           : Q2116MB3
'*  4. Program Name         : 불합격통지 등록 
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

Const Q245_I1_insp_result_no = 0  
Const Q245_I1_plant_cd = 1
Const Q245_I1_insp_class_cd = 2
    
Dim PQIG130																	'☆ : 조회용 ComProxy Dll 사용 변수 
Dim strInspReqNo
Dim I2_q_inspection_result
Dim I3_q_reject_report
Dim iErrorPosition

Redim I2_q_inspection_result(2)

strInspReqNo = Request("txtInspReqNo")

I2_q_inspection_result(Q245_I1_insp_result_no) = 1
I2_q_inspection_result(Q245_I1_plant_cd) = Request("hPlantCd")
I2_q_inspection_result(Q245_I1_insp_class_cd) = "R"
		
Set PQIG130 = Server.CreateObject("PQIG130.cQMtRejRptSvr")

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If
	 
Call PQIG130.Q_MAINT_REJECT_REPORT_SVR(gStrGlobalCollection, _
								       strInspReqNo, _
									   I2_q_inspection_result, _
									   I3_q_reject_report, _
									   "D")
		
If CheckSYSTEMError(Err,True) = True Then
	Set PQIG130 = Nothing
	Response.End
End If			
			              
Set PQIG130 = Nothing   
                            
%>
<Script Language=vbscript>
	Call parent.DbDeleteOk()
</Script>
