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
'*  3. Program ID           : Q2116MB2
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

'[CONVERSION INFORMATION]  IMPORTS View 상수 
Const Q245_I3_frame_dt = 0    '[CONVERSION INFORMATION]  View Name : import q_reject_report
Const Q245_I3_framer = 1
Const Q245_I3_defect_comment = 2
Const Q245_I3_defect_contents = 3
Const Q245_I3_required_improvement = 4

Dim PQIG130																	'☆ : 조회용 ComProxy Dll 사용 변수 
	
Dim lgIntFlgMode
Dim I2_q_inspection_result
Dim I3_q_reject_report
	
Dim I4_ief_supplied_select_char
Redim I2_q_inspection_result(2)
ReDim I3_q_reject_report(4)

I2_q_inspection_result(Q245_I1_insp_result_no) = 1
I2_q_inspection_result(Q245_I1_plant_cd) = Request("hPlantCd")
I2_q_inspection_result(Q245_I1_insp_class_cd) = "R"

lgIntFlgMode = CInt(Request("txtFlgMode"))			'☜: 저장시 Create/Update 판별 
	
If Len(Trim(Request("txtFrameDt"))) Then
	If UNIConvDate(Request("txtFrameDt")) = "" Then
		Call DisplayMsgBox("122116", vbinformation, "", "", I_MKSCRIPT)
		Response.End
	End If
End If
	
Set PQIG130 = Server.CreateObject("PQIG130.cQMtRejRptSvr")
	
'-----------------------
'Com action result check area(OS,internal)
'-----------------------

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

If lgIntFlgMode = OPMD_CMODE Then
	I4_ief_supplied_select_char = "C"
ElseIf lgIntFlgMode = OPMD_UMODE Then
	I4_ief_supplied_select_char = "U"
End If

I3_q_reject_report(Q245_I3_frame_dt) = UNIConvDate(Request("txtFrameDt"))
I3_q_reject_report(Q245_I3_framer) = Request("txtFramer")
I3_q_reject_report(Q245_I3_defect_comment) = Request("txtDefectComment")
I3_q_reject_report(Q245_I3_defect_contents) = Request("txtDefectContents")
I3_q_reject_report(Q245_I3_required_improvement) = Request("txtRequiredImprovement")

Dim strtxtInspReqNo2
strtxtInspReqNo2 = Request("txtInspReqNo2")

Call PQIG130.Q_MAINT_REJECT_REPORT_SVR(gStrGlobalCollection, _
								       strtxtInspReqNo2, _
									   I2_q_inspection_result, _
									   I3_q_reject_report, _
									   I4_ief_supplied_select_char)
	    
If CheckSYSTEMError(Err,True) = True Then
	Set PQIG130 = Nothing
	Response.End
End If		    
		              
Set PQIG130 = Nothing   
%>
<Script Language=vbscript>
With parent																		'☜: 화면 처리 ASP 를 지칭함 
	.DbSaveOk
End With
</Script>