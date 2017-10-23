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
'*  3. Program ID           : Q4111MB2
'*  4. Program Name         : 검사결과 신규/수정 저장 
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

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

Const I1_plant_cd = 0
Const I1_insp_req_no = 1
Const I1_insp_result_no = 2
Const I1_inspector = 3
Const I1_insp_dt = 4
Const I1_insp_qty = 5
Const I1_defect_qty = 6
Const I1_decision = 7
Const I1_remark = 8
    
Dim objPQIG280																	
	
Dim lgIntFlgMode	
Dim sCommandSent
	
Dim I1_q_inspection_result
	
lgIntFlgMode = CInt(Request("txtFlgMode"))										'☜: 저장시 Create/Update 판별 
If lgIntFlgMode = OPMD_CMODE Then
	sCommandSent = "CREATE"
ElseIf lgIntFlgMode = OPMD_UMODE Then
	sCommandSent = "UPDATE"
End If
    
ReDim I1_q_inspection_result(8)
    
I1_q_inspection_result(I1_plant_cd) = Trim(Request("txtPlantCd"))
I1_q_inspection_result(I1_insp_req_no) = Trim(Request("txtInspReqNo2"))
I1_q_inspection_result(I1_insp_result_no) = 1
I1_q_inspection_result(I1_inspector) = Trim(Request("txtInspectorCd"))
I1_q_inspection_result(I1_insp_dt) = UniConvDate(Request("txtInspDt"))
I1_q_inspection_result(I1_insp_qty) = UNIConvNum(Request("txtInspQty"), 0)
I1_q_inspection_result(I1_defect_qty) = UNIConvNum(Request("txtDefectQty"), 0)
I1_q_inspection_result(I1_decision) = Request("cboDecision")
I1_q_inspection_result(I1_remark) = Trim(Request("txtRemark"))
	
Set objPQIG280 = Server.CreateObject("PQIG280.cQMtInspResultSimple")
	
If CheckSYSTEMError(Err,True) = True Then
   Response.End
End if
	
Call objPQIG280.Q_MT_INSP_RESULT_SIMPLE_SVR(gStrGlobalCollection, sCommandSent, I1_q_inspection_result)
	
If CheckSYSTEMError(Err,True) = true Then
   Set objPQIG280 = Nothing
   Response.End
   
End if
	
Set objPQIG280 = Nothing
%>
<Script Language=vbscript>
	Parent.DbSaveOk
</Script>
