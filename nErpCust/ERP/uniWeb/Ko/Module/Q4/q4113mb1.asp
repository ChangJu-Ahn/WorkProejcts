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
'*  3. Program ID           : Q4113MB1
'*  4. Program Name         : 수입검사불합격통지 
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
Call HideStatusWnd

Dim PQIG350													'☆ : 조회용 ComProxy Dll 사용 변수 

'EXPORTS VIEW
'PLANT
Const Q350_E1_plant_cd = 0
Const Q350_E1_plant_nm = 1

'INSPECTION RESULT
Const Q350_E2_insp_req_no = 0
Const Q350_E2_item_cd = 1
Const Q350_E2_item_nm = 2
Const Q350_E2_spec = 3
Const Q350_E2_lot_no = 4
Const Q350_E2_lot_sub_no = 5
Const Q350_E2_lot_size = 6
Const Q350_E2_insp_dt = 7
Const Q350_E2_decision_cd = 8
Const Q350_E2_decision_nm = 9
Const Q350_E2_bp_cd = 10
Const Q350_E2_bp_nm = 11

'REJECT REPORT
Const Q350_E3_frame_dt = 0
Const Q350_E3_framer = 1
Const Q350_E3_defect_comment = 2
Const Q350_E3_defect_contents = 3
Const Q350_E3_required_improvement = 4

Dim iExportPlant
Dim iStrUnitCd
Dim iExportInspectionResult
Dim iExportRejectReport
Dim iStrPreNextError
Set PQIG350 = Server.CreateObject("PQIG350.cQLoRejReportSimple")

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call PQIG350.Q_LOOK_UP_REJECT_REPORT_SIMPLE_SVR(gStrGlobalCollection, _
										Request("PrevNextFlg"), _
										Request("txtPlantCd"), _
										Request("txtInspReqNo"), _
										1, _
										iExportPlant, _
										iStrUnitCd, _
										iExportInspectionResult, _
										iExportRejectReport, _
										iStrPreNextError)

If CheckSYSTEMError(Err,True) = True Then
	Set PQIG350 = Nothing
	Response.End
End If

If iStrPreNextError = "900011" Or iStrPreNextError = "900012" Then
	Call DisplayMsgBox(iStrPreNextError, vbOKOnly, "", "", I_MKSCRIPT)
End If
	
Set PQIG350 = Nothing
%>
<Script Language=vbscript>
With Parent.frm1
	.txtPlantCd.Value = "<%=ConvSPChars(Trim(iExportPlant(Q350_E1_plant_cd)))%>"
	.txtPlantNm.Value = "<%=ConvSPChars(Trim(iExportPlant(Q350_E1_plant_nm)))%>"
	.hPlantCd.value = "<%=ConvSPChars(Trim(iExportPlant(Q350_E1_plant_cd)))%>"
	
	.txtInspReqNo1.Value = "<%=ConvSPChars(Trim(iExportInspectionResult(Q350_E2_insp_req_no)))%>"
	.txtInspReqNo2.Value = "<%=ConvSPChars(Trim(iExportInspectionResult(Q350_E2_insp_req_no)))%>"
	.txtBpCd.Value = "<%=ConvSPChars(Trim(iExportInspectionResult(Q350_E2_bp_cd)))%>"
	.txtBpNm.Value = "<%=ConvSPChars(Trim(iExportInspectionResult(Q350_E2_bp_nm)))%>"
	.txtItemCd.Value = "<%=ConvSPChars(Trim(iExportInspectionResult(Q350_E2_item_cd)))%>"
	.txtItemNm.Value = "<%=ConvSPChars(Trim(iExportInspectionResult(Q350_E2_item_nm)))%>"
	.txtSpec.Value = "<%=ConvSPChars(Trim(iExportInspectionResult(Q350_E2_spec)))%>"
	.txtLotNo.value = "<%=ConvSPChars(Trim(iExportInspectionResult(Q350_E2_lot_no)))%>"
	If "<%=ConvSPChars(Trim(iExportInspectionResult(Q350_E2_lot_no)))%>" <> "" Then
		.txtLotSubNo.value = "<%=UniNumClientFormat(iExportInspectionResult(Q350_E2_lot_sub_no), 0 ,0)%>"
	End If
	.txtLotSize.Text = "<%=UniNumClientFormat(iExportInspectionResult(Q350_E2_lot_size), ggQty.DecPoint ,0)%>"
	
	.txtUnit.value = "<%=ConvSPChars(Trim(iStrUnitCd))%>"
	.txtInspDt.Text = "<%=UNIDateClientFormat(iExportInspectionResult(Q350_E2_insp_dt))%>"
	.hDecision.Value = "<%=ConvSPChars(Trim(iExportInspectionResult(Q350_E2_decision_cd)))%>"		
	.txtDecision.Value = "<%=ConvSPChars(Trim(iExportInspectionResult(Q350_E2_decision_nm)))%>"	
	
	.txtFramer.Value = "<%=ConvSPChars(Trim(iExportRejectReport(Q350_E3_framer)))%>"
	.txtFrameDt.Text = "<%=UNIDateClientFormat(iExportRejectReport(Q350_E3_frame_dt))%>"
	.txtDefectComment.Value = "<%=ConvSPChars(Trim(iExportRejectReport(Q350_E3_defect_comment)))%>"
	.txtDefectContents.Value = "<%=ConvSPChars(Trim(iExportRejectReport(Q350_E3_defect_contents)))%>"
	.txtRequiredImprovement.Value = "<%=ConvSPChars(Trim(iExportRejectReport(Q350_E3_required_improvement)))%>"
End with
	
Call Parent.DbQueryOk()
</Script>	
