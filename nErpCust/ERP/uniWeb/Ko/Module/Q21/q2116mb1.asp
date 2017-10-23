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
'*  3. Program ID           : Q2116MB1
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

Dim PQIG140													'☆ : 조회용 ComProxy Dll 사용 변수 
Dim strInspReqNo
Dim strPlantCd
Dim I1_q_inspection_result

strInspReqNo = Request("txtInspReqNo")
strPlantCd = Request("txtPlantCd")

ReDim I1_q_inspection_result(2)	
'[CONVERSION INFORMATION]  IMPORTS View 상수 
Const Q249_I1_insp_result_no = 0    '[CONVERSION INFORMATION]  View Name : import q_inspection_result
Const Q249_I1_plant_cd = 1
Const Q249_I1_insp_class_cd = 2

I1_q_inspection_result(Q249_I1_insp_result_no) = 1
I1_q_inspection_result(Q249_I1_insp_class_cd) = "R"
I1_q_inspection_result(Q249_I1_plant_cd) = strPlantCd

Dim E1_q_reject_report
Const Q249_E1_frame_dt = 0    '[CONVERSION INFORMATION]  View Name : export q_reject_report
Const Q249_E1_framer = 1
Const Q249_E1_defect_comment = 2
Const Q249_E1_defect_contents = 3
Const Q249_E1_required_improvement = 4

Dim E2_b_biz_partner
Const Q249_E2_bp_nm = 0    '[CONVERSION INFORMATION]  View Name : export b_biz_partner
    
Dim E3_b_minor
Const Q249_E3_minor_nm = 0    '[CONVERSION INFORMATION]  View Name : export_nm_for_inspector b_minor

Dim E4_b_minor
Const Q249_E4_minor_nm = 0    '[CONVERSION INFORMATION]  View Name : export_nm_for_decision b_minor

Dim E5_b_item
Const Q249_E5_item_nm = 0    '[CONVERSION INFORMATION]  View Name : export b_item
Const Q249_E5_spec = 1
Const Q249_E5_basic_unit = 2

Dim E6_b_plant
Const Q249_E6_plant_nm = 0    '[CONVERSION INFORMATION]  View Name : export b_plant
    
Dim E7_q_inspection_result
Const Q249_E7_insp_result_no = 0    '[CONVERSION INFORMATION]  View Name : export q_inspection_result
Const Q249_E7_insp_class_cd = 1
Const Q249_E7_insp_dt = 2
Const Q249_E7_insp_qty = 3
Const Q249_E7_defect_qty = 4
Const Q249_E7_decision = 5
Const Q249_E7_inspector_cd = 6
Const Q249_E7_rmk = 7
Const Q249_E7_bp_cd = 8
Const Q249_E7_wc_cd = 9
Const Q249_E7_item_cd = 10
Const Q249_E7_plant_cd = 11
Const Q249_E7_lot_no = 12
Const Q249_E7_lot_sub_no = 13
Const Q249_E7_lot_size = 14
Const Q249_E7_sl_cd = 15
Const Q249_E7_sl_cd_for_good = 16
Const Q249_E7_sl_cd_for_defect = 17
Const Q249_E7_status_flag = 18
Const Q249_E7_transfer_flag = 19
Const Q249_E7_release_dt = 20

Set PQIG140 = Server.CreateObject("PQIG140.cQLookupRejRptSvr")

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call PQIG140.Q_LOOK_UP_REJECT_REPORT_SVR(gStrGlobalCollection, _
										I1_q_inspection_result, _
										strInspReqNo, _
										E1_q_reject_report, _
										E2_b_biz_partner, _
										E3_b_minor, _
										E4_b_minor, _
										E5_b_item, _
										E6_b_plant, _
										E7_q_inspection_result)

If CheckSYSTEMError(Err,True) = True Then
	Set PQIG140 = Nothing
	Response.End
End If

Set PQIG140 = Nothing
%>
<Script Language=vbscript>
With Parent.frm1
	.hPlantCd.value = "<%=ConvSPChars(strPlantCd)%>"
	.txtPlantNm.Value = "<%=ConvSPChars(E6_b_plant(Q249_E6_plant_nm))%>"
	.txtInspReqNo2.Value = "<%=ConvSPChars(strInspReqNo)%>"
	.txtItemCd.Value = "<%=ConvSPChars(E7_q_inspection_result(Q249_E7_item_cd))%>"
	.txtItemNm.Value = "<%=ConvSPChars(E5_b_item(Q249_E5_item_nm))%>"
	.txtBpCd.Value = "<%=ConvSPChars(E7_q_inspection_result(Q249_E7_bp_cd))%>"
	.txtBpNm.Value = "<%=ConvSPChars(E2_b_biz_partner(Q249_E2_bp_nm))%>"
	.txtLotNo.Value	= "<%=ConvSPChars(E7_q_inspection_result(Q249_E7_lot_no))%>"
	.txtLotSubNo.Value = "<%=E7_q_inspection_result(Q249_E7_lot_sub_no)%>"
	.txtLotSize.Text = "<%=UniNumClientFormat(E7_q_inspection_result(Q249_E7_lot_size), ggQty.DecPoint ,0)%>"
		
	.txtInspDt.Text = "<%=UNIDateClientFormat(E7_q_inspection_result(Q249_E7_insp_dt))%>"
		
	.txtDecision.Value = "<%=ConvSPChars(E4_b_minor(Q249_E4_minor_nm))%>"		
	.txtFramer.Value = "<%=ConvSPChars(E1_q_reject_report(Q249_E1_framer))%>"
		
	.txtFrameDt.Text = "<%=UNIDateClientFormat(E1_q_reject_report(Q249_E1_frame_dt))%>"
		
	.txtDefectComment.Value = "<%=ConvSPChars(E1_q_reject_report(Q249_E1_defect_comment))%>"
	.txtDefectContents.Value = "<%=ConvSPChars(E1_q_reject_report(Q249_E1_defect_contents))%>"
	.txtRequiredImprovement.Value = "<%=ConvSPChars(E1_q_reject_report(Q249_E1_required_improvement))%>"
End with
	
parent.lgNextNo = ""		' 다음 키 값 넘겨줌 
parent.lgPrevNo = ""		' 이전 키 값 넘겨줌 , 현재 ComProxy가 제대로 안되 있음		
	
Parent.DbQueryOk
</Script>	
