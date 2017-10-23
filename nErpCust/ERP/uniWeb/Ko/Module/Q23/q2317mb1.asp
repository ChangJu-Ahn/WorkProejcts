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
'*  3. Program ID           : Q2317MB1
'*  4. Program Name         : Release
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
strinsp_class_cd = "F"	''###그리드 컨버전 주의부분###

Dim strInspReqNo
Dim strPlantCd

'HEADER
'EXPORT VIEW
'B_PLANT
Const E1_plant_cd = 0
Const E1_plant_nm = 1
	    
'Q_INSPECTION_REQUEST
Const E2_insp_req_no = 0
Const E2_insp_class_cd = 1
Const E2_insp_class_nm = 2
Const E2_item_cd = 3
Const E2_item_nm = 4
Const E2_item_spec = 5
Const E2_lot_no = 6
Const E2_lot_sub_no = 7
Const E2_lot_size = 8
Const E2_unit = 9
Const E2_insp_req_dt = 10
	    
'수입검사 
Const E2_r_bp_cd = 11
Const E2_r_bp_nm = 12
	    
'공정검사 
Const E2_p_rout_no = 13
Const E2_p_rout_no_desc = 14
Const E2_p_opr_no = 15
Const E2_p_opr_no_desc = 16
Const E2_p_wc_cd = 17
Const E2_p_wc_nm = 18
	    
'최종검사 
Const E2_f_sl_cd = 19
Const E2_f_sl_nm = 20
	    
'출하검사 
Const E2_s_bp_cd = 21
Const E2_s_bp_nm = 22

'자체 검사 여부 
Const E2_if_yesno = 23

'수입검사 
Const E2_r_sl_cd = 24
Const E2_r_sl_nm = 25
	    
'Q_INSPECTION_RESULT
Const E3_insp_result_no = 0
Const E3_lot_size = 1
Const E3_inspector_cd = 2
Const E3_inspector_nm = 3
Const E3_insp_dt = 4
Const E3_insp_qty = 5
Const E3_defect_qty = 6
Const E3_decision_cd = 7
Const E3_decision_nm = 8
Const E3_defective_rate = 9
Const E3_remark = 10
Const E3_status_flag_cd = 11
Const E3_status_flag_nm = 12
Const E3_transfer_flag_cd = 13
	    
'Release 정보 
Const E3_goods_qty = 14
Const E3_defectives_qty = 15
Const E3_release_dt = 16
Const E3_goods_sl_cd = 17
Const E3_goods_sl_nm = 18
Const E3_defectives_sl_cd = 19
Const E3_defectives_sl_nm = 20
	    
'Q_Configure
'공급처의 검사유형(입고전/후)
Const E4_gr_insp_type = 0
	    
'품질환경설정의 자동 입고/재고이동 
Const E4_pr_yn_before_receipt = 1
Const E4_st_yn_after_receipt = 2
	    
Dim objPQIG290
		
Dim EG1_b_plant
Dim EG2_q_inspection_request
Dim EG3_q_inspection_result
Dim EG4_q_configuration
Dim E5_PrevNextError

strInspReqNo = Request("txtInspReqNo")
strPlantCd = Request("txtPlantCd")

Set objPQIG290 = Server.CreateObject("PQIG290.cQLoInspResultSimple")    

If CheckSYSTEMError(Err,True) = True Then
   Response.End
End if
		    
Call objPQIG290.Q_LOOK_UP_INSP_RESULT_SIMPLE_SVR(gStrGlobalCollection, _
											"", _
											strPlantCd, _
											strInspReqNo, _
											1, _
											strinsp_class_cd, _
											EG1_b_plant, _
											EG2_q_inspection_request,_
											EG3_q_inspection_result,_
											EG4_q_configuration, _
											E5_PrevNextError)
		
If CheckSYSTEMError(Err,True) = true Then

	Set objPQIG290 = Nothing
	Response.End
End if
		
Set objPQIG290 = Nothing
%>
<Script Language=vbscript>
With Parent.frm1
	'Inspection Request
	.txtPlantCd.Value = "<%=ConvSPChars(Trim(EG1_b_plant(E1_plant_cd)))%>"
	.txtPlantNm.Value = "<%=ConvSPChars(Trim(EG1_b_plant(E1_plant_nm)))%>"
	.txtItemCd.Value = "<%=ConvSPChars(Trim(EG2_q_inspection_request(E2_item_cd)))%>"
	.txtItemNm.Value = "<%=ConvSPChars(Trim(EG2_q_inspection_request(E2_item_nm)))%>"
	
	.txtInspReqDt.Text = "<%=UNIDateClientFormat(EG2_q_inspection_request(E2_insp_req_dt))%>"
	
	.txtLotNo.Value = "<%=ConvSPChars(Trim(EG2_q_inspection_request(E2_lot_no)))%>"
	If "<%=Trim(EG2_q_inspection_request(E2_lot_no))%>" <> "" Then
		.txtLotSubNo.Value = "<%=EG2_q_inspection_request(E2_lot_sub_no)%>"
	End IF
	.txtLotSize.Text = "<%=UniNumClientFormat(EG2_q_inspection_request(E2_lot_size), ggQty.DecPoint ,0)%>"
		
	'Inspection Result
	.txtDecision.Value = "<%=ConvSPChars(EG3_q_inspection_result(E3_decision_nm))%>"
	.hStatusFlag.Value = "<%=ConvSPChars(EG3_q_inspection_result(E3_status_flag_cd))%>"		
	.txtStatusFlag.Value = "<%=ConvSPChars(EG3_q_inspection_result(E3_status_flag_nm))%>"	
				
	.txtReleaseDt.Text = "<%=UNIDateClientFormat(EG3_q_inspection_result(E3_release_dt))%>"
	
	If "<%=ConvSPChars(EG3_q_inspection_result(E3_status_flag_cd))%>" = "R" Then
		.txtSlCdForGood.Value = "<%=ConvSPChars(Trim(EG3_q_inspection_result(E3_goods_sl_cd)))%>"
		.txtSlNmForGood.Value = "<%=ConvSPChars(Trim(EG3_q_inspection_result(E3_goods_sl_nm)))%>"
		.txtSlCdForDefective.Value = "<%=ConvSPChars(Trim(EG3_q_inspection_result(E3_defectives_sl_cd)))%>"
		.txtSlNmForDefective.Value = "<%=ConvSPChars(Trim(EG3_q_inspection_result(E3_defectives_sl_nm)))%>"
	Else
		If parent.UNICDbl("<%=UniNumClientFormat(EG3_q_inspection_result(E3_goods_qty), ggQty.DecPoint ,0)%>") > 0 Then
			.txtSlCdForGood.Value = "<%=ConvSPChars(Trim(EG2_q_inspection_request(E2_f_sl_cd)))%>"
			.txtSlNmForGood.Value = "<%=ConvSPChars(Trim(EG2_q_inspection_request(E2_f_sl_nm)))%>"
		End If
		If parent.UNICDbl("<%=UniNumClientFormat(EG3_q_inspection_result(E3_defectives_qty), ggQty.DecPoint ,0)%>") > 0 Then
			.txtSlCdForDefective.Value = "<%=ConvSPChars(Trim(EG2_q_inspection_request(E2_f_sl_cd)))%>"
			.txtSlNmForDefective.Value = "<%=ConvSPChars(Trim(EG2_q_inspection_request(E2_f_sl_nm)))%>"
		End If
	End If
	
	.txtInspectorCd.Value = "<%=ConvSPChars(Trim(EG3_q_inspection_result(E3_inspector_cd)))%>"
	.txtInspectorNm.Value = "<%=ConvSPChars(Trim(EG3_q_inspection_result(E3_inspector_nm)))%>"
	.txtInspDt.Text = "<%=UNIDateClientFormat(EG3_q_inspection_result(E3_insp_dt))%>"
	.txtGoodQty.Text = "<%=UniNumClientFormat(EG3_q_inspection_result(E3_goods_qty), ggQty.DecPoint ,0)%>"
	.txtDefectQty.Value = "<%=UniNumClientFormat(EG3_q_inspection_result(E3_defectives_qty), ggQty.DecPoint ,0)%>"
End with
'입고전/후 검사, 재고이동 자동처리, 구매입고 자동처리 
With parent
	.lgNextNo = ""		' 다음 키 값 넘겨줌 
	.lgPrevNo = ""		' 이전 키 값 넘겨줌 , 현재 ComProxy가 제대로 안되 있음		
	
	.DbQueryOk
End With
</Script>

