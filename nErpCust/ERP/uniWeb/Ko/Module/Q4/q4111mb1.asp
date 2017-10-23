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
'*  3. Program ID           : Q4111MB1
'*  4. Program Name         : 검사결과조회 
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
		
Set objPQIG290 = Server.CreateObject("PQIG290.cQLoInspResultSimple")    

If CheckSYSTEMError(Err,True) = True Then
   Response.End
End if
		    
Call objPQIG290.Q_LOOK_UP_INSP_RESULT_SIMPLE_SVR(gStrGlobalCollection, _
											Request("PrevNextFlg"), _
											Request("txtPlantCd"), _
											Request("txtInspReqNo"), _
											1, _
											"", _
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
		
If E5_PrevNextError = "900011" Or E5_PrevNextError = "900012" Then
	Call DisplayMsgBox(E5_PrevNextError, vbOKOnly, "", "", I_MKSCRIPT)
End If

%>
<Script Language=vbscript>
With parent.frm1		
	'query condition
	.txtPlantNm.value = "<%=ConvSPChars(EG1_b_plant(E1_plant_nm))%>"
	.txtInspReqNo1.value = "<%=ConvSPChars(Trim(EG2_q_inspection_request(E2_insp_req_no)))%>"
	'content
	.txtInspReqNo2.value = "<%=ConvSPChars(EG2_q_inspection_request(E2_insp_req_no))%>"
	.txtInspClass.value = "<%=ConvSPChars(EG2_q_inspection_request(E2_insp_class_nm))%>"
	.txtItemCd.value = "<%=ConvSPChars(EG2_q_inspection_request(E2_item_cd))%>"
	.txtItemNm.value = "<%=ConvSPChars(EG2_q_inspection_request(E2_item_nm))%>"
	.txtSpec.value = "<%=ConvSPChars(EG2_q_inspection_request(E2_item_spec))%>"
	.txtLotNo.value = "<%=ConvSPChars(EG2_q_inspection_request(E2_lot_no))%>"
	If "<%=ConvSPChars(EG2_q_inspection_request(E2_lot_no))%>" <> "" Then
		.txtLotSubNo.Text = "<%=UniNumClientFormat(EG2_q_inspection_request(E2_lot_sub_no), 0 ,0)%>"
	End If
	.txtLotSize.Text = "<%=UniNumClientFormat(EG2_q_inspection_request(E2_lot_size), ggQty.DecPoint ,0)%>"
	.txtUnit.value = "<%=ConvSPChars(EG2_q_inspection_request(E2_unit))%>"
	.txtInspReqDt.Text = "<%=UniDateClientFormat(EG2_q_inspection_request(E2_insp_req_dt))%>"
			
	Select Case "<%=EG2_q_inspection_request(E2_insp_class_cd)%>"
		Case "R"
			.txtSupplierCd.value = "<%=ConvSPChars(EG2_q_inspection_request(E2_r_bp_cd))%>"
			.txtSupplierNm.value = "<%=ConvSPChars(EG2_q_inspection_request(E2_r_bp_nm))%>"
			.txtSLCd1.value = "<%=ConvSPChars(EG2_q_inspection_request(E2_r_sl_cd))%>"
			.txtSLNm1.value = "<%=ConvSPChars(EG2_q_inspection_request(E2_r_sl_nm))%>"	

			If "<%=EG3_q_inspection_result(E3_status_flag_cd)%>" = "R" Then
				.txtGoodsSLCd.value = "<%=ConvSPChars(EG3_q_inspection_result(E3_goods_sl_cd))%>"
				.txtGoodsSLNm.value = "<%=ConvSPChars(EG3_q_inspection_result(E3_goods_sl_nm))%>"
				.txtDefectivesSLCd.value = "<%=ConvSPChars(EG3_q_inspection_result(E3_defectives_sl_cd))%>"
				.txtDefectivesSLNm.value = "<%=ConvSPChars(EG3_q_inspection_result(E3_defectives_sl_nm))%>"
			Else
				.hGoodsSLCd.value = "<%=ConvSPChars(EG2_q_inspection_request(E2_r_sl_cd))%>"
				.hGoodsSLNm.value = "<%=ConvSPChars(EG2_q_inspection_request(E2_r_sl_nm))%>"
				.hDefectivesSLCd.value = "<%=ConvSPChars(EG2_q_inspection_request(E2_r_sl_cd))%>"
				.hDefectivesSLNm.value = "<%=ConvSPChars(EG2_q_inspection_request(E2_r_sl_nm))%>"
			End If					
		Case "P"
			.txtRoutNo.value = "<%=ConvSPChars(EG2_q_inspection_request(E2_p_rout_no))%>"
			.txtRoutNoDesc.value = "<%=ConvSPChars(EG2_q_inspection_request(E2_p_rout_no_desc))%>"
			.txtOprNo.value = "<%=ConvSPChars(EG2_q_inspection_request(E2_p_opr_no))%>"
			.txtOprNoDesc.value = "<%=ConvSPChars(EG2_q_inspection_request(E2_p_opr_no_desc))%>"
			.txtWcCd.value = "<%=ConvSPChars(EG2_q_inspection_request(E2_p_wc_cd))%>"
			.txtWcNm.value = "<%=ConvSPChars(EG2_q_inspection_request(E2_p_wc_nm))%>"

			If "<%=EG3_q_inspection_result(E3_status_flag_cd)%>" = "R" Then
				.txtGoodsSLCd.value = "<%=ConvSPChars(EG3_q_inspection_result(E3_goods_sl_cd))%>"
				.txtGoodsSLNm.value = "<%=ConvSPChars(EG3_q_inspection_result(E3_goods_sl_nm))%>"
				.txtDefectivesSLCd.value = "<%=ConvSPChars(EG3_q_inspection_result(E3_defectives_sl_cd))%>"
				.txtDefectivesSLNm.value = "<%=ConvSPChars(EG3_q_inspection_result(E3_defectives_sl_nm))%>"
			End If
		Case "F"
			.txtSLCd2.value = "<%=ConvSPChars(EG2_q_inspection_request(E2_f_sl_cd))%>"
			.txtSLNm2.value = "<%=ConvSPChars(EG2_q_inspection_request(E2_f_sl_nm))%>"

			If "<%=EG3_q_inspection_result(E3_status_flag_cd)%>" = "R" Then
				.txtGoodsSLCd.value = "<%=ConvSPChars(EG3_q_inspection_result(E3_goods_sl_cd))%>"
				.txtGoodsSLNm.value = "<%=ConvSPChars(EG3_q_inspection_result(E3_goods_sl_nm))%>"
				.txtDefectivesSLCd.value = "<%=ConvSPChars(EG3_q_inspection_result(E3_defectives_sl_cd))%>"
				.txtDefectivesSLNm.value = "<%=ConvSPChars(EG3_q_inspection_result(E3_defectives_sl_nm))%>"
			Else
				.hGoodsSLCd.value = "<%=ConvSPChars(EG2_q_inspection_request(E2_f_sl_cd))%>"
				.hGoodsSLNm.value = "<%=ConvSPChars(EG2_q_inspection_request(E2_f_sl_nm))%>"
				.hDefectivesSLCd.value = "<%=ConvSPChars(EG2_q_inspection_request(E2_f_sl_cd))%>"
				.hDefectivesSLNm.value = "<%=ConvSPChars(EG2_q_inspection_request(E2_f_sl_nm))%>"
			End If					
		Case "S"
			.txtBPCd.value = "<%=ConvSPChars(EG2_q_inspection_request(E2_s_bp_cd))%>"
			.txtBPNm.value = "<%=ConvSPChars(EG2_q_inspection_request(E2_s_bp_nm))%>"

			If "<%=EG3_q_inspection_result(E3_status_flag_cd)%>" = "R" Then
				.txtGoodsSLCd.value = "<%=ConvSPChars(EG3_q_inspection_result(E3_goods_sl_cd))%>"
				.txtGoodsSLNm.value = "<%=ConvSPChars(EG3_q_inspection_result(E3_goods_sl_nm))%>"
				.txtDefectivesSLCd.value = "<%=ConvSPChars(EG3_q_inspection_result(E3_defectives_sl_cd))%>"
				.txtDefectivesSLNm.value = "<%=ConvSPChars(EG3_q_inspection_result(E3_defectives_sl_nm))%>"
			End If					
	End Select
			
	.txtStatusFlag.value = "<%=ConvSPChars(EG3_q_inspection_result(E3_status_flag_nm))%>"
	.txtInspectorCd.value =	"<%=ConvSPChars(EG3_q_inspection_result(E3_inspector_cd))%>"
	.txtInspectorNm.value =	"<%=ConvSPChars(EG3_q_inspection_result(E3_inspector_nm))%>"
	.txtInspDt.Text = "<%=UniDateClientFormat(EG3_q_inspection_result(E3_insp_dt))%>"
	.txtInspQty.Text = "<%=UniNumClientFormat(EG3_q_inspection_result(E3_insp_qty), ggQty.DecPoint, 0)%>"
	.txtDefectQty.Text = "<%=UniNumClientFormat(EG3_q_inspection_result(E3_defect_qty), ggQty.DecPoint, 0)%>"
	.cboDecision.value = "<%=EG3_q_inspection_result(E3_decision_cd)%>"
	.txtDefectiveRate.Text = "<%=UniNumClientFormat(EG3_q_inspection_result(E3_defective_rate), 2,0)%>"
	.txtRemark.value = "<%=ConvSPChars(EG3_q_inspection_result(E3_remark))%>"
	
	If "<%=ConvSPChars(EG3_q_inspection_result(E3_status_flag_cd))%>" = "R" Then
		.txtGoodsQty.Text = "<%=UniNumClientFormat(EG3_q_inspection_result(E3_goods_qty), ggQty.DecPoint, 0)%>"
		.txtDefectivesQty.Text = "<%=UniNumClientFormat(EG3_q_inspection_result(E3_defectives_qty), ggQty.DecPoint, 0)%>"
		.txtReleaseDt.Text = "<%=UniDateClientFormat(EG3_q_inspection_result(E3_release_dt))%>"
	Else
		.hGoodsQty.value = "<%=UniNumClientFormat(EG3_q_inspection_result(E3_goods_qty), ggQty.DecPoint, 0)%>"
		.hDefectivesQty.value = "<%=UniNumClientFormat(EG3_q_inspection_result(E3_defectives_qty), ggQty.DecPoint, 0)%>"
		.hReleaseDt.value = "<%=UniDateClientFormat(EG3_q_inspection_result(E3_release_dt))%>"
	End If
		
End With
		
With parent
	.lgstatusflag = "<%=EG3_q_inspection_result(E3_status_flag_cd)%>"
	.lgInspClassCd = "<%=EG2_q_inspection_request(E2_insp_class_cd)%>"
	.lgReceivingInspType = "<%=EG4_q_configuration(E4_gr_insp_type)%>"
	.lgAutoPR = "<%=EG4_q_configuration(E4_pr_yn_before_receipt)%>"
	.lgAutoST = "<%=EG4_q_configuration(E4_st_yn_after_receipt)%>"
	.lgIFYesNo = "<%=EG2_q_inspection_request(E2_if_yesno)%>"
	
	Call .DbQueryOk()
	
End With
</Script>
