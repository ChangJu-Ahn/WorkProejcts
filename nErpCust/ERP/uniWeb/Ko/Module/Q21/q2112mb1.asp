<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q2112MB1
'*  4. Program Name         : 내역등록 
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
Call loadInfTB19029B("I", "*", "NOCOOKIE","MB")
Call LoadBasisGlobalInf

On Error Resume Next
Call HideStatusWnd

Const C_SHEETMAXROWS_D = 100
 
Dim strinsp_class_cd
strinsp_class_cd = "R"	'@@@주의 

Dim PQIG020
Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngRow
Dim StrNextKey1
Dim StrNextKey2
Dim strInspReqNo
Dim strPlantCd
Dim strHeaderQuery
Dim strResultNo
Dim lgStrPrevKey1
Dim lgStrPrevKey2

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

'DETAIL
Dim EG1_export_group
Dim E1_q_inspection_details

ReDim EG1_export_group(22)

Const Q210_EG1_E1_q_inspection_details_insp_item_cd = 0
Const Q210_EG1_E1_q_inspection_details_insp_series = 1
Const Q210_EG1_E1_q_inspection_details_insp_class_cd = 2
Const Q210_EG1_E1_q_inspection_details_insp_method_cd = 3
Const Q210_EG1_E1_q_inspection_details_sample_qty = 4
Const Q210_EG1_E1_q_inspection_details_accpt_decision_qty = 5
Const Q210_EG1_E1_q_inspection_details_rejt_decision_qty = 6
Const Q210_EG1_E1_q_inspection_details_accpt_decision_discreate = 7
Const Q210_EG1_E1_q_inspection_details_max_defect_ratio = 8
Const Q210_EG1_E1_q_inspection_details_defect_qty = 9
Const Q210_EG1_E1_q_inspection_details_decesion = 10
Const Q210_EG1_E1_q_inspection_details_measmt_equipmt_cd = 11
Const Q210_EG1_E1_q_inspection_details_measmt_unit_cd = 12
Const Q210_EG1_E1_q_inspection_details_insp_order = 13
Const Q210_EG1_E1_q_inspection_details_insp_unit_indctn = 14
Const Q210_EG1_E2_q_inspection_item_insp_item_nm = 15
Const Q210_EG1_E3_q_inspection_standard_by_item_insp_spec = 16
Const Q210_EG1_E3_q_inspection_standard_by_item_usl = 17
Const Q210_EG1_E3_q_inspection_standard_by_item_lsl = 18
Const Q210_EG1_E4_InspMethodCd_b_minor_minor_nm = 19
Const Q210_EG1_E5_q_measurement_equipment_measmt_equipmt_nm = 20
Const Q210_EG1_E6_Decision_b_minor_minor_nm = 21
Const Q210_EG1_E7_InspUnitIndctn_b_minor_minor_nm = 22

LngMaxRow = Request("txtMaxRows")
strPlantCd = Request("txtPlantCd")
strInspReqNo = Request("txtInspReqNo")

strHeaderQuery = "OK"

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
	strHeaderQuery = "ERR"
	Set objPQIG290 = Nothing
	Response.End
End if
		
Set objPQIG290 = Nothing
%>
<Script Language=vbscript>
With Parent.frm1
	.txtPlantCd.Value = "<%=ConvSPChars(Trim(EG1_b_plant(E1_plant_cd)))%>"
	.txtPlantNm.Value = "<%=ConvSPChars(Trim(EG1_b_plant(E1_plant_nm)))%>"
	.txtItemCd.Value = "<%=ConvSPChars(Trim(EG2_q_inspection_request(E2_item_cd)))%>"
	.txtItemNm.Value = "<%=ConvSPChars(Trim(EG2_q_inspection_request(E2_item_nm)))%>"
	.txtBpCd.Value = "<%=ConvSPChars(Trim(EG2_q_inspection_request(E2_r_bp_cd)))%>"		'@@@주의(txtWcCd) AND E4_q_inspection_result(9)
	.txtBpNm.Value = "<%=ConvSPChars(Trim(EG2_q_inspection_request(E2_r_bp_nm)))%>"		'@@@주의(E8_p_work_center)
	.txtLotNo.Value = "<%=ConvSPChars(Trim(EG2_q_inspection_request(E2_lot_no)))%>"
	If "<%=Trim(EG2_q_inspection_request(E2_lot_no))%>" <> "" Then
		.txtLotSubNo.Value = "<%=EG2_q_inspection_request(E2_lot_sub_no)%>"
	End IF
	.txtLotSize.Text = "<%=UniNumClientFormat(EG2_q_inspection_request(E2_lot_size), ggQty.DecPoint ,0)%>"
End with			
</Script>
<%
'Detail
Set PQIG020 = Server.CreateObject("PQIG020.cQListInspDetailSvr")

If CheckSystemError(Err,True) Then
	Response.End														'☜: 비지니스 로직 처리를 종료함 
End If

strResultNo = 1
lgStrPrevKey1 = Request("lgStrPrevKey1")
lgStrPrevKey2 = Request("lgStrPrevKey2")

'Import
Dim I1_q_inspection_details
Redim I1_q_inspection_details(1)

If lgStrPrevKey1 = ""  and lgStrPrevKey2 = "" then
	I1_q_inspection_details(0) = 1
	I1_q_inspection_details(1) = 1
Else
	I1_q_inspection_details(0) = lgStrPrevKey1
	I1_q_inspection_details(1) = lgStrPrevKey2
End If

Call PQIG020.Q_LIST_INSP_DETAIL_SVR(gStrGlobalCollection, _
									C_SHEETMAXROWS_D, _
									I1_q_inspection_details, _
									strResultNo, _
									strInspReqNo, _
									EG1_export_group, _
									E1_q_inspection_details)

If CheckSystemError(Err,True) Then
	If strHeaderQuery ="OK" Then
%>
<Script Language=vbscript>
		With Parent
			.frm1.hInspReqNo.value = "<%=ConvSPChars(strInspReqNo)%>"
			.frm1.hPlantCd.value = "<%=ConvSPChars(strPlantCd)%>"
			.frm1.hInspItemCd.value = "<%=StrNextKey1%>"
			.frm1.hInspSeries.value = "<%=StrNextKey2%>"
			.DbQueryOk
		End With
</Script>
<%
	End If
	Set PQIG020 = Nothing		
	Response.End														'☜: 비지니스 로직 처리를 종료함 
End If

' 변경 부분: Next Key값과 실제 데이타(그룹뷰안)의 마지막 값이 같으면 다음 데이타가 없으므로 키 전달자 변수의 값을 초기화함 
' 문자/숫자 일 경우, 문맥에 맞게 처리함 
If EG1_export_group(UBound(EG1_export_group),Q210_EG1_E1_q_inspection_details_insp_item_cd) = E1_q_inspection_details(0) _
	And EG1_export_group(UBound(EG1_export_group),Q210_EG1_E1_q_inspection_details_insp_series) = E1_q_inspection_details(1) then
	StrNextKey1 = ""
	StrNextKey2 = ""
Else
	StrNextKey1 = I1_q_inspection_details(0)
	StrNextKey2 = I1_q_inspection_details(1)
End If
%>
<Script Language=vbscript>
Dim strData
With Parent
<%      
	For LngRow = 0 To UBound(EG1_export_group)
		If LngRow < C_SHEETMAXROWS_D Then
%>
			strData = strData & Chr(11) & "<%=EG1_export_group(LngRow,Q210_EG1_E1_q_inspection_details_Insp_Order)%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(EG1_export_group(LngRow,Q210_EG1_E1_q_inspection_details_insp_item_cd))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(EG1_export_group(LngRow,Q210_EG1_E2_q_inspection_item_insp_item_nm))%>"
			strData = strData & Chr(11) & "<%=EG1_export_group(LngRow,Q210_EG1_E7_InspUnitIndctn_b_minor_minor_nm)%>"
			strData = strData & Chr(11) & "<%=EG1_export_group(LngRow,Q210_EG1_E1_q_inspection_details_insp_series)%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(EG1_export_group(LngRow,Q210_EG1_E1_q_inspection_details_sample_qty), ggQty.DecPoint ,0)%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(EG1_export_group(LngRow,Q210_EG1_E1_q_inspection_details_accpt_decision_qty), ggQty.DecPoint ,0)%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(EG1_export_group(LngRow,Q210_EG1_E1_q_inspection_details_rejt_decision_qty), ggQty.DecPoint ,0)%>"
			If "<%=EG1_export_group(LngRow,Q210_EG1_E1_q_inspection_details_accpt_decision_discreate)%>" <> "" Then
				strData = strData & Chr(11) & "<%=UniNumClientFormat(EG1_export_group(LngRow,Q210_EG1_E1_q_inspection_details_accpt_decision_discreate), 4, 0)%>"
			Else
				strData = strData & Chr(11) & ""
			End If
			
			If "<%=EG1_export_group(LngRow,Q210_EG1_E1_q_inspection_details_max_defect_ratio)%>" <> "" Then
				strData = strData & Chr(11) & "<%=UniNumClientFormat(EG1_export_group(LngRow,Q210_EG1_E1_q_inspection_details_max_defect_ratio), 4, 0)%>"
			Else
				strData = strData & Chr(11) & ""
			End If			

			strData = strData & Chr(11) & "<%=ConvSPChars(EG1_export_group(LngRow,Q210_EG1_E3_q_inspection_standard_by_item_insp_spec))%>"

			If "<%=EG1_export_group(LngRow,Q210_EG1_E3_q_inspection_standard_by_item_lsl)%>" <> "" Then
				strData = strData & Chr(11) & "<%=UniNumClientFormat(EG1_export_group(LngRow,Q210_EG1_E3_q_inspection_standard_by_item_lsl), 4, 0)%>"
			Else
				strData = strData & Chr(11) & ""
			End If

			If "<%=EG1_export_group(LngRow,Q210_EG1_E3_q_inspection_standard_by_item_usl)%>" <> "" Then
				strData = strData & Chr(11) & "<%=UniNumClientFormat(EG1_export_group(LngRow,Q210_EG1_E3_q_inspection_standard_by_item_usl), 4, 0)%>"
			Else
				strData = strData & Chr(11) & ""
			End If

			strData = strData & Chr(11) & "<%=UniNumClientFormat(EG1_export_group(LngRow,Q210_EG1_E1_q_inspection_details_defect_qty), ggQty.DecPoint ,0)%>"
			strData = strData & Chr(11) & "<%=EG1_export_group(LngRow,Q210_EG1_E1_q_inspection_details_insp_unit_indctn)%>"
			strData = strData & Chr(11) & "<%=LngMaxRow + LngRow + 1%>" 
       		strData = strData & Chr(11) & Chr(12)
<%
		Else
			StrNextKey1 = ConvSPChars(EG1_export_group(LngRow,Q210_EG1_E1_q_inspection_details_insp_order))
			StrNextKey2 = ConvSPChars(EG1_export_group(LngRow,Q210_EG1_E1_q_inspection_details_insp_series))
		End If 
	Next
%>    
	.ggoSpread.Source = .frm1.vspdData 
	.ggoSpread.SSShowDataByClip strData		

	.lgStrPrevKey1 = "<%=ConvSPChars(StrNextKey1)%>"
	.lgStrPrevKey2 = "<%=ConvSPChars(StrNextKey2)%>"
	
	If .frm1.vspdData.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData, 0) And .lgStrPrevKey1 <> "" And .lgStrPrevKey2 <> "" Then
		.DbQuery
	Else
		 <% ' Request값을 hidden input으로 넘겨줌 %>
		 .frm1.hInspReqNo.value = "<%=ConvSPChars(strInspReqNo)%>"
		 .frm1.hPlantCd.value = "<%=ConvSPChars(strPlantCd)%>"
		 .frm1.hInspItemCd.value = "<%=ConvSPChars(StrNextKey1)%>"
		 .frm1.hInspSeries.value = "<%=ConvSPChars(StrNextKey2)%>"
		 .DbQueryOk
	End If		
End with
</Script>