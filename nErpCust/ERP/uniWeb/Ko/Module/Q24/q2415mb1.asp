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
'*  3. Program ID           : Q2415MB1
'*  4. Program Name         : 판정 
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
 
Const C_SHEETMAXROWS_D = 100
Dim PQIG160		

Dim strData
Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngRow
Dim intGroupCount          											
Dim lgintPrevKey1
Dim lgintPrevKey2
Dim intNextKey1
Dim intNextKey2
Dim strPlantCd
Dim strInspReqNo
Dim intInspResultNo
Dim strInspClassCd

'EXPORT
'HEADER
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
    
    '수입검사 
Const E2_r_sl_cd = 23
Const E2_r_sl_nm = 24
    
    'Q_INSPECTION_RESULT
Const E3_insp_result_no = 0
Const E3_inspector_cd = 1
Const E3_inspector_nm = 2
Const E3_insp_dt = 3
Const E3_insp_qty = 4
Const E3_defect_qty = 5
Const E3_decision_cd = 6
Const E3_decision_nm = 7
Const E3_remark = 8
Const E3_status_flag_cd = 9
Const E3_status_flag_nm = 10

'DETAILS    
'Q_INSPECTION_DETAILS
Const EG1_insp_order = 0
Const EG1_insp_item_cd = 1
Const EG1_insp_item_nm = 2
Const EG1_insp_series = 3
Const EG1_decision_cd = 4
Const EG1_decision_nm = 5
Const EG1_insp_method_cd = 6
Const EG1_insp_method_nm = 7
Const EG1_sample_qty = 8
Const EG1_defect_qty = 9
Const EG1_accpt_decision_qty = 10
Const EG1_rejt_decision_qty = 11
Const EG1_accpt_decision_discreate = 12
Const EG1_max_defect_ratio = 13

Dim E1_b_plant
Dim E2_q_inspection_request
Dim E3_q_inspection_result
Dim E4_q_defect_ratio
Dim EG1_q_inspection_details

LngMaxRow = CLng(Request("txtMaxRows"))            

strPlantCd = Request("txtPlantCd")
strInspReqNo = Request("txtInspReqNo")
intInspResultNo = 1
strInspClassCd = "S"

If Request("lgStrPrevKey1") = "" Then
	lgintPrevKey1 = 0
Else
	lgintPrevKey1 = UNIConvNum(Request("lgStrPrevKey1"), 0)
End If

If Request("lgStrPrevKey2") = "" Then
	lgintPrevKey2 = 0
Else
	lgintPrevKey2 = UNIConvNum(Request("lgStrPrevKey2"), 0)
End If

Set PQIG160 = Server.CreateObject("PQIG160.cQListDecisionSvr")
If CheckSystemError(Err,True) Then											'☜: ComProxy Unload
	Response.End														'☜: 비지니스 로직 처리를 종료함 
End If

Call PQIG160.Q_LIST_DECISION_SVR(gStrGlobalCollection, _
								C_SHEETMAXROWS_D, _
								strPlantCd, _
								strInspReqNo, _
								intInspResultNo, _
								strInspClassCd, _
								lgintPrevKey1, _
								lgintPrevKey2, _
								E1_b_plant, _
								E2_q_inspection_request, _
								E3_q_inspection_result, _
								E4_q_defect_ratio, _
								EG1_q_inspection_details)

If CheckSystemError(Err,True) Then											
	Set PQIG160= Nothing
	Response.End														'☜: 비지니스 로직 처리를 종료함 
End If

Set PQIG160= Nothing
%>
<Script Language=vbscript>
	With Parent.frm1
 		'Header
 		.txtPlantCd.Value = "<%=ConvSPChars(Trim(E1_b_plant(E1_plant_cd)))%>"
		.txtPlantNm.Value = "<%=ConvSPChars(Trim(E1_b_plant(E1_plant_nm)))%>"
		
		.txtInspReqNo.Value = "<%=ConvSPChars(strInspReqNo)%>"
		.txtItemCd.Value = "<%=ConvSPChars(Trim(E2_q_inspection_request(E2_item_cd)))%>"
		.txtItemNm.Value = "<%=ConvSPChars(Trim(E2_q_inspection_request(E2_item_nm)))%>"
		.txtBpCd.Value = "<%=ConvSPChars(Trim(E2_q_inspection_request(E2_s_bp_cd)))%>"
		.txtBpNm.Value = "<%=ConvSPChars(Trim(E2_q_inspection_request(E2_s_bp_nm)))%>"
		.txtLotNo.Value = "<%=Trim(E2_q_inspection_request(E2_lot_no))%>"
		.txtLotSubNo.Value = "<%=E2_q_inspection_request(E2_lot_sub_no)%>"
		.txtLotSize.Text = "<%=UniNumClientFormat(E2_q_inspection_request(E2_lot_size), ggQty.DecPoint ,0)%>"
		.txtInspReqDt.Text = "<%=UniDateClientFormat(E2_q_inspection_request(E2_insp_req_dt))%>"
		
		.txtInspectorCd.Value = "<%=ConvSPChars(Trim(E3_q_inspection_result(E3_inspector_cd)))%>"
		.txtInspectorNm.Value = "<%=ConvSPChars(Trim(E3_q_inspection_result(E3_inspector_nm)))%>"
		.txtInspDt.Text = "<%=UniDateClientFormat(E3_q_inspection_result(E3_insp_dt))%>"	
		
		.txtInspQty.Text = "<%=UniNumClientFormat(E3_q_inspection_result(E3_insp_qty), ggQty.DecPoint ,0)%>"
		.txtDefectQty.Text = "<%=UniNumClientFormat(E3_q_inspection_result(E3_defect_qty), ggQty.DecPoint ,0)%>"
		.cboDecision.Value = "<%=ConvSPChars(Trim(E3_q_inspection_result(E3_decision_cd)))%>"
		.txtDefectRatio.Value = "<%=UniNumClientFormat(E4_q_defect_ratio, 2, 0)%>"
		.txtDefectRatioUnit.Value = "%"
		.txtRemark.Value = "<%=ConvSPChars(Trim(E3_q_inspection_result(E3_remark)))%>" 
		.hStatusFlag.value = "<%=ConvSPChars(Trim(E3_q_inspection_result(E3_status_flag_cd)))%>"
	End with
	
</Script>
<%
	strData = ""
	For LngRow = 0 To UBound(EG1_q_inspection_details)
		If LngRow < C_SHEETMAXROWS_D Then
			strData = strData & Chr(11) & ConvSPChars(Trim(EG1_q_inspection_details(LngRow, EG1_decision_nm))) _	
							  & Chr(11) & EG1_q_inspection_details(LngRow, EG1_insp_order) _
							  & Chr(11) & ConvSPChars(Trim(EG1_q_inspection_details(LngRow, EG1_insp_item_cd))) _
							  & Chr(11) & ConvSPChars(Trim(EG1_q_inspection_details(LngRow, EG1_insp_item_nm))) _ 	
							  & Chr(11) & EG1_q_inspection_details(LngRow, EG1_insp_series) _
							  & Chr(11) & UniNumClientFormat(EG1_q_inspection_details(LngRow, EG1_sample_qty), ggQty.DecPoint ,0) _
							  & Chr(11) & UniNumClientFormat(EG1_q_inspection_details(LngRow, EG1_defect_qty), ggQty.DecPoint ,0) _
							  & Chr(11) & UniNumClientFormat(EG1_q_inspection_details(LngRow, EG1_accpt_decision_qty), ggQty.DecPoint ,0) _
							  & Chr(11) & UniNumClientFormat(EG1_q_inspection_details(LngRow, EG1_rejt_decision_qty), ggQty.DecPoint ,0)
							  
			 				  If EG1_q_inspection_details(LngRow, EG1_accpt_decision_discreate) = "" Or IsNull(EG1_q_inspection_details(LngRow, EG1_accpt_decision_discreate)) Then
								strData = strData & Chr(11) & ""
							  Else
								strData = strData & Chr(11) & UniNumClientFormat(EG1_q_inspection_details(LngRow, EG1_accpt_decision_discreate), 4 ,0)
							  End If

			 				  If EG1_q_inspection_details(LngRow, EG1_max_defect_ratio) = "" Or IsNull(EG1_q_inspection_details(LngRow, EG1_max_defect_ratio)) Then
								strData = strData & Chr(11) & ""
							  Else
								strData = strData & Chr(11) & UniNumClientFormat(EG1_q_inspection_details(LngRow, EG1_max_defect_ratio), 4 ,0)
							  End If							  
							  
			strData = strData & Chr(11) & ConvSPChars(Trim(EG1_q_inspection_details(LngRow, EG1_insp_method_cd))) _
							  & Chr(11) & ConvSPChars(Trim(EG1_q_inspection_details(LngRow, EG1_insp_method_nm))) _	
							  & Chr(11) & ConvSPChars(Trim(EG1_q_inspection_details(LngRow, EG1_decision_cd))) _	
							  & Chr(11) & (LngMaxRow + LngRow) _
							  & Chr(11) & Chr(12)
		Else
			intNextKey1 = EG1_q_inspection_details(LngRow, EG1_insp_order)
			intNextKey2 = EG1_q_inspection_details(LngRow, EG1_insp_series)
		End If
	Next
%>
<Script Language=vbscript>
	With Parent
		.ggoSpread.Source = .frm1.vspdData
		.ggoSpread.SSShowDataByClip "<%=strData%>"
		
		'/* 2003-03 정기패치: Next Key 처리 관련 - START */		
		.lgStrPrevKey1 = "<%=intNextKey1%>"
		.lgStrPrevKey2 = "<%=intNextKey2%>"
			 
		If .lgStrPrevKey1 <> "" And .lgStrPrevKey2 <> "" Then
			.DbQuery
		Else
			<% ' Request값을 hidden input으로 넘겨줌 %>
			.frm1.hInspReqNo.value = "<%=ConvSPChars(strInspReqNo)%>"
			.frm1.hPlantCd.value = "<%=ConvSPChars(strPlantCd)%>"
			.DbQueryOk
	    End If		
	    '/* 2003-03 정기패치: Next Key 처리 관련 - END */
	End with
</Script>


