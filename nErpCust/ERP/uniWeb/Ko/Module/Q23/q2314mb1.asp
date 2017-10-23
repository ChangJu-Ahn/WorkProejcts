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
'*  3. Program ID           : Q2314MB1
'*  4. Program Name         : 불량원인등록 
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
strinsp_class_cd = "F"	'@@@주의 

Const C_SHEETMAXROWS_D = 100

Dim PQIG090		
Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngMaxRow2
Dim LngMaxRow3
Dim LngRow
Dim intGroupCount
Dim lgStrPrevKeyM         											'☆ : 조회용 ComProxy Dll 사용 변수 
Dim lgStrPrevKey1
Dim lgStrPrevKey2
Dim lgStrPrevKey3
Dim StrNextKey1
Dim StrNextKey2
Dim strNextKey3

Dim strHeaderQuery

Dim strInspReqNo
Dim strPlantCd
Dim strResultNo
Dim strInspSeries
Dim DefectTypeCd
Dim strData
Dim lglngHiddenRows
Dim lRow
Dim i

'HEADER
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

Dim import_q_inspection_result
Redim import_q_inspection_result(2)
Const Q234_I2_insp_result_no = 0
Const Q234_I2_plant_cd = 1
Const Q234_I2_insp_class_cd = 2

Redim E4_q_inspection_result(20)
Const Q212_E4_insp_result_no = 0
Const Q212_E4_insp_class_cd = 1
Const Q212_E4_insp_dt = 2
Const Q212_E4_insp_qty = 3
Const Q212_E4_defect_qty = 4
Const Q212_E4_decision = 5
Const Q212_E4_inspector_cd = 6
Const Q212_E4_rmk = 7
Const Q212_E4_bp_cd = 8
Const Q212_E4_wc_cd = 9
Const Q212_E4_item_cd = 10
Const Q212_E4_plant_cd = 11
Const Q212_E4_lot_no = 12
Const Q212_E4_lot_sub_no = 13
Const Q212_E4_lot_size = 14
Const Q212_E4_sl_cd = 15
Const Q212_E4_sl_cd_for_good = 16
Const Q212_E4_sl_cd_for_defect = 17
Const Q212_E4_status_flag = 18
Const Q212_E4_transfer_flag = 19
Const Q212_E4_release_dt = 20

Redim E5_b_item(2)
Const Q212_E5_item_nm = 0
Const Q212_E5_spec = 1
Const Q212_E5_basic_unit = 2    
    
Redim E6_b_plant(0)
Const Q212_E6_plant_nm = 0    

Redim E8_p_work_center(0)
Const Q212_E8_wc_nm = 0    

'DETAIL	
Redim I2_q_inspection_result(2)

Dim I3_q_inspection_details
Redim I3_q_inspection_details(1)
Const Q234_I3_insp_item_cd = 0
Const Q234_I3_insp_series = 1
    
Dim E1_q_inspection_result    
Dim E2_b_plant
Dim E3_b_item
Dim E4_p_work_center
Dim E5_b_biz_partner
Dim E6_b_minor
Dim E7_b_minor
Dim EG1_group_export
Dim EG2_group_export
Dim EG3_group_export
Dim EG4_group_export
Dim E8_q_inspection_defect_type
Dim E9_q_inspection_details
    
strInspReqNo	= Request("txtInspReqNo")
I3_q_inspection_details(Q234_I3_insp_item_cd)	= Request("txtInspItemCd")
I3_q_inspection_details(Q234_I3_insp_series)	= Request("txtInspSeries")

lgStrPrevKeyM	= Request("lgStrPrevKeyM")
lglngHiddenRows = CLng(Request("lglngHiddenRows"))
lRow = CLng(Request("lRow"))
LngMaxRow2 = Request("txtMaxRows2")
LngMaxRow3 = Request("txtMaxRows3")

LngMaxRow = Request("txtMaxRows")
strInspReqNo = Request("txtInspReqNo")
strPlantCd = Request("txtPlantCd")


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
	'Header
	'Inspection Request
	.txtPlantCd.Value = "<%=ConvSPChars(Trim(EG1_b_plant(E1_plant_cd)))%>"
	.txtPlantNm.Value = "<%=ConvSPChars(Trim(EG1_b_plant(E1_plant_nm)))%>"
	.txtItemCd.Value = "<%=ConvSPChars(Trim(EG2_q_inspection_request(E2_item_cd)))%>"
	.txtItemNm.Value = "<%=ConvSPChars(Trim(EG2_q_inspection_request(E2_item_nm)))%>"
	.txtLotNo.Value = "<%=ConvSPChars(Trim(EG2_q_inspection_request(E2_lot_no)))%>"
	If "<%=Trim(EG2_q_inspection_request(E2_lot_no))%>" <> "" Then
		.txtLotSubNo.Value = "<%=EG2_q_inspection_request(E2_lot_sub_no)%>"
	End IF
	.txtLotSize.Text = "<%=UniNumClientFormat(EG2_q_inspection_request(E2_lot_size), ggQty.DecPoint ,0)%>"
End with
</Script>
<%
'Detail
Set PQIG090 = Server.CreateObject("PQIG090.cQListInspDefTypeSvr")

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

import_q_inspection_result(Q234_I2_insp_result_no) = 1
import_q_inspection_result(Q234_I2_plant_cd) = strPlantCd
import_q_inspection_result(Q234_I2_insp_class_cd) = strinsp_class_cd

call PQIG090.Q_LIST_INSP_DEFECT_TYPE_SVR  (gStrGlobalCollection, C_SHEETMAXROWS_D, strInspReqNo, _
										import_q_inspection_result, I3_q_inspection_details , , DefectTypeCd , _
										E1_q_inspection_result, E2_b_plant, E3_b_item, E4_p_work_center, _
										E5_b_biz_partner, E6_b_minor, E7_b_minor, EG1_group_export, _
										EG2_group_export, EG3_group_export, EG4_group_export, _
										E8_q_inspection_defect_type, E9_q_inspection_details)
		
If CheckSYSTEMError(Err,True) = True Then
	Set PQIG090 = Nothing	
	Response.End														'☜: 비지니스 로직 처리를 종료함 
End If

Set PQIG090 = Nothing

	For i = 0 To UBound(E9_q_inspection_details,1)
		If i < C_SHEETMAXROWS_D Then
			strData = strData & Chr(11) & Trim(ConvSPChars(E9_q_inspection_details(i, 0)))
			strData = strData & Chr(11) & Trim(ConvSPChars(EG2_group_export(i, 0)))
			strData = strData & Chr(11) & Trim(ConvSPChars(E9_q_inspection_details(i, 1)))
			strData = strData & Chr(11) & Trim(ConvSPChars(EG3_group_export(i, 0)))
			strData = strData & Chr(11) & Trim(ConvSPChars(EG4_group_export(i, 0)))
			strData = strData & Chr(11) & Trim(ConvSPChars(EG3_group_export(i, 1)))
			strData = strData & Chr(11) & LngMaxRow + i + 1
       		strData = strData & Chr(11) & Chr(12)
		Else
		StrNextKey = E4_q_inspection_result(i, 0)
	End if	
Next
%>
<Script Language=vbscript>
With Parent
	.ggoSpread.Source = .frm1.vspdData 
	.ggoSpread.SSShowDataByClip "<%=strData%>"
		
	.lgStrPrevKey1 = "<%=ConvSPChars(StrNextKey1)%>"
	.lgStrPrevKey2 = "<%=ConvSPChars(StrNextKey2)%>"
	.lgStrPrevKey3 = "<%=ConvSPChars(StrNextKey3)%>"
	If .frm1.vspdData.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData, 0) And .lgStrPrevKey1 <> "" And .lgStrPrevKey2 <> "" And .lgStrPrevKey3 <> "" Then
		.DbQuery
	Else
		 <% ' Request값을 hidden input으로 넘겨줌 %>
		 .frm1.hInspReqNo.value = "<%=ConvSPChars(strInspReqNo)%>"
		 .frm1.hPlantCd.value = "<%=ConvSPChars(strPlantCd)%>"
		 .frm1.hInspItemCd.value = "<%=ConvSPChars(StrNextKey1)%>"
		 .frm1.hInspSeries.value = "<%=StrNextKey2%>"
		 .frm1.hDefectTypeCd.value = "<%=ConvSPChars(StrNextKey3)%>"
		 .DbQueryOk
    End If		
End with
</Script>
<%
Set PQIG090 = Nothing  
%>
