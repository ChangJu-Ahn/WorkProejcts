<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029B("I", "*", "NOCOOKIE","MB") %>
<%Call LoadBasisGlobalInf%>
<%
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q2115MB1
'*  4. Program Name         : 부적합처리 등록 
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

Dim strinsp_class_cd
strinsp_class_cd = "R"	'@@@주의 

Dim PQIG120													'☆ : 조회용 ComProxy Dll 사용 변수 

Dim strMode													'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim StrNextKey		' 다음 값 
Dim lgStrPrevKey	' 이전 값 
Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngRow
Dim intGroupCount          
Dim strPlantCd
Dim strInspReqNo
Dim strResultNo
Dim strData

Dim strHeaderQuery

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
Dim I1_q_inspection_disposition_disposition_cd					
Dim strDispositionCd
Dim EG1_export_group 

lgStrPrevKey = Request("lgStrPrevKey")
LngMaxRow = Request("txtMaxRows")
strPlantCd = Trim(Request("txtPlantCd"))
strInspReqNo = Trim(Request("txtInspReqNo"))

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
		.txtBpCd.Value = "<%=ConvSPChars(Trim(EG2_q_inspection_request(E2_r_bp_cd)))%>"	
		.txtBpNm.Value = "<%=ConvSPChars(Trim(EG2_q_inspection_request(E2_r_bp_nm)))%>"	
		.txtLotNo.Value = "<%=ConvSPChars(Trim(EG2_q_inspection_request(E2_lot_no)))%>"
		If "<%=Trim(EG2_q_inspection_request(E2_lot_no))%>" <> "" Then
			.txtLotSubNo.Value = "<%=EG2_q_inspection_request(E2_lot_sub_no)%>"
		End IF
		.txtLotSize.Text = "<%=UniNumClientFormat(EG2_q_inspection_request(E2_lot_size), ggQty.DecPoint ,0)%>"
		
		'Inspection Result
		.txtInspQty.Text = "<%=UniNumClientFormat(EG3_q_inspection_result(E3_insp_qty), ggQty.DecPoint ,0)%>"
		.txtDefectQty.Text = "<%=UniNumClientFormat(EG3_q_inspection_result(E3_defect_qty), ggQty.DecPoint ,0)%>"
		.hDecisionCd.Value = "<%=ConvSPChars(EG3_q_inspection_result(E3_decision_cd))%>"
		.txtDecision.Value = "<%=ConvSPChars(EG3_q_inspection_result(E3_decision_nm))%>"
		.hStatusFlag.Value = "<%=ConvSPChars(EG3_q_inspection_result(E3_status_flag_cd))%>"		
	End with
</Script>
<%
'Detail
Set PQIG120 = Server.CreateObject("PQIG120.cQListInspDispSvr")
'-----------------------
'Com action result check area(OS,internal)
'-----------------------
If CheckSystemError(Err,True) Then											'☜: ComProxy Unload
	Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)						
	Response.End														'☜: 비지니스 로직 처리를 종료함 
End If

strResultNo = 1

If lgStrPrevKey = ""then
	I1_q_inspection_disposition_disposition_cd = ""
Else
	I1_q_inspection_disposition_disposition_cd = lgStrPrevKey
End If

Call PQIG120.Q_LIST_INSP_DISPOSIT_SVR(gStrGlobalCollection,C_SHEETMAXROWS_D, _
									strDispositionCd, _
									strResultNo, strInspReqNo, _
									EG1_export_group)
									
If CheckSystemError(Err,True) Then
	If strHeaderQuery ="OK" Then		
%>
<Script Language=vbscript>
	With Parent
		If  "<%=E4_q_inspection_result(18)%>"<> "" Then
			 .frm1.hInspReqNo.value = "<%=ConvSPChars(strInspReqNo)%>"
			 .frm1.hPlantCd.value = "<%=ConvSPChars(strPlantCd)%>"
		 	.DbQueryOk
		End If
	End With
</Script>
<%
	End If
	Set PQIG120 = Nothing
	Response.End														'☜: 비지니스 로직 처리를 종료함 
End If

Set PQIG120 = Nothing
    
For LngRow = 0 To UBound(EG1_export_group)
	If LngRow < C_SHEETMAXROWS_D Then 
		strData = strData & Chr(11) & ConvSPChars(EG1_export_group(LngRow,0))	'disposition_cd
		strData = strData & Chr(11) & ""
		strData = strData & Chr(11) & ConvSPChars(EG1_export_group(LngRow,3))	'disposition_nm
		strData = strData & Chr(11) & UniNumClientFormat(EG1_export_group(LngRow,1), ggQty.DecPoint ,0)				'qty
		strData = strData & Chr(11) & ConvSPChars(EG1_export_group(LngRow,2))	'remark
		strData = strData & Chr(11) & LngMaxRow + LngRow + 1
		strData = strData & Chr(11) & Chr(12)
	Else
		StrNextKey = ConvSPChars(EG1_export_group(LngRow,0))
	End If									
Next
%>
<Script Language=vbscript>
With Parent
	'.frm1.txtPlantNm.Value 	= "<%=ConvSPChars(pq21518.ExportBPlantPlantNm)%>"
    
	.ggoSpread.Source = .frm1.vspdData 
	.ggoSpread.SSShowDataByClip "<%=strData%>"
		
	.lgStrPrevKey = "<%=ConvSPChars(StrNextKey)%>"
	If .frm1.vspdData.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData, 0) And .lgStrPrevKey <> "" Then
		.DbQuery
	Else
		<% ' Request값을 hidden input으로 넘겨줌 %>
		.frm1.hInspReqNo.value = "<%=ConvSPChars(strInspReqNo)%>"
		.frm1.hPlantCd.value = "<%=ConvSPChars(strPlantCd)%>"
		.DbQueryOk
	End If
End with
</Script>