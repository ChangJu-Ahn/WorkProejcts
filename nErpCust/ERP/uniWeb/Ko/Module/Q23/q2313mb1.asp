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
'*  3. Program ID           : Q2313MB1
'*  4. Program Name         : 불량유형등록 
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
strinsp_class_cd = "F"	'@@@주의 

Dim PQIG290
Dim PQIG020	

Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngRow
Dim iLngExportUBound
Dim iTotalStr
Dim TmpBuffer
Dim strData

Dim lgStrPrevKey1
Dim lgStrPrevKey2          											'☆ : 조회용 ComProxy Dll 사용 변수 
Dim StrNextKey1
Dim StrNextKey2

Dim strInspReqNo
Dim strPlantCd
Dim strResultNo

Dim strHeaderQuery
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
	    
'DETAIL					
Const E5_insp_order = 0
Const E5_insp_item_cd = 1
Const E5_insp_item_nm = 2
Const E5_insp_series = 3
Const E5_defect_qty = 4
		
Dim EG1_b_plant
Dim EG2_q_inspection_request
Dim EG3_q_inspection_result
Dim EG4_q_configuration
Dim E5_PrevNextError

Dim I1_q_inspection_details
Dim E1_q_inspection_details

Redim I1_q_inspection_details(1) 

strInspReqNo = Request("txtInspReqNo")
strPlantCd = Request("txtPlantCd")

strHeaderQuery = "OK"

Set PQIG290 = Server.CreateObject("PQIG290.cQLoInspResultSimple")    

If CheckSYSTEMError(Err,True) = True Then
   Response.End
End if
		    
Call PQIG290.Q_LOOK_UP_INSP_RESULT_SIMPLE_SVR(gStrGlobalCollection, _
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
	Set PQIG290 = Nothing
	Response.End
End if
		
Set PQIG290 = Nothing
%>			
<Script Language=vbscript>
	With Parent.frm1
		.txtPlantCd.Value = "<%=ConvSPChars(Trim(EG1_b_plant(E1_plant_cd)))%>"
		.txtPlantNm.Value = "<%=ConvSPChars(Trim(EG1_b_plant(E1_plant_nm)))%>"
		.txtItemCd.Value = "<%=ConvSPChars(Trim(EG2_q_inspection_request(E2_item_cd)))%>"
		.txtItemNm.Value = "<%=ConvSPChars(Trim(EG2_q_inspection_request(E2_item_nm)))%>"
		.txtLotNo.Value = "<%=ConvSPChars(Trim(EG2_q_inspection_request(E2_lot_no)))%>"
		If "<%=Trim(EG2_q_inspection_request(E2_lot_no))%>" <> "" Then
			.txtLotSubNo.Value = "<%=EG2_q_inspection_request(E2_lot_sub_no)%>"
		End If
		.txtLotSize.Text = "<%=UniNumClientFormat(EG2_q_inspection_request(E2_lot_size), ggQty.DecPoint ,0)%>"
	End with
</Script>
<%
'Detail
strResultNo = 1
lgStrPrevKey1 = Request("lgStrPrevKey1")
lgStrPrevKey2 = Request("lgStrPrevKey2")

If lgStrPrevKey1 = ""  and lgStrPrevKey2 = "" then
	I1_q_inspection_details(0) = 1
	I1_q_inspection_details(1) = 1
Else
	I1_q_inspection_details(0) = lgStrPrevKey1
	I1_q_inspection_details(1) = lgStrPrevKey2
End If

Set PQIG020 = Server.CreateObject("PQIG020.cQListInspDetailSvr")
If CheckSystemError(Err,True) Then	
	Set PQIS020= Nothing						
	Response.End														'☜: 비지니스 로직 처리를 종료함 
End If

Call PQIG020.Q_LIST_INSP_DETAIL_FOR_DEFECT_TYPE_SVR(gStrGlobalCollection, _
												C_SHEETMAXROWS_D, _
												strInspReqNo, _
												strResultNo, _
												I1_q_inspection_details, _
												E1_q_inspection_details)
		
If CheckSystemError(Err,True) Then
	Set PQIG020 = Nothing		
	Response.End														'☜: 비지니스 로직 처리를 종료함 
End If

Set PQIG020 = Nothing

' 변경 부분: Next Key값과 실제 데이타(그룹뷰안)의 마지막 값이 같으면 다음 데이타가 없으므로 키 전달자 변수의 값을 초기화함 
' 문자/숫자 일 경우, 문맥에 맞게 처리함 

iLngExportUBound = UBound(E1_q_inspection_details, 1)

ReDim TmpBuffer(iLngExportUBound)	

StrNextKey1 = ""
StrNextKey2 = ""

For LngRow = 0 to iLngExportUBound
	If LngRow < C_SHEETMAXROWS_D Then
		strData = Chr(11) & ConvSPChars(E1_q_inspection_details(LngRow, E5_insp_item_cd)) _
				& Chr(11) & ConvSPChars(E1_q_inspection_details(LngRow, E5_insp_item_nm)) _
				& Chr(11) & E1_q_inspection_details(LngRow, E5_insp_series) _
				& Chr(11) & UniNumClientFormat(E1_q_inspection_details(LngRow, E5_defect_qty), ggQty.DecPoint ,0) _
				& Chr(11) & LngMaxRow + LngRow + 1 _
   				& Chr(11) & Chr(12)
    		
		TmpBuffer(LngRow) = strData
		
	Else
		StrNextKey1 = E1_q_inspection_details(LngRow, E5_insp_order)
		StrNextKey2 = E1_q_inspection_details(LngRow, E5_insp_series)
		
	End If
Next

iTotalStr = Join(TmpBuffer, "")

%>
<Script Language=vbscript>
	With Parent
		.ggoSpread.Source = .frm1.vspdData 
		.ggoSpread.SSShowDataByClip "<%=iTotalStr%>"
		
		.lgStrPrevKey1 = "<%=ConvSPChars(StrNextKey1)%>"
		.lgStrPrevKey2 = "<%=ConvSPChars(StrNextKey2)%>"
		If .frm1.vspdData.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData, 0) And .lgStrPrevKey1 <> "" And .lgStrPrevKey2 <> "" Then	
			.DbQuery
		Else		
			<% ' Request값을 hidden input으로 넘겨줌 %>
			.frm1.hInspReqNo.value = "<%=ConvSPChars(strInspReqNo)%>"
			.frm1.hPlantCd.value = "<%=ConvSPChars(strPlantCd)%>"
			.frm1.hInspItemCd.value = "<%=ConvSPChars(StrNextKey1)%>"
			.frm1.hInspSeries.value = "<%=StrNextKey2%>"
			.DbQueryOk
		End If		
	End with
</Script>
<%
Set PQIG020 = Nothing  
%>