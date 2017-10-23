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
'*  3. Program ID           : Q1211MB1
'*  4. Program Name         : 품목별 검사기준 등록 
'*  5. Program Desc         : 
'*  6. Component List       : PQBG120
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

Dim PQBG120													'☆ : 조회용 ComProxy Dll 사용 변수 
Dim strMode													'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim StrNextKey		' 다음 값 
Dim lgStrPrevKey	' 이전 값 
Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngRow
Dim intGroupCount          
Dim strPlantCd
Dim strItemCd
Dim strInspClassCd
Dim strRoutNo
Dim strOprNo
Dim inspection_item
Dim E1_q_inspection_item
	
Dim E3_b_plant
Const E1_plant_cd = 0
Const E1_plant_nm = 1    

Dim E2_b_item
Const E2_item_cd = 0
Const E2_item_nm = 1

Dim E4_p_routing_header
Const E4_rout_no = 0
Const E4_rout_no_desc = 1

Dim E5_p_routing_detail
Const E5_opr_no = 0
Const E5_opr_no_desc = 1

Dim EG1_inspection_standard_by_item
Const EG1_E1_insp_item_cd = 0
Const EG1_E1_insp_item_nm = 1
Const EG1_E1_insp_char = 2
Const EG1_E1_insp_char_nm = 3
Const EG1_E1_insp_class_cd = 4
Const EG1_E1_insp_method_cd = 5
Const EG1_E1_insp_method_nm = 6
Const EG1_E1_weight_cd = 7
Const EG1_E1_weight_nm = 8
Const EG1_E1_insp_spec = 9
Const EG1_E1_usl = 10
Const EG1_E1_lsl = 11
Const EG1_E1_measmt_unit_cd = 12
Const EG1_E1_ucl = 13
Const EG1_E1_lcl = 14
Const EG1_E1_mthd_of_cl_cal_cd = 15
Const EG1_E1_mthd_of_cl_cal_nm = 16
Const EG1_E1_calculated_qty = 17
Const EG1_E1_insp_order = 18
Const EG1_E1_insp_unit_indctn_cd = 19
Const EG1_E1_insp_unit_indctn_nm = 20
Const EG1_E1_measmt_equipmt_cd = 21
Const EG1_E1_measmt_equipmt_nm = 22
Const EG1_E1_insp_process_desc = 23
Const EG1_E1_remark = 24

	
Dim strData
Dim i

Const C_SHEETMAXROWS_D = 100
		
	lgStrPrevKey	= Request("lgStrPrevKey")
	LngMaxRow		= Request("txtMaxRows")
	strPlantCd		= Request("txtplantCd")
	strItemCd		= Request("txtItemCd")
	strRoutNo		= Request("txtRoutNo")
	strOprNo		= Request("txtOprNo")
	
	strInspClassCd = Request("cboInspClassCd")

	If lgStrPrevKey <> "" Then
		inspection_item = lgStrPrevKey
	End If

	Set PQBG120 = Server.CreateObject("PQBG120.cQListInspStdItemSvr")

	If CheckSYSTEMError(Err,True) = True Then
		Response.End
	End If

	CALL PQBG120.Q_LIST_INSP_STD_BY_ITEM_SVR (gStrGlobalCollection, _
											  C_SHEETMAXROWS_D, _
											  inspection_item, _
											  strItemCd, _
											  strInspClassCd, _
											  strPlantCd, _
											  strRoutNo, _
											  strOprNo, _
											  E1_q_inspection_item, _
											  E2_b_item, _
											  E3_b_plant, _
											  E4_p_routing_header, _
											  E5_p_routing_detail, _
											  EG1_inspection_standard_by_item)

	If CheckSYSTEMError(Err,True) = True Then
		Set PQBG120 = Nothing
		Response.End
	End If
		
	Set PQBG120 = Nothing

	Dim TmpBuffer
	Dim iTotalStr
	ReDim TmpBuffer(UBound(EG1_inspection_standard_by_item, 1))

	For LngRow = 0 To UBound(EG1_inspection_standard_by_item, 1)
		If LngRow < C_SHEETMAXROWS_D Then	
			strData = Chr(11) & Trim(ConvSPChars(EG1_inspection_standard_by_item(LngRow, EG1_E1_insp_item_cd))) _
					& Chr(11) & "" _											
					& Chr(11) & Trim(ConvSPChars(EG1_inspection_standard_by_item(LngRow, EG1_E1_insp_item_nm))) _
					& Chr(11) & Trim(ConvSPChars(EG1_inspection_standard_by_item(LngRow, EG1_E1_insp_char_nm))) _
					& Chr(11) & Trim(ConvSPChars(EG1_inspection_standard_by_item(LngRow, EG1_E1_insp_order))) _
					& Chr(11) & Trim(ConvSPChars(EG1_inspection_standard_by_item(LngRow, EG1_E1_insp_method_cd))) _
					& Chr(11) & "" _
					& Chr(11) & Trim(ConvSPChars(EG1_inspection_standard_by_item(LngRow, EG1_E1_insp_method_nm))) _
					& Chr(11) & Trim(ConvSPChars(EG1_inspection_standard_by_item(LngRow, EG1_E1_insp_unit_indctn_nm))) _
					& Chr(11) & Trim(ConvSPChars(EG1_inspection_standard_by_item(LngRow, EG1_E1_weight_nm))) _
					& Chr(11) & Trim(ConvSPChars(EG1_inspection_standard_by_item(LngRow, EG1_E1_insp_spec)))

  			If EG1_inspection_standard_by_item(LngRow, EG1_E1_lsl) <> "" Then
				strData = strData & Chr(11) & UniNumClientFormat(EG1_inspection_standard_by_item(LngRow, EG1_E1_lsl), ggQty.DecPoint ,0)
			Else
				strData = strData & Chr(11) & ""
			End If

			If EG1_inspection_standard_by_item(LngRow, EG1_E1_usl) <> "" Then
				strData = strData & Chr(11) & UniNumClientFormat(EG1_inspection_standard_by_item(LngRow, EG1_E1_usl), ggQty.DecPoint ,0)
			Else
				strData = strData & Chr(11) & ""
			End If
											
			strData = strData & Chr(11) & Trim(ConvSPChars(EG1_inspection_standard_by_item(LngRow, EG1_E1_mthd_of_cl_cal_nm))) _
							  & Chr(11) & UniNumClientFormat(EG1_inspection_standard_by_item(LngRow, EG1_E1_calculated_qty), ggQty.DecPoint ,0)
			
			If EG1_inspection_standard_by_item(LngRow, EG1_E1_lcl) <> "" Then
				strData = strData & Chr(11) & UniNumClientFormat(EG1_inspection_standard_by_item(LngRow, EG1_E1_lcl), ggQty.DecPoint ,0)
			Else
				strData = strData & Chr(11) & ""
			End If

			If EG1_inspection_standard_by_item(LngRow, EG1_E1_ucl) <> "" Then
				strData = strData & Chr(11) & UniNumClientFormat(EG1_inspection_standard_by_item(LngRow, EG1_E1_ucl), ggQty.DecPoint ,0)
			Else
				strData = strData & Chr(11) & ""
			End If
			
			strData = strData & Chr(11) & Trim(ConvSPChars(EG1_inspection_standard_by_item(LngRow, EG1_E1_measmt_equipmt_cd))) _
							  & Chr(11) & "" _
							  & Chr(11) & Trim(ConvSPChars(EG1_inspection_standard_by_item(LngRow, EG1_E1_measmt_equipmt_nm))) _
							  & Chr(11) & Trim(ConvSPChars(EG1_inspection_standard_by_item(LngRow, EG1_E1_measmt_unit_cd))) _
							  & Chr(11) & "" _
							  & Chr(11) & Trim(ConvSPChars(EG1_inspection_standard_by_item(LngRow, EG1_E1_insp_process_desc))) _
							  & Chr(11) & Trim(ConvSPChars(EG1_inspection_standard_by_item(LngRow, EG1_E1_remark))) _
							  & Chr(11) & Trim(ConvSPChars(EG1_inspection_standard_by_item(LngRow, EG1_E1_insp_char))) _
							  & Chr(11) & Trim(ConvSPChars(EG1_inspection_standard_by_item(LngRow, EG1_E1_insp_unit_indctn_cd))) _
							  & Chr(11) & Trim(ConvSPChars(EG1_inspection_standard_by_item(LngRow, EG1_E1_weight_cd))) _
							  & Chr(11) & Trim(ConvSPChars(EG1_inspection_standard_by_item(LngRow, EG1_E1_mthd_of_cl_cal_cd))) _
							  & Chr(11) & LngMaxRow + LngRow + 1 _
							  & Chr(11) & Chr(12)
			TmpBuffer(LngRow) = strData
		ELSE
			StrNextKey = EG1_inspection_standard_by_item(LngRow, 0)
		End If
	Next

	iTotalStr = Join(TmpBuffer, "")
%>
<Script Language=vbscript>
	With Parent		
		.ggoSpread.Source = .frm1.vspdData 
		.ggoSpread.SSShowDataByClip "<%=iTotalStr%>"		
		.lgStrPrevKey = "<%=ConvSPChars(StrNextKey)%>"
		If .frm1.vspdData.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData, 0) And .lgStrPrevKey <> "" Then
			.DbQuery
		Else
			.frm1.txtPlantNm.Value = "<%=ConvSPChars(E3_b_plant(E1_plant_nm))%>"
			.frm1.txtItemNm.Value = "<%=ConvSPChars(E2_b_item(E2_item_nm))%>"
			.frm1.txtRoutNoDesc.Value = "<%=ConvSPChars(E4_p_routing_header(E4_rout_no_desc))%>"
			.frm1.txtOprNoDesc.Value = "<%=ConvSPChars(E5_p_routing_detail(E5_opr_no_desc))%>"
			.frm1.hPlantCd.value = "<%=ConvSPChars(strPlantCd)%>"
			.frm1.hInspClassCd.value = "<%=ConvSPChars(strInspClassCd)%>"
			.frm1.hItemCd.value = "<%=ConvSPChars(strItemCd)%>"
			.frm1.hRoutNo.value = "<%=ConvSPChars(strRoutNo)%>"
			.frm1.hOprNo.value = "<%=ConvSPChars(strOprNo)%>"
			.DbQueryOk
        End If		
	End with
</Script>
