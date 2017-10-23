<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%Call loadInfTB19029B("I", "*", "NOCOOKIE","MB")%>
<%Call LoadBasisGlobalInf%>
<%
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1212MB1
'*  4. Program Name         : 기타검사조건 등록 
'*  5. Program Desc         : Quality Configuration
'*  6. Component List       : PQBG140
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

Dim PQBG140													'☆ : 조회용 ComProxy Dll 사용 변수 
Dim strMode													'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim StrNextKey		' 다음 값 
Dim lgStrPrevKey	' 이전 값 
Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngRow
Dim intGroupCount          
Dim i
Dim strData
	
Dim strPlantCd
Dim strInspClassCd
Dim strItemCd
Dim strInspItemCd
Dim strRoutNo
Dim strOprNo
Dim strInspSeries

Dim E1_plant_nm
Dim E2_item_nm
Dim E3_insp_item_nm
Dim E4_rout_no_desc
Dim E5_opr_no_desc
Dim E6_insp_method
Dim EG1_group_export
                    	
'Q_inspection_condition
Const Q045_EG1_E1_insp_series = 0
Const Q045_EG1_E1_sample_qty = 1
Const Q045_EG1_E1_accpt_decision_qty = 2
Const Q045_EG1_E1_rejt_decision_qty = 3
Const Q045_EG1_E1_accpt_decision_discreate = 4
Const Q045_EG1_E1_max_defect_ratio = 5
    
'Q_inspection_standard_by_item 의 insp_method
Const Q045_E2_insp_method_cd = 0
Const Q045_E2_insp_method_nm = 1
    
    	    
lgStrPrevKey	= Request("lgStrPrevKey")
LngMaxRow		= Request("txtMaxRows")
strPlantCd		= Request("txtplantCd")
strInspClassCd	= Request("cboInspClassCd")
strItemCd		= Request("txtItemCd")
strInspItemCd	= Request("txtInspItemCd")
strRoutNo		= Request("txtRoutNo")
strOprNo		= Request("txtOprNo")

Const C_SHEETMAXROWS_D = 100

If lgStrPrevKey <> "" Then
	strInspSeries = lgStrPrevKey
End If

Set PQBG140 = Server.CreateObject("PQBG140.cQListInspCndtnSvr")

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call PQBG140.Q_LIST_INSP_CNDTN_SVR 	(gStrGlobalCollection , _
										C_SHEETMAXROWS_D, _
										strInspSeries, _
										strInspItemCd , _
										strItemCd , _
										strInspClassCd , _
										strPlantCd , _
										strRoutNo, _
										strOprNo, _
										E1_plant_nm , _
										E2_item_nm , _
										E3_insp_item_nm , _
										E4_rout_no_desc , _
										E5_opr_no_desc , _
										E6_insp_method , _
										EG1_group_export )

If CheckSYSTEMError(Err,True) = True Then
	Set PQBG140 = Nothing
	Response.End														'☜: 비지니스 로직 처리를 종료함 
End If

Set PQBG140 = Nothing

If IsEmpty(EG1_group_export) = true then
	Set PQBG140 = Nothing
	Response.End
End If

Dim TmpBuffer
Dim iTotalStr
		
For i = 0 To UBound(EG1_group_export, 1)
    If i < C_SHEETMAXROWS_D Then
		ReDim TmpBuffer(C_SHEETMAXROWS_D - 1)
		
		strData = strData & Chr(11) & Trim(ConvSPChars(EG1_group_export(i, Q045_EG1_E1_insp_series)))										'insp_series
		strData = strData & Chr(11) & UniNumClientFormat(EG1_group_export(i, Q045_EG1_E1_sample_qty), ggQty.DecPoint ,0)					'sample_qty		
		strData = strData & Chr(11) & UniNumClientFormat(EG1_group_export(i, Q045_EG1_E1_accpt_decision_qty), ggQty.DecPoint ,0)					'accpt_decision_qty
		strData = strData & Chr(11) & UniNumClientFormat(EG1_group_export(i, Q045_EG1_E1_rejt_decision_qty), ggQty.DecPoint ,0)					'rejt_decision_qty
		
		If EG1_group_export(i, 4) <> "" Or IsNull(EG1_group_export(i, Q045_EG1_E1_accpt_decision_discreate)) Then
			strData = strData & Chr(11) & UniNumClientFormat(EG1_group_export(i, Q045_EG1_E1_accpt_decision_discreate), ggQty.DecPoint ,0)				'accpt_decision_discreate
		Else
			strData = strData & Chr(11) & EG1_group_export(i, Q045_EG1_E1_accpt_decision_discreate)													'accpt_decision_discreate
		End If

		If EG1_group_export(i, 5) <> "" Or IsNull(EG1_group_export(i, Q045_EG1_E1_max_defect_ratio)) Then
			strData = strData & Chr(11) & UniNumClientFormat(EG1_group_export(i, Q045_EG1_E1_max_defect_ratio), ggQty.DecPoint ,0)				'accpt_decision_discreate
		Else
			strData = strData & Chr(11) & EG1_group_export(i, Q045_EG1_E1_max_defect_ratio)													'accpt_decision_discreate
		End If

		strData = strData & Chr(11) & LngMaxRow + i + 1
		strData = strData & Chr(11) & Chr(12)
		TmpBuffer(i) = strData
    ELSE 
		StrNextKey = EG1_group_export(i, Q045_EG1_E1_insp_series)
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
		.frm1.txtPlantNm.value = "<%=ConvSPChars(E1_plant_nm)%>"
		.frm1.txtItemNm.value = "<%=ConvSPChars(E2_item_nm)%>"
		.frm1.txtRoutNoDesc.value = "<%=ConvSPChars(E4_rout_no_desc)%>"
		.frm1.txtOprNoDesc.Value = "<%=ConvSPChars(E5_opr_no_desc)%>"
		.frm1.txtInspItemNm.value = "<%=ConvSPChars(E3_insp_item_nm)%>"
		.frm1.txtInspMthdCd.value = "<%=ConvSPChars(E6_insp_method(Q045_E2_insp_method_cd))%>"
		.frm1.txtInspMthdNm.value = "<%=ConvSPChars(E6_insp_method(Q045_E2_insp_method_nm))%>"
		
		.frm1.hPlantCd.value = "<%=ConvSPChars(strPlantCd)%>"
		.frm1.hInspClassCd.value = "<%=ConvSPChars(strInspClassCd)%>"
		.frm1.hItemCd.value = "<%=ConvSPChars(strItemCd)%>"
		.frm1.hInspItemCd.value = "<%=ConvSPChars(strInspItemCd)%>"
		.frm1.hRoutNo.value = "<%=ConvSPChars(strRoutNo)%>"
		.frm1.hOprNo.value = "<%=ConvSPChars(strOprNo)%>"
				 		 
		.DbQueryOk
	End If		
End with
</Script>	
<%
Set PQBG140 = Nothing
%>