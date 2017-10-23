<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%Call loadInfTB19029B("I", "*", "NOCOOKIE","MB")%>
<%Call LoadBasisGlobalInf%>
<%
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1213MB1
'*  4. Program Name         : 조정형 (공정외) 검사조건 등록 
'*  5. Program Desc         : 
'*  6. Component List       : PQBG160
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

Dim PQBG160													'☆ : 조회용 ComProxy Dll 사용 변수 
Dim strMode			
										'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim StrNextKey		' 다음 값 
Dim lgStrPrevKey	' 이전 값 
Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngRow
Dim intGroupCount   
	       
Dim strPlantCd
Dim strInspClassCd
Dim strItemCd
Dim strInspItemCd
Dim strData
Dim i

Dim E1_b_plant
Dim E2_b_item
Dim E3_q_inspection_item
Dim E4_b_minor_insp_mthd
Dim EG1_group_export	
	
Const C_SHEETMAXROWS_D = 100   
	   
lgStrPrevKey = Request("lgStrPrevKey")
LngMaxRow = Request("txtMaxRows")

strPlantCd = Request("txtplantCd")
strInspClassCd= Request("cboInspClassCd")
strItemCd = Request("txtItemCd")
strInspItemCd = Request("txtInspItemCd")

Set PQBG160 = Server.CreateObject("PQBG160.cQListInspStdDtl1Svr")

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call PQBG160.Q_LIST_INSP_STD_DTL1_SVR(gStrGlobalCollection, C_SHEETMAXROWS_D, _
					  		    	lgStrPrevKey, _
									strInspItemCd, _
									strInspClassCd, _
									strItemCd, _
									strPlantCd, _
									E1_b_plant, _
									E2_b_item, _
									E3_q_inspection_item, _
									E4_b_minor_insp_mthd, _
						 			EG1_group_export)

If CheckSYSTEMError(Err,True) = True Then
	Set PQBG160 = Nothing
	Response.End
End If

Const Q051_E1_plant_cd = 0
Const Q051_E1_plant_nm = 1

Const Q051_E2_item_cd = 0
Const Q051_E2_item_nm = 1

Const Q051_E3_insp_item_cd = 0
Const Q051_E3_insp_item_nm = 1

Const Q051_E4_minor_cd_insp_mthd_cd = 0
Const Q051_E4_minor_nm_insp_mthd_nm = 1


Const Q051_EG1_E1_minor_nm_aql_nm = 0
Const Q051_EG1_E2_minor_nm_insp_level_nm = 1
Const Q051_EG1_E3_minor_nm_switch_nm = 2
Const Q051_EG1_E4_minor_nm_substitute_nm = 3
Const Q051_EG1_E5_minor_nm_mthd_decision_nm = 4
Const Q051_EG1_E6_bp_nm_bp = 5
	    
Const Q051_EG1_E7_bp_cd = 6
Const Q051_EG1_E7_switch_cd = 7
Const Q051_EG1_E7_insp_level_cd = 8
Const Q051_EG1_E7_aql = 9
Const Q051_EG1_E7_substitute_for_sigma = 10
Const Q051_EG1_E7_mthd_of_decision = 11

Dim TmpBuffer
Dim iTotalStr

For i = 0 To UBound(EG1_group_export, 1)
    If i < C_SHEETMAXROWS_D Then
		ReDim TmpBuffer(C_SHEETMAXROWS_D - 1)
		
		strData = strData & Chr(11) & Trim(ConvSPChars(EG1_group_export(i, Q051_EG1_E7_bp_cd)))
		strData = strData & Chr(11) & ""
		strData = strData & Chr(11) & Trim(ConvSPChars(EG1_group_export(i, Q051_EG1_E6_bp_nm_bp)))
		strData = strData & Chr(11) & Trim(ConvSPChars(EG1_group_export(i, Q051_EG1_E3_minor_nm_switch_nm)))
		strData = strData & Chr(11) & Trim(ConvSPChars(EG1_group_export(i, Q051_EG1_E7_insp_level_cd)))
		strData = strData & Chr(11) & ""
		strData = strData & Chr(11) & Trim(ConvSPChars(UniConvNumDBToCompanyWithOutChange(EG1_group_export(i, Q051_EG1_E7_aql), 0)))		
		strData = strData & Chr(11) & ""
		strData = strData & Chr(11) & Trim(ConvSPChars(EG1_group_export(i, Q051_EG1_E4_minor_nm_substitute_nm)))
		strData = strData & Chr(11) & Trim(ConvSPChars(EG1_group_export(i, Q051_EG1_E5_minor_nm_mthd_decision_nm)))
			
		strData = strData & Chr(11) & Trim(ConvSPChars(EG1_group_export(i, Q051_EG1_E7_switch_cd)))
		strData = strData & Chr(11) & Trim(ConvSPChars(EG1_group_export(i, Q051_EG1_E7_substitute_for_sigma)))
		strData = strData & Chr(11) & Trim(ConvSPChars(EG1_group_export(i, Q051_EG1_E7_mthd_of_decision)))
		strData = strData & Chr(11) & LngMaxRow + i + 1
		strData = strData & Chr(11) & Chr(12)
		TmpBuffer(i) = strData
    Else
		StrNextKey = EG1_group_export(i,Q051_EG1_E7_bp_cd)
    End If
Next  

iTotalStr = Join(TmpBuffer, "")

Set PQBG160 = Nothing
%>
<Script Language=vbscript>
With Parent
	.frm1.txtInspMthdCd.Value = "<%=ConvSPChars(E4_b_minor_insp_mthd(Q051_E4_minor_cd_insp_mthd_cd))%>"
	.frm1.txtInspMthdNm.Value = "<%=ConvSPChars(E4_b_minor_insp_mthd(Q051_E4_minor_nm_insp_mthd_nm))%>"
			
	.frm1.txtPlantCd.Value = "<%=ConvSPChars(E1_b_plant(Q051_E1_plant_cd))%>"
	.frm1.txtPlantNm.Value = "<%=ConvSPChars(E1_b_plant(Q051_E1_plant_nm))%>"
			
	.frm1.txtItemCd.Value = "<%=ConvSPChars(E2_b_item(Q051_E2_item_cd))%>"
	.frm1.txtItemNm.Value = "<%=ConvSPChars(E2_b_item(Q051_E2_item_nm))%>"

	.frm1.txtInspItemCd.Value = "<%=ConvSPChars(E3_q_inspection_item(Q051_E3_insp_item_cd))%>"
	.frm1.txtInspItemNm.Value = "<%=ConvSPChars(E3_q_inspection_item(Q051_E3_insp_item_nm))%>"

	.ggoSpread.Source = .frm1.vspdData 
	.ggoSpread.SSShowDataByClip "<%=iTotalStr%>"		
	.lgStrPrevKey = "<%=ConvSPChars(StrNextKey)%>"
			
	If .frm1.vspdData.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData, 0) And .lgStrPrevKey <> "" Then
		.DbQuery
	Else
		.frm1.hPlantCd.value = "<%=ConvSPChars(strPlantCd)%>"
		.frm1.hInspClassCd.value 	= "<%=ConvSPChars(strInspClassCd)%>"
		.frm1.hItemCd.value 	= "<%=ConvSPChars(strItemCd)%>"
		.frm1.hInspItemCd.value 	= "<%=ConvSPChars(strInspItemCd)%>"
							 
		.DbQueryOk
	End If				
End with			
</Script>