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
'*  3. Program ID           : Q1214MB1
'*  4. Program Name         : 조정형 (공정) 검사조건 등록 
'*  5. Program Desc         : 
'*  6. Component List       : PQBG180
'*  7. Modified date(First) : 2004/05/07
'*  8. Modified date(Last)  : 2004/05/07
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Koh Jae Woo
'* 11. Comment
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
On Error Resume Next
Call HideStatusWnd 

Const EG1_switch_cd = 0
Const EG1_insp_level_cd = 1
Const EG1_aql = 2
Const EG1_insp_cnt = 3
Const EG1_rejt_cnt = 4
Const EG1_sigma = 5
Const EG1_substitute_for_sigma = 6
Const EG1_mthd_of_decision = 7

Dim PQBG170													'☆ : 조회용 ComProxy Dll 사용 변 
Dim strMode													'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

Dim strPlantCd
Dim strInspClassCd
Dim strItemCd
Dim strRoutNo
Dim strOprNo
Dim strInspItemCd

Dim E1_plant_nm
Dim E2_item_nm
Dim E3_rout_no_desc
Dim E4_opr_no_desc
Dim E5_insp_item_nm
Dim E6_insp_method_cd
Dim E7_insp_method_nm

Dim EG1_q_inspection_standard_detail2

strPlantCd		= Request("txtplantCd")
strInspClassCd	= Request("cboInspClassCd")
strItemCd		= Request("txtItemCd")
strRoutNo		= Request("txtRoutNo")
strOprNo		= Request("txtOprNo")
strInspItemCd	= Request("txtInspItemCd")

Set PQBG170 = Server.CreateObject("PQBG170.cQLookInspStdDtl2Svr")

Call PQBG170.Q_LOOK_UP_INSP_STAND_DETAIL2(gStrGlobalCollection, strPlantCd, strItemCd, strInspClassCd, strRoutNo, strOprNo, strInspItemCd, _
									E1_plant_nm, E2_item_nm, E3_rout_no_desc, E4_opr_no_desc, E5_insp_item_nm, E6_insp_method_cd, E7_insp_method_nm, _
									EG1_q_inspection_standard_detail2)

If CheckSYSTEMError(Err,True) Then
	Set PQBG170 = Nothing
	Response.End 
End If 

Set PQBG170 = Nothing
%>
<Script Language=vbscript>
Dim strData	
With Parent
	.frm1.txtPlantNm.Value = "<%=ConvSPChars(E1_plant_nm)%>"
	.frm1.txtItemNm.Value = "<%=ConvSPChars(E2_item_nm)%>" 
	.frm1.txtRoutNoDesc.Value = "<%=ConvSPChars(E3_rout_no_desc)%>"
	.frm1.txtOprNoDesc.Value = "<%=ConvSPChars(E4_opr_no_desc)%>"
	.frm1.txtInspItemNm.Value = "<%=ConvSPChars(E5_insp_item_nm)%>" 
	.frm1.txtInspMthdCd.Value = "<%=ConvSPChars(E6_insp_method_cd)%>"
	.frm1.txtInspMthdNm.Value = "<%=ConvSPChars(E7_insp_method_nm)%>"
	.frm1.cboSwitch.Value = "<%=ConvSPChars(EG1_q_inspection_standard_detail2(EG1_switch_cd))%>"
	.frm1.txtInspLevel.Value = "<%=ConvSPChars(EG1_q_inspection_standard_detail2(EG1_insp_level_cd))%>"
	.frm1.txtAQL.Text = "<%=UniConvNumDBToCompanyWithOutChange(EG1_q_inspection_standard_detail2(EG1_aql), 0)%>"
	.frm1.cboSubstituteForSigma.Value = "<%=ConvSPChars(EG1_q_inspection_standard_detail2(EG1_substitute_for_sigma))%>"
	.frm1.cboMthdOfDecision.Value = "<%=ConvSPChars(EG1_q_inspection_standard_detail2(EG1_mthd_of_decision))%>"
End with
Parent.DbQueryOk
</Script>
