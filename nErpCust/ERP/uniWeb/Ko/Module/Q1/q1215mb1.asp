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
'*  3. Program ID           : Q1215MB1
'*  4. Program Name         : 선별형검사조건 등록 
'*  5. Program Desc         : Quality Configuration
'*  6. Component List       : PQBS201
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

Dim PQBS201													'☆ : 조회용 ComProxy Dll 사용 변 
Dim strMode													'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

Dim strPlantCd
Dim strInspClassCd
Dim strItemCd
Dim strInspItemCd
Dim strRoutNo
Dim strOprNo

Dim E1_q_inspection_standard_by_item
Dim E7_q_inspection_standard_detail3

Const Q064_E1_plant_cd = 0
Const Q064_E1_plant_nm = 1
Const Q064_E1_item_cd = 2
Const Q064_E1_item_nm = 3
Const Q064_E1_insp_item_cd = 4
Const Q064_E1_insp_item_nm = 5
Const Q064_E1_insp_class_cd = 6
Const Q064_E1_insp_method_cd = 7
Const Q064_E1_insp_method_nm = 8
Const Q064_E1_rout_no = 9
Const Q064_E1_rout_no_desc = 10
Const Q064_E1_opr_no = 11
Const Q064_E1_opr_no_desc = 12
    
Const Q064_E2_quality_assurance = 0
Const Q064_E2_qa_value = 1
Const Q064_E2_pbar = 2
    
strPlantCd		= Request("txtplantCd")
strInspClassCd	= Request("cboInspClassCd")
strItemCd		= Request("txtItemCd")
strInspItemCd	= Request("txtInspItemCd")
strRoutNo		= Request("txtRoutNo")
strOprNo		= Request("txtOprNo")

Set PQBS201 = Server.CreateObject("PQBS201.cQLookInspStdDtl3Svr")

Call PQBS201.Q_LOOK_UP_INSP_STAND_DETAIL3(gStrGlobalCollection, strPlantCd, strItemCd, strInspItemCd,strInspClassCd, strRoutNo, strOprNo, _
										E1_q_inspection_standard_by_item,  E7_q_inspection_standard_detail3)

If CheckSYSTEMError(Err,True) Then
	Set PQBG201 = Nothing
	Response.End 
End If 

Set PQBG201 = Nothing
%>
<Script Language=vbscript>
Dim strData	
With Parent
	.frm1.txtPlantNm.Value			= "<%=ConvSPChars(E1_q_inspection_standard_by_item(Q064_E1_plant_nm))%>"
	.frm1.txtItemNm.Value			= "<%=ConvSPChars(E1_q_inspection_standard_by_item(Q064_E1_item_nm))%>" 
	.frm1.txtInspItemNm.Value		= "<%=ConvSPChars(E1_q_inspection_standard_by_item(Q064_E1_insp_item_nm))%>" 
	.frm1.txtInspMthdCd.Value		= "<%=ConvSPChars(E1_q_inspection_standard_by_item(Q064_E1_insp_method_cd))%>"
	.frm1.txtInspMthdNm.Value		= "<%=ConvSPChars(E1_q_inspection_standard_by_item(Q064_E1_insp_method_nm))%>"
	.frm1.txtRoutNo.Value			= "<%=ConvSPChars(E1_q_inspection_standard_by_item(Q064_E1_rout_no))%>"
	.frm1.txtRoutNoDesc.Value		= "<%=ConvSPChars(E1_q_inspection_standard_by_item(Q064_E1_rout_no_desc))%>"
	.frm1.txtOprNo.Value			= "<%=ConvSPChars(E1_q_inspection_standard_by_item(Q064_E1_opr_no))%>"
	.frm1.txtOprNoDesc.Value		= "<%=ConvSPChars(E1_q_inspection_standard_by_item(Q064_E1_opr_no_desc))%>"
	.frm1.cboLotQualityIndex.Value	= "<%=ConvSPChars(E7_q_inspection_standard_detail3(Q064_E2_quality_assurance))%>"

	If "<%=ConvSPChars(E7_q_inspection_standard_detail3(Q064_E2_quality_assurance))%>" = "A" Then
		.frm1.cboAOQL.Value = parent.UniConvNumPCToCompanyWithoutRound("<%=E7_q_inspection_standard_detail3(Q064_E2_qa_value)%>", 0)
	Else
		.frm1.cboLTPD.Value = parent.UniConvNumPCToCompanyWithoutRound("<%=E7_q_inspection_standard_detail3(Q064_E2_qa_value)%>", 0)
	End If

	.frm1.txtPBar.Value = "<%=UniNumClientFormat(E7_q_inspection_standard_detail3(Q064_E2_pbar), 4, 0)%>"
End with
Parent.DbQueryOk
</Script>	
