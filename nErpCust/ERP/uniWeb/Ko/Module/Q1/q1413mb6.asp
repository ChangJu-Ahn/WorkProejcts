<%@LANGUAGE = VBScript%> 
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadinfTB19029.asp" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->	
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1413MB6
'*  4. Program Name         : 계량 조정형 n,k 선정 
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
%>
<!-- ChartFX용 상수를 사용하기 위한 Include 지정 -->
<!-- #include file="../../inc/CfxIE.inc" -->
<%													
On Error Resume Next

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
Err.Clear

Call LoadBasisGlobalInf
Call LoadinfTB19029B("I", "*", "NOCOOKIE", "QB") 

Dim PQBG280
Dim strRigor
Dim strInsLevel
Dim strSTDack
Dim strInsCri

Dim strLotSize
Dim strAQL
Dim strSD
Dim strLevel

Dim strUpperBound
Dim strLowerBound

Dim I1_qm_input_workset
Const Q094_I1_quality_assurance = 0
Const Q094_I1_insp_method_cd = 1
Const Q094_I1_alpha = 2
Const Q094_I1_beta = 3
Const Q094_I1_p1 = 4
Const Q094_I1_p2 = 5
Const Q094_I1_lot_size = 6
Const Q094_I1_lq = 7
Const Q094_I1_aql = 8
Const Q094_I1_switch_cd = 9
Const Q094_I1_insp_level_cd = 10
Const Q094_I1_substitute_for_sigma = 11
Const Q094_I1_ltpd_pt = 12
Const Q094_I1_aoql = 13
Const Q094_I1_pbar = 14
Const Q094_I1_qa_value = 15


Dim EG1_export_group
Const Q094_EG1_E1_qm_output_workset_m = 0
Const Q094_EG1_E1_qm_output_workset_k = 1
Const Q094_EG1_E1_qm_output_workset_n_qty = 2
Const Q094_EG1_E1_qm_output_workset_ac_qty = 3
Const Q094_EG1_E1_qm_output_workset_re_qty = 4

strRigor = Request("txtRigor")
strLevel = Request("txtDefectMode")

strSTDack = Request("txtSTDack")
strInsLevel = Request("txtInsCri")

strLotsize = Request("txtLotSize")
strAQL = Request("txtAQL")
strSD = Request("txtSD")

strUpperBound = Request("txtUpperBound")
strLowerBound = Request("txtLowerBound")

Set PQBG280 = Server.CreateObject("PQBG280.cQLookUpNCSvr")

If CheckSYSTEMError(Err,True) Then
   Response.End
End if  

Redim I1_qm_input_workset(15)
I1_qm_input_workset(Q094_I1_insp_method_cd) = "2311"
I1_qm_input_workset(Q094_I1_switch_cd) = strRigor
I1_qm_input_workset(Q094_I1_substitute_for_sigma) = strSTDack
I1_qm_input_workset(Q094_I1_insp_level_cd) = strLevel
I1_qm_input_workset(Q094_I1_lot_size) = UNIConvNum(strLotsize, 0)
I1_qm_input_workset(Q094_I1_aql) = UNIConvNum(strAQL, 0)

EG1_export_group = PQBG280.Q_LOOK_UP_NC_SVR(gStrGlobalCollection,I1_qm_input_workset)

If CheckSYSTEMError(Err,True) Then
   Set PQBG280 = Nothing
   Response.End
End if  

Set PQBG280 = Nothing
%>
<Script Language=vbscript>
With Parent.frm1	
	.txtSampleSize.Value = "<%=UniNumClientFormat(EG1_export_group(0,Q094_EG1_E1_qm_output_workset_n_qty), ggQty.DecPoint ,0)%>"
	If .rdoForm.rdoForm1.checked = true then
		.txtAcceptSize.Value = "<%=UniNumClientFormat(EG1_export_group(0,Q094_EG1_E1_qm_output_workset_k), ggQty.DecPoint ,0)%>"
	Else
		.txtMaxTolPercent.Value = "<%=UniNumClientFormat(EG1_export_group(0,Q094_EG1_E1_qm_output_workset_m), ggQty.DecPoint ,0)%>"
	End If
End with
</Script>	
