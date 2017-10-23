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
'*  3. Program ID           : Q1214MB2
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

Const IG1_plant_cd = 0
Const IG1_item_cd = 1
Const IG1_insp_class_cd = 2
Const IG1_rout_no = 3
Const IG1_opr_no = 4
Const IG1_insp_item_cd = 5
Const IG1_switch_cd = 6
Const IG1_insp_level_cd = 7
Const IG1_aql = 8
Const IG1_insp_cnt = 9
Const IG1_rejt_cnt = 10
Const IG1_sigma = 11
Const IG1_substitute_for_sigma = 12
Const IG1_mthd_of_decision = 13
    		
Dim PQBG180																	'☆ : 조회용 ComProxy Dll 사용 변수 
Dim iCommandSent
	
Dim lgIntFlgMode

Dim IG1_q_inspection_standard_detail2

Redim IG1_q_inspection_standard_detail2(13)
	
lgIntFlgMode	= CInt(Request("txtFlgMode"))					'☜: 저장시 Create/Update 판별 
	
IG1_q_inspection_standard_detail2(IG1_plant_cd) = UCase(Request("txtPlantCd"))
IG1_q_inspection_standard_detail2(IG1_item_cd) = UCase(Request("txtItemCd"))
IG1_q_inspection_standard_detail2(IG1_insp_class_cd) = UCase(Request("cboInspClassCd"))
IG1_q_inspection_standard_detail2(IG1_rout_no) = UCase(Request("txtRoutNo"))
IG1_q_inspection_standard_detail2(IG1_opr_no) = UCase(Request("txtOprNo"))
IG1_q_inspection_standard_detail2(IG1_insp_item_cd) = UCase(Request("txtInspItemCd"))
IG1_q_inspection_standard_detail2(IG1_switch_cd) = UCase(Request("cboSwitch"))
IG1_q_inspection_standard_detail2(IG1_insp_level_cd) = UCase(Request("txtInspLevel"))
IG1_q_inspection_standard_detail2(IG1_aql) = UNIConvNum(Request("txtAQL"), 0)
IG1_q_inspection_standard_detail2(IG1_substitute_for_sigma) = UCase(Request("cboSubstituteForSigma"))
IG1_q_inspection_standard_detail2(IG1_mthd_of_decision) = UCase(Request("cboMthdOfDecision"))

If lgIntFlgMode = OPMD_CMODE Then
	iCommandSent = "CREATE"
ElseIf lgIntFlgMode = OPMD_UMODE Then
	iCommandSent = "UPDATE"
End If
	
Set PQBG180 = Server.CreateObject("PQBG180.cQMaintInspDtl2Svr")

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If	
	
Call PQBG180.Q_MAINT_INSP_DTL2_SVR(gStrGlobalCollection, iCommandSent, IG1_q_inspection_standard_detail2)
	 
If CheckSYSTEMError(Err,True) Then
	Set PQBG180 = Nothing
	Response.End 
End If	
	
Set PQBG180 = Nothing                                                   '☜: Unload Comproxy
%>
<Script Language=vbscript>
With parent																		'☜: 화면 처리 ASP 를 지칭함 
	.DbSaveOk
End With
</Script>
