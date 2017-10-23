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
'*  3. Program ID           : Q1215MB2
'*  4. Program Name         : 선별형검사조건 등록 
'*  5. Program Desc         : Quality Configuration
'*  6. Component List       : PQBG190
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

Dim PQBG190																	'☆ : 조회용 ComProxy Dll 사용 변수 
	
Dim lgIntFlgMode
Dim LngMaxRow
Dim iCommandSent
Dim strPlantCd
Dim strItemCd
Dim strInspClassCd
Dim strInspItemCd
Dim strRoutNo
Dim strOprNo
Dim strLotQualityIndex
Dim strPBar
Dim strValue
Dim I6_q_inspection_standard_detail3
Const Q059_I6_quality_assurance = 0
Const Q059_I6_qa_value = 1
Const Q059_I6_pbar = 2
     
lgIntFlgMode		= CInt(Request("txtFlgMode"))			'☜: 저장시 Create/Update 판별 
strPlantCd			= UCase(Request("txtPlantCd"))
strItemCd			= UCase(Request("txtItemCd"))
strInspClassCd		= UCase(Request("cboInspClassCd"))
strInspItemCd		= UCase(Request("txtInspItemCd"))
strRoutNo			= UCase(Request("txtRoutNo"))
strOprNo			= UCase(Request("txtOprNo"))
strLotQualityIndex	= Request("cboLotQualityIndex")
strPBar				= Request("txtPBar")	
	
Redim I6_q_inspection_standard_detail3(2)
	
Set PQBG190 = Server.CreateObject("PQBG190.cQMaintInspDtl3Svr")

I6_q_inspection_standard_detail3(Q059_I6_quality_assurance) = strLotQualityIndex
If strLotQualityIndex = "A" Then
	I6_q_inspection_standard_detail3(Q059_I6_qa_value) = UNIConvNum(Request("cboAOQL"), 0)
Else
	I6_q_inspection_standard_detail3(Q059_I6_qa_value) = UNIConvNum(Request("cboLTPD"), 0)
End If
I6_q_inspection_standard_detail3(Q059_I6_pbar) = UNIConvNum(strPBar, 0)
	
If lgIntFlgMode = OPMD_CMODE Then
	iCommandSent = "CREATE"
ElseIf lgIntFlgMode = OPMD_UMODE Then
	iCommandSent = "UPDATE"
End If
	
Call PQBG190.Q_MAINT_INSP_DTL3_SVR(gStrGlobalCollection, _
									iCommandSent, _
									strPlantCd, strItemCd, strInspItemCd, strInspClassCd, strRoutNo, strOprNo, _
									I6_q_inspection_standard_detail3)
	
If CheckSYSTEMError(Err,True) Then
	Set PQBG190 = Nothing
	Response.End 
End If

Set PQBG190 = Nothing
%>
<Script Language=vbscript>
With parent																		'☜: 화면 처리 ASP 를 지칭함 
	.DbSaveOk		
End With
</Script>
