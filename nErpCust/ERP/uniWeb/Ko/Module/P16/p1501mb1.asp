<%@LANGUAGE = VBScript%>
<%'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1501mb1.asp
'*  4. Program Name         : ManageResource 조회 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/04/04
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : RYU SUNG WON
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'**********************************************************************************************%>

<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("I", "*", "NOCOOKIE", "MB")

Dim oPP1G605

Dim strCode																	'☆ : Lookup 용 코드 저장 변수 
Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim strPrevNextFlg
Dim strPlantCd
Dim GroupCount, GroupCount1
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter 선언 
Dim rs0, rs1, rs2							'DBAgent Parameter 선언 
Dim ADF										'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg								'Record Set Return Message 변수선언 

Dim I1_P_Resource
Dim E1_P_Resource
Dim E2_B_Plant
Dim E3_P_Resource_Group
Dim E4_StatusCodeOfPrevNext

Const C_I1_resource_cd = 0
Const C_I1_valid_to_dt = 1

Const C_E1_resource_cd = 0
Const C_E1_description = 1
Const C_E1_resource_type = 2
Const C_E1_cost_type = 3
Const C_E1_max_over_run = 4
Const C_E1_over_time_multi = 5
Const C_E1_selection_rule = 6
Const C_E1_super_flg = 7
Const C_E1_valid_from_dt = 8
Const C_E1_valid_to_dt = 9
Const C_E1_sequence_rule = 10
'Added
Const C_E1_No_Of_Resource = 11
Const C_E1_efficiency = 12
Const C_E1_utilization = 13
Const C_E1_run_rccp = 14
Const C_E1_run_crp = 15
Const C_E1_overload_tol = 16
Const C_E1_mfg_cost = 17
Const C_E1_rsc_base_qty = 18
Const C_E1_rsc_base_unit = 19

Const C_E2_plant_cd = 0
Const C_E2_plant_nm = 1

Const C_E3_resource_group_cd = 0
Const C_E3_description = 1

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

On Error Resume Next														'☜: 
Err.Clear																	'☜: Protect system from crashing

	strMode = Request("txtMode")												'☜ : 현재 상태를 받음 
	strPrevNextFlg = Request("PrevNextFlg")
	strPlantCd = Request("txtPlantCd")
 
	Redim UNISqlId(1)
	Redim UNIValue(1, 1)

	UNISqlId(0) = "180000saa"
	UNISqlId(1) = "180000san"

	UNIValue(0, 0) = " " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & " "
	
	UNIValue(1, 0) = " " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & " "
	UNIValue(1, 1) = " " & FilterVar(UCase(Request("txtResourceCd1")), "''", "S") & " "
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
	Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)
	
	%>
	<Script Language=vbscript>
		parent.frm1.txtPlantNm.value = ""
		parent.frm1.txtResourceNm1.value = ""
		parent.frm1.txtCurCd.value = ""		'통화코드	'ex : KRW
	</Script>	
	<%

	' Plant 명 Display      
	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
		%>
		<Script Language=vbscript>
		parent.frm1.txtPlantCd.Focus()
		</Script>	
		<%
		rs0.Close
		Set rs0 = Nothing
		Set ADF = Nothing
		Response.End
	Else
		%>
		<Script Language=vbscript>
		parent.frm1.txtPlantNm.value = "<%=ConvSPChars(rs0("PLANT_NM"))%>"
		parent.frm1.txtCurCd.value = "<%=ConvSPChars(rs0("CUR_CD"))%>"		'통화코드	'ex : KRW
		</Script>	
		<%
		rs0.Close
		Set rs0 = Nothing
	End If

	' 자원명 Display
	If (rs1.EOF And rs1.BOF) Then
		Call DisplayMsgBox("181604", vbOKOnly, "", "", I_MKSCRIPT)	'#####
		%>
		<Script Language=vbscript>
		parent.frm1.txtResourceCd1.Focus()
		</Script>	
		<%
		rs1.Close
		Set rs1 = Nothing
		Set ADF = Nothing
		Response.End
	Else
		%>
		<Script Language=vbscript>
		parent.frm1.txtResourceNm1.value = "<%=ConvSPChars(rs1("Description"))%>"
		</Script>	
		<%
		rs1.Close
		Set rs1 = Nothing
	End If
	
	
	
	ReDim I1_p_resource(1)
	I1_p_resource(C_I1_resource_cd) = Request("txtResourceCd1")
	I1_p_resource(C_I1_valid_to_dt) = ""

	Set oPP1G605 = Server.CreateObject("PP1G605.cPLkUpRsrcSvr")

    If CheckSYSTEMError(Err,True) = True Then
		Response.End 
    End If

	Call oPP1G605.P_LOOK_UP_RESOURCE_SVR(gStrGlobalCollection, _
										strPrevNextFlg, _
										strPlantCd, _
										I1_P_Resource, _
										E1_P_Resource, _
										E2_B_Plant, _
										E3_P_Resource_Group, _
										E4_StatusCodeOfPrevNext)

    If CheckSYSTEMError(Err,True) = True Then
       Set oPP1G605 = Nothing
       Response.End 
    End If
    
    Set oPP1G605 = Nothing															'☜: Unload Comproxy
    
	If (E4_StatusCodeOfPrevNext = "900011" Or E4_StatusCodeOfPrevNext = "900012") Then
		Call DisplayMsgBox(E4_StatusCodeOfPrevNext, vbInformation, "", "", I_MKSCRIPT)	'⊙: DB 에러코드, 메세지타입, %처리, 스크립트유형 
	End If
	
	Response.Write "<Script Language = VBScript> " & vbCr
	Response.Write "Dim LngRow " & vbCr
	
	Response.Write "With parent.frm1 " & vbCr
	Response.Write "	.txtPlantCd.value			= """ & ConvSPChars(E2_b_plant(C_E2_plant_cd)) & """" & vbCr
	Response.Write "	.txtPlantNm.value			= """ & ConvSPChars(E2_b_plant(C_E2_plant_nm)) & """" & vbCr
	Response.Write "	.txtResourceCd1.value		= """ & ConvSPChars(E1_p_resource(C_E1_resource_cd)) & """" & vbCr
	Response.Write "	.txtResourceNm1.value		= """ & ConvSPChars(E1_p_resource(C_E1_description)) & """" & vbCr
	Response.Write "	.txtResourceCd2.value		= """ & ConvSPChars(E1_p_resource(C_E1_resource_cd)) & """" & vbCr
	Response.Write "	.txtResourceNm2.value		= """ & ConvSPChars(E1_p_resource(C_E1_description)) & """" & vbCr
	Response.Write "	.cboResourceType.value		= """ & ConvSPChars(UCase(E1_p_resource(C_E1_resource_type))) & """" & vbCr
	Response.Write "	.txtResourceGroupCd.value	= """ & ConvSPChars(E3_p_resource_group(C_E3_resource_group_cd)) & """" & vbCr
	Response.Write "	.txtResourceGroupNm.value   = """ & ConvSPChars(E3_p_resource_group(C_E3_description)) & """" & vbCr
	Response.Write "	.txtNoOfResource.value		= """ & UNINumClientFormat(E1_p_resource(C_E1_No_Of_Resource),ggQty.DecPoint,0) & """" & vbCr
	Response.Write "	.txtEfficiency.value		= """ & UNINumClientFormat(E1_p_resource(C_E1_efficiency),ggQty.DecPoint,0) & """" & vbCr
	Response.Write "	.txtUtilization.value		= """ & UNINumClientFormat(E1_p_resource(C_E1_utilization),ggQty.DecPoint,0) & """" & vbCr
	
	Response.Write "	If """ & E1_p_resource(C_E1_run_rccp) & """ = ""Y"" Then  " & vbCr 'adsf
	Response.Write "		.rdoRunRccp1.checked = True " & vbCr
	Response.Write "	Else " & vbCr
	Response.Write "		.rdoRunRccp2.checked = True " & vbCr
	Response.Write "	End If	" & vbCr
	Response.Write "	If """ & E1_p_resource(C_E1_run_crp) & """ = ""Y"" Then  " & vbCr 'adsf
	Response.Write "		.rdoRunCrp1.checked = True " & vbCr
	Response.Write "	Else " & vbCr
	Response.Write "		.rdoRunCrp2.checked = True " & vbCr
	Response.Write "	End If	" & vbCr	
	Response.Write "	.txtOverloadTol.value	= """ & UNINumClientFormat(E1_p_resource(C_E1_overload_tol),ggQty.DecPoint,0) & """" & vbCr
	Response.Write "	.txtMfgCost.text		= """ & UniNumClientFormat(E1_p_resource(C_E1_mfg_cost),ggUnitCost.DecPoint,0) & """" & vbCr   
	Response.Write "	.txtResourceEa.text		= """ & UniNumClientFormat(E1_p_resource(C_E1_rsc_base_qty),ggQty.DecPoint,0) & """" & vbCr   
	Response.Write "	.txtResourceUnitCd.value= """ & ConvSPChars(E1_p_resource(C_E1_rsc_base_unit)) & """" & vbCr
	Response.Write "	.txtResourceUnitCd1.value= """ & ConvSPChars(E1_p_resource(C_E1_rsc_base_unit)) & """" & vbCr
	
	Response.Write "	.txtValidFromDt.text	= """ & UniDateClientFormat(E1_p_resource(C_E1_valid_from_dt)) & """" & vbCr
	Response.Write "	.txtValidToDt.text		= """ & UniDateClientFormat(E1_p_resource(C_E1_valid_to_dt)) & """" & vbCr
	Response.Write "	.hPlantCd.value			= """ & ConvSPChars(E2_b_plant(C_E2_plant_cd)) & """" & vbCr
	Response.Write "	.hResourceCd.value		= """ & ConvSPChars(E1_p_resource(C_E1_resource_cd)) & """" & vbCr
	Response.Write "	parent.lgNextNo = """"" & vbCr		' 다음 키 값 넘겨줌 
	Response.Write "	parent.lgPrevNo = """"" & vbCr		' 이전 키 값 넘겨줌 , 현재 ComProxy가 제대로 안되 있음 
	Response.Write "	parent.DbQueryOk " & vbCr								'☜: 조회가 성공 
	Response.Write "End With " & vbCr
	Response.Write "</Script> " & vbCr
	Response.End																	'☜: Process End

'==============================================================================
' 사용자 정의 서버 함수 
'==============================================================================

	Response.Write "<SCRIPT LANGUAGE=VBSCRIPT RUNAT=SERVER> " & vbCr
	Response.Write "" & vbCr
	Response.Write "</SCRIPT> " & vbCr
%>

