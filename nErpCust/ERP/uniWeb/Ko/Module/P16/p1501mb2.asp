<%@LANGUAGE = VBScript%>
<%'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1501mb2.asp
'*  4. Program Name         : ManageResource 저장 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/04/07
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : RYU SUNG WON
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'**********************************************************************************************%>

<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<%
Call LoadBasisGlobalInf

Dim oPP1G606

Dim strCode																	'☆ : Lookup 용 코드 저장 변수 
Dim lgIntFlgMode
Dim iCommandSent
Dim I1_P_Resource			'Array
Dim I2_P_Resource_Group_Cd	'String
Dim I3_B_Plant_Cd			'String

Const C_I1_Resource_Cd = 0
Const C_I1_Description = 1
Const C_I1_Resource_Type = 2
Const C_I1_Cost_Type = 3
Const C_I1_Max_Over_Run = 4
Const C_I1_Over_Time_Multi = 5
Const C_I1_Selection_Rule = 6
Const C_I1_Super_Flg = 7
Const C_I1_Valid_From_Dt = 8
Const C_I1_Valid_To_Dt = 9
Const C_I1_Sequence_Rule = 10
'Added
Const C_I1_No_Of_Resource = 11
Const C_I1_efficiency = 12
Const C_I1_utilization = 13
Const C_I1_run_rccp = 14
Const C_I1_run_crp = 15
Const C_I1_overload_tol = 16
Const C_I1_mfg_cost = 17
Const C_I1_rsc_base_qty = 18
Const C_I1_rsc_base_unit = 19
ReDim I1_P_Resource(C_I1_rsc_base_unit)

Call HideStatusWnd

On Error Resume Next
Err.Clear  

    If Trim(Request("txtResourceCd2")) = "" Then												'⊙: 저장을 위한 값이 들어왔는지 체크 
		Call DisplayMsgBox("900010", vbOKOnly, "", "", I_MKSCRIPT)   '⊙: 에러메세지는 DB화 한다.           
		Response.End 
	End If


	lgIntFlgMode = CInt(Request("txtFlgMode"))										'☜: 저장시 Create/Update 판별 
	
	If lgIntFlgMode = OPMD_CMODE Then
		iCommandSent = "CREATE"
    ElseIf lgIntFlgMode = OPMD_UMODE Then
		iCommandSent = "UPDATE"
    End If

	I3_B_Plant_Cd = UCase(Trim(Request("txtPlantCd")))
	I2_P_Resource_Group_Cd = UCase(Trim(Request("txtResourceGroupCd")))
	I1_P_Resource(C_I1_Resource_Cd) = UCase(Trim(Request("txtResourceCd2")))
	I1_P_Resource(C_I1_Description) = Trim(Request("txtResourceNm2"))
	I1_P_Resource(C_I1_Resource_Type) = Request("cboResourceType")
	I1_P_Resource(C_I1_Cost_Type) = Request("txtCostType")
	I1_P_Resource(C_I1_Max_Over_Run) = UniConvNum(Request("txtMaxWorkTime"),0)
	I1_P_Resource(C_I1_Over_Time_Multi) = UniConvNum(Request("txtWorkOutRate"),0)
	I1_P_Resource(C_I1_Super_Flg) = Request("rdoInfiniteResourceFlg")
	I1_P_Resource(C_I1_Selection_Rule) = UniConvNum(Request("txtSelectionRule"),0)
	I1_P_Resource(C_I1_Sequence_Rule) = UniConvNum(Request("txtSequenceRule"),0)
	'Added
	I1_P_Resource(C_I1_No_Of_Resource) = UniConvNum(Request("txtNoOfResource"),0)
	I1_P_Resource(C_I1_efficiency) = UniConvNum(Request("txtEfficiency"),0)
	I1_P_Resource(C_I1_utilization) = UniConvNum(Request("txtUtilization"),0)
	I1_P_Resource(C_I1_run_rccp) = Request("rdoRunRccp")
	I1_P_Resource(C_I1_run_crp) = "N"
	I1_P_Resource(C_I1_overload_tol) = UniConvNum(Request("txtOverloadTol"),0)
	I1_P_Resource(C_I1_mfg_cost) = UniConvNum(Request("txtMfgCost"),0)
	I1_P_Resource(C_I1_rsc_base_qty) = UniConvNum(Request("txtResourceEa"),0)
	I1_P_Resource(C_I1_rsc_base_unit) = UCase(Request("txtResourceUnitCd"))

	If Len(Trim(Request("txtValidFromDt"))) Then
		If UniConvDate(Request("txtValidFromDt")) = "" Then	 
			Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
			Call LoadTab("parent.frm1.txtValidFromDt", 0, I_MKSCRIPT)
			Response.End	
		Else
			I1_P_Resource(C_I1_Valid_From_Dt) = UniConvDate(Request("txtValidFromDt"))
		End If
	End If
	
	If Len(Trim(Request("txtValidToDt"))) Then
		If UniConvDate(Request("txtValidToDt")) = "" Then	 
			Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
			Call LoadTab("parent.frm1.txtValidToDt", 0, I_MKSCRIPT)
			Response.End	
		Else
			I1_P_Resource(C_I1_Valid_To_Dt) = UniConvDate(Request("txtValidToDt"))
		End If
	End If

	Set oPP1G606 = Server.CreateObject("PP1G606.cPMngRsrc")

    If CheckSYSTEMError(Err,True) = True Then
       Response.End 
    End If

    Call oPP1G606.P_MANAGE_RESOURCE(gStrGlobalCollection, _
								iCommandSent, _
								I1_P_Resource, _
								I2_P_Resource_Group_Cd, _
								I3_B_Plant_Cd)
    

    If CheckSYSTEMError(Err,True) = True Then
       Set oPP1G606 = Nothing		                                                 '☜: Unload Comproxy DLL
       Response.End		
    End If
	
	Set oPP1G606 = Nothing
	
	'-----------------------
	'Result data display area
	'----------------------- 

	Response.Write "<Script Language=vbscript>	" & vbCr
	Response.Write "	With parent				" & vbCr																
	Response.Write "		.DbSaveOk			" & vbCr
	Response.Write "	End With				" & vbCr
	Response.Write "</Script>					" & vbCr
					

	Response.End																				'☜: Process End

	'==============================================================================
	' 사용자 정의 서버 함수 
	'==============================================================================

	Response.Write "<SCRIPT LANGUAGE=VBSCRIPT RUNAT=SERVER>	" & vbCr
	Response.Write "										" & vbCr
	Response.Write "</SCRIPT>								" & vbCr

%>