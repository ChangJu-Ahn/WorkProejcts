<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!--'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: p2213mb2.asp
'*  4. Program Name			: 
'*  5. Program Desc			: MPS Approval
'*  6. Comproxy List		: PP2G102.cPCnfmMrpSvr
'*  7. Modified date(First)	:
'*  8. Modified date(Last) 	: 2002/06/19
'*  9. Modifier (First)		: Lee Hyun Jae
'* 10. Modifier (Last)		: Jung Yu Kyung
'* 11. Comment				:
'**********************************************************************************************-->

<% 
Call LoadBasisGlobalInf
Call HideStatusWnd

On Error Resume Next

Dim ADF
Dim strRetMsg
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1

Dim lgStrPrevKey	' ÀÌÀü °ª 
Dim strStatus

	IF Request("txtMPSHistoryNo") = "" THEN
		Call DisplayMsgBox("187742", vbInformation, "", "", I_MKSCRIPT)
		Response.End
	END IF

	lgStrPrevKey = Request("lgStrPrevKey")
	
	Redim UNISqlId(1)
	Redim UNIValue(1, 1)
	
	UNISqlId(0) = "184000sab"
	UNISqlId(1) = "184000saa"

	UNIValue(0, 0) = " " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & ""
	UNIValue(0, 1) = " " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & ""	

	UNIValue(1, 0) = " " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & ""		
		
	UNILock = DISCONNREAD :	UNIFlag = "1"

    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)

	If NOT(rs1.EOF And rs1.BOF) Then
		strStatus = rs0("status")
		If strStatus = "2" Then
			Call DisplayMsgBox("187733", vbInformation, "", "", I_MKSCRIPT)
			rs0.Close
			Set rs0 = Nothing
			rs1.Close
			Set rs1 = Nothing							
			Set ADF = Nothing			         
			Response.End
		ElseIF  strStatus = "3" Or strStatus = "" Then
			Call DisplayMsgBox("187738", vbInformation, "", "", I_MKSCRIPT)
			rs0.Close
			Set rs0 = Nothing
			rs1.Close
			Set rs1 = Nothing				
			Set ADF = Nothing			             
			Response.End
		End If
    Else		
		Call DisplayMsgBox("125000", vbInformation, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing	
		rs1.Close
		Set rs1 = Nothing	
		Set ADF = Nothing				
		Response.End		
	End If
    
	rs0.Close
	Set rs0 = Nothing
	rs1.Close
	Set rs1 = Nothing	
	Set ADF = Nothing
	
'--------------------------------------------------------------------------------------------------------------------	
    Dim pPP2G102
    Dim I1_plant_cd, I2_mrp_parameter, I3_select_char
    
    Const P206_I2_plant_cd = 0    
    Const P206_I2_safe_flg = 1
    Const P206_I2_inv_flg = 2
    Const P206_I2_idep_flg = 3
    Const P206_I2_forward = 4
    Const P206_I2_mpsscope = 5

    ReDim I2_mrp_parameter(P206_I2_mpsscope)
    
	I1_plant_cd			= UCase(Trim(Request("txtPlantCd")))
	
	I2_mrp_parameter(P206_I2_plant_cd)	= UCase(Trim(Request("txtPlantCd")))
	I2_mrp_parameter(P206_I2_safe_flg)	= "Y"
	I2_mrp_parameter(P206_I2_inv_flg)	= "P"
	I2_mrp_parameter(P206_I2_idep_flg)	= "M" 
	I2_mrp_parameter(P206_I2_forward)	= UCase(Trim(Request("txtMPSHistoryNo")))
	I2_mrp_parameter(P206_I2_mpsscope)	= "" 

	I3_select_char = "M"

    Set pPP2G102 = Server.CreateObject("PP2G102.cPCnfmMrpSvr")
	    
    If CheckSYSTEMError(Err,True) = True Then
		Set pPP2G102 = Nothing		
		Response.End
	End If
	
	Call pPP2G102.P_CONFIRM_MRP_SRV(gStrGlobalCollection, I1_plant_cd, I2_mrp_parameter, I3_select_char, "")

	If CheckSYSTEMError(Err, True) = True Then
		Set pPP2G102 = Nothing
		Response.End
	End If

	Set pPP2G102 = Nothing      		
	Call DisplayMsgBox("183114", vbInformation, "", "", I_MKSCRIPT)
%>
