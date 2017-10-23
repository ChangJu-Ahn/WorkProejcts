<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!--
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: p2345mb2.asp
'*  4. Program Name			: MRP일괄전환 
'*  5. Program Desc			: MRP Conversion
'*  6. Comproxy List		: PP2G102.cPCnfmMrpSvr
'*  7. Modified date(First)	:
'*  8. Modified date(Last) 	: 
'*  9. Modifier (First)     : Lee Hyun Jae
'* 10. Modifier (Last)      : Jung Yu Kyung
'* 11. Comment				:
'**********************************************************************************************-->
<% 

Call LoadBasisGlobalInf
Call HideStatusWnd

On Error Resume Next									

'--------------------------------------------------------------------------------------------------------------------
Dim ADF	
Dim strRetMsg
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1

Dim strStatus

	IF Request("txtMRPHisNo") = "" THEN
		Call DisplayMsgBox("187742", vbInformation, "", "", I_MKSCRIPT)
		Response.End
	END IF
	

	Redim UNISqlId(0)
	Redim UNIValue(0, 1)
	
	UNISqlId(0) = "185000saa"

	UNIValue(0, 0) = " " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & ""
	UNIValue(0, 1) = " " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & ""	
		
	UNILock = DISCONNREAD :	UNIFlag = "1"

    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

    strStatus = rs0("status")
		
    If strStatus = "1" Then
		Call DisplayMsgBox("187740", vbInformation, "", "", I_MKSCRIPT)
		rs0.Close		
		Set rs0 = Nothing		
		Set ADF = Nothing 
		Response.End
	ElseIF  strStatus = "4" Or strStatus = "5" Or strStatus = "6" Or strStatus = "" Then
		Call DisplayMsgBox("187734", vbInformation, "", "", I_MKSCRIPT)
		rs0.Close		
		Set rs0 = Nothing		
		Set ADF = Nothing   
		Response.End
    End If
	rs0.Close		
	Set rs0 = Nothing		
	Set ADF = Nothing												

'--------------------------------------------------------------------------------------------------------------------									
	
	Dim pPP2G102
	Dim I2_mrp_parameter
	Dim I1_plant_cd, I3_select_char
	    
	Const P206_I2_plant_cd	= 0    
	Const P206_I2_safe_flg	= 1
	Const P206_I2_inv_flg	= 2
	Const P206_I2_idep_flg	= 3
	Const P206_I2_forward	= 4
	Const P206_I2_mpsscope	= 5

 	Err.Clear

    ReDim I2_mrp_parameter(P206_I2_mpsscope)
    
	I1_plant_cd			= UCase(Request("txtPlantCd"))
	
	I2_mrp_parameter(P206_I2_plant_cd)	= UCase(Request("txtPlantCd"))
	I2_mrp_parameter(P206_I2_safe_flg)	= "Y"
	I2_mrp_parameter(P206_I2_inv_flg)	= "M"
	I2_mrp_parameter(P206_I2_idep_flg)	= "M" 
	I2_mrp_parameter(P206_I2_forward)	= UCase(Request("txtMRPHisNo"))
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

'------------------------------------------------------------------------------------------------------------------
	
	Redim UNISqlId(1)
	Redim UNIValue(1, 1)
	
	UNISqlId(0) = "185000saa"
	UNISqlId(1) = "185000sac"

	UNIValue(0, 0) = " " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & ""
	UNIValue(0, 1) = " " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & ""
	
	UNIValue(1, 0) = " " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & ""
		
	UNILock = DISCONNREAD :	UNIFlag = "1"

    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)
    
    Response.Write "<Script Language=VBScript>" & vbCrLf
	Response.Write "With parent" & vbCrLf
	Response.Write "	.frm1.txtConvQty.value = """ & rs0("conv_order_qty") & """" & vbCrLf
	Response.Write "	.frm1.txtErrQty.value = """ & rs0("error_qty") & """" & vbCrLf
			
	IF 	rs0("STATUS") = "1" THEN
		Response.Write "	.frm1.txtStatus.value = ""전개""" & vbCrLf
	ELSEIF rs0("STATUS") = "2" THEN
		Response.Write "	.frm1.txtStatus.value = ""승인""" & vbCrLf
	ELSEIF rs0("STATUS") = "3" THEN
		Response.Write "	.frm1.txtStatus.value = ""부분전환""" & vbCrLf
	ELSEIF rs0("STATUS") = "4" THEN
		Response.Write "	.frm1.txtStatus.value = ""전환완료""" & vbCrLf
	ELSEIF rs0("STATUS") = "5" THEN
		Response.Write "	.frm1.txtStatus.value = ""전개취소""" & vbCrLf
	ELSE
		Response.Write "	.frm1.txtStatus.value = ""승인취소""" & vbCrLf
	END IF
	
	Response.Write "End With" & vbCrLf
	Response.Write "</Script>" & vbCrLf
 
   
	rs0.Close
	Set rs0 = Nothing
    Set ADF = Nothing
    Call DisplayMsgBox("183114", vbInformation, "", "", I_MKSCRIPT)
%>
