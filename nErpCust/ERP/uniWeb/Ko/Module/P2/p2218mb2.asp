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
'*  3. Program ID			: p2218mb2.asp
'*  4. Program Name			: MPSȮ����� 
'*  5. Program Desc			: Cancel MPS(save)
'*  6. Comproxy List		: PP2G102.cPCnfmMrpSvr
'*  7. Modified date(First)	: 2002/04/24
'*  8. Modified date(Last) 	: 2002/06/19
'*  9. Modifier (First)		: Jung Yu Kyung
'* 10. Modifier (Last)		: Jung Yu Kyung
'* 11. Comment				:
'**********************************************************************************************-->

<% Call LoadBasisGlobalInf
   Call LoadInfTB19029B("I", "*", "NOCOOKIE", "MB")

Call HideStatusWnd

On Error Resume Next
'--------------------------------------------------------------------------------------------------------------------
Dim ADF
Dim strRetMsg
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1

Dim strStatus
Dim strStatus_mrp
	
	Redim UNISqlId(1)
	Redim UNIValue(1, 1)
	
	UNISqlId(0) = "184000sab"
	UNISqlId(1) = "185000saa"

	UNIValue(0, 0) = "'" & Ucase(Trim(Request("txtPlantCd"))) & "'"
	UNIValue(0, 1) = "'" & Ucase(Trim(Request("txtPlantCd"))) & "'"

	UNIValue(1, 0) = "'" & Ucase(Trim(Request("txtPlantCd"))) & "'"
	UNIValue(1, 1) = "'" & Ucase(Trim(Request("txtPlantCd"))) & "'"
		
	UNILock = DISCONNREAD :	UNIFlag = "1"

    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)  

    IF NOT(rs0.EOF And rs0.BOF) Then
		strStatus = rs0("status")
		If strStatus = "1" Then
		    Call DisplayMsgBox("187730", vbInformation, "", "", I_MKSCRIPT)
		    rs0.Close
			rs1.Close
			Set rs0 = Nothing
			Set rs1 = Nothing
			Set ADF = Nothing         
			Response.End
		End If
	END IF
    
    IF NOT(rs1.EOF And rs1.BOF) Then
        strStatus_mrp = rs1("status")
		If strStatus_mrp = "1" Or strStatus_mrp = "6" Then
			Call DisplayMsgBox("187731", vbInformation, "", "", I_MKSCRIPT)
		    rs0.Close
			rs1.Close
			Set rs0 = Nothing
			Set rs1 = Nothing
			Set ADF = Nothing         
			Response.End
		ElseIF  strStatus_mrp = "2" Or strStatus_mrp = "3" Then
			Call DisplayMsgBox("187732", vbInformation, "", "", I_MKSCRIPT)
		    rs0.Close
			rs1.Close
			Set rs0 = Nothing
			Set rs1 = Nothing
			Set ADF = Nothing         
			Response.End
		End If
	END IF   
    
	rs0.Close
	Set rs0 = Nothing
	rs1.Close
	Set rs1 = Nothing	
'---------------------------------------------------------------------------------------------------------------------	
Dim pPP2G102
Dim I2_mrp_parameter
Dim I1_plant_cd, I3_select_char
Dim strSpread
    
Const P206_I2_plant_cd = 0    
Const P206_I2_safe_flg = 1
Const P206_I2_inv_flg = 2
Const P206_I2_idep_flg = 3
Const P206_I2_forward = 4
Const P206_I2_mpsscope = 5

 	Err.Clear
 	
 	ReDim I2_mrp_parameter(P206_I2_mpsscope)
    
	I1_plant_cd			= UCase(Trim(Request("txtPlantCd")))
	
	I2_mrp_parameter(P206_I2_plant_cd)	= UCase(Trim(Request("txtPlantCd")))
	I2_mrp_parameter(P206_I2_safe_flg)	= "Y"
	I2_mrp_parameter(P206_I2_inv_flg) = "C"
	I2_mrp_parameter(P206_I2_idep_flg)	= "M" 
	I2_mrp_parameter(P206_I2_forward)	= ""
	I2_mrp_parameter(P206_I2_mpsscope)	= "" 

	I3_select_char = "S"
	strSpread = Request("txtSpread")

	Set pPP2G102 = Server.CreateObject("PP2G102.cPCnfmMrpSvr")
		    
	If CheckSYSTEMError(Err,True) = True Then
		Set pPP2G102 = Nothing		
		Response.End
	End If
	
	Call pPP2G102.P_CONFIRM_MRP_SRV(gStrGlobalCollection, I1_plant_cd, I2_mrp_parameter, I3_select_char, strSpread)

	If CheckSYSTEMError(Err, True) = True Then
		Set pPP2G102 = Nothing
		Response.End
	End If
	
	Set pPP2G102 = Nothing      															
%>
<Script Language=vbscript>
	With parent																		
		.DbSaveOk
	End With
</Script>