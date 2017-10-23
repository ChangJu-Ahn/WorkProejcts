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
'*  3. Program ID			: p2213mb3.asp
'*  4. Program Name			: 
'*  5. Program Desc			: Cancel MPS
'*  6. Comproxy List		: PP2G103.cPCnclMpsSvr
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
			Call DisplayMsgBox("187735", vbInformation, "", "", I_MKSCRIPT)
			rs0.Close
			Set rs0 = Nothing
			rs1.Close
			Set rs1 = Nothing			
			Set ADF = Nothing			             
			Response.End
		End If
    ELSE		
		Call DisplayMsgBox("125000", vbInformation, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing	
		rs1.Close
		Set rs1 = Nothing	
		Set ADF = Nothing				
		Response.End		
	END IF		
    
	rs0.Close
	Set rs0 = Nothing
	rs1.Close
	Set rs1 = Nothing		
	Set ADF = Nothing

'--------------------------------------------------------------------------------------------------------------------									
	Dim pPP2G103
    Dim I1_plant_cd, I2_p_mps_history_no
   
	I1_plant_cd			= UCase(Trim(Request("txtPlantCd")))
	I2_p_mps_history_no	= UCase(Trim(Request("txtMPSHistoryNo")))
	
    Set pPP2G103 = Server.CreateObject("PP2G103.cPCnclMpsSvr")
	
    If CheckSYSTEMError(Err,True) = True Then
		Set pPP2G103 = Nothing		
		Response.End
	End If
	
	Call pPP2G103.P_CANCEL_MPS(gStrGlobalCollection, I1_plant_cd, I2_p_mps_history_no)

	If CheckSYSTEMError(Err, True) = True Then
		Set pPP2G103 = Nothing
		Response.End
	End If

	Set pPP2G103 = Nothing      		
	
	Call DisplayMsgBox("183114", vbInformation, "", "", I_MKSCRIPT)
%>
