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
'*  3. Program ID			: p2345mb3.asp
'*  4. Program Name			: MRP일괄전환 
'*  5. Program Desc			: cancel approve MRP
'*  6. Comproxy List		: PP3G103.cPCnclApproveMrpSvr
'*  7. Modified date(First)	: 2000/05/2
'*  8. Modified date(Last)  : 2002/06/18
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : Jung Yu Kyung
'* 11. Comment		:
'**********************************************************************************************-->
<% 

Call LoadBasisGlobalInf
Call HideStatusWnd

On Error Resume Next

'--------------------------------------------------------------------------------------------------------------------
Dim ADF
Dim strRetMsg
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0

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
		Call DisplayMsgBox("187741", vbInformation, "", "", I_MKSCRIPT)
		rs0.Close		
		Set rs0 = Nothing		
		Set ADF = Nothing
		Response.End
    End If
    
	rs0.Close		
	Set rs0 = Nothing		   
	Set ADF = Nothing
'--------------------------------------------------------------------------------------------------------------------									
    Dim pPP3G103
    Dim I1_plant_cd, I2_run_no

    Err.Clear
    
    I1_plant_cd = UCase(Request("txtPlantCd"))    
    I2_run_no = UCase(Request("txtMRPHisNo"))
    
	Set pPP3G103 = Server.CreateObject("PP3G103.cPCnclApproveMrpSvr")  
	
	If CheckSYSTEMError(Err,True) = True Then
		Set pPP3G103 = Nothing		
		Response.End
	End If
	
	Call pPP3G103.P_CANCEL_APPROVE_MRP(gStrGlobalCollection, I1_plant_cd, I2_run_no)

	If CheckSYSTEMError(Err, True) = True Then
		Set pPP3G103 = Nothing
		Response.End
	End If

	Set pPP3G103 = Nothing
'------------------------------------------------------------------------------------------------------------------
	
	Redim UNISqlId(0)
	Redim UNIValue(0, 1)
	
	UNISqlId(0) = "185000saa"

	UNIValue(0, 0) = " " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & ""
	UNIValue(0, 1) = " " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & ""
		
	UNILock = DISCONNREAD :	UNIFlag = "1"

    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    
	Response.Write "<Script Language=VBScript>" & vbCrLf
	Response.Write "With parent" & vbCrLf
			
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
