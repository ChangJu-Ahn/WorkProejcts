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
'*  3. Program ID			: p2217mb2.asp
'*  4. Program Name			: MPS°ü¸® 
'*  5. Program Desc			: MPS Save
'*  6. Comproxy List		: PP2G104.cPMngMps
'*  7. Modified date(First)	:
'*  8. Modified date(Last) 	: 2002/06/21
'*  9. Modifier (First)		: Lee Hyun Jae
'* 10. Modifier (Last)		: Jung Yu Kyung
'* 11. Comment				:
'**********************************************************************************************-->
<% Call LoadBasisGlobalInf
   Call LoadInfTB19029B("I", "*", "NOCOOKIE", "MB")


Call HideStatusWnd

On Error Resume Next

Dim ADF
Dim strRetMsg
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1

Dim strStatus
Dim strStatus_mrp

	lgStrPrevKey = Request("lgStrPrevKey")
	
	Redim UNISqlId(1)
	Redim UNIValue(1, 1)
	
	UNISqlId(0) = "184000sab"
	UNISqlId(1) = "185000saa"

	UNIValue(0, 0) = FilterVar(Trim(Request("txtPlantCd"))	, "''", "S")
	UNIValue(0, 1) = FilterVar(Trim(Request("txtPlantCd"))	, "''", "S")

	UNIValue(1, 0) = FilterVar(Trim(Request("txtPlantCd"))	, "''", "S")
	UNIValue(1, 1) = FilterVar(Trim(Request("txtPlantCd"))	, "''", "S")
		
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
		If strStatus_mrp = "1" Then	
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
Dim pPP2G104
Dim I1_plant_cd
Dim strSpread

 	Err.Clear
 	
	I1_plant_cd = UCase(Trim(Request("txtPlantCd")))
	strSpread = Request("txtSpread")

	Set pPP2G104 = Server.CreateObject("PP2G104.cPMngMps")
		    
	If CheckSYSTEMError(Err,True) = True Then
		Set pPP2G104 = Nothing		
		Response.End
	End If
	
	Call pPP2G104.P_MANAGE_MPS(gStrGlobalCollection, I1_plant_cd, strSpread)

	If CheckSYSTEMError(Err, True) = True Then
		Set pPP2G104 = Nothing
		Response.End
	End If
	
	Set pPP2G104 = Nothing      			
    
%>
<Script Language=vbscript>
	With parent
		.DbSaveOk
	End With
</Script>
