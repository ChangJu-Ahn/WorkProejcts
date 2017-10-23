<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%	Call LoadBasisGlobalInf
	Call LoadInfTB19029B("Q", "P", "NOCOOKIE", "BB")
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: 
'*  4. Program Name			: MRP Explosion
'*  5. Program Desc			: query Plant
'*  6. Comproxy List		: 
'*  7. Modified date(First)	:
'*  8. Modified date(Last) 	: 
'*  9. Modifier (First)		:
'* 10. Modifier (Last)		: Jung Yu Kyung
'* 11. Comment				:
'**********************************************************************************************

Call HideStatusWnd

On Error Resume Next

'--------------------------------------------------------------------------------------------------------------------
Dim ADF
Dim strRetMsg
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1

Dim lgStrPrevKey	' 이전 값 
Dim lGetSvrDate

lGetSvrDate = GetSvrDate

IF Request("txtPlantCd") = "" THEN
   Response.End
END IF

	lgStrPrevKey = Request("lgStrPrevKey")
	
	Redim UNISqlId(1)
	Redim UNIValue(1, 0)
	
	UNISqlId(0) = "184000saa"
	UNISqlId(1) = "185000sac"

	UNIValue(0, 0) = " " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & ""
	UNIValue(1, 0) = " " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & ""
		
	UNILock = DISCONNREAD :	UNIFlag = "1"

    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)

	If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("125000", vbInformation, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing		
		
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "With parent.frm1" & vbCrLf	
				Response.Write ".txtPlantNm.value = """"" & vbCrLf
				Response.Write ".txtMRPHisNo.value = """"" & vbCrLf
				Response.Write ".txtFixExecToDt.text = """"" & vbCrLf
				Response.Write ".txtPlanExecToDt.text = """"" & vbCrLf
				Response.Write ".txtPlantCd.focus" & vbCrLf	
			Response.Write "End With" & vbCrLf
		Response.Write "</Script>" & vbCrLf

		Response.End
	End If
	
	Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "With parent.frm1" & vbCrLf
			Response.Write ".txtMRPHisNo.value = """"" & vbCrLf
			Response.Write ".txtFixExecToDt.text = """ & UNIDateAdd("d",UniCInt(rs0("ptf_for_mrp"), 0), UNIDateClientFormat(lGetSvrDate), gDateFormat) & """" & vbCrLf	'☜: 화면 처리 ASP 를 지칭함 
			Response.Write ".txtPlanExecToDt.text = """ & UNIDateAdd("d",UniCInt(rs0("plan_hrzn"), 0), UNIDateClientFormat(lGetSvrDate), gDateFormat) & """" & vbCrLf	'☜: 화면 처리 ASP 를 지칭함 
			Response.Write ".txtPlantNm.value = """ & ConvSPChars(rs0("plant_nm")) & """" & vbCrLf
			Response.Write "parent.lgInvCloseDt = """ & UNIDateClientFormat(rs0("inv_cls_dt")) & """" & vbCrLf		

			If rs1.EOF And rs1.BOF Then
				Response.Write ".txtErrorQty.text = """"" & vbCrLf
			Else
				Response.Write ".txtErrorQty.text = """ & ConvSPChars(rs1("error_qty")) & """" & vbCrLf	
			End If
			
			Call .DBQueryOk
		Response.Write "End With" & vbCrLf
	Response.Write "</Script>" & vbCrLf

	rs0.Close
	Set rs0 = Nothing	
	rs1.Close
	Set rs1 = Nothing
    Set ADF = Nothing
%>
