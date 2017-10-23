<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->

<% Call LoadBasisGlobalInf
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: p2210mb1.asp
'*  4. Program Name			: MPS일괄생성 
'*  5. Program Desc			: Plant query
'*  6. Comproxy List		: 
'*  7. Modified date(First)	:
'*  8. Modified date(Last) 	: 
'*  9. Modifier (First)		:
'* 10. Modifier (Last)		: Jung Yu Kyoung
'* 11. Comment				:
'**********************************************************************************************
Call HideStatusWnd

On Error Resume Next

'--------------------------------------------------------------------------------------------------------------------
Dim ADF
Dim strRetMsg
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0

	Redim UNISqlId(0)
	Redim UNIValue(0, 0)
	
	UNISqlId(0) = "184000saa"

	UNIValue(0, 0) =FilterVar(Trim(Request("txtPlantCd"))	, "''", "S")
		
	UNILock = DISCONNREAD :	UNIFlag = "1"

    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

	If rs0.EOF And rs0.BOF Then
		
		Call DisplayMsgBox("125000", vbInformation, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing	
		
		Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "With parent.frm1" & vbCrLf
		Response.Write "	.txtPlantNm.value = """"" & vbCrLf
		Response.Write "	.txtMPSHistoryNo.value = """"" & vbCrLf
'20080304::hanc		Response.Write "	.txtPlantCd.focus" & vbCrLf
		Response.Write "	.txtPlanDt.text = """"" & vbCrLf
		Response.Write "	.txtDTF.text = """"" & vbCrLf
		Response.Write "	.txtPTF.text = """"" & vbCrLf
		Response.Write "End With" & vbCrLf
		Response.Write "</Script>" & vbCrLf
		
		Response.End 
	End If
	
	Response.Write "<Script Language=VBScript>" & vbCrLf
	Response.Write "With parent.frm1" & vbCrLf
	Response.Write "	.txtPlantNm.value = """ & ConvSPChars(rs0("plant_nm")) & """" & vbCrLf
	Response.Write "	.txtMPSHistoryNo.value = """"" & vbCrLf
'20080304::hanc	Response.Write "	.txtPlanDt.text = """ & UniDateClientFormat(UNIDateAdd("d",UniCInt(rs0("plan_hrzn"), 0),GetSvrDate, gServerDateFormat)) & """" & vbCrLf	'☜: 화면 처리 ASP 를 지칭함 
	Response.Write "	.txtDTF.text = """ & UniDateClientFormat(UNIDateAdd("d",UniCInt(rs0("dtf_for_mps"), 0),GetSvrDate, gServerDateFormat)) & """" & vbCrLf	'☜: 화면 처리 ASP 를 지칭함 
	Response.Write "	.txtPTF.text = """ & UniDateClientFormat(UNIDateAdd("d",UniCInt(rs0("ptf_for_mps"), 0),GetSvrDate, gServerDateFormat)) & """" & vbCrLf	'☜: 화면 처리 ASP 를 지칭함 
    Response.Write  "    Parent.DBQueryOk   " & vbCr      
	Response.Write "End With" & vbCrLf
	Response.Write "</Script>" & vbCrLf	
	
	rs0.Close
	Set rs0 = Nothing
    Set ADF = Nothing
%>
