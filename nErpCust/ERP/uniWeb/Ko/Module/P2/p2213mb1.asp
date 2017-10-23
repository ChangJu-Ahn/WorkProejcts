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
'*  3. Program ID			: p2213mb1.asp
'*  4. Program Name			: Approve MPS
'*  5. Program Desc			: plant query
'*  6. Comproxy List		: 
'*  7. Modified date(First)	:
'*  8. Modified date(Last) 	: 
'*  9. Modifier (First)		:
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

	lgStrPrevKey = Request("lgStrPrevKey")
	
	Redim UNISqlId(1)
	Redim UNIValue(1, 1)
	
	UNISqlId(0) = "184000sab"
	UNISqlId(1) = "184000saa"

	UNIValue(0, 0) = FilterVar(Trim(Request("txtPlantCd"))	, "''", "S")
	UNIValue(0, 1) = FilterVar(Trim(Request("txtPlantCd"))	, "''", "S")
	
	UNIValue(1, 0) = FilterVar(Trim(Request("txtPlantCd"))	, "''", "S")
	
	UNILock = DISCONNREAD :	UNIFlag = "1"

    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)

	If rs1.EOF And rs1.BOF Then
		Call DisplayMsgBox("125000", vbInformation, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing	
		rs1.Close
		Set rs1 = Nothing	
		
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtPlantNm.value = """"" & vbCrLf
			Response.Write "parent.frm1.txtPlantCd.focus" & vbCrLf
		Response.Write "</Script>" & vbCrLf			
		
		Response.End
	End If
%>

<Script Language=vbscript>
With parent
	.frm1.txtPlantNm.value		= "<%= ConvSPChars(rs1("plant_nm")) %>"
    .frm1.txtMPSHistoryNo.value = "<%= ConvSPChars(rs0("mps_history_no")) %>"
    .frm1.txtDTF.text			= "<%=UNIDateAdd("d",UniCInt(rs0("mps_dtf"),0), UniDateClientFormat(rs0("start_dt")),gDateFormat)%>"
    .frm1.txtPTF.text			= "<%=UNIDateAdd("d",UniCInt(rs0("mps_ptf"),0), UniDateClientFormat(rs0("start_dt")),gDateFormat)%>"
    .frm1.txtPlanDt.text		= "<%=UniDateClientFormat(rs0("plan_dt"))%>"
    .frm1.txtStartDt.text		= "<%=UniDateClientFormat(rs0("start_dt"))%>"
    IF "<%= UCase(rs0("max_order_flg")) %>" = "Y" Then
        .frm1.rdoMaxFlg1.Checked = True
    Else
        .frm1.rdoMaxFlg2.Checked = True
    End If

    IF "<%= UCase(rs0("min_order_flg")) %>" = "Y" Then
        .frm1.rdoMinFlg1.Checked = True
    Else
        .frm1.rdoMinFlg2.Checked = True
    End If
    
    IF "<%= UCase(rs0("round_flg")) %>" = "Y" Then
        .frm1.rdoRoundFlg1.Checked = True
    Else
        .frm1.rdoRoundFlg2.Checked = True
    End If
    
    IF "<%= UCase(rs0("inv_flg")) %>" = "Y" Then
        .frm1.rdoAvailInvFlg1.Checked = True
    Else
        .frm1.rdoAvailInvFlg2.Checked = True
    End If
    
    IF "<%= UCase(rs0("ss_flg")) %>" = "Y" Then
        .frm1.rdoSafeInvFlg1.Checked = True
    Else
        .frm1.rdoSafeInvFlg2.Checked = True
    End If
    
    IF "<%= UCase(rs0("start_flg")) %>" = "D" Then
        .frm1.rdoStartDtFlg1.Checked = True
    Else
        .frm1.rdoStartDtFlg2.Checked = True
    End If
    
End With
</Script>
<%
	rs0.Close
	rs1.Close
	Set rs0 = Nothing
	Set rs1 = Nothing
	Set ADF = Nothing
%>
