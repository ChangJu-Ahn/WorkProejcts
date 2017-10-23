<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!--'*****************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: 
'*  4. Program Name			: MRP 승인/전개취소 
'*  5. Program Desc			: Plant query
'*  6. Comproxy List		: 
'*  7. Modified date(First)	:
'*  8. Modified date(Last) 	: 
'*  9. Modifier (First)		:
'* 10. Modifier (Last)		: Jung Yu Kyung
'* 11. Comment				:
'********************************************************************************************-->
<% 

Call LoadBasisGlobalInf
Call HideStatusWnd

On Error Resume Next	
'--------------------------------------------------------------------------------------------------------------------
Dim ADF
Dim strRetMsg
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1

	
	Redim UNISqlId(1)
	Redim UNIValue(1, 1)
	
	UNISqlId(0) = "185000saa"
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
		Set ADF = Nothing		
		
		Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "	parent.frm1.txtPlantNm.value = """"" & vbCrLf
		Response.Write "	parent.frm1.txtPlantCd.focus" & vbCrLf
		Response.Write "</Script>" & vbCrLf

		Response.End
	End If    

%>
<Script Language=vbscript>
With parent
	.frm1.txtPlantNm.value		= "<%=ConvSPChars(rs1("plant_nm"))%>" 
    .frm1.txtMRPHisNo.value		= "<%=ConvSPChars(rs0("run_no"))%>"
    .frm1.txtFixExecFromDt.text = "<%=UNIDateClientFormat(rs0("from_during_dt"))%>"
    .frm1.txtFixExecToDt.text	= "<%=UNIDateClientFormat(rs0("firm_during_dt"))%>"
    .frm1.txtPlanExecToDt.text	= "<%=UNIDateClientFormat(rs0("to_during_dt"))%>"
    .frm1.txtStartDt.text		= "<%=UNIDateClientFormat(rs0("start_dt"))%>"
    .frm1.txtOrderQty.Text		= "<%=rs0("order_qty")%>"    
    
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
    
<%		IF 	rs0("STATUS") = "1" THEN%>
			.frm1.txtStatus.value = "전개"
<% 		ELSEIF rs0("STATUS") = "2" THEN%>
			.frm1.txtStatus.value = "승인"
<% 		ELSEIF rs0("STATUS") = "3" THEN%>
			.frm1.txtStatus.value = "부분전환"
<% 		ELSEIF rs0("STATUS") = "4" THEN%>
			.frm1.txtStatus.value = "전환완료"
<% 		ELSEIF rs0("STATUS") = "5" THEN%>
			.frm1.txtStatus.value = "전개취소"
<% 		ELSE%>
			.frm1.txtStatus.value = "승인취소"
<%		END IF%>		
End With
</Script>
<%
	rs0.Close
	Set rs0 = Nothing
	rs1.Close
	Set rs1 = Nothing	
	Set ADF = Nothing
%>
