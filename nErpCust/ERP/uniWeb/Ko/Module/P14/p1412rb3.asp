<%@LANGUAGE = VBScript%>
<%'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1412rb3.asp
'*  4. Program Name         : Look Up Ecn Info
'*  5. Program Desc         :
'*  6. Comproxy List        : + 
'*  7. Modified date(First) : 2003/03/27
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Ryu Sung Won
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'*************************************************************************************************%>

<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("I", "*", "NOCOOKIE", "RB")

Dim ADF												'ActiveX Data Factory ���� �������� 
Dim strRetMsg										'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag			'DBAgent Parameter ���� 
Dim rs0, rs1, rs2, rs3								'DBAgent Parameter ���� 
Dim strQryMode
Dim strMode																	'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim strEcnNo
Dim strReasonCd
Dim strReasonNm
Dim strEcnDesc
Dim blnResult

Call HideStatusWnd

On Error Resume Next
Err.Clear

	blnResult = False
	strMode = Request("txtMode")												'�� : ���� ���¸� ���� 

	Redim UNISqlId(0)
	Redim UNIValue(0, 0)

	UNISqlId(0) = "p1412mb3a"
	
	UNIValue(0, 0) = FilterVar(Request("txtEcnNo"), "''", "S")

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

	If (rs0.EOF and rs0.BOF)Then
		%>
		<Script Language=vbscript>
			parent.frm1.txtEcnDesc.value = ""
			parent.frm1.txtReasonCd.value = ""
			parent.frm1.txtReasonNm.value = ""
		</Script>	
		<%
		rs0.Close
		Set rs0 = Nothing
		blnResult = False
	Else
		%>
		<Script Language=vbscript>
			parent.frm1.txtEcnDesc.value = "<%=rs0("ECN_DESC")%>"
			parent.frm1.txtReasonCd.value = "<%=UCase(Trim(rs0("REASON_CD")))%>"
			parent.frm1.txtReasonNm.value = "<%=rs0("REASON_NM")%>"
		</Script>	
		<%
		rs0.Close
		Set rs0 = Nothing
		blnResult = True
	End If
%>
	<Script Language=vbscript>
		Call parent.LookUpEcnInfoOk("<%=blnResult%>")
	</Script>
<%
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
