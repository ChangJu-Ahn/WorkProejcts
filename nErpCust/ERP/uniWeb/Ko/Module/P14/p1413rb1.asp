<%@LANGUAGE = VBScript%>
<%'*******************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1413rb1.asp
'*  4. Program Name         : BOM Mass Replacement
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/03/14
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : RYU SUNG WON
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'********************************************************************************************%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call LoadBasisGlobalInf
Call LoadInfTB19029B("I", "P", "NOCOOKIE", "RB")

Dim ADF										'ActiveX Data Factory ���� �������� 
Dim strRetMsg								'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter ���� 
Dim rs0, rs1, rs2, rs3, rs4					'DBAgent Parameter ���� 
Dim strQryMode

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Call HideStatusWnd

Dim strPlantCd
Dim strItemCd
Dim strBomType

On Error Resume Next
Err.Clear																	'��: Protect system from crashing

	strQryMode = Request("lgIntFlgMode")
	
	Redim UNISqlId(3)
	Redim UNIValue(3, 1)

	UNISqlId(0) = "180000saa"
	UNISqlId(1) = "180000saq"
	UNISqlId(2) = "AMINORNM"
	UNISqlId(3) = "p1412mb1b"	'bom history flg

	UNIValue(0, 0) = FilterVar(Request("txtPlantCd"), "''", "S")
	UNIValue(1, 0) = FilterVar(Request("txtItemCd"), "''", "S")
	UNIValue(1, 1) = FilterVar(Request("txtPlantCd"), "''", "S")
	UNIValue(2, 0) = FilterVar("P1401","''","S")
	UNIValue(2, 1) = FilterVar(Request("txtBomType"),"''","S")
	UNIValue(3, 0) = FilterVar(Request("txtPlantCd"), "''", "S")
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1, rs2, rs3, rs4)

	' Plant �� Display      
	If (rs1.EOF And rs1.BOF) Then
		Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
		rs1.Close
		Set rs1 = Nothing
		Set ADF = Nothing
		Response.End
	Else
		rs1.Close
		Set rs1 = Nothing
	End If

	' ǰ��� Display
	IF Request("txtItemCd") <> "" Then
		If (rs2.EOF And rs2.BOF) Then
			Call DisplayMsgBox("122700", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
				parent.frm1.hItemCd.value = ""
				parent.frm1.txtItemNm.value = ""
				parent.frm1.txtAcct.value = ""
				parent.frm1.txtProcurType.value = ""
				parent.frm1.txtSpec.value = ""
				'parent.frm1.txtValidFromDt.Text = ""
				parent.frm1.txtValidToDt.Text = ""
				parent.frm1.txtItemCd.focus
			</Script>
			<%
			rs2.Close
			Set rs2 = Nothing
			Set ADF = Nothing
			Response.End
		Else
			%>
			<Script Language=vbscript>
				parent.frm1.hItemCd.value = "<%=ConvSPChars(rs2("ITEM_CD"))%>"
				parent.frm1.txtItemNm.value = "<%=ConvSPChars(rs2("ITEM_NM"))%>"
				parent.frm1.txtAcct.value = "<%=ConvSPChars(rs2("ITEM_ACCT_NM"))%>"
				parent.frm1.txtProcurType.value = "<%=ConvSPChars(rs2("PROCUR_TYPE_NM"))%>"
				parent.frm1.txtSpec.value = "<%=ConvSPChars(rs2("SPEC"))%>"
				'parent.frm1.txtValidFromDt.Text = "<%=UniDateClientFormat(rs2("PLANT_VALID_FROM_DT"))%>"
				parent.frm1.txtValidToDt.Text = "<%=UniDateClientFormat(rs2("PLANT_VALID_TO_DT"))%>"	'2003-09-13
			</Script>	
			<%
			rs2.Close
			Set rs2 = Nothing
		End If
	End IF
	
	' BOM No Check
	IF Request("txtBomType") <> "" Then
		If (rs3.EOF And rs3.BOF) Then
			Call DisplayMsgBox("182622", vbOKOnly, "", "", I_MKSCRIPT)
			rs3.Close
			Set rs3 = Nothing
			Set ADF = Nothing
			Response.End
		Else
			rs3.Close
			Set rs3 = Nothing
		End If
	End IF

	' BOM HISTORY FLG - P_Plant_Configuration
	If (rs4.EOF And rs4.BOF) Then
		%>
		<Script Language=vbscript>
			parent.frm1.hBomHistoryFlg.value = "N"
		</Script>
		<%
		rs4.Close
		Set rs4 = Nothing
		Set ADF = Nothing
		Response.End
	Else
		%>
		<Script Language=vbscript>
			parent.frm1.hBomHistoryFlg.value = "<%=Trim(rs4("BOM_HISTORY_FLG"))%>"
		</Script>
		<%
		rs4.Close
		Set rs4 = Nothing
	End If
	
	Set ADF = Nothing
	
	Response.Write "<Script Language=vbscript>" & vbCrLf
	Response.Write "parent.DbQueryOk()" & vbCrLf
	Response.Write "</Script>" & vbCrLf

%>
