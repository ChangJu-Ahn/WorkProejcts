<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4712mb0.asp
'*  4. Program Name         : List production order infomation	
'*  5. Program Desc         : 
'*  6. Modified date(First) : 2001.12.03
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Jeon, Jaehyun
'*  9. Modifier (Last)      : 
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call LoadBasisGlobalInf
Call loadInfTB19029B("I", "*", "NOCOOKIE","MB")

On Error Resume Next

Dim ADF										'ActiveX Data Factory ���� �������� 
Dim strRetMsg								'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter ���� 
Dim rs0, rs1, rs2								'DBAgent Parameter ���� 

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================

Call HideStatusWnd

Dim StrProdOrderNo
Dim StrOprNo

Err.Clear																	'��: Protect system from crashing

	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	Redim UNISqlId(1)
	Redim UNIValue(1, 0)

	UNISqlId(0) = "180000saa"
	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1)
    Set ADF = Nothing
	
	Response.Write "<Script Language=vbscript>" & vbcr
	Response.Write "	parent.frm1.txtPlantNm.value = """"" & vbcr
	Response.Write "</Script>" & vbcr

	' Plant �� Display      
	If (rs1.EOF And rs1.BOF) Then
		rs1.Close
		Set rs1 = Nothing
		strFlag = "ERROR_PLANT"
		Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
		Response.Write "<Script Language=vbscript>" & vbcr
		Response.Write "	parent.frm1.txtPlantNm.value = """"" & vbcr
		Response.Write "	parent.frm1.txtPlantCd.Focus()" & vbcr
		Response.Write "</Script>" & vbcr
		Response.End
	Else
		Response.Write "<Script Language=vbscript>" & vbcr
		Response.Write "	parent.frm1.txtPlantNm.value = """ & ConvSPChars(rs1("PLANT_NM")) & """" & vbcr
		Response.Write "</Script>" & vbcr
		rs1.Close
		Set rs1 = Nothing
	End If
	
	' Order information Display
	Redim UNISqlId(0)
	Redim UNIValue(0, 3)

	UNISqlId(0) = "189752saa"
	
	IF Request("txtProdOrdNo") = "" Then
		strProdOrderNo = "|"
	Else
		StrProdOrderNo = FilterVar(UCase(Request("txtProdOrdNo")), "''", "S")
	End IF

	IF Request("txtOprNo") = "" Then
		strOprNo = "|"
	Else
		StrOprNo = FilterVar(UCase(Request("txtOprNo")), "''", "S")
	End IF

	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 2) = StrProdOrderNo 
	UNIValue(0, 3) = strOprNo 
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    Set ADF = Nothing
    
	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("189200", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		Response.Write "<Script Language=vbscript>" & vbcr
		Response.Write "	parent.frm1.txtProdOrderNo.focus()" & vbcr
		Response.Write "</Script>" & vbcr
		Response.End													'��: �����Ͻ� ���� ó���� ������ 
	End If
	
%>

<Script Language=vbscript>

With parent	
	.frm1.txtItemCd.Value = Trim("<%=ConvSPChars(rs0("ITEM_CD"))%>")
	.frm1.hItemCd.Value = Trim("<%=ConvSPChars(rs0("ITEM_CD"))%>")
	.frm1.txtItemNm.Value = Trim("<%=ConvSPChars(rs0("ITEM_NM"))%>")
    .frm1.txtWCCd.Value = Trim("<%=ConvSPChars(rs0("WC_CD"))%>")
	.frm1.txtWCNm.Value = Trim("<%=ConvSPChars(rs0("WC_NM"))%>")
	.frm1.txtOrderQty.Value = Trim("<%=UniConvNumberDBToCompany(rs0("PRODT_ORDER_QTY"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>") 
	.frm1.txtProdQty.Value = Trim("<%=UniConvNumberDBToCompany(rs0("PROD_QTY_IN_ORDER_UNIT"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>")
	.frm1.txtOrderUnit.Value = Trim("<%=ConvSPChars(rs0("PRODT_ORDER_UNIT"))%>")
	.frm1.txtGoodQty.Value = Trim("<%=UniConvNumberDBToCompany(rs0("GOOD_QTY_IN_ORDER_UNIT"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>")
	.frm1.txtRoutingNo.Value = Trim("<%=ConvSPChars(rs0("Rout_No"))%>")
	.frm1.txtTrackingNo.Value = Trim("<%=ConvSPChars(rs0("TRACKING_NO"))%>")
	.frm1.hRoutNo.value		= Trim("<%=ConvSPChars(rs0("Rout_No"))%>")
	.frm1.hPlantCd.value		= "<%=ConvSPChars(Request("txtPlantCd"))%>"
	.frm1.hProdOrderNo.value	= "<%=ConvSPChars(Request("txtProdOrdNo"))%>"
	.frm1.hOprNo.value			= "<%=ConvSPChars(Request("txtOprNo"))%>"

<%	
	rs0.Close
	Set rs0 = Nothing

%>
	.DbQueryOk()

End With

	
</Script>	
