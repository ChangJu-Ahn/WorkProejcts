<%@LANGUAGE = VBScript%>
<%'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1601mb3.asp
'*  4. Program Name         : Look Up VAT Type 
'*  5. Program Desc         :
'*  6. Comproxy List        : + 
'*  7. Modified date(First) : 2002/01/15
'*  8. Modified date(Last)  : 2002/10/07
'*  9. Modifier (First)     : Jung Yu Kyung
'* 10. Modifier (Last)      : Ryu Sung Won
'* 11. Comment              :
'*************************************************************************************************%>

<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("I", "*", "NOCOOKIE", "MB")

Dim ADF												'ActiveX Data Factory ���� �������� 
Dim strRetMsg										'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag			'DBAgent Parameter ���� 
Dim rs0, rs1, rs2, rs3								'DBAgent Parameter ���� 
Dim strQryMode
Dim strMode																	'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim strVatType
Dim strVatRate
Dim Row 

strMode = Request("txtMode")												'�� : ���� ���¸� ���� 
strVatType = Trim(Request("txtVatType"))
'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Call HideStatusWnd

'On Error Resume Next
Err.Clear
	
	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	Redim UNISqlId(1)
	Redim UNIValue(1, 1)

	UNISqlId(0) = "p1601mb3a"
	UNISqlId(1) = "p1601mb3b"
	
	UNIValue(0, 0) = "" & FilterVar("B9001", "''", "S") & ""
	UNIValue(0, 1) = " " & FilterVar(strVatType, "''", "S") & ""
	UNIValue(1, 0) = "" & FilterVar("B9001", "''", "S") & ""	
	UNIValue(1, 1) = " " & FilterVar(strVatType, "''", "S") & ""

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)

	If (rs0.EOF and rs0.BOF)Then
		Call DisplayMsgBox("115100", vbOKOnly, strVatType, "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		Set ADF = Nothing
		Response.End
	End If

	If (rs1.EOF and rs1.BOF)Then
		Call DisplayMsgBox("115100", vbOKOnly, strVatType, "", I_MKSCRIPT)
		rs1.Close
		Set rs1 = Nothing
		Set ADF = Nothing
		Response.End
	End If

	strVatRate = UniConvNumberDBToCompany(rs0("reference"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
	Row = Request("Row")

%>
<Script Language=vbscript>
	With parent.frm1.vspdData
		Call parent.LookUpVatTypeOk("<%=strVatType%>","<%=strVatRate%>","<%=Row%>")
	End With
<%			
		rs0.Close
		Set rs0 = Nothing
		rs1.Close
		Set rs1 = Nothing
%>
</Script>	
<%
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
