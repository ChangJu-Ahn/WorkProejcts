<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4114mb4.asp	
'*  4. Program Name         : look up Work Center
'*  5. Program Desc         :
'*  6. Comproxy List        : DB Agent
'*  7. Modified date(First) : 2000/03/31
'*  8. Modified date(Last)  : 2002/06/29
'*  9. Modifier (First)     : Mr  Kim
'* 10. Modifier (Last)      : Park, BumSoo
'* 11. Comment              :
'**********************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

Call LoadBasisGlobalInf

On Error Resume Next								'��: 

Dim ADF												'ActiveX Data Factory ���� �������� 
Dim strRetMsg										'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0		'DBAgent Parameter ���� 
Dim strWcCd, strWcNm
Dim Row, Row1
Dim strProdtOrderNo, strOprNo, strInsideFlg

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Call HideStatusWnd

    Err.Clear															'��: Protect system from crashing

	Row = Request("Row")
	Row1 = Request("Row1")
	strProdtOrderNo = Request("txtProdtOrderNo")
	strOprNo = Request("txtOprNo")
	strWcCd = Trim(Request("txtWcCd"))

	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	Redim UNISqlId(0)
	Redim UNIValue(0, 1)
	
	UNISqlId(0) = "p4114mb4"
	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 1) = FilterVar(UCase(Request("txtWcCd")), "''", "S")

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
	
	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("182100", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		Set ADF = Nothing
		%>
		<Script Language=vbscript>
			Call parent.LookUpWcNotOk("<%=Row%>")
		</Script>
		<%	   
		Response.End													'��: �����Ͻ� ���� ó���� ������ 
	Else
		strWcNm		 = ConvSPChars(rs0("Wc_Nm"))
		strInsideFlg = UCase(rs0("Inside_Flg"))
		%>
		<Script Language=vbscript>
			Call parent.LookUpWcOk("<%=strWcCd%>", "<%=strWcNm%>","<%=strInsideFlg%>","<%=Row%>","<%=Row1%>","<%=strProdtOrderNo%>","<%=strOprNo%>")
		</Script>
		<%
	End If

Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
