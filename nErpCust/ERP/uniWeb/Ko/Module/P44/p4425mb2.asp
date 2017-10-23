<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4425mb2.asp
'*  4. Program Name         : ������������ȸ 
'*  5. Program Desc         :
'*  6. Modified date(First) : 2003-02-27
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Chen, Jae Hyun
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================
%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "P", "NOCOOKIE","MB")

Dim ADF												'ActiveX Data Factory ���� �������� 
Dim strRetMsg										'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag			'DBAgent Parameter ���� 
Dim rs0, rs1, rs2, rs3, rs4							'DBAgent Parameter ���� 

Dim strPlantCd
Dim strReportFromDt
Dim strReportToDt
Dim strProdtOrderNo
Dim strShiftCd

Dim strMode											'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Dim i

Call HideStatusWnd

On Error Resume Next
Err.Clear

	strMode = Request("txtMode")						'�� : ���� ���¸� ���� 
	strPlantCd = Request("txtPlantCd")
	strReportFromDt = Request("txtReportFromDt")
	strReportToDt = Request("txtReportToDt")
	strProdtOrderNo = Request("txtProdOrderNo")
	strShiftCd = Request("txtShiftCd")
	
	strPlantCd = FilterVar(UCase(strPlantCd), "''", "S")
	strReportFromDt = FilterVar(UniConVDate(strReportFromDt), "''", "S")
	strReportToDt = FilterVar(UniConVDate(strReportToDt), "''", "S")
	strProdtOrderNo = FilterVar(UCase(strProdtOrderNo), "''", "S")
	
	IF Trim(strShiftCd) = "" Then
	   strShiftCd = "|"
	ELSE
	   strShiftCd = FilterVar(UCase(strShiftCd), "''", "S")
	END IF
		
	Redim UNISqlId(0)
	Redim UNIValue(0, 5)

	UNISqlId(0) = "p4425mb1D"
	
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = strPlantCd
	UNIValue(0, 2) = strProdtOrderNo
	UNIValue(0, 3) = strReportFromDt
	UNIValue(0, 4) = strReportToDt 	
	UNIValue(0, 5) = strShiftCd
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

	If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT) 	
		rs0.Close
		Set rs0 = Nothing					
		Response.End													'��: �����Ͻ� ���� ó���� ������ 
	End If
%>

<Script Language=vbscript>
Dim LngMaxRow 
Dim strData
Dim TmpBuffer
Dim iTotalStr
    	
With parent
																'��: ȭ�� ó�� ASP �� ��Ī�� 
	LngMaxRow = .frm1.vspdData2.MaxRows									'Save previous Maxrow
	
	ReDim TmpBuffer(<%=rs0.RecordCount - 1%>)
<%  		
    For i=0 to rs0.RecordCount-1 
%>
		strData = ""
		strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("REPORT_DT"))%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("PRODT_QTY_IN_ORDER_UNIT"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("GOOD_QTY_IN_ORDER_UNIT"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("BAD_QTY_IN_ORDER_UNIT"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("RCPT_QTY_IN_ORDER_UNIT"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(UCase(Trim(rs0("PRODT_ORDER_UNIT"))))%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("PRODT_QTY_IN_BASE_UNIT"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("GOOD_QTY_IN_BASE_UNIT"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("BAD_QTY_IN_BASE_UNIT"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("RCPT_QTY_IN_BASE_UNIT"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(UCase(Trim(rs0("BASE_UNIT"))))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(UCase(Trim(rs0("PRODT_ORDER_NO"))))%>"	
		strData = strData & Chr(11) & LngMaxRow + <%=i%>
		strData = strData & Chr(11) & Chr(12)
		
		TmpBuffer(<%=i%>) = strData
		
<%		
		rs0.MoveNext
	Next
%>
	iTotalStr = Join(TmpBuffer, "")
	.ggoSpread.Source = .frm1.vspdData2
	.ggoSpread.SSShowDataByClip iTotalStr
		
<%		
	rs0.Close
	Set rs0 = Nothing
%>
	
End With	
</Script>	
<%
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
