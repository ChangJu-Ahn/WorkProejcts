<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4212mb3.asp
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) :
'*  7. Modified date(Last)  : 2001.11.16
'*  8. Modifier (First)     : Park, BumSoo
'*  9. Modifier (Last)      : Park, BumSoo
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call LoadBasisGlobalInf
Call loadInfTB19029B("I", "*", "NOCOOKIE","MB")

On Error Resume Next								'��: 

Dim ADF														'ActiveX Data Factory ���� �������� 
Dim strRetMsg												'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1			'DBAgent Parameter ���� 

Const C_SHEETMAXROWS = 100

Dim strMode											'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Dim lgNextFlg
Dim i 

Call HideStatusWnd

On Error Resume Next

Dim strProdOrdNo
	
	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	Redim UNISqlId(0)
	Redim UNIValue(0, 0)
	
	UNISqlId(0) = "p4212mb3"
		
	IF Trim(Request("txtProdOrderNo")) = "" Then
	   strProdOrdNo = ""
	ELSE
	   strProdOrdNo = FilterVar(UCase(Request("txtProdOrderNo")), "''", "S")
	END IF

	UNIValue(0, 0) = strProdOrdNo
			
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

	'-------------------------------------------
	' Display Spread 2
	'-------------------------------------------
	If rs0.EOF And rs0.BOF Then
		rs0.Close
		Set rs0 = Nothing
		Response.End 
	End If

%>
<Script Language=vbscript>
	Dim LngMaxRow
	Dim strData
	Dim TmpBuffer
    Dim iTotalStr
    	
	With parent																	'��: ȭ�� ó�� ASP �� ��Ī�� 
		LngMaxRow = .frm1.vspdData2.MaxRows										'Save previous Maxrow
		ReDim TmpBuffer(<%=rs0.RecordCount-1%>)
<%  
		For i=0 to rs0.RecordCount-1 
%>
			strData = ""
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("CHILD_ITEM_CD"))%>"		
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEM_NM"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("SPEC"))%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("REQ_QTY"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("BASE_UNIT"))%>"
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(ConvSPChars(rs0("REQ_DT")))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("TRACKING_NO"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("SL_CD"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("SL_NM"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("OPR_NO"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("WC_CD"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("WC_NM"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ISSUE_MTHD"))%>"
			strData = strData & Chr(11) & LngMaxRow + <%=i%>
			strData = strData & Chr(11) & Chr(12)
			
			TmpBuffer(<%=i%>) = strData
<%		
			rs0.MoveNext
		Next
%>
		iTotalStr= Join(TmpBuffer, "")
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
