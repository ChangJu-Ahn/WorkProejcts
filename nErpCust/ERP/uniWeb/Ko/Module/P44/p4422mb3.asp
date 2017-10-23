<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4213mb3.asp
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) : 2001/10/31
'*  7. Modified date(Last)  : 2003/05/26
'*  8. Modifier (First)     : Park, Bumsoo
'*  9. Modifier (Last)      : Chen, Jae Hyun
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


On Error Resume Next								'��: 

Dim ADF														'ActiveX Data Factory ���� �������� 
Dim strRetMsg												'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1			'DBAgent Parameter ���� 
Dim strQryMode

Dim strMode											'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Dim i

Dim lgStrPrevKey3	
Dim strPlantCd, strItemCd, strWcCd, strProdtOrderNo, strOrderStatus, strTrackingNo

Call HideStatusWnd

strMode = Request("txtMode")						'�� : ���� ���¸� ���� 
strQryMode = Request("lgIntFlgMode")

On Error Resume Next

	lgStrPrevKey3 = UCase(Trim(Request("lgStrPrevKey3")))	

	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	Redim UNISqlId(0)
	Redim UNIValue(0, 6)
	
	UNISqlId(0) = "p4422mb3"
			
	IF Trim(Request("txtPlantCd")) = "" Then
	   strPlantCd = "|"
	ELSE
	   strPlantCd = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	END IF
	
	IF Trim(Request("txtItemCd")) = "" Then
	   strItemCd = "|"
	ELSE
	   strItemCd = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	END IF

	IF Trim(Request("txtWcCd")) = "" Then
	   strWcCd = "|"
	ELSE
	   strWcCd = FilterVar(UCase(Request("txtWcCd")), "''", "S")
	END IF
		
	IF Trim(Request("txtProdtOrderNo")) = "" Then
	   strProdtOrderNo = "|"
	ELSE
	   strProdtOrderNo = FilterVar(UCase(Request("txtProdtOrderNo")), "''", "S")
	END IF
		
	IF Trim(Request("cboOrderStatus")) = "" Then
	   strOrderStatus = "|"
	ELSE
	   strOrderStatus = " " & FilterVar(UCase(Request("cboOrderStatus")), "''", "S") & ""
	END IF
		
	IF Trim(Request("txtTrackingNo")) = "" Then
	   strTrackingNo = "|"
	ELSE
	   strTrackingNo = FilterVar(UCase(Request("txtTrackingNo")), "''", "S")
	END IF
	
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = strPlantCd
	UNIValue(0, 2) = strItemCd
	UNIValue(0, 3) = strWcCd
	UNIValue(0, 4) = strProdtOrderNo
	UNIValue(0, 5) = strOrderStatus
	UNIValue(0, 6) = strTrackingNo
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

	If rs0.EOF And rs0.BOF Then
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

With parent																	'��: ȭ�� ó�� ASP �� ��Ī�� 
	LngMaxRow = .frm1.vspdData3.MaxRows										'Save previous Maxrow
	
	ReDim TmpBuffer(<%=rs0.RecordCount - 1%>)	
<%  
    For i=0 to rs0.RecordCount-1 
%>
		strData = ""
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PRODT_ORDER_NO"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("OPR_NO"))%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("wipinorderunit"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("prodt_order_unit"))%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("prevprodqtyinorderunit"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("prevgoodqtyinorderunit"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("prodqtyinorderunit"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("goodqtyinorderunit"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("badqtyinorderunit"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("wipinbaseunit"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("base_unit"))%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("prevprodqtyinbaseunit"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("prevgoodqtyinbaseunit"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("prodqtyinbaseunit"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("goodqtyinbaseunit"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("badqtyinbaseunit"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("tracking_no"))%>"						
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ORDER_STATUS"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ORDER_STATUS"))%>"
		strData = strData & Chr(11) & LngMaxRow + <%=i%>
		strData = strData & Chr(11) & Chr(12)
		
		TmpBuffer(<%=i%>) = strData
<%		
		rs0.MoveNext
	Next
%>
	iTotalStr = Join(TmpBuffer, "")
	.ggoSpread.Source = .frm1.vspdData3
	.ggoSpread.SSShowDataByClip iTotalStr
		
	.lgStrPrevKey3 = "<%=Trim(rs0("PRODT_ORDER_NO"))%>"
<%			
	rs0.Close
	Set rs0 = Nothing
%>
	.DbQuery3Ok(LngMaxRow)

End With	
</Script>	
<%
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
