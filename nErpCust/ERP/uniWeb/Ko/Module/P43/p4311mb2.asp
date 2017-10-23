<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4311mb2.asp
'*  4. Program Name         : 
'*  5. Program Desc         : List Onhand Stock Detail
'*  6. Modified date(First) : ?
'*  7. Modified date(Last)  : 2002/08/21
'*  8. Modifier (First)     : Park, BumSoo
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
Call LoadInfTB19029B("I", "*","NOCOOKIE","MB")

On Error Resume Next								'��: 

Dim ADF												'ActiveX Data Factory ���� �������� 
Dim strRetMsg										'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1	'DBAgent Parameter ���� 

Dim strMode											'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Dim i
Dim strPlantCd
Dim strChildItemCd
Dim strChildItemSeq
Dim	strProdOrderNo
Dim strOprNo
Dim strMRPReqNo
Dim strUnit
Dim strWcCd

'@Var_Declare

Call HideStatusWnd

strMode = Request("txtMode")						'�� : ���� ���¸� ���� 

On Error Resume Next

Dim strConvType
	
	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	
	Redim UNISqlId(0)
	Redim UNIValue(0, 2)

	strPlantCd		= UCase(Trim(Request("txtPlantCd")))
	strChildItemCd	= UCase(Trim(Request("txtChildItemCd")))
	strChildItemSeq = Request("txtReqSeqNo")
	strProdOrderNo	= Request("txtProdtOrderNo")
	strOprNo		= Request("txtOprNo")
	strMRPReqNo		= Request("txtMRPReqNo")
	strUnit			= Request("txtUnit")
	strWcCd			= Request("txtWcCd")

	UNISqlId(0) = "189510sab"

	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 2) = FilterVar(UCase(Request("txtChildItemCd")), "''", "S")
			
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
      
	If (rs0.EOF And rs0.BOF) Then
		rs0.Close
		Set rs0 = Nothing
		Response.End													'��: �����Ͻ� ���� ó���� ������ 
	End If
	
%>

<Script Language=vbscript>
Dim LngMaxRow
Dim LngMaxRows
Dim strData, strData1
Dim TmpBuffer, TmpBuffer1
Dim iTotalStr, iTotalStr1
   	
With parent																'��: ȭ�� ó�� ASP �� ��Ī�� 
	LngMaxRow = .frm1.vspdData2.MaxRows									'Save previous Maxrow
	LngMaxRows = .frm1.vspdData3.MaxRows
	
<%  
	If Not(rs0.EOF And rs0.BOF) Then		
%>			
		ReDim TmpBuffer(<%=rs0.RecordCount - 1%>)
		ReDim TmpBuffer1(<%=rs0.RecordCount - 1%>)
<%
		For i=0 to rs0.RecordCount-1 
%>
			strData = ""
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("BLOCK_INDICATOR"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("SL_CD"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("SL_NM"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("TRACKING_NO"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("LOT_NO"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("LOT_SUB_NO"))%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("GOOD_ON_HAND_QTY"),ggQty.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat("0",ggQty.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("STK_ON_INSP_QTY"),ggQty.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("STK_ON_TRNS_QTY"),ggQty.DecPoint,0)%>"
			strData = strData & Chr(11) & LngMaxRow + <%=i%>
			strData = strData & Chr(11) & Chr(12)
			
			TmpBuffer(<%=i%>) = strData
			
			strData1 = ""
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(strChildItemCd)%>"
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(strChildItemSeq)%>"
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("BLOCK_INDICATOR"))%>"
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("SL_CD"))%>"
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("SL_NM"))%>"
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("TRACKING_NO"))%>"
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("LOT_NO"))%>"
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("LOT_SUB_NO"))%>"
			strData1 = strData1 & Chr(11) & "<%=UniNumClientFormat(rs0("GOOD_ON_HAND_QTY"),ggQty.DecPoint,0)%>"
			strData1 = strData1 & Chr(11) & "<%=UniNumClientFormat("0",ggQty.DecPoint,0)%>"
			strData1 = strData1 & Chr(11) & "<%=UniNumClientFormat(rs0("STK_ON_INSP_QTY"),ggQty.DecPoint,0)%>"
			strData1 = strData1 & Chr(11) & "<%=UniNumClientFormat(rs0("STK_ON_TRNS_QTY"),ggQty.DecPoint,0)%>"
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(strPlantCd)%>"
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(strProdOrderNo)%>"
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(strOprNo)%>"
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(strMRPReqNo)%>"
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(strUnit)%>"
			strData1 = strData1 & Chr(11) & ""
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(strWcCd)%>"
			strData1 = strData1 & Chr(11) & LngMaxRows + <%=i+1%>
			strData1 = strData1 & Chr(11) & Chr(12)
			
			TmpBuffer1(<%=i%>) = strData1	
<%		
			rs0.MoveNext
		Next
%>
		iTotalStr1 = Join(TmpBuffer1, "")
		.ggoSpread.Source = .frm1.vspdData3
		.ggoSpread.SSShowDataByClip iTotalStr1
		
		iTotalStr = Join(TmpBuffer, "")
		.ggoSpread.Source = .frm1.vspdData2
		.ggoSpread.SSShowDataByClip iTotalStr
<%	
	End If

	rs0.Close
	Set rs0 = Nothing

%>	

	.frm1.hPlantCd.value = "<%=ConvSPChars(Request("txtPlantCd"))%>"
	.frm1.hProdOrderNo.value = "<%=ConvSPChars(Request("txtProdtOrderNo"))%>"			

	.DbDtlQueryOk(LngMaxRow)									'��: ��ȸ ������ ������� 


End With

</Script>	
<%
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
