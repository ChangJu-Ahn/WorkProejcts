<%'======================================================================================================
'*  1. Module Name          : Purchase
'*  2. Function Name        : 
'*  3. Program ID           : m5314mb1.asp
'*  4. Program Name         : ���ڼ��ݰ�꼭 
'*  5. Program Desc         : 
'*  6. Modified date(First) : 2009-07-07
'*  7. Modified date(Last)  : 2009-07-07
'*  8. Modifier (First)     : Lee Min Hyung
'*  9. Modifier (Last)      : Lee Min Hyung
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
Call LoadInfTB19029B("Q", "M", "NOCOOKIE","MB")
Call LoadBNumericFormatB("Q","M", "NOCOOKIE", "MB")

Dim ADF												'ActiveX Data Factory ���� �������� 
Dim strRetMsg										'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter ���� 
Dim rs0					                     'DBAgent Parameter ���� 
Dim strMode											'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

Call HideStatusWnd

Dim i

Err.Clear											'��: Protect system from crashing

' Order Header Display
Redim UNISqlId(0)
Redim UNIValue(0, 0)

UNISqlId(0) = "D1311MA12"

strMode = Request("txtMode")											'�� : ���� ���¸� ���� 

'UNIValue(0, 0) = "^"
UNIValue(0, 0) = FilterVar(UCase(Request("txtTaxBillNo")), "''", "S")

UNILock = DISCONNREAD :	UNIFlag = "1"

Set ADF = Server.CreateObject("prjPublic.cCtlTake")
strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

If (rs0.EOF And rs0.BOF) Then
	Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
	rs0.Close
	Set rs0 = Nothing
	Set ADF = Nothing
	Response.End																		'��: �����Ͻ� ���� ó���� ������ 
End If	%>
	
<Script Language=vbscript>
	Dim LngMaxRow
	Dim strData
	Dim TmpBuffer
	Dim iTotalStr1

	With parent																			'��: ȭ�� ó�� ASP �� ��Ī�� 
		LngMaxRow = .frm1.vspdData2.MaxRows										'Save previous Maxrow
		ReDim TmpBuffer(<%=rs0.RecordCount - 1%>)

<%		Dim iDx
		For iDx = 0 to rs0.RecordCount-1 %>
			strData = ""
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_std"))%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("item_prc"),ggUnitCost.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("item_qty"),ggQty.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("item_date"))%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("item_amt"),ggAmtOfMoney.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("item_tax"),ggAmtOfMoney.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_memo"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("inv_no"))%>"
			strData = strData & Chr(11) & LngMaxRow + <%=iDx%>
			strData = strData & Chr(11) & Chr(12)

			TmpBuffer(<%=iDx%>) = strData
<%			rs0.MoveNext
		Next %>

		iTotalStr1 = Join(TmpBuffer, "")
		.ggoSpread.Source = .frm1.vspdData2
		.ggoSpread.SSShowDataByClip iTotalStr1

<%		rs0.Close
		Set rs0 = Nothing	%>
	End With
</Script>	
<%
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>

