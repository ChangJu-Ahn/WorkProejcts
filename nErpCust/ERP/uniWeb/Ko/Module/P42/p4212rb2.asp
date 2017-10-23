<%@LANGUAGE = VBScript%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4212rb2.asp
'*  4. Program Name         : List Production Order Header
'*  5. Program Desc         : 
'*  6. Modified date(First) : ?
'*  7. Modified date(Last)  : 2002/12/20
'*  8. Modifier (First)     : Park, Bum Soo
'*  9. Modifier (Last)      : Ryu Sung Won
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "P", "NOCOOKIE", "RB")

Dim ADF										'ActiveX Data Factory ���� �������� 
Dim strRetMsg								'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter ���� 
Dim rs0										'DBAgent Parameter ���� 
Dim strMode									'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Call HideStatusWnd

strMode = Request("txtMode")												'�� : ���� ���¸� ���� 

Dim strItemCd
Dim StrProdOrderNo
Dim StrWcCd
Dim StrTrackingNo
Dim strOrderType
Dim strFlag
Dim strSlCd

On Error Resume Next
Err.Clear																	'��: Protect system from crashing

	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	Dim i
	
	Redim UNISqlId(0)
	Redim UNIValue(0, 3)

	UNISqlId(0) = "P4212RB2"
	
	IF Request("txtItemCd") = "" Then
		strItemCd = "|"
	Else
		StrItemCd = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	End IF

	IF Request("txtSlCd") = "" Then
		strSlCd = "|"
	Else
		StrSlCd = FilterVar(UCase(Request("txtSlCd")), "''", "S")
	End IF

	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 2) = strItemCd 
	UNIValue(0, 3) = strSlCd

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
      
	If (rs0.EOF And rs0.BOF) Then
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
    	
With parent																'��: ȭ�� ó�� ASP �� ��Ī�� 
	LngMaxRow = .vspdData2.MaxRows										'Save previous Maxrow
		
<%  
	If Not(rs0.EOF And rs0.BOF) Then
%>			
		ReDim TmpBuffer(<%=rs0.RecordCount - 1%>)
<%
		For i=0 to rs0.RecordCount-1 
%>
			strData = ""
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("tracking_no"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("lot_no"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("lot_sub_no"))%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("good_on_hand_qty"),ggQty.DecPoint,0)%>"							'�������� 
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("block_indicator"))%>"
			strData = strData & Chr(11) & LngMaxRow + <%=i%>
			strData = strData & Chr(11) & Chr(12)
			
			TmpBuffer(<%=i%>) = strData
<%		
			rs0.MoveNext
		Next
%>
		iTotalStr = Join(TmpBuffer, "")
		.ggoSpread.Source = .vspdData2
		.ggoSpread.SSShowDataByClip iTotalStr
		
<%	
	End If

	rs0.Close
	Set rs0 = Nothing

%>	
	.hPlantCd.value		= "<%=ConvSPChars(Request("txtPlantCd"))%>"
	.hItemCd.value		= "<%=ConvSPChars(Request("txtItemCd"))%>"
	.hSlCd.value		= "<%=ConvSPChars(Request("txtSlCd"))%>"
		
	.DbDtlQueryOk(LngMaxRow)

End With

</Script>	
<%
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
