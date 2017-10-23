<%@LANGUAGE = VBScript%>
<%'====================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4511rb1.asp
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) :
'*  7. Modified date(Last)  : 2002/12/12
'*  8. Modifier (First)     : 
'*  9. Modifier (Last)      : Ryu Sung Won
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=====================================================================================================%>

<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "P", "NOCOOKIE", "RB")

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

Call HideStatusWnd

strMode = Request("txtMode")						'�� : ���� ���¸� ���� 

Dim strItemCd
Dim StrProdOrderNo
Dim StrWcCd
Dim StrTrackingNo
Dim StrSlCd

On Error Resume Next	
Err.Clear                                    		'��: Protect system from crashing

	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	Redim UNISqlId(0)
	Redim UNIValue(0, 4)

	UNISqlId(0) = "p4111mb1"
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = " " & FilterVar(UCase(Request("lgPlantCD")), "''", "S") & ""
	UNIValue(0, 2) = " " & FilterVar(UCase(Request("lgProdOrderNo")), "''", "S") & ""
	UNIValue(0, 3) = "|"
	UNIValue(0, 4) = "|"

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
	Set ADF = Server.CreateObject("prjPublic.cCtlTake")
	strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1)	
	
	If (rs1.EOF And rs1.BOF) Then
		Call DisplayMsgBox("189200", vbOKOnly, "", "", I_MKSCRIPT)
		rs1.Close
		Set rs1 = Nothing
		Set ADF = Nothing
		Response.End
	Else
		%>
		<Script Language=vbscript>
			parent.txtItemCd.value = "<%=ConvSPChars(rs1("ITEM_CD"))%>"
			parent.txtItemNm.value = "<%=ConvSPChars(rs1("ITEM_NM"))%>"
		</Script>	
		<%
		Set rs1 = Nothing
	End If
	
	Redim UNISqlId(0)
	Redim UNIValue(0, 9)

	UNISqlId(0) = "189660saa"
	
	IF Request("txtItemCd") = "" Then
		strItemCd = "|"
	Else
		StrItemCd = " " & FilterVar(UCase(Request("txtItemCd")), "''", "S") & ""
	End IF

	IF Request("txtWcCd") = "" Then
		strWcCd = "|"
	Else
		StrWcCd = " " & FilterVar(UCase(Request("txtWcCd")), "''", "S") & ""
	End IF

	IF Request("txtTrackingNo") = "" Then
		strTrackingNo = "|"
	Else
		StrTrackingNo = " " & FilterVar(UCase(Request("txtTrackingNo")), "''", "S") & ""
	End IF

	IF Request("txtSlCd") = "" Then
		strSlCd = "|"
	Else
		strSlCd = " " & FilterVar(UCase(Request("txtSlCd")), "''", "S") & ""
	End IF

	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = "" & FilterVar("MR", "''", "S") & ""
	UNIValue(0, 2) = " " & FilterVar(Request("lgPlantCd"), "''", "S") & ""
	UNIValue(0, 3) = " " & FilterVar(Request("lgProdOrderNo"), "''", "S") & ""
	UNIValue(0, 4) = "|" 'strItemCd 
	UNIValue(0, 5) = "|" 'strSlCd
	UNIValue(0, 6) = "|" 'strWcCd
	UNIValue(0, 7) = "|" 'strTrackingNo
	UNIValue(0, 8) = "|" 'strTrackingNo
	UNIValue(0, 9) = "" & FilterVar("Y", "''", "S") & " "
		
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
      
	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		%>
		<Script Language=vbscript>
		Parent.DbQueryNotOk
		</Script>
		<%
		Response.End													'��: �����Ͻ� ���� ó���� ������ 
	End If
	
%>

<Script Language=vbscript>
Dim LngMaxRow
Dim strData
    	
With parent																	'��: ȭ�� ó�� ASP �� ��Ī�� 

	LngMaxRow = .vspdData.MaxRows
	
	ReDim TmpBuffer(<%=rs0.RecordCount - 1%>)	
<%  
	If Not(rs0.EOF And rs0.BOF) Then
		For i=0 to rs0.RecordCount-1 
%>
			strData = ""
			strData = strData & Chr(11) & "<%=UniDateClientFormat(rs0("POS_DT"))%>"				'�԰��� 
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("QTY"),ggQty.DecPoint,0)%>"'�԰� 
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("BASE_UNIT"))%>"								'���� 
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("LOT_NO"))%>"									'LOT��ȣ 
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("LOT_SUB_NO"))%>"								'LOT ���� 
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEM_DOCUMENT_NO"))%>"						'��ǥ��ȣ 
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("TRNS_TYPE"))%>"								'������� 
			strData = strData & Chr(11) & LngMaxRow + <%=i%>
			strData = strData & Chr(11) & Chr(12)
			
			TmpBuffer(<%=i%>) = strData
<%		
			rs0.MoveNext
		Next
%>
		iTotalStr = Join(TmpBuffer, "")
		.ggoSpread.Source = .vspdData
		.ggoSpread.SSShowDataByClip iTotalStr
		
<%	
	End If

	rs0.Close
	Set rs0 = Nothing

%>	
	.DbQueryOk

End With

</Script>	
<%
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
