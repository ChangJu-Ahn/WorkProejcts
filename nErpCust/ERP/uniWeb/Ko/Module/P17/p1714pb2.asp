<%@LANGUAGE = VBScript%>
<%'*******************************************************************************************
'*  1. Module Name          : ����BOM���� 
'*  2. Function Name        :
'*  3. Program ID           : p1714pb2.asp
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : +P32118ListProdOrderHeader
'*  7. Modified date(First) : 2005-02-18
'*  8. Modified date(Last)  :
'*  9. Modifier (First)     : yjw
'* 10. Modifier (Last)      :
'* 11. Comment              :
'********************************************************************************************%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "P", "NOCOOKIE", "PB")

Dim ADF										'ActiveX Data Factory ���� �������� 
Dim strRetMsg								'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter ���� 
Dim rs0, rs1, rs2							'DBAgent Parameter ���� 
Dim strQryMode
Dim i

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Const C_SHEETMAXROWS_D = 50

Call HideStatusWnd

Dim strReqTransNo

strQryMode = Request("lgIntFlgMode")
strReqTransNo = Trim(Request("txtReqTransNo"))

On Error Resume Next
Err.Clear																	'��: Protect system from crashing

	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
'	Redim UNISqlId(2)
'	Redim UNIValue(2, 0)
'
'	UNISqlId(0) = "180000saa"
'	UNISqlId(1) = "180000sab"
'
'	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
'	UNIValue(1, 0) = FilterVar(UCase(Request("txtItemCd")), "''", "S")
'
'	UNILock = DISCONNREAD :	UNIFlag = "1"
'
'    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
'    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1, rs2)

	' Order Header Display
	Redim UNISqlId(0)
	Redim UNIValue(0, 0)

	UNISqlId(0) = "p1714pb2"

'	UNIValue(0, 0) = "^"
	UNIValue(0, 0) = FilterVar(strReqTransNo,"''","S")

	UNILock = DISCONNREAD :	UNIFlag = "1"

    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		Set ADF = Nothing
		Response.End													'��: �����Ͻ� ���� ó���� ������ 
	End If

%>
<Script Language=vbscript>
    Dim LngMaxRow
    Dim strData
    Dim TmpBuffer
    Dim iTotalStr

    With parent												'��: ȭ�� ó�� ASP �� ��Ī�� 

 	LngMaxRow = .vspdData.MaxRows							'Save previous Maxrow

<%
	If Not(rs0.EOF And rs0.BOF) Then

		If C_SHEETMAXROWS_D < rs0.RecordCount Then
%>
			ReDim TmpBuffer(<%=C_SHEETMAXROWS_D - 1%>)
<%
		Else
%>
			ReDim TmpBuffer(<%=rs0.RecordCount - 1%>)
<%
		End If

		For i=0 to rs0.RecordCount-1
			If i < C_SHEETMAXROWS_D Then
%>
				strData = ""
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("REQ_TRANS_NO"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PLANT_CD"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PLANT_NM"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("CHILD_ITEM_CD"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEM_NM"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ERROR_CD"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ERROR_DESC"))%>"

				strData = strData & Chr(11) & LngMaxRow + <%=i%>
				strData = strData & Chr(11) & Chr(12)

				TmpBuffer(<%=i%>) = strData
<%
			rs0.MoveNext
			End If
		Next
%>

	iTotalStr = Join(TmpBuffer, "")

	.ggoSpread.Source = .vspdData
	.ggoSpread.SSShowDataByClip iTotalStr

'	.lgStrPrevKey = "<%=Trim(rs0("Prodt_Order_No"))%>"

<%
	End If

	rs0.Close
	Set rs0 = Nothing

%>

	If .vspdData.MaxRows < .PopupParent.VisibleRowCnt(.vspdData,0) and .lgStrPrevKey <> "" Then	 ' GroupView ������� ȭ�� Row������ ������ ������ �ٽ� ������ 
		.InitData(LngMaxRow)
		.DbQuery
	Else
		.hReqTransNo.value		= "<%=ConvSPChars(Request("txtReqTransNo"))%>"
		.DbQueryOk(LngMaxRow)
	End If

    End With
</Script>
