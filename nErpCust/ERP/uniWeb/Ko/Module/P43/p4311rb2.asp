<%@LANGUAGE = VBScript%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4311rb2.asp
'*  4. Program Name         : 
'*  5. Program Desc         : List Onhand Stock Detail
'*  6. Modified date(First) : 
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Park, BumSoo
'*  9. Modifier (Last)      : 
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

Dim ADF												'ActiveX Data Factory ���� �������� 
Dim strRetMsg										'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0		'DBAgent Parameter ���� 
Dim strMode											'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim strQryMode
Dim strSlCd
Dim strNextSlCd		
Dim strLotNo
Dim strLotSubNo

Dim i

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Const C_SHEETMAXROWS_D = 50

Call HideStatusWnd

strMode = Request("txtMode")												'�� : ���� ���¸� ���� 
strQryMode = Request("lgIntFlgMode")

On Error Resume Next

	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	If Request("txtSlCd") <> "" Then
		Redim UNISqlId(0)
		Redim UNIValue(0, 0)

		UNISqlId(0) = "180000sad"	
	
		UNIValue(0, 0) = FilterVar(UCase(Request("txtSlCd")), "''", "S")

		UNILock = DISCONNREAD :	UNIFlag = "1"
	
		Set ADF = Server.CreateObject("prjPublic.cCtlTake")
		strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

		If (rs0.EOF And rs0.BOF) Then
			rs0.Close
			Set rs0 = Nothing
			%>
			<Script Language=vbscript>
				Parent.txtSLNm.value = ""
			</Script>	
			<%
			Call DisplayMsgBox("125700", vbOKOnly, "", "", I_MKSCRIPT)
			Response.End													'��: �����Ͻ� ���� ó���� ������ 
		Else
			%>
			<Script Language=vbscript>
				Parent.txtSLNm.value = "<%=ConvSPChars(rs0("SL_NM"))%>"
			</Script>	
			<%
			rs0.Close
			Set rs0 = Nothing
		End If
	Else
		%>
		<Script Language=vbscript>
			Parent.txtSLNm.value = ""
		</Script>	
		<%
	End IF
	
	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	
	Redim UNISqlId(0)
	Redim UNIValue(0, 4)

	UNISqlId(0) = "p4311rb2"

	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 2) = FilterVar(UCase(Request("txtChildItemCd")), "''", "S")
	UNIValue(0, 3) = FilterVar(UCase(Request("txtTrackingNo")), "''", "S")
	
	strSlCd = FilterVar(UCase(Request("txtMajorSlCd")), "''", "S")
	
	If Request("lgStrPrevKey3") <> "" Then
		strNextSlCd = FilterVar(UCase(Request("lgStrPrevKey3")), "''", "S")
		strLotNo = FilterVar(UCase(Request("lgStrPrevKey4")), "''", "S")
		strLotSubNo = FilterVar(UCase(Trim(Request("lgStrPrevKey5"))),"" & FilterVar("0", "''", "S") & " ","S")
	Else
		strNextCd = "''"
		strLotNo = "''"
		strLotSubNo = "''"
	End If

	If strSlCd <> "''" Then	
		If strLotNo <> "''" Then
			UNIValue(0, 4) = " a.sl_cd = " & strSlCd & " and (a.lot_no >= " & strLotNo & " or (a.lot_no = " & strLotNo & " and a.lot_sub_no >= " & strLotSubNo & " ))"
		Else
			UNIValue(0, 4) = " a.sl_cd = " & strSlCd
		End If
	Else
		If strLotNo <> "''" Then
			UNIValue(0, 4) = " (a.sl_cd > " & strNextSlCd & " or (a.sl_cd >= " & strNextSlCd & " and a.lot_no > " & strLotNo & " ) or (a.sl_cd >= " & strNextSlCd & " and a.lot_no >= " & strLotNo & " and a.lot_sub_no >= " & strLotSubNo & " )) "
		Else
			UNIValue(0, 4) = "|"
		End If
	End If
	
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
Dim strData
Dim TmpBuffer
Dim iTotalStr
    	
With parent																'��: ȭ�� ó�� ASP �� ��Ī�� 
	LngMaxRow = .vspdData2.MaxRows									'Save previous Maxrow
		
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
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("BLOCK_INDICATOR"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("SL_CD"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("SL_NM"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("TRACKING_NO"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("LOT_NO"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("LOT_SUB_NO"))%>"
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("GOOD_ON_HAND_QTY"),ggQty.DecPoint,0)%>"
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("PREV_GOOD_QTY"),ggQty.DecPoint,0)%>"
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("STK_ON_INSP_QTY"),ggQty.DecPoint,0)%>"
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("STK_ON_TRNS_QTY"),ggQty.DecPoint,0)%>"
				strData = strData & Chr(11) & "0"
				strData = strData & Chr(11) & LngMaxRow + <%=i%>
				strData = strData & Chr(11) & Chr(12)
				
				TmpBuffer(<%=i%>) = strData
<%		
			rs0.MoveNext
			End If
		Next
%>
		iTotalStr = Join(TmpBuffer, "")
		.ggoSpread.Source = .vspdData2
		.ggoSpread.SSShowDataByClip iTotalStr
		
		.lgStrPrevKey3 = "<%=ConvSPChars(rs0("SL_CD"))%>"
		.lgStrPrevKey4 = "<%=ConvSPChars(rs0("LOT_NO"))%>"
		.lgStrPrevKey5 = "<%=ConvSPChars(rs0("LOT_SUB_NO"))%>"
		
<%	
	End If

	rs0.Close
	Set rs0 = Nothing

%>	
	.DbDtlQueryOk(LngMaxRow)									'��: ��ȸ ������ ������� 


End With

</Script>	
<%
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
