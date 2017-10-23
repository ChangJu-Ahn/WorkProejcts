<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4321mb1
'*  4. Program Name         : List BackLog 
'*  5. Program Desc         : 
'*  6. Modified date(First) : 2006-04-11
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     :HJO
'*  9. Modifier (Last)      : 
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "P", "NOCOOKIE","MB")

On Error Resume Next

Dim ADF										'ActiveX Data Factory ���� �������� 
Dim strRetMsg								'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter ���� 
Dim rs0, rs1, rs2, rs3, rs4, rs5			'DBAgent Parameter ���� 
Dim strMode									'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim strQryMode
Dim lgStrPrevKey,lgStrPrevKey1
Dim lgStrPrevKey2,lgStrPrevKey3,lgStrPrevKey4
Dim i

Const C_SHEETMAXROWS_D = 100
'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Call HideStatusWnd

strMode = Request("txtMode")												'�� : ���� ���¸� ���� 
strQryMode = Request("lgIntFlgMode")
lgStrPrevKey = FilterVar(UCase(Request("lgStrPrevKey")), "''", "S")
lgStrPrevKey1 = FilterVar(UCase(Request("lgStrPrevKey1")), "''", "S")
lgStrPrevKey2 = FilterVar(UCase(Request("lgStrPrevKey2")), "''", "S")
lgStrPrevKey3 = FilterVar(UCase(Request("lgStrPrevKey3")), "''", "S")
lgStrPrevKey4 = FilterVar(UCase(Request("lgStrPrevKey4")), "''", "S")

On Error Resume Next

Dim strItemCd
Dim StrProdOrderNo
Dim StrTrackingNo
Dim strCompleteFlag
Dim strStartDt
Dim strEndDt
Dim strFlag

Err.Clear																	'��: Protect system from crashing
	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	Redim UNISqlId(3)
	Redim UNIValue(3, 0)

	UNISqlId(0) = "180000saa"
	UNISqlId(1) = "180000sab"	
	UNISqlId(2) = "180000sam" 
	
	
	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(1, 0) = FilterVar(UCase(Request("txtItemCd")), "''", "S")	
	UNIValue(2, 0) = FilterVar(UCase(Request("txtTrackingNo")), "''", "S")


	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1, rs2, rs3)

	%>
	<Script Language=vbscript>
		parent.frm1.txtPlantNm.value = ""
		parent.frm1.txtItemNm.value = ""
	</Script>	
	<%

	' Plant �� Display  
	strFlag=""    
	If (rs1.EOF And rs1.BOF) Then
		strFlag = "ERROR_PLANT"		
	Else
		%>
		<Script Language=vbscript>
			parent.frm1.txtPlantNm.value = "<%=ConvSPChars(rs1("PLANT_NM"))%>"
		</Script>	
		<%    		
	End If
	
		'ǰ��� Display
	IF strFlag ="" and Request("txtItemCd") <> "" Then
		If (rs2.EOF And rs2.BOF) Then
			strFlag = "ERROR_ITEM"
		Else
			%>
			<Script Language=vbscript>
				parent.frm1.txtItemNm.value = "<%=ConvSPChars(rs2("ITEM_NM"))%>"
			</Script>	
			<%
		End If
	End IF
    'tracking no display
	IF strFlag="" and Request("txtTrackingNo") <> "" Then
		If (rs3.EOF And rs3.BOF) Then
			strFlag = "ERROR_TRACK"
		End If
	End IF
	
	rs1.Close	:	Set rs1 = Nothing
	rs2.Close	:	Set rs2 = Nothing
	rs3.Close	:	Set rs3 = Nothing
	
	If strFlag <> "" Then
		If strFlag = "ERROR_PLANT" Then
			Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
			parent.frm1.txtPlantCd.Focus()
			</Script>	
			<%
			Set ADF = Nothing
			Response.End
			
		ElseIf strFlag = "ERROR_ITEM" Then
			Call DisplayMsgBox("122600", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
			parent.frm1.txtItemCd.Focus()
			</Script>	
			<%
			Set ADF = Nothing
			Response.End
			
		ElseIf strFlag = "ERROR_TRACK" Then
			Call DisplayMsgBox("203045", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
			parent.frm1.txtTrackingNo.Focus()
			</Script>	
			<%
			Set ADF = Nothing
			Response.End		
		End If
	End IF
	
	' Order Header Display
	Redim UNISqlId(0)
	Redim UNIValue(0, 7)

	UNISqlId(0) = "P4321MA1S"
	
	
	IF Request("txtItemCd") = "" Then
		strItemCd = "|"
	Else
		StrItemCd = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	End IF

	IF Request("txtTrackingNo") = "" Then
		strTrackingNo = "|"
	Else
		StrTrackingNo = FilterVar(UCase(Request("txtTrackingNo")), "''", "S")
	End IF

	IF Request("txtProdFromDt") = "" Then
		strStartDt = "" & FilterVar("1900-01-01", "''", "S") & ""
	Else
		strStartDt = " " & FilterVar(UniConvDate(Request("txtProdFromDt")), "''", "S") & ""
	End IF

	IF Request("txtProdTODt") = "" Then
		strEndDt = "" & FilterVar("2999-12-31", "''", "S") & ""
	Else
		strEndDt = " " & FilterVar(UniConvDate(Request("txtProdTODt")), "''", "S") & ""
	End IF

	IF Request("txtProdOrdNo") <> "" Then
		StrProdOrderNo = FilterVar(UCase(Request("txtProdOrdNo")), "''", "S")
	End IF
	
	IF Request("txtrdoflag") = "A" Then
		strCompleteFlag = " (a.status <> " & FilterVar("R", "''", "S") & " AND a.status <> "  & FilterVar("D", "''", "S") & ") "
	ElseIf Request("txtrdoflag")="N" Then 
		strCompleteFlag = " a.status in ( " & FilterVar(Request("txtrdoflag"), "''", "S") & " ,'I' ) "
	Else
		strCompleteFlag = " a.status = " & FilterVar(Request("txtrdoflag"), "''", "S") & "  "
	End IF	
	
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 2) = strStartDt
	UNIValue(0, 3) = strEndDt	
	UNIValue(0, 4) = strItemCd 
	UNIValue(0, 5) = strTrackingNo

	Select Case strQryMode
		Case CStr(OPMD_CMODE)
			If Request("txtProdOrdNo") = "" Then
				UNIValue(0, 6) = "|"
			Else 
				UNIValue(0, 6) = " a.prodt_order_no >= " & strProdOrderNo	
			End If
		Case CStr(OPMD_UMODE) 
			UNIValue(0, 6) = " (a.report_dt > " & lgStrPrevKey2  
			UNIValue(0, 6) = UNIValue(0, 6) & " or (a.report_dt = " & lgStrPrevKey2 
			UNIValue(0, 6) = UNIValue(0, 6) & " and a.prodt_order_no > " & lgStrPrevKey & ")"	
			UNIValue(0, 6) = UNIValue(0, 6) & " or (a.report_dt = " & lgStrPrevKey2 
			UNIValue(0, 6) = UNIValue(0, 6) & " and a.prodt_order_no = " & lgStrPrevKey		
			UNIValue(0, 6) = UNIValue(0, 6) & " and a.opr_no > " & lgStrPrevKey1  & ")"	
			UNIValue(0, 6) = UNIValue(0, 6) & " or (a.report_dt >= " & lgStrPrevKey2 
			UNIValue(0, 6) = UNIValue(0, 6) & " and a.prodt_order_no = " & lgStrPrevKey		
			UNIValue(0, 6) = UNIValue(0, 6) & " and a.opr_no = " & lgStrPrevKey1 
			UNIValue(0, 6) = UNIValue(0, 6) & " and a.resv_seq > " & lgStrPrevKey3  & ")"	
			UNIValue(0, 6) = UNIValue(0, 6) & " or (a.report_dt >= " & lgStrPrevKey2 
			UNIValue(0, 6) = UNIValue(0, 6) & " and a.prodt_order_no = " & lgStrPrevKey		
			UNIValue(0, 6) = UNIValue(0, 6) & " and a.opr_no = " & lgStrPrevKey1 
			UNIValue(0, 6) = UNIValue(0, 6) & " and a.resv_seq = " & lgStrPrevKey3
			UNIValue(0, 6) = UNIValue(0, 6) & " and a.result_seq >= " & lgStrPrevKey4 & ")) "
	End Select	
	UNIValue(0, 7) = strCompleteFlag
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
      
	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		%>
		<Script Language=vbscript>
			parent.DbQueryNotOk()
		</Script>	
		<%		
		Response.End													'��: �����Ͻ� ���� ó���� ������ 
	End If
	
%>

<Script Language=vbscript>
Dim LngMaxRow
Dim strData
Dim TmpBuffer
Dim iTotalStr
    	
With parent																	'��: ȭ�� ó�� ASP �� ��Ī�� 
	LngMaxRow = .frm1.vspdData1.MaxRows										'Save previous Maxrow
		
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
				strData = strData & Chr(11) & Trim("<%=ConvSPChars(rs0("CHK"))%>")									'������ȣ 
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("REPORT_DT"))%>"									'���������� 
				strData = strData & Chr(11) & Trim("<%=ConvSPChars(rs0("PRODT_ORDER_NO"))%>")									'������ȣ 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("OPR_NO"))%>"													'�������� 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEM_CD"))%>"												'ǰ�� 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEM_NM"))%>"												'ǰ��� 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("SPEC"))%>"													'�԰� 
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("ISSUE_QTY"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"				'�������� 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("REMARK"))%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("GOOD_QTY"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"						'�ܷ� 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("BASIC_UNIT"))%>"										'�������� 
							
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("SL_CD"))%>"												'����� 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("SL_NM"))%>"													'�۾��� 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("TRACKING_NO"))%>"													'�۾���� 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("REQ_NO"))%>"													'�۾� 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("RESV_SEQ"))%>"													'�۾��� 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("RESULT_SEQ"))%>"												'�۾����� 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEM_DOCUMENT_NO"))%>"											'���û��� 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("STATUS"))%>"											'���� 				
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("STATUS_NM"))%>"	
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ERROR_DESC"))%>"														'MIlestone
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PLANT_CD"))%>"	
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("DOCUMENT_YEAR"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("WC_CD"))%>"	
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("LOT_NO"))%>"	
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("LOT_SUB_NO"))%>"	
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("COST_CD"))%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("SCHD_QTY"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"				'�������� 
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("ORIGIN_QTY"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"				'��������  

				strData = strData & Chr(11) & LngMaxRow + <%=i%>
				strData = strData & Chr(11) & Chr(12)
				
				TmpBuffer(<%=i%>) = strData
<%		
				rs0.MoveNext
			End If
		Next
%>
		iTotalStr = Join(TmpBuffer,"")
		.ggoSpread.Source = .frm1.vspdData1
		.ggoSpread.SSShowDataByClip iTotalStr
		
		.lgStrPrevKey = Trim("<%=ConvSPChars(rs0("PRODT_ORDER_NO"))%>")
		.lgStrPrevKey1 = "<%=ConvSPChars(rs0("OPR_NO"))%>"
		.lgStrPrevKey2 = "<%=ConvSPChars(rs0("REPORT_DT"))%>"
		.lgStrPrevKey3 = "<%=ConvSPChars(rs0("RESV_SEQ"))%>"
		.lgStrPrevKey4 = "<%=ConvSPChars(rs0("RESULT_SEQ"))%>"
		
<%	
	End If

	rs0.Close
	Set rs0 = Nothing

%>	
	.frm1.hPlantCd.value		= "<%=ConvSPChars(Request("txtPlantCd"))%>"
	.frm1.hItemCd.value			= "<%=ConvSPChars(Request("txtItemCd"))%>"
	.frm1.hProdOrderNo.value	= "<%=ConvSPChars(Request("txtProdOrderNo"))%>"
	.frm1.hTrackingNo.value		= "<%=ConvSPChars(Request("txtTrackingNo"))%>"
	.frm1.hProdFromDt.value		= "<%=Request("txtProdFromDt")%>"
	.frm1.hProdTODt.value		= "<%=Request("txtProdTODt")%>"
	.frm1.hrdoFlag.value		= "<%=Request("txtrdoflag")%>"
	
	.DbQueryOk(LngMaxRow+1)

End With

</Script>	
<%
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
