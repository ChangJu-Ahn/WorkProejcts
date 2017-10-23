<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4117mb1_KO119.asp
'*  4. Program Name         : List Production Order Header
'*  5. Program Desc         : 
'*  6. Modified date(First) : 2006-04-11
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : 
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
Call LoadInfTB19029B("I", "*", "NOCOOKIE","MB")

On Error Resume Next

Dim ADF										'ActiveX Data Factory ���� �������� 
Dim strRetMsg								'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter ���� 
Dim rs0, rs1, rs2, rs3, rs4, rs5			'DBAgent Parameter ���� 
Dim strMode									'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim strQryMode
Dim lgStrPrevKey
Dim lgStrPrevKey1

Const C_SHEETMAXROWS_D = 50
'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Call HideStatusWnd

strMode = Request("txtMode")												'�� : ���� ���¸� ���� 
strQryMode = Request("lgIntFlgMode")
lgStrPrevKey = FilterVar(UCase(Request("lgStrPrevKey")), "", "SNM")
'lgStrPrevKey1 = FilterVar(UCase(Request("lgStrPrevKey1")), "''", "S")

On Error Resume Next

Dim strItemCd
Dim StrProdOrderNo
Dim StrWcCd
Dim StrTrackingNo
Dim strOrderType
Dim strFlag
Dim strCompleteFlag
Dim strStartDt
Dim strEndDt
Dim strItemGroupCd
Dim i

Err.Clear																	'��: Protect system from crashing

	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	Redim UNISqlId(2)
	Redim UNIValue(2, 0)

	UNISqlId(0) = "180000saa"
	UNISqlId(1) = "180000sab"
'	UNISqlId(2) = "180000sac"
	UNISqlId(2) = "180000sam"    
'	UNISqlId(4) = "180000sas"
	
	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(1, 0) = FilterVar(UCase(Request("txtItemCd")), "''", "S")
'	UNIValue(2, 0) = FilterVar(UCase(Request("txtWcCd")), "''", "S")
	UNIValue(2, 0) = FilterVar(UCase(Request("txtTrackingNo")), "''", "S")
'	UNIValue(4, 0) = FilterVar(UCase(Request("txtItemGroupCd")), "''", "S")

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1, rs2, rs3)

	%>
	<Script Language=vbscript>
		parent.frm1.txtPlantNm.value = ""
		parent.frm1.txtItemNm.value = ""
'		parent.frm1.txtWCNm.value = ""
'		parent.frm1.txtItemGroupNm.value = ""
	</Script>	
	<%

	'Plant �� Display      
	If (rs1.EOF And rs1.BOF) Then
		strFlag = "ERROR_PLANT"
	Else
		%>
		<Script Language=vbscript>
			parent.frm1.txtPlantNm.value = "<%=ConvSPChars(rs1("PLANT_NM"))%>"
		</Script>	
		<%    	
	End If

    'tracking no display
	IF Request("txtTrackingNo") <> "" Then
		If (rs3.EOF And rs3.BOF) Then
			strFlag = "ERROR_TRACK"
		End If
	End IF

	
	' ǰ��� Display
	IF Request("txtItemCd") <> "" Then
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
	
	rs1.Close	:	Set rs1 = Nothing
	rs2.Close	:	Set rs2 = Nothing
	rs3.Close	:	Set rs3 = Nothing
'	rs4.Close	:	Set rs4 = Nothing
'	rs5.Close	:	Set rs5 = Nothing
	
	If strFlag <> "" Then
		%>
		<Script Language=vbscript>
			Call parent.SetFieldColor(False)
		</Script>	
		<%
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
'		ElseIf strFlag = "ERROR_WCCD" Then		
'			Call DisplayMsgBox("182100", vbOKOnly, "", "", I_MKSCRIPT)
			%>
'			<Script Language=vbscript>
'			parent.frm1.txtWCCd.Focus()
'			</Script>	
			<%
'			Set ADF = Nothing
'			Response.End
		ElseIf strFlag = "ERROR_TRACK" Then
			Call DisplayMsgBox("203045", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
			parent.frm1.txtTrackingNo.Focus()
			</Script>	
			<%
			Set ADF = Nothing
			Response.End
'		ElseIf strFlag = "ERROR_GROUP" Then
'			Call DisplayMsgBox("127400", vbInformation, "", "", I_MKSCRIPT)
'			Response.Write "<Script Language=VBScript>" & vbCrLf
'				Response.Write "parent.frm1.txtItemGroupCd.focus" & vbCrLf
'			Response.Write "</Script>" & vbCrLf
'			Set ADF = Nothing
'			Response.End	
		End If
	End IF
	' Order Header Display
	Redim UNISqlId(0)
	Redim UNIValue(0, 2)

	'UNISqlId(0) = "P4412MB1H"
	UNISqlId(0) = "p4117ma101ko119"
	
	IF Request("txtItemCd") = "" Then
'		strItemCd = "|"
	Else
		StrItemCd = FilterVar(UCase(Request("txtItemCd")), "", "SNM")
	End IF

'	IF Request("txtWcCd") = "" Then
'		strWcCd = "|"
'	Else
'		StrWcCd = FilterVar(UCase(Request("txtWcCd")), "''", "S")
'	End IF

	IF Request("txtTrackingNo") = "" Then
'		strTrackingNo = "|"
	Else
		StrTrackingNo = UCase(FilterVar(Request("txtTrackingNo"),"","SNM"))
	End IF

	IF Request("txtProdFromDt") = "" Then
		strStartDt = "" & FilterVar("1900-01-01", "", "SNM") & ""
	Else
		strStartDt = " " & FilterVar(UniConvDate(Request("txtProdFromDt")), "", "SNM") & ""
	End IF

'	IF Request("txtProdTODt") = "" Then
'		strEndDt = "" & FilterVar("2999-01-01", "''", "S") & ""
'	Else
'		strEndDt = " " & FilterVar(UniConvDate(Request("txtProdTODt")), "''", "S") & ""
'	End IF

'	IF Request("txtOrderType") = "" Then
'		strOrderType = "|"
'	Else
'		strOrderType = " " & FilterVar(UCase(Request("txtOrderType")), "''", "S") & ""
'	End IF
	
	IF Request("txtProdOrdNo") <> "" Then
	    StrProdOrderNo = FilterVar(Request("txtProdOrdNo"),"","SNM") 
	End IF
	
'	IF Request("txtrdoflag") = "N" Then
'		strCompleteFlag = " (a.prodt_order_qty - b.prod_qty_in_order_unit) > " & FilterVar("0", "''", "S") & "  "
'	Else 
'		strCompleteFlag = " (a.prodt_order_qty - b.prod_qty_in_order_unit) <= " & FilterVar("0", "''", "S") & " "
'	End IF
	
'	IF Request("txtItemGroupCd") = "" Then
'		strItemGroupCd = "|"
'	Else
'		strItemGroupCd = "c.item_group_cd in (select item_group_cd from ufn_P_ListItemGrp(" & FilterVar(Trim(Request("txtItemGroupCd"))	, "''", "S") & " ))"
'	End IF
	
'	UNIValue(0, 0) = "^"
	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "", "SNM")
	UNIValue(0, 1) = FilterVar(UniConvDate(Request("txtProdFromDt")), "", "SNM") 
	UNIValue(0, 2) = ""
	
	Select Case strQryMode
		Case CStr(OPMD_CMODE)
			If Request("txtProdOrdNo") = "" Then
'				UNIValue(0, 2) = "|"
			Else 
				UNIValue(0, 2) = UNIValue(0, 2) & " and a.prodt_order_no >= '" & strProdOrderNo	& "'"
			End If
		Case CStr(OPMD_UMODE) 
			UNIValue(0, 2) = UNIValue(0, 2) & " and a.prodt_order_no > '" & lgStrPrevKey & "'"  
			UNIValue(0, 2) = UNIValue(0, 2) & " or (a.prodt_order_no = " & lgStrPrevKey 
			UNIValue(0, 2) = UNIValue(0, 2) & " and b.opr_no >= " & lgStrPrevKey1 & ")) "
	End Select
	
	If strItemCd <> "" then
	UNIValue(0, 2) = UNIValue(0, 2) & " and a.item_cd = '" & strItemCd & "'"
	End if
	
	If strTrackingNo <> "" then
	UNIValue(0, 2) = UNIValue(0, 2) & " and a.tracking_no = '" & strTrackingNo & "'" 
	End If
	
		
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
      
	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		%>
		<Script Language=vbscript>
			parent.DbQueryNotOk
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
				strData = strData & Chr(11) & Trim("<%=ConvSPChars(rs0("PRODT_ORDER_NO"))%>")
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEM_CD"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEM_NM"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("SPEC"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("SEC_ITEM_CD"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ROUT_NO"))%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("PRODT_ORDER_QTY"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Prodt_Order_SumQty"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("BASE_UNIT"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("TRACKING_NO"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("LAMP_MAKER"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("INVERT_MAKER"))%>"
				strData = strData & Chr(11) & LngMaxRow + <%=i%>
				strData = strData & Chr(11) & Chr(12)
				
				TmpBuffer(<%=i%>) = strData
<%		
				rs0.MoveNext
			End If
		Next
%>
		iTotalStr = Join(TmpBuffer, "")
		.ggoSpread.Source = .frm1.vspdData1
		.ggoSpread.SSShowDataByClip iTotalStr

		.lgStrPrevKey = Trim("<%=ConvSPChars(rs0("PRODT_ORDER_NO"))%>")
'		.lgStrPrevKey1 = "<%=ConvSPChars(rs0("OPR_NO"))%>"
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
		
	.DbQueryOk(LngMaxRow+1)

End With

</Script>	
<%
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
