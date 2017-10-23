<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4416mb1.asp
'*  4. Program Name         : List Production Order Header
'*  5. Program Desc         : 
'*  6. Modified date(First) : 2000-05-14
'*  7. Modified date(Last)  : 2003-09-13
'*  8. Modifier (First)     : Park, Bum Soo
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
Call LoadInfTB19029B("I", "*", "NOCOOKIE","MB")


On Error Resume Next

Dim ADF										'ActiveX Data Factory ���� �������� 
Dim strRetMsg								'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter ���� 
Dim rs0, rs1, rs2, rs3, rs4					'DBAgent Parameter ���� 
Dim strMode									'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim strQryMode
Dim lgStrPrevKey

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
lgStrPrevKey = FilterVar(UCase(Request("lgStrPrevKey")), "''", "S")

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
	Redim UNISqlId(3)
	Redim UNIValue(3, 0)

	UNISqlId(0) = "180000saa"
	UNISqlId(1) = "180000sab"
	UNISqlId(2) = "180000sas"
	UNISqlId(3) = "180000sam"   
	 
	
	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(1, 0) = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	UNIValue(2, 0) = FilterVar(UCase(Request("txtItemGroupCd")), "''", "S")
	UNIValue(3, 0) = FilterVar(UCase(Request("txtTrackingNo")), "''", "S")
	

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1, rs2, rs3, rs4)

	%>
	<Script Language=vbscript>
		parent.frm1.txtPlantNm.value = ""
		parent.frm1.txtItemNm.value = ""
		parent.frm1.txtItemGroupNm.value = ""
	</Script>	
	<%

	' Plant �� Display      
	If (rs1.EOF And rs1.BOF) Then
		strFlag = "ERROR_PLANT"
	Else
		%>
		<Script Language=vbscript>
			parent.frm1.txtPlantNm.value = "<%=ConvSPChars(rs1("PLANT_NM"))%>"
		</Script>	
		<%    	
	End If
	
	' Item Group Check
	IF Request("txtItemGroupCd") <> "" Then
	 	If rs3.EOF AND rs3.BOF Then
			strFlag = "ERROR_GROUP"
		Else
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtItemGroupNm.value = """ & ConvSPChars(rs3("item_group_nm")) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		End If
	End If

	'tracking no display
	IF Request("txtTrackingNo") <> "" Then
		If (rs4.EOF And rs4.BOF) Then
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
	rs4.Close	:	Set rs4 = Nothing
	
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
		ElseIf strFlag = "ERROR_TRACK" Then
			Call DisplayMsgBox("203045", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
			parent.frm1.txtTrackingNo.Focus()
			</Script>	
			<%
			Set ADF = Nothing
			Response.End
		ElseIf strFlag = "ERROR_GROUP" Then
			Call DisplayMsgBox("127400", vbInformation, "", "", I_MKSCRIPT)
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtItemGroupCd.focus" & vbCrLf
			Response.Write "</Script>" & vbCrLf
			Set ADF = Nothing
			Response.End	
		End If
	End IF

	' Order Header Display
	Redim UNISqlId(0)
	Redim UNIValue(0, 9)

	UNISqlId(0) = "P4413MB1H"
	
	IF Request("txtItemCd") = "" Then
		strItemCd = "|"
	Else
		StrItemCd = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	End IF

	IF Request("txtProdOrdNo") = "" Then
		strProdOrderNo = "|"
	Else
		StrProdOrderNo = FilterVar(UCase(Request("txtProdOrdNo")), "''", "S")
	End IF

	IF Request("txtTrackingNo") = "" Then
		strTrackingNo = "|"
	Else
		StrTrackingNo = FilterVar(UCase(Request("txtTrackingNo")), "''", "S")
	End IF

	IF Request("txtProdFromDt") = "" Then
		strStartDt = "" & FilterVar("1900-01-01", "''", "S") & ""
	Else
		strStartDt = " " & FilterVar(UCase(UniConvDate(Request("txtProdFromDt"))), "''", "S") & ""
	End IF

	IF Request("txtProdTODt") = "" Then
		strEndDt = "" & FilterVar("2999-12-31", "''", "S") & ""
	Else
		strEndDt = " " & FilterVar(UCase(UniConvDate(Request("txtProdTODt"))), "''", "S") & ""
	End IF

	IF Request("txtOrderType") = "" Then
		strOrderType = "|"
	Else
		strOrderType = " " & FilterVar(UCase(Request("txtOrderType")), "''", "S") & ""
	End IF
	
	IF Request("txtrdoflag") = "N" Then
		strCompleteFlag = " (a.prodt_order_qty - a.prod_qty_in_order_unit) > " & FilterVar("0", "''", "S") & "  "
	Else 
		strCompleteFlag = " (a.prodt_order_qty - a.prod_qty_in_order_unit) <= " & FilterVar("0", "''", "S") & "  "
	End IF
	
	IF Request("txtItemGroupCd") = "" Then
		strItemGroupCd = "|"
	Else
		strItemGroupCd = " c.item_group_cd in (select item_group_cd from ufn_P_ListItemGrp(" & FilterVar(Trim(Request("txtItemGroupCd"))	, "''", "S") & " ))"
	End IF
	
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 2) = strItemCd 
	UNIValue(0, 3) = strTrackingNo
	UNIValue(0, 4) = strOrderType
	Select Case strQryMode
		Case CStr(OPMD_CMODE)
			UNIValue(0, 5) = StrProdOrderNo
		Case CStr(OPMD_UMODE) 
			UNIValue(0, 5) = lgStrPrevKey 
	End Select	
	UNIValue(0, 6) = strStartDt
	UNIValue(0, 7) = strEndDt
	UNIValue(0, 8) = strCompleteFlag
	UNIValue(0, 9) = strItemGroupCd
		
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
    	
With parent
	LngMaxRow = .frm1.vspdData1.MaxRows
		
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
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("PRODT_ORDER_QTY"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PRODT_ORDER_UNIT"))%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("PROD_QTY_IN_ORDER_UNIT"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("REMAIN_QTY"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("GOOD_QTY_IN_ORDER_UNIT"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("BAD_QTY_IN_ORDER_UNIT"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("INSP_GOOD_QTY_IN_ORDER_UNIT"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
 				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("INSP_BAD_QTY_IN_ORDER_UNIT"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("RCPT_QTY_IN_ORDER_UNIT"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("PLAN_START_DT"))%>"
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("PLAN_COMPT_DT"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("STATUS_NM"))%>"
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("RELEASE_DT"))%>"
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("REAL_START_DT"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ROUT_NO"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("WC_CD"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("WC_NM"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("TRACKING_NO"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ORDER_TYPE_NM"))%>"
				strData = strData & Chr(11) & "<%=rs0("AUTO_RCPT_FLG")%>"
				strData = strData & Chr(11) & "<%=rs0("LOT_FLG")%>"
				strData = strData & Chr(11) & "<%=rs0("LOT_GEN_MTHD")%>"
				strData = strData & Chr(11) & "<%=rs0("PROD_INSPEC_FLG")%>"
				strData = strData & Chr(11) & "<%=rs0("FINAL_INSPEC_FLG")%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_group_cd"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_group_nm"))%>"
				strData = strData & Chr(11) & Trim("<%=ConvSPChars(rs0("PARENT_ORDER_NO"))%>")
				strData = strData & Chr(11) & Trim("<%=ConvSPChars(rs0("PARENT_OPR_NO"))%>")
				strData = strData & Chr(11) & Trim("<%=ConvSPChars(rs0("ORIGINAL_ORDER_NO"))%>")
				strData = strData & Chr(11) & Trim("<%=ConvSPChars(rs0("ORIGINAL_OPR_NO"))%>")
				strData = strData & Chr(11) & Trim("<%=ConvSPChars(rs0("OPR_NO"))%>")
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("TOTAL_REWORK_QTY"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
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
<%	
	End If

	rs0.Close
	Set rs0 = Nothing

%>	
	.frm1.hPlantCd.value		= "<%=ConvSPChars(Request("txtPlantCd"))%>"
	.frm1.hItemCd.value			= "<%=ConvSPChars(Request("txtItemCd"))%>"
	.frm1.hProdOrderNo.value	= "<%=ConvSPChars(Request("txtProdOrderNo"))%>"
	.frm1.hTrackingNo.value		= "<%=ConvSPChars(Request("txtTrackingNo"))%>"
	.frm1.hOrderType.value		= "<%=ConvSPChars(Request("txtOrderType"))%>"
	.frm1.hProdFromDt.value		= "<%=ConvSPChars(Request("txtProdFromDt"))%>"
	.frm1.hProdTODt.value		= "<%=ConvSPChars(Request("txtProdTODt"))%>"
	.frm1.hrdoFlag.value		= "<%=Request("txtrdoflag")%>"
	.frm1.hItemGroupCd.value	= "<%=ConvSPChars(Request("txtItemGroupCd"))%>"
		
	.DbQueryOk(LngMaxRow+1)

End With

</Script>	
<%
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>