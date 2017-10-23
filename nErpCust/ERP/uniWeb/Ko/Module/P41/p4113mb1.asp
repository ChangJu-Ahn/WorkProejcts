<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%
'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4113mb1.asp
'*  4. Program Name         : List Production Order Header
'*  5. Program Desc         : 
'*  6. Modified date(First) : 2002-05-08
'*  7. Modified date(Last)  : 2002-05-08
'*  8. Modifier (First)     : Park, BumSoo
'*  9. Modifier (Last)      : Park, BumSoo
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call LoadBasisGlobalInf
Call loadInfTB19029B("Q", "P", "NOCOOKIE","MB")
On Error Resume Next

Dim ADF										'ActiveX Data Factory ���� �������� 
Dim strRetMsg								'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter ���� 
Dim rs0, rs1, rs2, rs3, rs4					'DBAgent Parameter ���� 
Dim strQryMode								'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Dim i

Const C_SHEETMAXROWS_D = 100

Call HideStatusWnd

strQryMode = Request("lgIntFlgMode")

Dim strStartDt
Dim strEndDt
Dim strItemCd
Dim strProdOrderNo
Dim strTrackingNo
Dim strOrderType
Dim strOrderStatus
Dim strItemGroupCd
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
	UNISqlId(3) = "180000sas"

	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(1, 0) = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	UNIValue(2, 0) = FilterVar(UCase(Request("txtTrackingNo")), "''", "S")
	UNIValue(3, 0) = FilterVar(UCase(Request("txtItemGroupCd")), "''", "S")

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1, rs2, rs3, rs4)

	' Plant �� Display      
	If (rs1.EOF And rs1.BOF) Then
		rs1.Close
		Set rs1 = Nothing
		strFlag = "ERROR_PLANT"
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtPlantNm.value = """"" & vbCrLf	
		Response.Write "</Script>" & vbCrLf
	Else
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtPlantNm.value = """ & ConvSPChars(rs1("PLANT_NM")) & """" & vbCrLf
		Response.Write "</Script>" & vbCrLf
		rs1.Close
		Set rs1 = Nothing
	End If
	' ǰ��� Display
	IF Request("txtItemCd") <> "" Then
		If (rs2.EOF And rs2.BOF) Then
			rs2.Close
			Set rs2 = Nothing
			strFlag = "ERROR_ITEM"
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtItemNm.value = """"" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		Else
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtItemNm.value = """ & ConvSPChars(rs2("ITEM_NM")) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf
			rs2.Close
			Set rs2 = Nothing
		End If
	Else
		rs2.Close
		Set rs2 = Nothing
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtItemNm.value = """"" & vbCrLf
		Response.Write "</Script>" & vbCrLf
	End IF
	' Tracking No. Check
	IF Request("txtTrackingNo") <> "" Then
		If (rs3.EOF And rs3.BOF) Then
			rs3.Close
			Set rs3 = Nothing
			strFlag = "ERROR_TRACK"
		Else
			rs3.Close
			Set rs3 = Nothing
		End If
	Else
		rs3.Close
		Set rs3 = Nothing
	End IF
	' Item Group Check
	IF Request("txtItemGroupCd") <> "" Then
	 	If rs4.EOF AND rs4.BOF Then
			rs4.Close
			Set rs4 = Nothing
			strFlag = "ERROR_GROUP"
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtItemGroupNm.value = """"" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		Else
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtItemGroupNm.value = """ & ConvSPChars(rs4("item_group_nm")) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf
			rs4.Close
			Set rs4 = Nothing
		End If
	Else
		rs4.Close
		Set rs4 = Nothing
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtItemGroupNm.value = """"" & vbCrLf
		Response.Write "</Script>" & vbCrLf
	End If
		
	If strFlag <> "" Then
		If strFlag = "ERROR_PLANT" Then
			Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtPlantCd.focus" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		ElseIf strFlag = "ERROR_ITEM" Then
			Call DisplayMsgBox("122600", vbOKOnly, "", "", I_MKSCRIPT)
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtItemCd.focus" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		ElseIf strFlag = "ERROR_TRACK" Then
			Call DisplayMsgBox("203045", vbOKOnly, "", "", I_MKSCRIPT)
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtTrackingNo.focus" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		ElseIf strFlag = "ERROR_GROUP" Then
			Call DisplayMsgBox("127400", vbInformation, "", "", I_MKSCRIPT)
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtItemGroupCd.focus" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		End If
		Set ADF = Nothing
		Response.End
	End IF
	Set ADF = Nothing

	' Order Header Display
	Redim UNISqlId(0)
	Redim UNIValue(0, 9)

	UNISqlId(0) = "p4113mb1"
	
	IF Request("txtFromDt") = "" Then
		strStartDt = "|"
	Else
		strStartDt = " " & FilterVar(UNIConvDate(Request("txtFromDt")), "''", "S") & ""
	End IF

	IF Request("txtToDt") = "" Then
		strEndDt = "|"
	Else
		strEndDt = " " & FilterVar(UNIConvDate(Request("txtToDt")), "''", "S") & ""
	End IF
	
	IF Request("txtItemCd") = "" Then
		strItemCd = "|"
	Else
		StrItemCd = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	End IF

	IF Request("txtTrackingNo") = "" Then
		strTrackingNo = "|"
	Else
		strTrackingNo = FilterVar(UCase(Request("txtTrackingNo")), "''", "S")
	End IF

	Select Case strQryMode
		Case CStr(OPMD_CMODE)
			If Request("txtProdOrderNo") = "" Then
				strProdOrderNo = "|"
			Else
				strProdOrderNo = FilterVar(UCase(Request("txtProdOrderNo")), "''", "S")
			End If	
		Case CStr(OPMD_UMODE) 
			strProdOrderNo = FilterVar(UCase(Request("lgStrPrevKey")), "''", "S")
	End Select 

	IF Request("cboOrderType") = "" Then
		strOrderType = "|"
	Else
		strOrderType = " " & FilterVar(UCase(Request("cboOrderType")), "''", "S") & ""
	End IF

	IF Request("cboOrderStatus") = "" Then
		strOrderStatus = "|"
	Else
		strOrderStatus = " " & FilterVar(UCase(Request("cboOrderStatus")), "''", "S") & ""
	End IF

	IF Request("txtItemGroupCd") = "" Then
		strItemGroupCd = "|"
	Else
		strItemGroupCd = "b.item_group_cd in (select item_group_cd from ufn_P_ListItemGrp(" & FilterVar(Trim(Request("txtItemGroupCd"))	, "''", "S") & " ))"
	End IF

	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 2) = strStartDt
	UNIValue(0, 3) = strEndDt
	UNIValue(0, 4) = strItemCd 
	UNIValue(0, 5) = strTrackingNo
	UNIValue(0, 6) = strProdOrderNo		
	UNIValue(0, 7) = strOrderType
	UNIValue(0, 8) = strOrderStatus
	UNIValue(0, 9) = strItemGroupCd
	
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
    	
With parent																	'��: ȭ�� ó�� ASP �� ��Ī�� 
	LngMaxRow = .frm1.vspdData.MaxRows										'Save previous Maxrow

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
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Prodt_Order_No"))%>"											'��: Production Order No
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Item_Cd"))%>"												'��: Item Code
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Item_Nm"))%>"												'��: Item Name
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Spec"))%>"												'��: Specification
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Prodt_Order_Qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Prodt_Order_Unit"))%>"										'��: Unit
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Prod_Qty_In_Order_Unit"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Good_Qty_In_Order_Unit"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Bad_Qty_In_Order_Unit"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Insp_Good_Qty_In_Order_Unit"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Insp_Bad_Qty_In_Order_Unit"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Rcpt_Qty_In_Order_Unit"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("Plan_Start_Dt"))%>"									'��: Planned Start Date
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("Plan_Compt_Dt"))%>"									'��: Planned Completion Date
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("Real_Start_Dt"))%>"									'��: Real Start Date
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("Real_Compt_Dt"))%>"									'��: Real Start Date
				strData = strData & Chr(11) & "<%=rs0("Order_Status")%>"
				strData = strData & Chr(11) & "<%=rs0("Order_Status")%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Re_Work_Flg"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Rout_No"))%>"												'��: Routing
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Sl_Cd"))%>"													'��: Storage Location Code
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Sl_Nm"))%>"													'��: Storage Location Name
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("Schd_Start_Dt"))%>"									'��: Scheduled Start Date
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("Schd_Compt_Dt"))%>"									'��: Scheduled Completion Date
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("Release_Dt"))%>"										'��: Release Date
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Tracking_No"))%>"											'��: Tracking No.
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Order_Qty_In_Base_Unit"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Base_Unit"))%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Prod_Qty_In_Base_Unit"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Good_Qty_In_Base_Unit"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Bad_Qty_In_Base_Unit"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Insp_Good_Qty_In_Base_Unit"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Insp_Bad_Qty_In_Base_Unit"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Rcpt_Qty_In_Base_Unit"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=rs0("Prodt_Order_Type")%>"
				strData = strData & Chr(11) & "<%=rs0("Prodt_Order_Type")%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_group_cd"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_group_nm"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("cost_cd"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("cost_nm"))%>"
				
				strData = strData & Chr(11) & LngMaxRow + <%=i%>
				strData = strData & Chr(11) & Chr(12)
				
				TmpBuffer(<%=i%>) = strData
<%		
				rs0.MoveNext
			End If
		Next
%>
		iTotalStr = Join(TmpBuffer, "")
		.ggoSpread.Source = .frm1.vspdData
		.ggoSpread.SSShowDataByClip iTotalStr
		
		.lgStrPrevKey = "<%=Trim(rs0("PRODT_ORDER_NO"))%>"
		
<%	
	End If

	rs0.Close
	Set rs0 = Nothing

%>	
	If .frm1.vspdData.MaxRows < .VisibleRowCnt(.frm1.vspdData,0) and .lgStrPrevKey <> "" Then	<% ' GroupView ������� ȭ�� Row������ ������ ������ �ٽ� ������ %>
		Call .InitData(LngMaxRow)
		.DbQuery
	Else
		.frm1.hPlantCd.value	= "<%=ConvSPChars(Request("txtPlantCd"))%>"
		.frm1.hItemCd.value		= "<%=ConvSPChars(Request("txtItemCd"))%>"
		.frm1.hProdOrderNo.value= "<%=ConvSPChars(Request("txtProdOrderNo"))%>"
		.frm1.hProdFromDt.value	= "<%=Request("txtFromDt")%>"
		.frm1.hProdToDt.value	= "<%=Request("txtToDt")%>"
		.frm1.hOrderType.value	= "<%=ConvSPChars(Request("cboOrderType"))%>"
		.frm1.hOrderStatus.value= "<%=ConvSPChars(Request("cboOrderStatus"))%>"
		.frm1.hTrackingNo.value= "<%=ConvSPChars(Request("txtTrackingNo"))%>"
		.frm1.hItemGroupCd.value= "<%=ConvSPChars(Request("txtItemGroupCd"))%>"
		
		.DbQueryOk(LngMaxRow+1)
	End If

End With

</Script>	
<%
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
