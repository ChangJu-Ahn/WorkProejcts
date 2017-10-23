<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4425mb1.asp
'*  4. Program Name         : ������������ȸ 
'*  5. Program Desc         :
'*  6. Modified date(First) : 2003-02-19
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Ryu Sung Won
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
Call LoadInfTB19029B("Q", "P", "NOCOOKIE","MB")

Dim ADF												'ActiveX Data Factory ���� �������� 
Dim strRetMsg										'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag			'DBAgent Parameter ���� 
Dim rs0, rs1, rs2, rs3, rs4, rs5					'DBAgent Parameter ���� 
Dim strQryMode

Dim strPlantCd
Dim strReportFromDt
Dim strReportToDt
Dim strProdtOrderNo
Dim strItemCd
Dim strTrackingNo
Dim strShiftCd
Dim strOrderStatus
Dim strItemGroupCd
Dim strFlag

Const C_SHEETMAXROWS_D = 100

Dim strMode											'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Dim i

Call HideStatusWnd

On Error Resume Next
Err.Clear

	strMode = Request("txtMode")						'�� : ���� ���¸� ���� 
	strQryMode = Request("lgIntFlgMode")
	
	strPlantCd = Request("txtPlantCd")
	strReportFromDt = Request("txtReportFromDt")
	strReportToDt = Request("txtReportToDt")
	strProdtOrderNo = Request("txtProdOrderNo")
	strItemCd = Request("txtItemCd")
	strTrackingNo = Request("txtTrackingNo")
	strShiftCd = Request("txtShiftCd")
	strOrderStatus = Request("cboOrderStatus")
	strItemGroupCd = Request("txtItemGroupCd")
	
	IF Trim(strPlantCd) = "" Then
	   Call DisplayMsgBox("900010", vbOKOnly, "", "", I_MKSCRIPT)
	   Response.End
	ELSE
	   strPlantCd = FilterVar(UCase(strPlantCd), "''", "S")
	END IF
	
	IF Trim(strReportFromDt) = "" Then
	   Call DisplayMsgBox("900010", vbOKOnly, "", "", I_MKSCRIPT)
	   Response.End
	ELSE
	   strReportFromDt = FilterVar(UniConVDate(strReportFromDt), "''", "S")
	END IF
	
	IF Trim(strReportToDt) = "" Then
	   Call DisplayMsgBox("900010", vbOKOnly, "", "", I_MKSCRIPT)
	   Response.End
	ELSE
	   strReportToDt = FilterVar(UniConVDate(strReportToDt), "''", "S")
	END IF
	
	IF Trim(strProdtOrderNo) = "" Then
	   strProdtOrderNo = "|"
	ELSE
	   strProdtOrderNo = FilterVar(UCase(strProdtOrderNo), "''", "S")
	END IF
	
	IF Trim(strItemCd) = "" Then
	   strItemCd = "|"
	ELSE
	   strItemCd = FilterVar(UCase(strItemCd), "''", "S")
	END IF

	IF Trim(strTrackingNo) = "" Then
	   strTrackingNo = "|"
	ELSE
	   strTrackingNo = FilterVar(UCase(strTrackingNo), "''", "S")
	END IF
	
	IF Trim(strShiftCd) = "" Then
	   strShiftCd = "|"
	ELSE
	   strShiftCd = FilterVar(UCase(strShiftCd), "''", "S")
	END IF
	
	IF Trim(strOrderStatus) = "" Then
	   strOrderStatus = "|"
	ELSE
	   strOrderStatus = FilterVar(UCase(strOrderStatus), "''", "S")
	END IF
	
	IF Trim(strItemGroupCd) = "" Then
		strItemGroupCd = "|"
	Else
		strItemGroupCd = "c.item_group_cd in (select item_group_cd from ufn_P_ListItemGrp(" & FilterVar(Trim(Request("txtItemGroupCd"))	, "''", "S") & " ))"
	End IF
	
	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	Redim UNISqlId(4)
	Redim UNIValue(4, 1)

	UNISqlId(0) = "180000saa"
	UNISqlId(1) = "180000sab"
	UNISqlId(2) = "180000sam"
	UNISqlId(3) = "180000sao"
	UNISqlId(4) = "180000sas"

	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(1, 0) = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	UNIValue(2, 0) = FilterVar(UCase(Request("txtTrackingNo")), "''", "S")
	UNIValue(3, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(3, 1) = FilterVar(UCase(Request("txtShiftCd")), "''", "S")
	UNIValue(4, 0) = FilterVar(UCase(Request("txtItemGroupCd")), "''", "S")

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1, rs2, rs3, rs4, rs5)

   	' Plant �� Display      
	%>
	<Script Language=vbscript>
		parent.frm1.txtPlantNm.value = ""
		parent.frm1.txtItemNm.value = ""
		parent.frm1.txtItemGroupNm.value = ""
	</Script>	
	<%    	
	' Plant �� Display      
	If (rs1.EOF And rs1.BOF) Then
		Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
		rs1.Close
		Set rs1 = Nothing
		Set ADF = Nothing
		%>
		<Script Language=vbscript>
			parent.frm1.txtPlantCd.Focus()
		</Script>	
		<%
		Response.End
	Else
		%>
		<Script Language=vbscript>
			parent.frm1.txtPlantNm.value = "<%=ConvSPChars(rs1("PLANT_NM"))%>"
		</Script>	
		<%    	
		rs1.Close
		Set rs1 = Nothing
	End If
	
	' ǰ��� Display
	IF Request("txtItemCd") <> "" Then
		If (rs2.EOF And rs2.BOF) Then
			Call DisplayMsgBox("122600", vbOKOnly, "", "", I_MKSCRIPT)
			rs2.Close
			Set rs2 = Nothing
			Set ADF = Nothing
			%>
			<Script Language=vbscript>
				parent.frm1.txtItemCd.Focus()
			</Script>	
			<%
			Response.End
		Else
			%>
			<Script Language=vbscript>
				parent.frm1.txtItemNm.value = "<%=ConvSPChars(rs2("ITEM_NM"))%>"
			</Script>	
			<%
			rs2.Close
			Set rs2 = Nothing
		End If
	Else
		rs2.Close
		Set rs2 = Nothing
	End IF
	
	'Tracking_No
	IF Request("txtTrackingNo") <> "" Then
		If (rs3.EOF And rs3.BOF) Then
			Call DisplayMsgBox("203045", vbOKOnly, "", "", I_MKSCRIPT)
			rs3.Close
			Set rs3 = Nothing
			Set ADF = Nothing
			%>
			<Script Language=vbscript>
				parent.frm1.txtTrackingNo.Focus()
			</Script>	
			<%
			Response.End
		Else
			rs3.Close
			Set rs3 = Nothing
		End If
	Else
		rs3.Close
		Set rs3 = Nothing
	End IF
	
	'Shift_Cd
	IF Request("txtShiftCd") <> "" Then
		If (rs4.EOF And rs4.BOF) Then
			Call DisplayMsgBox("180400", vbOKOnly, "", "", I_MKSCRIPT)
			rs4.Close
			Set rs4 = Nothing
			Set ADF = Nothing
			%>
			<Script Language=vbscript>
				parent.frm1.txtShiftCd.Focus()
			</Script>	
			<%
			Response.End
		Else
			rs4.Close
			Set rs4 = Nothing
		End If
	Else
		rs4.Close
		Set rs4 = Nothing
	End IF
	
	' Item Group Check
	IF Request("txtItemGroupCd") <> "" Then
	 	If rs5.EOF AND rs5.BOF Then
			rs5.Close
			Set rs5 = Nothing
			Set ADF = Nothing
			strFlag = "ERROR_GROUP"
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtItemGroupNm.Focus() " & vbCrLf
			Response.Write "</Script>" & vbCrLf
			Response.End
		Else
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtItemGroupNm.value = """ & ConvSPChars(rs5("item_group_nm")) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf
			rs5.Close
			Set rs5 = Nothing
		End If
	Else
		rs5.Close
		Set rs5 = Nothing
	End If
	
		
	Redim UNISqlId(0)
	Redim UNIValue(0, 9)

	UNISqlId(0) = "p4425mb1H"
	
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = strPlantCd
	
	If CInt(strQryMode) = Cint(OPMD_UMODE) Then
		UNIValue(0, 2) = FilterVar(UCase(Request("lgStrPrevKey1")), "''", "S")
	Else
		UNIValue(0, 2) = strProdtOrderNo
	End If
	
	UNIValue(0, 3) = strReportFromDt
	UNIValue(0, 4) = strReportToDt 	
	UNIValue(0, 5) = strItemCd
	
	UNIValue(0, 6) = strTrackingNo
	UNIValue(0, 7) = strShiftCd
	UNIValue(0, 8) = strOrderStatus
	UNIValue(0, 9) = strItemGroupCd
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

	If rs0.EOF And rs0.BOF Then
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
	LngMaxRow = .frm1.vspdData1.MaxRows									'Save previous Maxrow
			
<%  
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
			strData = strData & Chr(11) & "<%=ConvSPChars(UCase(Trim(rs0("PRODT_ORDER_NO"))))%>"	
			strData = strData & Chr(11) & "<%=ConvSPChars(UCase(Trim(rs0("ITEM_CD"))))%>"	
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEM_NM"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(UCase(Trim(rs0("SPEC"))))%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("PRODT_ORDER_QTY"),ggQty.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("PROD_QTY_IN_ORDER_UNIT"),ggQty.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("GOOD_QTY_IN_ORDER_UNIT"),ggQty.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("BAD_QTY_IN_ORDER_UNIT"),ggQty.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("RCPT_QTY_IN_ORDER_UNIT"),ggQty.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(UCase(Trim(rs0("PRODT_ORDER_UNIT"))))%>"
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("PLAN_START_DT"))%>"
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("PLAN_COMPT_DT"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(UCase(Trim(rs0("TRACKING_NO"))))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ORDER_STATUS"))%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("ORDER_QTY_IN_BASE_UNIT"),ggQty.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("PROD_QTY_IN_BASE_UNIT"),ggQty.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("GOOD_QTY_IN_BASE_UNIT"),ggQty.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("BAD_QTY_IN_BASE_UNIT"),ggQty.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("RCPT_QTY_IN_BASE_UNIT"),ggQty.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(UCase(Trim(rs0("BASE_UNIT"))))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_group_cd"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_group_nm"))%>"
			strData = strData & Chr(11) & LngMaxRow + <%=i%>
			strData = strData & Chr(11) & Chr(12)
				
			TmpBuffer(<%=i%>) =  strData 
<%		
			rs0.MoveNext
			
		End If
	Next
%>
		iTotalStr = Join(TmpBuffer, "")
		.ggoSpread.Source = .frm1.vspdData1
		.ggoSpread.SSShowDataByClip iTotalStr
			
		.lgStrPrevKey1 = "<%=Trim(rs0("PRODT_ORDER_NO"))%>"
			
<%		
		rs0.Close
		Set rs0 = Nothing
%>
		
	.frm1.hPlantCd.value		= "<%=ConvSPChars(Request("txtPlantCd"))%>"
	.frm1.hReportFromDt.value	= "<%=Request("txtReportFromDt")%>"
	.frm1.hReportToDt.value		= "<%=Request("txtReportToDt")%>"
	.frm1.hProdOrderNo.value	= "<%=ConvSPChars(Request("txtProdOrderNo"))%>"
	.frm1.hItemCd.value			= "<%=ConvSPChars(Request("txtItemCd"))%>"
	.frm1.hTrackingNo.value		= "<%=ConvSPChars(Request("txtTrackingNo"))%>"
	.frm1.hShiftCd.value		= "<%=Request("txtShiftCd")%>"
	.frm1.hOrderStatus.value	= "<%=Request("cboOrderStatus")%>"
	.frm1.hItemGroupCd.value	= "<%=ConvSPChars(Request("txtItemGroupCd"))%>"
	.DbQueryOk	
		
End With	
</Script>	
<%
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
