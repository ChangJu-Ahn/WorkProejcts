<%@LANGUAGE = VBScript%>
<%'*******************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4111pb1.asp
'*  4. Program Name         : List Production Order Header (Query)
'*  5. Program Desc         :
'*  6. Comproxy List        : +P32118ListProdOrderHeader
'*  7. Modified date(First) : 2000/03/28
'*  8. Modified date(Last)  : 2002/12/10
'*  9. Modifier (First)     : Park, BumSoo
'* 10. Modifier (Last)      : RYU SUNG WON
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
Dim rs0, rs1, rs2, rs3, rs4, rs5			'DBAgent Parameter ���� 
Dim strQryMode
Dim i

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Const C_SHEETMAXROWS_D = 100

Call HideStatusWnd

Dim strStartDt
Dim strEndDt
Dim strItemCd
Dim strProdOrderNo
Dim strTrackingNo
Dim strOrderStatus
Dim strOrderStatus1, strOrderStatus2, strOrderStatus3
Dim strItemGroupCd
Dim strNextKey
Dim strFlag

strQryMode = Request("lgIntFlgMode")

On Error Resume Next
Err.Clear																	'��: Protect system from crashing

	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	Redim UNISqlId(4)
	Redim UNIValue(4, 0)

	UNISqlId(0) = "180000saa"
	UNISqlId(1) = "180000sab"
	UNISqlId(2) = "180000sas"
	UNISqlId(3) = "180000sab"
	UNISqlId(4) = "180000sac"
	
	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(1, 0) = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	UNIValue(2, 0) = FilterVar(Ucase(Request("txtItemGroupCd")), "''", "S")
	UNIValue(3, 0) = FilterVar(UCase(Request("txtChildItemCd")), "''", "S")
	UNIValue(4, 0) = FilterVar(UCase(Request("txtWcCd")), "''", "S")
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1, rs2, rs3, rs4, rs5)

	%>
	<Script Language=vbscript>
		parent.txtItemNm.value = ""
	</Script>	
	<%    	

	' Plant �� Display      
	If (rs1.EOF And rs1.BOF) Then
		rs1.Close
		Set rs1 = Nothing
		strFlag = "ERROR_PLANT"
	Else
		rs1.Close
		Set rs1 = Nothing
	End If

	' ��ǰ��� Display
	IF Request("txtItemCd") <> "" Then
		If (rs2.EOF And rs2.BOF) Then
			rs2.Close
			Set rs2 = Nothing
			strFlag = "ERROR_ITEM"
			%>
			<Script Language=vbscript>
				parent.txtItemNm.value = ""
			</Script>	
			<%
		Else
			%>
			<Script Language=vbscript>
				parent.txtItemNm.value = "<%=ConvSPChars(rs2("ITEM_NM"))%>"
			</Script>	
			<%
			rs2.Close
			Set rs2 = Nothing
		End If
	End IF
	
	' ��ǰ��� Display
	If (rs4.EOF And rs4.BOF) Then
		rs4.Close
		Set rs4 = Nothing
		strFlag = "ERROR_CHILD_ITEM"
		%>
		<Script Language=vbscript>
			parent.txtChildItemNm.value = ""
		</Script>	
		<%
	Else
		%>
		<Script Language=vbscript>
			parent.txtChildItemNm.value = "<%=ConvSPChars(rs4("ITEM_NM"))%>"
		</Script>	
		<%
		rs4.Close
		Set rs4 = Nothing
	End If
	
	' �۾��� Display
	If (rs5.EOF And rs5.BOF) Then
		rs5.Close
		Set rs5 = Nothing
		strFlag = "ERROR_WC"
		%>
		<Script Language=vbscript>
			parent.txtWcNm.value = ""
		</Script>	
		<%
	Else
		%>
		<Script Language=vbscript>
			parent.txtWcNm.value = "<%=ConvSPChars(rs5("WC_NM"))%>"
		</Script>	
		<%
		rs5.Close
		Set rs5 = Nothing
	End If
		
	If strFlag <> "" Then
		If strFlag = "ERROR_PLANT" Then
			Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
			Set ADF = Nothing
			Response.End
		ElseIf strFlag = "ERROR_ITEM" Then
			Call DisplayMsgBox("122600", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
			parent.txtItemCd.Focus()
			</Script>	
			<%
			Set ADF = Nothing
			Response.End
		ElseIf strFlag = "ERROR_CHILD_ITEM" Then
			Call DisplayMsgBox("122600", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
			parent.txtChildItemCd.Focus()
			</Script>	
			<%
			Set ADF = Nothing
			Response.End
		ElseIf strFlag = "ERROR_WC" Then
			Call DisplayMsgBox("182100", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
			parent.txtWcCd.Focus()
			</Script>	
			<%
			Set ADF = Nothing
			Response.End	
		End If
	End IF
	
	' Item Group Check
	IF Request("txtItemGroupCd") <> "" Then
	 	If rs3.EOF AND rs3.BOF Then
			rs3.Close
			Set rs3 = Nothing
			strFlag = "ERROR_GROUP"
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.txtItemGroupNm.value = """"" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		Else
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.txtItemGroupNm.value = """ & ConvSPChars(rs3("item_group_nm")) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf
			rs3.Close
			Set rs3 = Nothing
		End If
	Else
		rs3.Close
		Set rs3 = Nothing
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.txtItemGroupNm.value = """"" & vbCrLf
		Response.Write "</Script>" & vbCrLf
	End If

	' Order Header Display
	Redim UNISqlId(0)
	Redim UNIValue(0, 11)

	UNISqlId(0) = "p4211pb1"
	
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

	If Request("txtProdOrderNo") = "" Then
		strProdOrderNo = "|"
	Else
		strProdOrderNo = FilterVar(UCase(Request("txtProdOrderNo")), "''", "S")
	End If	

	IF Request("txtFromStstus") = "" Then
		strOrderStatus1 = "|"
	Else
		strOrderStatus1 = " " & FilterVar(UCase(Request("txtFromStstus")), "''", "S") & ""
	End IF

	IF Request("txtToStstus") = "" Then
		strOrderStatus2 = "|"
	Else
		strOrderStatus2 = " " & FilterVar(UCase(Request("txtToStstus")), "''", "S") & ""
	End IF
	
	IF Request("txtThirdStstus") = "" Then
		strOrderStatus3 = "|"
	Else
		strOrderStatus3 = " " & FilterVar(UCase(Request("txtThirdStstus")), "''", "S") & ""
	End IF

	IF strOrderStatus1 = "|" and strOrderStatus2 = "|" and strOrderStatus3 = "|" Then
		strOrderStatus = "|"
	Else
		strOrderStatus = ""
		strOrderStatus = "a.order_status in ( " & "" & strOrderStatus1 & "" & "," & "" & strOrderStatus2 & "" & "," & "" & strOrderStatus3 & "" & " ) "
	End If
	
	IF Request("txtItemGroupCd") = "" Then
		strItemGroupCd = "|"
	Else
		strItemGroupCd = "b.item_group_cd in (select item_group_cd from ufn_P_ListItemGrp(" & FilterVar(Trim(Request("txtItemGroupCd"))	,"''", "S") & " ))"
	End IF
	
	Select Case strQryMode
		Case CStr(OPMD_CMODE)
			strNextKey = "|"	
		Case CStr(OPMD_UMODE) 
			strNextKey = " ( a.prodt_order_no > " &  FilterVar(UCase(Request("lgStrPrevKey1")), "''", "S") 
			strNextKey = strNextKey & " Or (a.prodt_order_no = " &  FilterVar(UCase(Request("lgStrPrevKey1")), "''", "S") 
			strNextKey = strNextKey & " And b.opr_no > " & FilterVar(Request("lgStrPrevKey2"), "''", "S") & ")"
			strNextKey = strNextKey & " Or (a.prodt_order_no = " &  FilterVar(UCase(Request("lgStrPrevKey1")), "''", "S") 
			strNextKey = strNextKey & " And b.opr_no = " & FilterVar(Request("lgStrPrevKey2"), "''", "S") 
			strNextKey = strNextKey & " And b.seq >= " & FilterVar(Request("lgStrPrevKey3"), "''", "S") & ") ) "
	End Select

	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 2) = FilterVar(UCase(Request("txtChildItemCd")), "''", "S")
	UNIValue(0, 3) = FilterVar(UCase(Request("txtWcCd")), "''", "S")
	UNIValue(0, 4) = strItemCd
	UNIValue(0, 5) = strTrackingNo
	UNIValue(0, 6) = strProdOrderNo	
	UNIValue(0, 7) = strStartDt
	UNIValue(0, 8) = strEndDt
	UNIValue(0, 9) = strOrderStatus
	UNIValue(0, 10) = strItemGroupCd
	UNIValue(0, 11) = strNextKey

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
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Prodt_Order_No"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("opr_no"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("job_nm"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("sl_cd"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("sl_nm"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("tracking_no"))%>"
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("req_dt"))%>"
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("req_qty"),ggQty.DecPoint,0)%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("basic_unit"))%>"
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("issued_qty"),ggQty.DecPoint,0)%>"
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("consumed_qty"),ggQty.DecPoint,0)%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("req_no"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_cd"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_nm"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("spec"))%>"
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("plan_start_dt"))%>"
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("plan_compt_dt"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("order_status"))%>"
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("prodt_order_qty"),ggQty.DecPoint,0)%>"
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("prod_qty_in_order_unit"),ggQty.DecPoint,0)%>"
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("good_qty_in_order_unit"),ggQty.DecPoint,0)%>"
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("rcpt_qty_in_order_unit"),ggQty.DecPoint,0)%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("prodt_order_unit"))%>"
				If "<%=rs0("Re_Work_Flg")%>" = "N" Then
					strData = strData & Chr(11) & "�۾�"
				Else
					strData = strData & Chr(11) & "���۾�"
				End If
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
		
	.lgStrPrevKey1 = "<%=Trim(rs0("Prodt_Order_No"))%>"
	.lgStrPrevKey2 = "<%=Trim(rs0("opr_no"))%>"
	.lgStrPrevKey3 = "<%=Trim(rs0("seq"))%>"
	
<%	
	End If

	rs0.Close
	Set rs0 = Nothing

%>	
	
	If .vspdData.MaxRows < .PopupParent.VisibleRowCnt(.vspdData,0) and .lgStrPrevKey <> "" Then	 ' GroupView ������� ȭ�� Row������ ������ ������ �ٽ� ������ 
		.InitData(LngMaxRow)
		.DbQuery
	Else
		.hItemCd.value		= "<%=ConvSPChars(Request("txtItemCd"))%>"
		.hProdOrderNo.value	= "<%=ConvSPChars(Request("txtProdOrderNo"))%>"
		.hProdFromDt.value	= "<%=Request("txtFromDt")%>"
		.hProdToDt.value	= "<%=Request("txtToDt")%>"
		.hFromStatus.value	= "<%=ConvSPChars(Request("txtFromStstus"))%>"
		.hToStatus.value	= "<%=ConvSPChars(Request("txtToStstus"))%>"
		.hTrackingNo.value	= "<%=ConvSPChars(Request("txtTrackingNo"))%>"
		.hitemGroupCd.value	= "<%=ConvSPChars(Request("txtItemGroupCd"))%>"
		
		.DbQueryOk(LngMaxRow)
	End If  
    
    End With
</Script>	
