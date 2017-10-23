<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : mc800qb1
'*  4. Program Name         : 납입지시현황조회 
'*  5. Program Desc         : 납입지시현황조회 
'*  6. Component List       : 
'*  7. Modified date(First) : 2003/02/27
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Lee Woo Guen
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call LoadBasisGlobalInf
Call loadInfTB19029B("Q", "M", "NOCOOKIE","MB")
On Error Resume Next

Dim ADF										'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg								'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter 선언 
Dim rs0, rs1, rs2, rs3, rs4, rs5			'DBAgent Parameter 선언 
Dim strQryMode								'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================
Dim i

Const C_SHEETMAXROWS = 50

Call HideStatusWnd

strQryMode = Request("lgIntFlgMode")

Dim strPlantCd
Dim strReqFromDt
Dim strReqToDt
Dim strItemCd
Dim strBpCd
Dim strProdOrderNo
Dim strPoNo
Dim strTrackingNo
Dim strDlvyOrderStatus
Dim strFlag
Dim PvArr

Err.Clear																	'☜: Protect system from crashing

'=======================================================================================================
'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
'=======================================================================================================
	Redim UNISqlId(4)
	Redim UNIValue(4, 0)

	UNISqlId(0) = "180000saa"					' Plant Check
	UNISqlId(1) = "180000sab"					' Item Code Check
	UNISqlId(2) = "m3111pa03"					' Biz Partner Check
	UNISqlId(3) = "mc300mb101"					' PO No Check	
	UNISqlId(4) = "180000sam"					' Tracking No Check
	
	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(1, 0) = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	UNIValue(2, 0) = FilterVar(UCase(Request("txtBpCd")), "''", "S")
	UNIValue(3, 0) = FilterVar(UCase(Request("txtPoNo")), "''", "S")
	UNIValue(4, 0) = FilterVar(UCase(Request("txtTrackingNo")), "''", "S")

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1, rs2, rs3, rs4, rs5)

	%>
	<Script Language=vbscript>
		parent.frm1.txtPlantNm.value = ""
		parent.frm1.txtItemNm.value = ""
		parent.frm1.txtBpNm.value = ""
	</Script>	
	<%    	
	
	' Plant 명 Display      
	If (rs1.EOF And rs1.BOF) Then
		rs1.Close
		Set rs1 = Nothing
		strFlag = "ERROR_PLANT"    	
	Else
		%>
		<Script Language=vbscript>
			parent.frm1.txtPlantNm.value = "<%=ConvSPChars(rs1("PLANT_NM"))%>"
		</Script>	
		<%    	
		rs1.Close
		Set rs1 = Nothing
	End If

		' 품목명 Display
	IF Request("txtItemCd") <> "" Then
		If (rs2.EOF And rs2.BOF) Then
			rs2.Close
			Set rs2 = Nothing
			strFlag = "ERROR_ITEM"
		Else
			%>
			<Script Language=vbscript>
				parent.frm1.txtItemNm.value = "<%=ConvSPChars(rs2("ITEM_NM"))%>"
			</Script>	
			<%
			rs2.Close
			Set rs2 = Nothing
		End If
	End IF

		' 공급처명 Display
	IF Request("txtBpCd") <> "" Then
		If (rs3.EOF And rs3.BOF) Then
			rs3.Close
			Set rs3 = Nothing
			strFlag = "ERROR_BPCd"
		Else
			%>
			<Script Language=vbscript>
				parent.frm1.txtBpNm.value = "<%=ConvSPChars(rs2("BP_NM"))%>"
			</Script>	
			<%
			rs3.Close
			Set rs3 = Nothing
		End If
	End IF

	' PO No Display
	IF Request("txtPoNo") <> "" Then
		If (rs4.EOF And rs4.BOF) Then
			rs4.Close
			Set rs4 = Nothing
			strFlag = "ERROR_PONo"
		Else
			rs4.Close
			Set rs4 = Nothing
		End If
	End IF
			
	' Tracking No Display
	IF Request("txtTrackingNo") <> "" Then
		If (rs5.EOF And rs5.BOF) Then
			rs5.Close
			Set rs5 = Nothing
			strFlag = "ERROR_TRACK"
		Else
			rs5.Close
			Set rs5 = Nothing
		End If
	End IF

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
		ElseIf strFlag = "ERROR_BPCd" Then
			Call DisplayMsgBox("179021", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
			parent.frm1.txtBpCd.Focus()
			</Script>	
			<%
			Set ADF = Nothing
			Response.End
		ElseIf strFlag = "ERROR_PONo" Then
			Call DisplayMsgBox("173100", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
			parent.frm1.txtPoNo.Focus()
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
	Redim UNIValue(0, 9)

	UNISqlId(0) = "mc800qb1a"
	
	StrPlantCd = FilterVar(UCase(Request("txtPlantCd")), "''", "S")

	IF Request("txtReqFromDt") = "" Then
		strReqFromDt = "|"
	Else
		strReqFromDt = " " & FilterVar(UNIConvDate(Request("txtReqFromDt")), "''", "S") & ""
	End IF

	IF Request("txtReqToDt") = "" Then
		strReqToDt = "|"
	Else
		strReqToDt = " " & FilterVar(UNIConvDate(Request("txtReqToDt")), "''", "S") & ""
	End IF

	IF Request("txtItemCd") = "" Then
		strItemCd = "|"
	Else
		strItemCd = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	End IF

	IF Request("txtBpCd") = "" Then
		strBpCd = "|"
	Else
		strBpCd = FilterVar(UCase(Request("txtBpCd")), "''", "S")
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
		
	IF Request("txtPoNo") = "" Then
		strPoNo = "|"
	Else
		StrPoNo = FilterVar(UCase(Request("txtPoNo")), "''", "S")
	End IF

	IF Request("txtTrackingNo") = "" Then
		strTrackingNo = "|"
	Else
		strTrackingNo = FilterVar(UCase(Request("txtTrackingNo")), "''", "S")
	End IF

	IF Request("cboDlvyOrderStatus") = "" Then
		strDlvyOrderStatus = "|"
	Else
		strDlvyOrderStatus = " " & FilterVar(UCase(Request("cboDlvyOrderStatus")), "''", "S") & ""
	End IF

	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 2) = strReqFromDt
	UNIValue(0, 3) = strReqToDt
	UNIValue(0, 4) = strItemCd 
	UNIValue(0, 5) = strBpCd
	UNIValue(0, 6) = strProdOrderNo		
	UNIValue(0, 7) = strPoNo
	UNIValue(0, 8) = strTrackingNo
	UNIValue(0, 9) = strDlvyOrderStatus

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
      
	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		Set ADF = Nothing
		Response.Write "<Script Language=VBScript>" & vbCr
		Response.Write "	Parent.frm1.txtPlantCd.focus " & vbCr
		Response.Write "	Set Parent.gActiveElement = Parent.document.activeElement    " & vbCr
		Response.Write "</Script>" & vbCr
		Response.End													'☜: 비지니스 로직 처리를 종료함 
	End If
	
%>

<Script Language=vbscript>
Dim LngMaxRow
Dim strData
    	
With parent																	'☜: 화면 처리 ASP 를 지칭함 
	LngMaxRow = .frm1.vspdData.MaxRows										'Save previous Maxrow

<%  
    ReDim PvArr(C_SHEETMAXROWS - 1)
	If Not(rs0.EOF And rs0.BOF) Then
		For i=0 to rs0.RecordCount-1
			If i < C_SHEETMAXROWS Then 
%>
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Prodt_Order_No"))%>"						'☆: Production Order No
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Item_Cd"))%>"							'☆: Item Code
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Item_Nm"))%>"							'☆: Item Description
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Spec"))%>"								'☆: Specification
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("Req_Dt"))%>"						'☆: Required Date
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Req_Qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Base_Unit"))%>"							'☆: Base Unit
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Do_Qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Rcpt_Qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Bp_Cd"))%>"								'☆: Biz Partner
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Bp_Nm"))%>"								'☆: Biz Partner Description
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("Do_Date"))%>"					'☆: Delivery Date
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Do_Time"))%>"							'☆: Delivery Time
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Minor_Nm_Do_Time"))%>"					'☆: Delivery Time Description
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Do_Status"))%>"							'☆: Delivery Status
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Minor_Nm_Do_Status"))%>"					'☆: Delivery Status Description
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Tracking_No"))%>"						'☆: Tracking No
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Po_No"))%>"								'☆: PO No
			strData = strData & Chr(11) & "<%=rs0("Po_Seq_No")%>"										'☆: PO Seq No
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Do_Qty_Po_Unit"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Rcpt_Qty_Po_Unit"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Po_Unit"))%>"							'☆: PO Unit
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Opr_No"))%>"								'☆: Operation No
			strData = strData & Chr(11) & "<%=rs0("Seq")%>"												'☆: Operation Seq No
			strData = strData & Chr(11) & "<%=rs0("Sub_Seq")%>"											'☆: Operation Sub Seq No
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Wc_Cd"))%>"								'☆: Work Center
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Wc_Nm"))%>"								'☆: Work Center Description
            strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("Plan_Start_Dt"))%>"				'☆: Plan Start Date
            strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("Plan_Compt_Dt"))%>"				'☆: Plan Compt Date
            strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("Release_Dt"))%>"					'☆: Release Date
			
			strData = strData & Chr(11) & LngMaxRow + <%=i%>
			strData = strData & Chr(11) & Chr(12)

<%		
			rs0.MoveNext

			PvArr(i) = strData	
			strData = ""
			End If
		Next

		strData  = Join(PvArr, "")
%>
		.ggoSpread.Source = .frm1.vspdData
		.ggoSpread.SSShowData strData
		
		.lgStrPrevKey = "<%=Trim(rs0("PRODT_ORDER_NO"))%>"
		
<%	
	End If

	rs0.Close
	Set rs0 = Nothing

%>	
	If .frm1.vspdData.MaxRows < .VisibleRowCnt(.frm1.vspdData,0) and .lgStrPrevKey <> "" Then	<% ' GroupView 사이즈로 화면 Row수보다 쿼리가 작으면 다시 쿼리함 %>
		Call .InitData(LngMaxRow)
		.DbQuery
	Else
		.frm1.hPlantCd.value		= "<%=ConvSPChars(Request("txtPlantCd"))%>"
		.frm1.hReqFromDt.value		= "<%=UNIDateClientFormat(Request("txtReqFromDt"))%>"
		.frm1.hReqToDt.value		= "<%=UNIDateClientFormat(Request("txtReqToDt"))%>"
		.frm1.hItemCd.value			= "<%=ConvSPChars(Request("txtItemCd"))%>"
		.frm1.hBpCd.value			= "<%=ConvSPChars(Request("txtBpCd"))%>"
		.frm1.hProdOrderNo.value	= "<%=ConvSPChars(Request("txtProdOrderNo"))%>"
		.frm1.hPoNo.value			= "<%=ConvSPChars(Request("txtPoNo"))%>"
		.frm1.hTrackingNo.value		= "<%=ConvSPChars(Request("txtTrackingNo"))%>"
		.frm1.hDlvyOrderStatus.value	= "<%=ConvSPChars(Request("cboDlvyOrderStatus"))%>"		
		.DbQueryOk(LngMaxRow+1)
	End If

End With

</Script>	
<%
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
