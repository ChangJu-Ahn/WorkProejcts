<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'======================================================================================================
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
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call LoadBasisGlobalInf
Call loadInfTB19029B("I", "*", "NOCOOKIE","MB")

Dim ADF										'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg								'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter 선언 
Dim rs0, rs1, rs2, rs3, rs4, rs5, rs6		'DBAgent Parameter 선언 
Dim strQryMode
'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================
Dim i

Const C_SHEETMAXROWS_D = 100

Call HideStatusWnd

strQryMode = Request("lgIntFlgMode")

On Error Resume Next

Dim strStartDt
Dim strEndDt
Dim strItemCd
Dim strProdOrderNo
Dim strTrackingNo
Dim strOrderType
Dim strOrderStatus
Dim strMRPRunNo
Dim strItemGroupCd
Dim strFlag
Dim strBsitemcd
Err.Clear																	'☜: Protect system from crashing

	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================
	Redim UNISqlId(5)
	Redim UNIValue(5, 1)

	UNISqlId(0) = "180000saa"
	UNISqlId(1) = "180000sab"
	UNISqlId(2) = "180000sam"
	UNISqlId(3) = "180000sar"
	UNISqlId(4) = "180000sas"
	UNISqlId(5) = "180000sat"
	
	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(1, 0) = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	UNIValue(2, 0) = FilterVar(UCase(Request("txtTrackingNo")), "''", "S")
	UNIValue(3, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(3, 1) = FilterVar(UCase(Request("txtMRPRunNo")), "''", "S")
	UNIValue(4, 0) = FilterVar(UCase(Request("txtItemGroupCd")), "''", "S")
	UNIValue(5, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1, rs2, rs3, rs4, rs5, rs6)

	' Plant 명 Display      
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
	
	' Get OprCostFlag      
	If (rs6.EOF And rs6.BOF) Then
		rs6.Close
		Set rs6 = Nothing
		strFlag = "ERROR_PLANT_CONFIG"
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.hOprCostFlag.value = """"" & vbCrLf	
		Response.Write "</Script>" & vbCrLf
	Else
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.hOprCostFlag.value = """ & ConvSPChars(rs6("OPR_COST_FLAG")) & """" & vbCrLf
		Response.Write "</Script>" & vbCrLf
		rs6.Close
		Set rs6 = Nothing
	End If
	
	' 품목명 Display
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
	' MRP Run No. Check
	IF Request("txtMRPRunNo") <> "" Then
		If (rs4.EOF And rs4.BOF) Then
			rs4.Close
			Set rs4 = Nothing
			strFlag = "ERROR_RUNNO"
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
			strFlag = "ERROR_GROUP"
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtItemGroupNm.value = """"" & vbCrLf
			Response.Write "</Script>" & vbCrLf
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
		ElseIf strFlag = "ERROR_PLANT_CONFIG" Then
			Call DisplayMsgBox("180600", vbOKOnly, "", "", I_MKSCRIPT)
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
		ElseIf strFlag = "ERROR_RUNNO" Then
			Call DisplayMsgBox("187600", vbOKOnly, "", "", I_MKSCRIPT)
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtMRPRunNo.focus" & vbCrLf
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
	Redim UNIValue(0,11)

	UNISqlId(0) = "p4112mb1_ko441"  '2008-01-08::hanc
	
	IF Request("txtStartDt") = "" Then
		strStartDt = "|"
	Else
		strStartDt = " " & FilterVar(UNIConvDate(Request("txtStartDt")), "''", "S") & ""
	End IF

	IF Request("txtEndDt") = "" Then
		strEndDt = "|"
	Else
		strEndDt = " " & FilterVar(UNIConvDate(Request("txtEndDt")), "''", "S") & ""
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

	IF Request("rdoOrderStatus") = "" Then
		strOrderStatus = "|"
	Else
		strOrderStatus = FilterVar(UCase(Request("rdoOrderStatus")), "''", "S")
	End If
	
	IF Request("txtMRPRunNo") = "" Then
		strMRPRunNo = "|"
	Else
		strMRPRunNo = FilterVar(UCase(Request("txtMRPRunNo")), "''", "S")
	End IF

	IF Request("txtItemGroupCd") = "" Then
		strItemGroupCd = "|"
	Else
		strItemGroupCd = "b.item_group_cd in (select item_group_cd from ufn_P_ListItemGrp(" & FilterVar(Trim(Request("txtItemGroupCd"))	, "''", "S") & " ))"
	End IF

	If Trim(Request("txtBsitemcd")) = "" Then
		strBsitemcd = "|"
	Else
		strBsitemcd = strBsitemcd & " b.BASE_ITEM_CD = " & FilterVar(Trim(Request("txtBsitemcd")), " " , "S") 
	End If

	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 2) = strStartDt
	UNIValue(0, 3) = strEndDt
	UNIValue(0, 4) = strItemCd 
	UNIValue(0, 5) = strTrackingNo
	UNIValue(0, 6) = strProdOrderNo		
	UNIValue(0, 7) = strOrderType
	UNIValue(0, 8) = strOrderStatus
	UNIValue(0, 9) = strMRPRunNo
	UNIValue(0,10) = strItemGroupCd
	UNIValue(0,11) = strBsitemcd


	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		Set ADF = Nothing
		Response.End
	End If
	
%>

<Script Language=vbscript>
Dim LngMaxRow
Dim strData
Dim TmpBuffer
Dim iTotalStr
    	
With parent																	'☜: 화면 처리 ASP 를 지칭함 
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
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Prodt_Order_No"))%>"											'☆: Production Order No
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Item_Cd"))%>"												'☆: Item Code
				strData = strData & Chr(11) & ""																				'☆: Item Popup
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Item_Nm"))%>"												'☆: Item Name
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Spec"))%>"													'☆: Spec
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Prodt_Order_Qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Prodt_Order_Unit"))%>"										'☆: Unit
				strData = strData & Chr(11) & ""																				'☆: Unit Popup
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Order_Qty_In_Base_Unit"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Base_Unit"))%>"			
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("Plan_Start_Dt"))%>"									'☆: Planned Start Date
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("Plan_Compt_Dt"))%>"									'☆: Planned Completion Date
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Rout_No"))%>"												'☆: Routing
				strData = strData & Chr(11) & ""																				'☆: Routing Popup
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Sl_Cd"))%>"													'☆: Storage Location Code
				strData = strData & Chr(11) & ""																				'☆: Storage Location Popup
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Sl_Nm"))%>"													'☆: Storage Location Name			
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("WC_CD"))%>"
				strData = strData & Chr(11) & ""
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("WC_NM"))%>"
				strData = strData & Chr(11) & "<%=rs0("Order_Status")%>"
				strData = strData & Chr(11) & "<%=rs0("Order_Status")%>"
				strData = strData & Chr(11) & "<%=rs0("Re_Work_flg")%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Remark"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Bom_No"))%>"
				strData = strData & Chr(11) & "<%=rs0("Prodt_Order_Type")%>"
				strData = strData & Chr(11) & "<%=rs0("Prodt_Order_Type")%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Plan_Order_No"))%>"											'☆: Plan Order No.
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Tracking_No"))%>"											'☆: Tracking No.			
				strData = strData & Chr(11) & ""																				'☆: Tracking No Popup
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("Schd_Start_Dt"))%>"									'☆: Scheduled Start Date
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("Schd_Compt_Dt"))%>"									'☆: Scheduled Completion Date
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("Valid_From_DT"))%>"
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("Valid_To_DT"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Order_Unit_MFG"))%>"
				strData = strData & Chr(11) & "<%=rs0("Order_Lt_MFG")%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Fixed_MRP_Qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Min_MRP_Qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Max_MRP_Qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Round_Qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Scrap_Rate_MFG"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("MPS_Mgr"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("MRP_Mgr"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Prod_Mgr"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_group_cd"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_group_nm"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("MRP_RUN_NO"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PARENT_ORDER_NO"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PARENT_OPR_NO"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("COST_CD"))%>"
				strData = strData & Chr(11) & ""
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("COST_NM"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("OPR_NO"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("BASE_ITEM_CD"))%>"												'☆: Item Code
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("BASE_ITEM_NM"))%>"												'☆: Item Name
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
	If .frm1.vspdData.MaxRows < .VisibleRowCnt(.frm1.vspdData,0) and .lgStrPrevKey <> "" Then	' GroupView 사이즈로 화면 Row수보다 쿼리가 작으면 다시 쿼리함 
		Call .InitData(LngMaxRow)
		
		.DbQuery
	Else
		.frm1.hPlantCd.value	= "<%=ConvSPChars(Request("txtPlantCd"))%>"
		.frm1.hItemCd.value		= "<%=ConvSPChars(Request("txtItemCd"))%>"
		.frm1.hProdOrderNo.value= "<%=ConvSPChars(Request("txtProdOrderNo"))%>"
		.frm1.hProdFromDt.value	= "<%=Request("txtStartDt")%>"
		.frm1.hProdToDt.value	= "<%=Request("txtEndDt")%>"
		.frm1.hOrderType.value	= "<%=ConvSPChars(Request("cboOrderType"))%>"
		.frm1.hOrderStatus.value= "<%=ConvSPChars(Request("rdoOrderStatus"))%>"
		.frm1.hTrackingNo.value = "<%=ConvSPChars(Request("txtTrackingNo"))%>"
		.frm1.hMRPRunNo.value	= "<%=ConvSPChars(Request("txtMRPRunNo"))%>"
		.frm1.hItemGroupCd.value= "<%=ConvSPChars(Request("txtItemGroupCd"))%>"
		
		.DbQueryOk(LngMaxRow+1)
	End If

End With

</Script>	
<%
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
