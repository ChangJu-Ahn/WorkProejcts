<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4611mb1.asp
'*  4. Program Name         : List Production Order (Query)
'*  5. Program Desc         : List ASP used by Order Closing
'*  6. Comproxy List        : DB Agent (p4611mb1)
'*  7. Modified date(First) : 2000/03/28
'*  8. Modified date(Last)  : 2003/09/12
'*  9. Modifier (First)     : Park, Bum Soo
'* 10. Modifier (Last)      : Park, Bum Soo
'* 11. Comment              : 
'**********************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call LoadBasisGlobalInf
Call loadInfTB19029B("I", "*", "NOCOOKIE", "MB")

On Error Resume Next

Dim ADF
Dim strRetMsg
Dim UNISqlId, UNIValue, UNILock, UNIFlag
Dim rs0, rs1, rs2, rs3, rs4
Dim strQryMode
Dim strFlag
Dim strItemCd
Dim strProdOrderNo
Dim strTrackingNo
Dim strOrderType
Dim strOrderStatus
Dim strProdMgr
Dim strItemGroupCd

Dim i

	Const C_SHEETMAXROWS_D = 100

	Call HideStatusWnd

	strQryMode = Request("lgIntFlgMode")

	On Error Resume Next

	Err.Clear

	'=======================================================================================================
	'	Handle Description
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
	'=======================================================================================================
	'	Main Query - Order Header Display
	'=======================================================================================================
	Redim UNISqlId(0)
	Redim UNIValue(0, 9)

	UNISqlId(0) = "p4611mb1_nepes"
	
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

	IF Request("cboProdMgr") = "" Then
		strProdMgr = "|"
	Else
		strProdMgr = " " & FilterVar(UCase(Request("cboProdMgr")), "''", "S") & ""
	End IF

	IF Request("txtItemGroupCd") = "" Then
		strItemGroupCd = "|"
	Else
		strItemGroupCd = "b.item_group_cd in (select item_group_cd from ufn_P_ListItemGrp(" & FilterVar(Trim(Request("txtItemGroupCd"))	, "''", "S") & " ))"
	End IF

	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 2) = FilterVar(UCase(Request("txtFromDt")), "''", "S")
	UNIValue(0, 3) = FilterVar(UCase(Request("txtToDt")), "''", "S")
	UNIValue(0, 4) = strItemCd 
	UNIValue(0, 5) = strTrackingNo
	UNIValue(0, 6) = strProdOrderNo		
	UNIValue(0, 7) = strOrderType
	UNIValue(0, 8) = strProdMgr
	UNIValue(0, 9) = strItemGroupCd
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
	
	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		Response.End													'☜: 비지니스 로직 처리를 종료함 
	End If
	
%>

<Script Language=vbscript>
Dim LngMaxRow
Dim strData
Dim DblRemain
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
			strData = strData & Chr(11) & ""																		'☆: Check Box
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Prodt_Order_No"))%>"
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("Real_Compt_Dt"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Item_Cd"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Item_Nm"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Spec"))%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Prodt_Order_Qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Prodt_Order_Unit"))%>"
			DblRemain = "<%=UniConvNumberDBToCompany(CDbl(rs0("Prodt_Order_Qty")) - CDbl(rs0("Prod_Qty_In_Order_Unit")),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
			If DblRemain >= "0" Then
				strData = strData & Chr(11) & DblRemain
			Else
				strData = strData & Chr(11) & "0"
			End If
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("Plan_Start_Dt"))%>"
			strData = strData	 & Chr(11) & "<%=UNIDateClientFormat(rs0("Plan_Compt_Dt"))%>"
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("Schd_Start_Dt"))%>"
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("Schd_Compt_Dt"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Rout_No"))%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Prod_Qty_In_Order_Unit"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Good_Qty_In_Order_Unit"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Rcpt_Qty_In_Order_Unit"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("Release_Dt"))%>"
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("Real_Start_Dt"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Tracking_No"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Order_Status_Desc"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Order_Type_Desc"))%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Bad_Qty_In_Order_Unit"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Insp_Good_Qty_In_Order_Unit"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Insp_Bad_Qty_In_Order_Unit"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany((Cdbl(rs0("Good_Qty_In_Order_Unit")) - Cdbl(rs0("Rcpt_Qty_In_Order_Unit"))),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Base_Unit"))%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Order_Qty_In_Base_Unit"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Prod_Qty_In_Base_Unit"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Good_Qty_In_Base_Unit"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Bad_Qty_In_Base_Unit"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Insp_Good_Qty_In_Base_Unit"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Insp_Bad_Qty_In_Base_Unit"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Rcpt_Qty_In_Base_Unit"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany((Cdbl(rs0("Good_Qty_In_Base_Unit")) - Cdbl(rs0("Rcpt_Qty_In_Base_Unit"))),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("prod_mgr"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_group_cd"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_group_nm"))%>"
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
	If .frm1.vspdData.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData,0) and .lgStrPrevKey <> "" Then
		.DbQuery
	Else
		.frm1.hPlantCd.value	= "<%=ConvSPChars(Request("txtPlantCd"))%>"
		.frm1.hItemCd.value		= "<%=ConvSPChars(Request("txtItemCd"))%>"
		.frm1.hProdOrderNo.value= "<%=ConvSPChars(Requset("txtProdOrderNo"))%>"
		.frm1.hOrderType.value	= "<%=Request("cboOrderType")%>"
		.frm1.hProdFromDt.value	= "<%=Request("txtFromDt")%>"
		.frm1.hProdToDt.value	= "<%=Request("txtToDt")%>"
		.frm1.hTrackingNo.value = "<%=ConvSPChars(Request("txtTrackingNo"))%>"
		.frm1.hProdMgr.value	= "<%=Request("cboProdMgr")%>"
		.frm1.hItemGroupCd.value= "<%=ConvSPChars(Request("txtItemGroupCd"))%>"		
		.DbQueryOk
	End If

End With

</Script>	
