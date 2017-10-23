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
On Error Resume Next

Dim ADF
Dim strRetMsg
Dim UNISqlId, UNIValue, UNILock, UNIFlag
Dim rs0, rs1, rs2, rs3, rs4
Dim strQryMode
Dim StrNextKey
Dim strItemCd
Dim strProdOrderNo
Dim strTrackingNo
Dim strItemGroupCd
Dim strInfNo
Dim strFlag
Dim i

	Const C_SHEETMAXROWS_D = 100

	Call HideStatusWnd

	strQryMode = Request("lgIntFlgMode")

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
	
	UNIValue(0, 0) = FilterVar(Ucase(Trim(Request("txtPlantCd"))),"''","S")
	UNIValue(1, 0) = FilterVar(Ucase(Trim(Request("txtItemCd"))),"''","S")
	UNIValue(2, 0) = FilterVar(Ucase(Trim(Request("txtTrackingNo"))),"''","S")
	UNIValue(3, 0) = FilterVar(Ucase(Trim(Request("txtItemGroupCd"))), "''", "S")
	
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
	Redim UNIValue(0, 8)

	UNISqlId(0) = "p44B1mb1"

	IF Request("txtItemCd") = "" Then
		strItemCd = "|"
	Else
		StrItemCd = FilterVar(Ucase(Trim(Request("txtItemCd"))),"''","S")
	End IF

	IF Request("txtTrackingNo") = "" Then
		strTrackingNo = "|"
	Else
		strTrackingNo = FilterVar(Ucase(Trim(Request("txtTrackingNo"))),"''","S")
	End IF
	
	If Request("txtProdOrderNo") = "" Then
		strProdOrderNo = "|"
	Else
		strProdOrderNo = FilterVar(Ucase(Trim(Request("txtProdOrderNo"))),"''","S")
	End If
	
	Select Case strQryMode
		Case CStr(OPMD_CMODE)
			strInfNo = "|"	
		Case CStr(OPMD_UMODE) 
			strInfNo = FilterVar(Ucase(Trim(Request("lgStrPrevKey"))),"''","S")
	End Select 

	IF Request("txtItemGroupCd") = "" Then
		strItemGroupCd = "|"
	Else
		strItemGroupCd = " c.item_group_cd in (select item_group_cd from ufn_P_ListItemGrp(" & FilterVar(Trim(Request("txtItemGroupCd"))	,"''", "S") & " ))"
	End IF

	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(Ucase(Trim(Request("txtPlantCd"))),"''","S")
	UNIValue(0, 2) = FilterVar(UNIConvDate(Request("txtFromDt")),"''","S")
	UNIValue(0, 3) = FilterVar(UNIConvDate(Request("txtToDt")),"''","S")
	UNIValue(0, 4) = strItemCd 
	UNIValue(0, 5) = strTrackingNo
	UNIValue(0, 6) = strProdOrderNo	
	UNIValue(0, 7) = strInfNo
	UNIValue(0, 8) = strItemGroupCd

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
      
	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		Set ADF = Nothing
		Response.End													'☜: 비지니스 로직 처리를 종료함 
	End If
	
%>

<Script Language=vbscript>
Dim LngMaxRow
Dim strData
Dim TmpBuffer
Dim iTotalStr
Dim strCur
    	
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
			strData = strData & Chr(11) & ""
			strData = strData & Chr(11) & Trim("<%=ConvSPChars(rs0("pop_inf_key"))%>")									
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("plant_cd"))%>"
<%			
			'If rs0("input_type") = "H" Then
%>											
			'	strData = strData & Chr(11) & "오더별"
<%			'ElseIf  rs0("input_type") = "D" Then			
%>
			'	strData = strData & Chr(11) & "공정별"			
<%			'Else
%>
			'	strData = strData & Chr(11) & ""	
<%			'End If
%>			
			strData = strData & Chr(11) & "<%=rs0("input_type")%>"
															
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("prodt_order_no"))%>"											
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_cd"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_nm"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("spec"))%>"											
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("opr_no"))%>"											
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("report_type"))%>"											
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("pop_qty"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("pop_unit"))%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("prodt_order_qty"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("prod_qty"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("report_dt"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("shift_cd"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("description"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("reason_cd"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("reason_nm"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("tracking_no"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("lot_no"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("lot_sub_no"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("rcpt_item_document_no"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("remark"))%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("subcontract_prc"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("C_subcontract_amt"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
			'strData = strData & Chr(11) & "<%=UniConvNumDBToCompanyWithOutChange(rs0("subcontract_prc"), 0)%>"
			'strData = strData & Chr(11) & "<%=UniConvNumDBToCompanyWithOutChange(rs0("subcontract_amt"), 0)%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("cur_cd"))%>"
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
		
		'Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData, LngMaxRow + 1, LngMaxRow + <%=i%>, .parent.gCurrency,.C_subcontract_prc, "C", "I", "X", "X")
		'Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData, LngMaxRow + 1, LngMaxRow + <%=i%>, .parent.gCurrency,.C_subcontract_amt, "A", "I", "X", "X")
		
		.lgStrPrevKey = "<%=Trim(rs0("pop_inf_key"))%>"
		
<%	
	End If

	rs0.Close
	Set rs0 = Nothing

%>	
	
	.frm1.hPlantCd.value	= "<%=ConvSPChars(Request("txtPlantCd"))%>"
	.frm1.hItemCd.value		= "<%=ConvSPChars(Request("txtItemCd"))%>"
	.frm1.hProdOrderNo.value= "<%=ConvSPChars(Request("txtProdOrderNo"))%>"
	.frm1.hFromDt.value		= "<%=Request("txtFromDt")%>"
	.frm1.hToDt.value		= "<%=Request("txtToDt")%>"
	.frm1.hTrackingNo.value	= "<%=ConvSPChars(Request("txtTrackingNo"))%>"
	.frm1.hItemGroupCd.value= "<%=ConvSPChars(Request("txtItemGroupCd"))%>"		
	.DbQueryOk(LngMaxRow)

End With

</Script>	
<%
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>