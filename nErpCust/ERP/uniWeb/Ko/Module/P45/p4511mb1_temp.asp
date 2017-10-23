<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4511mb1.asp
'*  4. Program Name         : List Production Results for Goods Receipt
'*  5. Program Desc         : 
'*  6. Modified date(First) : 2000/04/21
'*  7. Modified date(Last)  : 2003/09/12
'*  8. Modifier (First)     : Park, Bum Soo
'*  9. Modifier (Last)      : Park, Bum Soo
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
Dim rs0, rs1, rs2, rs3, rs4, rs5, rs6
Dim strQryMode
Dim lgStrPrevKey
Dim lgStrPrevKey2
Dim strFlag
Dim strItemCd
Dim StrProdOrderNo
Dim StrWcCd
Dim StrTrackingNo
Dim StrSlCd
Dim strItemGroupCd
Dim strTemp
Dim i

	Const C_SHEETMAXROWS_D = 100

	Call HideStatusWnd

	strQryMode = Request("lgIntFlgMode")

	lgStrPrevKey = FilterVar(UCase(Request("lgStrPrevKey")), "''", "S")
	lgStrPrevKey2 = FilterVar(UCase(Request("lgStrPrevKey2")), "''", "S")

	Err.Clear

	'=======================================================================================================
	'	Handle Description and Check Existence
	'=======================================================================================================
	Redim UNISqlId(5)
	Redim UNIValue(5, 1)

	UNISqlId(0) = "180000saa"
	UNISqlId(1) = "180000saf"
	UNISqlId(2) = "180000sac"
	UNISqlId(3) = "180000sad"
	UNISqlId(4) = "180000sam"
	UNISqlId(5) = "180000sas"
	
	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(1, 0) = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	UNIValue(1, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(2, 0) = FilterVar(UCase(Request("txtWcCd")), "''", "S")
	UNIValue(3, 0) = FilterVar(UCase(Request("txtSlCd")), "''", "S")
	UNIValue(4, 0) = FilterVar(UCase(Request("txtTrackingNo")), "''", "S")
	UNIValue(5, 0) = FilterVar(UCase(Request("txtItemGroupCd")), "''", "S")

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1, rs2, rs3, rs4, rs5, rs6)
	Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing

	' Plant Check
	If (rs1.EOF And rs1.BOF) Then
		rs1.Close
		Set rs1 = Nothing
		strFlag = "ERROR_PLANT"
		Response.Write "<Script Language=vbscript>" & vbCrLf
		Response.Write "parent.frm1.txtPlantNm.value = """"" & vbCrLf
		Response.Write "</Script>" & vbCrLf
	Else
		Response.Write "<Script Language=vbscript>" & vbCrLf
		Response.Write "parent.frm1.txtPlantNm.value = """ & ConvSPChars(rs1("PLANT_NM")) & """" & vbCrLf
		Response.Write "</Script>" & vbCrLf
		rs1.Close
		Set rs1 = Nothing
	End If
	' Item Check
	IF Request("txtItemCd") <> "" Then
		If (rs2.EOF And rs2.BOF) Then
			rs2.Close
			Set rs2 = Nothing
			strFlag = "ERROR_ITEM"
			Response.Write "<Script Language=vbscript>" & vbCrLf
			Response.Write "parent.frm1.txtItemNm.value = """"" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		Else
			Response.Write "<Script Language=vbscript>" & vbCrLf
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
	' Work Center Check
	IF Request("txtWcCd") <> "" Then
		If (rs3.EOF And rs3.BOF) Then
			rs3.Close
			Set rs3 = Nothing
			strFlag = "ERROR_WCCD"
			Response.Write "<Script Language=vbscript>" & vbCrLf
			Response.Write "parent.frm1.txtWCNm.value = """"" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		Else
			Response.Write "<Script Language=vbscript>" & vbCrLf
			Response.Write "parent.frm1.txtWCNm.value = """ & ConvSPChars(rs3("WC_NM")) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf
			rs3.Close
			Set rs3 = Nothing
		End If
	Else
		rs3.Close
		Set rs3 = Nothing
		Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtWCNm.value = """"" & vbCrLf
		Response.Write "</Script>" & vbCrLf
	End IF
	' Storage Location Check
	IF Request("txtSlCd") <> "" Then
		If (rs4.EOF And rs4.BOF) Then
			rs4.Close
			Set rs4 = Nothing
			strFlag = "ERROR_SLCD"
			Response.Write "<Script Language=vbscript>" & vbCrLf
			Response.Write "parent.frm1.txtSLNm.value = """"" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		Else
			Response.Write "<Script Language=vbscript>" & vbCrLf
			Response.Write "parent.frm1.txtSLNm.value = """ & ConvSPChars(rs4("SL_NM")) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf
			rs4.Close
			Set rs4 = Nothing
		End If
	Else
		rs4.Close
		Set rs4 = Nothing
		Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtSLNm.value = """"" & vbCrLf
		Response.Write "</Script>" & vbCrLf
	End IF
	' Tracking No. Check
	IF Request("txtTrackingNo") <> "" Then
		If (rs5.EOF And rs5.BOF) Then
			rs5.Close
			Set rs5 = Nothing
			strFlag = "ERROR_TRACK"
		Else
			rs5.Close
			Set rs5 = Nothing
		End If
	Else
		rs5.Close
		Set rs5 = Nothing
	End IF
	' Item Group Check
	IF Request("txtItemGroupCd") <> "" Then
	 	If rs6.EOF AND rs6.BOF Then
			rs6.Close
			Set rs6 = Nothing
			strFlag = "ERROR_GROUP"
			Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtItemGroupNm.value = """"" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		Else
			Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtItemGroupNm.value = """ & ConvSPChars(rs6("item_group_nm")) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf
			rs6.Close
			Set rs6 = Nothing
		End If
	Else
		rs6.Close
		Set rs6 = Nothing
		Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtItemGroupNm.value = """"" & vbCrLf
		Response.Write "</Script>" & vbCrLf
	End If
	' Error Hnadling	
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
		ElseIf strFlag = "ERROR_WCCD" Then
			Call DisplayMsgBox("182100", vbOKOnly, "", "", I_MKSCRIPT)
			Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtWCCd.focus" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		ElseIf strFlag = "ERROR_SLCD" Then
			Call DisplayMsgBox("125700", vbOKOnly, "", "", I_MKSCRIPT)
			Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtSlCd.focus" & vbCrLf
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
	'	Main Query - Production Results Display
	'=======================================================================================================
	Redim UNISqlId(0)
	Redim UNIValue(0, 15)

	UNISqlId(0) = "P4511MB1"
	
	IF Request("txtItemCd") = "" Then
		strItemCd = "|"
	Else
		StrItemCd = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	End IF

	IF Request("txtProdOrderNo") = "" Then
		strProdOrderNo = "|"
	Else
		StrProdOrderNo = FilterVar(UCase(Request("txtProdOrderNo")), "''", "S")
	End IF

	IF Request("txtWcCd") = "" Then
		strWcCd = "|"
	Else
		StrWcCd = FilterVar(UCase(Request("txtWcCd")), "''", "S")
	End IF

	IF Request("txtTrackingNo") = "" Then
		strTrackingNo = "|"
	Else
		StrTrackingNo = FilterVar(UCase(Request("txtTrackingNo")), "''", "S")
	End IF

	IF Request("txtSlCd") = "" Then
		strSlCd = "|"
	Else
		strSlCd = FilterVar(UCase(Request("txtSlCd")), "''", "S")
	End IF

	IF Request("txtItemGroupCd") = "" Then
		strItemGroupCd = "|"
	Else
		strItemGroupCd = "f.item_group_cd in (select item_group_cd from ufn_P_ListItemGrp(" & FilterVar(Trim(Request("txtItemGroupCd"))	, "''", "S") & " ))"
	End IF

	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 2) = "" & FilterVar("N", "''", "S") & " "
	UNIValue(0, 3) = "" & FilterVar("N", "''", "S") & " "
	UNIValue(0, 4) = StrProdOrderNo
	UNIValue(0, 5) = strItemCd 
	UNIValue(0, 6) = strTrackingNo
	UNIValue(0, 7) = strSlCd
	UNIValue(0, 8) = FilterVar(UNIConvDate(Request("txtFromDt")), "''", "S")
	UNIValue(0, 9) = FilterVar(UNIConvDate(Request("txtToDt")), "''", "S")
	UNIValue(0,10) = strWcCd
	UNIValue(0,11) = "" & FilterVar("F", "''", "S") & " "
	UNIValue(0,12) = "" & FilterVar("I", "''", "S") & " "
	UNIValue(0,13) = "" & FilterVar("G", "''", "S") & " "
	
	Select Case strQryMode
	
		Case CStr(OPMD_CMODE)
			UNIValue(0,14) = "|"
		Case CStr(OPMD_UMODE) 
			 strTemp = ""
			 strTemp = "(a.prodt_order_no > " & lgStrPrevKey 
			 strTemp = strTemp  & " or (a.prodt_order_no = " & lgStrPrevKey		  'second condition  for group view
			 strTemp = strTemp  & " and b.seq >= " & lgStrPrevKey2 & " )) "       'second condition  for group view
			
			UNIValue(0,14) = strTemp
	End Select
	UNIValue(0,15) = strItemGroupCd
			
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing

	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		Response.Write "<Script Language=vbscript>" & vbCrLf
		Response.Write "parent.DbQueryNotOk()" & vbCrLf
		Response.Write "</Script>" & vbCrLf
		Response.End													'☜: 비지니스 로직 처리를 종료함 
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
				strData = strData & Chr(11) & ""																		'Check Box
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PRODT_ORDER_NO"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEM_CD"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEM_NM"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("SPEC"))%>"
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("REPORT_DT"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("SHIFT_CD"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("REPORT_TYPE"))%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("PROD_QTY"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PRODT_ORDER_UNIT"))%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("RCPT_QTY"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("BASE_UNIT"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("SL_CD"))%>"
				strData = strData & Chr(11) & ""																		'창고팝업 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("LOT_FLG"))%>"										'Lot관리여부 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("LOT_GEN_MTHD"))%>"									'Lot 부여방법 
				strData = strData & Chr(11) & ""																		'Lot번호 
				strData = strData & Chr(11) & "0"																		'Lot Sub 번호 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("OPR_NO"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("WC_CD"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("SEQ"))%>"
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("PLAN_START_DT"))%>"
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("PLAN_COMPT_DT"))%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("PRODT_ORDER_QTY"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("PROD_QTY_IN_ORDER_UNIT"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("GOOD_QTY_IN_ORDER_UNIT"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("RCPT_QTY_IN_ORDER_UNIT"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("ORDER_QTY_IN_BASE_UNIT"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("PROD_QTY_IN_BASE_UNIT"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("GOOD_QTY_IN_BASE_UNIT"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("RCPT_QTY_IN_BASE_UNIT"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("SCHD_START_DT"))%>"
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("SCHD_COMPT_DT"))%>"
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("RELEASE_DT"))%>"
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("REAL_START_DT"))%>"
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("REAL_COMPT_DT"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ORDER_STATUS"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("TRACKING_NO"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEM_GROUP_CD"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEM_GROUP_NM"))%>"
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
		
		.lgStrPrevKey = "<%=ConvSPChars(rs0("PRODT_ORDER_NO"))%>"
		.lgStrPrevKey2 = "<%=ConvSPChars(rs0("SEQ"))%>"
<%	
	End If

	rs0.Close
	Set rs0 = Nothing

%>	
	.frm1.hPlantCd.value		= "<%=ConvSPChars(Request("txtPlantCd"))%>"
	.frm1.hItemCd.value			= "<%=ConvSPChars(Request("txtItemCd"))%>"
	.frm1.hProdOrderNo.value	= "<%=ConvSPChars(Request("txtProdOrderNo"))%>"
	.frm1.hWcCd.value			= "<%=ConvSPChars(Request("txtWcCd"))%>"
	.frm1.hTrackingNo.value		= "<%=ConvSPChars(Request("txtTrackingNo"))%>"
	.frm1.hSlCd.value			= "<%=ConvSPChars(Request("txtSlCd"))%>"
	.frm1.hFromDt.value			= "<%=Request("txtFromDt")%>"
	.frm1.hToDt.value			= "<%=Request("txtToDt")%>"
	.frm1.hItemGroupCd.value	= "<%=ConvSPChars(Request("txtItemGroupCd"))%>"
	.DbQueryOk

End With

</Script>	
