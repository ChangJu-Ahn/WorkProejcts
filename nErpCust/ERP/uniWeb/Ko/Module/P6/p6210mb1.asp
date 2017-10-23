<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p6210mb1.asp
'*  4. Program Name         : List Cast Result
'*  5. Program Desc         : 
'*  6. Modified date(First) : 2005/04/21
'*  7. Modified date(Last)  : 2005/10/12
'*  8. Modifier (First)     : Chen, Jae Hyun
'*  9. Modifier (Last)      : Chen, Jae Hyun
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
Dim lgStrPrevKey1
Dim lgStrPrevKey2
Dim strFlag
Dim strFromItemCd, strToItemCd
Dim strProdOrderNo
Dim strOprNo
Dim StrWcCd
Dim StrTrackingNo
Dim strItemGroupCd
Dim strCastRslt
Dim strTemp
Dim i

	Const C_SHEETMAXROWS_D = 100

	Call HideStatusWnd

	strQryMode = Request("lgIntFlgMode")

	lgStrPrevKey = FilterVar(Ucase(Trim(Request("lgStrPrevKey"))),"''","S")
	lgStrPrevKey1 = FilterVar(Ucase(Trim(Request("lgStrPrevKey1"))),"''","S")
	lgStrPrevKey2 = FilterVar(Ucase(Trim(Request("lgStrPrevKey2"))),"''","S")

	Err.Clear

	'=======================================================================================================
	'	Handle Description and Check Existence
	'=======================================================================================================
	Redim UNISqlId(5)
	Redim UNIValue(5, 1)

	UNISqlId(0) = "180000saa"
	UNISqlId(1) = "180000saf"
	UNISqlId(2) = "180000sac"
	UNISqlId(3) = "180000saf"
	UNISqlId(4) = "180000sam"
	UNISqlId(5) = "180000sas"
	
	UNIValue(0, 0) = FilterVar(Ucase(Trim(Request("txtPlantCd"))),"''","S")
	UNIValue(1, 0) = FilterVar(Ucase(Trim(Request("txtFromItemCd"))),"''","S")
	UNIValue(1, 1) = FilterVar(Ucase(Trim(Request("txtPlantCd"))),"''","S")
	UNIValue(2, 0) = FilterVar(Ucase(Trim(Request("txtWcCd"))),"''","S")
	UNIValue(3, 0) = FilterVar(Ucase(Trim(Request("txtToItemCd"))),"''","S")
	UNIValue(3, 1) = FilterVar(Ucase(Trim(Request("txtPlantCd"))),"''","S")
	UNIValue(4, 0) = FilterVar(Ucase(Trim(Request("txtTrackingNo"))),"''","S")
	UNIValue(5, 0) = FilterVar(Ucase(Trim(Request("txtItemGroupCd"))), "''", "S")

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
	IF Request("txtFromItemCd") <> "" Then
		If (rs2.EOF And rs2.BOF) Then
			rs2.Close
			Set rs2 = Nothing
			Response.Write "<Script Language=vbscript>" & vbCrLf
			Response.Write "parent.frm1.txtFromItemNm.value = """"" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		Else
			Response.Write "<Script Language=vbscript>" & vbCrLf
			Response.Write "parent.frm1.txtFromItemNm.value = """ & ConvSPChars(rs2("ITEM_NM")) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf
			rs2.Close
			Set rs2 = Nothing
		End If
	Else
		rs2.Close
		Set rs2 = Nothing
		Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtFromItemNm.value = """"" & vbCrLf
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
	' Item Check
	IF Request("txtToItemCd") <> "" Then
		If (rs4.EOF And rs4.BOF) Then
			rs4.Close
			Set rs4 = Nothing
			Response.Write "<Script Language=vbscript>" & vbCrLf
			Response.Write "parent.frm1.txtToItemNm.value = """"" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		Else
			Response.Write "<Script Language=vbscript>" & vbCrLf
			Response.Write "parent.frm1.txtToItemNm.value = """ & ConvSPChars(rs4("ITEM_NM")) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf
			rs4.Close
			Set rs4 = Nothing
		End If
	Else
		rs4.Close
		Set rs4 = Nothing
		Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtToItemNm.value = """"" & vbCrLf
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
		ElseIf strFlag = "ERROR_WCCD" Then
			Call DisplayMsgBox("182100", vbOKOnly, "", "", I_MKSCRIPT)
			Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtWCCd.focus" & vbCrLf
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
	Redim UNIValue(0, 13)
	
	IF Request("cboResultFlg") = "Y" Then
		UNISqlId(0) = "p6210mb1y"
		strCastRslt = "|"
	ElseIf Request("cboResultFlg") = "N" Then
		UNISqlId(0) = "p6210mb1n"
		strCastRslt = " NOT EXISTS (SELECT PRODT_ORDER_NO, OPR_NO, SEQ FROM P_CAST_RESULTS WHERE A.PRODT_ORDER_NO = PRODT_ORDER_NO AND A.OPR_NO = OPR_NO AND A.SEQ = SEQ)"
	Else	
		UNISqlId(0) = "p6210mb1A"
		strCastRslt = "|"
	End If	
	
	IF Request("txtFromItemCd") = "" Then
		strFromItemCd = "|"
	Else
		strFromItemCd = FilterVar(Ucase(Trim(Request("txtItemCd"))),"''","S")
	End IF
	
	IF Request("txtToItemCd") = "" Then
		strToItemCd = "|"
	Else
		strToItemCd = FilterVar(Ucase(Trim(Request("txtItemCd"))),"''","S")
	End IF

	IF Request("txtProdOrderNo") = "" Then
		strProdOrderNo = "|"
	Else
		StrProdOrderNo = FilterVar(Ucase(Trim(Request("txtProdOrderNo"))),"''","S")
	End IF
	
	IF Request("txtOprCD") = "" Then
		strOprNo = "|"
	Else
		strOprNo = FilterVar(Ucase(Trim(Request("txtOprCD"))),"''","S")
	End IF

	IF Request("txtWcCd") = "" Then
		strWcCd = "|"
	Else
		StrWcCd = FilterVar(Ucase(Trim(Request("txtWcCd"))),"''","S")
	End IF

	IF Request("txtTrackingNo") = "" Then
		strTrackingNo = "|"
	Else
		StrTrackingNo = FilterVar(Ucase(Trim(Request("txtTrackingNo"))),"''","S")
	End IF
	
	IF Request("txtItemGroupCd") = "" Then
		strItemGroupCd = "|"
	Else
		strItemGroupCd = "f.item_group_cd in (select item_group_cd from ufn_P_ListItemGrp(" & FilterVar(Trim(Request("txtItemGroupCd"))	,"''", "S") & " ))"
	End IF

 
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(Ucase(Trim(Request("txtPlantCd"))),"''","S")
	UNIValue(0, 2) = FilterVar(UNIConvDate(Request("txtFromDt")),"''","S")
	UNIValue(0, 3) = FilterVar(UNIConvDate(Request("txtToDt")),"''","S")
	UNIValue(0, 4) = strFromItemCd
	UNIValue(0, 5) = strToItemCd
	UNIValue(0, 6) = strProdOrderNo
	UNIValue(0, 7) = strOprNo
	UNIValue(0, 8) = strWcCd
	UNIValue(0, 9) = strTrackingNo
	UNIValue(0,10) = "|"
	UNIValue(0,11) = strItemGroupCd
	UNIValue(0,12) = strCastRslt
	
	Select Case strQryMode
	
		Case CStr(OPMD_CMODE)
			UNIValue(0,13) = "|"
		Case CStr(OPMD_UMODE) 
			strTemp = ""
			strTemp = "(a.prodt_order_no > " & lgStrPrevKey 
			strTemp = strTemp  & " or (a.prodt_order_no = " & lgStrPrevKey		  
			strTemp = strTemp  & " and a.opr_no > " & lgStrPrevKey1  & " ) "    
			strTemp = strTemp  & " or (a.prodt_order_no = " & lgStrPrevKey		  
			strTemp = strTemp  & " and a.opr_no = " & lgStrPrevKey1 
			strTemp = strTemp  & " and a.seq >= " & lgStrPrevKey2 & " )) "    
			
			UNIValue(0,13) = strTemp
	End Select
			
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
				If Trim("<%=rs0("RESULT_FLAG")%>") = "Y" Then
					strData = strData & Chr(11) & "1"
				Else
					strData = strData & Chr(11) & "0"
				End If	
				strData = strData & Chr(11) & "<%=rs0("RESULT_FLAG")%>"	
				strData = strData & Chr(11) & "<%=rs0("DEL_FLG")%>"	
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PRODT_ORDER_NO"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("OPR_NO"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("SEQ"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("SHIFT_CD"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEM_CD"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEM_NM"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("SPEC"))%>"
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("REPORT_DT"))%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("PROD_QTY_IN_ORDER_UNIT"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PRODT_ORDER_UNIT"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("FACILITY_CD"))%>"
				strData = strData & Chr(11) & ""
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("FACILITY_NM"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("CAST_CD"))%>"
				strData = strData & Chr(11) & ""
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("CAST_NM"))%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("CUR_COUNT"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("CAVI"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("INPUT_QTY"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("REMARK"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("WC_CD"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("WC_NM"))%>"
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
		.lgStrPrevKey1 = "<%=ConvSPChars(rs0("OPR_NO"))%>"
		.lgStrPrevKey2 = "<%=ConvSPChars(rs0("SEQ"))%>"
<%	
	End If

	rs0.Close
	Set rs0 = Nothing

%>	
	.frm1.hPlantCd.value		= "<%=ConvSPChars(Request("txtPlantCd"))%>"
	.frm1.hFromItemCd.value		= "<%=ConvSPChars(Request("txtFromItemCd"))%>"
	.frm1.hToItemCd.value		= "<%=ConvSPChars(Request("txtToItemCd"))%>"
	.frm1.hProdOrderNo.value	= "<%=ConvSPChars(Request("txtProdOrderNo"))%>"
	.frm1.hOprCD.value			= "<%=ConvSPChars(Request("txtOprCD"))%>"
	.frm1.hWcCd.value			= "<%=ConvSPChars(Request("txtWcCd"))%>"
	.frm1.hTrackingNo.value		= "<%=ConvSPChars(Request("txtTrackingNo"))%>"
	.frm1.hFromDt.value			= "<%=Request("txtFromDt")%>"
	.frm1.hToDt.value			= "<%=Request("txtToDt")%>"
	.frm1.hItemGroupCd.value	= "<%=ConvSPChars(Request("txtItemGroupCd"))%>"
	.frm1.hResultFlg.value		= "<%=ConvSPChars(Request("cboCastRsltFlg"))%>"
	.DbQueryOk(LngMaxRow + 1)

End With

</Script>	
