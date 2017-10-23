<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        :
'*  3. Program ID           : p4915mb1.asp
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) : 2005-01-17
'*  7. Modified date(Last)  :
'*  8. Modifier (First)     : Yoon, Jeong Woo
'*  9. Modifier (Last)      :
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
Call loadInfTB19029B("Q", "P", "NOCOOKIE", "MB")

On Error Resume Next

Dim ADF
Dim strRetMsg
Dim UNISqlId, UNIValue, UNILock, UNIFlag
Dim rs0, rs1, rs2, rs3, rs4, rs5, rs6
Dim strQryMode
'Dim lgStrPrevKey
'Dim lgStrPrevKey2
Dim strFlag
Dim strItemCd
Dim StrSlCd
Dim strDeleteFlag
Dim strItemGroupCd
Dim strStatus
Dim strTemp
Dim i


Dim strFromDt
Dim strToDt
Dim StrProdOrderNo
Dim strWcCd
Dim strTrackingNo
Dim strShiftCd

	Const C_SHEETMAXROWS_D = 100

	Call HideStatusWnd

	strQryMode = Request("lgIntFlgMode")

'	lgStrPrevKey = FilterVar(Ucase(Trim(Request("lgStrPrevKey"))),"''","S")
'	lgStrPrevKey2 = FilterVar(Ucase(Trim(Request("lgStrPrevKey2"))),"''","S")

	'=======================================================================================================
	'	Handle Description and Check Existence
	'=======================================================================================================
	Redim UNISqlId(2)
	Redim UNIValue(1, 0)

	UNISqlId(0) = "180000saa"
	UNISqlId(1) = "180000sac"
	UNISqlId(2) = "180000saf"

	UNIValue(0, 0) = FilterVar(Ucase(Trim(Request("txtPlantCd"))),"''","S")
	UNIValue(1, 0) = FilterVar(Ucase(Trim(Request("txtWcCd"))),"''","S")

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
	' Work Center Check
	IF Request("txtWcCd") <> "" Then
		If (rs2.EOF And rs2.BOF) Then
			rs2.Close
			Set rs2 = Nothing
			strFlag = "ERROR_WCCD"
			Response.Write "<Script Language=vbscript>" & vbCrLf
			Response.Write "parent.frm1.txtWcNm.value = """"" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		Else
			Response.Write "<Script Language=vbscript>" & vbCrLf
			Response.Write "parent.frm1.txtWcNm.value = """ & ConvSPChars(rs2("WC_NM")) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf
			rs2.Close
			Set rs2 = Nothing
		End If
	Else
		rs2.Close
		Set rs2 = Nothing
		Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtWCNm.value = """"" & vbCrLf
		Response.Write "</Script>" & vbCrLf
	End IF
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
			Response.Write "parent.frm1.txtWcCd.focus" & vbCrLf
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
	Redim UNIValue(0, 7)

	UNISqlId(0) = "P4915MA1"

	If Request("txtFromDt") = "" Then
		strFromDt = "|"
	Else
		strFromDt = FilterVar(UNIConvDate(Request("txtFromDt")),"''","S")
	End If

	If Request("txtToDt") = "" Then
		strToDt = "|"
	Else
		strToDt = FilterVar(UNIConvDate(Request("txtToDt")),"''","S")
	End If

	IF Request("txtProdOrderNo") = "" Then
		strProdOrderNo = "|"
	Else
		StrProdOrderNo = FilterVar(Ucase(Trim(Request("txtProdOrderNo"))),"''","S")
	End IF

	IF Request("txtWcCd") = "" Then
		strWcCd = "|"
	Else
		StrWcCd = FilterVar(Ucase(Trim(Request("txtWcCd"))),"''","S")
	End IF

	If Request("txtTrackingNo") = "" Then
		strTrackingNo = "|"
	Else
		strTrackingNo = FilterVar(Ucase(Trim(Request("txtTrackingNo"))),"''","S")
	End If

	If Request("txtShiftCd") = "" Then
		strShiftCd = "|"
	Else
		strShiftCd = FilterVar(Ucase(Trim(Request("txtShiftCd"))),"''","S")
	End If

	UNIValue(0, 0) = "^"
'	UNIValue(0, 1) = "'MR'"
	UNIValue(0, 1) = FilterVar(Ucase(Trim(Request("txtPlantCd"))),"''","S")
	UNIValue(0, 2) = strFromDt
	UNIValue(0, 3) = strToDt
	UNIValue(0, 4) = strProdOrderNo
	UNIValue(0, 5) = strWcCd
	UNIValue(0, 6) = strTrackingNo
	UNIValue(0, 7) = strShiftCd

'	Response.Write "aa" & UNIValue(0, 0) & "<P>"
'	Response.Write "bb" & UNIValue(0, 1) & "<P>"
'	Response.Write "cc" & UNIValue(0, 2) & "<P>"
'	Response.Write "dd" & UNIValue(0, 3) & "<P>"
'	Response.Write "ee" & UNIValue(0, 4) & "<P>"
'	Response.Write "ff" & UNIValue(0, 5) & "<P>"
'	Response.Write "gg" & UNIValue(0, 6) & "<P>"
'	Response.Write "hh" & UNIValue(0, 7) & "<P>"

'	Select Case strQryMode
'
'		Case CStr(OPMD_CMODE)
'			UNIValue(0, 12) = "|"
'		Case CStr(OPMD_UMODE)
'			 strTemp = ""
'			 strTemp = "(a.prodt_order_no > " & lgStrPrevKey
'			 strTemp = strTemp  & " or (a.prodt_order_no = " & lgStrPrevKey
'			 strTemp = strTemp  & " and c.seq >= " & lgStrPrevKey2 & " )) "
'
'			UNIValue(0, 12) = strTemp
'	End Select
'
'	UNIValue(0,13) = strItemGroupCd

	UNILock = DISCONNREAD :	UNIFlag = "1"

    Set ADF = Server.CreateObject("prjPublic.cCtlTake")

    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

	Response.Write strRetMsg & "<P>"
    Set ADF = Nothing

	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		Response.End
	End If

%>

<Script Language=vbscript>
Dim LngMaxRow
Dim strData
Dim TmpBuffer
Dim iTotalStr

With parent

	LngMaxRow = .frm1.vspdData.MaxRows

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
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("WC_CD"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("WC_NM"))%>"
				strData = strData & Chr(11) & "<%=UniDateClientFormat(rs0("REPORT_DT"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PRODT_ORDER_NO"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("OPR_NO"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("RESOURCE_CD"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("RESOURCE_NM"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEM_CD"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEM_NM"))%>"
				strData = strData & Chr(11) & "<%=ConvToTimeFormat(rs0("ST_TIME"))%>"
				strData = strData & Chr(11) & "<%=ConvToTimeFormat(rs0("END_TIME"))%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("LOSS_MAN"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=ConvToTimeFormat(rs0("WK_LOSS_QTY"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("WK_LOSS_CD"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("WK_LOSS_NM"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("RT_DEPT_CD"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("RT_DEPT_NM"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("NOTES"))%>"
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

'		.lgStrPrevKey = "<%=ConvSPChars(rs0("prodt_order_no"))%>"
'		.lgStrPrevKey2 = "<%=ConvSPChars(rs0("seq"))%>"

		.frm1.hPlantCd.value		= "<%=ConvSPChars(Request("txtPlantCd"))%>"
		.frm1.hFromDt.value			= "<%=Request("txtFromDt")%>"
		.frm1.hToDt.value			= "<%=Request("txtToDt")%>"
		.frm1.hProdOrderNo.value	= "<%=ConvSPChars(Request("txtProdOrderNo"))%>"
		.frm1.hWcCd.value			= "<%=ConvSPChars(Request("txtWcCd"))%>"
		.frm1.hTrackingNo.value		= "<%=ConvSPChars(Request("txtTrackingNo"))%>"
		.frm1.hShiftCd.value		= "<%=ConvSPChars(Request("txtShiftCd"))%>"

<%
	End If

	rs0.Close
	Set rs0 = Nothing

%>
	.DbQueryOk

End With

</Script>
<script Language = vbscript RUNAT = server>
Function ConvToTimeFormat(ByVal iVal)
	Dim iTime
	Dim iMin
	Dim iSec
			
	If IVal = 0 Then
		ConvToTimeFormat = "00:00:00"
	Else
		iMin = Fix(IVal Mod 3600)
		iTime = Fix(IVal /3600)
		
		iSec = Fix(iMin Mod 60)
		iMin = Fix(iMin / 60)
		
		ConvToTimeFormat = Right("0" & CStr(iTime),2) & ":" & Right("0" & CStr(iMin),2) & ":" & Right("0" & CStr(iSec),2)
		 
	End If
End Function
</script>