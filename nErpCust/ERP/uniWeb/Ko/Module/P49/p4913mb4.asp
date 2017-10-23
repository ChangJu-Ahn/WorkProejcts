<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        :
'*  3. Program ID           : p4913mb4.asp
'*  4. Program Name         : 간접공수 현황 
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
<%
Call LoadBasisGlobalInf
Call loadInfTB19029B("I", "*", "NOCOOKIE","MB")
On Error Resume Next

Dim ADF
Dim strRetMsg
Dim UNISqlId, UNIValue, UNILock, UNIFlag
Dim rs0, rs1, rs2, rs3, rs4
Dim strQryMode
Dim lgStrPrevKey
Dim strProdOrdNo
Dim strFlag
Dim strWcCd
Dim i

	Const C_SHEETMAXROWS_D = 100

	Call HideStatusWnd

	strQryMode = Request("lgIntFlgMode")

'	lgStrPrevKey = FilterVar(Ucase(Trim(Request("lgStrPrevKey1"))),"''","S")

	'======================================================================================================
	'	Handle Description
	'======================================================================================================
	Redim UNISqlId(1)
	Redim UNIValue(1, 0)

	UNISqlId(0) = "180000saa"
	UNISqlId(1) = "180000sac"
'	UNISqlId(2) = "180000sam"
'	UNISqlId(3) = "180000sas"

	UNIValue(0, 0) = FilterVar(Ucase(Trim(Request("txtPlantCd"))),"''","S")
	UNIValue(1, 0) = FilterVar(Ucase(Trim(Request("txtWcCd"))),"''","S")

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
		Response.Write "parent.frm1.txtWcNm.value = """"" & vbCrLf
		Response.Write "</Script>" & vbCrLf
	End If

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
	'	Main Query - Order Header Display
	'=======================================================================================================
	Redim UNISqlId(1)
	Redim UNIValue(1, 3)

	UNISqlId(0) = "P4913MA5"
	UNISqlId(1) = "P4913MA6"

'	UNIValue(0, 0) = "^"
	UNIValue(0, 0) = FilterVar(Ucase(Trim(Request("txtPlantCd"))),"''","S")
	UNIValue(0, 1) = FilterVar(Ucase(Trim(Request("txtWcCd"))),"''","S")
	UNIValue(0, 2) = FilterVar(UNIConvDate(Request("txtprodDt")),"''","S")
'	UNIValue(0, 3) = FilterVar(Ucase(Trim(Request("txtProdOrderNo"))),"''","S")

	UNIValue(1, 0) = FilterVar(Ucase(Trim(Request("txtPlantCd"))),"''","S")
	UNIValue(1, 1) = FilterVar(Ucase(Trim(Request("txtWcCd"))),"''","S")
	UNIValue(1, 2) = FilterVar(UNIConvDate(Request("txtprodDt")),"''","S")
'	UNIValue(1, 3) = FilterVar(Ucase(Trim(Request("txtProdOrderNo"))),"''","S")

	UNILock = DISCONNREAD :	UNIFlag = "1"

    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)

	If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)	'⊙: DB 에러코드, 메세지타입, %처리, 스크립트유형 
		rs0.Close
		Set rs0 = Nothing
		Response.End													'☜: 비지니스 로직 처리를 종료함 
	End If

%>

<Script Language=vbscript>
Dim LngMaxRow
Dim strData
Dim TmpBuffer
Dim iTotalStr

With parent																	'☜: 화면 처리 ASP 를 지칭함 
	LngMaxRow = .frm1.vspdData4.MaxRows										'Save previous Maxrow
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
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PLANT_CD"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("WC_CD"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("REPORT_DT"))%>"
'			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PRODT_ORDER_NO"))%>"
			strData = strData & Chr(11) & "<%=rs0("SEQ_NO")%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("EMP_NO"))%>"
			strData = strData & Chr(11) & ""
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("EMP_NM"))%>"
			strData = strData & Chr(11) & "<%=ConvToTimeFormat(rs0("TIME"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("NOTES"))%>"

'			msgbox "<%=rs0("STD_TIME")%>"

			strData = strData & Chr(11) & LngMaxRow + <%=i%>
			strData = strData & Chr(11) & Chr(12)

			TmpBuffer(<%=i%>) = strData
<%
			rs0.MoveNext
		End If
	Next
%>
		iTotalStr = Join(TmpBuffer, "")
		.ggoSpread.Source = .frm1.vspdData4
		.ggoSpread.SSShowDataByClip iTotalStr

'		.lgStrPrevKey1 = "<%=Trim(rs0("PRODT_ORDER_NO"))%>"

		.frm1.hPlantCd.value	 = "<%=ConvSPChars(Request("txtPlantCd"))%>"
		.frm1.hProdDt.value      = "<%=Request("txtprodDt")%>"
		.frm1.hWcCd.value        = "<%=ConvSPChars(Request("txtWcCd"))%>"
'		.frm1.hProdOrderNo.value = "<%=ConvSPChars(Request("txtProdOrderNo"))%>"

<%
		rs0.Close
		Set rs0 = Nothing
%>
'	.DbQueryOk(LngMaxRow)
End With


With parent																	'☜: 화면 처리 ASP 를 지칭함 
	LngMaxRow = .frm1.vspdData5.MaxRows										'Save previous Maxrow
<%
	If C_SHEETMAXROWS_D < rs1.RecordCount Then
%>
		ReDim TmpBuffer(<%=C_SHEETMAXROWS_D - 1%>)
<%
	Else
%>
		ReDim TmpBuffer(<%=rs1.RecordCount - 1%>)
<%
	End If

    For i=0 to rs1.RecordCount-1
		If i < C_SHEETMAXROWS_D Then
%>
			strData = ""
			strData = strData & Chr(11) & "<%=ConvSPChars(rs1("PLANT_CD"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs1("WC_CD"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs1("REPORT_DT"))%>"
'			strData = strData & Chr(11) & "<%=ConvSPChars(rs1("PRODT_ORDER_NO"))%>"
			strData = strData & Chr(11) & "<%=rs1("SEQ_NO")%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs1("EMP_NO"))%>"
			strData = strData & Chr(11) & ""
			strData = strData & Chr(11) & "<%=ConvSPChars(rs1("EMP_NM"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs1("ACT_CD"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs1("ACT_NM"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs1("NOTES"))%>"

'			msgbox "<%=rs1("STD_TIME")%>"

			strData = strData & Chr(11) & LngMaxRow + <%=i%>
			strData = strData & Chr(11) & Chr(12)

			TmpBuffer(<%=i%>) = strData
<%
			rs1.MoveNext
		End If
	Next
%>
		iTotalStr = Join(TmpBuffer, "")
		.ggoSpread.Source = .frm1.vspdData5
		.ggoSpread.SSShowDataByClip iTotalStr

'		.lgStrPrevKey1 = "<%=Trim(rs1("PRODT_ORDER_NO"))%>"

		.frm1.hPlantCd.value	 = "<%=ConvSPChars(Request("txtPlantCd"))%>"
		.frm1.hProdDt.value      = "<%=Request("txtprodDt")%>"
		.frm1.hWcCd.value        = "<%=ConvSPChars(Request("txtWcCd"))%>"
'		.frm1.hProdOrderNo.value = "<%=ConvSPChars(Request("txtProdOrderNo"))%>"

<%
		rs1.Close
		Set rs1 = Nothing
%>
	.DbQueryOk(LngMaxRow)
End With
</Script>
<%
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
<Script Language=vbscript RUNAT=server>
'==============================================================================
' Function : ConvToTimeFormat
' Description : 시간 형식으로 변경 
'==============================================================================
Function ConvToTimeFormat(ByVal iVal)
	Dim iTime
	Dim iMin
	Dim iSec

	Dim iVal2

	iVal2 = Fix(iVal)

	If iVal2 = 0 Then
		ConvToTimeFormat = "00:00:00"
	ElseIf iVal2 > 0 Then
		iMin = Fix(iVal2 Mod 3600)
		iTime = Fix(iVal2 /3600)

		iSec = Fix(iMin Mod 60)
		iMin = Fix(iMin / 60)

		ConvToTimeFormat = Right("0" & CStr(iTime),2) & ":" & Right("0" & CStr(iMin),2) & ":" & Right("0" & CStr(iSec),2)
	Else
		iVal2 = Replace(iVal2, "-", "")
		iMin = Fix(iVal2 Mod 3600)
		iTime = Fix(iVal2 /3600)

		iSec = Fix(iMin Mod 60)
		iMin = Fix(iMin / 60)
		ConvToTimeFormat = Right("0" & CStr(iTime),2) & ":" & Right("0" & CStr(iMin),2) & ":" & Right("0" & CStr(iSec),2)
		ConvToTimeFormat = "-" & ConvToTimeFormat

	End If
End Function
</script>