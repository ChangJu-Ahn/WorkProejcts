<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        :
'*  3. Program ID           : p4911mb1.asp
'*  4. Program Name         : 표준시간등록 
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
Dim lgStrPrevKey
Dim lgStrPrevKey2
Dim strFlag
Dim strItemCd
Dim StrProdOrderNo
Dim StrWcCd
Dim StrTrackingNo
Dim StrSlCd
Dim strDeleteFlag
Dim strItemGroupCd
Dim strStatus
Dim strTemp
Dim i

	Const C_SHEETMAXROWS_D = 100

	Call HideStatusWnd

	strQryMode = Request("lgIntFlgMode")

'	lgStrPrevKey = FilterVar(Ucase(Trim(Request("lgStrPrevKey"))),"''","S")
'	lgStrPrevKey2 = FilterVar(Ucase(Trim(Request("lgStrPrevKey2"))),"''","S")

	'=======================================================================================================
	'	Handle Description and Check Existence
	'=======================================================================================================
	Redim UNISqlId(3)
	Redim UNIValue(2, 1)

	UNISqlId(0) = "180000saa"
	UNISqlId(1) = "180000sac"
	UNISqlId(2) = "180000saf"

	UNIValue(0, 0) = FilterVar(Ucase(Trim(Request("txtPlantCd"))),"''","S")
	UNIValue(1, 0) = FilterVar(Ucase(Trim(Request("txtWcCd"))),"''","S")
	UNIValue(2, 0) = FilterVar(Ucase(Trim(Request("txtItemCd"))),"''","S")
	UNIValue(2, 1) = FilterVar(Ucase(Trim(Request("txtPlantCd"))),"''","S")
	
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
		Response.Write "parent.frm1.txtWcNm.value = """"" & vbCrLf
		Response.Write "</Script>" & vbCrLf
	End If

	' 품목명 Display
	IF Request("txtItemCd") <> "" Then
		If (rs3.EOF And rs3.BOF) Then
			rs3.Close
			Set rs3 = Nothing
			strFlag = "ERROR_ITEM"
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtItemNm.value = """"" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		Else
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtItemNm.value = """ & ConvSPChars(rs3("ITEM_NM")) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf
			rs3.Close
			Set rs3 = Nothing
		End If
	Else
		rs3.Close
		Set rs3 = Nothing
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtItemNm.value = """"" & vbCrLf
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
			Response.Write "parent.frm1.txtWCCd.focus" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		ElseIf strFlag = "ERROR_ITEM" Then
			Call DisplayMsgBox("122600", vbOKOnly, "", "", I_MKSCRIPT)
			Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtItemCd.focus" & vbCrLf
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
	Redim UNIValue(0, 4)

	UNISqlId(0) = "P4911MA1"

	IF Request("txtWcCd") = "" Then
		strWcCd = "|"
	Else
		StrWcCd = FilterVar(Ucase(Trim(Request("txtWcCd"))),"''","S")
	End IF

	IF Request("txtItemCd") = "" Then
	   strItemCd = "|"
	ELSE
	   strItemCd = FilterVar(Ucase(Trim(Request("txtItemCd"))),"''","S")
	END IF

	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(Ucase(Trim(Request("txtPlantCd"))),"''","S")
	UNIValue(0, 2) = FilterVar(UNIConvDate(Request("txtBaseDt")),"''","S")
	UNIValue(0, 3) = strWcCd
	UNIValue(0, 4) = strItemCd

	UNILock = DISCONNREAD :	UNIFlag = "1"

    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
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
'				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PLANT_CD"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("WC_CD"))%>"
				strData = strData & Chr(11) & ""
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("WC_NM"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEM_CD"))%>"
				strData = strData & Chr(11) & ""
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEM_NM"))%>"
				strData = strData & Chr(11) & "<%=ConvToTimeFormat(rs0("ST_APPLY"))%>"
				strData = strData & Chr(11) & "<%=ConvToTimeFormat(rs0("ST_FIXED"))%>"
				strData = strData & Chr(11) & "<%=ConvToTimeFormat(rs0("AST_FIXED"))%>"
				strData = strData & Chr(11) & "<%=UniDateClientFormat(rs0("FR_DT"))%>"
				strData = strData & Chr(11) & "<%=UniDateClientFormat(rs0("TO_DT"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("USE_YN"))%>"
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
		.frm1.hWcCd.value			= "<%=ConvSPChars(Request("txtWcCd"))%>"
		.frm1.hItemCd.value			= "<%=ConvSPChars(Request("txtItemCd"))%>"
		.frm1.hBaseDt.value			= "<%=Request("txtBaseDt")%>"

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