<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : 설계BOM관리 
'*  2. Function Name        :
'*  3. Program ID           : p1714mb1.asp
'*  4. Program Name         : 제조BOM이관 
'*  5. Program Desc         :
'*  6. Modified date(First) : 2005-02-14
'*  7. Modified date(Last)  :
'*  8. Modifier (First)     : Yoon, Jeong Woo
'*  9. Modifier (Last)      :
'* 10. Comment              : 이관요청된 BOM을 제조BOM으로 이관처리 
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

Dim strBasePlantCd
Dim strItemCd
Dim strReqTransNo

Dim i

	Const C_SHEETMAXROWS_D = 100

	Call HideStatusWnd

	strQryMode = Request("lgIntFlgMode")

'	lgStrPrevKey = FilterVar(Ucase(Trim(Request("lgStrPrevKey"))),"''","S")
'	lgStrPrevKey2 = FilterVar(Ucase(Trim(Request("lgStrPrevKey2"))),"''","S")

	'=======================================================================================================
	'	Handle Description and Check Existence
	'=======================================================================================================
	Redim UNISqlId(2)
	Redim UNIValue(2, 1)

	UNISqlId(0) = "P1714MA2"	' Design Plant : 설계공장 
	UNISqlId(1) = "180000saa"
	UNISqlId(2) = "180000saf"

	UNIValue(0, 0) = FilterVar(Ucase(Trim(Request("txtBasePlantCd"))),"''","S")		' Design Plant : 설계공장 
	UNIValue(1, 0) = FilterVar(Ucase(Trim(Request("txtDestPlantCd"))),"''","S")		' 대상공장 
	UNIValue(2, 0) = FilterVar(Ucase(Trim(Request("txtItemCd"))),"''","S")
	UNIValue(2, 1) = FilterVar(Ucase(Trim(Request("txtDestPlantCd"))),"''","S")

	UNILock = DISCONNREAD :	UNIFlag = "1"

    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1, rs2, rs3, rs4, rs5, rs6)
	Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing

	' txtBasePlantCd Check
	If (rs1.EOF And rs1.BOF) Then
		rs1.Close
		Set rs1 = Nothing
		strFlag = "ERROR_BasePLANT"
		Response.Write "<Script Language=vbscript>" & vbCrLf
		Response.Write "parent.frm1.txtBasePlantNm.value = """"" & vbCrLf
		Response.Write "</Script>" & vbCrLf
	Else
		Response.Write "<Script Language=vbscript>" & vbCrLf
		Response.Write "parent.frm1.txtBasePlantNm.value = """ & ConvSPChars(rs1("PLANT_NM")) & """" & vbCrLf
		Response.Write "</Script>" & vbCrLf
		rs1.Close
		Set rs1 = Nothing
	End If

	' txtDestPlantCd Check
	If (rs2.EOF And rs2.BOF) Then
		rs2.Close
		Set rs2 = Nothing
		strFlag = "ERROR_DestPLANT"
		Response.Write "<Script Language=vbscript>" & vbCrLf
		Response.Write "parent.frm1.txtDestPlantNm.value = """"" & vbCrLf
		Response.Write "</Script>" & vbCrLf
	Else
		Response.Write "<Script Language=vbscript>" & vbCrLf
		Response.Write "parent.frm1.txtDestPlantNm.value = """ & ConvSPChars(rs2("PLANT_NM")) & """" & vbCrLf
		Response.Write "</Script>" & vbCrLf
		rs2.Close
		Set rs2 = Nothing
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
		If strFlag = "ERROR_BasePLANT" Then
			Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
			Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtBasePlantCd.focus" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		ElseIf strFlag = "ERROR_DestPLANT" Then
			Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
			Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtDestPlantCd.focus" & vbCrLf
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

	UNISqlId(0) = "p1714ma1"

'		strVal = strVal & "&txtBasePlantCd="	& iStrBasePlantCd				'☆: 조회 조건 데이타 
'		strVal = strVal & "&txtDestPlantCd="	& iStrDestPlantCd				'☆: 조회 조건 데이타 
'		strVal = strVal & "&txtItemCd="			& iStrItemCd					'☆: 조회 조건 데이타 
'		strVal = strVal & "&txtReqTransNo="		& iStrReqTransNo 				'☆: 조회 조건 데이타 

	IF Request("txtBasePlantCd") = "" Then
		strBasePlantCd = "|"
	Else
		StrBasePlantCd = FilterVar(Ucase(Trim(Request("txtBasePlantCd"))),"''","S")
	End IF

	IF Request("txtItemCd") = "" Then
	   strItemCd = "|"
	ELSE
	   strItemCd = FilterVar(Ucase(Trim(Request("txtItemCd"))),"''","S")
	END IF

	IF Request("txtReqTransNo") = "" Then
	   strReqTransNo = "|"
	ELSE
	   strReqTransNo = FilterVar(Ucase(Trim(Request("txtReqTransNo"))),"''","S")
	END IF

	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(Ucase(Trim(Request("txtDestPlantCd"))),"''","S")
	UNIValue(0, 2) = StrBasePlantCd
	UNIValue(0, 3) = strItemCd
	UNIValue(0, 4) = strReqTransNo

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
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("CHK"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("REQ_TRANS_NO"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PLANT_CD"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PLANT_NM"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("DESIGN_PLANT_CD"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("DESIGN_PLANT_NM"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEM_CD"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEM_NM"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("SPEC"))%>"
				strData = strData & Chr(11) & "<%=UniDateClientFormat(rs0("REQ_TRANS_DT"))%>"
				strData = strData & Chr(11) & "<%=UniDateClientFormat(rs0("TRANS_DT"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("DESCRIPTION"))%>"
				strData = strData & Chr(11) & "<%=UniDateClientFormat(rs0("VALID_FROM_DT"))%>"
				strData = strData & Chr(11) & "<%=UniDateClientFormat(rs0("VALID_TO_DT"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("DRAWING_PATH"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("STATUS"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("BOM_NO"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("MAJOR_FLG"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("RETURNDESC"))%>"
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

'		.frm1.hPlantCd.value		= "<%=ConvSPChars(Request("txtPlantCd"))%>"
'		.frm1.hWcCd.value			= "<%=ConvSPChars(Request("txtWcCd"))%>"
'		.frm1.hItemCd.value			= "<%=ConvSPChars(Request("txtItemCd"))%>"
'		.frm1.hBaseDt.value			= "<%=Request("txtBaseDt")%>"

<%
	End If

	rs0.Close
	Set rs0 = Nothing

%>
	.DbQueryOk

End With

</Script>