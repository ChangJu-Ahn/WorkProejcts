<%@LANGUAGE = VBScript%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4600mb1.asp
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) : 2002/01/02
'*  7. Modified date(Last)  : 2002/02/21
'*  8. Modifier (First)     : Park, BumSoo 
'*  9. Modifier (Last)      : Park, BumSoo 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=====================================================================================================%>

<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "P", "NOCOOKIE", "PB")

Dim ADF														'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg												'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag					'DBAgent Parameter 선언 
Dim	rs0, rs1, rs2, rs3, rs4
Dim strQryMode
Dim strConsumedDtFrom, strConsumedDtTo, strItemCd, strWcCd, strProdtOrderNo, strResourceCd, strResourceGroupCd
Dim strMode											'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

Const C_SHEETMAXROWS = 100

Call HideStatusWnd


'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================
Dim StrNextKey		' 다음 값 
Dim lgStrPrevKey	
Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngRow
Dim GroupCount

Dim strEcnNo
Dim strEcnDesc
Dim strReasonCd
Dim strValidDt
Dim strStatus
Dim strEBomFlg
Dim strMBomFlg

On Error Resume Next
Err.Clear																'☜: Protect system from crashing
	
	strQryMode = Request("lgIntFlgMode")
	
	'--------------------------------------------
	' 변경근거가 존재하는지 체크 
	'--------------------------------------------
	If Trim(Request("txtReasonCd")) <> "" Then
		Redim UNISqlId(0)
		Redim UNIValue(0, 1)

		UNISqlId(0) = "s0000qa000"
	
		UNIValue(0, 0) = FilterVar("P1402","''","S")						'major_cd
		UNIValue(0, 1) = FilterVar(Request("txtReasonCd"),"''","S")			'minor_cd

		UNILock = DISCONNREAD :	UNIFlag = "1"
	
		Set ADF = Server.CreateObject("prjPublic.cCtlTake")
		strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1)

		If (rs1.EOF And rs1.BOF) Then
			Call DisplayMsgBox("182803", vbOKOnly, "", "", I_MKSCRIPT)
			rs1.Close
			Set rs1 = Nothing
			Set ADF = Nothing
			Response.Write "<Script Language=vbscript>		" & vbCr
			Response.Write "	parent.txtReasonCd.focus()	" & vbCr															
			Response.Write "</Script>						" & vbCr
			Response.End
		End If
		rs1.Close
		Set rs1 = Nothing
	End If

		
	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================
	Redim UNISqlId(0)
	Redim UNIValue(0, 9)

	UNISqlId(0) = "p1410mb9a"

	IF Trim(Request("txtEcnNo")) = "" Then
	   strEcnNo = "|"
	ELSE
	   strEcnNo = FilterVar(UCase(Request("txtEcnNo")), "''", "S")
	END IF

	IF Trim(Request("lgStrPrevKey")) = "" Then
		lgStrPrevKey = "|"
	ELSE
		lgStrPrevKey = FilterVar(UCase(Request("lgStrPrevKey")), "''", "S")
	END IF
	
	IF Trim(Request("txtEcnDesc")) = "" Then
	   strEcnDesc = "|"
	ELSE
		strEcnDesc = Replace(Trim(Request("txtEcnDesc")), "[", "[[]")
		strEcnDesc = "%" & Replace(strEcnDesc, "%", "[%]") & "%"
		strEcnDesc = FilterVar(strEcnDesc, "''", "S")
	END IF	
	
	IF Trim(Request("txtReasonCd")) = "" Then
	   strReasonCd = "|"
	ELSE
	   strReasonCd = FilterVar(UCase(Request("txtReasonCd")), "''", "S")
	END IF		
		
	IF Trim(Request("txtValidDt")) = "" Then
	   strValidDt = "|"
	ELSE
	   strValidDt = " " & FilterVar(UNIConvDate(Request("txtValidDt")), "''", "S") & "" 
	END IF

	IF Trim(Request("cboStatus")) = "" Then
	   strStatus = "|"
	ELSE
	   strStatus = " " & FilterVar(Request("cboStatus"), "''", "S") & ""
	END IF	

	IF Trim(Request("cboEBomFlg")) = "" Then
	   strEBomFlg = "|"
	ELSE
	   strEBomFlg = " " & FilterVar(Request("cboEBomFlg"), "''", "S") & ""
	END IF

	IF Trim(Request("cboMBomFlg")) = "" Then
	   strMBomFlg = "|"
	ELSE
	   strMBomFlg = " " & FilterVar(Request("cboMBomFlg"), "''", "S") & ""
	END IF

	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = strEcnNo
	UNIValue(0, 2) = lgStrPrevKey
	UNIValue(0, 3) = strEcnDesc
	UNIValue(0, 4) = strReasonCd
	UNIValue(0, 5) = strValidDt
	UNIValue(0, 6) = strValidDt
	UNIValue(0, 7) = strStatus
	UNIValue(0, 8) = strEBomFlg
	UNIValue(0, 9) = strMBomFlg
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

	If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)	'⊙: DB 에러코드, 메세지타입, %처리, 스크립트유형 
		rs0.Close
		Set rs0 = Nothing
		Set ADF = Nothing
		Response.End
	End If
	
%>

<Script Language=vbscript>
Dim LngLastRow
Dim LngMaxRow
Dim LngRow
Dim strTemp
Dim strData

With parent																	'☜: 화면 처리 ASP 를 지칭함 
	LngMaxRow = .vspdData.MaxRows										'Save previous Maxrow
		
<%  
	Dim i
    For i=0 to rs0.RecordCount-1 
		If i < C_SHEETMAXROWS Then
%>
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ECN_NO"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ECN_DESC"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("REASON_CD"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ISSUEDBY"))%>"
			
			If <%=ConvSPChars(rs0("ECN_STATUS"))%> = "1" Then
				strData = strData & Chr(11) & "Active"
			Else	'2
				strData = strData & Chr(11) & "Inactive"
			End If
			
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ECN_EBOM_FLG"))%>"
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("EBOM_DT"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ECN_MBOM_FLG"))%>"
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("MBOM_DT"))%>"
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("VALID_FROM_DT"))%>"
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("VALID_TO_DT"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("REMARK"))%>"
			strData = strData & Chr(11) & LngMaxRow + <%=i%>
			strData = strData & Chr(11) & Chr(12)
<%		
			rs0.MoveNext
		End If
	Next
%>
		.ggoSpread.Source = .vspdData
		.ggoSpread.SSShowDataByClip strData

		.lgStrPrevKey = "<%=Trim(rs0("ECN_NO"))%>"		
<%			
		rs0.Close
		Set rs0 = Nothing
%>
		.hEcnNo.value		= "<%=ConvSPChars(Request("txtEcnNo"))%>"
		.hReasonCd.value	= "<%=ConvSPChars(Request("txtReasonCd"))%>"
		.hValidDt.value		= "<%=UNIDateClientFormat(Request("txtValidDt"))%>"
		.hStatus.value		= "<%=ConvSPChars(Request("cboStatus"))%>"
		.hEBomFlg.value		= "<%=ConvSPChars(Request("cboEBomFlg"))%>"
		.hMBomFlg.value		= "<%=ConvSPChars(Request("cboMBomFlg"))%>"

	.DbQueryOk
End With	
</Script>	
<%
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
