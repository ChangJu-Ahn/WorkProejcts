<%@LANGUAGE = VBScript%>
<%'======================================================================================================
'*  1. Module Name          : Product
'*  2. Function Name        : 
'*  3. Program ID           : p1410mb9.asp 
'*  4. Program Name         : 
'*  5. Program Desc         : Query ECN Info.
'*  6. Modified date(First) : 2003/03/07
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Ryu Sung Won
'*  9. Modifier (Last)      : 
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================%>

<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadinfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadinfTB19029B("Q", "P", "NOCOOKIE", "MB")

Dim ADF														'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg
Dim UNISqlId, UNIValue, UNILock, UNIFlag
Dim rs0, rs1, rs2, rs3
Dim iIntCnt, strQryMode
Dim strData
Dim TmpBuffer
Dim iTotalStr

Const C_SHEETMAXROWS_D = 100

Call HideStatusWnd

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
Err.Clear

	strQryMode = Request("lgIntFlgMode")
	lgStrPrevKey = Request("lgStrPrevKey")
	LngMaxRow = Request("txtMaxRows")


	Redim UNISqlId(1)
	Redim UNIValue(1, 1)

	UNISqlId(0) = "p1410mb9b"
	UNISqlId(1) = "s0000qa000"
	
	UNIValue(0, 0) = FilterVar(UCase(Request("txtECNNo")), "''", "S")
	UNIValue(1, 0) = FilterVar("P1402","''","S")
	UNIValue(1, 1) = FilterVar(Request("txtReasonCd"),"''","S")

	UNILock = DISCONNREAD :	UNIFlag = "1"

	Set ADF = Server.CreateObject("prjPublic.cCtlTake")
	strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1, rs2)

	'======================================================================================================
	' ECN NO가 존재하는지 체크해서 있으면 명을 넣어주고 없으면 SKIP한다.
	'======================================================================================================
	If (rs1.EOF And rs1.BOF) Then
		Response.Write "<Script Language=vbscript>		" & vbCr
		Response.Write "parent.frm1.txtECNNoDesc.value = """"" & vbCr
		Response.Write "</Script>						" & vbCr
	Else
		Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtECNNoDesc.value = """ & ConvSPChars(rs1("ECN_DESC")) & """" & vbCrLf
		Response.Write "</Script>" & vbCrLf
	End If

	rs1.Close
	Set rs1 = Nothing

	'======================================================================================================
	' 설계변경근거가 존재하는지 체크 
	'======================================================================================================
	If (rs2.EOF And rs2.BOF) Then
		IF Request("txtReasonCd") = "" Then
			Response.Write "<Script Language=vbscript>		" & vbCr
			Response.Write "parent.frm1.txtReasonDesc.value = """"" & vbCr													
			Response.Write "</Script>						" & vbCr
		ELSE			
			Call DisplayMsgBox("182803", vbOKOnly, "", "", I_MKSCRIPT)
			rs2.Close
			Set rs2 = Nothing
			Response.Write "<Script Language=vbscript>		" & vbCr
			Response.Write "parent.frm1.txtReasonDesc.value = """"" & vbCr													
			Response.Write "</Script>						" & vbCr
			Response.End												'☜: 비지니스 로직 처리를 종료함	
		END IF			
	Else
		Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtReasonDesc.value = """ & ConvSPChars(rs2("MINOR_NM")) & """" & vbCrLf
		Response.Write "</Script>" & vbCrLf
	End If

	rs2.Close
	Set rs2 = Nothing
	Set ADF = Nothing

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
	
	IF Trim(Request("txtEcnNoDesc")) = "" Then
	   strEcnDesc = "|"
	ELSE
	   strEcnDesc = FilterVar(UCase(Request("txtEcnNoDesc")), "''", "S")
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
	
	Response.Write "<Script Language = VBScript>" & vbCrLf
	Response.Write "With parent" & vbCrLf
	
    If Not(rs0.EOF And rs0.BOF) Then

		If C_SHEETMAXROWS_D < rs0.RecordCount Then 

			ReDim TmpBuffer(C_SHEETMAXROWS_D - 1)

		Else

			ReDim TmpBuffer(rs0.RecordCount - 1)

		End If

		For iIntCnt = 0 To rs0.RecordCount - 1 
			If iIntCnt < C_SHEETMAXROWS_D Then
				strData = ""
				strData = strData & Chr(11) & ConvSPChars(rs0("ECN_NO"))
				strData = strData & Chr(11) & ConvSPChars(rs0("ECN_DESC"))
				strData = strData & Chr(11) & ConvSPChars(rs0("REASON_CD"))
				strData = strData & Chr(11) & ConvSPChars(rs0("ISSUEDBY"))
				
				If ConvSPChars(rs0("ECN_STATUS")) = "1" Then
				strData = strData & Chr(11) & "Active"
				Else '2
				strData = strData & Chr(11) & "Inactive"
				End If
				
				strData = strData & Chr(11) & ConvSPChars(rs0("ECN_EBOM_FLG"))
				strData = strData & Chr(11) & UNIDateClientFormat(rs0("EBOM_DT"))
				strData = strData & Chr(11) & ConvSPChars(rs0("ECN_MBOM_FLG"))
				strData = strData & Chr(11) & UNIDateClientFormat(rs0("MBOM_DT"))
				strData = strData & Chr(11) & UNIDateClientFormat(rs0("VALID_FROM_DT"))
				strData = strData & Chr(11) & UNIDateClientFormat(rs0("VALID_TO_DT"))		
				strData = strData & Chr(11) & ConvSPChars(rs0("INSRT_USER_ID"))
				strData = strData & Chr(11) & UNIDateClientFormat(rs0("INSRT_DT"))
				strData = strData & Chr(11) & ConvSPChars(rs0("REMARK"))
				
		        strData = strData & Chr(11) & (LngMaxRow + iIntCnt)
				strData = strData & Chr(11) & Chr(12)
				
				TmpBuffer(iIntCnt) = strData
				
				rs0.MoveNext
			End If
		Next
		
		iTotalStr = Join(TmpBuffer, "")
		
		Response.Write ".ggoSpread.Source = .frm1.vspdData" & vbCrLf
		Response.Write ".ggoSpread.SSShowDataByClip """ & iTotalStr & """" & vbCrLf
		
		If rs0("ECN_NO") = Null Then
			Response.Write ".lgStrPrevKey = """"" & vbCrLf
		Else
			Response.Write ".lgStrPrevKey = """ & Trim(rs0("ECN_NO")) & """" & vbCrLf
		End If
	End If	

	rs0.Close
	Set rs0 = Nothing
	
	Response.Write ".frm1.hEcnNo.value		= """ & ConvSPChars(Request("txtEcnNo")) & """" & vbCrLf
	Response.Write ".frm1.hReasonCd.value	= """ & ConvSPChars(Request("txtReasonCd")) & """" & vbCrLf
	Response.Write ".frm1.hValidDt.value	= """ & UNIDateClientFormat(Request("txtValidDt")) & """" & vbCrLf
	Response.Write ".frm1.hStatus.value		= """ & ConvSPChars(Request("cboStatus")) & """" & vbCrLf
	Response.Write ".frm1.hEBomFlg.value	= """ & ConvSPChars(Request("cboEBomFlg")) & """" & vbCrLf
	Response.Write ".frm1.hMBomFlg.value	= """ & ConvSPChars(Request("cboMBomFlg")) & """" & vbCrLf

	Response.Write ".DbQueryOk" & vbCrLf
	Response.Write "End With" & vbCrLf
	Response.Write "</Script>" & vbCrLf

Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
