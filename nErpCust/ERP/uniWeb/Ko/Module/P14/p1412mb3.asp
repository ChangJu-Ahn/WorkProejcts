<%@LANGUAGE = VBScript%>
<%'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1412mb3.asp
'*  4. Program Name         : Look Up Ecn Info
'*  5. Program Desc         :
'*  6. Comproxy List        : + 
'*  7. Modified date(First) : 2003/03/20
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Ryu Sung Won
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'*************************************************************************************************%>

<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("I", "*", "NOCOOKIE", "MB")

Dim ADF												'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg										'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag			'DBAgent Parameter 선언 
Dim rs0, rs1, rs2, rs3								'DBAgent Parameter 선언 
Dim strQryMode
Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim strEcnNo
Dim strReasonCd
Dim strReasonNm
Dim strEcnDesc
Dim strTarget
Dim Row 

Call HideStatusWnd

On Error Resume Next
Err.Clear
	
	strMode = Request("txtMode")												'☜ : 현재 상태를 받음 
	strTarget = Request("txtTarget")
	Row = Request("Row")	

	Redim UNISqlId(1)
	Redim UNIValue(1, 1)

	UNISqlId(0) = "p1412mb3a"
	UNISqlId(1) = "AMINORNM"
	
	UNIValue(0, 0) = FilterVar(Request("txtEcnNo"), "''", "S")
	UNIValue(1, 0) = FilterVar("P1402","''","S")
	UNIValue(1, 1) = FilterVar(Request("txtReasonCd"),"''","S")


	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)

	If (rs0.EOF and rs0.BOF)Then
		'Call DisplayMsgBox("182801", vbOKOnly, "", "", I_MKSCRIPT)
		strReasonCd = ""
		strReasonNm = ""
		strEcnDesc = ""
		rs0.Close
		Set rs0 = Nothing
	Else
		strReasonCd = UCase(Trim(rs0("REASON_CD")))
		strReasonNm = rs0("REASON_NM")
		strEcnDesc = rs0("ECN_DESC")
		rs0.Close
		Set rs0 = Nothing
	End If
	
	If Trim(strTarget) = "REASON" Then
		If (rs1.EOF and rs1.BOF)Then
			Call DisplayMsgBox("182803", vbOKOnly, "", "", I_MKSCRIPT)
			strReasonNm = ""
			rs1.Close
			Set rs1 = Nothing
			Set ADF = Nothing
		Else
			strReasonNm = ConvSPChars(rs1("MINOR_NM"))
			rs1.Close
			Set rs1 = Nothing
		End If
	End If
%>
<Script Language=vbscript>
	Call parent.LookUpEcnInfoOk("<%=strReasonCd%>","<%=ConvSPChars(strReasonNm)%>","<%=ConvSPChars(strEcnDesc)%>","<%=ConvSPChars(strTarget)%>","<%=Row%>")
</Script>	
<%
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
