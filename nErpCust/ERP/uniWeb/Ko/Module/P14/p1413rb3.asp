<%@LANGUAGE = VBScript%>
<%'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1412rb3.asp
'*  4. Program Name         : Look Up Ecn Info
'*  5. Program Desc         :
'*  6. Comproxy List        : + 
'*  7. Modified date(First) : 2003/03/27
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
Call LoadInfTB19029B("I", "*", "NOCOOKIE", "RB")

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

Call HideStatusWnd

On Error Resume Next
Err.Clear


	strMode = Request("txtMode")												'☜ : 현재 상태를 받음 

	Redim UNISqlId(0)
	Redim UNIValue(0, 0)

	UNISqlId(0) = "p1412mb3a"
	
	UNIValue(0, 0) = FilterVar(Request("txtEcnNo"), "''", "S")

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

	If (rs0.EOF and rs0.BOF)Then
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

%>
<Script Language=vbscript>
	Call parent.LookUpEcnInfoOk("<%=strReasonCd%>","<%=strReasonNm%>","<%=strEcnDesc%>")
</Script>	
<%
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
