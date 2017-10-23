<%@LANGUAGE = VBScript%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : b3b22mb4.asp 
'*  4. Program Name         : Called By B3B22MA1 (Class Management)
'*  5. Program Desc         : Manage Class Information
'*  6. Modified date(First) : 2003/02/12
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Ryu Sung Won
'*  9. Modifier (Last)      : 
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================%>

<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "P", "NOCOOKIE", "MB")

Dim ADF												'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg										'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag			'DBAgent Parameter 선언 
Dim rs0, rs1, rs2, rs3								'DBAgent Parameter 선언 
Dim strQryMode
Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim strCharCd
Dim strCharChoice

strMode = Request("txtMode")												'☜ : 현재 상태를 받음 
strCharCd = UCase(Trim(Request("txtCharCd")))
strCharChoice = Trim(Request("CharChoice"))
'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================
Call HideStatusWnd

On Error Resume Next
Err.Clear
	
	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================
	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================
	Redim UNISqlId(1)
	Redim UNIValue(1, 0)

	UNISqlId(0) = "b3b21mb1a"
	
	UNIValue(0, 0) = " " & FilterVar(strCharCd, "''", "S") & ""

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

	If (rs0.EOF and rs0.BOF)Then
		'사양항목이 존재하지 않습니다.
		Call DisplayMsgBox("122630", vbOKOnly, "", "", I_MKSCRIPT)
		If strCharChoice = "1" Then
		%>
		<Script Language=vbscript>
			parent.frm1.txtCharCd1.focus
		</Script>
		<%		
		Else
		%>
		<Script Language=vbscript>
			parent.frm1.txtCharCd2.focus
		</Script>
		<%
		End If
		rs0.Close
		Set rs0 = Nothing
		Set ADF = Nothing
		Response.End
	End If
%>
<Script Language=vbscript>
	With parent.frm1
<%
	If strCharChoice = "1" Then
%>	
		.txtCharCd1.value = "<%=UCase(Trim(rs0("CHAR_CD")))%>"
		.txtCharNm1.value = "<%=rs0("CHAR_NM")%>"
		.txtCharValueDigit1.value = <%=rs0("CHAR_VALUE_DIGIT")%>
<%
	Else
%>
		.txtCharCd2.value = "<%=UCase(Trim(rs0("CHAR_CD")))%>"
		.txtCharNm2.value = "<%=rs0("CHAR_NM")%>"
		.txtCharValueDigit2.value = "<%=rs0("CHAR_VALUE_DIGIT")%>"
<%
	End If
%>
	End With
	
	Call parent.SetClassDigit
<%			
		rs0.Close
		Set rs0 = Nothing
%>
</Script>	
<%
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
