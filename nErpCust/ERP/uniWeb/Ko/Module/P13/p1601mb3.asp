<%@LANGUAGE = VBScript%>
<%'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1601mb3.asp
'*  4. Program Name         : Look Up VAT Type 
'*  5. Program Desc         :
'*  6. Comproxy List        : + 
'*  7. Modified date(First) : 2002/01/15
'*  8. Modified date(Last)  : 2002/10/07
'*  9. Modifier (First)     : Jung Yu Kyung
'* 10. Modifier (Last)      : Ryu Sung Won
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
Dim strVatType
Dim strVatRate
Dim Row 

strMode = Request("txtMode")												'☜ : 현재 상태를 받음 
strVatType = Trim(Request("txtVatType"))
'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================
Call HideStatusWnd

'On Error Resume Next
Err.Clear
	
	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================
	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================
	Redim UNISqlId(1)
	Redim UNIValue(1, 1)

	UNISqlId(0) = "p1601mb3a"
	UNISqlId(1) = "p1601mb3b"
	
	UNIValue(0, 0) = "" & FilterVar("B9001", "''", "S") & ""
	UNIValue(0, 1) = " " & FilterVar(strVatType, "''", "S") & ""
	UNIValue(1, 0) = "" & FilterVar("B9001", "''", "S") & ""	
	UNIValue(1, 1) = " " & FilterVar(strVatType, "''", "S") & ""

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)

	If (rs0.EOF and rs0.BOF)Then
		Call DisplayMsgBox("115100", vbOKOnly, strVatType, "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		Set ADF = Nothing
		Response.End
	End If

	If (rs1.EOF and rs1.BOF)Then
		Call DisplayMsgBox("115100", vbOKOnly, strVatType, "", I_MKSCRIPT)
		rs1.Close
		Set rs1 = Nothing
		Set ADF = Nothing
		Response.End
	End If

	strVatRate = UniConvNumberDBToCompany(rs0("reference"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
	Row = Request("Row")

%>
<Script Language=vbscript>
	With parent.frm1.vspdData
		Call parent.LookUpVatTypeOk("<%=strVatType%>","<%=strVatRate%>","<%=Row%>")
	End With
<%			
		rs0.Close
		Set rs0 = Nothing
		rs1.Close
		Set rs1 = Nothing
%>
</Script>	
<%
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
