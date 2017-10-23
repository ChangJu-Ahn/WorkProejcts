<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4114mb4.asp	
'*  4. Program Name         : look up Work Center
'*  5. Program Desc         :
'*  6. Comproxy List        : DB Agent
'*  7. Modified date(First) : 2000/03/31
'*  8. Modified date(Last)  : 2002/06/29
'*  9. Modifier (First)     : Mr  Kim
'* 10. Modifier (Last)      : Park, BumSoo
'* 11. Comment              :
'**********************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

Call LoadBasisGlobalInf

On Error Resume Next								'☜: 

Dim ADF												'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg										'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0		'DBAgent Parameter 선언 
Dim strWcCd, strWcNm
Dim Row, Row1
Dim strProdtOrderNo, strOprNo, strInsideFlg

'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================
Call HideStatusWnd

    Err.Clear															'☜: Protect system from crashing

	Row = Request("Row")
	Row1 = Request("Row1")
	strProdtOrderNo = Request("txtProdtOrderNo")
	strOprNo = Request("txtOprNo")
	strWcCd = Trim(Request("txtWcCd"))

	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================
	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================
	Redim UNISqlId(0)
	Redim UNIValue(0, 1)
	
	UNISqlId(0) = "p4114mb4"
	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 1) = FilterVar(UCase(Request("txtWcCd")), "''", "S")

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
	
	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("182100", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		Set ADF = Nothing
		%>
		<Script Language=vbscript>
			Call parent.LookUpWcNotOk("<%=Row%>")
		</Script>
		<%	   
		Response.End													'☜: 비지니스 로직 처리를 종료함 
	Else
		strWcNm		 = ConvSPChars(rs0("Wc_Nm"))
		strInsideFlg = UCase(rs0("Inside_Flg"))
		%>
		<Script Language=vbscript>
			Call parent.LookUpWcOk("<%=strWcCd%>", "<%=strWcNm%>","<%=strInsideFlg%>","<%=Row%>","<%=Row1%>","<%=strProdtOrderNo%>","<%=strOprNo%>")
		</Script>
		<%
	End If

Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
