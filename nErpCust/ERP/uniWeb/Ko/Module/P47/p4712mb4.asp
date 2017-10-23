<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4712mb4.asp
'*  4. Program Name         : List RESOURCE infomation	
'*  5. Program Desc         : 
'*  6. Modified date(First) : 2001.12.12
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Jeon, Jaehyun
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
'Call loadInfTB19029("I", "*")
Call loadInfTB19029B("I", "*", "NOCOOKIE","MB")

On Error Resume Next

Dim ADF										'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg								'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter 선언 
Dim rs0, rs1, rs2								'DBAgent Parameter 선언 

'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'======================================================================================================
Dim LngRow

Call HideStatusWnd

Dim StrResourceCd
Dim strRcdNm
Dim strRcdtype
Dim strRgcd
Dim strRgcdNm

Err.Clear																	'☜: Protect system from crashing

	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================	
	
	' Order information Display
	Redim UNISqlId(0)
	Redim UNIValue(0, 2)

	UNISqlId(0) = "P4712MB4"
	
	IF Request("txtResourceCd") = "" Then
		StrResourceCd = "|"
	Else
		StrResourceCd = FilterVar(UCase(Request("txtResourceCd")), "''", "S")
	End IF

	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 2) = StrResourceCd 
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    
	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		LngRow = Request("Row")
		rs0.Close
		Set rs0 = Nothing
		%>
		<Script Language=vbscript>
		Call	parent.LookupRcNotOk("<%=LngRow%>")
		</Script>	
		<%		
		Response.End													'☜: 비지니스 로직 처리를 종료함 
	End If
	
	LngRow = Request("Row")
	strRcdNm = Trim(ConvSPChars(rs0("DESCRIPTION")))
	strRcdtype = Trim(ConvSPChars(rs0("MINOR_NM")))
	strRgcd = Trim(ConvSPChars(rs0("RESOURCE_GROUP_CD")))
    strRgcdNm = Trim(ConvSPChars(rs0("GROUP_NM")))

	
%>

<Script Language=vbscript>
	
	
<%	
	rs0.Close
	Set rs0 = Nothing

%>

  Call parent.LookupRcOk("<%=StrResourceCd%>", "<%=strRcdNm%>", "<%=strRcdtype%>", "<%=strRgcd%>", "<%=strRgcdNm%>", "<%=LngRow%>")


	
</Script>	
<%
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
