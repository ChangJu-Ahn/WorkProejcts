<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4221mb2.asp
'*  4. Program Name         : Resource Plan
'*  5. Program Desc         : List Resource Plan
'*  6. Modified date(First) : 
'*  7. Modified date(Last)  : 2002/08/21
'*  8. Modifier (First)     : Hong, EunSook
'*  9. Modifier (Last)      : Park, BumSoo
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call LoadBasisGlobalInf
On Error Resume Next								'☜: 

Dim ADF												'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg										'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0		'DBAgent Parameter 선언 

Dim strMode											'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================
Dim StrNextKey		' 다음 값 
Dim lgStrPrevKey	' 이전 값 
Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngRow
Dim strStartDt
Dim strEndDt
Dim i

Call HideStatusWnd

	lgStrPrevKey = Request("lgStrPrevKey2")
	
	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================
	
	Redim UNISqlId(0)
	Redim UNIValue(0, 4)
	
	UNISqlId(0) = "189700sac"
	
	If Request("txtStartDt") = "" Then
		strStartDt = "|"
	Else
		strStartDt =  " " & FilterVar(UniConvDate(Request("txtStartDt")), "''", "S") & ""
	End If
	
	If Request("txtEndDt") = "" Then
		strEndDt = "|"
	Else
		strEndDt   =  " " & FilterVar(UniConvDate(Request("txtEndDt")), "''", "S") & ""
	End If

	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 2) = FilterVar(UCase(Request("txtResourceGroupCd")), "''", "S")
	UNIValue(0, 3) = strStartDt
	UNIValue(0, 4) = strEndDt
		
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
      
	If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		
		rs0.Close
		Set rs0 = Nothing
					
		Response.End													'☜: 비지니스 로직 처리를 종료함 
	End If
	
%>
<Script Language=vbscript>
Dim LngLastRow
Dim LngMaxRow
Dim LngRow
Dim strTemp
Dim strData
    	
With parent																	'☜: 화면 처리 ASP 를 지칭함 
	LngMaxRow = .frm1.vspdData2.MaxRows										'Save previous Maxrow
		
<%  
    For i=0 to rs0.RecordCount-1 
%>
		
		strData = strData & Chr(11) & "<%=UniDateClientFormat(rs0("start_dt"))%>"
		strData = strData & Chr(11) & "<%=UniDateClientFormat(rs0("end_dt"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("prodt_order_no"))%>"
		StrDate = strData & Chr(11) & "<%=ConvSPChars(rs0("load_no"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_cd"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("rout_no"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("opr_no"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("wc_cd"))%>"
		strData = strData & Chr(11) & "<%=rs0("start_flg")%>"
		strData = strData & Chr(11) & "<%=rs0("end_flg")%>"
		strData = strData & Chr(11) & LngMaxRow + <%=i%>
		strData = strData & Chr(11) & Chr(12)
		
<%		
		rs0.MoveNext
	Next
%>
	
		.ggoSpread.Source = .frm1.vspdData2
		.ggoSpread.SSShowDataByClip strData		

		.lgStrPrevKey2 = ""	
		.frm1.hStartDt.value = "<%=Request("txtStartDt")%>"
		.frm1.hEndDt.value = "<%=Request("txtEndDt")%>"
		
<%			
		rs0.Close
		Set rs0 = Nothing
%>
	.DbDtlQueryOk
End With	
</Script>	
<%
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
