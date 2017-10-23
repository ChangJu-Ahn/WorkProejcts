<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4220mb1.asp 
'*  4. Program Name         : Resource Plan By Production Order
'*  5. Program Desc         : List Production Order
'*  6. Modified date(First) : 2002/03/04
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Park, Bum Soo
'*  9. Modifier (Last)      : Chen, Jae Hyun
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call LoadBasisGlobalInf
On Error Resume Next								'☜: 

Dim ADF												'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg										'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1	'DBAgent Parameter 선언 
Dim strQryMode

Const C_SHEETMAXROWS = 100

Dim strMode											'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================
Dim StrNextKey		' 다음 값 
Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngRow
Dim strStartDt
Dim strEndDt
Dim strProdOrderNo
Dim i

Call HideStatusWnd

strQryMode = Request("lgIntFlgMode")

On Error Resume Next
	
	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================
	
	Redim UNISqlId(1)
	Redim UNIValue(1, 4)
	
	UNISqlId(0) = "180000saa"
	UNISqlId(1) = "189701saa"	
	
	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	
	If Request("txtStartDt") = "" Then
		strStartDt = "|"
	Else
		strStartDt = " " & FilterVar(UniConvDate(Request("txtStartDt")), "''", "S") & ""
	End If
	
	If Request("txtEndDt") = "" Then
		strEndDt = "|"
	Else
		strEndDt = " " & FilterVar(UniConvDate(Request("txtEndDt")), "''", "S") & ""
	End If
	
	If Request("lgStrPrevKey1") = "" Then
		strProdOrderNo = "|"
	Else
		strProdOrderNo = FilterVar(UCase(Request("lgStrPrevKey1")), "''", "S")
	End If
	
	UNIValue(1, 0) = "^"
	UNIValue(1, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(1, 2) = strStartDt
	UNIValue(1, 3) = strEndDt
	
	Select Case strQryMode
		Case CStr(OPMD_CMODE)
			UNIValue(1, 4) = "|"
		Case CStr(OPMD_UMODE) 
			UNIValue(1, 4) = strProdOrderNo
	End Select

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)

	' Plant 명 Display      
	If (rs0.EOF And rs0.BOF) Then
		rs0.Close
		Set rs0 = Nothing
		Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
		%>
		<Script Language=vbscript>
			parent.frm1.txtPlantNm.value = ""
			parent.frm1.txtPlantCd.Focus()
		</Script>	
		<%    	
		Response.End													'☜: 비지니스 로직 처리를 종료함 
	Else
		%>
		<Script Language=vbscript>
			parent.frm1.txtPlantNm.value = "<%=ConvSPChars(rs0("PLANT_NM"))%>"
		</Script>	
		<%    	
		rs0.Close
		Set rs0 = Nothing
	End If
      
	If rs1.EOF And rs1.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		Set rs1 = Nothing					
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
	LngMaxRow = .frm1.vspdData1.MaxRows										'Save previous Maxrow
		
<%  
    For i=0 to rs1.RecordCount-1 
		If i < C_SHEETMAXROWS Then
%>
			strData = strData & Chr(11) & "<%=ConvSPChars(rs1("prodt_order_no"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs1("item_cd"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs1("item_nm"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs1("rout_no"))%>"
			strData = strData & Chr(11) & "<%=UniDateClientFormat(rs1("plan_start_dt"))%>"
			strData = strData & Chr(11) & "<%=UniDateClientFormat(rs1("plan_compt_dt"))%>"
			strData = strData & Chr(11) & "<%=UniDateClientFormat(rs1("schd_start_dt"))%>"
			strData = strData & Chr(11) & "<%=UniDateClientFormat(rs1("schd_compt_dt"))%>"
			strData = strData & Chr(11) & "<%=UniDateClientFormat(rs1("real_start_dt"))%>"
			strData = strData & Chr(11) & LngMaxRow + <%=i%>			
			strData = strData & Chr(11) & Chr(12)
<%		
			rs1.MoveNext
		End If
	Next
%>
	
		.ggoSpread.Source = .frm1.vspdData1
		.ggoSpread.SSShowDataByClip strData
		
		
		.lgStrPrevKey1 = "<%=Trim(rs1("PRODT_ORDER_NO"))%>"
		
		.frm1.hPlantCd.value = "<%=ConvSPChars(Request("txtPlantCd"))%>"
		.frm1.hStartDt.value = "<%=Request("txtStartDt")%>"
		.frm1.hEndDt.value = "<%=Request("txtEndDt")%>"		
<%			
		rs1.Close
		Set rs1 = Nothing
%>
	.DbQueryOk
End With	
</Script>	
<%
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
