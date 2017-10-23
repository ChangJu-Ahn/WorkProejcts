<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4321mb1
'*  4. Program Name         : List BackLog 
'*  5. Program Desc         : 
'*  6. Modified date(First) : 2006-04-11
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     :HJO
'*  9. Modifier (Last)      : 
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "P", "NOCOOKIE","MB")

On Error Resume Next

Dim ADF										'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg								'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter 선언 
Dim rs0, rs1, rs2, rs3, rs4, rs5			'DBAgent Parameter 선언 
Dim strMode									'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim strQryMode
Dim lgStrPrevKey,lgStrPrevKey1
Dim lgStrPrevKey2,lgStrPrevKey3,lgStrPrevKey4
Dim i

Const C_SHEETMAXROWS_D = 100
'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================
Call HideStatusWnd

strMode = Request("txtMode")												'☜ : 현재 상태를 받음 
strQryMode = Request("lgIntFlgMode")
lgStrPrevKey = FilterVar(UCase(Request("lgStrPrevKey")), "''", "S")
lgStrPrevKey1 = FilterVar(UCase(Request("lgStrPrevKey1")), "''", "S")
lgStrPrevKey2 = FilterVar(UCase(Request("lgStrPrevKey2")), "''", "S")
lgStrPrevKey3 = FilterVar(UCase(Request("lgStrPrevKey3")), "''", "S")
lgStrPrevKey4 = FilterVar(UCase(Request("lgStrPrevKey4")), "''", "S")

On Error Resume Next

Dim strItemCd
Dim StrProdOrderNo
Dim StrTrackingNo
Dim strCompleteFlag
Dim strStartDt
Dim strEndDt
Dim strFlag

Err.Clear																	'☜: Protect system from crashing
	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================
	Redim UNISqlId(3)
	Redim UNIValue(3, 0)

	UNISqlId(0) = "180000saa"
	UNISqlId(1) = "180000sab"	
	UNISqlId(2) = "180000sam" 
	
	
	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(1, 0) = FilterVar(UCase(Request("txtItemCd")), "''", "S")	
	UNIValue(2, 0) = FilterVar(UCase(Request("txtTrackingNo")), "''", "S")


	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1, rs2, rs3)

	%>
	<Script Language=vbscript>
		parent.frm1.txtPlantNm.value = ""
		parent.frm1.txtItemNm.value = ""
	</Script>	
	<%

	' Plant 명 Display  
	strFlag=""    
	If (rs1.EOF And rs1.BOF) Then
		strFlag = "ERROR_PLANT"		
	Else
		%>
		<Script Language=vbscript>
			parent.frm1.txtPlantNm.value = "<%=ConvSPChars(rs1("PLANT_NM"))%>"
		</Script>	
		<%    		
	End If
	
		'품목명 Display
	IF strFlag ="" and Request("txtItemCd") <> "" Then
		If (rs2.EOF And rs2.BOF) Then
			strFlag = "ERROR_ITEM"
		Else
			%>
			<Script Language=vbscript>
				parent.frm1.txtItemNm.value = "<%=ConvSPChars(rs2("ITEM_NM"))%>"
			</Script>	
			<%
		End If
	End IF
    'tracking no display
	IF strFlag="" and Request("txtTrackingNo") <> "" Then
		If (rs3.EOF And rs3.BOF) Then
			strFlag = "ERROR_TRACK"
		End If
	End IF
	
	rs1.Close	:	Set rs1 = Nothing
	rs2.Close	:	Set rs2 = Nothing
	rs3.Close	:	Set rs3 = Nothing
	
	If strFlag <> "" Then
		If strFlag = "ERROR_PLANT" Then
			Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
			parent.frm1.txtPlantCd.Focus()
			</Script>	
			<%
			Set ADF = Nothing
			Response.End
			
		ElseIf strFlag = "ERROR_ITEM" Then
			Call DisplayMsgBox("122600", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
			parent.frm1.txtItemCd.Focus()
			</Script>	
			<%
			Set ADF = Nothing
			Response.End
			
		ElseIf strFlag = "ERROR_TRACK" Then
			Call DisplayMsgBox("203045", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
			parent.frm1.txtTrackingNo.Focus()
			</Script>	
			<%
			Set ADF = Nothing
			Response.End		
		End If
	End IF
	
	' Order Header Display
	Redim UNISqlId(0)
	Redim UNIValue(0, 7)

	UNISqlId(0) = "P4321MA1S"
	
	
	IF Request("txtItemCd") = "" Then
		strItemCd = "|"
	Else
		StrItemCd = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	End IF

	IF Request("txtTrackingNo") = "" Then
		strTrackingNo = "|"
	Else
		StrTrackingNo = FilterVar(UCase(Request("txtTrackingNo")), "''", "S")
	End IF

	IF Request("txtProdFromDt") = "" Then
		strStartDt = "" & FilterVar("1900-01-01", "''", "S") & ""
	Else
		strStartDt = " " & FilterVar(UniConvDate(Request("txtProdFromDt")), "''", "S") & ""
	End IF

	IF Request("txtProdTODt") = "" Then
		strEndDt = "" & FilterVar("2999-12-31", "''", "S") & ""
	Else
		strEndDt = " " & FilterVar(UniConvDate(Request("txtProdTODt")), "''", "S") & ""
	End IF

	IF Request("txtProdOrdNo") <> "" Then
		StrProdOrderNo = FilterVar(UCase(Request("txtProdOrdNo")), "''", "S")
	End IF
	
	IF Request("txtrdoflag") = "A" Then
		strCompleteFlag = " (a.status <> " & FilterVar("R", "''", "S") & " AND a.status <> "  & FilterVar("D", "''", "S") & ") "
	ElseIf Request("txtrdoflag")="N" Then 
		strCompleteFlag = " a.status in ( " & FilterVar(Request("txtrdoflag"), "''", "S") & " ,'I' ) "
	Else
		strCompleteFlag = " a.status = " & FilterVar(Request("txtrdoflag"), "''", "S") & "  "
	End IF	
	
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 2) = strStartDt
	UNIValue(0, 3) = strEndDt	
	UNIValue(0, 4) = strItemCd 
	UNIValue(0, 5) = strTrackingNo

	Select Case strQryMode
		Case CStr(OPMD_CMODE)
			If Request("txtProdOrdNo") = "" Then
				UNIValue(0, 6) = "|"
			Else 
				UNIValue(0, 6) = " a.prodt_order_no >= " & strProdOrderNo	
			End If
		Case CStr(OPMD_UMODE) 
			UNIValue(0, 6) = " (a.report_dt > " & lgStrPrevKey2  
			UNIValue(0, 6) = UNIValue(0, 6) & " or (a.report_dt = " & lgStrPrevKey2 
			UNIValue(0, 6) = UNIValue(0, 6) & " and a.prodt_order_no > " & lgStrPrevKey & ")"	
			UNIValue(0, 6) = UNIValue(0, 6) & " or (a.report_dt = " & lgStrPrevKey2 
			UNIValue(0, 6) = UNIValue(0, 6) & " and a.prodt_order_no = " & lgStrPrevKey		
			UNIValue(0, 6) = UNIValue(0, 6) & " and a.opr_no > " & lgStrPrevKey1  & ")"	
			UNIValue(0, 6) = UNIValue(0, 6) & " or (a.report_dt >= " & lgStrPrevKey2 
			UNIValue(0, 6) = UNIValue(0, 6) & " and a.prodt_order_no = " & lgStrPrevKey		
			UNIValue(0, 6) = UNIValue(0, 6) & " and a.opr_no = " & lgStrPrevKey1 
			UNIValue(0, 6) = UNIValue(0, 6) & " and a.resv_seq > " & lgStrPrevKey3  & ")"	
			UNIValue(0, 6) = UNIValue(0, 6) & " or (a.report_dt >= " & lgStrPrevKey2 
			UNIValue(0, 6) = UNIValue(0, 6) & " and a.prodt_order_no = " & lgStrPrevKey		
			UNIValue(0, 6) = UNIValue(0, 6) & " and a.opr_no = " & lgStrPrevKey1 
			UNIValue(0, 6) = UNIValue(0, 6) & " and a.resv_seq = " & lgStrPrevKey3
			UNIValue(0, 6) = UNIValue(0, 6) & " and a.result_seq >= " & lgStrPrevKey4 & ")) "
	End Select	
	UNIValue(0, 7) = strCompleteFlag
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
      
	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		%>
		<Script Language=vbscript>
			parent.DbQueryNotOk()
		</Script>	
		<%		
		Response.End													'☜: 비지니스 로직 처리를 종료함 
	End If
	
%>

<Script Language=vbscript>
Dim LngMaxRow
Dim strData
Dim TmpBuffer
Dim iTotalStr
    	
With parent																	'☜: 화면 처리 ASP 를 지칭함 
	LngMaxRow = .frm1.vspdData1.MaxRows										'Save previous Maxrow
		
<%  
	If Not(rs0.EOF And rs0.BOF) Then		
		If C_SHEETMAXROWS_D < rs0.RecordCount Then 
%>
			ReDim TmpBuffer(<%=C_SHEETMAXROWS_D - 1%>)
<%
		Else
%>			
			ReDim TmpBuffer(<%=rs0.RecordCount - 1%>)
<%
		End If
		
		For i=0 to rs0.RecordCount-1
			If i < C_SHEETMAXROWS_D Then
%>				
				strData = ""
				strData = strData & Chr(11) & Trim("<%=ConvSPChars(rs0("CHK"))%>")									'오더번호 
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("REPORT_DT"))%>"									'착수예정일 
				strData = strData & Chr(11) & Trim("<%=ConvSPChars(rs0("PRODT_ORDER_NO"))%>")									'오더번호 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("OPR_NO"))%>"													'공정순서 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEM_CD"))%>"												'품목 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEM_NM"))%>"												'품목명 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("SPEC"))%>"													'규격 
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("ISSUE_QTY"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"				'오더수량 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("REMARK"))%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("GOOD_QTY"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"						'잔량 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("BASIC_UNIT"))%>"										'오더단위 
							
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("SL_CD"))%>"												'라우팅 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("SL_NM"))%>"													'작업장 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("TRACKING_NO"))%>"													'작업장명 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("REQ_NO"))%>"													'작업 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("RESV_SEQ"))%>"													'작업명 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("RESULT_SEQ"))%>"												'작업순서 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEM_DOCUMENT_NO"))%>"											'지시상태 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("STATUS"))%>"											'지시 				
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("STATUS_NM"))%>"	
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ERROR_DESC"))%>"														'MIlestone
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PLANT_CD"))%>"	
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("DOCUMENT_YEAR"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("WC_CD"))%>"	
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("LOT_NO"))%>"	
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("LOT_SUB_NO"))%>"	
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("COST_CD"))%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("SCHD_QTY"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"				'오더수량 
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("ORIGIN_QTY"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"				'오더수량  

				strData = strData & Chr(11) & LngMaxRow + <%=i%>
				strData = strData & Chr(11) & Chr(12)
				
				TmpBuffer(<%=i%>) = strData
<%		
				rs0.MoveNext
			End If
		Next
%>
		iTotalStr = Join(TmpBuffer,"")
		.ggoSpread.Source = .frm1.vspdData1
		.ggoSpread.SSShowDataByClip iTotalStr
		
		.lgStrPrevKey = Trim("<%=ConvSPChars(rs0("PRODT_ORDER_NO"))%>")
		.lgStrPrevKey1 = "<%=ConvSPChars(rs0("OPR_NO"))%>"
		.lgStrPrevKey2 = "<%=ConvSPChars(rs0("REPORT_DT"))%>"
		.lgStrPrevKey3 = "<%=ConvSPChars(rs0("RESV_SEQ"))%>"
		.lgStrPrevKey4 = "<%=ConvSPChars(rs0("RESULT_SEQ"))%>"
		
<%	
	End If

	rs0.Close
	Set rs0 = Nothing

%>	
	.frm1.hPlantCd.value		= "<%=ConvSPChars(Request("txtPlantCd"))%>"
	.frm1.hItemCd.value			= "<%=ConvSPChars(Request("txtItemCd"))%>"
	.frm1.hProdOrderNo.value	= "<%=ConvSPChars(Request("txtProdOrderNo"))%>"
	.frm1.hTrackingNo.value		= "<%=ConvSPChars(Request("txtTrackingNo"))%>"
	.frm1.hProdFromDt.value		= "<%=Request("txtProdFromDt")%>"
	.frm1.hProdTODt.value		= "<%=Request("txtProdTODt")%>"
	.frm1.hrdoFlag.value		= "<%=Request("txtrdoflag")%>"
	
	.DbQueryOk(LngMaxRow+1)

End With

</Script>	
<%
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
