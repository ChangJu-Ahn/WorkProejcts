<%'**********************************************************************************************
'*  1. Module Name			: INTERFACE
'*  2. Function Name		: 
'*  3. Program ID			: xi312mb1_ko119.asp
'*  4. Program Name			: 
'*  5. Program Desc			: 
'*  6. Comproxy List		: ADO 180000saa, 180000sab, 180000sad
'*						    :
'*  7. Modified date(First)	: 2006/04/24
'*  8. Modified date(Last)	: 
'*  9. Modifier (First)		: HJO
'* 10. Modifier (Last)		: 
'* 11. Comment				:
'**********************************************************************************************%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%					

Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "P","NOCOOKIE","MB")

Dim ADF												'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg										'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag			'DBAgent Parameter 선언 
Dim rs0, rs1, rs2, rs3, rs4							'DBAgent Parameter 선언 
Dim strQryMode

Const C_SHEETMAXROWS_D = 100

'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================
Dim StrNextKey										' 다음 값 
Dim lgStrPrevKey1
Dim lgStrPrevKey2
Dim lgStrPrevKey3
Dim lgStrPrevKey4
Dim lgStrPrevKey5
Dim lgStrPrevKey6
Dim LngMaxRow										' 현재 그리드의 최대Row
Dim LngRow1
Dim GroupCount1
Dim strMode											'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim strFlag
Dim LngRow

Call HideStatusWnd

strMode = Request("txtMode")						'☜ : 현재 상태를 받음 
strQryMode = Request("lgIntFlgMode")

On Error Resume Next
Err.Clear

	
	lgStrPrevKey1 = Request("lgStrPrevKey1")
	lgStrPrevKey2 = Request("lgStrPrevKey2")
	lgStrPrevKey3 = Request("lgStrPrevKey3")
	lgStrPrevKey4 = Request("lgStrPrevKey4")
	lgStrPrevKey5 = Request("lgStrPrevKey5")
	lgStrPrevKey6 = Request("lgStrPrevKey6")

	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================
	Redim UNISqlId(1)
	Redim UNIValue(1, 0)

	UNISqlId(0) = "180000saa"
	UNISqlId(1) = "180000sab"

	
	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(1, 0) = FilterVar(UCase(Request("txtItemCd")), "''", "S")

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1, rs2)
	
	%>
		<Script Language=vbscript>
			parent.frm1.txtPlantNm.value = ""
			parent.frm1.txtItemNm.value = ""
		</Script>	
	<%
	strFlag=""
		
	' Plant 명 Display      
	If (rs1.EOF And rs1.BOF) Then
		rs1.Close
		Set rs1 = Nothing		
		strFlag = "ERROR_PLANT"
	Else
		%>
		<Script Language=vbscript>
			parent.frm1.txtPlantNm.value = "<%=ConvSPChars(rs1("PLANT_NM"))%>"
		</Script>	
		<%    	
		rs1.Close
		Set rs1 = Nothing
	End If
	' 품목명 Display
	IF strFlag="" and Request("txtItemCd") <> "" Then
		If (rs2.EOF And rs2.BOF) Then
			rs2.Close
			Set rs2 = Nothing
			strFlag = "ERROR_ITEM"
		Else
			%>
			<Script Language=vbscript>
				parent.frm1.txtItemNm.value = "<%=ConvSPChars(rs2("ITEM_NM"))%>"
			</Script>	
			<%
			rs2.Close
			Set rs2 = Nothing
		End If
	End IF
	
	If strFlag <> "" Then
		If strFlag = "ERROR_PLANT" Then
			Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
			parent.frm1.txtPlantCd.Focus()
			</Script>	
			<%
			set ADF = Nothing
			Response.End	
		ElseIf strFlag = "ERROR_ITEM" Then
		   Call DisplayMsgBox("122600", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
			parent.frm1.txtItemCd.Focus()
			</Script>	
			<%
			set ADF = Nothing
			Response.End
		End If
	End IF
 
    
	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=====================================================================================================
	
	Redim UNISqlId(0)
	Redim UNIValue(0, 19)
	
	UNISqlId(0) = "xi312mb1s_ko441"	'main query change id

	
	UNIValue(0, 0) = C_SHEETMAXROWS_D + 1
	UNIValue(0, 1) = FilterVar(UCase(Trim(Request("txtPlantCd"))),"''","S")
	UNIValue(0, 2) = FilterVar(UniConvDate(Trim(Request("txtSendStartDt"))) ,"''","S")
	UNIValue(0, 3) = FilterVar(UniConvDate(Trim(Request("txtSendEndDt"))) ,"''","S")
	
	If Trim(Request("txtItemCd")) = "" Then
		UNIValue(0, 4) = FilterVar("%","''","S")
	Else
		UNIValue(0, 4) = FilterVar(UCase(Trim(Request("txtItemCd"))),"''","S")
	End If 
	
	If Trim(Request("txtPlanStartDt")) = "" Then
		UNIValue(0, 5) = FilterVar("1900-01-01","''","S")
	Else
		UNIValue(0, 5) = FilterVar(UniConvDate(Trim(Request("txtPlanStartDt"))),"''","S")
	End If
	
	If Trim(Request("txtPlanEndDt")) = "" Then
		UNIValue(0, 6) = FilterVar("2999-12-31","''","S")
	Else
		UNIValue(0, 6) = FilterVar(UniConvDate(Trim(Request("txtPlanEndDt"))),"''","S")
	End If

	If Trim(Request("rdoFlag")) = "A" Then
		UNIValue(0, 7) = FilterVar("%", "''", "S") 
	Else
		UNIValue(0, 7) = FilterVar(UCase(Trim(Request("rdoFlag"))), "''", "S")
	End If
	
	If Trim(Request("rdoFlagF")) = "A" Then
		UNIValue(0, 8) = FilterVar("%", "''", "S") 
	Else
		UNIValue(0, 8) = FilterVar(UCase(Trim(Request("rdoFlagF"))), "''", "S")
	End If
	
	If Trim(Request("txtProdtOrderNo")) <>"" Then 
		UNIValue(0, 9) = FilterVar(UCase(Trim(Request("txtProdtOrderNo"))),"''","S")
	Else
		UNIValue(0, 9) = FilterVar("%", "''", "S") 
	End If
	
'20080122::hanc 	If Trim(Request("lgStrPrevKey1")) <>"" Then 
'20080122::hanc 		UNIValue(0, 9) = FilterVar(UCase(Trim(Request("lgStrPrevKey1"))),"''","S")
'20080122::hanc 		UNIValue(0, 10) = FilterVar(UCase(Trim(Request("lgStrPrevKey1"))),"''","S")
'20080122::hanc 	Else
'20080122::hanc 		UNIValue(0, 9) = FilterVar("", "''", "S")
'20080122::hanc 		UNIValue(0, 10) = FilterVar("", "''", "S")  
'20080122::hanc 	End If
'20080122::hanc 	
'20080122::hanc 	If Trim(Request("lgStrPrevKey2")) <>"" Then 
'20080122::hanc 		UNIValue(0, 11) = FilterVar(UCase(Trim(Request("lgStrPrevKey2"))),"''","S")
'20080122::hanc 		UNIValue(0, 12) = FilterVar(UCase(Trim(Request("lgStrPrevKey2"))),"''","S")
'20080122::hanc 	Else
'20080122::hanc 		UNIValue(0, 11) = FilterVar("", "''", "S")
'20080122::hanc 		UNIValue(0, 12) = FilterVar("", "''", "S")  
'20080122::hanc 	End If
'20080122::hanc 	
'20080122::hanc 	If Trim(Request("lgStrPrevKey3")) <>"" Then 
'20080122::hanc 		UNIValue(0, 13) = FilterVar(UniConvDate(Trim(Request("lgStrPrevKey3"))),"''","S")
'20080122::hanc 		UNIValue(0, 14) = FilterVar(UniConvDate(Trim(Request("lgStrPrevKey3"))),"''","S")
'20080122::hanc 	Else
'20080122::hanc 		UNIValue(0, 13) = FilterVar("1900-01-01","''","S")
'20080122::hanc 		UNIValue(0, 14) = FilterVar("1900-01-01","''","S")
'20080122::hanc 	End If
'20080122::hanc 	
'20080122::hanc 	If Trim(Request("lgStrPrevKey4")) <>"" Then 
'20080122::hanc 		UNIValue(0, 15) = FilterVar(UCase(Trim(Request("lgStrPrevKey4"))),"''","S")
'20080122::hanc 		UNIValue(0, 16) = FilterVar(UCase(Trim(Request("lgStrPrevKey4"))),"''","S")
'20080122::hanc 	Else
'20080122::hanc 		UNIValue(0, 15) = FilterVar("", "''", "S")
'20080122::hanc 		UNIValue(0, 16) = FilterVar("", "''", "S")  
'20080122::hanc 	End If
'20080122::hanc 	
'20080122::hanc 	If Trim(Request("lgStrPrevKey5")) <>"" Then 
'20080122::hanc 		UNIValue(0, 17) = FilterVar(UCase(Trim(Request("lgStrPrevKey5"))),"''","S")
'20080122::hanc 		UNIValue(0, 18) = FilterVar(UCase(Trim(Request("lgStrPrevKey5"))),"''","S")
'20080122::hanc 	Else
'20080122::hanc 		UNIValue(0, 17) = FilterVar("", "''", "S")
'20080122::hanc 		UNIValue(0, 18) = FilterVar("", "''", "S")  
'20080122::hanc 	End If
'20080122::hanc 	
'20080122::hanc 	If Trim(Request("lgStrPrevKey6")) = "" Then
'20080122::hanc 		UNIValue(0, 19) = "0"
'20080122::hanc 	Else
'20080122::hanc 		UNIValue(0, 19) = FilterVar(Trim(Request("lgStrPrevKey6")),"0","SNM")
'20080122::hanc 	End If
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    'Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
      
	If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		
		rs0.Close
		Set rs0 = Nothing
		%>
			<Script Language=vbscript>
			parent.DbQueryOk(0)
			</Script>	
			<%
		set ADF = Nothing
		Response.End														'☜: 비지니스 로직 처리를 종료함 
	End If
%>

<Script Language=vbscript>
Dim LngMaxRow
Dim strData
Dim TmpBuffer
Dim iTotalStr
Dim strTime

With parent																	'☜: 화면 처리 ASP 를 지칭함 
	LngMaxRow = .frm1.vspdData.MaxRows										'Save previous Maxrow
	.lgStrPrevKey1 = ""
	.lgStrPrevKey2 = ""
	.lgStrPrevKey3 = ""
	.lgStrPrevKey4 = ""
	.lgStrPrevKey5 = ""
	.lgStrPrevKey6 = ""
		
<%
	If C_SHEETMAXROWS_D < rs0.RecordCount Then 
%>
		ReDim TmpBuffer(<%=C_SHEETMAXROWS_D - 1%>)
<%
	Else
%>			
		ReDim TmpBuffer(<%=rs0.RecordCount - 1%>)
<%
	End If
	 
	LngRow = 0
	
    Do While Not rs0.EOF
		If LngRow < C_SHEETMAXROWS_D Then
			
%>
			strTime = "<%=Trim(ConvSPChars(rs0("job_time")))%>"
			
			strData = ""
			strData = strData & Chr(11) & "<%=Trim(ConvSPChars(rs0("prodt_order_no")))%>"
			strData = strData & Chr(11) & "<%=Trim(ConvSPChars(rs0("tag_no")))%>"
			strData = strData & Chr(11) & "<%=Trim(ConvSPChars(rs0("result_if_seq")))%>"
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("report_dt")) %>" 
			strData = strData & Chr(11) & "<%=Trim(ConvSPChars(rs0("item_cd")))%>"
			strData = strData & Chr(11) & "<%=Trim(ConvSPChars(rs0("item_nm")))%>"
			strData = strData & Chr(11) & "<%=Trim(ConvSPChars(rs0("lot_no")))%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("prod_qty_in_order_unit"),ggQty.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=Trim(ConvSPChars(rs0("job_status")))%>"	
			if strTime = "" Then
				strData = strData & Chr(11) & ""
			Else	
				strData = strData & Chr(11) & left(strTime,2) & ":" & right(strTime,2)
			End IF	
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Job_line")) %>" 
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("create_type")) %>" 				
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("result_seq")) %>" 
			strData = strData & Chr(11) & "<%=Trim((rs0("send_dt")))%>"
			strData = strData & Chr(11) & "<%=Trim(ConvSPChars(rs0("erp_receive_dt")))%>"	
			strData = strData & Chr(11) & "<%=Trim(ConvSPChars(rs0("erp_apply_flag")))%>"
			strData = strData & Chr(11) & "<%=Trim(ConvSPChars(rs0("err_desc")))%>"
			strData = strData & Chr(11) & "<%=Trim(ConvSPChars(rs0("POP_INF_KEY")))%>"
			strData = strData & Chr(11) & "<%=Trim(ConvSPChars(rs0("plant_cd")))%>"					
			
			strData = strData & Chr(11) & LngMaxRow + <%=LngRow%>
			strData = strData & Chr(11) & Chr(12)	
			
			TmpBuffer(<%=LngRow%>) = strData
<%			
		Else
%>
			.lgStrPrevKey1 = "<%=Trim(ConvSPChars(rs0("plant_cd")))%>"
			.lgStrPrevKey2 = "<%=Trim(ConvSPChars(rs0("prodt_order_no")))%>"
			.lgStrPrevKey3 = "<%=Trim(UNIDateClientFormat(rs0("report_dt")))%>"
			.lgStrPrevKey4 = "<%=Trim(ConvSPChars(rs0("create_type")))%>"
			.lgStrPrevKey5 = "<%=Trim(ConvSPChars(rs0("tag_no")))%>"
			.lgStrPrevKey6 = "<%=Trim(ConvSPChars(rs0("result_if_seq")))%>"				
<%		
		End If
		LngRow = LngRow + 1
		rs0.MoveNext
	Loop
%>
		iTotalStr = Join(TmpBuffer, "")
		.ggoSpread.Source = .frm1.vspdData
		.ggoSpread.SSShowDataByClip iTotalStr
		
		.frm1.hPlantCd.value = "<%=ConvSPChars(Request("txtPlantCd"))%>"
		.frm1.hSendStartDt.value  = "<%=Request("txtSendStartDt")%>"
		.frm1.hSendEndDt.value	 = "<%=Request("txtSendEndDt")%>"
		.frm1.hPlanStartDt.value  = "<%=Request("txtPlanStartDt")%>"
		.frm1.hPlanEndDt.value	 = "<%=Request("txtPlanEndDt")%>"
		.frm1.hItemCd.value = "<%=ConvSPChars(Request("txtItemCd"))%>"
		.frm1.hProdOrderNo.value = "<%=ConvSPChars(Request("txtProdtOrderNo"))%>"
		.frm1.hRdoFlag.value = "<%=ConvSPChars(UCase(Trim(Request("rdoFlag"))))%>"


<%			
		rs0.Close
		Set rs0 = Nothing
%>
	.DbQueryOk(LngMaxRow+1)
End With
</Script>	
<%
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>


