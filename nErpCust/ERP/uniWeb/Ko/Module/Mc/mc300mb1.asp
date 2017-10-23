<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : MC300MB1
'*  4. Program Name         : 납입지시확정 
'*  5. Program Desc         : 납입지시확정 
'*  6. Component List       : 
'*  7. Modified date(First) : 2003-02-25
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Ryu Sung Won
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "M", "NOCOOKIE","MB")

Dim ADF												'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg										'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag			'DBAgent Parameter 선언 
Dim rs0, rs1, rs2, rs3, rs4, rs5					'DBAgent Parameter 선언 
Dim strQryMode

Dim strPlantCd
Dim strFromReqDt
Dim strToReqDt
Dim strItemCd
Dim strBpCd
Dim strProdtOrderNo
Dim strPoNo
Dim strWcCd
Dim strTrackingNo
	
Const C_SHEETMAXROWS = 50

Dim strMode											'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================
Dim StrNextKey
Dim strPrevKey_ProdOrderNo
Dim strPrevKey_OprNo
Dim strPrevKey_Seq
Dim strPrevKey_SubSeq
Dim LngMaxRow				' 현재 그리드의 최대Row
Dim GroupCount
Dim i

Call HideStatusWnd

On Error Resume Next
Err.Clear

	strMode = Request("txtMode")						'☜ : 현재 상태를 받음 
	strQryMode = Request("lgIntFlgMode")
	strPrevKey_ProdOrderNo = Request("lgStrPrevKey1")
	strPrevKey_OprNo = Request("lgStrPrevKey2")
	strPrevKey_Seq = Request("lgStrPrevKey3")
	strPrevKey_SubSeq = Request("lgStrPrevKey4")
	
	strPlantCd = Request("txtPlantCd")
	strFromReqDt = Request("txtFromReqDt")
	strToReqDt = Request("txtToReqDt")
	strItemCd = Request("txtItemCd")
	strBpCd = Request("txtBpCd")
	strProdtOrderNo = Request("txtProdOrderNo")
	strPoNo = Request("txtPoNo")
	strWcCd = Request("txtWcCd")
	strTrackingNo = Request("txtTrackingNo")

	IF Trim(strPlantCd) = "" Then
	   Call DisplayMsgBox("900010", vbOKOnly, "", "", I_MKSCRIPT)
	   Response.End
	ELSE
	   strPlantCd = FilterVar(UCase(strPlantCd), "''", "S")
	END IF
	
	IF Trim(strFromReqDt) = "" Then
	   Call DisplayMsgBox("900010", vbOKOnly, "", "", I_MKSCRIPT)
	   Response.End
	ELSE
	   strFromReqDt = FilterVar(UniConVDate(strFromReqDt), "''", "S")
	END IF
	
	IF Trim(strToReqDt) = "" Then
	   Call DisplayMsgBox("900010", vbOKOnly, "", "", I_MKSCRIPT)
	   Response.End
	ELSE
	   strToReqDt = FilterVar(UniConVDate(strToReqDt), "''", "S")
	END IF
	
	IF Trim(strItemCd) = "" Then
	   strItemCd = "|"
	ELSE
	   strItemCd = FilterVar(UCase(strItemCd), "''", "S")
	END IF

	IF Trim(strBpCd) = "" Then
	   strBpCd = "|"
	ELSE
	   strBpCd = FilterVar(UCase(strBpCd), "''", "S")
	END IF
	
	IF Trim(strProdtOrderNo) = "" Then
	   strProdtOrderNo = "|"
	ELSE
	   strProdtOrderNo = FilterVar(UCase(strProdtOrderNo), "''", "S")
	END IF
	
	IF Trim(strPoNo) = "" Then
	   strPoNo = "|"
	ELSE
	   strPoNo = FilterVar(UCase(strPoNo), "''", "S")
	END IF
	
	IF Trim(strWcCd) = "" Then
	   strWcCd = "|"
	ELSE
	   strWcCd = FilterVar(UCase(strWcCd), "''", "S")
	END IF
	
	IF Trim(strTrackingNo) = "" Then
	   strTrackingNo = "|"
	ELSE
	   strTrackingNo = FilterVar(UCase(strTrackingNo), "''", "S")
	END IF
	
	If Cint(strQryMode) = Cint(OPMD_CMODE) Then
		'=======================================================================================================
		'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
		'=======================================================================================================
		Redim UNISqlId(4)
		Redim UNIValue(4, 0)

		UNISqlId(0) = "180000saa"	'plant_cd
		UNISqlId(1) = "180000sak"	'BpCd
		UNISqlId(2) = "mc300mb101"	'PoNo
		UNISqlId(3) = "180000sac"	'WcCd
		UNISqlId(4) = "180000sam"	'tracking_no

		UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
		UNIValue(1, 0) = FilterVar(UCase(Request("txtBpCd")), "''", "S")
		UNIValue(2, 0) = FilterVar(UCase(Request("txtPoNo")), "''", "S")
		UNIValue(3, 0) = FilterVar(UCase(Request("txtWcCd")), "''", "S")
		UNIValue(4, 0) = FilterVar(UCase(Request("txtTrackingNo")), "''", "S")

		UNILock = DISCONNREAD :	UNIFlag = "1"
	
		Set ADF = Server.CreateObject("prjPublic.cCtlTake")
		strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1, rs2, rs3, rs4, rs5)

   		' Plant 명 Display      
		%>
		<Script Language=vbscript>
			parent.frm1.txtPlantNm.value = ""
			parent.frm1.txtItemNm.value = ""
		</Script>	
		<%    	
		' Plant 명 Display      
		If (rs1.EOF And rs1.BOF) Then
			Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
			rs1.Close
			Set rs1 = Nothing
			Set ADF = Nothing
			%>
			<Script Language=vbscript>
				parent.frm1.txtPlantNm.value = ""
				parent.frm1.txtPlantCd.Focus()
			</Script>	
			<%
			Response.End
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
		IF Request("txtBpCd") <> "" Then
			If (rs2.EOF And rs2.BOF) Then
				Call DisplayMsgBox("229927", vbOKOnly, "", "", I_MKSCRIPT)
				rs2.Close
				Set rs2 = Nothing
				Set ADF = Nothing
				%>
				<Script Language=vbscript>
					parent.frm1.txtBpNm.value = ""
					parent.frm1.txtBpCd.Focus()
				</Script>	
				<%
				Response.End
			Else
				%>
				<Script Language=vbscript>
					parent.frm1.txtBpNm.value = "<%=ConvSPChars(rs2("BP_NM"))%>"
				</Script>	
				<%
				rs2.Close
				Set rs2 = Nothing
			End If
		End IF
	
		' 발주번호 
		IF Request("txtPoNo") <> "" Then
			If (rs3.EOF And rs3.BOF) Then
				Call DisplayMsgBox("173132", vbOKOnly, "발주번호", "", I_MKSCRIPT)
				rs3.Close
				Set rs3 = Nothing
				Set ADF = Nothing
				%>
				<Script Language=vbscript>
					parent.frm1.txtPoNo.Focus()
				</Script>	
				<%
				Response.End
			Else
				%>
				<Script Language=vbscript>
					parent.frm1.txtPoNo.value = "<%=ConvSPChars(rs3("PO_NO"))%>"
				</Script>	
				<%
				rs3.Close
				Set rs3 = Nothing
			End If
		End IF
	
		'Shift_Cd
		IF Request("txtWcCd") <> "" Then
			If (rs4.EOF And rs4.BOF) Then
				Call DisplayMsgBox("182100", vbOKOnly, "", "", I_MKSCRIPT)
				rs4.Close
				Set rs4 = Nothing
				Set ADF = Nothing
				%>
				<Script Language=vbscript>
					parent.frm1.txtWcNm.value = ""
					parent.frm1.txtWcCd.Focus()
				</Script>	
				<%
				Response.End
			Else
				%>
				<Script Language=vbscript>
					parent.frm1.txtWcNm.value = "<%=ConvSPChars(rs4("WC_NM"))%>"
				</Script>	
				<%
				rs4.Close
				Set rs4 = Nothing
			End If
		End IF

		'Tracking_No
		IF Request("txtTrackingNo") <> "" Then
			If (rs5.EOF And rs5.BOF) Then
				Call DisplayMsgBox("203045", vbOKOnly, "", "", I_MKSCRIPT)
				rs5.Close
				Set rs5 = Nothing
				Set ADF = Nothing
				%>
				<Script Language=vbscript>
					parent.frm1.txtTrackingNo.Focus()
				</Script>	
				<%
				Response.End
			Else
				rs5.Close
				Set rs5 = Nothing
			End If
		End IF
	End If
	
	Redim UNISqlId(0)
	Redim UNIValue(0, 10)

	UNISqlId(0) = "mc300mb102"
	
	UNIValue(0, 0) = "^"		'Allocated
	UNIValue(0, 1) = "" & FilterVar("AL", "''", "S") & ""		'Allocated
	UNIValue(0, 2) = strPlantCd
	UNIValue(0, 3) = strFromReqDt
	UNIValue(0, 4) = strToReqDt 	
	UNIValue(0, 5) = strItemCd
	UNIValue(0, 6) = strBpCd
	UNIValue(0, 7) = strPoNo
	UNIValue(0, 8) = strWcCd
	UNIValue(0, 9) = strTrackingNo
	
	If Cint(strQryMode) = Cint(OPMD_UMODE) Then
		If Trim(strPrevKey_ProdOrderNo) <> "" Then
			UNIValue(0, 10) = "(a.prodt_order_no > " & FilterVar(UCase(strPrevKey_ProdOrderNo), "''", "S") & " OR " & _
				"(a.prodt_order_no = " & FilterVar(UCase(strPrevKey_ProdOrderNo), "''", "S") & " and " & _
				" b.opr_no = " & FilterVar(UCase(strPrevKey_OprNo), "''", "S") & ") OR " & _
				"(a.prodt_order_no = " & FilterVar(UCase(strPrevKey_ProdOrderNo), "''", "S") & " and " & _
				" b.opr_no = " & FilterVar(UCase(strPrevKey_OprNo), "''", "S") & " and " & _
				" b.seq = " & FilterVar(UCase(strPrevKey_Seq), "''", "S") & ") OR " & _
				"(a.prodt_order_no = " & FilterVar(UCase(strPrevKey_ProdOrderNo), "''", "S") & " and " & _
				" b.opr_no = " & FilterVar(UCase(strPrevKey_OprNo), "''", "S") & " and " & _
				" b.seq = " & FilterVar(UCase(strPrevKey_Seq), "''", "S") & " and " & _
				" b.subseq = " & FilterVar(UCase(strPrevKey_SubSeq), "''", "S") & "))"
		End If
	Else
		If strProdtOrderNo = "|" Then
			UNIValue(0, 10) = "|"
		Else
			UNIValue(0, 10) = "a.prodt_order_no >= " & strProdtOrderNo
		End If
	End If

	UNILock = DISCONNREAD :	UNIFlag = "1"

    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

	If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT) 	
		rs0.Close
		Set rs0 = Nothing					
		Response.Write "<Script Language=vbscript>"	& vbCr
		Response.Write "parent.frm1.txtPlantCd.focus"	& vbCr
		Response.Write "</Script>" & vbCr
		Response.End													'☜: 비지니스 로직 처리를 종료함 
	End If
%>

<Script Language=vbscript>
Dim LngLastRow
Dim LngMaxRow 
Dim strTemp
Dim strData
Dim PvArr

With parent																'☜: 화면 처리 ASP 를 지칭함 
	LngMaxRow = .frm1.vspdData.MaxRows									'Save previous Maxrow
	ReDim PvArr(<%=C_SHEETMAXROWS%> - 1)
		
<%  		
    For i=0 to rs0.RecordCount-1 
		If i < C_SHEETMAXROWS Then
%>
		strData = strData & Chr(11) & ""
		strData = strData & Chr(11) & "<%=ConvSPChars(UCase(Trim(rs0("PRODT_ORDER_NO"))))%>"	
		strData = strData & Chr(11) & "<%=ConvSPChars(UCase(Trim(rs0("ITEM_CD"))))%>"	
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEM_NM"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(UCase(Trim(rs0("SPEC"))))%>"
		strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("REQ_DT"))%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("REQ_QTY"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(UCase(Trim(rs0("BASE_UNIT"))))%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("DO_QTY"),ggQty.DecPoint,0)%>"
		'strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("RCPT_QTY"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(UCase(Trim(rs0("BP_CD"))))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("BP_NM"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(UCase(Trim(rs0("PO_NO"))))%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("PO_SEQ_NO"),0,0)%>"	'정수부만 표시 
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("DO_QTY_PO_UNIT"),ggQty.DecPoint,0)%>"
		'strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("RCPT_QTY_PO_UNIT"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(UCase(Trim(rs0("PO_UNIT"))))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(UCase(Trim(rs0("OPR_NO"))))%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("SEQ"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("SUB_SEQ"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(UCase(Trim(rs0("WC_CD"))))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("WC_NM"))%>"
		strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("PLAN_START_DT"))%>"
		strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("PLAN_COMPT_DT"))%>"
		strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("RELEASE_DT"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("TRACKING_NO"))%>"
		strData = strData & Chr(11) & LngMaxRow + <%=i%>
		strData = strData & Chr(11) & Chr(12)

        PvArr(<%=i%>) = strData	
		strData = ""
<%		
		rs0.MoveNext
		
		End If
	Next
%>
		strData = join(PvArr,"")
		.ggoSpread.Source = .frm1.vspdData
		.ggoSpread.SSShowData strData
		
		.lgStrPrevKey1 = "<%=Trim(rs0("PRODT_ORDER_NO"))%>"
		.lgStrPrevKey2 = "<%=Trim(rs0("OPR_NO"))%>"
		.lgStrPrevKey3 = "<%=Trim(rs0("SEQ"))%>"
		.lgStrPrevKey4 = "<%=Trim(rs0("SUB_SEQ"))%>"
		
<%		
		rs0.Close
		Set rs0 = Nothing
%>
	If .frm1.vspdData.MaxRows < .VisibleRowCnt(.frm1.vspdData,0) and .lgStrPrevKey1 <> "" Then	<% ' GroupView 사이즈로 화면 Row수보다 쿼리가 작으면 다시 쿼리함 %>
		.DbQuery
	Else
		.frm1.hPlantCd.value		= "<%=ConvSPChars(Request("txtPlantCd"))%>"
		.frm1.hFromReqDt.value		= "<%=UNIDateClientFormat(Request("txtFromReqDt"))%>"
		.frm1.hToReqDt.value		= "<%=UNIDateClientFormat(Request("txtToReqDt"))%>"
		.frm1.hItemCd.value			= "<%=ConvSPChars(Request("txtItemCd"))%>"
		.frm1.hBpCd.value			= "<%=ConvSPChars(Request("txtBpCd"))%>"
		.frm1.hProdOrderNo.value	= "<%=ConvSPChars(Request("txtProdOrderNo"))%>"
		.frm1.hPoNo.value			= "<%=ConvSPChars(Request("txtPoNo"))%>"
		.frm1.hWcCd.value			= "<%=ConvSPChars(Request("txtWcCd"))%>"
		.frm1.hTrackingNo.value		= "<%=ConvSPChars(Request("txtTrackingNo"))%>"
		.DbQueryOk	'(LngMaxRow+1)
	End If
End With	
</Script>	
<%
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
