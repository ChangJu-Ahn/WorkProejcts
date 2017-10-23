<%
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: p4311mb1.asp
'*  4. Program Name			: List Component Requirement (Reservation) (Query)
'*  5. Program Desc			: Used By Goods Issue For Production Order
'*  6. Comproxy List		: ADO 180000saa, 180000sab, 180000sad
'*						    :
'*  7. Modified date(First)	: 2000/03/28
'*  8. Modified date(Last)	: 2003/02/25
'*  9. Modifier (First)		: Park, BumSoo
'* 10. Modifier (Last)		: Chen, Jae Hyun
'* 11. Comment				:
'**********************************************************************************************
%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("I", "*", "NOCOOKIE", "MB")

On Error Resume Next								'☜: 

Dim ADF												'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg										'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag			'DBAgent Parameter 선언 
Dim rs0, rs1, rs2, rs3, rs4, rs5, rs6, rs7			'DBAgent Parameter 선언 
Dim strQryMode

Const C_SHEETMAXROWS_D = 100

'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================
Dim lgStrPrevKey									' 이전 값 
Dim LngRow

Dim strMode											'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim strFlag


Call HideStatusWnd

strMode = Request("txtMode")						'☜ : 현재 상태를 받음 
strQryMode = Request("lgIntFlgMode")

On Error Resume Next

	lgStrPrevKey = UCase(Trim(Request("lgStrPrevKey")))	
	lgStrPrevKey1 = UCase(Trim(Request("lgStrPrevKey1")))
	
	
	Err.Clear										'☜: Protect system from crashing
	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================
	Redim UNISqlId(6)
	Redim UNIValue(6, 0)

	UNISqlId(0) = "180000saa"
	UNISqlId(1) = "180000sab"
	UNISqlId(2) = "180000sad"
	UNISqlId(3) = "180000sam"
	UNISqlId(4) = "180000sac"
	UNISqlId(5) = "180000sas"
	UNISqlId(6) = "180000sab"

	
	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(1, 0) = FilterVar(UCase(Request("txtChildItemCd")), "''", "S")
	UNIValue(2, 0) = FilterVar(UCase(Request("txtSLCd")), "''", "S")
	UNIValue(3, 0) = FilterVar(UCase(Request("txtTrackingNo")), "''", "S")
	UNIValue(4, 0) = FilterVar(UCase(Request("txtWcCd")), "''", "S")
	UNIValue(5, 0) = FilterVar(UCase(Request("txtItemGroupCd")), "''", "S")
	UNIValue(6, 0) = FilterVar(UCase(Request("txtItemCd")), "''", "S")

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1, rs2, rs3 ,rs4, rs5, rs6, rs7)
	
	%>
		<Script Language=vbscript>
			parent.frm1.txtPlantNm.value = ""
			parent.frm1.txtChildItemNm.value = ""
			parent.frm1.txtSLNm.value = ""
			parent.frm1.txtWcNm.value = ""
			parent.frm1.txtItemGroupNm.value = ""
			parent.frm1.txtItemNm.value = ""
		</Script>	
	<%
	
	'Item Group Display
	IF Request("txtItemGroupCd") <> "" Then
		If (rs6.EOF And rs6.BOF) Then
			
			strFlag = "ERROR_GROUP"
		Else
			%>
			<Script Language=vbscript>
				parent.frm1.txtItemGroupNm.value = "<%=ConvSPChars(rs6("item_group_nm"))%>"
			</Script>	
			<%
		End If	
	End IF
	rs6.Close
	Set rs6 = Nothing
	
	'Work Center Display
	IF Request("txtWcCd") <> "" Then
		If (rs5.EOF And rs5.BOF) Then
			
			strFlag = "ERROR_WCCD"
		Else
			%>
			<Script Language=vbscript>
				parent.frm1.txtWCNm.value = "<%=ConvSPChars(rs5("wc_nm"))%>"
			</Script>	
			<%
		End If	
	End IF
	rs5.Close
	Set rs5 = Nothing
	
	'Item Code Display
	IF Request("txtItemCd") <> "" Then
		If (rs7.EOF And rs7.BOF) Then
			strFlag = "ERROR_ITEM"
		Else
			%>
			<Script Language=vbscript>
				parent.frm1.txtItemNm.value = "<%=ConvSPChars(rs7("item_nm"))%>"
			</Script>	
			<%
		End If	
	End IF
	rs7.Close
	Set rs7 = Nothing
	
	'Tracking Display	
	IF Request("txtTrackingNo") <> "" Then
		If (rs4.EOF And rs4.BOF) Then
			strFlag = "ERROR_TRACK"
		End If	
	End IF
	rs4.Close
	Set rs4 = Nothing
			
	' 창고명 Display
	IF Request("txtSLCd") <> "" Then
		If (rs3.EOF And rs3.BOF) Then
			strFlag = "ERROR_SLCD"
		Else
			%>
			<Script Language=vbscript>
				parent.frm1.txtSLNm.value = "<%=ConvSPChars(rs3("SL_NM"))%>"
			</Script>	
			<%
		End If	
	End IF
	rs3.Close
	Set rs3 = Nothing
	
	' 품목명 Display
	IF Request("txtChildItemCd") <> "" Then
		If (rs2.EOF And rs2.BOF) Then
			strFlag = "ERROR_CHILD_ITEM"
		Else
			%>
			<Script Language=vbscript>
				parent.frm1.txtChildItemNm.value = "<%=ConvSPChars(rs2("ITEM_NM"))%>"
			</Script>	
			<%
		End If
	End IF
	rs2.Close
	Set rs2 = Nothing
	
	' Plant 명 Display      
	If (rs1.EOF And rs1.BOF) Then
		strFlag = "ERROR_PLANT"
	Else
		%>
		<Script Language=vbscript>
			parent.frm1.txtPlantNm.value = "<%=ConvSPChars(rs1("PLANT_NM"))%>"
		</Script>	
		<%    	
	End If
	rs1.Close
	Set rs1 = Nothing

	If strFlag <> "" Then
		If strFlag = "ERROR_PLANT" Then
			Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
			parent.frm1.txtPlantCd.Focus()
			</Script>	
			<%
			Response.End	
		ElseIf strFlag = "ERROR_CHILD_ITEM" Then
		   Call DisplayMsgBox("122600", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
			parent.frm1.txtChildItemCd.Focus()
			</Script>	
			<%
			Response.End
		ElseIf strFlag = "ERROR_ITEM" Then
			Call DisplayMsgBox("122600", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
			If parent.document.all("DetailCondition1").style.display = "none" Then
				parent.document.all("DetailCondition1").style.display = ""
				parent.document.all("DetailCondition2").style.display = ""
			End If
			parent.frm1.txtItemCd.Focus()
			</Script>	
			<%
			Response.End
		ElseIf strFlag = "ERROR_GROUP" Then
			Call DisplayMsgBox("127400", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
			parent.frm1.txtIteGroupCd.Focus()
			</Script>	
			<%
			Response.End	
		ElseIf strFlag = "ERROR_SLCD" Then
			Call DisplayMsgBox("125700", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
			parent.frm1.txtSLCd.Focus()
			</Script>	
			<%
			Response.End
		ElseIf strFlag = "ERROR_WCCD" Then
			Call DisplayMsgBox("182100", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
			parent.frm1.txtWCCd.Focus()
			</Script>	
			<%
			Response.End	
		ElseIf strFlag = "ERROR_TRACK" Then
			Call DisplayMsgBox("203045", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
			If parent.document.all("DetailCondition1").style.display = "none" Then
				parent.document.all("DetailCondition1").style.display = ""
				parent.document.all("DetailCondition2").style.display = ""
			End If
			parent.frm1.txtTrackingNo.Focus()
			</Script>	
			<%
			Set ADF = Nothing
			Response.End	
		End If
	End IF
 
    
	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=====================================================================================================
	
	Redim UNISqlId(0)
	Redim UNIValue(0, 17)
	
	UNISqlId(0) = "p4311mb1"
	
	
	UNIValue(0, 0) = FilterVar(UCase(Trim(Request("txtPlantCd"))),"''","S")
	UNIValue(0, 1) = FilterVar(UniConvDate(Request("txtReqStartDt")),"''","S")
	UNIValue(0, 2) = FilterVar(UniConvDate(Request("txtReqEndDt")),"''","S")
	IF Request("txtChildItemCd") = "" Then
		UNIValue(0, 3) = FilterVar("%", "''", "S")
	Else
	UNIValue(0, 3) = FilterVar(UCase(Request("txtChildItemCd")),"''","S")
	End IF 
	
	IF Request("txtSLCd") = "" Then
		UNIValue(0, 4) = FilterVar("%", "''", "S")
	Else
	UNIValue(0, 4) = FilterVar(UCase(Request("txtSLCd")),"''","S")
	End IF
	IF Request("txtTrackingNo") = "" Then
		UNIValue(0, 5) =  FilterVar("%", "''", "S")
	Else
	UNIValue(0, 5) = FilterVar(UCase(Request("txtTrackingNo")),"''","S")
	End IF 
	
	IF Trim(Request("cboProdMgr")) = "" Then
		UNIValue(0, 6) = FilterVar("%", "''", "S") & " or isNULL(i.prod_mgr,'') = '' "
	Else
		UNIValue(0, 6) = FilterVar(UCase(Request("cboProdMgr")), "''", "S")
	End IF 
	
	IF Trim(Request("cboInvMgr")) = "" Then
		UNIValue(0, 7) = FilterVar("%", "''", "S") & " or isNULL(h.inv_mgr,'') = '' "
	Else
		UNIValue(0, 7) = FilterVar(UCase(Request("cboInvMgr")), "''", "S")
	End IF 
	
	IF Trim(Request("txtWcCd")) = "" Then
		UNIValue(0, 8) = FilterVar("%", "''", "S")
	Else
		UNIValue(0, 8) = FilterVar(UCase(Request("txtWcCd")), "''", "S")
	End IF 
	
	IF Trim(Request("cboJobCd")) = "" Then
		UNIValue(0, 9) = FilterVar("%", "''", "S") & " or isNULL(c.job_cd,'') = '' "
	Else
		UNIValue(0, 9) = FilterVar(UCase(Request("cboJobCd")), "''", "S")
	End IF 
	
	IF Trim(Request("txtItemCd")) = "" Then
		UNIValue(0, 10) = FilterVar("%", "''", "S")
	Else
		UNIValue(0, 10) = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	End IF 
	
	IF Trim(Request("txtItemGroupCd")) = "" Then
		UNIValue(0, 11) = "" 
	Else
		UNIValue(0, 11) = " and d.item_group_cd in (select item_group_cd from ufn_P_ListItemGrp(" & FilterVar(Trim(Request("txtItemGroupCd"))	, "''", "S") & " )) "
	End IF 
	
	Select Case strQryMode
		Case CStr(OPMD_CMODE)
	
			UNIValue(0, 12) = FilterVar(UCase(Request("txtProdtOrderNo")),"''","S")
			
			UNIValue(0, 13) = FilterVar(UCase(Request("txtProdtOrderNo")),"''","S")
			
			UNIValue(0, 14) = "''"
			
			UNIValue(0, 15) = "''"
			
			UNIValue(0, 16) = "''"
			
			UNIValue(0, 17) = "''"
			
		Case CStr(OPMD_UMODE)

			UNIValue(0, 12) = FilterVar(UCase(Request("lgStrPrevKey")),"''","S")
			UNIValue(0, 13) = FilterVar(UCase(Request("lgStrPrevKey")),"''","S")
			UNIValue(0, 14) = FilterVar(UCase(Request("lgStrPrevKey1")),"''","S")
			UNIValue(0, 15) = FilterVar(UCase(Request("lgStrPrevKey")),"''","S")
			UNIValue(0, 16) = FilterVar(UCase(Request("lgStrPrevKey1")),"''","S")
			UNIValue(0, 17) = FilterVar(UCase(Request("lgStrPrevKey2")),"''","S")
			
	End Select
		
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
       
	If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		
		rs0.Close
		Set rs0 = Nothing
		Response.End														'☜: 비지니스 로직 처리를 종료함 
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
	If C_SHEETMAXROWS_D < rs0.RecordCount Then 
%>
		ReDim TmpBuffer(<%=C_SHEETMAXROWS_D - 1%>)
<%
	Else
%>			
		ReDim TmpBuffer(<%=rs0.RecordCount - 1%>)
<%
	End If
	
    For LngRow=0 to rs0.RecordCount-1 
		If LngRow < C_SHEETMAXROWS_D Then
%>
			strData = ""
			strData = strData & Chr(11) & "<%=Trim(ConvSPChars(rs0("prodt_order_no")))%>"
			strData = strData & Chr(11) & "<%=Trim(ConvSPChars(rs0("child_item_cd")))%>"
			strData = strData & Chr(11) & "<%=Trim(ConvSPChars(rs0("child_item_nm")))%>"
			strData = strData & Chr(11) & "<%=Trim(ConvSPChars(rs0("spec")))%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("req_qty"),ggQty.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=Trim(ConvSPChars(rs0("base_unit")))%>"		 
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("Issued_qty"),ggQty.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("remain_qty"),ggQty.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("issue_qty"),ggQty.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("req_dt")) %>" 
			strData = strData & Chr(11) & "<%=Trim(ConvSPChars(rs0("resv_status")))%>"
			strData = strData & Chr(11) & "<%=Trim(ConvSPChars(rs0("resv_desc")))%>"
			strData = strData & Chr(11) & "<%=Trim(ConvSPChars(rs0("sl_cd")))%>"	
			strData = strData & Chr(11) & "<%=Trim(ConvSPChars(rs0("sl_nm")))%>"	
			strData = strData & Chr(11) & "<%=Trim(ConvSPChars(rs0("tracking_no")))%>"
			strData = strData & Chr(11) & "<%=Trim(ConvSPChars(rs0("opr_no")))%>"
			strData = strData & Chr(11) & "<%=Trim(ConvSPChars(rs0("wc_cd")))%>"
			strData = strData & Chr(11) & "<%=Trim(ConvSPChars(rs0("seq")))%>"
			strData = strData & Chr(11) & "<%=Trim(ConvSPChars(rs0("req_no")))%>"
			strData = strData & Chr(11) & "<%=Trim(ConvSPChars(rs0("item_cd")))%>"		
			strData = strData & Chr(11) & "<%=Trim(ConvSPChars(rs0("item_nm")))%>"
			strData = strData & Chr(11) & "<%=Trim(ConvSPChars(rs0("parent_item_spec")))%>"
			strData = strData & Chr(11) & "<%=Trim(ConvSPChars(rs0("job_nm")))%>"
			strData = strData & Chr(11) & LngMaxRow + <%=LngRow%>
			strData = strData & Chr(11) & Chr(12)
			
			TmpBuffer(<%=LngRow%>) = strData
<%		
			rs0.MoveNext
		End If
	Next
%>
		iTotalStr = Join(TmpBuffer, "")
		.ggoSpread.Source = .frm1.vspdData1 
		.ggoSpread.SSShowDataByClip iTotalStr
		
		.lgStrPrevKey = "<%=Trim(ConvSPChars(rs0("Prodt_order_no")))%>"
		.lgStrPrevKey1 = "<%=Trim(ConvSPChars(rs0("opr_no")))%>"
		.lgStrPrevKey2 = "<%=Trim(ConvSPChars(rs0("seq")))%>"

		.frm1.hPlantCd.value	 = "<%=ConvSPChars(Request("txtPlantCd"))%>"
		.frm1.hReqStartDt.value  = "<%=Request("txtReqStartDt")%>"
		.frm1.hReqEndDt.value	 = "<%=Request("txtReqEndDt")%>"
		.frm1.hChildItemCd.value = "<%=ConvSPChars(Request("txtChildItemCd"))%>"
		.frm1.hProdOrderNo.value = "<%=ConvSPChars(Request("txtProdtOrderNo"))%>"
		.frm1.hSLCd.value		 = "<%=ConvSPChars(Request("txtSLCd"))%>"
		.frm1.hTrackingNo.value  = "<%=ConvSPChars(Request("txtTrackingNo"))%>"
		.frm1.hProdMgr.value  = "<%=ConvSPChars(Request("cboProdMgr"))%>"
		.frm1.hInvMgr.value  = "<%=ConvSPChars(Request("cboInvMgr"))%>"
		.frm1.hWcCd.value = "<%=ConvSPChars(Request("txtWcCd"))%>"
		.frm1.hItemCd.value = "<%=ConvSPChars(Request("txtItemCd"))%>"
		.frm1.hItemGroupCd.value = "<%=ConvSPChars(Request("txtItemGroupCd"))%>"
		.frm1.hJobCd.value = "<%=ConvSPChars(Request("cboJobCd"))%>"
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
