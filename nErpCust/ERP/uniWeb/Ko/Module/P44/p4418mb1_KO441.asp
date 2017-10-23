<%
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: p4417mb1.asp
'*  4. Program Name			: List Production Results (Query)
'*  5. Program Desc			: Called By Confirm By Operation and Confirm By Order
'*  6. Comproxy List		: DB Agent (p4417mb1)
'*  7. Modified date(First)	: 2000/03/28
'*  8. Modified date(Last)	: 2002/06/26
'*  9. Modifier (First)		: Park, BumSoo
'* 10. Modifier (Last)		: Park, BumSoo
'* 11. Comment				: COOL:GEN -> DB Agent
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
Call LoadInfTB19029B("Q", "P", "NOCOOKIE","MB")

On Error Resume Next

Dim ADF										'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg								'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter 선언 
Dim rs0, rs1, rs2, rs3							'DBAgent Parameter 선언 
Dim strMode									'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================
Dim i
Dim strSQL

Const C_SHEETMAXROWS_D = 50

Call HideStatusWnd

strMode = Request("txtMode")												'☜ : 현재 상태를 받음 
strQryMode = Request("lgIntFlgMode")
lgStrPrevKey = Request("lgStrPrevKey")

On Error Resume Next

Err.Clear																	'☜: Protect system from crashing

	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================
	strSQL = " A.OPR_NO <= " &  FilterVar(UCase(Request("txtOprCd")), "''", "S")
	strSQL = strSQL & " AND A.OPR_NO >= (SELECT ISNULL(MAX(OPR_NO), '') FROM P_PRODUCTION_ORDER_DETAIL "
	strSQL = strSQL & " WHERE MILESTONE_FLG = 'Y' "
	strSQL = strSQL & "  AND PRODT_ORDER_NO = "  & FilterVar(UCase(Request("txtProdOrderNo")), "''", "S")
	strSQL = strSQL & "  AND OPR_NO < " &  FilterVar(UCase(Request("txtOprCd")), "''", "S") & ")"
	
	Redim UNISqlId(1)
	Redim UNIValue(1, 2)

	UNISqlId(0) = "180000saa"
		'Add 2005-10-05
	UNISqlId(1) = "P4419MB1S"

	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
		'Add 2005-10-05
	UNIValue(1, 0) = "^"
	UNIValue(1, 1) = FilterVar(UCase(Request("txtProdOrderNo")), "''", "S")
	UNIValue(1, 2) = strSQL
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs3)

	' Plant 명 Display      
	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
		%>
		<Script Language=vbscript>
		parent.frm1.txtPlantNm.value = ""	
		parent.frm1.txtPlantCd.Focus()
		</Script>	
		<%
		rs0.Close
		Set rs0 = Nothing
		Set ADF = Nothing
		Response.End
	Else
		%>
		<Script Language=vbscript>
			parent.frm1.txtPlantNm.value = "<%=ConvSPChars(rs0("PLANT_NM"))%>"
		</Script>	
		<%
		rs0.Close
		Set rs0 = Nothing
	End If
	
	' Add 2005-10-05 Check Damper Components      
	If (rs3.EOF And rs3.BOF) Then
		%>
		<Script Language=vbscript>
		parent.frm1.hDamperFlag.value = "N"	
		</Script>	
		<%
		rs3.Close
		Set rs3 = Nothing
	Else
		If CLng(rs3("DamperCount")) > 0 Then
			%>
			<Script Language=vbscript>
				parent.frm1.hDamperFlag.value = "Y"
			</Script>	
			<%
		Else
			%>
			<Script Language=vbscript>
				parent.frm1.hDamperFlag.value = "N"
			</Script>	
			<%
		End If	
		rs3.Close
		Set rs3 = Nothing
	End If
	
	' Order Header Display
	If strQryMode = CStr(OPMD_CMODE) Then

		Redim UNISqlId(0)
		Redim UNIValue(0, 2)

		UNISqlId(0) = "p4418mb1h"
	
		UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
		UNIValue(0, 1) = FilterVar(UCase(Request("txtProdOrderNo")), "''", "S")
		UNIValue(0, 2) = FilterVar(UCase(Request("txtOprCd")), "''", "S")

		UNILock = DISCONNREAD :	UNIFlag = "1"
	
		Set ADF = Server.CreateObject("prjPublic.cCtlTake")
		strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1)
			  
		If (rs1.EOF And rs1.BOF) Then
			Call DisplayMsgBox("189300", vbOKOnly, "", "", I_MKSCRIPT)
			rs1.Close
			Set rs1 = Nothing
			Set ADF = Nothing
			Response.End
		Else
			%>
			<Script Language=vbscript>
			With parent.frm1
			
				.txtItemCd.value		= "<%=ConvSPChars(rs1("Item_Cd"))%>"
				.txtItemNm.value		= "<%=ConvSPChars(rs1("Item_Nm"))%>"
				.txtOrderQty.value		= "<%=UniNumClientFormat(rs1("Prodt_Order_Qty"),ggQty.DecPoint,0)%>"
				.txtPlndStartDt.text	= "<%=UNIDateClientFormat(rs1("Plan_Start_Dt"))%>"
				.txtPlndComptDt.text	= "<%=UNIDateClientFormat(rs1("Plan_Compt_Dt"))%>"
				.txtProdQty.Value		= "<%=UniNumClientFormat(rs1("Prod_Qty_In_Order_Unit"),ggQty.DecPoint,0)%>"
				.txtInspQty.Value		= "<%=UniNumClientFormat(rs1("Good_Qty_In_Order_Unit"),ggQty.DecPoint,0)%>"
				.txtRcptQty.Value 		= "<%=UniNumClientFormat(rs1("Rcpt_Qty_In_Order_Unit"),ggQty.DecPoint,0)%>"
				.txtStatus.value		= "<%=ConvSPChars(rs1("Order_Status"))%>"
				.txtUnit.value			= "<%=ConvSPChars(rs1("Prodt_Order_Unit"))%>"

			End With
			</Script>	
			<%
			rs1.Close
			Set rs1 = Nothing
		End If

	End If
	
	' Results Display
	Redim UNISqlId(0)
	Redim UNIValue(0, 2)

	UNISqlId(0) = "p4418mb1d"
	
	UNIValue(0, 0) = FilterVar(UCase(Request("txtProdOrderNo")), "''", "S")
	UNIValue(0, 1) = FilterVar(UCase(Request("txtOprCd")), "''", "S")
	
	If Request("lgStrPrevKey") <> "" Then
		UNIValue(0, 2) = FilterVar(UCase(Request("lgStrPrevKey")), "''", "S")
	Else
		UNIValue(0, 2) = "0"
	End If

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
	Set ADF = Server.CreateObject("prjPublic.cCtlTake")
	strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs2)

	If (rs2.EOF And rs2.BOF) Then
		Call DisplayMsgBox("189600", vbOKOnly, "", "", I_MKSCRIPT)
		rs2.Close
		Set rs2 = Nothing
		Set ADF = Nothing
		Response.End
	End If
	
%>

<Script Language=vbscript>
Dim LngMaxRow
Dim strData
Dim TmpBuffer
Dim iTotalStr
    	
With parent
	LngMaxRow = .frm1.vspdData.MaxRows

<%  
	If Not(rs2.EOF And rs2.BOF) Then
		If C_SHEETMAXROWS_D < rs2.RecordCount Then 
%>
			ReDim TmpBuffer(<%=C_SHEETMAXROWS_D - 1%>)
<%
		Else
%>			
			ReDim TmpBuffer(<%=rs2.RecordCount - 1%>)
<%
		End If

		For i=0 to rs2.RecordCount-1
			If i < C_SHEETMAXROWS_D Then 
%>
				strData = ""
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs2("Report_Dt"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs2("Shift_Cd"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs2("Report_Type"))%>"
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs2("Prod_Qty_In_Order_Unit"),ggQty.DecPoint,0)%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs2("Reason_Cd"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs2("Remark"))%>"
				strData = strData & Chr(11) & "<%=rs2("Seq")%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs2("Milestone_Flg"))%>"
				strData = strData & Chr(11) & LngMaxRow + <%=i%>
				strData = strData & Chr(11) & Chr(12)
				
				TmpBuffer(<%=i%>) = strData
<%		
				rs2.MoveNext
			End If
		Next
%>
		iTotalStr = Join(TmpBuffer, "")
		.ggoSpread.Source = .frm1.vspdData
		.ggoSpread.SSShowDataByClip iTotalStr
		
		.lgStrPrevKey = "<%=Trim(rs2("Seq"))%>"
		
<%	
	End If

	rs2.Close
	Set rs2 = Nothing

%>	

	If .frm1.vspdData.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData,0) and .lgStrPrevKey <> "" Then
		.DbQuery
	Else
		.frm1.hPlantCd.value	 = "<%=ConvSPChars(Request("txtPlantCd"))%>"
		.frm1.hProdOrderNo.value = "<%=ConvSPChars(Request("txtProdOrderNo"))%>"
		.frm1.hOprCd.value		 = "<%=ConvSPChars(Request("txtOprCd"))%>"    
		.DbQueryOk
	End If

End With

</Script>	
<%
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
