<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: p4211mb1.asp
'*  4. Program Name			: List Production Order Detail (Query)
'*  5. Program Desc			:
'*  6. Comproxy List		: DB Agent (p4211mb1)
'*  7. Modified date(First)	: 2000/03/28
'*  8. Modified date(Last)	: 2000/06/25
'*  9. Modifier (First)		: Park, BumSoo
'* 10. Modifier (Last)		: Park, BumSoo
'* 11. Comment				: COOL -> DB Agent
'**********************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call LoadBasisGlobalInf
Call loadInfTB19029B("I", "*", "NOCOOKIE","MB")

Dim ADF										'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg								'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter 선언 
Dim rs0, rs1, rs2							'DBAgent Parameter 선언 
Dim lgStrPrevKey
Dim i

'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================
Const C_SHEETMAXROWS_D = 50

Call HideStatusWnd

lgStrPrevKey = Request("lgStrPrevKey")

On Error Resume Next

Err.Clear																	'☜: Protect system from crashing

	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================
	Redim UNISqlId(0)
	Redim UNIValue(0, 0)

	UNISqlId(0) = "180000saa"

	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

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

	' Order Header Display
	If strQryMode = CStr(OPMD_CMODE) Then

		Redim UNISqlId(0)
		Redim UNIValue(0, 1)

		UNISqlId(0) = "p4211mb1h"
	
		UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
		UNIValue(0, 1) = FilterVar(UCase(Request("txtProdOrdNo")), "''", "S")

		UNILock = DISCONNREAD :	UNIFlag = "1"
	
		Set ADF = Server.CreateObject("prjPublic.cCtlTake")
		strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1)
			  
		If (rs1.EOF And rs1.BOF) Then
			Call DisplayMsgBox("189200", vbOKOnly, "", "", I_MKSCRIPT)
			rs1.Close
			Set rs1 = Nothing
			Set ADF = Nothing
			Response.End
		Else
			%>
			<Script Language=vbscript>
			With parent.frm1
				.txtItemCd.value				= "<%=ConvSPChars(rs1("Item_Cd"))%>"
				.txtItemNm.value				= "<%=ConvSPChars(rs1("Item_Nm"))%>"
				.txtOrderQty.value				= "<%=UniConvNumberDBToCompany(rs1("Prodt_Order_Qty"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				.txtOrderUnit.value				= "<%=ConvSPChars(rs1("Prodt_Order_Unit"))%>"
				.txtOrderQtyInBaseUnit.value	= "<%=UniConvNumberDBToCompany(rs1("Order_Qty_In_Base_Unit"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				.txtBaseUnit.value				= "<%=ConvSPChars(rs1("Base_Unit"))%>"
				.txtPlanStartDt.text			= "<%=UNIDateClientFormat(rs1("Plan_Start_Dt"))%>"
				.txtPlanComptDt.text			= "<%=UNIDateClientFormat(rs1("Plan_Compt_Dt"))%>"
				.txtReWorkFlag.value			= "<%=ConvSPChars(rs1("Re_Work_Flg"))%>"
				.txtOrderStatus.value			= "<%=ConvSPChars(rs1("Order_Status"))%>"
				.txtTrackingNo.value			= "<%=ConvSPChars(rs1("Tracking_No"))%>"
				.txtRoutingNo.value				= "<%=ConvSPChars(rs1("Rout_No"))%>"
			End With
			</Script>	
			<%
			rs1.Close
			Set rs1 = Nothing
		End If

	End If
	
	Redim UNISqlId(0)
	Redim UNIValue(0, 1)

	UNISqlId(0) = "p4211mb1d"
	
	UNIValue(0, 0) = " " & FilterVar(UCase(Request("txtProdOrdNo")), "''", "S") & ""
	
	If Request("lgStrPrevKey") <> "" Then
		UNIValue(0, 1) = " " & FilterVar(UCase(Request("lgStrPrevKey")), "''", "S") & ""
	Else
		UNIValue(0, 1) = "''"
	End If

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
	Set ADF = Server.CreateObject("prjPublic.cCtlTake")
	strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs2)

	If (rs2.EOF And rs2.BOF) Then
		Call DisplayMsgBox("189300", vbOKOnly, "", "", I_MKSCRIPT)
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
	LngMaxRow = .frm1.vspdData1.MaxRows

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
				strData = strData & Chr(11) & "<%=ConvSPChars(rs2("Opr_No"))%>"												'☆: Operation No.
				strData = strData & Chr(11) & "<%=ConvSPChars(rs2("Job_Cd"))%>"												'☆: Job Code
				strData = strData & Chr(11) & ""
				strData = strData & Chr(11) & "<%=ConvSPChars(rs2("Wc_Cd"))%>"												'☆: Work Center Code
				strData = strData & Chr(11) & "<%=ConvSPChars(rs2("Wc_Nm"))%>"												'☆: Work Center Name
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs2("Plan_Start_Dt"))%>"								'☆: Planned Start Date
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs2("Plan_Compt_Dt"))%>"								'☆: Planned Completion Date
				strData = strData & Chr(11) & "<%=ConvSPChars(rs2("Order_Status"))%>"										'☆: Order Status
				strData = strData & Chr(11) & ""
				strData = strData & Chr(11) & "<%=rs2("Inside_Flg")%>"
				If "<%=rs2("Inside_Flg")%>" = "Y" Then
					strData = strData & Chr(11) & "사내"
				Else
					strData = strData & Chr(11) & "외주"
				End If
				strData = strData & Chr(11) & "<%=rs2("Milestone_Flg")%>"

				strData = strData & Chr(11) & LngMaxRow + <%=i%>
				strData = strData & Chr(11) & Chr(12)
				TmpBuffer(<%=i%>) = strData
<%		
				rs2.MoveNext
			End If
		Next
%>
		iTotalStr = Join(TmpBuffer, "")
		.ggoSpread.Source = .frm1.vspdData1
		.ggoSpread.SSShowDataByClip iTotalStr
		
		.lgStrPrevKey = "<%=Trim(rs2("Opr_No"))%>"
		
<%	
	End If

	rs2.Close
	Set rs2 = Nothing

%>	
	If .frm1.vspdData1.MaxRows < .VisibleRowCnt(.frm1.vspdData1,0) and .lgStrPrevKey <> "" Then
		.initData(LngMaxRow+1)
		.DbQuery
	Else
		.frm1.hPlantCd.value = "<%=ConvSPChars(Request("txtPlantCd"))%>"
		.frm1.hProdOrderNo.value = "<%=ConvSPChars(Request("txtProdOrdNo"))%>"

		.DbQueryOk(LngMaxRow+1)
	End If

End With

</Script>	
<%
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
