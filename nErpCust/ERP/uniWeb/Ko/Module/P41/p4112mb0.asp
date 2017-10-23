<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: Production Order (제조오더)
'*  3. Program ID			: p4112mb0.asp
'*  4. Program Name			: Lookup Item By Plant
'*  5. Program Desc			: Production Order Manage (Query)
'*  6. Comproxy List		: +B1b119LookUpItemByPlant
'*  7. Modified date(First)	: 2000/09/28
'*  8. Modified date(Last)	: 2001/03/28
'*  9. Modifier (First)		: Park, Bum Soo
'* 10. Modifier (Last)		: Park, Bum Soo
'* 11. Comment				: 제조오더관리(Multi)에서 품목을 입력하였을 경우 Lookup하는 프로그램 
'**********************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call LoadBasisGlobalInf

On Error Resume Next								'☜: 

Dim ADF												'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg										'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1		'DBAgent Parameter 선언 
Dim strProdtOrderNo, strProdtOrderNo_Next, strProdtOrderNo_Previous

'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================
Call HideStatusWnd

    Err.Clear															'☜: Protect system from crashing

	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================
	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================
	Redim UNISqlId(1)
	Redim UNIValue(1, 1)

	UNISqlId(0) = "180000sab"
	UNISqlId(1) = "p4112mb0"
	UNIValue(0, 0) = " " & FilterVar(UCase(Request("txtItemCd")), "''", "S") & ""
	UNIValue(1, 0) = " " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & ""
	UNIValue(1, 1) = " " & FilterVar(UCase(Request("txtItemCd")), "''", "S") & ""
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)
	
	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("122600", vbOKOnly, "", "", I_MKSCRIPT)
		%>
		<Script Language=vbscript>
		Parent.LookUpItemByPlantFail(CInt("<%=Request("txtRow")%>"))
		</Script>
		<%
		rs0.Close
		Set rs0 = Nothing
		Set ADF = Nothing
		Response.End													'☜: 비지니스 로직 처리를 종료함 
	End If

	If (rs1.EOF And rs1.BOF) Then
		Call DisplayMsgBox("122700", vbOKOnly, "", "", I_MKSCRIPT)
		%>
		<Script Language=vbscript>
		Parent.LookUpItemByPlantFail(CInt("<%=Request("txtRow")%>"))
		</Script>
		<%
		rs1.Close
		Set rs1 = Nothing
		Set ADF = Nothing
		Response.End													'☜: 비지니스 로직 처리를 종료함 
	End If

    If rs1("Procur_Type") = "P" Then
    
		Call DisplayMsgBox("189209", vbInformation, "", "", I_MKSCRIPT)	'⊙: DB 에러코드, 메세지타입, %처리, 스크립트유형 
		%>
		<Script Language=vbscript>
		Parent.LookUpItemByPlantFail(CInt("<%=Request("txtRow")%>"))
		</Script>
		<%
		rs1.Close
		Set rs1 = Nothing
		Set ADF = Nothing
		Response.End																'☜: Process End
    End If
    
%>
<Script Language=vbscript>

	With parent.frm1.vspdData

		.Row = CLng("<%=Request("txtRow")%>")

		If "<%=rs1("Item_Valid_Flg")%>" = "N" or "<%=rs1("Plant_Valid_Flg")%>" = "N" Then 'VALID_FLG
			' Block invalid item
			.Col = parent.C_ItemCode
			.Text = ""
					
			Call parent.DisplayMsgBox("122619", "x", "x", "x")
			
		Else
			If "<%=rs1("Tracking_Flg")%>" = "N" Then 'TRACKING_FLG
				' Block input of tracking no. when item is not a tracking managed item.
				parent.ggoSpread.SpreadLock parent.C_TrackingNo, .Row, parent.C_TrackingNoPopup, .Row
				parent.ggoSpread.SSSetProtected parent.C_TrackingNo,	.Row, .Row
				parent.ggoSpread.SSSetProtected parent.C_TrackingNoPopup, .Row, .Row			
				.Col = parent.C_TrackingNo
				.Text = "*"
			Else
				' Prepare input of tracking no. when item is a tracking managed item.
			    parent.ggoSpread.SpreadUnLock parent.C_TrackingNo, .Row, parent.C_TrackingNoPopup, .Row
				parent.ggoSpread.SSSetRequired parent.C_TrackingNo,	.Row, .Row			
				.Col = parent.C_TrackingNo
				.Text = ""
			End If

			If "<%=rs1("Phantom_Flg")%>" = "Y" Then 'PHANTOM_FLG
				
				Call parent.DisplayMsgBox("189214", "x", "x", "x")
				' Phantom Item can not be ordered to produce.
				
			Else
				' Display Default Values
				.Col = parent.C_ItemName
				.text = "<%=ConvSPChars(rs1("Item_Nm"))%>"
				.Col = parent.C_Specification
				.text = "<%=ConvSPChars(rs1("Spec"))%>"
				.Col = parent.C_OrderUnit
				.value = "<%=ConvSPChars(rs1("Order_Unit_Mfg"))%>"
				.Col = parent.C_BaseUnit
				.value = "<%=ConvSPChars(rs1("Basic_Unit"))%>"
				.Col = parent.C_SLCD
				.value = "<%=ConvSPChars(rs1("Sl_Cd"))%>"
				.Col = parent.C_SLNM
				.value = "<%=ConvSPChars(rs1("Sl_Nm"))%>"
				' Hidden fields for displaying the item information at the bottom screen when row changes on the top grid.
				.Col = parent.C_OrderLtMFG
				.value = "<%=rs1("Order_Lt_Mfg")%>"
				.Col = parent.C_MaxMRPQty
				.value = "<%=rs1("Max_Mrp_Qty")%>"
				.Col = parent.C_MinMRPQty
				.value = "<%=rs1("Min_Mrp_Qty")%>"
				.Col = parent.C_RoundQty
				.value = "<%=rs1("Round_Qty")%>"
				' Display item information at the bottom screen
				parent.frm1.txtOrderUnitMFG.value	= "<%=ConvSPChars(rs1("OrderUnitMfg"))%>"
				parent.frm1.txtOrderLtMFG.value		= "<%=ConvSPChars(rs1("OrderUnitMfg"))%>"
				parent.frm1.txtMaxMRPQty.value		= "<%=rs1("MaxMrpQty")%>"
				parent.frm1.txtMinMRPQty.value		= "<%=rs1("MinMrpQty")%>"
				parent.frm1.txtRoundQty.value		= "<%=rs1("RoundQty")%>"
			End If
		End If

		Call parent.LookUpItemByPlantSuccess("<%=Request("txtItemCd")%>", CInt("<%=Request("txtRow")%>"))

	End With
</Script>
<%
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
