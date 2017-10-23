<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4111mb1.asp 
'*  4. Program Name         : Called By P4111MA1 (Order Management - Single)
'*  5. Program Desc         : Lookup Production Order Header
'*  6. Modified date(First) : 2002/05/07
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Park, BumSoo
'*  9. Modifier (Last)      : Park, BumSoo
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
Call loadInfTB19029B("I", "*", "NOCOOKIE","MB")
On Error Resume Next								'☜: 

Dim ADF												'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg										'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag			'DBAgent Parameter 선언 
Dim rs0, rs1, rs2, rs3									'DBAgent Parameter 선언 
Dim strQryMode
Dim strMode											'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim strProdtOrderNo, strProdtOrderNo_Next, strProdtOrderNo_Previous

'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================
Call HideStatusWnd

On Error Resume Next
	
	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================
	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================
	Redim UNISqlId(2)
	Redim UNIValue(2, 1)

	UNISqlId(0) = "180000saa"
	UNISqlId(1) = "p4419mb1h"
	UNISqlId(2) = "180000sat"
	
	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	
	UNIValue(1, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(1, 1) = FilterVar(UCase(Request("txtProdOrderNo")), "''", "S")
	
	UNIValue(2, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs2, rs3)
	
	Response.Write "<Script Language=vbscript>"
	Response.Write "	parent.frm1.txtPlantNm.value = """""
	Response.Write "</Script>"
		

	' Plant 명 Display      
	If (rs0.EOF And rs0.BOF) Then
		rs0.Close
		Set rs0 = Nothing
		strFlag = "ERROR_PLANT"
		Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
		
		Response.Write "<Script Language=vbscript>"
		Response.Write "	parent.frm1.txtPlantCd.Focus()"
		Response.Write "</Script>"
		
		Set ADF = Nothing
		Response.End
	Else
		Response.Write "<Script Language=vbscript>"
		Response.Write "	parent.frm1.txtPlantNm.value = """ & ConvSPChars(rs0("PLANT_NM")) & """"
		Response.Write "</Script>"
		rs0.Close
		Set rs0 = Nothing
		Set ADF = Nothing
	End If
	
	' GET OPR. COST FLAG    
	If (rs3.EOF And rs3.BOF) Then
		Call DisplayMsgBox("180600", vbOKOnly, "", "", I_MKSCRIPT)
		Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtPlantCd.Focus()" & vbCrLf
		Response.Write "</Script>" & vbCrLf
		rs3.Close
		Set rs3 = Nothing
		Set ADF = Nothing
		Response.End
	Else
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.hOprCostFlag.value = """ & ConvSPChars(rs3("OPR_COST_FLAG")) & """" & vbCrLf
		Response.Write "</Script>" & vbCrLf
		rs3.Close
		Set rs3 = Nothing
	End If
	
	If (rs2.EOF And rs2.BOF) Then
		rs2.Close
		Set rs2 = Nothing
		strFlag = "ERROR_ORDER"
		Call DisplayMsgBox("189200", vbOKOnly, "", "", I_MKSCRIPT)
		
		Response.Write "<Script Language=vbscript>" & vbcrlf
		Response.Write "	parent.frm1.txtProdOrderNo.Focus()" & vbcrlf
		Response.Write "	parent.dbQueryNotOk()" & vbcrlf
		Response.Write "</Script>"
		
		Set ADF = Nothing
		Response.End
	Else
		rs2.Close
		Set rs2 = Nothing
		Set ADF = Nothing
	End If
	

	Redim UNISqlId(0)
	Redim UNIValue(0, 4)

	Select Case Request("txtQueryType")
	
	Case "" , "R"	'Common & Release
		strProdtOrderNo = FilterVar(UCase(Request("txtProdOrderNo")), "''", "S")
		strProdtOrderNo_Next = "|"
		strProdtOrderNo_Previous = "|"

	Case "N"		'Next
		strProdtOrderNo = "|"
		strProdtOrderNo_Next = FilterVar(UCase(Request("txtProdOrderNo")), "''", "S")
		strProdtOrderNo_Previous = "|"
	
	Case "P"		'Previous
		strProdtOrderNo = "|"
		strProdtOrderNo_Next = "|"
		strProdtOrderNo_Previous = FilterVar(UCase(Request("txtProdOrderNo")), "''", "S")

	End Select
	
	Select Case Request("txtQueryType")
		'After Order Release
		Case "R"
			UNISqlId(0) = "p4111mb1H"
		'Else	
		Case Else
			UNISqlId(0) = "p4111mb2H"
	End Select	
	
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 2) = strProdtOrderNo
	UNIValue(0, 3) = strProdtOrderNo_Next
	UNIValue(0, 4) = strProdtOrderNo_Previous

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1)
	Set ADF = Nothing
	
	If (rs1.EOF and rs1.BOF) and Request("txtQueryType") = "" Then
		Call DisplayMsgBox("189226", vbOKOnly, "", "", I_MKSCRIPT)	'Data Not found
		rs1.Close
		Set rs1 = Nothing
		Response.Write "<Script Language=vbscript>"
		Response.Write "parent.dbQueryNotOk"
		Response.Write "</SCRIPT>"
		Response.End
	End If
	
	If (rs1.EOF and rs1.BOF) and Request("txtQueryType") <> "" Then	'When txt is P or N
		Call DisplayMsgBox("900012", vbOKOnly, "", "", I_MKSCRIPT)	'This is the edge data
		rs1.Close
		Set rs1 = Nothing
		
		Redim UNISqlId(0)
		Redim UNIValue(0, 4)
		'---------------------------------------
		' ****  CAUTION ********
		' SqlId is Not p4111mb1.
		' p4111mb1 is used for REFERENCE POPUP
		'----------------------------------------
		UNISqlId(0) = "p4111mb2H" 
		UNIValue(0, 0) = "^"
		UNIValue(0, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
		UNIValue(0, 2) = FilterVar(UCase(Request("txtProdOrderNo")), "''", "S")
		UNIValue(0, 3) = "|"
		UNIValue(0, 4) = "|"
		
		UNILock = DISCONNREAD :	UNIFlag = "1"
	
		Set ADF = Server.CreateObject("prjPublic.cCtlTake")
		strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1)
		Set ADF = Nothing
		If (rs1.EOF and rs1.BOF)Then
			Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
			rs1.Close
			Set rs1 = Nothing
			'Set ADF = Nothing
			Response.End
		End If
		
	End If

%>
<Script Language=vbscript>
	
	With parent.frm1
		.txtProdOrderNo.value	= "<%=ConvSPChars(rs1("Prodt_Order_No"))%>"									'☆: Production Order No
		.txtProdOrderNo1.value	= "<%=ConvSPChars(rs1("Prodt_Order_No"))%>"									'☆: Production Order No
		.txtStatus.value		= "<%=rs1("Order_Status")%>"												'☆: Order Status
		.txtItemCd.value		= "<%=ConvSPChars(rs1("Item_Cd"))%>"										'☆: Item Code
		.txtItemNm.value		= "<%=ConvSPChars(rs1("Item_Nm"))%>"										'☆: Item Name
		.txtSpecification.value	= "<%=ConvSPChars(rs1("Spec"))%>"											'☆: Specification
		.txtOrderQty.value		= "<%=UniConvNumberDBToCompany(rs1("Prodt_Order_Qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"		'☆: Order Quantity
		.txtUnit.value			= "<%=ConvSPChars(rs1("Prodt_Order_Unit"))%>"								'☆: Unit
		.txtBaseOrderQty.value	= "<%=UniConvNumberDBToCompany(rs1("Order_Qty_In_Base_Unit"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"	'☆: Order Quantity in Base Unit
		.txtBaseUnit.value		= "<%=ConvSPChars(rs1("Base_Unit"))%>"										'☆: Basic Unit
		.txtSLCd.value			= "<%=ConvSPChars(rs1("Sl_Cd"))%>"											'☆: Storage Location Code
		.txtSLNm.Value			= "<%=ConvSPChars(rs1("Sl_Nm"))%>"											'☆: Storage Location Code
		.cboReWork.Value		= "<%=rs1("Re_Work_Flg")%>"													'☆: Valid From Date
		.txtParentOrderNo.value = "<%=ConvSPChars(rs1("parent_order_no"))%>"								'☆: ParentOrderNo
		.txtParentOprNo.value	= "<%=ConvSPChars(rs1("parent_opr_no"))%>"								'☆: ParentOprNo
		.txtOrderType.value		= "<%=rs1("Prodt_Order_Type")%>"											'☆: Order Type
		.txtRouting.value		= "<%=ConvSPChars(rs1("Rout_No"))%>"										'☆: Routing
		.txtPlanOrderNo.value	= "<%=ConvSPChars(rs1("Plan_Order_No"))%>"									'☆: Plan Order No		
		.txtTrackingNo.value	= "<%=ConvSPChars(rs1("Tracking_No"))%>"									'☆: Tracking No		
		.txtRemark.value		= "<%=ConvSPChars(rs1("Remark"))%>"											'☆: Remark

		.txtPlanStartDt.text	= "<%=UNIDateClientFormat(rs1("Plan_Start_Dt"))%>"							'☆: BOM Last Updated Date
		.txtPlanEndDt.text		= "<%=UNIDateClientFormat(rs1("Plan_Compt_Dt"))%>"							'☆: Inv Closing Date(년월까지)
		.txtPlannedStartDt.text	= "<%=UNIDateClientFormat(rs1("Schd_Start_Dt"))%>"							'☆: Inv Open Date(년월까지)
		.txtPlannedEndDt.text	= "<%=UNIDateClientFormat(rs1("Schd_Compt_Dt"))%>"							'☆: Valid From Date
		.txtBOMNo.value			= "<%=ConvSPChars(rs1("Bom_No"))%>"											'☆: BOM No.
		.txtReleaseDt.text		= "<%=UNIDateClientFormat(rs1("Release_Dt"))%>"								'☆: Release Date

		.txtProdLT.value		= "<%=ConvSPChars(rs1("Order_Lt_Mfg"))%>"									'☆: Production Lead Time
		.txtMaxLotQty.value		= "<%=UniConvNumberDBToCompany(rs1("Max_Mrp_Qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"			'☆: Maximum Lot Quantity
		.txtMinLotQty.value		= "<%=UniConvNumberDBToCompany(rs1("Min_Mrp_Qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"			'☆: Minimum Lot Quantity
		.txtRoundingQty.value	= "<%=UniConvNumberDBToCompany(rs1("Round_Qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"				'☆: Rounding Quantity
		.txtProdMgr.value		= "<%=ConvSPChars(rs1("Prod_Mgr"))%>"										'☆: Item By Plant
		
		'Add 2005-09-27
		.txtCostCd.value		= "<%=ConvSPChars(rs1("cost_cd"))%>"
		.txtCostNm.value		= "<%=ConvSPChars(rs1("cost_nm"))%>"

		parent.DbQueryOk																					'☜: 조화가 성공 
	
	End With

</Script>	

<%			
		rs0.Close
		Set rs0 = Nothing
		'Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
