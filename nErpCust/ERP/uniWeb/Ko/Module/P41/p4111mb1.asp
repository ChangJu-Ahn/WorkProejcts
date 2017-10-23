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
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call LoadBasisGlobalInf
Call loadInfTB19029B("I", "*", "NOCOOKIE","MB")
On Error Resume Next								'��: 

Dim ADF												'ActiveX Data Factory ���� �������� 
Dim strRetMsg										'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag			'DBAgent Parameter ���� 
Dim rs0, rs1, rs2, rs3									'DBAgent Parameter ���� 
Dim strQryMode
Dim strMode											'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim strProdtOrderNo, strProdtOrderNo_Next, strProdtOrderNo_Previous

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Call HideStatusWnd

On Error Resume Next
	
	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
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
		

	' Plant �� Display      
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
		.txtProdOrderNo.value	= "<%=ConvSPChars(rs1("Prodt_Order_No"))%>"									'��: Production Order No
		.txtProdOrderNo1.value	= "<%=ConvSPChars(rs1("Prodt_Order_No"))%>"									'��: Production Order No
		.txtStatus.value		= "<%=rs1("Order_Status")%>"												'��: Order Status
		.txtItemCd.value		= "<%=ConvSPChars(rs1("Item_Cd"))%>"										'��: Item Code
		.txtItemNm.value		= "<%=ConvSPChars(rs1("Item_Nm"))%>"										'��: Item Name
		.txtSpecification.value	= "<%=ConvSPChars(rs1("Spec"))%>"											'��: Specification
		.txtOrderQty.value		= "<%=UniConvNumberDBToCompany(rs1("Prodt_Order_Qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"		'��: Order Quantity
		.txtUnit.value			= "<%=ConvSPChars(rs1("Prodt_Order_Unit"))%>"								'��: Unit
		.txtBaseOrderQty.value	= "<%=UniConvNumberDBToCompany(rs1("Order_Qty_In_Base_Unit"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"	'��: Order Quantity in Base Unit
		.txtBaseUnit.value		= "<%=ConvSPChars(rs1("Base_Unit"))%>"										'��: Basic Unit
		.txtSLCd.value			= "<%=ConvSPChars(rs1("Sl_Cd"))%>"											'��: Storage Location Code
		.txtSLNm.Value			= "<%=ConvSPChars(rs1("Sl_Nm"))%>"											'��: Storage Location Code
		.cboReWork.Value		= "<%=rs1("Re_Work_Flg")%>"													'��: Valid From Date
		.txtParentOrderNo.value = "<%=ConvSPChars(rs1("parent_order_no"))%>"								'��: ParentOrderNo
		.txtParentOprNo.value	= "<%=ConvSPChars(rs1("parent_opr_no"))%>"								'��: ParentOprNo
		.txtOrderType.value		= "<%=rs1("Prodt_Order_Type")%>"											'��: Order Type
		.txtRouting.value		= "<%=ConvSPChars(rs1("Rout_No"))%>"										'��: Routing
		.txtPlanOrderNo.value	= "<%=ConvSPChars(rs1("Plan_Order_No"))%>"									'��: Plan Order No		
		.txtTrackingNo.value	= "<%=ConvSPChars(rs1("Tracking_No"))%>"									'��: Tracking No		
		.txtRemark.value		= "<%=ConvSPChars(rs1("Remark"))%>"											'��: Remark

		.txtPlanStartDt.text	= "<%=UNIDateClientFormat(rs1("Plan_Start_Dt"))%>"							'��: BOM Last Updated Date
		.txtPlanEndDt.text		= "<%=UNIDateClientFormat(rs1("Plan_Compt_Dt"))%>"							'��: Inv Closing Date(�������)
		.txtPlannedStartDt.text	= "<%=UNIDateClientFormat(rs1("Schd_Start_Dt"))%>"							'��: Inv Open Date(�������)
		.txtPlannedEndDt.text	= "<%=UNIDateClientFormat(rs1("Schd_Compt_Dt"))%>"							'��: Valid From Date
		.txtBOMNo.value			= "<%=ConvSPChars(rs1("Bom_No"))%>"											'��: BOM No.
		.txtReleaseDt.text		= "<%=UNIDateClientFormat(rs1("Release_Dt"))%>"								'��: Release Date

		.txtProdLT.value		= "<%=ConvSPChars(rs1("Order_Lt_Mfg"))%>"									'��: Production Lead Time
		.txtMaxLotQty.value		= "<%=UniConvNumberDBToCompany(rs1("Max_Mrp_Qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"			'��: Maximum Lot Quantity
		.txtMinLotQty.value		= "<%=UniConvNumberDBToCompany(rs1("Min_Mrp_Qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"			'��: Minimum Lot Quantity
		.txtRoundingQty.value	= "<%=UniConvNumberDBToCompany(rs1("Round_Qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"				'��: Rounding Quantity
		.txtProdMgr.value		= "<%=ConvSPChars(rs1("Prod_Mgr"))%>"										'��: Item By Plant
		
		'Add 2005-09-27
		.txtCostCd.value		= "<%=ConvSPChars(rs1("cost_cd"))%>"
		.txtCostNm.value		= "<%=ConvSPChars(rs1("cost_nm"))%>"

		parent.DbQueryOk																					'��: ��ȭ�� ���� 
	
	End With

</Script>	

<%			
		rs0.Close
		Set rs0 = Nothing
		'Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
