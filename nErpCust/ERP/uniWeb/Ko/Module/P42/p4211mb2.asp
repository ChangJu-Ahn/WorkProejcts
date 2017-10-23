<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: p4211mb2.asp
'*  4. Program Name			: List Component Requirement (Reservation) (Query)
'*  5. Program Desc			: Used By Goods Issue For Production Order
'*  6. Comproxy List		: DB AGENT							  
'*  7. Modified date(First)	: 2000/03/28
'*  8. Modified date(Last)	: 2002/08/21
'*  9. Modifier (First)		: Park, BumSoo
'* 10. Modifier (Last)		: Chen, Jae Hyun
'* 11. Comment				:
'**********************************************************************************************
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

Dim ADF														'ActiveX Data Factory ���� �������� 
Dim strRetMsg												'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0				'DBAgent Parameter ���� 

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Dim i

Call HideStatusWnd

On Error Resume Next

	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	Redim UNISqlId(0)
	Redim UNIValue(0, 2)

	UNISqlId(0) = "p4211mb2"

	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtProdtOrderNo")), "''", "S")
	UNIValue(0, 2) = FilterVar(UCase(Request("txtOprNo")), "''", "S")
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

	If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)	'��: DB �����ڵ�, �޼���Ÿ��, %ó��, ��ũ��Ʈ���� 
		rs0.Close
		Set rs0 = Nothing				
		%>
		<Script Language=vbscript>
		    LngMaxRow = 0
			Parent.DbDtlQueryNotOk(LngMaxRow)
		</Script>
		<%
		Response.End													'��: �����Ͻ� ���� ó���� ������ 
	End If
%>

<Script Language=vbscript>
Dim LngMaxRow
Dim LngMaxRows
Dim strData, strData1
Dim TmpBuffer, TmpBuffer1
Dim iTotalStr, iTotalStr1
    	
With parent																	'��: ȭ�� ó�� ASP �� ��Ī�� 
	LngMaxRow = .frm1.vspdData2.MaxRows										'Save previous Maxrow
	LngMaxRows = .frm1.vspdData3.MaxRows
	ReDim TmpBuffer(<%=rs0.RecordCount-1 %>)
	ReDim TmpBuffer1(<%=rs0.RecordCount-1 %>)
<%  
    For i=0 to rs0.RecordCount-1 
%>
		strData = ""
		strData = strData & Chr(11) & Trim("<%=ConvSPChars(rs0("Item_Cd"))%>")								'��: Item Code
		strData = strData & Chr(11) & ""																	'��: Item Code Popup	
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Item_Nm"))%>"									'��: Item Name
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("spec"))%>"										'��: Spec
		strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Req_Qty"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"			'��: Required Quantity
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Base_Unit"))%>"									'��: Base Unit
		strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Issued_Qty"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"			'��: Required Quantity
		strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("Req_Dt"))%>"								'��: Required Date	
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Tracking_No"))%>"								'��: Tracking No.
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Sl_Cd"))%>"										'��: Storage Location Code
		strData = strData & Chr(11) & ""																	'��: Storage Location Popup
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Sl_Nm"))%>"										'��: Storage Location Name
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Resv_Status"))%>"								'��: Reseve Status
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Resv_Desc"))%>"									'��: Reseve Status
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Issue_Mthd"))%>"									'��: Issue Method
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Issue_Mthd_Desc"))%>"							'��: Issue Method
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Req_No"))%>"										'��: Required No.
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Seq"))%>"										'��: Sequence
		strData = strData & Chr(11) & "<%=ConvSPChars(Request("txtPlantCd"))%>"								'��: Plant Code
		strData = strData & Chr(11) & Trim("<%=ConvSPChars(rs0("Prodt_Order_No"))%>")						'��: Production Order No.
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Wc_Cd"))%>"										'��: Work Center
		strData = strData & Chr(11) & Trim("<%=ConvSPChars(rs0("Opr_No"))%>")								'��: Operation No.	
		strData = strData & Chr(11) & Trim("<%=ConvSPChars(rs0("Item_Cd"))%>")								'��: Item Code
		strData = strData & Chr(11) & Trim("<%=ConvSPChars(rs0("Order_Status"))%>")							'��: Item Code
		strData = strData & Chr(11) & LngMaxRow + <%=i%>
		strData = strData & Chr(11) & Chr(12)
		
		TmpBuffer(<%=i%>) = strData

		' Insert Into Hidden Grid
		strData1 = ""
		strData1 = strData1 & Chr(11) & Trim("<%=ConvSPChars(rs0("Item_Cd"))%>")								'��: Item Code
		strData1 = strData1 & Chr(11) & ""																		'��: Item Code Popup	
		strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("Item_Nm"))%>"										'��: Item Name
		strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("spec"))%>"											'��: Spec
		strData1 = strData1 & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Req_Qty"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"			'��: Required Quantity
		strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("Base_Unit"))%>"									'��: Base Unit
		strData1 = strData1 & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Issued_Qty"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"			'��: Required Quantity
		strData1 = strData1 & Chr(11) & "<%=UNIDateClientFormat(rs0("Req_Dt"))%>"								'��: Required Date	
		strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("Tracking_No"))%>"									'��: Tracking No.
		strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("Sl_Cd"))%>"										'��: Storage Location Code
		strData1 = strData1 & Chr(11) & ""																		'��: Storage Location Popup
		strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("Sl_Nm"))%>"										'��: Storage Location Name
		strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("Resv_Status"))%>"									'��: Reserve Status
		strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("Resv_Desc"))%>"									'��: Reserve Status
		strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("Issue_Mthd"))%>"									'��: Issue Method
		strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("Issue_Mthd_Desc"))%>"									'��: Issue Method
		strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("Req_No"))%>"										'��: Required No.
		strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("Seq"))%>"											'��: Sequence
		strData1 = strData1 & Chr(11) & "<%=ConvSPChars(Request("txtPlantCd"))%>"								'��: Plant Code
		strData1 = strData1 & Chr(11) & Trim("<%=ConvSPChars(rs0("Prodt_Order_No"))%>")							'��: Production Order No.
		strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("Wc_Cd"))%>"										'��: Work Center
		strData1 = strData1 & Chr(11) & Trim("<%=ConvSPChars(rs0("Opr_No"))%>")									'��: Operation No.	
		strData1 = strData1 & Chr(11) & Trim("<%=ConvSPChars(rs0("Item_Cd"))%>")								'��: Item Code
		strData1 = strData1 & Chr(11) & Trim("<%=ConvSPChars(rs0("Order_Status"))%>")							'��: Item Code
		strData1 = strData1 & Chr(11) & LngMaxRows + <%=i%>
		strData1 = strData1 & Chr(11) & Chr(12)
		
		TmpBuffer1(<%=i%>) = strData1
<%		
		rs0.MoveNext
	Next
%>
		iTotalStr = Join(TmpBuffer, "")
		.ggoSpread.Source = .frm1.vspdData2
		.ggoSpread.SSShowDataByClip iTotalStr
		
		iTotalStr1 = Join(TmpBuffer1, "")
		.ggoSpread.Source = .frm1.vspdData3
		.ggoSpread.SSShowDataByClip iTotalStr1
		
		.frm1.hProdOrderNo.value= "<%=ConvSPChars(Request("txtProdtOrderNo"))%>"
		
<%			
		rs0.Close
		Set rs0 = Nothing
%>
	.DbDtlQueryOk(LngMaxRow+1)
End With	
</Script>	
<%
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
