<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4713mb4.asp
'*  4. Program Name         : Lookup Production Info.
'*  5. Program Desc         : 
'*  6. Modified date(First) : 2001-12-10
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Park, Bumsoo
'*  9. Modifier (Last)      : Jeon, Jaehyun
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call LoadBasisGlobalInf
Call loadInfTB19029B("I", "*", "NOCOOKIE", "MB")

On Error Resume Next

Dim ADF										'ActiveX Data Factory ���� �������� 
Dim strRetMsg								'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter ���� 
Dim rs0										'DBAgent Parameter ���� 

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================

Call HideStatusWnd

Dim strItemCd
Dim strOprNo
Dim strRoutNo
Dim strFlag

Err.Clear																	'��: Protect system from crashing

	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================

	' Order Header Display
	Redim UNISqlId(0)
	Redim UNIValue(0, 1)

	UNISqlId(0) = "p4713mb4"
	
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtProdtOrderNo")), "''", "S")

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
      
	If (rs0.EOF And rs0.BOF) Then
		rs0.Close
		Set rs0 = Nothing
		Set ADF = Nothing
		Call DisplayMsgBox("189200", vbInformation, "", "", I_MKSCRIPT)	'��: DB �����ڵ�, �޼���Ÿ��, %ó��, ��ũ��Ʈ���� 
		%>
		<Script Language=vbscript>
			Call parent.LookUpOrderHeaderFail(CInt("<%=Request("txtRow")%>"))
		</Script>	
		<%		
		Response.End													'��: �����Ͻ� ���� ó���� ������ 
	End If
	
%>
<Script Language=vbscript>

	With parent.frm1.vspdData1

		.Row = CLng("<%=Request("txtRow")%>")

			.Col = parent.C_ItemCd
			.value = "<%=ConvSPChars(rs0("item_cd"))%>"
			.Col = parent.C_ItemNm
			.value = "<%=ConvSPChars(rs0("item_nm"))%>"
			.Col = parent.C_Spec
			.value = "<%=ConvSPChars(rs0("spec"))%>"
			.Col = parent.C_ProdtOrderUnit
			.value = "<%=ConvSPChars(rs0("prodt_order_unit"))%>"
			.Col = parent.C_RoutNo
			.value = "<%=ConvSPChars(rs0("rout_no"))%>"
			.Col = parent.C_TrackingNo
			.value = "<%=ConvSPChars(rs0("tracking_no"))%>"
			.Col = parent.C_OrderType
			.value = "<%=ConvSPChars(rs0("prodt_order_type"))%>"
			.Col = parent.C_ProdtOrderQty
			.value = "<%=UniConvNumberDBToCompany(rs0("prodt_order_qty"),ggQty.DecPoint,ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
			
			Call parent.LookUpOrderHeaderSuccess(CInt("<%=Request("txtRow")%>"))

	End With
	
<%
	rs0.Close
	Set rs0 = Nothing
%>
</Script>	
<%
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
