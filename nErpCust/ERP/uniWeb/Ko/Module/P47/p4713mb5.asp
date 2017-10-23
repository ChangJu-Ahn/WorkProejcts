<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4713mb5.asp
'*  4. Program Name         : Lookup Production Info.
'*  5. Program Desc         : 
'*  6. Modified date(First) : 2001/12/12
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Chen, Jae Hyun
'*  9. Modifier (Last)      : 
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call LoadBasisGlobalInf

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
	Redim UNIValue(0, 2)

	UNISqlId(0) = "p4713mb5"
	
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtProdtOrderNo")), "''", "S")
	UNIValue(0, 2) = FilterVar(UCase(Request("txtOprNo")), "''", "S")

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
      
	If (rs0.EOF And rs0.BOF) Then
		rs0.Close
		Set rs0 = Nothing
		Set ADF = Nothing
		'Call DisplayMsgBox("189300", vbInformation, "", "", I_MKSCRIPT)	'��: DB �����ڵ�, �޼���Ÿ��, %ó��, ��ũ��Ʈ���� 
		%>
		<Script Language=vbscript>
			Call parent.LookUpOrderDetailFail(CInt("<%=Request("txtRow")%>"))
		</Script>	
		<%		
		Response.End													'��: �����Ͻ� ���� ó���� ������ 
	End If
	
%>
<Script Language=vbscript>

	With parent.frm1.vspdData1

		.Row = CLng("<%=Request("txtRow")%>")


			.Col = parent.C_JobCd
			.text = "<%=ConvSPChars(rs0("job_cd"))%>"
			.Col = parent.C_JobNm
			.text = "<%=ConvSPChars(rs0("job_cd"))%>"
			.Col = parent.C_WcCd
			.text = "<%=ConvSPChars(rs0("wc_cd"))%>"
			.Col = parent.C_WcNm
			.text = "<%=ConvSPChars(rs0("wc_nm"))%>"
			.Col = parent.C_OrderStatus
			.text = "<%=ConvSPChars(rs0("order_status"))%>"
			.Col = parent.C_OrderStatusDesc
			.text = "<%=ConvSPChars(rs0("order_status"))%>"
			.Col = parent.C_PlanStartDt
			.text = "<%=UNIDateClientFormat(rs0("plan_start_dt"))%>"
			.Col = parent.C_PlanComptDt
			.text = "<%=UNIDateClientFormat(rs0("plan_compt_dt"))%>"
			.Col = parent.C_ReleaseDt
			.text = "<%=UNIDateClientFormat(rs0("release_dt"))%>"
			.Col = parent.C_RealStartDt
			.text = "<%=UNIDateClientFormat(rs0("real_start_dt"))%>"
			
			Call parent.LookUpOrderDetailSuccess(CInt("<%=Request("txtRow")%>"))

	End With
	
<%
	rs0.Close
	Set rs0 = Nothing
%>
</Script>	
<%
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
