<%@ Language=vbscript%>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: p4711mb1.asp
'*  4. Program Name			: List Shift (Query)
'*  5. Program Desc			: List Shift (Called By Confirm By Operation and Confirm By Order)
'*  6. Comproxy List		: +
'*  7. Modified date(First)	: 2000/03/28
'*  8. Modified date(Last)	: 2002/06/26
'*  9. Modifier (First)		: Park, BumSoo
'* 10. Modifier (Last)		: Park, BumSoo
'* 11. Comment		:
'**********************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call LoadBasisGlobalInf()

On Error Resume Next
Dim ADF										'ActiveX Data Factory ���� �������� 
Dim strRetMsg								'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter ���� 
Dim rs0, rs1										'DBAgent Parameter ���� 
'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Dim i

Call HideStatusWnd

Err.Clear																	'��: Protect system from crashing

	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	Redim UNISqlId(1)
	Redim UNIValue(1, 0)

	UNISqlId(0) = "p4400mb1"
	UNISqlId(1) = "184000saa"
	
	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(1, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)
	
	If rs1.EOF And rs1.BOF Then
		Call DisplayMsgBox("125000", vbInformation, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing	
		rs1.Close
		Set rs1 = Nothing
		Set ADF = Nothing
		Response.Write "<Script Language=vbscript>" & vbCrLf
		Response.Write "	parent.frm1.txtPlantNm.value = """"" & vbCrLf
		Response.Write "	parent.frm1.txtPlantCd.Focus()"
		Response.Write "</Script>" & vbCrLf			
		Response.End													'��: �����Ͻ� ���� ó���� ������ 
	End If
	
	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("180400", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		rs1.Close
		Set rs1 = Nothing
		Set ADF = Nothing
		Response.End
	End If


%>

<Script Language=vbscript>
    	
With parent
	.frm1.txtPlantNm.value		= "<%= ConvSPChars(rs1("plant_nm")) %>"

<%	rs1.Close
	Set rs1 = Nothing
	  
	If Not(rs0.EOF And rs0.BOF) Then
		For i=0 to rs0.RecordCount-1
		
%>
			Call .SetCombo(.frm1.cboShiftCdFrom,"<%=ConvSPChars(rs0("Shift_Cd"))%>","<%=ConvSPChars(rs0("Shift_Cd"))%>")
			Call .SetCombo(.frm1.cboShiftCdTo,"<%=ConvSPChars(rs0("Shift_Cd"))%>","<%=ConvSPChars(rs0("Shift_Cd"))%>")
<%		
			rs0.MoveNext
		Next
	End If

	rs0.Close
	Set rs0 = Nothing

%>	
	.lgShiftCnt = "<%=i%>"

End With

</Script>	
<%
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
