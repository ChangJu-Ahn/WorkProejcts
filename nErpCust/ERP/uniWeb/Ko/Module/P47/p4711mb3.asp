<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4711mb3.asp
'*  4. Program Name         : �ڿ��Һ������� 
'*  5. Program Desc         :
'*  6. Comproxy List        : +P11011ManageLotPeriod
'*  7. Modified date(First) : 2001-12-07
'*  8. Modified date(Last)  : 2001-12-07
'*  9. Modifier (First)     : Park, BumSoo
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'**********************************************************************************************
'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call LoadBasisGlobalInf
Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide
	
On Error Resume Next														'��: 

Dim ADF												'ActiveX Data Factory ���� �������� 
Dim strRetMsg										'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag			'DBAgent Parameter ���� 
Dim rs0, rs1, rs2
Dim strReturnVal

'-----------------------------------------------------------
' SQL Server, APS DB Server Information Read
'-----------------------------------------------------------
 	Err.Clear																'��: Protect system from crashing

	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	Redim UNISqlId(1)
	Redim UNIValue(1,1)
	
	UNISqlId(0) = "180000saa"
	UNISqlId(1) = "180000sal"
	
	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(1, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(1, 1) = FilterVar(UCase(Request("txtBatchRunNo")), "''", "S")
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1, rs2)
	Set ADF = Nothing

	' Plant �� Display      
	If (rs1.EOF And rs1.BOF) Then
		Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
		Response.Write "<Script Language=vbscript>" & vbcr
		Response.Write "	parent.frm1.txtPlantNm.value = """"" & vbcr
		Response.Write "	parent.frm1.txtPlantCd.Focus()" & vbcr
		Response.Write "</Script>" & vbcr
		Response.End    	
	Else
		Response.Write "<Script Language=vbscript>" & vbcr
		Response.Write "	parent.frm1.txtPlantNm.value = """ & ConvSPChars(rs1("PLANT_NM")) & """" & vbcr
		Response.Write "</Script>" & vbcr
	End If
	
	'�̷¹�ȣ ��ȸ 
	If (rs2.EOF And rs2.BOF) Then
		rs2.Close
		Set rs2 = Nothing
		Call DisplayMsgBox("189719", vbOKOnly, "", "", I_MKSCRIPT) '�̷¹�ȣ�� �������� �ʽ��ϴ�.
		Response.Write "<Script Language=vbscript>" & vbcr
		Response.Write "	parent.frm1.txtBatchRunNo.Focus()" & vbcr
		Response.Write "</Script>" & vbcr
		Response.End
	Else
		rs2.Close
		Set rs2 = Nothing
	End If
		
	Redim UNISqlId(0)
	Redim UNIValue(0,2)

	UNISqlId(0) = "p4711mb3"
	
	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 1) = FilterVar(UCase(Request("txtBatchRunNo")), "''", "S")
	UNIValue(0, 2) = FilterVar(UCase(gUsrID), "''", "S")

	UNILock = DISCONNREAD :	UNIFlag = "1"	
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
	Set ADF = Nothing
	
	'If strRetMsg <> "0;Success" Then
	strReturnVal = split(strRetMsg,gColSep)
	If strReturnVal(0) <> "0" Then
		Call DisplayMsgBox(strRetMsg, vbOKOnly, "", "", I_MKSCRIPT)
	Else
		Call DisplayMsgBox(rs0("error_msg"), vbOKOnly, "", "", I_MKSCRIPT)
	End If
		
%>

<Script Language=vbscript>

parent.frm1.cboStatus.value	= "<%=ConvSPChars(rs0("status"))%>"
parent.frm1.txtBatchRunNo.Value = ""
parent.frm1.txtSuccessCnt.Value = ""
parent.frm1.txtErrorCnt.Value = ""
<%			
	rs0.Close
	Set rs0 = Nothing
%>
	
</Script>
