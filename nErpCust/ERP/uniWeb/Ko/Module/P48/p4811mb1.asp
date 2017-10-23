<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4811mb2.asp
'*  4. Program Name         : MPS Plan & MFG. Prod
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002.10.10
'*  8. Modified date(Last)  : 2002.10.10
'*  9. Modifier (First)     : Lee Woo Guen
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'**********************************************************************************************

'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 

%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<%																			'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call LoadBasisGlobalInf
Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide

On Error Resume Next														'��: 
	
Dim ADF												'ActiveX Data Factory ���� �������� 
Dim strRetMsg										'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag			'DBAgent Parameter ���� 
Dim rs0, rs1
Dim strReturnVal
'-----------------------------------------------------------
' SQL Server, APS DB Server Information Read
'-----------------------------------------------------------
 	Err.Clear																'��: Protect system from crashing

	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	Redim UNISqlId(1)
	Redim UNIValue(1, 0)

	UNISqlId(0) = "180000saa"
	
	UNIValue(0, 0) = " " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & ""

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1)

	%>
	<Script Language=vbscript>
		parent.frm1.txtPlantNm.value = ""
	</Script>	
	<%    	

	' Plant �� Display      
	If (rs1.EOF And rs1.BOF) Then
		Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
		%>
		<Script Language=vbscript>
			parent.frm1.txtPlantNm.value = ""
			parent.frm1.txtPlantNm.Focus()
			rs1.Close
			Set rs1 = Nothing
			Set ADF = Nothing
			Response.End
		</Script>	
		<%    	
	Else
		%>
		<Script Language=vbscript>
			parent.frm1.txtPlantNm.value = "<%=ConvSPChars(rs1("PLANT_NM"))%>"
		</Script>	
		<%    	
	End If
	rs1.Close
	Set rs1 = Nothing
	Set ADF = Nothing

	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	Redim UNISqlId(0)
	Redim UNIValue(0,3)

	UNISqlId(0) = "p4811mb1a"							'��: Statements�� �̿��Ͽ� SP���� 
	
	UNIValue(0, 0) = " " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & ""
	UNIValue(0, 1) = " " & FilterVar(UCase(Request("cboYear")), "''", "S") & ""
	UNIValue(0, 2) = " " & FilterVar(UCase(Request("cboMonth")), "''", "S") & ""
	UNIValue(0, 3) = " " & FilterVar(gUsrID, "''", "S") & ""

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    
	strReturnVal = split(strRetMsg,gColSep)
	If strReturnVal(0) <> "0" Then
		Call DisplayMsgBox(strRetMsg, vbOKOnly, "", "", I_MKSCRIPT)
	Else
		Call DisplayMsgBox(rs0("error_msg"), vbOKOnly, "", "", I_MKSCRIPT)
	End If
	
%>

<Script Language=vbscript>

<%			
	rs0.Close
	Set rs0 = Nothing
%>
	
</Script>	
<%
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
