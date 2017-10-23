<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4221mb1
'*  4. Program Name         : 
'*  5. Program Desc         : List Resource
'*  6. Modified date(First) : ?
'*  7. Modified date(Last)  : 2002/08/21
'*  8. Modifier (First)     : Park, Bum Soo
'*  9. Modifier (Last)      : Chen, Jae Hyun
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call LoadBasisGlobalInf

On Error Resume Next								'��: 

Dim ADF												'ActiveX Data Factory ���� �������� 
Dim strRetMsg										'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2		'DBAgent Parameter ���� 

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Dim StrNextKey		' ���� �� 
Dim strFlag
Dim LngMaxRow		' ���� �׸����� �ִ�Row
Dim LngRow
Dim i

'@Var_Declare

Call HideStatusWnd
	
	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	
	Redim UNISqlId(2)
	Redim UNIValue(2, 1)
	
	UNISqlId(0) = "180000saa"
	UNISqlId(1) = "180000sae"
	UNISqlId(2) = "189700saa"
	
	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(1, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(1, 1) = FilterVar(UCase(Request("txtResourceGroupCd")), "''", "S")
	UNIValue(2, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(2, 1) = FilterVar(UCase(Request("txtResourceGroupCd")), "''", "S")
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)
    
    ' �ڿ��׷� �� Display      
	If (rs1.EOF And rs1.BOF) Then
		rs1.Close
		Set rs1 = Nothing
		strFlag = "ERROR_RGGRP"
		%>
		<Script Language=vbscript>
			parent.frm1.txtResourceGroupNm.value = ""
		</Script>	
		<%    	
	Else
		%>
		<Script Language=vbscript>
			parent.frm1.txtResourceGroupNm.value = "<%=ConvSPChars(rs1("DESCRIPTION"))%>"
		</Script>	
		<%    	
		rs1.Close
		Set rs1 = Nothing
	End If
    
    ' Plant �� Display      
	If (rs0.EOF And rs0.BOF) Then
		rs0.Close
		Set rs0 = Nothing
		strFlag = "ERROR_PLANT"
		%>
		<Script Language=vbscript>
			parent.frm1.txtPlantNm.value = ""
		</Script>	
		<%    	
	Else
		%>
		<Script Language=vbscript>
			parent.frm1.txtPlantNm.value = "<%=ConvSPChars(rs0("PLANT_NM"))%>"
		</Script>	
		<%    	
		rs0.Close
		Set rs0 = Nothing
	End If

	If strFlag <> "" Then
		If strFlag = "ERROR_PLANT" Then
			Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>			
			parent.frm1.txtPlantCd.Focus()
			</Script>	
			<%
			Response.End	
		End If
		If strFlag = "ERROR_RGGRP" Then
			Call DisplayMsgBox("181700", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
			parent.frm1.txtResourceGroupCd.Focus()
			</Script>	
			<%
			Response.End	
		End If
	End IF
      
	If rs2.EOF And rs2.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		Set rs2 = Nothing					
		Response.End														'��: �����Ͻ� ���� ó���� ������ 
	End If
%>
<Script Language=vbscript>
Dim LngLastRow
Dim LngMaxRow
Dim LngRow
Dim strTemp
Dim strData
    	
With parent																	'��: ȭ�� ó�� ASP �� ��Ī�� 
	LngMaxRow = .frm1.vspdData1.MaxRows										'Save previous Maxrow
		
<%  
    For i=0 to rs2.RecordCount-1 
%>
			strData = strData & Chr(11) & "<%=ConvSPChars(rs2("resource_cd"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs2("description"))%>"
			strData = strData & Chr(11) & LngMaxRow + <%=i%>
			strData = strData & Chr(11) & Chr(12)
<%		
			rs2.MoveNext
	Next
%>
	
		.ggoSpread.Source = .frm1.vspdData1
		.ggoSpread.SSShowDataByClip strData
		
		.frm1.hPlantCd.value = "<%=ConvSPChars(Request("txtPlantCd"))%>"
		.frm1.hResourceGroupCd.value = "<%=ConvSPChars(Request("txtResourceGroupCd"))%>"
<%			
		rs2.Close
		Set rs2 = Nothing
%>
	.DbQueryOk
End With	
</Script>	
<%
Set ADF = Nothing															'��: ActiveX Data Factory Object Nothing
%>
