<%@LANGUAGE = VBScript%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : b3b22mb4.asp 
'*  4. Program Name         : Called By B3B22MA1 (Class Management)
'*  5. Program Desc         : Manage Class Information
'*  6. Modified date(First) : 2003/02/12
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Ryu Sung Won
'*  9. Modifier (Last)      : 
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================%>

<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "P", "NOCOOKIE", "MB")

Dim ADF												'ActiveX Data Factory ���� �������� 
Dim strRetMsg										'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag			'DBAgent Parameter ���� 
Dim rs0, rs1, rs2, rs3								'DBAgent Parameter ���� 
Dim strQryMode
Dim strMode																	'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim strCharCd
Dim strCharChoice

strMode = Request("txtMode")												'�� : ���� ���¸� ���� 
strCharCd = UCase(Trim(Request("txtCharCd")))
strCharChoice = Trim(Request("CharChoice"))
'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Call HideStatusWnd

On Error Resume Next
Err.Clear
	
	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	Redim UNISqlId(1)
	Redim UNIValue(1, 0)

	UNISqlId(0) = "b3b21mb1a"
	
	UNIValue(0, 0) = " " & FilterVar(strCharCd, "''", "S") & ""

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

	If (rs0.EOF and rs0.BOF)Then
		'����׸��� �������� �ʽ��ϴ�.
		Call DisplayMsgBox("122630", vbOKOnly, "", "", I_MKSCRIPT)
		If strCharChoice = "1" Then
		%>
		<Script Language=vbscript>
			parent.frm1.txtCharCd1.focus
		</Script>
		<%		
		Else
		%>
		<Script Language=vbscript>
			parent.frm1.txtCharCd2.focus
		</Script>
		<%
		End If
		rs0.Close
		Set rs0 = Nothing
		Set ADF = Nothing
		Response.End
	End If
%>
<Script Language=vbscript>
	With parent.frm1
<%
	If strCharChoice = "1" Then
%>	
		.txtCharCd1.value = "<%=UCase(Trim(rs0("CHAR_CD")))%>"
		.txtCharNm1.value = "<%=rs0("CHAR_NM")%>"
		.txtCharValueDigit1.value = <%=rs0("CHAR_VALUE_DIGIT")%>
<%
	Else
%>
		.txtCharCd2.value = "<%=UCase(Trim(rs0("CHAR_CD")))%>"
		.txtCharNm2.value = "<%=rs0("CHAR_NM")%>"
		.txtCharValueDigit2.value = "<%=rs0("CHAR_VALUE_DIGIT")%>"
<%
	End If
%>
	End With
	
	Call parent.SetClassDigit
<%			
		rs0.Close
		Set rs0 = Nothing
%>
</Script>	
<%
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
