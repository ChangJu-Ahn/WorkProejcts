<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4712mb4.asp
'*  4. Program Name         : List RESOURCE infomation	
'*  5. Program Desc         : 
'*  6. Modified date(First) : 2001.12.12
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Jeon, Jaehyun
'*  9. Modifier (Last)      : 
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
'Call loadInfTB19029("I", "*")
Call loadInfTB19029B("I", "*", "NOCOOKIE","MB")

On Error Resume Next

Dim ADF										'ActiveX Data Factory ���� �������� 
Dim strRetMsg								'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter ���� 
Dim rs0, rs1, rs2								'DBAgent Parameter ���� 

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'======================================================================================================
Dim LngRow

Call HideStatusWnd

Dim StrResourceCd
Dim strRcdNm
Dim strRcdtype
Dim strRgcd
Dim strRgcdNm

Err.Clear																	'��: Protect system from crashing

	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================	
	
	' Order information Display
	Redim UNISqlId(0)
	Redim UNIValue(0, 2)

	UNISqlId(0) = "P4712MB4"
	
	IF Request("txtResourceCd") = "" Then
		StrResourceCd = "|"
	Else
		StrResourceCd = FilterVar(UCase(Request("txtResourceCd")), "''", "S")
	End IF

	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 2) = StrResourceCd 
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    
	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		LngRow = Request("Row")
		rs0.Close
		Set rs0 = Nothing
		%>
		<Script Language=vbscript>
		Call	parent.LookupRcNotOk("<%=LngRow%>")
		</Script>	
		<%		
		Response.End													'��: �����Ͻ� ���� ó���� ������ 
	End If
	
	LngRow = Request("Row")
	strRcdNm = Trim(ConvSPChars(rs0("DESCRIPTION")))
	strRcdtype = Trim(ConvSPChars(rs0("MINOR_NM")))
	strRgcd = Trim(ConvSPChars(rs0("RESOURCE_GROUP_CD")))
    strRgcdNm = Trim(ConvSPChars(rs0("GROUP_NM")))

	
%>

<Script Language=vbscript>
	
	
<%	
	rs0.Close
	Set rs0 = Nothing

%>

  Call parent.LookupRcOk("<%=StrResourceCd%>", "<%=strRcdNm%>", "<%=strRcdtype%>", "<%=strRgcd%>", "<%=strRgcdNm%>", "<%=LngRow%>")


	
</Script>	
<%
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
