<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<%Call LoadBasisGlobalInf%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : b1b01mb5.asp
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) :
'*  7. Modified date(Last)  : 2002/11/15
'*  8. Modifier (First)     : 
'*  9. Modifier (Last)      : Hong Chang Ho
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

On Error Resume Next								'��: 

Dim ADF														'ActiveX Data Factory ���� �������� 
Dim strRetMsg												'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0				'DBAgent Parameter ���� 
Dim strQryMode

Dim strMode											'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Dim StrNextKey		' ���� �� 
Dim lgStrPrevKey1	
Dim LngMaxRow		' ���� �׸����� �ִ�Row
Dim LngRow
Dim GroupCount

Call HideStatusWnd

strMode = Request("txtMode")						'�� : ���� ���¸� ���� 
strQryMode = Request("lgIntFlgMode")

On Error Resume Next

Dim strHsCd

lgStrPrevKey1 = UCase(Trim(Request("lgStrPrevKey1")))	
	
'=======================================================================================================
'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
'=======================================================================================================
	
Redim UNISqlId(0)
Redim UNIValue(0, 0)
	
UNISqlId(0) = "122600sab"
IF Request("txtHsCd") = "" Then
   strHsCd = "|"
ELSE
   strHsCd = FilterVar(Request("txtHsCd") , "''", "S")
END IF
	
UNIValue(0, 0) = strHsCd
	
UNILock = DISCONNREAD :	UNIFlag = "1"
	
Set ADF = Server.CreateObject("prjPublic.cCtlTake")
strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

If rs0.EOF And rs0.BOF Then
'	Call DisplayMsgBox("126700", vbInformation, "", "", I_MKSCRIPT)	'��: DB �����ڵ�, �޼���Ÿ��, %ó��, ��ũ��Ʈ���� 
	Set rs0 = Nothing					
	Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "parent.LookUpHsNotOk" & vbCrLf
	Response.Write "</Script>" & vbCrLf
	Response.End													'��: �����Ͻ� ���� ó���� ������ 
End If

Response.Write "<Script Language=VBScript>" & vbCrLf
	Response.Write "parent.frm1.txtHsUnit.value = """ & ConvSPChars(rs0(3)) & """" & vbCrLf
Response.Write "</Script>" & vbCrLf
rs0.Close
Set rs0 = Nothing
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
Response.End
%>
