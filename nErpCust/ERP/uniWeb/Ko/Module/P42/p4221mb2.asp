<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4221mb2.asp
'*  4. Program Name         : Resource Plan
'*  5. Program Desc         : List Resource Plan
'*  6. Modified date(First) : 
'*  7. Modified date(Last)  : 2002/08/21
'*  8. Modifier (First)     : Hong, EunSook
'*  9. Modifier (Last)      : Park, BumSoo
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
On Error Resume Next								'��: 

Dim ADF												'ActiveX Data Factory ���� �������� 
Dim strRetMsg										'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0		'DBAgent Parameter ���� 

Dim strMode											'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Dim StrNextKey		' ���� �� 
Dim lgStrPrevKey	' ���� �� 
Dim LngMaxRow		' ���� �׸����� �ִ�Row
Dim LngRow
Dim strStartDt
Dim strEndDt
Dim i

Call HideStatusWnd

	lgStrPrevKey = Request("lgStrPrevKey2")
	
	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	
	Redim UNISqlId(0)
	Redim UNIValue(0, 4)
	
	UNISqlId(0) = "189700sac"
	
	If Request("txtStartDt") = "" Then
		strStartDt = "|"
	Else
		strStartDt =  " " & FilterVar(UniConvDate(Request("txtStartDt")), "''", "S") & ""
	End If
	
	If Request("txtEndDt") = "" Then
		strEndDt = "|"
	Else
		strEndDt   =  " " & FilterVar(UniConvDate(Request("txtEndDt")), "''", "S") & ""
	End If

	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 2) = FilterVar(UCase(Request("txtResourceGroupCd")), "''", "S")
	UNIValue(0, 3) = strStartDt
	UNIValue(0, 4) = strEndDt
		
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
      
	If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		
		rs0.Close
		Set rs0 = Nothing
					
		Response.End													'��: �����Ͻ� ���� ó���� ������ 
	End If
	
%>
<Script Language=vbscript>
Dim LngLastRow
Dim LngMaxRow
Dim LngRow
Dim strTemp
Dim strData
    	
With parent																	'��: ȭ�� ó�� ASP �� ��Ī�� 
	LngMaxRow = .frm1.vspdData2.MaxRows										'Save previous Maxrow
		
<%  
    For i=0 to rs0.RecordCount-1 
%>
		
		strData = strData & Chr(11) & "<%=UniDateClientFormat(rs0("start_dt"))%>"
		strData = strData & Chr(11) & "<%=UniDateClientFormat(rs0("end_dt"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("prodt_order_no"))%>"
		StrDate = strData & Chr(11) & "<%=ConvSPChars(rs0("load_no"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_cd"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("rout_no"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("opr_no"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("wc_cd"))%>"
		strData = strData & Chr(11) & "<%=rs0("start_flg")%>"
		strData = strData & Chr(11) & "<%=rs0("end_flg")%>"
		strData = strData & Chr(11) & LngMaxRow + <%=i%>
		strData = strData & Chr(11) & Chr(12)
		
<%		
		rs0.MoveNext
	Next
%>
	
		.ggoSpread.Source = .frm1.vspdData2
		.ggoSpread.SSShowDataByClip strData		

		.lgStrPrevKey2 = ""	
		.frm1.hStartDt.value = "<%=Request("txtStartDt")%>"
		.frm1.hEndDt.value = "<%=Request("txtEndDt")%>"
		
<%			
		rs0.Close
		Set rs0 = Nothing
%>
	.DbDtlQueryOk
End With	
</Script>	
<%
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
