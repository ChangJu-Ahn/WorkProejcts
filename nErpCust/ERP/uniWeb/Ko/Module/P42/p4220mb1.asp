<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4220mb1.asp 
'*  4. Program Name         : Resource Plan By Production Order
'*  5. Program Desc         : List Production Order
'*  6. Modified date(First) : 2002/03/04
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Park, Bum Soo
'*  9. Modifier (Last)      : Chen, Jae Hyun
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call LoadBasisGlobalInf
On Error Resume Next								'��: 

Dim ADF												'ActiveX Data Factory ���� �������� 
Dim strRetMsg										'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1	'DBAgent Parameter ���� 
Dim strQryMode

Const C_SHEETMAXROWS = 100

Dim strMode											'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Dim StrNextKey		' ���� �� 
Dim LngMaxRow		' ���� �׸����� �ִ�Row
Dim LngRow
Dim strStartDt
Dim strEndDt
Dim strProdOrderNo
Dim i

Call HideStatusWnd

strQryMode = Request("lgIntFlgMode")

On Error Resume Next
	
	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	
	Redim UNISqlId(1)
	Redim UNIValue(1, 4)
	
	UNISqlId(0) = "180000saa"
	UNISqlId(1) = "189701saa"	
	
	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	
	If Request("txtStartDt") = "" Then
		strStartDt = "|"
	Else
		strStartDt = " " & FilterVar(UniConvDate(Request("txtStartDt")), "''", "S") & ""
	End If
	
	If Request("txtEndDt") = "" Then
		strEndDt = "|"
	Else
		strEndDt = " " & FilterVar(UniConvDate(Request("txtEndDt")), "''", "S") & ""
	End If
	
	If Request("lgStrPrevKey1") = "" Then
		strProdOrderNo = "|"
	Else
		strProdOrderNo = FilterVar(UCase(Request("lgStrPrevKey1")), "''", "S")
	End If
	
	UNIValue(1, 0) = "^"
	UNIValue(1, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(1, 2) = strStartDt
	UNIValue(1, 3) = strEndDt
	
	Select Case strQryMode
		Case CStr(OPMD_CMODE)
			UNIValue(1, 4) = "|"
		Case CStr(OPMD_UMODE) 
			UNIValue(1, 4) = strProdOrderNo
	End Select

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)

	' Plant �� Display      
	If (rs0.EOF And rs0.BOF) Then
		rs0.Close
		Set rs0 = Nothing
		Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
		%>
		<Script Language=vbscript>
			parent.frm1.txtPlantNm.value = ""
			parent.frm1.txtPlantCd.Focus()
		</Script>	
		<%    	
		Response.End													'��: �����Ͻ� ���� ó���� ������ 
	Else
		%>
		<Script Language=vbscript>
			parent.frm1.txtPlantNm.value = "<%=ConvSPChars(rs0("PLANT_NM"))%>"
		</Script>	
		<%    	
		rs0.Close
		Set rs0 = Nothing
	End If
      
	If rs1.EOF And rs1.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		Set rs1 = Nothing					
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
	LngMaxRow = .frm1.vspdData1.MaxRows										'Save previous Maxrow
		
<%  
    For i=0 to rs1.RecordCount-1 
		If i < C_SHEETMAXROWS Then
%>
			strData = strData & Chr(11) & "<%=ConvSPChars(rs1("prodt_order_no"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs1("item_cd"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs1("item_nm"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs1("rout_no"))%>"
			strData = strData & Chr(11) & "<%=UniDateClientFormat(rs1("plan_start_dt"))%>"
			strData = strData & Chr(11) & "<%=UniDateClientFormat(rs1("plan_compt_dt"))%>"
			strData = strData & Chr(11) & "<%=UniDateClientFormat(rs1("schd_start_dt"))%>"
			strData = strData & Chr(11) & "<%=UniDateClientFormat(rs1("schd_compt_dt"))%>"
			strData = strData & Chr(11) & "<%=UniDateClientFormat(rs1("real_start_dt"))%>"
			strData = strData & Chr(11) & LngMaxRow + <%=i%>			
			strData = strData & Chr(11) & Chr(12)
<%		
			rs1.MoveNext
		End If
	Next
%>
	
		.ggoSpread.Source = .frm1.vspdData1
		.ggoSpread.SSShowDataByClip strData
		
		
		.lgStrPrevKey1 = "<%=Trim(rs1("PRODT_ORDER_NO"))%>"
		
		.frm1.hPlantCd.value = "<%=ConvSPChars(Request("txtPlantCd"))%>"
		.frm1.hStartDt.value = "<%=Request("txtStartDt")%>"
		.frm1.hEndDt.value = "<%=Request("txtEndDt")%>"		
<%			
		rs1.Close
		Set rs1 = Nothing
%>
	.DbQueryOk
End With	
</Script>	
<%
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
