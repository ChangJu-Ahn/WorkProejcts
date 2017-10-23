<%@LANGUAGE = VBScript%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4711rb2.asp
'*  4. Program Name         : 
'*  5. Program Desc         : 
'*  6. Modified date(First) : 2000/12/19
'*  7. Modified date(Last)  : 2002/12/12
'*  8. Modifier (First)     : Park, Bum Soo
'*  9. Modifier (Last)      : Ryu Sung Won
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'======================================================================================================%>

<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "P", "NOCOOKIE", "RB")

Dim ADF												'ActiveX Data Factory ���� �������� 
Dim strRetMsg										'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1		'DBAgent Parameter ���� 
Dim strQryMode
Dim strMode											'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Dim i

Dim strItemCd
Dim strItemAcct
Dim strWcCd
Dim strShiftCd

Call HideStatusWnd

strMode = Request("txtMode")						'�� : ���� ���¸� ���� 
strQryMode = Request("lgIntFlgMode")

On Error Resume Next
Err.Clear
	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	Redim UNISqlId(1)
	Redim UNIValue(1,1)
	
	UNISqlId(0) = "180000sal"
	UNISqlId(1) = "180000saa"
	
	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 1) = FilterVar(UCase(Request("txtBatchRunNo")), "''", "S")
	UNIValue(1, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)
	Set ADF = Nothing
	
	' Plant �� Display      
	If (rs1.EOF And rs1.BOF) Then
		rs1.Close
		Set rs1 = Nothing
		Response.Write "<Script Language=vbscript>"
		Response.Write "parent.frm1.txtPlantNm.value = """""
		Response.Write "</Script>"
		Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
		Response.End
	Else
		Response.Write "<Script Language=vbscript>"
		Response.Write "parent.frm1.txtPlantNm.value = """ & ConvSPChars(rs1("PLANT_NM")) & """"
		Response.Write "</Script>"
	End If
	rs1.Close
	Set rs1 = Nothing
	
	'�̷¹�ȣ ��ȸ 
	If (rs0.EOF And rs0.BOF) Then
		rs0.Close
		Set rs0 = Nothing
		Call DisplayMsgBox("189719", vbOKOnly, "", "", I_MKSCRIPT)
		Response.End
	Else
		rs0.Close
		Set rs0 = Nothing
	End If
	
	Redim UNISqlId(0)
	Redim UNIValue(0, 3)
	
	UNISqlId(0) = "p4711rb2"
	
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 2) = FilterVar(UCase(Request("txtBatchRunNo")), "''", "S")
	UNIValue(0, 3) = " " & FilterVar(gLang, "''", "S") & ""

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
Dim LngMaxRow
Dim strData
Dim TmpBuffer
Dim iTotalStr
    	
With parent																	'��: ȭ�� ó�� ASP �� ��Ī�� 
	LngMaxRow = .frm1.vspdData.MaxRows										'Save previous Maxrow
	ReDim TmpBuffer(<%=rs0.RecordCount-1%>)	
<%  
    For i=0 to rs0.RecordCount-1 
%>
		strData = ""
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("prodt_order_no"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("opr_no"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_cd"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_nm"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("spec"))%>"
		If "<%=ConvSPChars(rs0("error_cd"))%>" = "123800" Then
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("error_desc1"))%>" + " -> " + "<%=ConvSPChars(rs0("error_desc2"))%>" + " : " + "<%=ConvSPChars(rs0("error_nm"))%>"
		ElseIf "<%=ConvSPChars(rs0("error_cd"))%>" = "189718" Then
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("error_desc1"))%>" + " �� " + " : " + "<%=ConvSPChars(rs0("error_nm"))%>"
		ElseIf "<%=ConvSPChars(rs0("error_cd"))%>" = "181500" Then
			strData = strData & Chr(11) & Trim("<%=ConvSPChars(rs0("error_desc1"))%>") + " ����� " + " : " + "<%=ConvSPChars(rs0("error_nm"))%>"
		Else
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("error_nm"))%>"
		End If
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("rout_no"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("report_type"))%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("prod_qty_in_order_unit"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("prodt_order_unit"))%>"
		strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("consumed_dt"))%>"	
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("job_cd"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("job_nm"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("wc_cd"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("wc_nm"))%>"
		strData = strData & Chr(11) & LngMaxRow + <%=i%>
		strData = strData & Chr(11) & Chr(12)
		
		TmpBuffer(<%=i%>) = strData
<%		
		rs0.MoveNext
	Next
%>
	iTotalStr = Join(TmpBuffer, "")
	.ggoSpread.Source = .frm1.vspdData
	.ggoSpread.SSShowDataByClip iTotalStr
		
<%			
	rs0.Close
	Set rs0 = Nothing
%>

End With	
</Script>	
<%
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
