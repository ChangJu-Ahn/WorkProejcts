<%@LANGUAGE = VBScript%>
<%'======================================================================================================
'*  1. Module Name          : Purchase
'*  2. Function Name        : Reference Popup Ȯ�������ȸ	
'*  3. Program ID           : m3110rb1.asp
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) : 2005/07/05
'*  7. Modified date(Last)  : 2005/07/05
'*  8. Modifier (First)     : Chen, Jaehyun
'*  9. Modifier (Last)      : Chen, Jaehyun
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
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1	'DBAgent Parameter ���� 
Dim strQryMode
Dim strMode

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Dim i
Dim j

Dim strItemCd
Dim strTrackingNo
Dim strConvType1
Dim strConvType2

Call HideStatusWnd

On Error Resume Next
Err.Clear
	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	Redim UNISqlId(1)
	Redim UNIValue(1, 1)
	
	UNISqlId(0) = "189702sac"
	UNISqlId(1) = "189702sad"
	
	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 1) = FilterVar(UCase(Request("txtPlanOrderNo")), "''", "S")
	
	UNIValue(1, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(1, 1) = FilterVar(UCase(Request("txtPlanOrderNo")), "''", "S")
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)
      
	If (rs0.EOF And rs0.BOF) and (rs1.EOF And rs1.BOF) Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		rs1.Close
		Set rs0 = Nothing
		Set rs1 = Nothing			
		Response.End													'��: �����Ͻ� ���� ó���� ������ 
	End If
%>

<Script Language=vbscript>

Dim LngMaxRow1
Dim LngMaxRow2
Dim strData1
Dim strData2
Dim TmpBuffer1, TmpBuffer2
Dim iTotalStr1, iTotalStr2
    	
With parent																	'��: ȭ�� ó�� ASP �� ��Ī�� 
	LngMaxRow1 = .frm1.vspdData1.MaxRows										'Save previous Maxrow
	LngMaxRow2 = .frm1.vspdData2.MaxRows										'Save previous Maxrow
<%  
	If Not(rs0.EOF And rs0.BOF) Then
%>			
		ReDim TmpBuffer1(<%=rs0.RecordCount - 1%>)
<%
		For i=0 to rs0.RecordCount-1
%>			
			strData1 = ""
            strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("order_no"))%>"
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("item_cd"))%>"
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("item_nm"))%>"			
			strData1 = strData1 & Chr(11) & "<%=UNIDateClientFormat(rs0("start_plan_dt"))%>"
			strData1 = strData1 & Chr(11) & "<%=UNIDateClientFormat(rs0("end_plan_dt"))%>"
			strData1 = strData1 & Chr(11) & "<%=UNINumClientFormat(rs0("plan_qty"), ggQty.DecPoint, 0)%>"
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("tracking_no"))%>"
			strData1 = strData1 & Chr(11) & LngMaxRow + <%=i%>
			strData1 = strData1 & Chr(11) & Chr(12)
			
			TmpBuffer1(<%=i%>) = strData1
<%		
			rs0.MoveNext
		Next
%>
		iTotalStr1 = Join(TmpBuffer1, "")
		.ggoSpread.Source = .frm1.vspdData1
		.ggoSpread.SSShowDataByClip iTotalStr1
		
		
<%	
	End If
%>	

<%  
	If Not(rs1.EOF And rs1.BOF) Then
%>			
		ReDim TmpBuffer2(<%=rs1.RecordCount - 1%>)
<%
		For j=0 to rs1.RecordCount-1
%>
			strData2 = ""
			strData2 = strData2 & Chr(11) & "<%=ConvSPChars(rs1("order_no"))%>"
			strData2 = strData2 & Chr(11) & "<%=ConvSPChars(rs1("item_cd"))%>"
			strData2 = strData2 & Chr(11) & "<%=ConvSPChars(rs1("item_nm"))%>"			
			strData2 = strData2 & Chr(11) & "<%=UNIDateClientFormat(rs1("start_plan_dt"))%>"
			strData2 = strData2 & Chr(11) & "<%=UNIDateClientFormat(rs1("end_plan_dt"))%>"
			strData2 = strData2 & Chr(11) & "<%=UNINumClientFormat(rs1("plan_qty"), ggQty.DecPoint, 0)%>"
			strData2 = strData2 & Chr(11) & "<%=ConvSPChars(rs1("tracking_no"))%>"
			strData2 = strData2 & Chr(11) & LngMaxRow + <%=j%>
			strData2 = strData2 & Chr(11) & Chr(12)
			
			TmpBuffer2(<%=j%>) = strData2
<%		
			rs1.MoveNext
		Next
%>
		iTotalStr2 = Join(TmpBuffer2, "") 
		.ggoSpread.Source = .frm1.vspdData2
		.ggoSpread.SSShowDataByClip iTotalStr2
		

<%
	End If

	rs0.Close
	Set rs0 = Nothing

	rs1.Close
	Set rs1 = Nothing
%>
	.DbQueryOk

End With

</Script>	
<%
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
