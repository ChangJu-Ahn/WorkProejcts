<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4412mb2.asp
'*  4. Program Name         : List Production Results
'*  5. Program Desc         : 
'*  6. Modified date(First) : ?
'*  7. Modified date(Last)  : 2003/03/26
'*  8. Modifier (First)     : Park, Bum Soo
'*  9. Modifier (Last)      : Chen, Jae Hyun
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================
%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "P", "NOCOOKIE","MB")

On Error Resume Next

Dim ADF										'ActiveX Data Factory ���� �������� 
Dim strRetMsg								'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter ���� 
Dim rs0										'DBAgent Parameter ���� 
Dim strMode									'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Dim StrNextKey		' ���� �� 
Dim GroupCount
Dim i

Call HideStatusWnd

strMode = Request("txtMode")												'�� : ���� ���¸� ���� 

Dim StrProdOrderNo
Dim strOprNo

Err.Clear																	'��: Protect system from crashing

	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	' Production Results Display
	Redim UNISqlId(0)
	Redim UNIValue(0, 3)

	UNISqlId(0) = "P4412MB2_KO441"      '2008-03-25 2:30���� :: hanc
	
	IF Request("txtProdOrderNo") = "" Then
		strProdOrderNo = "|"
	Else
		StrProdOrderNo = FilterVar(UCase(Request("txtProdOrderNo")), "''", "S")
	End IF

	IF Request("txtOprNo") = "" Then
		strOprNo = "|"
	Else
		StrOprNo = FilterVar(UCase(Request("txtOprNo")), "''", "S")
	End IF

	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 2) = StrProdOrderNo
	UNIValue(0, 3) = StrOprNo
		
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
      
	If (rs0.EOF And rs0.BOF) Then
'		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		Response.End													'��: �����Ͻ� ���� ó���� ������ 
	End If
	
%>

<Script Language=vbscript>
Dim LngMaxRow, LngMaxRows		' ���� �׸����� �ִ�Row
Dim strData, strData1
Dim TmpBuffer1, TmpBuffer2
Dim iTotalStr1, iTotalStr2
    	
With parent																	'��: ȭ�� ó�� ASP �� ��Ī�� 
	LngMaxRow = .frm1.vspdData2.MaxRows										'Save previous Maxrow
	LngMaxRows = .frm1.vspdData3.MaxRows

	.frm1.vspdData3.MaxRows = .frm1.vspdData3.MaxRows + Clng(<%=rs0.RecordCount%>)
		
<%  
	If Not(rs0.EOF And rs0.BOF) Then
%>	
		ReDim TmpBuffer1(<%=rs0.RecordCount - 1%>)
		ReDim TmpBuffer2(<%=rs0.RecordCount - 1%>)
<%		
		For i=0 to rs0.RecordCount-1 
%>
			strData = ""
			strData = strData & Chr(11) & "<%=ConvSPChars(UNIDateClientFormat(rs0("report_dt")))%>"
			strData = strData & Chr(11) & Trim("<%=ConvSPChars(rs0("report_type"))%>")
			strData = strData & Chr(11) & Trim("<%=ConvSPChars(rs0("shift_cd"))%>")
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("prod_qty_in_order_unit"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("reason_cd"))%>" '5
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("minor_nm"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("remark"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("MACHINE_CD"))%>"     '2008-03-25 2:31���� :: hanc
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("MACHINE_NM"))%>"     '2008-03-25 2:30���� :: hanc
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("lot_no"))%>"															'Lot No.
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("lot_sub_no"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("rcpt_item_document_no"))%>" '10
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("iss_item_document_no"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("insp_req_no"))%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("insp_good_qty_in_order_unit"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("insp_bad_qty_in_order_unit"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("rcpt_qty_in_order_unit"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>" '15
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("prodt_order_no"))%>" '20
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("opr_no"))%>"
			strData = strData & Chr(11) & "<%=rs0("seq")%>"
			strData = strData & Chr(11) & ""
			strData = strData & Chr(11) & ""
			strData = strData & Chr(11) & "" '25
			strData = strData & Chr(11) & "" '26
			strData = strData & Chr(11) & ""
			strData = strData & Chr(11) & ""
			strData = strData & Chr(11) & ""
			strData = strData & Chr(11) & LngMaxRow + <%=i%>
			strData = strData & Chr(11) & Chr(12)
			
			TmpBuffer1(<%=i%>) = strData 
				
			' Insert Into Hidden Grid
			
			strData1 = ""
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(UNIDateClientFormat(rs0("report_dt")))%>"
			strData1 = strData1 & Chr(11) & Trim("<%=ConvSPChars(rs0("report_type"))%>")
			strData1 = strData1 & Chr(11) & Trim("<%=ConvSPChars(rs0("shift_cd"))%>")
			strData1 = strData1 & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("prod_qty_in_order_unit"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("reason_cd"))%>" '5
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("minor_nm"))%>"
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("remark"))%>"
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("MACHINE_CD"))%>"     '2008-03-25 2:31���� :: hanc
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("MACHINE_NM"))%>"     '2008-03-25 2:30���� :: hanc
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("lot_no"))%>"															'Lot No.
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("lot_sub_no"))%>"
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("rcpt_item_document_no"))%>" '10
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("iss_item_document_no"))%>"
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("insp_req_no"))%>"
			strData1 = strData1 & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("insp_good_qty_in_order_unit"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
			strData1 = strData1 & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("insp_bad_qty_in_order_unit"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
			strData1 = strData1 & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("rcpt_qty_in_order_unit"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>" '15
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("prodt_order_no"))%>" '20
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("opr_no"))%>"
			strData1 = strData1 & Chr(11) & "<%=rs0("seq")%>"
			strData1 = strData1 & Chr(11) & ""
			strData1 = strData1 & Chr(11) & ""
			strData1 = strData1 & Chr(11) & "" '25
			strData1 = strData1 & Chr(11) & "" '26
			strData1 = strData1 & Chr(11) & ""
			strData1 = strData1 & Chr(11) & ""
			strData1 = strData1 & Chr(11) & ""
			strData1 = strData1 & Chr(11) & LngMaxRows + <%=i + 1%>
			strData1 = strData1 & Chr(11) & Chr(12)
			
			TmpBuffer2(<%=i%>) = strData1 
<%		
			rs0.MoveNext
		Next
%>
		iTotalStr1 = Join(TmpBuffer1, "")
		iTotalStr2 = Join(TmpBuffer2, "")
		
		.ggoSpread.Source = .frm1.vspdData2
		.ggoSpread.SSShowDataByClip iTotalStr1
		.ggoSpread.Source = .frm1.vspdData3
		.ggoSpread.SSShowDataByClip iTotalStr2
<%	
	End If

	rs0.Close
	Set rs0 = Nothing

%>	
	.frm1.hPlantCd.value		= "<%=ConvSPChars(Request("txtPlantCd"))%>"
	.frm1.hProdOrderNo.value	= "<%=ConvSPChars(Request("txtProdOrderNo"))%>"

End With

</Script>	
<%
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
