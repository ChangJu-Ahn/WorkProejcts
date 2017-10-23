<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4114mb2.asp
'*  4. Program Name         : List Production Order Detail (Lower Grid)
'*  5. Program Desc         : 
'*  6. Modified date(First) : 2001. 5. 22
'*  7. Modified date(Last)  : 2001. 9. 04
'*  8. Modifier (First)     : JaeHyun Chen
'*  9. Modifier (Last)      : Park, BumSoo
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
Call loadInfTB19029B("I", "*", "NOCOOKIE","MB")
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
Dim i

Call HideStatusWnd

strMode = Request("txtMode")												'�� : ���� ���¸� ���� 

Dim StrProdOrderNo

Err.Clear																	'��: Protect system from crashing

	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	' Production Results Display
	Redim UNISqlId(0)
	Redim UNIValue(0, 1)

	UNISqlId(0) = "189300sab"
	
	IF Request("txtProdOrderNo") = "" Then
		strProdOrderNo = "|"
	Else
		StrProdOrderNo = FilterVar(UCase(Request("txtProdOrderNo")), "''", "S")
	End IF
	

	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = StrProdOrderNo
		
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
      
	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		Response.End													'��: �����Ͻ� ���� ó���� ������ 
	End If
	
%>

<Script Language=vbscript>
Dim LngMaxRow, LngMaxRows
Dim strTemp
Dim strData, strData1
Dim TmpBuffer, TmpBuffer1
Dim iTotalStr, iTotalStr1
    	
With parent																	'��: ȭ�� ó�� ASP �� ��Ī�� 

	LngMaxRow = .frm1.vspdData2.MaxRows										'Save previous Maxrow
	LngMaxRows = .frm1.vspdData3.MaxRows

<%  
	If Not(rs0.EOF And rs0.BOF) Then
%>	
		ReDim TmpBuffer(<%=rs0.RecordCount - 1%>)
		ReDim TmpBuffer1(<%=rs0.RecordCount - 1%>)
<%		
		For i=0 to rs0.RecordCount-1 
%>
			strData = ""
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("opr_no"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("wc_cd"))%>"
			strData = strData & Chr(11) & ""	' Work Center Popup
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("wc_nm"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("job_cd"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("minor_nm"))%>"
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("plan_start_dt"))%>"
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("plan_compt_dt"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("bp_cd"))%>"
			strData = strData & Chr(11) & ""	' Currency Bp Pupup
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("bp_nm"))%>"

			strData = strData & Chr(11) & "<%=UniConvNumDBToCompanyWithOutChange(rs0("subcontract_prc"), 0)%>"
			strData = strData & Chr(11) & "<%=UniConvNumDBToCompanyWithOutChange(rs0("subcontract_amt"), 0)%>"
			
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("cur_cd"))%>"
			strData = strData & Chr(11) & ""	' Currency Code Pupup
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("tax_type"))%>"
			strData = strData & Chr(11) & ""	' Tax Code Pupup
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("rout_order"))%>"
			strData = strData & Chr(11) & "<%=rs0("milestone_flg")%>"
			strData = strData & Chr(11) & "<%=rs0("insp_flg")%>"
			strData = strData & Chr(11) & "<%=rs0("inside_flg")%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(Trim(Request("txtProdOrderNo")))%>"
			strData = strData & Chr(11) & "<%=rs0("order_status")%>"
			strData = strData & Chr(11) & LngMaxRow + <%=i%>
			strData = strData & Chr(11) & Chr(12)
			
			TmpBuffer(<%=i%>) = strData

			' Insert Into Hidden Grid
			strData1 = ""
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("opr_no"))%>"
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("wc_cd"))%>"
			strData1 = strData1 & Chr(11) & ""	' Work Center Popup
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("wc_nm"))%>"
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("job_cd"))%>"
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("minor_nm"))%>"
			strData1 = strData1 & Chr(11) & "<%=UNIDateClientFormat(rs0("plan_start_dt"))%>"
			strData1 = strData1 & Chr(11) & "<%=UNIDateClientFormat(rs0("plan_compt_dt"))%>"
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("bp_cd"))%>"
			strData1 = strData1 & Chr(11) & ""	' Currency Bp Pupup
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("bp_nm"))%>"

			strData1 = strData1 & Chr(11) & "<%=UniConvNumDBToCompanyWithOutChange(rs0("subcontract_prc"), 0)%>"
			strData1 = strData1 & Chr(11) & "<%=UniConvNumDBToCompanyWithOutChange(rs0("subcontract_amt"), 0)%>"

			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("cur_cd"))%>"
			strData1 = strData1 & Chr(11) & ""	' Currency Code Pupup
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("tax_type"))%>"
			strData1 = strData1 & Chr(11) & ""	' Tax Code Pupup
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("rout_order"))%>"
			strData1 = strData1 & Chr(11) & "<%=rs0("milestone_flg")%>"
			strData1 = strData1 & Chr(11) & "<%=rs0("insp_flg")%>"
			strData1 = strData1 & Chr(11) & "<%=rs0("inside_flg")%>"
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(Trim(Request("txtProdOrderNo")))%>"
			strData1 = strData1 & Chr(11) & "<%=rs0("order_status")%>"
			strData1 = strData1 & Chr(11) & LngMaxRows + <%=i+1%>
			strData1 = strData1 & Chr(11) & Chr(12)
			
			TmpBuffer1(<%=i%>) = strData1
<%		
			rs0.MoveNext
		Next
%>
		iTotalStr = Join(TmpBuffer, "")
		iTotalStr1 = Join(TmpBuffer1, "")
		
		.ggoSpread.Source = .frm1.vspdData2
		Call .ggoSpread.SSShowDataByClip(iTotalStr ,"F")
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData2, LngMaxRow + 1, LngMaxRow + <%=i%>, .C_Currency2,.C_CCFCost2, "C", "I", "X", "X")
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData2, LngMaxRow + 1, LngMaxRow + <%=i%>, .C_Currency2,.C_CCFAmt2, "A", "I", "X", "X")
		.ggoSpread.Source = .frm1.vspdData3
		Call .ggoSpread.SSShowDataByClip(iTotalStr1 ,"F")
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData3, LngMaxRows + 1, LngMaxRows + <%=i%>, .C_Currency3,.C_CCFCost3, "C", "I", "X", "X")
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData3, LngMaxRows + 1, LngMaxRows + <%=i%>, .C_Currency3,.C_CCFAmt3, "A", "I", "X", "X")
<%	
	End If

	rs0.Close
	Set rs0 = Nothing

%>	
	.DbDtlQueryOk()

End With

</Script>	
<%
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
