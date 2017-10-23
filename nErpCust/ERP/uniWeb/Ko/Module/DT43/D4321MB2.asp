<%'======================================================================================================
'*  1. Module Name          : E-TAX
'*  2. Function Name        : 
'*  3. Program ID           : D4211mb2.asp
'*  4. Program Name         : ���ڼ��ݰ�꼭
'*  5. Program Desc         : 
'*  6. Modified date(First) : 2011-05-17
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : 
'*  9. Modifier (Last)      : 
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
Call LoadInfTB19029B("Q", "M", "NOCOOKIE","MB")
Call LoadBNumericFormatB("Q","M", "NOCOOKIE", "MB")

Dim ADF												'ActiveX Data Factory ���� �������� 
Dim strRetMsg										'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter ����
Dim rs0					                     'DBAgent Parameter ����
Dim i

 On Error Resume Next                                                              '��: Protect system from crashing
 Err.Clear                                                                         '��: Clear Error status

' Order Header Display
Redim UNISqlId(0)
Redim UNIValue(0, 0)

UNISqlId(0) = "D4321MA12"


'UNIValue(0, 0) = "^"
UNIValue(0, 0) = FilterVar(UCase(Request("txtTaxBillNo")), "''", "S")

UNILock = DISCONNREAD :	UNIFlag = "1"

Set ADF = Server.CreateObject("prjPublic.cCtlTake")
strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

If (rs0.EOF And rs0.BOF) Then
	Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
	rs0.Close
	Set rs0 = Nothing
	Set ADF = Nothing
	Response.End																		'��: �����Ͻ� ���� ó���� ������ 
End If	%>
	
<Script Language=vbscript>
	Dim LngMaxRow
	Dim strData
	Dim TmpBuffer
	Dim iTotalStr1

	With parent																			'��: ȭ�� ó�� ASP �� ��Ī�� 
		LngMaxRow = .frm1.vspdData2.MaxRows										'Save previous Maxrow
		ReDim TmpBuffer(<%=rs0.RecordCount - 1%>)

<%		Dim iDx
		For iDx = 0 to rs0.RecordCount-1 %>
			strData = ""
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_cd"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_nm"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("spec"))%>"
			strData = strData & Chr(11) & "<%=UNIConvNumDBToCompanyByCurrency(rs0("IV_QTY"), gCurrency, ggAmtOfMoneyNo, "X", "X")%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("IV_UNIT"))%>"
			strData = strData & Chr(11) & "<%=UNIConvNumDBToCompanyByCurrency(rs0("IV_PRC"), gCurrency, ggAmtOfMoneyNo, "X", "X")%>"
			strData = strData & Chr(11) & "<%=UNIConvNumDBToCompanyByCurrency(rs0("total_amt"), gCurrency, ggAmtOfMoneyNo, "X", "X")%>"
			strData = strData & Chr(11) & "<%=UNIConvNumDBToCompanyByCurrency(rs0("IV_DOC_AMT"), gCurrency, ggAmtOfMoneyNo, "X", "X")%>"
			strData = strData & Chr(11) & "<%=UNIConvNumDBToCompanyByCurrency(rs0("VAT_DOC_AMT"), gCurrency, ggAmtOfMoneyNo, "X", "X")%>"
			strData = strData & Chr(11) & "<%=UNIConvNumDBToCompanyByCurrency(rs0("TOTAL_AMT_LOC"), gCurrency, ggAmtOfMoneyNo, "X", "X")%>"
			strData = strData & Chr(11) & "<%=UNIConvNumDBToCompanyByCurrency(rs0("IV_LOC_AMT"), gCurrency, ggAmtOfMoneyNo, "X", "X")%>"
			strData = strData & Chr(11) & "<%=UNIConvNumDBToCompanyByCurrency(rs0("VAT_LOC_AMT"), gCurrency, ggAmtOfMoneyNo, "X", "X")%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("IV_NO"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("IV_SEQ_NO"))%>"
			strData = strData & Chr(11) & LngMaxRow + <%=iDx%>
			strData = strData & Chr(11) & Chr(12)

			TmpBuffer(<%=iDx%>) = strData
<%			rs0.MoveNext
		Next %>

		iTotalStr1 = Join(TmpBuffer, "")
		.ggoSpread.Source = .frm1.vspdData2
		.ggoSpread.SSShowDataByClip iTotalStr1

<%		rs0.Close
		Set rs0 = Nothing	%>
	End With
</Script>	
<%
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>


