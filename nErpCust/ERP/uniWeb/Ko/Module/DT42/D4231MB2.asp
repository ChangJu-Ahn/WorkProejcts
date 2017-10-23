<%'======================================================================================================
'*  1. Module Name          : E-TAX
'*  2. Function Name        : 
'*  3. Program ID           : D4211mb2.asp
'*  4. Program Name         : 전자세금계산서
'*  5. Program Desc         : 
'*  6. Modified date(First) : 2011-05-17
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : 
'*  9. Modifier (Last)      : 
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
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

Dim ADF												'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg										'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter 선언
Dim rs0					                     'DBAgent Parameter 선언
Dim i

On error resume next
Err.Clear			                        '☜: Protect system from crashing
				
											
' Order Header Display
Redim UNISqlId(0)
Redim UNIValue(0, 0)

'If Trim(Request("txtWhereFlag")) = "SD" Then
	'UNISqlId(0) = "D4211MA12"
	UNISqlId(0) = "D4231MA12"
'Else
'	UNISqlId(0) = "D2211MA13"
'End If


'UNIValue(0, 0) = "^"
UNIValue(0, 0) = FilterVar(UCase(Request("txtTaxBillNo")), "''", "S")

UNILock = DISCONNREAD :	UNIFlag = "1"

Set ADF = Server.CreateObject("prjPublic.cCtlTake")
strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

If (rs0.EOF And rs0.BOF) Then
	'Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
	rs0.Close
	Set rs0 = Nothing
	Set ADF = Nothing
	Response.End																		'☜: 비지니스 로직 처리를 종료함 
End If	%>
	
<Script Language=vbscript>
	Dim LngMaxRow
	Dim strData
	Dim TmpBuffer
	Dim iTotalStr1

	With parent																			'☜: 화면 처리 ASP 를 지칭함 
		LngMaxRow = .frm1.vspdData2.MaxRows										'Save previous Maxrow
		ReDim TmpBuffer(<%=rs0.RecordCount - 1%>)

<%		Dim iDx
		For iDx = 0 to rs0.RecordCount-1 %>
			strData = ""
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("item_md"))%>"			
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_name"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_size"))%>"			
			strData = strData & Chr(11) & "<%=UNIConvNumDBToCompanyByCurrency(rs0("item_qty"), gCurrency, ggAmtOfMoneyNo, "X", "X")%>"
			strData = strData & Chr(11) & "<%=UNIConvNumDBToCompanyByCurrency(rs0("unit_price"), gCurrency, ggAmtOfMoneyNo, "X", "X")%>"												
			strData = strData & Chr(11) & "<%=UNIConvNumDBToCompanyByCurrency(rs0("sup_amount"), gCurrency, ggAmtOfMoneyNo, "X", "X")%>"
			strData = strData & Chr(11) & "<%=UNIConvNumDBToCompanyByCurrency(rs0("tax_amount"), gCurrency, ggAmtOfMoneyNo, "X", "X")%>"
			strData = strData & Chr(11) & "<%=UNIConvNumDBToCompanyByCurrency(rs0("total_amount"), gCurrency, ggAmtOfMoneyNo, "X", "X")%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("remark"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("vat_no"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("vat_seq"))%>"
									
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
		
		.DbQueryOk2()
	End With
</Script>	
<%
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>


