<%'======================================================================================================
'*  1. Module Name          : Purchase
'*  2. Function Name        : 
'*  3. Program ID           : m5314mb1.asp
'*  4. Program Name         : 전자세금계산서 
'*  5. Program Desc         : 
'*  6. Modified date(First) : 2009-07-07
'*  7. Modified date(Last)  : 2009-07-07
'*  8. Modifier (First)     : Lee Min Hyung
'*  9. Modifier (Last)      : Lee Min Hyung
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
Dim strMode											'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim strWhereFlag

Call HideStatusWnd

Dim i

Err.Clear											'☜: Protect system from crashing

strWhereFlag = Request("strWhereFlag")

' Order Header Display
Redim UNISqlId(0)
Redim UNIValue(0, 0)

If strWhereFlag = "SD" Then 
    UNISqlId(0) = "D1211MA12"
Else
    UNISqlId(0) = "D1211MA13"
End If

strMode = Request("txtMode")											'☜ : 현재 상태를 받음 

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
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_cd"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_nm"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("spec"))%>"
			strData = strData & Chr(11) & "<%=UNIConvNumDBToCompanyByCurrency(rs0("bill_qty"), gCurrency, ggAmtOfMoneyNo, "X", "X")%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("bill_unit"))%>"
			strData = strData & Chr(11) & "<%=UNIConvNumDBToCompanyByCurrency(rs0("bill_price"), gCurrency, ggAmtOfMoneyNo, "X", "X")%>"
			strData = strData & Chr(11) & "<%=UNIConvNumDBToCompanyByCurrency(rs0("total_amt"), gCurrency, ggAmtOfMoneyNo, "X", "X")%>"
			strData = strData & Chr(11) & "<%=UNIConvNumDBToCompanyByCurrency(rs0("bill_amt"), gCurrency, ggAmtOfMoneyNo, "X", "X")%>"
			strData = strData & Chr(11) & "<%=UNIConvNumDBToCompanyByCurrency(rs0("vat_amt"), gCurrency, ggAmtOfMoneyNo, "X", "X")%>"
			strData = strData & Chr(11) & "<%=UNIConvNumDBToCompanyByCurrency(rs0("total_amt_loc"), gCurrency, ggAmtOfMoneyNo, "X", "X")%>"
			strData = strData & Chr(11) & "<%=UNIConvNumDBToCompanyByCurrency(rs0("bill_amt_loc"), gCurrency, ggAmtOfMoneyNo, "X", "X")%>"
			strData = strData & Chr(11) & "<%=UNIConvNumDBToCompanyByCurrency(rs0("vat_amt_loc"), gCurrency, ggAmtOfMoneyNo, "X", "X")%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("bill_no"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("bill_seq"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("tax_bill_no"))%>"
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
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>


