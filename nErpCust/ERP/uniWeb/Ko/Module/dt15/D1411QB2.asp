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

Call HideStatusWnd

Dim i

Err.Clear											'☜: Protect system from crashing

' Order Header Display
Redim UNISqlId(0)
Redim UNIValue(0, 1)

UNISqlId(0) = "D1411MA12"

strMode = Request("txtMode")											'☜ : 현재 상태를 받음 

UNIValue(0, 0) = FilterVar(UCase(Request("txtVatNo")), "''", "S")
UNIValue(0, 1) = FilterVar(UCase(Request("txtVatNo")), "''", "S")

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
	Dim LngMaxRow, LngMaxRows
	Dim strData1, strData2
	Dim TmpBuffer1, TmpBuffer2
	Dim iTotalStr1, iTotalStr2

	With parent																			'☜: 화면 처리 ASP 를 지칭함 
		LngMaxRow = .frm1.vspdData2.MaxRows	
		LngMaxRows = .frm1.vspdData3.MaxRows
<%  
		If Not(rs0.EOF And rs0.BOF) Then
%>	
			ReDim TmpBuffer1(<%=rs0.RecordCount - 1%>)
			ReDim TmpBuffer2(<%=rs0.RecordCount - 1%>)

<%			Dim iDx
			For iDx = 0 to rs0.RecordCount-1 %>
				strData1 = ""
				
				strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("item"))%>"
				strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("item_std"))%>"
				strData1 = strData1 & Chr(11) & "<%=UniNumClientFormat(rs0("item_prc"),ggAmtOfMoney.DecPoint,0)%>"
				strData1 = strData1 & Chr(11) & "<%=UniNumClientFormat(rs0("item_qty"), ggQty.DecPoint,0)      %>"
				strData1 = strData1 & Chr(11) & "<%=UNIDateClientFormat(rs0("item_date"))%>"
				strData1 = strData1 & Chr(11) & "<%=UniNumClientFormat(rs0("item_amt"),ggAmtOfMoney.DecPoint,0)%>"
				strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("item_tax"))%>"
				strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("item_memo"))%>"

				strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("inv_no"))%>"
				strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("vat_no"))%>"
				strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("inv_item_seq_no"))%>"
				
				
				strData1 = strData1 & Chr(11) & LngMaxRow + <%=iDx%>
				strData1 = strData1 & Chr(11) & Chr(12)

				TmpBuffer1(<%=iDx%>) = strData1
				
				' Insert Into Hidden Grid 가끔 히든과 그리드2가 다를 수 있다.
				strData2 = ""
				
				strData2 = strData2 & Chr(11) & "<%=ConvSPChars(rs0("item"))%>"
				strData2 = strData2 & Chr(11) & "<%=ConvSPChars(rs0("item_std"))%>"
				strData2 = strData2 & Chr(11) & "<%=UniNumClientFormat(rs0("item_prc"),ggAmtOfMoney.DecPoint,0)%>"
				strData2 = strData2 & Chr(11) & "<%=UniNumClientFormat(rs0("item_qty"), ggQty.DecPoint,0)      %>"
				strData2 = strData2 & Chr(11) & "<%=UNIDateClientFormat(rs0("item_date"))%>"
				strData2 = strData2 & Chr(11) & "<%=UniNumClientFormat(rs0("item_amt"),ggAmtOfMoney.DecPoint,0)%>"
				strData2 = strData2 & Chr(11) & "<%=ConvSPChars(rs0("item_tax"))%>"
				strData2 = strData2 & Chr(11) & "<%=ConvSPChars(rs0("item_memo"))%>"

				strData2 = strData2 & Chr(11) & "<%=ConvSPChars(rs0("inv_no"))%>"
				strData2 = strData2 & Chr(11) & "<%=ConvSPChars(rs0("vat_no"))%>"
				strData2 = strData2 & Chr(11) & "<%=ConvSPChars(rs0("inv_item_seq_no"))%>"
				
				
				strData2 = strData2 & Chr(11) & LngMaxRows + <%=iDx%>
				strData2 = strData2 & Chr(11) & Chr(12)

				TmpBuffer2(<%=iDx%>) = strData2
				
<%				rs0.MoveNext
			Next %>

			iTotalStr1 = Join(TmpBuffer1, "")
			iTotalStr2 = Join(TmpBuffer2, "")
			.ggoSpread.Source = .frm1.vspdData2
			.ggoSpread.SSShowDataByClip iTotalStr1
			.ggoSpread.Source = .frm1.vspdData3
			.ggoSpread.SSShowDataByClip iTotalStr2

<%		
		End If
		rs0.Close
		Set rs0 = Nothing	%>
		.DbDtlQueryOk()
	End With
</Script>	
<%
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>


