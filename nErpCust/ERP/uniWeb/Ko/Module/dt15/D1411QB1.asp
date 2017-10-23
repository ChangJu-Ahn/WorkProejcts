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

'On Error Resume Next

Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "M", "NOCOOKIE","MB")
Call LoadBNumericFormatB("Q","M", "NOCOOKIE", "MB")

Dim ADF												'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg										'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter 선언 
Dim rs0					                     'DBAgent Parameter 선언 
Dim strMode											'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

Call HideStatusWnd

Dim StrSupplierCd
Dim StrcbobillStatus
Dim StrhdtxtRadio
Dim StrcboTransferStatus
Dim strIssuedFromDt
Dim strIssuedToDt
Dim i

Err.Clear																	'☜: Protect system from crashing



' Order Header Display
Redim UNISqlId(0)
Redim UNIValue(0, 7)

UNISqlId(0) = "D1411MA11"

strMode = Request("txtMode")											'☜ : 현재 상태를 받음 

UNIValue(0, 0) = "^"
UNIValue(0, 1) = FilterVar(UniConvDate(Request("txtIssuedFromDt")), "''", "S")
UNIValue(0, 2) = FilterVar(UniConvDate(Request("txtIssuedToDt")), "''", "S")

If Request("popTaxRecipient") = "" Then
	Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtSupplierNm.value = """"" & vbCrLf
	Response.Write "</Script>" & vbCrLf
	UNIValue(0, 3) = "|"
Else
	UNIValue(0, 3) = FilterVar(UCase(Request("popTaxRecipient")), "''", "S")
End If

If Request("popTaxBizArea") = "" Then
	Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtBizAreaNm.value = """"" & vbCrLf
	Response.Write "</Script>" & vbCrLf
	UNIValue(0, 4) = "|"
Else
	UNIValue(0, 4) = FilterVar(UCase(Request("popTaxBizArea")), "''", "S")
End If

If Request("cboTaxDocumentType") = "" Then
	UNIValue(0, 5) = "|"
Else
	If Request("cboTaxDocumentType") = "00" Then
		UNIValue(0, 5) = FilterVar("", "''", "S")
	Else
		UNIValue(0, 5) = FilterVar(UCase(Request("cboTaxDocumentType")), "''", "S")
	End if	
End If

If Request("cboTransmitStatus") = "" Then
	UNIValue(0, 6) = "|"
Else
	UNIValue(0, 6) = FilterVar(UCase(Request("cboTransmitStatus")), "''", "S")
End If

UNIValue(0, 7) = "|"

UNILock = DISCONNREAD :	UNIFlag = "1"

Set ADF = Server.CreateObject("prjPublic.cCtlTake")

strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)


If (rs0.EOF And rs0.BOF) Then
	Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
	rs0.Close
	Set rs0 = Nothing
	Set ADF = Nothing
	Response.End													'☜: 비지니스 로직 처리를 종료함 
End If	%>

<Script Language=vbscript>
	Dim LngMaxRow
	Dim strData
	Dim TmpBuffer
	Dim iTotalStr1
	Dim aaa

	With parent																	'☜: 화면 처리 ASP 를 지칭함 
		LngMaxRow = .frm1.vspdData.MaxRows										'Save previous Maxrow
		ReDim TmpBuffer(<%=rs0.RecordCount - 1%>)

<%		Dim iDx
		For iDx = 0 to rs0.RecordCount - 1 %>
			aaa = <%=iDx%>
				
			strData = ""
			strData = strData & Chr(11) & ""

			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PROC_FLAG_NM"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("IS_SEND_NM"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("GL_NO"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("VAT_NO"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("INV_NO"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("INV_AMEND_TYPE"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("INV_AMEND_TYPE_NM"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("REMARK"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("REMARK2"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("REMARK3"))%>"
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("ISSUED_DT"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("IO_FG"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("BP_CD"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("BP_NM"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("MADE_VAT_FG"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("DOC_CUR"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("VAT_TYPE"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("VAT_TYPE_NM"))%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("VAT_RATE")   ,ggExchRate.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("NET_AMT")    ,ggAmtOfMoney.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("NET_LOC_AMT"),ggAmtOfMoney.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("VAT_AMT")    ,ggAmtOfMoney.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("VAT_LOC_AMT"),ggAmtOfMoney.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("CREDIT_CD"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("REPORT_BIZ_AREA_CD"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("TAX_BIZ_AREA_NM"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("BIZ_AREA_CD"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("BIZ_AREA_NM"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("REF_NO"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("DISUSE_REASON"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("LEGACY_PK"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("SALE_NO"))%>"			

			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PROC_FLAG"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("IS_SEND"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ISSUE_DT_FG"))%>"

			
			strData = strData & Chr(11) & LngMaxRow + <%=iDx%>
			strData = strData & Chr(11) & Chr(12)

			TmpBuffer(<%=iDx%>) = strData
<%			rs0.MoveNext
		Next %>

		iTotalStr1 = Join(TmpBuffer,"")
		.ggoSpread.Source = .frm1.vspdData
		.ggoSpread.SSShowDataByClip iTotalStr1

<%		rs0.Close
		Set rs0 = Nothing	%>

		.DbQueryOk()
	End With
</Script>
<%
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
