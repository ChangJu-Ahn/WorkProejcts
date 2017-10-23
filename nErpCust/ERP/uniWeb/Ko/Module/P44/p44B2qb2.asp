<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4113mb1.asp
'*  4. Program Name         : List Production Order Header
'*  5. Program Desc         : 
'*  6. Modified date(First) : 2002-05-08
'*  7. Modified date(Last)  : 2002-05-08
'*  8. Modifier (First)     : Park, BumSoo
'*  9. Modifier (Last)      : Park, BumSoo
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call LoadBasisGlobalInf
Call loadInfTB19029B("I", "*", "NOCOOKIE","MB")
On Error Resume Next

Dim ADF
Dim strRetMsg
Dim UNISqlId, UNIValue, UNILock, UNIFlag
Dim rs0, rs1, rs2, rs3, rs4
Dim strQryMode
Dim StrNextKey
Dim strItemCd
Dim strProdOrderNo
Dim strTrackingNo
Dim strItemGroupCd
Dim strInfNo
Dim strErpApplyFlag
Dim strFlag
Dim i

	Const C_SHEETMAXROWS_D = 100

	Call HideStatusWnd

	strQryMode = Request("lgIntFlgMode")

	Err.Clear
	'=======================================================================================================
	'	Handle Description
	'=======================================================================================================

	'=======================================================================================================
	'	Main Query - Order Header Display
	'=======================================================================================================
	Redim UNISqlId(0)
	Redim UNIValue(0, 1)

	UNISqlId(0) = "p44B2qb2"


	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = Filtervar(Request("txtInfNo"), "''", "S")

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
      
	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		Set ADF = Nothing
		Response.End													'☜: 비지니스 로직 처리를 종료함 
	End If
	
%>

<Script Language=vbscript>
Dim LngMaxRow
Dim strData
Dim TmpBuffer
Dim iTotalStr
Dim strCur
    	
With parent																	'☜: 화면 처리 ASP 를 지칭함 
	LngMaxRow = .frm1.vspdData2.MaxRows										'Save previous Maxrow

<%  
	If Not(rs0.EOF And rs0.BOF) Then
		If C_SHEETMAXROWS_D < rs0.RecordCount Then 
%>
			ReDim TmpBuffer(<%=C_SHEETMAXROWS_D - 1%>)
<%
		Else
%>			
			ReDim TmpBuffer(<%=rs0.RecordCount - 1%>)
<%
		End If
		
		For i=0 to rs0.RecordCount-1
			If i < C_SHEETMAXROWS_D Then 
%>

			strData = ""
<%			
			If rs0("input_type") = "H" Then
%>											
				strData = strData & Chr(11) & "오더별"
<%			ElseIf  rs0("input_type") = "D" Then			
%>
				strData = strData & Chr(11) & "공정별"			
<%			Else
%>
				strData = strData & Chr(11) & ""	
<%			End If
%>																
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("prodt_order_no"))%>"											
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_cd"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_nm"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("spec"))%>"											
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("opr_no"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("result_seq"))%>"											
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("report_type"))%>"											
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("pop_qty"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("pop_unit"))%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("prodt_order_qty"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("prod_qty"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("report_dt"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("shift_cd"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("description"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("reason_cd"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("reason_nm"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("tracking_no"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("lot_no"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("lot_sub_no"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("rcpt_item_document_no"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("remark"))%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("subcontract_prc"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("C_subcontract_amt"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
			'strData = strData & Chr(11) & "<%=UniConvNumDBToCompanyWithOutChange(rs0("subcontract_prc"), 0)%>"
			'strData = strData & Chr(11) & "<%=UniConvNumDBToCompanyWithOutChange(rs0("subcontract_amt"), 0)%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("cur_cd"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_group_cd"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_group_nm"))%>"
			strData = strData & Chr(11) & LngMaxRow + <%=i%>
			strData = strData & Chr(11) & Chr(12)
			
			TmpBuffer(<%=i%>) = strData
<%		
			rs0.MoveNext
			End If
		Next
%>
		iTotalStr = Join(TmpBuffer, "")
		
		.ggoSpread.Source = .frm1.vspdData2
		.ggoSpread.SSShowDataByClip iTotalStr
		
		'Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData2, LngMaxRow + 1, LngMaxRow + <%=i%>, .parent.gCurrency,.C_subcontract_prc, "C", "I", "X", "X")
		'Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData2, LngMaxRow + 1, LngMaxRow + <%=i%>, .parent.gCurrency,.C_subcontract_amt, "A", "I", "X", "X")
		
<%	
	End If

	rs0.Close
	Set rs0 = Nothing

%>	

End With

</Script>	
<%
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>