<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4414mb2.asp
'*  4. Program Name         : List Production Results
'*  5. Program Desc         : 
'*  6. Modified date(First) : ?
'*  7. Modified date(Last)  : 2002/08/21
'*  8. Modifier (First)     : Park, Bum Soo
'*  9. Modifier (Last)      : Chen, Jae Hyun
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
Call LoadInfTB19029B("Q", "P", "NOCOOKIE","MB")

On Error Resume Next

Dim ADF										'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg								'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter 선언 
Dim rs0										'DBAgent Parameter 선언 
Dim strMode									'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================
Dim StrNextKey		' 다음 값 
Dim GroupCount
Dim i

Call HideStatusWnd

strMode = Request("txtMode")												'☜ : 현재 상태를 받음 

On Error Resume Next

Dim StrProdOrderNo
Dim strOprNo

Err.Clear																	'☜: Protect system from crashing

	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================
	' Production Results Display
	Redim UNISqlId(0)
	Redim UNIValue(0, 3)

	UNISqlId(0) = "P4412MB2"
	
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
		rs0.Close
		Set rs0 = Nothing
		Response.End													'☜: 비지니스 로직 처리를 종료함 
	End If
	
%>

<Script Language=vbscript>
Dim LngMaxRow, LngMaxRows
Dim strTemp
Dim strData, strData1
Dim TmpBuffer1, TmpBuffer2
Dim iTotalStr1, iTotalStr2
    	
With parent																	'☜: 화면 처리 ASP 를 지칭함 
	LngMaxRow = .frm1.vspdData2.MaxRows										'Save previous Maxrow
	LngMaxRows = .frm1.vspdData3.MaxRows
		
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
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("shift_cd"))%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("prod_qty_in_order_unit"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("reason_cd"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("minor_nm"))%>"
			strData = strData & Chr(11) & ""																	'Lot No.
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("rcpt_item_document_no"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("iss_item_document_no"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("insp_req_no"))%>"
			strData = strData & Chr(11) & "<%=UniConvNumDBToCompanyWithOutChange(rs0("subcontract_prc"), 0)%>"
			strData = strData & Chr(11) & "<%=UniConvNumDBToCompanyWithOutChange(rs0("subcontract_amt"), 0)%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("remark"))%>"										'Remark
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("cur_cd"))%>"
			strData = strData & Chr(11) & ""																	'Popup
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("tax_type"))%>"
			strData = strData & Chr(11) & ""																	'Popup
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("prodt_order_no"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("opr_no"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("seq"))%>"
			strData = strData & Chr(11) & ""
			strData = strData & Chr(11) & ""
			strData = strData & Chr(11) & ""
			strData = strData & Chr(11) & ""
			strData = strData & Chr(11) & ""
			strData = strData & Chr(11) & ""
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("insp_good_qty_in_order_unit"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("insp_bad_qty_in_order_unit"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("rcpt_qty_in_order_unit"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
			strData = strData & Chr(11) & LngMaxRow + <%=i%>
			strData = strData & Chr(11) & Chr(12)
			
			TmpBuffer1(<%=i%>) = strData
			
			' Insert Into Hidden Grid
			strData1 = ""
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(UNIDateClientFormat(rs0("report_dt")))%>"
			strData1 = strData1 & Chr(11) & Trim("<%=ConvSPChars(rs0("report_type"))%>")
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("shift_cd"))%>"
			strData1 = strData1 & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("prod_qty_in_order_unit"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("reason_cd"))%>"
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("minor_nm"))%>"
			strData1 = strData1 & Chr(11) & ""																	'Lot No.
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("rcpt_item_document_no"))%>"
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("iss_item_document_no"))%>"
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("insp_req_no"))%>"
			strData1 = strData1 & Chr(11) & "<%=UniConvNumDBToCompanyWithOutChange(rs0("subcontract_prc"), 0)%>"
			strData1 = strData1 & Chr(11) & "<%=UniConvNumDBToCompanyWithOutChange(rs0("subcontract_amt"), 0)%>"
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("remark"))%>"										'Remark
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("cur_cd"))%>"
			strData1 = strData1 & Chr(11) & ""																	'Popup
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("tax_type"))%>"
			strData1 = strData1 & Chr(11) & ""																	'Popup
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("prodt_order_no"))%>"
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("opr_no"))%>"
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("seq"))%>"
			strData1 = strData1 & Chr(11) & ""
			strData1 = strData1 & Chr(11) & ""
			strData1 = strData1 & Chr(11) & ""
			strData1 = strData1 & Chr(11) & ""
			strData1 = strData1 & Chr(11) & ""
			strData1 = strData1 & Chr(11) & ""
			strData1 = strData1 & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("insp_good_qty_in_order_unit"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
			strData1 = strData1 & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("insp_bad_qty_in_order_unit"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
			strData1 = strData1 & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("rcpt_qty_in_order_unit"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
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
		Call .ggoSpread.SSShowDataByClip(iTotalStr1 ,"F")
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData2, 1, .Frm1.vspdData2.MaxRows, .C_CurrencyCode1,.C_SubcontractPrice1, "C", "I", "X", "X")
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData2, 1, .Frm1.vspdData2.MaxRows, .C_CurrencyCode1,.C_SubcontractAmt1, "A", "I", "X", "X")
		
		.ggoSpread.Source = .frm1.vspdData3
		Call .ggoSpread.SSShowDataByClip(iTotalStr2 ,"F")
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData3, 1, .Frm1.vspdData3.MaxRows, .C_CurrencyCode2,.C_SubcontractPrice2, "C", "I", "X", "X")
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData3, 1, .Frm1.vspdData3.MaxRows, .C_CurrencyCode2,.C_SubcontractAmt2, "A", "I", "X", "X")
		
		
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
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
