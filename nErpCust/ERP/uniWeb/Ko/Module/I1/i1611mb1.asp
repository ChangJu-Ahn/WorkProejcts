<%@  LANGUAGE = VBSCript%>
<% Option Explicit%>
<!--'======================================================================================================
'*  1. Module Name          : Inventory
'*  2. Function Name        : ADO Template (Save)
'*  3. Program ID           : i1611mb1.asp
'*  4. Program Name         : 수불현황조회 
'*  5. Program Desc         :
'*  6. Modified date(First) : 2001/01/17
'*  7. Modified date(Last)  : 2005/01/26
'*  8. Modifier (First)     : Lee Seung Wook
'*  9. Modifier (Last)      : Lee Seung Wook
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->

<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%                                                                        
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "I","NOCOOKIE","MB")   
Call HideStatusWnd 

On Error Resume Next

Dim ADF
Dim strRetMsg
Dim UNISqlId, UNIValue, UNILock, UNIFlag
Dim rs0, rs1, rs2, rs3, rs4
Dim strQryMode

Dim i

Const C_SHEETMAXROWS_D = 100

Call HideStatusWnd

strQryMode = Request("lgIntFlgMode")

Dim strTrnsFrDt	                                               
Dim strTrnsToDt	                                               
Dim strFrSlCd	                                               
Dim strToSlCd
Dim strMovType	                                               
Dim strItemCd	                                               
Dim strItemAcct	                                               
Dim strTrnsType                                                
Dim strWcCd
Dim strdocumentNo
Dim strSeqNo
Dim strSubSeqNo
Dim strkeyval

Dim strTrackingNo

Err.Clear																	'☜: Protect system from crashing

	'Order Header Display
	Redim UNISqlId(0)
	Redim UNIValue(0, 12)

	UNISqlId(0) = "160900saa"
	
	Select Case strQryMode
		Case CStr(OPMD_CMODE)
			strTrnsFrDt = FilterVar(Request("txtTrnsFrDt"), "''", "S")
			strkeyval = "|"
		Case CStr(OPMD_UMODE) 
			strTrnsFrDt = FilterVar(UCase(Request("lgStrPrevKey")), "''", "S")
			strdocumentNo = FilterVar(UCase(Request("lgStrPrevKey2")), "''", "S")
			strSeqNo = FilterVar(UCase(Request("lgStrPrevKey3")), "''", "S")
			strSubSeqNo = FilterVar(UCase(Request("lgStrPrevKey4")), "''", "S")
			strkeyval = " ( A.DOCUMENT_DT > "& strTrnsFrDt & _
					" Or ( A.DOCUMENT_DT = "& strTrnsFrDt &" AND B.ITEM_DOCUMENT_NO > "& strdocumentNo &")" & _
					" Or ( A.DOCUMENT_DT = "& strTrnsFrDt &" AND B.ITEM_DOCUMENT_NO = "& strdocumentNo &" and B.SEQ_NO > "& strSeqNo &" )" & _
					" Or ( A.DOCUMENT_DT = "& strTrnsFrDt &" AND B.ITEM_DOCUMENT_NO = "& strdocumentNo &" and B.SEQ_NO = "& strSeqNo &" and B.SUB_SEQ_NO >= "& strSubSeqNo &")) "
	End Select 
	
	IF Request("txtFrSlCd") = "" Then
		strFrSlCd = "|"
	Else
		strFrSlCd = FilterVar(UCase(Request("txtFrSlCd")), "''", "S")
	End IF
	
	IF Request("txtToSlCd") = "" Then
		strToSlCd = "|"
	Else
		strToSlCd = FilterVar(UCase(Request("txtToSlCd")), "''", "S")
	End IF
	
	IF Request("txtItemCd") = "" Then
		strItemCd = "|"
	Else
		StrItemCd = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	End IF

	IF Request("cboItemAcct") = "" Then
		strItemAcct = "|"
	Else
		strItemAcct = FilterVar(UCase(Request("cboItemAcct")), "''", "S")
	End IF

	IF Request("txtWcCd") = "" Then
		strWcCd = "|"
	Else
		strWcCd = " " & FilterVar(UCase(Request("txtWcCd")), "''", "S") & ""
	End IF

	IF Request("txtMovType") = "" Then
		strMovType = "|"
	Else
		strMovType = " " & FilterVar(UCase(Request("txtMovType")), "''", "S") & ""
	End IF
	
	IF Request("cboTrnsType") = "" Then
		strTrnsType = "|"
	Else
		strTrnsType = " " & FilterVar(UCase(Request("cboTrnsType")), "''", "S") & ""
	End IF
	
	IF Request("txtTrackingNo") = "" Then
		strTrackingNo = "|"	
	Else
		strTrackingNo = FilterVar(UCase(Request("txtTrackingNo")), "''", "S")
	End If

	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 2) = FilterVar((Trim(Request("txtTrnsFrDt"))),"''","S")
	UNIValue(0, 3) = FilterVar((Trim(Request("txtTrnsToDt"))), "''", "S")
	UNIValue(0, 4) = strkeyval
	UNIValue(0, 5) = strItemCd 
	UNIValue(0, 6) = strFrSlCd
	UNIValue(0, 7) = strToSlCd		
	UNIValue(0, 8) = strItemAcct
	UNIValue(0, 9) = strWcCd
	UNIValue(0, 10) = strMovType
	UNIValue(0, 11) = strTrnsType
	UNIValue(0, 12) = strTrackingNo
	
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
    	
With parent																	'☜: 화면 처리 ASP 를 지칭함 
	LngMaxRow = .frm1.vspdData.MaxRows										'Save previous Maxrow

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

				strData = "" _
				& Chr(11) & "<%=ConvSPChars(rs0("ITEM_CD"))%>" _
				& Chr(11) & "<%=ConvSPChars(rs0("ITEM_NM"))%>" _
				& Chr(11) & "<%=ConvSPChars(rs0("SL_CD"))%>" _ 
				& Chr(11) & "<%=UNIDateClientFormat(rs0("DOCUMENT_DT"))%>" _
				& Chr(11) & "<%=ConvSPChars(rs0("ORDER_UNIT"))%>" _ 
				& Chr(11) & "<%=ConvSPChars(rs0("BASE_UNIT"))%>" _
				& Chr(11) & "<%=UniConvNumberDBToCompany(rs0("QTY"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>" _
				& Chr(11) & "<%=UniConvNumberDBToCompany(rs0("PRICE"),ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit,0)%>" _
				& Chr(11) & "<%=UniConvNumberDBToCompany(rs0("AMOUNT"),ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit,0)%>" _
				& Chr(11) & "<%=UniConvNumberDBToCompany(rs0("COST_OF_DEVY"),ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit,0)%>" _
				& Chr(11) & "<%=ConvSPChars(rs0("TRACKING_NO"))%>" _ 
				& Chr(11) & "<%=ConvSPChars(rs0("LOT_NO"))%>" _
				& Chr(11) & "<%=ConvSPChars(rs0("LOT_SUB_NO"))%>" _
				& Chr(11) & "<%=ConvSPChars(rs0("TRNS_NM"))%>" _
				& Chr(11) & "<%=ConvSPChars(rs0("MOV_NM"))%>" _
				& Chr(11) & "<%=ConvSPChars(rs0("ITEM_DOCUMENT_NO"))%>" _
				& Chr(11) & "<%=ConvSPChars(rs0("SEQ_NO"))%>" _
				& Chr(11) & "<%=ConvSPChars(rs0("SO_NO"))%>" _
				& Chr(11) & "<%=ConvSPChars(rs0("SO_SEQ"))%>" _
				& Chr(11) & "<%=ConvSPChars(rs0("PO_NO"))%>" _
				& Chr(11) & "<%=ConvSPChars(rs0("PO_SEQ_NO"))%>" _
				& Chr(11) & "<%=ConvSPChars(rs0("PRODT_ORDER_NO"))%>" _
				& Chr(11) & "<%=ConvSPChars(rs0("SPEC"))%>" _
				& Chr(11) & "<%=ConvSPChars(rs0("DEBIT_CREDIT_FLAG"))%>" _
				& Chr(11) & "<%=ConvSPChars(rs0("BP_CD"))%>" _
				& Chr(11) & "<%=ConvSPChars(rs0("BP_NM"))%>" _
				& Chr(11) & "<%=ConvSPChars(rs0("TRNS_SL_CD"))%>" _
				& Chr(11) & "<%=ConvSPChars(rs0("WC_CD"))%>" _
				& Chr(11) & "<%=ConvSPChars(rs0("WC_NM"))%>" _
				& Chr(11) & "<%=ConvSPChars(rs0("DOCUMENT_TEXT"))%>" _
				& Chr(11) & "<%=ConvSPChars(rs0("SUB_SEQ_NO"))%>" _
				& Chr(11) & LngMaxRow + <%=i%> _
				& Chr(11) & Chr(12)
				
				TmpBuffer(<%=i%>) = strData
<%		
				rs0.MoveNext
			End If
		Next
%>
		iTotalStr = Join(TmpBuffer, "")
		.ggoSpread.Source = .frm1.vspdData
		.ggoSpread.SSShowDataByClip iTotalStr
		
		.lgStrPrevKey = "<%=Trim(rs0("DOCUMENT_DT"))%>"
		.lgStrPrevKey2 = "<%=Trim(rs0("ITEM_DOCUMENT_NO"))%>"
		.lgStrPrevKey3 = "<%=Trim(rs0("SEQ_NO"))%>"
		.lgStrPrevKey4 = "<%=Trim(rs0("SUB_SEQ_NO"))%>"
		
<%	
	End If

	rs0.Close
	Set rs0 = Nothing

%>	
	on error resume next
	If .frm1.vspdData.MaxRows < .VisibleRowCnt(.frm1.vspdData,0) and .lgStrPrevKey <> "" and .lgStrPrevKey2 <> "" _
	 and .lgStrPrevKey3 <> "" and .lgStrPrevKey4 <> "" Then
		.DbQuery()
	Else
		.frm1.hPlantCd.value = "<%=ConvSPChars(Request("txtPlantCd"))%>"
		.frm1.hItemCd.value	= "<%=ConvSPChars(Request("txtItemCd"))%>"
		.frm1.hTrnsFrDt.value = "<%=Request("txtTrnsFrDt")%>"
		.frm1.hTrnsToDt.value = "<%=Request("txtTrnsToDt")%>"
		.frm1.hFrSlCd.value	= "<%=ConvSPChars(Request("txtFrSlCd"))%>"
		.frm1.hToSlCd.value	= "<%=ConvSPChars(Request("txtToSlCd"))%>"
		.frm1.hItemAcct.value = "<%=ConvSPChars(Request("cboItemAcct"))%>"
		.frm1.hWcCd.value = "<%=ConvSPChars(Request("txtWcCd"))%>"
		.frm1.hMovType.value = "<%=ConvSPChars(Request("txtMovType"))%>"
		.frm1.hTrnsType.value = "<%=ConvSPChars(Request("cboTrnsType"))%>"
		.frm1.hTrackingNo.value = "<%=ConvSPChars(Request("txtTrackingNo"))%>"
		.DbQueryOk()
	End If

End With

</Script>	
<%
Set ADF = Nothing
%>

