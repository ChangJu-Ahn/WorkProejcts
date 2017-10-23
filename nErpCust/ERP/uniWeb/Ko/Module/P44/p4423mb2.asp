<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4423mb1.asp
'*  4. Program Name         : ���ְ����񳻿� ��ȸ 
'*  5. Program Desc         :
'*  6. Modified date(First) : 2001.11.28
'*  7. Modified date(Last)  : 2002/11/21
'*  8. Modifier (First)     : Chen, Jae Hyun
'*  9. Modifier (Last)      : Kang Hyo Ku
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

Dim ADF												'ActiveX Data Factory ���� �������� 
Dim strRetMsg										'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag			'DBAgent Parameter ���� 
Dim rs0												'DBAgent Parameter ���� 
Dim strQryMode, strFlag

Dim strBpCd
Dim strFromDt
Dim strToDt
Dim StrPlantCd
Dim StrWcCd
Dim StrCurCd
Dim StrTaxType
Dim strTemp

Const C_SHEETMAXROWS_D = 100

Dim strMode											'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim lgStrPrevKey5
Dim lgStrPrevKey6
Dim lgStrPrevKey7
Dim i

Call HideStatusWnd

On Error Resume Next								'��: 

strMode = Request("txtMode")						'�� : ���� ���¸� ���� 
strQryMode = Request("lgIntFlgMode")
lgStrPrevKey5 = FilterVar(UCase(Request("lgStrPrevKey5")), "''", "S")
lgStrPrevKey6 = FilterVar(UCase(Request("lgStrPrevKey6")), "''", "S")
lgStrPrevKey7 = FilterVar(UCase(Request("lgStrPrevKey7")), "''", "S")
		        
'=======================================================================================================
'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
'=======================================================================================================
	Redim UNISqlId(0)
	Redim UNIValue(0, 8)

	UNISqlId(0) = "p4423mb2a"
	
	IF Request("txtBpCd") = "" Then
		strBpCd = "|"
	Else
		strBpCd = FilterVar(UCase(Request("txtBpCd")), "''", "S")
	End IF
	
	IF UNIConvDate(Request("txtFromDt")) = UNIConvDate("") Then
		strFromDt = "|"
	Else
		strFromDt = " " & FilterVar(UniConvDate(Request("txtFromDt")), "''", "S") & ""
	End IF
	
	IF UNIConvDate(Request("txtToDt")) = UNIConvDate("") Then
		strToDt = "|"
	Else
		strToDt = " " & FilterVar(UniConvDate(Request("txtToDt")), "''", "S") & ""
	End IF

	IF Request("txtPlantCd") = "" Then
		strPlantCd = "|"
	Else
		strPlantCd = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	End IF

	IF Request("txtWcCd") = "" Then
		strWcCd = "|"
	Else
		StrWcCd = FilterVar(UCase(Request("txtWcCd")), "''", "S")
	End IF
	
	IF Request("txtCurCd") = "" Then
		strCurCd = "|"
	Else
		StrCurCd = FilterVar(UCase(Request("txtCurCd")), "''", "S")
	End IF
	
	IF Request("txtTaxType") = "" Then
		strTaxType = "|"
	Else
		StrTaxType = FilterVar(UCase(Request("txtTaxType")), "''", "S")
	End IF
		
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = strBpCd
	UNIValue(0, 2) = strFromDt
	UNIValue(0, 3) = strToDt
	UNIValue(0, 4) = strPlantCd
	UNIValue(0, 5) = strWcCd
	UNIValue(0, 6) = strCurCd
	UNIValue(0, 7) = strTaxType
	
	Select Case strQryMode
		Case CStr(OPMD_CMODE)		
			UNIValue(0, 8) = "|" 
		Case CStr(OPMD_UMODE)
			 strTemp = ""
			 strTemp = "(A.PRODT_ORDER_NO > " & lgStrPrevKey5 
			 strTemp = strTemp  & " or (A.PRODT_ORDER_NO = " & lgStrPrevKey5   'second condition  for group view
			 strTemp = strTemp  & " and F.OPR_NO > " & lgStrPrevKey6 & ") "  'second condition  for group view
			 strTemp = strTemp  & " or (A.PRODT_ORDER_NO = " & lgStrPrevKey5    'third condition  for group view
			 strTemp = strTemp  & " and F.OPR_NO = " & lgStrPrevKey6 		'third condition  for group view
			 strTemp = strTemp  & " and F.SEQ >= " & lgStrPrevKey7 & ")) "  'third condition  for group view 
			UNIValue(0, 8) = strTemp
	End Select	
	
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
Dim LngMaxRow
Dim strData
Dim TmpBuffer
Dim iTotalStr
    	
With parent																						'��: ȭ�� ó�� ASP �� ��Ī�� 

	LngMaxRow = .frm1.vspdData2.MaxRows															'Save previous Maxrow
		
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
			If i < C_SHEETMAXROWS_D THEN
%>
				strData = ""
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PRODT_ORDER_NO"))%>"							'������ȣ 
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("PRODT_ORDER_QTY"),ggQty.DecPoint,0)%>" '�������� 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PRODT_ORDER_UNIT"))%>"						'�������� 
				strData = strData & Chr(11) & "<%=UniDateClientFormat(rs0("REPORT_DT"))%>"						'�԰��� 
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("PROD_QTY_IN_ORDER_UNIT"),ggQty.DecPoint,0)%>"	'�԰���� 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("CUR_CD"))%>"									'�԰�	
				strData = strData & Chr(11) & "<%=UniConvNumDBToCompanyWithOutChange(rs0("SUBCONTRACT_PRC"), 0)%>"	'���ִܰ� 
				strData = strData & Chr(11) & "<%=UniConvNumDBToCompanyWithOutChange(rs0("SUBCONTRACT_AMT"), 0)%>"	'���ֱݾ� 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("TAX_TYPE"))%>"								'VAT���� 
				strData = strData & Chr(11) & "<%=UniConvNumDBToCompanyWithOutChange(rs0("TAX_AMT"), 0)%>"		'VAT�ݾ� 
				strData = strData & Chr(11) & "<%=UniConvNumDBToCompanyWithOutChange(rs0("TOTAL_COST"), 0)%>"	'�ѱݾ� 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("WC_CD"))%>"			'�۾��� 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("WC_NM"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEM_CD"))%>"								'ǰ�� 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEM_NM"))%>"								'ǰ��� 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("SPEC"))%>"									'�԰�	
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("TRACKING_NO"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("OPR_NO"))%>"	'���� 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("SEQ"))%>"	'���� 
				
				strData = strData & Chr(11) & LngMaxRow + <%=i%>
				strData = strData & Chr(11) & Chr(12)
				
				TmpBuffer(<%=i%>) = strData
<%		
				rs0.MoveNext
			END IF
		Next
		
%>
		iTotalStr = Join(TmpBuffer, "")
		.ggoSpread.Source  = .frm1.vspdData2
		Call .ggoSpread.SSShowDataByClip(iTotalStr, "F")
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData2, LngMaxRow + 1 , LngMaxRow + <%=i%> ,.C_CurCd2,.C_SubContractPrc2, "C" ,"I","X","X")
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData2, LngMaxRow + 1 , LngMaxRow + <%=i%> ,.C_CurCd2,.C_SubcontractAmt2, "A" ,"I","X","X")
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData2, LngMaxRow + 1 , LngMaxRow + <%=i%> ,.C_CurCd2,.C_TaxAmt2, "A" ,"I","X","X")
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData2, LngMaxRow + 1 , LngMaxRow + <%=i%> ,.C_CurCd2,.C_TotalCost2, "A" ,"I","X","X")
		
		.lgStrPrevKey5 = "<%=ConvSPChars(rs0("prodt_order_no"))%>"
		.lgStrPrevKey6 = "<%=ConvSPChars(rs0("opr_no"))%>"
		.lgStrPrevKey7 = "<%=ConvSPChars(rs0("seq"))%>"
		     
<%	
	End If

	rs0.Close
	Set rs0 = Nothing

%>	
	.DbDtlQueryOk												'��: ��ȸ ������ ������� 

End With

</Script>	
<%
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
