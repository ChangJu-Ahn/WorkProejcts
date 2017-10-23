<%'======================================================================================================
'*  1. Module Name          : E-TAX
'*  2. Function Name        : 
'*  3. Program ID           : D4231mb1.asp
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
Dim strMode											'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

Call HideStatusWnd

Dim StrSupplierCd
Dim StrcbobillStatus
Dim StrhdtxtRadio
Dim StrcboTransferStatus
Dim strIssuedFromDt
Dim strIssuedToDt
Dim i

on error resume next
Err.Clear																	'��: Protect system from crashing

' Order Header Display
Redim UNISqlId(0)
Redim UNIValue(0, 11)

UNISqlId(0) = "D4231MA11"

strMode = Request("txtMode")											'�� : ���� ���¸� ���� 

UNIValue(0, 0) = "^"
UNIValue(0, 1) = "^"
UNIValue(0, 2) = FilterVar(UniConvDate(Request("txtIssuedFromDt")), "''", "S")
UNIValue(0, 3) = FilterVar(UniConvDate(Request("txtIssuedToDt")), "''", "S")


If Request("txtSupplierCd") = "" Then
	UNIValue(0, 4) = "|"
Else
	UNIValue(0, 4) = FilterVar(UCase(Request("txtSupplierCd")), "''", "S")
End If


UNIValue(0, 5) = "|"


If Request("txtBizAreaCd") = "" Then
	UNIValue(0, 6) = "%"
Else
	UNIValue(0, 6) = FilterVar(UCase(Request("txtBizAreaCd")), "''", "S")
End If

if Request("cboBillStatus") = "" Then
     UNIValue(0, 7) = "|"
Else
    If Request("cboBillStatus") = "X" Then
	    UNIValue(0, 7) = " is null "
    Else
	    UNIValue(0, 7) = "=" & FilterVar(UCase(Request("cboBillStatus")), "''", "S")
    End If
End IF    

If Request("rdoStatusFlag") = "A" Then
	UNIValue(0, 8) = "|"
Else
     if Request("rdoStatusFlag") = "D" Then
        UNIValue(0, 8) = "18"
     Else
        UNIValue(0, 8) = "01"
     end if  	
End If

If Request("txtTaxDocNo") = "" Then
	UNIValue(0, 9) = "|"
Else
	UNIValue(0, 9) = FilterVar(UCase(Request("txtTaxDocNo")) & "%", "''", "S") 
End If

If Request("txtTaxBillNo") = "" Then
	UNIValue(0, 10) = "|"
Else
	UNIValue(0, 10) = FilterVar(UCase(Request("txtTaxBillNo")) & "%", "''", "S")
End If

If Request("txtTempGLNo") = "" Then
	UNIValue(0, 11) = "|"
Else
	UNIValue(0, 11) = FilterVar(UCase(Request("txtTempGLNo")) & "%", "''", "S")
End If


UNILock = DISCONNREAD :	UNIFlag = "1"

Set ADF = Server.CreateObject("prjPublic.cCtlTake")
strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

If (rs0.EOF And rs0.BOF) Then
	Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
	rs0.Close
	Set rs0 = Nothing
	Set ADF = Nothing
	Response.End													'��: �����Ͻ� ���� ó���� ������ 
End If	%>

<Script Language=vbscript>
	Dim LngMaxRow
	Dim strData
	Dim TmpBuffer
	Dim iTotalStr1
	Dim aaa

	With parent																	'��: ȭ�� ó�� ASP �� ��Ī�� 
		LngMaxRow = .frm1.vspdData.MaxRows										'Save previous Maxrow
		ReDim TmpBuffer(<%=rs0.RecordCount - 1%>)

<%		Dim iDx
		For iDx = 0 to rs0.RecordCount - 1 %>
			aaa = <%=iDx%>
				
			strData = ""
			strData = strData & Chr(11) & ""
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("where_flag"))%>"			
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("dti_wdate"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("conversation_id"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("tax_bill_no"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("tax_doc_no"))%>"	
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("temp_gl_no"))%>"			
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("dti_status"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("dti_status_nm"))%>"			
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("tax_bill_type_nm"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("amend_code"))%>"
			strData = strData & Chr(11) & ""
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("amend_code_nm"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("bp_cd"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("bp_nm"))%>"						
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("byr_emp_name"))%>"
			strData = strData & Chr(11) & ""			
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("byr_dept_name"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("byr_tel_num"))%>"			
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("byr_email"))%>"
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("issued_dt"))%>"			
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("vat_calc_type_nm"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("vat_inc_flag_nm"))%>"			
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("vat_type"))%>"			
            strData = strData & Chr(11) & "<%=ConvSPChars(rs0("vat_type_nm"))%>"
            strData = strData & Chr(11) & "<%=UNIConvNumDBToCompanyByCurrency(rs0("vat_rate"), gCurrency, ggAmtOfMoneyNo, "X", "X")%>"
            strData = strData & Chr(11) & "<%=ConvSPChars(rs0("cur"))%>"            
            strData = strData & Chr(11) & "<%=UNIConvNumDBToCompanyByCurrency(rs0("total_amt"), gCurrency, ggAmtOfMoneyNo, "X", "X")%>"
            strData = strData & Chr(11) & "0"    
			strData = strData & Chr(11) & "<%=UNIConvNumDBToCompanyByCurrency(rs0("net_amt"), gCurrency, ggAmtOfMoneyNo, "X", "X")%>"
			strData = strData & Chr(11) & "0"    
			strData = strData & Chr(11) & "<%=UNIConvNumDBToCompanyByCurrency(rs0("vat_amt"), gCurrency, ggAmtOfMoneyNo, "X", "X")%>"
			strData = strData & Chr(11) & "0"    
			strData = strData & Chr(11) & "<%=UNIConvNumDBToCompanyByCurrency(rs0("total_loc_amt"), gCurrency, ggAmtOfMoneyNo, "X", "X")%>"
			strData = strData & Chr(11) & "0"    
			strData = strData & Chr(11) & "<%=UNIConvNumDBToCompanyByCurrency(rs0("net_loc_amt"), gCurrency, ggAmtOfMoneyNo, "X", "X")%>"
			strData = strData & Chr(11) & "0"    
			strData = strData & Chr(11) & "<%=UNIConvNumDBToCompanyByCurrency(rs0("vat_loc_amt"), gCurrency, ggAmtOfMoneyNo, "X", "X")%>"
			strData = strData & Chr(11) & "0"    
            strData = strData & Chr(11) & "<%=ConvSPChars(rs0("report_biz_area"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("tax_biz_area_nm"))%>"
            strData = strData & Chr(11) & "<%=ConvSPChars(rs0("sales_grp"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("sales_grp_nm"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("remark"))%>"  			                      																								
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
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
