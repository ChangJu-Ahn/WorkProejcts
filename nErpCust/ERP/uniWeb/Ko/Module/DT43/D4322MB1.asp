<%'======================================================================================================
'*  1. Module Name          : E-TAX
'*  2. Function Name        : 
'*  3. Program ID           : D4211mb1.asp
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
Dim strMode											'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

Call HideStatusWnd

Dim StrSupplierCd
Dim StrcbobillStatus
Dim StrhdtxtRadio
Dim StrcboTransferStatus
Dim strIssuedFromDt
Dim strIssuedToDt
Dim i

On Error Resume Next
Err.Clear																	'☜: Protect system from crashing

' Order Header Display
Redim UNISqlId(0)
Redim UNIValue(0, 6)


UNISqlId(0) = "D4322MA101"

strMode = Request("txtMode")											'☜ : 현재 상태를 받음 

UNIValue(0, 0) = FilterVar(UniConvDate(Request("txtIssuedFromDt")), "''", "S")
UNIValue(0, 1) = FilterVar(UniConvDate(Request("txtIssuedToDt")), "''", "S")
UNIValue(0, 2) = FilterVar(UCase(Request("txtBizAreaCd")), "''", "S")

If Request("txtSalesGrpCd") = "" Then
	UNIValue(0, 3) = "'%'"
Else
	UNIValue(0, 3) = FilterVar(UCase(Request("txtSalesGrpCd")), "''", "S")
End If


If Request("txtSupplierCd") = "" Then
	UNIValue(0, 4) = "'%'"
Else
	UNIValue(0, 4) = FilterVar(UCase(Request("txtSupplierCd")), "''", "S")
End If


if Request("cboBillStatus") = "" Then
     UNIValue(0, 5) = "'%'"
Else
    UNIValue(0, 5) =  FilterVar(UCase(Request("cboBillStatus")), "''", "S")
End IF    

if Request("cboAmendCode") = "" Then
     UNIValue(0, 6) = "'%'"
Else   
     UNIValue(0, 6) = FilterVar(UCase(Request("cboAmendCode")), "''", "S")
End IF    


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
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("dti_wdate"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("conversation_id"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(replace(rs0("SBDESCRIPTION"),Chr(13) & Chr(10), ""))%>" 			
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("iv_no"))%>"						
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("posted_flg"))%>"	
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("dti_status"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("dti_status_nm"))%>"								
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("build_cd"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("bp_nm"))%>"	
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("iv_cur"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("amend_code"))%>"            
            strData = strData & Chr(11) & "<%=ConvSPChars(rs0("amend_code_nm"))%>"            
            strData = strData & Chr(11) & "<%=UNIConvNumDBToCompanyByCurrency(rs0("net_doc_amt"), gCurrency, ggAmtOfMoneyNo, "X", "X")%>"
			strData = strData & Chr(11) & "<%=UNIConvNumDBToCompanyByCurrency(rs0("fi_net_amt"), gCurrency, ggAmtOfMoneyNo, "X", "X")%>"						
			strData = strData & Chr(11) & "<%=UNIConvNumDBToCompanyByCurrency(rs0("tot_vat_doc_amt"), gCurrency, ggAmtOfMoneyNo, "X", "X")%>"
			strData = strData & Chr(11) & "<%=UNIConvNumDBToCompanyByCurrency(rs0("fi_vat_amt"), gCurrency, ggAmtOfMoneyNo, "X", "X")%>"						
			strData = strData & Chr(11) & "<%=UNIConvNumDBToCompanyByCurrency(rs0("total_amt"), gCurrency, ggAmtOfMoneyNo, "X", "X")%>"			
			strData = strData & Chr(11) & "<%=UNIConvNumDBToCompanyByCurrency(rs0("fi_total_amt"), gCurrency, ggAmtOfMoneyNo, "X", "X")%>"			
			strData = strData & Chr(11) & "<%=UNIConvNumDBToCompanyByCurrency(rs0("net_loc_amt"), gCurrency, ggAmtOfMoneyNo, "X", "X")%>"
			strData = strData & Chr(11) & "<%=UNIConvNumDBToCompanyByCurrency(rs0("fi_net_loc_amt"), gCurrency, ggAmtOfMoneyNo, "X", "X")%>"						
			strData = strData & Chr(11) & "<%=UNIConvNumDBToCompanyByCurrency(rs0("tot_vat_loc_amt"), gCurrency, ggAmtOfMoneyNo, "X", "X")%>"
			strData = strData & Chr(11) & "<%=UNIConvNumDBToCompanyByCurrency(rs0("fi_vat_loc_amt"), gCurrency, ggAmtOfMoneyNo, "X", "X")%>" 			           			
			strData = strData & Chr(11) & "<%=UNIConvNumDBToCompanyByCurrency(rs0("total_loc_amt"), gCurrency, ggAmtOfMoneyNo, "X", "X")%>"			
			strData = strData & Chr(11) & "<%=UNIConvNumDBToCompanyByCurrency(rs0("fi_total_loc_amt"), gCurrency, ggAmtOfMoneyNo, "X", "X")%>"
            strData = strData & Chr(11) & "<%=ConvSPChars(rs0("vat_inc_flag"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("vat_inc_flag_nm"))%>"			
            strData = strData & Chr(11) & "<%=ConvSPChars(rs0("vat_type"))%>"			
            strData = strData & Chr(11) & "<%=ConvSPChars(rs0("vat_type_nm"))%>"
            strData = strData & Chr(11) & "<%=UNIConvNumDBToCompanyByCurrency(rs0("vat_rt"), gCurrency, ggAmtOfMoneyNo, "X", "X")%>"                        
            strData = strData & Chr(11) & "<%=ConvSPChars(rs0("sup_emp_name"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("sup_dept_name"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("sup_tel_num"))%>"			
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("sup_email"))%>"                        
            strData = strData & Chr(11) & "<%=ConvSPChars(rs0("tax_biz_area"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("tax_biz_area_nm"))%>"            
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("pur_grp"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("pur_grp_nm"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("remark"))%>"			
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("vat_calc_type"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("vat_calc_type_nm"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("issue_dt_flag"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("where_flag"))%>"														
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("issued_dt"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("return_code"))%>" 
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("return_description"))%>"						            															                                                			                        																								                                                			                        
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
