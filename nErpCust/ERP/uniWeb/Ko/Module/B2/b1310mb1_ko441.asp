<%@ LANGUAGE=VBSCript%>
<%Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../comasp/loadinftb19029.asp" -->
<%

Call LoadBasisGlobalInf() 
Call LoadInfTB19029B("I","*","NOCOOKIE", "MB")
Call HideStatusWnd

On Error Resume Next

Dim strMode																	'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

strMode = Request("txtMode")												'�� : ���� ���¸� ���� 


If strMode = "" Then
	Response.End 
	Call HideStatusWnd
ElseIf strMode <> CStr(UID_M0001) Then											'��: ��ȸ ���� Biz �̹Ƿ� �ٸ����� �׳� ������ 
	Call DisplayMsgBox("700118", vbOKOnly, "", "", I_MKSCRIPT)	'��ȸ��û�� �� �� �ֽ��ϴ�.
	Response.End 
	Call HideStatusWnd
ElseIf Request("txtBankCd") = "" Then												'��: ��ȸ�� ���� ���� ���Դ��� üũ 
	Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)	'��ȸ ���ǰ��� ����ֽ��ϴ�!
	Response.End 
	Call HideStatusWnd
End If


Dim pPB2SA05 														'�� : ��ȸ�� ComProxy Dll ��� ���� 
Dim IntRows
Dim GroupCount
Dim LngMaxRow
Dim StrNextKey
Dim lgStrPrevKey

' Com+ Conv. ���� ���� 
Dim pvStrGlobalCollection 
    
Dim import_next_b_bank_acct
Dim import_b_bank
    
Dim export_par_b_bank 
Dim export_b_country 
Dim export_next_b_bank_acct 
Dim export_group 
Dim export_b_bank 
Dim iIntQueryCount
Const C_SHEETMAXROWS_D  = 100
dim iStrPrevKey

Dim arrCount
DIM iIntLoopCount

' ÷�� ���� 
    Const C_import_b_bank_bank_cd = 0

    Const C_import_next_b_bank_acct_bank_acct_no = 0

    Const C_export_b_bank_bank_cd = 0
    Const C_export_b_bank_bank_nm = 1
    Const C_export_b_bank_bank_full_nm = 2
    Const C_export_b_bank_bank_eng_nm = 3
    Const C_export_b_bank_zip_cd = 4
    Const C_export_b_bank_addr1 = 5
    Const C_export_b_bank_addr2 = 6
    Const C_export_b_bank_addr3 = 7
    Const C_export_b_bank_eng_addr1 = 8
    Const C_export_b_bank_eng_addr2 = 9
    Const C_export_b_bank_eng_addr3 = 10
    Const C_export_b_bank_bank_type = 11
    Const C_export_b_bank_country_cd = 12
    Const C_export_b_bank_par_bank_cd = 13
    Const C_export_b_bank_bank_fg = 14
    Const C_export_b_bank_addr4 = 15

    Const C_export_b_country_country_nm = 0
    Const C_export_b_country_country_cd = 1

    Const C_export_next_b_bank_acct_bank_acct_no = 0

    Const C_export_par_b_bank_bank_cd = 0
    Const C_export_par_b_bank_bank_nm = 1

    Const C_export_group_export_item_b_bank_acct_bank_acct_no = 0
    Const C_export_group_export_item_b_bank_acct_bank_acct_type = 1
    Const C_export_group_export_item_b_bank_acct_dpst_type = 2
    Const C_export_group_export_item_b_bank_c_acct_use = 3
    Const C_export_group_export_item_b_bank_acct_bp_cd = 4
    Const C_export_group_export_item_b_bank_acct_limit_amt = 5
    Const C_export_group_export_item_b_biz_partner_bp_cd = 6
    Const C_export_group_export_item_b_biz_partner_bp_nm = 7
    Const C_export_group_export_item_b_bank_c_acct_prnt = 8	'>>air ����¿���    

lgStrPrevKey = Request("lgStrPrevKey")

Set pPB2SA05 = Server.CreateObject("PB2SA05_KO441.cBListBankSvr")	    	    

'-----------------------
'Com action result check area(OS,internal)
'-----------------------

If Err.Number <> 0 Then
	Set pPB2SA05 = Nothing													'��: ComProxy Unload
	Call ServerMesgBox(Err.description , vbInformation, I_MKSCRIPT)			'��:
	Response.End															'��: �����Ͻ� ���� ó���� ������ 
	Call HideStatusWnd
End If

'-----------------------
'Data manipulate  area(import view match)
'-----------------------

'Redim import_b_bank(C_import_b_bank_bank_cd)
'ReDim import_next_b_bank_acct(C_import_next_b_bank_acct_bank_acct_no)

'import_b_bank(0) = Trim(Request("txtBankCd"))
'import_next_b_bank_acct(0) = lgStrPrevKey

import_b_bank = Request("txtBankCd")
import_next_b_bank_acct = lgStrPrevKey

Call pPB2SA05.B_LIST_BANK_SVR(gStrGlobalCollection,C_SHEETMAXROWS_D, CStr(import_b_bank), CStr(import_next_b_bank_acct),export_b_bank,  export_b_country,export_par_b_bank ,  export_group, export_next_b_bank_acct)
Set pPB2SA05 = Nothing

'-----------------------
'Com action result check area(OS,internal)
'-----------------------
	If CheckSYSTEMError(Err,True) = True Then
		Set pPB2SA05 = Nothing																	'��: ComProxy Unload
		Response.End																			'��: �����Ͻ� ���� ó���� ������ 
	End If



	IF Trim(lgStrPrevKey)=  "" then
%>

<Script Language=vbscript>
	With parent.frm1																	'��: ȭ�� ó�� ASP �� ��Ī�� 
		.txtBankCd.Value		= "<%=ConvSPChars(export_b_bank(C_export_b_bank_bank_cd))%>"
		.txtBankCd1.Value		= "<%=ConvSPChars(export_b_bank(C_export_b_bank_bank_cd))%>"
		.txtBankNm.value		= "<%=ConvSPChars(export_b_bank(C_export_b_bank_bank_nm))%>"
		.txtBankShNm.Value		= "<%=ConvSPChars(export_b_bank(C_export_b_bank_bank_nm))%>"
		.txtBankFullNm.Value	= "<%=ConvSPChars(export_b_bank(C_export_b_bank_bank_full_nm))%>"
		.cboBankType.Value	    = "<%=ConvSPChars(export_b_bank(C_export_b_bank_bank_type))%>"
		.txtBankEngNm.Value		= "<%=ConvSPChars(export_b_bank(C_export_b_bank_bank_eng_nm))%>"
		.txtCountryCd.Value		= "<%=ConvSPChars(export_b_country(C_export_b_country_country_cd))%>"
		.txtCountryNm.Value		= "<%=ConvSPChars(export_b_country(C_export_b_country_country_nm))%>"
		.txtZipCd.Value			= "<%=ConvSPChars(export_b_bank(C_export_b_bank_zip_cd))%>"
		.txtAddr1.Value		    = "<%=ConvSPChars(export_b_bank(C_export_b_bank_addr1))%>"
		.txtAddr2.Value			= "<%=ConvSPChars(export_b_bank(C_export_b_bank_addr2))%>"
		.txtAddr3.Value		    = "<%=ConvSPChars(export_b_bank(C_export_b_bank_addr3))%>"
		.txtEngAddr1.Value	    = "<%=ConvSPChars(export_b_bank(C_export_b_bank_eng_addr1))%>"
		.txtEngAddr2.Value	    = "<%=ConvSPChars(export_b_bank(C_export_b_bank_eng_addr2))%>"
		.txtEngAddr3.Value		= "<%=ConvSPChars(export_b_bank(C_export_b_bank_eng_addr3))%>"
		
		.hBankCd.value = "<%=ConvSPChars(export_b_bank(C_export_b_bank_bank_cd))%>"
	End With
	
	Call parent.DbQueryOk
</Script>

<%
end if

Response.Write Err.Description

If CheckSYSTEMError(Err,True) = True Then

	Set pPB2SA05 = Nothing																	'��: ComProxy Unload
	Response.End																			'��: �����Ͻ� ���� ó���� ������	
End If

GroupCount = 0

If IsEmpty(export_group) = False and IsArray(export_group) = True Then    

	GroupCount = UBound(export_group,1)
	
	If GroupCount > 0 Then
		If Trim(export_group(GroupCount,C_export_group_export_item_b_bank_acct_bank_acct_no)) = Trim(import_next_b_bank_acct) Then
			StrNextKey = ""

		Else
			StrNextKey = Trim(import_next_b_bank_acct)
		End If
		
		If CheckSYSTEMError(Err,True) = True Then

			Set pPB2SA05 = Nothing																	'��: ComProxy Unload
			Response.End																			'��: �����Ͻ� ���� ó���� ������	
		End If
		
	End If
   
%>
<Script Language=vbscript>
	Dim LngMaxRow, strData
	
	With parent
	
		LngMaxRow = .frm1.vspdData.MaxRows
<%
	For arrCount = 0 To GroupCount		
		iIntLoopCount = iIntLoopCount + 1
	    
	    If  iIntLoopCount < (C_SHEETMAXROWS_D + 1) Then

'			For arrCount = 0 To GroupCount

%>
				strData = strData & Chr(11) & "<%=ConvSPChars(export_group(arrCount,C_export_group_export_item_b_bank_acct_bank_acct_no))%>"
				strData = strData & Chr(11) & "<%=export_group(arrCount,C_export_group_export_item_b_bank_acct_bank_acct_type)%>"
				strData = strData & Chr(11) & ""
				strData = strData & Chr(11) & "<%=export_group(arrCount,C_export_group_export_item_b_bank_acct_dpst_type)%>"
				strData = strData & Chr(11) & ""
				strData = strData & Chr(11) & "<%=ConvSPChars(Trim(export_group(arrCount,C_export_group_export_item_b_bank_c_acct_use)))%>"
				strData = strData & Chr(11) & ""
				strData = strData & Chr(11) & "<%=ConvSPChars(Trim(export_group(arrCount,C_export_group_export_item_b_bank_c_acct_prnt)))%>"	'>>air ����¿���				
				strData = strData & Chr(11) & ""
				strData = strData & Chr(11) & "<%=ConvSPChars(Trim(export_group(arrCount,C_export_group_export_item_b_bank_acct_bp_cd)))%>"
				strData = strData & Chr(11) & ""
				strData = strData & Chr(11) & "<%=ConvSPChars(export_group(arrCount,C_export_group_export_item_b_biz_partner_bp_nm))%>"
				strData = strData & Chr(11) & "<%=UNINumClientFormat(export_group(arrCount,C_export_group_export_item_b_bank_acct_limit_amt), ggAmtOfMoney.DecPoint, 0)%>"
				strData = strData & Chr(11) & LngMaxRow + <%=arrCount%>  + 1                               '11
				strData = strData & Chr(11) & Chr(12)
			
<%
'			Next 
	    Else
			iStrPrevKey = Export_group(UBound(Export_group, 1), 0)
			iIntQueryCount = iIntQueryCount + 1
			Exit For
			  
		End If
	Next
	If  iIntLoopCount < (C_SHEETMAXROWS_D + 1) Then
		iStrPrevKey = ""
	    iIntQueryCount = ""
	End If
%>

		.ggoSpread.Source = .frm1.vspdData 
		.ggoSpread.SSShowData strData

		.lgStrPrevKey = "<%=ConvSPChars(iStrPrevKey)%>"

'
'        If .frm1.vspdData.MaxRows < .C_SHEETMAXROWS And .lgStrPrevKey <> "" Then
'            .DbQuery
'        Else
''           .frm1.hBankCd.value = "<%=Request("txtBankCd")%>"
            .DbQueryOk
'        End If

	End With	
</Script>
<%  
End If
    Set pPB2SA05 = Nothing
    Response.End
%>    
