<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
'********************************************************************************************************
'*  1. Module Name          : Procuremant																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m4212mb1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : 수입통관 내역등록 Query Transaction 처리용 ASP							*
'*  7. Modified date(First) : 2000/03/22																*
'*  8. Modified date(Last)  : 2000/03/22																*
'*  9. Modifier (First)     :																			*
'* 10. Modifier (Last)      :																			*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/22 : Coding Start												*
'********************************************************************************************************

	Dim lgOpModeCRUD
	Dim lgCurrency
	
	On Error Resume Next					'☜: Protect system from crashing
	Err.Clear 						'☜: Clear Error status
				
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*","NOCOOKIE", "MB")
	Call LoadBNumericFormatB("I", "*","NOCOOKIE", "MB")
		
	Call HideStatusWnd

	lgOpModeCRUD	=	Request("txtMode")
				
	Select Case lgOpModeCRUD
		Case CStr(UID_M0001)                                                         '☜: Query
		'**수정(2003.06.11)-추가 100건을 조회시는 헤더는 조회하지 않는다**
		If Request("txtMaxRows") = 0 Then
			Call SubBizQueryMulti_CC_HDR()
		Else
			lgCurrency = request("txtCurrency")
		End If
			
		Call SubBizQueryMulti_CC_LAN()
	End Select

'============================================================================================================
' Name : SubBizQueryMulti_CC_HDR
' Desc : Query Data from Db - Dev. Defined 
'============================================================================================================
Sub SubBizQueryMulti_CC_HDR()
		
	Dim OBJ_PM6G119																' Master L/C Header 조회용 Object

	Dim E1_m_cc_hdr 
    Const M418_E1_cc_no = 0    
	Const M418_E1_id_no = 1
	Const M418_E1_ip_no = 2
	Const M418_E1_id_dt = 13
	Const M418_E1_currency = 38
	Const M418_E1_doc_amt = 39
	Const M418_E1_ip_dt = 51
	Const M418_E1_usd_xch_rate = 66
	Const M418_E1_pur_grp = 101
	Const M418_E1_beneficiary = 105
	Const M418_E1_beneficiary_nm = 106
	
	Dim strCcNo	
	On Error Resume Next
	Err.Clear																'☜: Protect system from crashing
		
    '---------------------------------- 통관 Header Data Query ----------------------------------
		
	If Request("txtCCNo") = "" Then									
		Call DisplayMsgBox("700112", vbInformation,	"", "",	I_MKSCRIPT)
		Exit Sub 
	End If
	
	Set OBJ_PM6G119 = Server.CreateObject("PM6G119.cMLkImportCcHdrS")
		
	If CheckSYSTEMError(Err,True) = True Then
		Response.End
		'Exit Sub
	End If
	
	strCcNo = Trim(Request("txtCCNo"))
	Call OBJ_PM6G119.M_LOOKUP_IMPORT_CC_HDR_SVR(gStrGlobalCollection, strCcNo, E1_m_cc_hdr)

	If CheckSYSTEMError(Err,True) = True Then
		Set OBJ_PM6G119 = Nothing
		Response.End
	End If
	
	Set OBJ_PM6G119 = Nothing
	'-----------------------
	'Result data display area
	'-----------------------
	lgCurrency = ConvSPChars(E1_m_cc_hdr(M418_E1_currency))

	Response.Write "<Script Language=VBScript>" & vbCr
	Response.Write "With parent.frm1" & vbCr

	'========= TAB 1 (수입신고) ==========

	'##### Rounding Logic #####
	'항상 거래화폐가 우선 
	Response.Write ".txtCurrency.value = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_currency)) & """" & vbCr
	Response.Write " parent.CurFormatNumericOCX" & vbCr
	Response.Write ".txtIDNo.value = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_id_no)) & """" & vbCr				'신고번호 
	Response.Write ".txtIDDt.Text = """ & UNIDateClientFormat(E1_m_cc_hdr(M418_E1_id_dt)) & """" & vbCr			'신고일 
	Response.Write ".txtIPNo.value = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_ip_no)) & """" & vbCr				'면허번호 
	Response.Write ".txtIPDt.Text = """ & UNIDateClientFormat(E1_m_cc_hdr(M418_E1_ip_dt)) & """" & vbCr			'면허일 
	Response.Write "If """ & E1_m_cc_hdr(M418_E1_doc_amt) & """ = 0 Then " & vbCr								'통관금액 
	Response.Write "	.txtDocAmt.Text = """ & E1_m_cc_hdr(M418_E1_doc_amt) & """" & vbCr
	Response.Write "Else" & vbCr
	Response.Write "	.txtDocAmt.text = """ & UNIConvNumDBToCompanyByCurrency(E1_m_cc_hdr(M418_E1_doc_amt),lgCurrency,ggAmtOfMoneyNo,"X","X") & """" & vbCr
	Response.Write "End If" & vbCr		
	Response.Write ".txtBeneficiary.value = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_beneficiary)) & """" & vbCr		'수출자 
	Response.Write ".txtBeneficiaryNm.value = """ & ConvSPChars(E1_m_cc_hdr(M418_E1_beneficiary_nm)) & """" & vbCr	'수출자명 
	Response.Write ".hdnXchRt.value = """ & UNINumClientFormat(E1_m_cc_hdr(M418_E1_usd_xch_rate), ggExchRate.DecPoint, 0) & """" & vbCr	'USD환율 
		
	'-----------------------
	' Rounding Column Set               '☜:화폐별 라운딩 스프레드 포매팅 
	'----------------------- 		
	Response.Write "If Trim(parent.lgIntFlgMode) = Trim(parent.parent.OPMD_CMODE) Then parent.CurFormatNumSprSheet() " & vbCr             
	Response.Write "parent.lgIntFlgMode = """ & OPMD_UMODE & """ " & vbCr
	Response.Write "Call parent.DbQueryOk()	" & vbCr												'☜: 조회가 성공 
	Response.Write "End With" & vbCr
	Response.Write "</Script>" & vbCr
	
End Sub

'============================================================================================================
' Name : SubBizQueryMulti_CC_LAN
' Desc : Query Data from Db - Dev. Defined 
'============================================================================================================
Sub SubBizQueryMulti_CC_LAN()

	Dim OBJ_PM6G138
	Dim PvArr
	Dim lGrpCnt
	Dim iTotstrData
		
	Dim E1_m_cc_lan_lan_no
		
	Dim EG1_export_group 
	Const M385_EG1_E1_m_cc_lan_lan_no = 0
	Const M385_EG1_E1_m_cc_lan_hs_cd = 1
	Const M385_EG1_E1_m_cc_lan_qty = 2
	Const M385_EG1_E1_m_cc_lan_doc_amt = 3
	Const M385_EG1_E1_m_cc_lan_cif_doc_amt = 4
	Const M385_EG1_E1_m_cc_lan_loc_amt = 5
	Const M385_EG1_E1_m_cc_lan_cif_loc_amt = 6
	Const M385_EG1_E1_m_cc_lan_unit = 7
	Const M385_EG1_E1_m_cc_lan_net_weight = 8
	Const M385_EG1_E1_m_cc_lan_tariff_rate = 9
	Const M385_EG1_E1_m_cc_lan_redu_rate = 10
	Const M385_EG1_E1_m_cc_lan_tax_loc_amt = 11
	Const M385_EG1_E1_m_cc_lan_insrt_user_id = 12
	Const M385_EG1_E1_m_cc_lan_insrt_dt = 13
	Const M385_EG1_E1_m_cc_lan_updt_user_id = 14
	Const M385_EG1_E1_m_cc_lan_updt_dt = 15
	Const M385_EG1_E1_m_cc_lan_ext1_qty = 16
	Const M385_EG1_E1_m_cc_lan_ext1_amt = 17
	Const M385_EG1_E1_m_cc_lan_ext1_cd = 18
	Const M385_EG1_E2_b_hs_code_hs_nm = 19
    
		
	Dim iLngMaxRow
	Dim iLngRow
	Dim iStrPrevKey
	Dim istrData
	Dim istrTemp
	Dim iStrNextKey  	
	Dim iarrValue
	Dim iStrCcNo
	
	Const C_SHEETMAXROWS_D  = 100
		
	'---------------------------------- 통관 Detail Data Query ----------------------------------
		
	Set OBJ_PM6G138 = Server.CreateObject("PM6G138.cMListImportCcLanS")
		
	If CheckSYSTEMError(Err,True) = True Then
		Set OBJ_PM6G138 = Nothing
		Exit Sub
	End If													
	
	iStrCcNo = Trim(Request("txtCCNo"))
	iStrPrevKey = UNIConvNum(Request("lgStrPrevKey"),0)
		
	call OBJ_PM6G138.M_LIST_IMPORT_CC_LAN_SVR(gStrGlobalCollection,C_SHEETMAXROWS_D, iStrCcNo, iStrPrevKey, _
                                               E1_m_cc_lan_lan_no,EG1_export_group)

	
	If CheckSYSTEMError2(Err,True,"","","","","") = True Then
		Set OBJ_PM6G138 = Nothing
		Response.Write "<Script Language=vbscript>" & vbCr
		Response.Write "parent.dbQueryOk" & chr(13)
		Response.Write "</Script>"
		Exit Sub
	End If
		
    Set OBJ_PM6G138 = Nothing														'☜: Unload Comproxy
        
	iLngMaxRow = CLng(Request("txtMaxRows"))
	ReDim PvArr(UBound(Eg1_export_group,1))	

	lGrpCnt = 0
	
	For iLngRow = 0 To UBound(EG1_export_group,1)
	
		If  iLngRow < C_SHEETMAXROWS_D  Then
		Else
		   iStrNextKey = ConvSPChars(EG1_export_group(iLngRow,M385_EG1_E1_m_cc_lan_lan_no)) 
           Exit For
        End If  	
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M385_EG1_E1_m_cc_lan_lan_no)) _
		                    & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M385_EG1_E1_m_cc_lan_hs_cd)) _
		                    & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M385_EG1_E2_b_hs_code_hs_nm)) _
		                    & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M385_EG1_E1_m_cc_lan_unit)) _
		                    & Chr(11) & UNIConvNumDBToCompanyByCurrency (EG1_export_group(iLngRow,M385_EG1_E1_m_cc_lan_cif_doc_amt),lgCurrency,ggAmtOfMoneyNo,"X","X") _
		                    & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow,M385_EG1_E1_m_cc_lan_cif_loc_amt), ggAmtOfMoney.DecPoint, 0) _
		                    & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow,M385_EG1_E1_m_cc_lan_tariff_rate), ggExchRate.DecPoint, 0) _
		                    & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow,M385_EG1_E1_m_cc_lan_redu_rate), ggExchRate.DecPoint, 0)
		If UNINumClientFormat(EG1_export_group(iLngRow,M385_EG1_E1_m_cc_lan_tax_loc_amt), ggAmtOfMoney.DecPoint, 0) = 0 Then												'세액	
			istrData = istrData & Chr(11) & "0"
		Else
			istrData = istrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow,M385_EG1_E1_m_cc_lan_tax_loc_amt), ggAmtOfMoney.DecPoint, 0)
		End If
		istrData = istrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow,M385_EG1_E1_m_cc_lan_net_weight), ggQty.DecPoint, 0) _
		                    & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow,M385_EG1_E1_m_cc_lan_qty), ggQty.DecPoint, 0) _
		                    & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group(iLngRow,M385_EG1_E1_m_cc_lan_doc_amt),lgCurrency,ggAmtOfMoneyNo,"X","X") _
		                    & Chr(11) & iLngMaxRow + iLngRow _
		                    & Chr(11) & Chr(12)
		
		PvArr(lGrpCnt) = istrData
		lGrpCnt = lGrpCnt + 1
		istrData = ""
    Next
    
    iTotstrData = Join(PvArr, "")
	    
	Response.Write "<Script Language=VBScript>"					& vbCr
	Response.Write "With parent"								& vbCr
	Response.Write "	.SetSpreadColor -1,-1	"				& vbCr
	Response.Write "	.ggoSpread.Source = .frm1.vspdData "								& vbCr
	Response.Write "	.ggoSpread.SSShowData		""" & iTotstrData				& """ " & vbCr
	Response.Write "		.lgStrPrevKey =	""" & iStrNextKey						& """ " & vbCr
	Response.Write "		.frm1.txtHCCNo.value = """ & ConvSPChars(Request("txtCCNo"))		& """ "	& vbCr
	Response.Write "		.DbQueryOk "		& vbCr 
	Response.Write "End With"					& vbCr
	Response.Write "</Script>"					& vbCr
End Sub

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
			
End Sub

%>

