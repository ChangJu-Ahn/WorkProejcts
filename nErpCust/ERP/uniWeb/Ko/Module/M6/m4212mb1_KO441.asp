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
'*  8. Modified date(Last)  : 2003/06/03																*
'*  9. Modifier (First)     : Sun-joung Lee																*
'* 10. Modifier (Last)      : Lee Eun Hee																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/22 : Coding Start
'*							  2. 2002/06/21 : Com Plus Conv.											*
'*							  3. 2003/04/28 : 성능개선													*
'********************************************************************************************************
Dim lgOpModeCRUD
Dim lgCurrency
	
On Error Resume Next																	'☜: Protect system from crashing
Err.Clear 																				'☜: Clear Error status

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*","NOCOOKIE", "MB")
Call LoadBNumericFormatB("I", "*","NOCOOKIE", "MB")

Call HideStatusWnd

lgOpModeCRUD	=	Request("txtMode")													'☜: Read Operation Mode (CRUD)
				
Select Case lgOpModeCRUD
    Case CStr(UID_M0001)														'☜: Query
		'**수정(2003.06.11)-추가 100건을 조회시는 헤더는 조회하지 않는다**
		If Request("txtMaxRows") = 0 Then
		
			Call SubBizQueryMulti_CC_HDR()
		Else
			lgCurrency = request("txtCurrency")
		End If	
		Call SubBizQueryMulti_CC_DTL()
    Case CStr(UID_M0002)
    Case CStr(UID_M0003)															'☜: Delete
End Select

'============================================================================================================
' Name : SubBizQueryMulti_CC_HDR
' Desc : Query Data from Db - Dev.Defined
'============================================================================================================
Sub SubBizQueryMulti_CC_HDR()
	
	Dim OBJ_PM6G119
		
	Dim L_E1_m_cc_hdr
	'ReDim L_E1_m_cc_hdr, 상수 인덱스 유지할것.
	Const E1_id_no = 1
	Const E1_ip_no = 2
	Const E1_id_dt = 13
	Const E1_currency = 38
	Const E1_doc_amt= 39
	Const E1_xch_rate = 41
	Const E1_ip_dt = 51
	Const E1_pur_grp = 101
	Const E1_beneficiary = 105
	Const E1_beneficiary_nm = 106
	
	Dim iStrCcNo
	
	On Error Resume Next
	Err.Clear																'☜: Protect system from crashing
	
	Set OBJ_PM6G119 = Server.CreateObject("PM6G119.cMLkImportCcHdrS")

	If CheckSYSTEMError(Err,True) = True Then
		Response.End
		'Exit Sub
	End If
	
	iStrCcNo = UCase(Trim(Request("txtCCNo")))
		
	Call OBJ_PM6G119.M_LOOKUP_IMPORT_CC_HDR_SVR(gStrGlobalCollection, iStrCcNo, L_E1_m_cc_hdr)
				 
	If CheckSYSTEMError2(Err,True,"","","","","") = true then 
			
		Set OBJ_PM6G119 = Nothing												
		Response.Write "<Script Language=vbscript>" & vbCr
		Response.write " parent.frm1.vspdData.MaxRows = 0 " & vbCr
		Response.Write " parent.dbQueryOk " & vbCr
		Response.Write "</Script>"
		Response.End
																
	End If
		

	Set OBJ_PM6G119 = Nothing									'☜: ComProxy UnLoad
	
	'-----------------------
	'Result data display area
	'-----------------------
	lgCurrency = ConvSPChars(L_E1_m_cc_hdr(E1_currency))
	
	Response.Write "<Script Language=VBScript>" & vbCr
	Response.Write "With parent.frm1"		 & vbCr
	
	'##### Rounding Logic #####
	'항상 거래화폐가 우선 
	Response.Write ".txtCurrency.value		  = """ & ConvSPChars(L_E1_m_cc_hdr(E1_currency))	& """" & vbCr
	Response.Write "parent.CurFormatNumericOCX" & vbCr
	'##########################
	Response.Write ".txtHCCNo.value			= """ & ConvSPChars(Request("txtCCNo"))				& """" & vbCr
	Response.Write ".txtIDNo.value			= """ & ConvSPChars(L_E1_m_cc_hdr(E1_id_no))		& """" & vbCr		'신고번호 
																			
	Response.Write ".txtIDDt.text			= """ & UNIDateClientFormat(L_E1_m_cc_hdr(E1_id_dt)) & """" & vbCr		'신고일 
	Response.Write ".txtIPNo.value			= """ & ConvSPChars(L_E1_m_cc_hdr(E1_ip_no))		& """" & vbCr		'면허번호 
	Response.Write ".txtIPDt.text			= """ & UNIDateClientFormat(L_E1_m_cc_hdr(E1_ip_dt)) & """" & vbCr		'면허일 
		
	Response.Write ".txtDocAmt.text			= """ & UNIConvNumDBToCompanyByCurrency(L_E1_m_cc_hdr(E1_doc_amt),lgCurrency,ggAmtOfMoneyNo,"X","X") & """" & vbCr '통관금액 
	Response.Write ".txtBeneficiary.value	= """ & ConvSPChars(L_E1_m_cc_hdr(E1_beneficiary))			& """" & vbCr		'수출자 
	Response.Write ".txtBeneficiaryNm.value = """ & ConvSPChars(L_E1_m_cc_hdr(E1_beneficiary_nm))			& """" & vbCr		'수출자명 
	Response.Write ".txtPurGrp.value		= """ & ConvSPChars(L_E1_m_cc_hdr(E1_pur_grp))		& """" & vbCr		'수입담당 

	Response.Write "If """ & L_E1_m_cc_hdr(E1_xch_rate)							& """ = 0 Then		 " & vbCr		'환율 
	Response.Write ".txtXchRate.value = """ & L_E1_m_cc_hdr(E1_xch_rate)	& """	"	& vbCr
	Response.Write "Else" & vbCr
	Response.Write ".txtXchRate.value = """ & UNINumClientFormat(L_E1_m_cc_hdr(E1_xch_rate), ggExchRate.DecPoint, 0) & """"	& vbCr
	Response.Write "End If" & vbCr
            
    '-----------------------
	' Rounding Column Set               '☜:화폐별 라운딩 스프레드 포매팅 
	'----------------------- 		
	Response.Write "If parent.lgIntFlgMode =  parent.parent.OPMD_CMODE Then parent.CurFormatNumSprSheet" & vbCr             
	Response.Write "parent.lgIntFlgMode = """ & OPMD_UMODE & """ " & vbCr
	Response.Write "End With" & vbCr
		
	Response.Write "</Script>" & vbCr	
End Sub

'============================================================================================================
' Name : SubBizQueryMulti_CC_DTL
' Desc : Query Data from Db - Dev.Defined
'============================================================================================================
Sub SubBizQueryMulti_CC_DTL()
	Dim istrData
	Dim iStrNextKey		' 다음 값 
	Dim lgStrPrevKey	' 이전 값 
	Dim iStrPrevKey
	Dim iLngMaxRow		' 현재 그리드의 최대Row
	Dim iLngRow
	Dim PvArr
	Dim lGrpCnt
	Dim iTotstrData
		
	Const C_SHEETMAXROWS_D  = 100
		
	Dim OBJ_PM6G128
		
	Dim imp_next_cc_seq				'Next Key
	Dim Eg1_export_group
	Dim E1_m_cc_dtl_cc_seq
	Dim E2_m_cc_dtl_max_seq
	Dim E3_m_cc_dtl_total_amt
		
	Const EG1_E1_cc_seq = 0    
	Const EG1_E1_lc_no = 1
	Const EG1_E1_lc_seq = 2
	Const EG1_E1_po_no = 3
	Const EG1_E1_po_seq = 4
	Const EG1_E1_hs_cd = 5
	Const EG1_E1_net_weight = 6
	Const EG1_E1_weight_unit = 7
	Const EG1_E1_qty = 8
	Const EG1_E1_price = 9
	Const EG1_E1_unit = 10
	Const EG1_E1_doc_amt = 11
	Const EG1_E1_cif_doc_amt = 12
	Const EG1_E1_loc_amt = 13
	Const EG1_E1_cif_loc_amt = 14
	Const EG1_E1_lan_no = 15
	Const EG1_E1_receipt_flg = 16
	Const EG1_E1_insrt_user_id = 17
	Const EG1_E1_insrt_dt = 18
	Const EG1_E1_updt_user_id = 19
	Const EG1_E1_updt_dt = 20
	Const EG1_E1_receipt_qty = 21
	Const EG1_E1_ext1_amt = 22
	Const EG1_E1_ext1_cd = 23
	Const EG1_E1_biz_area = 24
	Const EG1_E1_ext1_qty = 25
	Const EG1_E1_ext2_qty = 26
	Const EG1_E1_ext3_qty = 27
	Const EG1_E1_ext2_amt = 28
	Const EG1_E1_ext3_amt = 29
	Const EG1_E1_ext2_cd = 30
	Const EG1_E1_ext3_cd = 31
	Const EG1_E1_ext1_rt = 32
	Const EG1_E1_ext2_rt = 33
	Const EG1_E1_ext3_rt = 34
	Const EG1_E1_ext1_dt = 35
	Const EG1_E1_ext2_dt = 36
	Const EG1_E1_ext3_dt = 37
	Const EG1_E1_tracking_no = 38
	Const EG1_E2_item_cd = 39    
	Const EG1_E2_item_nm = 40
	Const EG1_E2_spec = 41
	Const EG1_E2_item_acct = 42
	Const EG1_E3_plant_cd = 43   
	Const EG1_E3_plant_nm = 44
	Const EG1_E4_hs_nm = 45    
	Const EG1_E5_bl_no = 46    
	Const EG1_E5_bl_doc_no = 47
	Const EG1_E6_bl_seq = 48   
	Const EG1_E6_qty = 49
	Const EG1_E6_cc_qty = 50
	Const EG1_E6_doc_amt = 51
	Const EG1_E7_lc_doc_no = 52 
	Const EG1_E7_lc_amend_seq = 53
    
    Dim iStrCcNo
	On Error Resume Next
	Err.Clear																'☜: Protect system from crashing
	
	Set OBJ_PM6G128 = Server.CreateObject("PM6G128.cMListImportCcDtlS")
		
	If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
	End If
	
	iStrCcNo = 	UCase(Trim(Request("txtCCNo")))
	imp_next_cc_seq = UCase(Trim(Request("lgStrPrevKey")))
	Call OBJ_PM6G128.M_LIST_IMPORT_CC_DTL_SVR(gStrglobalcollection,CLng(C_SHEETMAXROWS_D),imp_next_cc_seq,iStrCcNo, _
											      Eg1_export_group,E1_m_cc_dtl_cc_seq,E2_m_cc_dtl_max_seq,E3_m_cc_dtl_total_amt)
		
	if Request("txtQuerytype") <> "Query" And Err.Description = "B_MESSAGE174400" then
		Response.Write "<Script Language=vbscript>" & vbCr
		Response.write "parent.frm1.vspdData.MaxRows = 0" & vbCr
		Response.write "	parent.lgIntFlgMode = """ & OPMD_UMODE & """ " & vbCr
		Response.Write "parent.dbQueryOk" & vbCr
		Response.Write "</Script>"
		IF UBound(EG1_exp_group,1) <= 0 Then
			Set OBJ_PM6G128 = Nothing
			Exit Sub												'☜: ComProxy Unload	
		End If
		
	Else 
		If CheckSYSTEMError2(Err,True,"","","","","") = true then 		
			Set OBJ_PM6G128 = Nothing												'☜: ComProxy Unload
			'Detail항목이 없을 경우 Header정보만 보여줌 
			Response.Write "<Script Language=vbscript>" & vbCr
			Response.write "parent.frm1.vspdData.MaxRows = 0" & vbCr
			Response.write "	parent.lgIntFlgMode = """ & OPMD_UMODE & """ " & vbCr
			Response.Write "parent.dbQueryOk" & vbCr
			Response.Write "</Script>"
			Exit Sub															'☜: 비지니스 로직 처리를 종료함 
		End If
	End if

			
	Set OBJ_PM6G128 = Nothing														'☜: Unload Comproxy
		
	iLngMaxRow = CInt(Request("txtMaxRows"))											'Save previous Maxrow   
	
	ReDim PvArr(UBound(Eg1_export_group,1))	
	'-----------------------
	'Result data display area
	'-----------------------
	lGrpCnt = 0
	
	For iLngRow = 0 To UBound(Eg1_export_group,1)
		If iLngRow>=C_SHEETMAXROWS_D Then Exit For
		istrData = istrData & Chr(11) & ConvSPChars(Eg1_export_group(iLngRow,EG1_E2_item_cd)) _
		                    & Chr(11) & ConvSPChars(Eg1_export_group(iLngRow,EG1_E2_item_nm)) _
		                    & Chr(11) & ConvSPChars(Eg1_export_group(iLngRow,EG1_E2_spec)) _
		                    & Chr(11) & ConvSPChars(Eg1_export_group(iLngRow,EG1_E1_tracking_no)) _
		                    & Chr(11) & ConvSPChars(Eg1_export_group(iLngRow,EG1_E1_unit)) _
		                    & Chr(11) & UNINumClientFormat(Eg1_export_group(iLngRow,EG1_E1_qty), ggQty.DecPoint, 0) _
		                    & Chr(11) & UNIConvNumDBToCompanyByCurrency(Eg1_export_group(iLngRow,EG1_E1_price), lgCurrency, ggUnitCostNo,"X","X") _
		                    & Chr(11) & UNIConvNumDBToCompanyByCurrency(Eg1_export_group(iLngRow,EG1_E1_doc_amt),lgCurrency,ggAmtOfMoneyNo,"X","X") _
		                    & Chr(11) & UNINumClientFormat(Eg1_export_group(iLngRow,EG1_E1_net_weight), ggQty.DecPoint, 0)
		'& Chr(11) & UNINumClientFormat(Eg1_export_group(iLngRow,EG1_E1_cif_loc_amt), ggAmtOfMoney.DecPoint, 0) _
		istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(Eg1_export_group(iLngRow,EG1_E1_cif_doc_amt),lgCurrency,ggAmtOfMoneyNo,"X","X") _
		                    & Chr(11) & UNIConvNumDBToCompanyByCurrency(Eg1_export_group(iLngRow,EG1_E1_cif_loc_amt),gCurrency,ggAmtOfMoneyNo,"X","X") _
		                    & Chr(11) & ConvSPChars(Eg1_export_group(iLngRow,EG1_E1_hs_cd)) _
		                    & Chr(11) & ConvSPChars(Eg1_export_group(iLngRow,EG1_E4_hs_nm)) _
		                    & Chr(11) & UNINumClientFormat(Eg1_export_group(iLngRow,EG1_E6_qty), ggQty.DecPoint, 0) _
		                    & Chr(11) & UNINumClientFormat(Eg1_export_group(iLngRow,EG1_E1_receipt_qty), ggQty.DecPoint, 0) _
		                    & Chr(11) & ConvSPChars(Eg1_export_group(iLngRow,EG1_E1_cc_seq)) _
		                    & Chr(11) & ConvSPChars(Eg1_export_group(iLngRow,EG1_E5_bl_no)) _
		                    & Chr(11) & ConvSPChars(Eg1_export_group(iLngRow,EG1_E6_bl_seq)) _
		                    & Chr(11) & ConvSPChars(Eg1_export_group(iLngRow,EG1_E5_bl_doc_no)) _
		                    & Chr(11) & ConvSPChars(Eg1_export_group(iLngRow,EG1_E1_po_no)) _
		                    & Chr(11) & ConvSPChars(Eg1_export_group(iLngRow,EG1_E1_po_seq)) _
		                    & Chr(11) & ConvSPChars(Eg1_export_group(iLngRow,EG1_E1_lc_no))
		istrData = istrData & Chr(11) & ConvSPChars(Eg1_export_group(iLngRow,EG1_E1_lc_seq)) _
		                    & Chr(11) & UNIConvNumDBToCompanyByCurrency (Eg1_export_group(iLngRow,EG1_E6_doc_amt),lgCurrency,ggAmtOfMoneyNo,"X","X") _
		                    & Chr(11) & UNINumClientFormat(Eg1_export_group(iLngRow,EG1_E6_cc_qty), ggQty.DecPoint, 0) _
		                    & Chr(11) & UNIConvNumDBToCompanyByCurrency(Eg1_export_group(iLngRow,EG1_E1_doc_amt),lgCurrency,ggAmtOfMoneyNo,"X","X") _
		                    & Chr(11) & UNIConvNumDBToCompanyByCurrency(Eg1_export_group(iLngRow,EG1_E1_doc_amt),lgCurrency,ggAmtOfMoneyNo,"X","X") _
		                    & Chr(11) & iLngMaxRow + iLngRow _			
		                    & Chr(11) & Chr(12)
		
		PvArr(lGrpCnt) = istrData
		lGrpCnt = lGrpCnt + 1
		istrData = ""
	Next
	
	iTotstrData = Join(PvArr, "")
		
	If  iStrPrevKey = E1_m_cc_dtl_cc_seq  Then
		iStrPrevKey = ""
	Else
		iStrNextKey = E1_m_cc_dtl_cc_seq
	End If
		
	Response.Write "<Script Language=VBScript>"					& vbCr
	Response.Write "With parent"							& vbCr
	Response.Write "	.ggoSpread.Source = .frm1.vspdData "				& vbCr
	'Response.Write  "    .frm1.vspdData.Redraw = False   "                  & vbCr      
    Response.Write "	.ggoSpread.SSShowData        """ & iTotstrData	    & """" & vbCr	
		
	'Response.Write  "    Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData," & iLngMaxRow + 1 & "," & iLngMaxRow + iLngRow & "	,lgCurrency	,.C_Price		,""C"" ,""I"",""X"",""X"")" & vbCr
	'Response.Write  "    Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData," & iLngMaxRow + 1 & "," & iLngMaxRow + iLngRow & "	,lgCurrency	,.C_DocAmt		,""A"" ,""I"",""X"",""X"")" & vbCr
	'Response.Write  "    Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData," & iLngMaxRow + 1 & "," & iLngMaxRow + iLngRow & "	,.parent.gCurrency	,.C_CIFLocAmt		,""A"" ,""I"",""X"",""X"")" & vbCr
	'Response.Write  "    Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData," & iLngMaxRow + 1 & "," & iLngMaxRow + iLngRow & "	,lgCurrency	,.C_BlAmt		,""A"" ,""I"",""X"",""X"")" & vbCr
    'Response.Write  "    Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData," & iLngMaxRow + 1 & "," & iLngMaxRow + iLngRow & "	,lgCurrency	,.C_OrgDocAmt		,""A"" ,""I"",""X"",""X"")" & vbCr
    'Response.Write  "    Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData," & iLngMaxRow + 1 & "," & iLngMaxRow + iLngRow & "	,lgCurrency	,.C_OrgDocAmt1		,""A"" ,""I"",""X"",""X"")" & vbCr
    
	Response.Write "		.lgStrPrevKey =	""" & iStrNextKey						    & """ " & vbCr
	'Response.Write "	If .frm1.vspdData.MaxRows < .C_SHEETMAXROWS And .lgStrPrevKey <> """" Then	  " & vbCr
	'Response.Write "		.DbQuery "		    			& vbCr 
	'Response.Write "	Else															"		& vbCr
	Response.Write "		.frm1.txtHCCNo.value = """ & ConvSPChars(Request("txtCCNo"))		& """ "	& vbCr
	Response.Write "		.DbQueryOk "		    		& vbCr 
	'Response.Write "	End If							"	& vbCr
	'Response.Write  "    .frm1.vspdData.Redraw = True   "   & vbCr
	Response.Write "End With"								& vbCr
	Response.Write "</Script>"								& vbCr

End Sub

%>
