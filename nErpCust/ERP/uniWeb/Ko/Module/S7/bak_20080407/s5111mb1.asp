<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 매출채권관리 
'*  3. Program ID           : S5111MB1
'*  4. Program Name         : 매출채권등록 
'*  5. Program Desc         :
'*  6. Comproxy List        : PS7G111.cSBillHdrSvr,PS3G102.cLookupSoHdrSvr,PB5CS41.cLookupBizPartnerSvr
'*							  PS4G119.cSLkLcHdrSvr,PB5CS41.cLookupBizPartnerSvr	
'*							  PS7G115.cSPostOpenArSvr
'*  7. Modified date(First) : 2000/04/19
'*  8. Modified date(Last)  : 2002/11/15
'*  9. Modifier (First)     : Cho song hyon
'* 10. Modifier (Last)      : Ahn Tae Hee
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/04/19 : 3rd 화면 Layout & ASP Coding
'*                            -2000/08/11 : 4th 화면 Layout
'*                            -2001/12/18 : Date 표준적용 
'*                            -2002/11/15 : UI 표준적용 
'**********************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%Call LoadBasisGlobalInf()
Call loadInfTB19029B("I", "*","NOCOOKIE","MB") 
Call LoadBNumericFormatB("I","*","NOCOOKIE","MB") %>

<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<%		
Dim strMode
On Error Resume Next                                                             '☜: Protect system from crashing
Err.Clear                                                                        '☜: Clear Error status

'for 구주Tax
'@@@@@@@@@@@
Dim pvCB 
'@@@@@@@@@@@
Call HideStatusWnd                                                               '☜: Hide Processing message
'---------------------------------------Common-----------------------------------------------------------

'------ Developer Coding part (Start ) ------------------------------------------------------------------

'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

strMode      = Request("txtMode")                                               '☜: Read Operation Mode (CRUD)
Select Case strMode
    Case CStr(UID_M0001)                                                        '☜: Query
         Call SubBizQuery()
    Case CStr(UID_M0002)
         Call SubBizSave()
    Case CStr(UID_M0003)                                                        '☜: Delete
         Call SubBizDelete()
    Case CStr("SoNoHdr")	     												'☜: 현재 수주헤더관련조회를 요청받음 
         Call SubBizSoNoHdr
    Case CStr("LCNoHdr")														'☜: 현재 LC헤더관련조회를 요청받음 
         Call SubBizLCNoHdr
    Case CStr("PostFlag")		     											'☜: 확정 요청 
	     Call SubBizPostFlag
End Select
'============================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================
Sub SubBizQuery()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear 
    Call HideStatusWnd
    Dim lgCurrency   
    Err.Clear                                                               '☜: Protect system from crashing

    Dim lgStrGlFlag
    Dim lgArrGlFlag
	
	        
    Call SubOpenDB(lgObjConn)
    call SubMakeSQLStatements
 
	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                 'If data not exists
	    lgObjRs.Close
	    lgObjConn.Close
	    Set lgObjRs = Nothing
	    Set lgObjConn = Nothing
	    If Request("txtPrevNext") = "Q" Then
			'매출채권정보가 없습니다.
		    Call DisplayMsgBox("205100", vbInformation, "", "", I_MKSCRIPT)             '☜: No data is found. 
		Elseif Request("txtPrevNext") = "P" Then
			'이전 자료가 없습니다.
		    Call DisplayMsgBox("200002", vbInformation, "", "", I_MKSCRIPT)             '☜: No data is found. 
		Else
			'이후 자료가 없습니다.
		    Call DisplayMsgBox("200003", vbInformation, "", "", I_MKSCRIPT)             '☜: No data is found. 
		End If
		
	    Exit Sub
	End If

	lgCurrency = ConvSPChars(lgObjRs("Cur")) 
	'-----------------------
	'Result data display area
	'----------------------- 
    Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent.frm1"           & vbCr
	
		'##### Rounding Logic #####
		'항상 거래화폐가 우선 
		
	Response.Write ".txtDocCur1.value		= """ & UCase(lgCurrency)		     					& """" & vbCr
	Response.Write ".txtDocCur2.value	    = """ & UCase(lgCurrency)		      				    & """" & vbCr
	   
	Response.Write " parent.CurFormatNumericOCX " & vbCr 		
    
		'##########################
	Response.Write ".txtBillDt.Text		    = """ & UNIDateClientFormat(lgObjRs("bill_dt"))          & """" & vbCr
	Response.Write ".txtPlanIncomeDt.Text   = """ & UNIDateClientFormat(lgObjRs("income_plan_dt"))   & """" & vbCr	
	Response.Write ".txtConBillNo.value     = """ & ConvSPChars(lgObjRs("bill_no"))                  & """" & vbCr		
	Response.Write ".txtBillNo.value        = """ & ConvSPChars(lgObjRs("bill_no"))                  & """" & vbCr
	Response.Write ".txtHBillNo.value       = """ & ConvSPChars(lgObjRs("bill_no"))                  & """" & vbCr
	Response.Write ".txtTaxBillNo.value     = """ & ConvSPChars(lgObjRs("tax_bill_no"))              & """" & vbCr
	Response.Write ".txtPayerCd.value       = """ & ConvSPChars(lgObjRs("payer"))                    & """" & vbCr
	Response.Write ".txtPayerNm.value       = """ & ConvSPChars(lgObjRs("payer_nm"))                 & """" & vbCr			
				
		'수금영업그룹 
	Response.Write ".txtToBizAreaCd.value   = """ & ConvSPChars(lgObjRs("to_sales_grp"))             & """" & vbCr
	Response.Write ".txtToBizAreaNm.value   = """ & ConvSPChars(lgObjRs("to_sales_grp_nm"))          & """" & vbCr
	Response.Write ".txtTaxBizAreaCd.value  = """ & ConvSPChars(lgObjRs("tax_biz_area"))             & """" & vbCr
	Response.Write ".txtTaxBizAreaNm.value  = """ & ConvSPChars(lgObjRs("tax_biz_area_nm"))          & """" & vbCr
	Response.Write ".txtPayTypeCd.value     = """ & ConvSPChars(lgObjRs("pay_type"))                 & """" & vbCr
	Response.Write ".txtPayTypeNm.value     = """ & ConvSPChars(lgObjRs("pay_type_nm"))              & """" & vbCr
	Response.Write ".txtPayTermsCd.value    = """ & ConvSPChars(lgObjRs("pay_meth"))                 & """" & vbCr
	Response.Write ".txtPayTermsNm.value    = """ & ConvSPChars(lgObjRs("pay_meth_nm"))              & """" & vbCr
	
	If Trim(lgObjRs("pay_dur")) = "0" Then
	
	Response.Write ".txtPayDur.Text         = """"" & vbCr
			
	Else
	
	Response.Write ".txtPayDur.Text         = """ & lgObjRs("pay_dur")                               & """" & vbCr

	End If

	Response.Write ".txtXchgRate.Text      = """ & UNINumClientFormat(lgObjRs("Xchg_Rate"), ggExchRate.DecPoint, 0)      & """" & vbCr


		'매출채권금액 
    Response.Write ".txtBillAmt.Text       = """ & UNIConvNumDBToCompanyByCurrency(lgObjRs("bill_amt"), lgCurrency, ggAmtOfMoneyNo, "X" , "X")                       & """" & vbCr
    Response.Write ".txtBillAmtLoc.Text    = """ & UNIConvNumDBToCompanyByCurrency(lgObjRs("bill_amt_loc"), lgCurrency, ggAmtOfMoneyNo, gLocRndPolicyNo , "X")       & """" & vbCr
		
		'VAT금액 
    Response.Write ".txtVatAmt.Text        = """ & UNIConvNumDBToCompanyByCurrency(lgObjRs("vat_amt"), lgCurrency, ggAmtOfMoneyNo, gLocRndPolicyNo , "X")                & """" & vbCr
    Response.Write ".txtVatLocAmt.Text     = """ & UNIConvNumDBToCompanyByCurrency(lgObjRs("vat_amt_loc"), lgCurrency, ggAmtOfMoneyNo, gLocRndPolicyNo , "X")         & """" & vbCr
		
		'적립금액 
    Response.Write ".txtDepositAmt.Text    = """ & UNIConvNumDBToCompanyByCurrency(lgObjRs("deposit_amt"), lgCurrency, ggAmtOfMoneyNo, "X" , "X")                    & """" & vbCr
    Response.Write ".txtDepositAmtLoc.Text = """ & UNIConvNumDBToCompanyByCurrency(lgObjRs("deposit_amt_loc"), gCurrency, ggAmtOfMoneyNo, gLocRndPolicyNo , "X")  & """" & vbCr
	
		'총매출채권금액 
    Response.Write ".txtTotBillAmt.Text    = """ & UNIConvNumDBToCompanyByCurrency(CDbl(lgObjRs("bill_amt")) + CDbl(lgObjRs("vat_amt")) + CDbl(lgObjRs("deposit_amt")), lgCurrency, ggAmtOfMoneyNo, "X" , "X")  & """" & vbCr
    Response.Write ".txtTotBillAmtLoc.Text = """ & UNIConvNumDBToCompanyByCurrency(CDbl(lgObjRs("bill_amt_loc")) + CDbl(lgObjRs("vat_amt_loc")) + CDbl(lgObjRs("deposit_amt_loc")), gCurrency, ggAmtOfMoneyNo, gLocRndPolicyNo, "X")  & """" & vbCr

			
		'총수금액 
    Response.Write ".txtIncomeAmt.Text     = """ & UNIConvNumDBToCompanyByCurrency(lgObjRs("collect_amt"), lgCurrency, ggAmtOfMoneyNo, "X" , "X")  & """" & vbCr
    Response.Write ".txtIncomeLocAmt.Text  = """ & UNIConvNumDBToCompanyByCurrency(lgObjRs("collect_amt_loc"), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")  & """" & vbCr

		'통관FOB 금액 
    Response.Write ".txtAcceptFobAmt.Text  = """ & UNIConvNumDBToCompanyByCurrency(lgObjRs("accept_fob_amt"), lgCurrency, ggAmtOfMoneyNo, "X" , "X")  & """" & vbCr

	Response.Write ".txtVatType.value    = """ & ConvSPChars(lgObjRs("vat_type"))              & """" & vbCr
	Response.Write ".txtVatTypeNm.value    = """ & ConvSPChars(lgObjRs("vat_type_nm"))              & """" & vbCr
	Response.Write ".txtVATRate.Text    = """ & UNINumClientFormat(lgObjRs("vat_rate"), ggExchRate.DecPoint, 0)              & """" & vbCr
	Response.Write ".txtVatCalcType.value    = """ & ConvSPChars(lgObjRs("vat_calc_type"))              & """" & vbCr

	Response.Write ".txtBillTypeCd.value    = """ & ConvSPChars(lgObjRs("bill_type"))              & """" & vbCr
	Response.Write ".txtBillTypeNm.value    = """ & ConvSPChars(lgObjRs("bill_type_nm"))              & """" & vbCr

	Response.Write ".txtBillToPartyCd.value    = """ & ConvSPChars(lgObjRs("bill_to_party"))              & """" & vbCr
	Response.Write ".txtBillToPartyNm.value    = """ & ConvSPChars(lgObjRs("bill_to_party_nm"))              & """" & vbCr
	Response.Write ".txtSoldtoPartyCd.value    = """ & ConvSPChars(lgObjRs("sold_to_party"))              & """" & vbCr
	Response.Write ".txtSoldtoPartyNm.value    = """ & ConvSPChars(lgObjRs("sold_to_party_nm"))              & """" & vbCr

	Response.Write ".txtSoNo.value    = """ & ConvSPChars(lgObjRs("so_no"))              & """" & vbCr
		
	If Trim(ConvSPChars(lgObjRs("so_no"))) <> "" Then

	Response.Write ".chkSoNo.checked = True             " & vbCr	
	End If

	Response.Write ".txtSalesGrpCd.value      = """ & ConvSPChars(lgObjRs("sales_grp"))          & """" & vbCr
	Response.Write ".txtSalesGrpNm.value      = """ & ConvSPChars(lgObjRs("sales_grp_nm"))       & """" & vbCr

	Response.Write ".txtPaytermsTxt.value     = """ & ConvSPChars(Trim(lgObjRs("pay_terms_txt")))       & """" & vbCr
	Response.Write ".txtRemark.value          = """ & ConvSPChars(lgObjRs("remark"))             & """" & vbCr
	Response.Write ".txtBeneficiaryCd.value   = """ & ConvSPChars(lgObjRs("beneficiary"))        & """" & vbCr
	Response.Write ".txtBeneficiaryNm.value   = """ & ConvSPChars(lgObjRs("beneficiary_nm"))     & """" & vbCr
	Response.Write ".txtApplicantCd.value     = """ & ConvSPChars(lgObjRs("applicant"))          & """" & vbCr
	Response.Write ".txtApplicantNm.value     = """ & ConvSPChars(lgObjRs("applicant_nm"))       & """" & vbCr
	
	Response.Write ".txtLCNo.value            = """ & ConvSPChars(lgObjRs("lc_no"))              & """" & vbCr
	Response.Write ".txtLCAmendSeq.value      = """ & ConvSPChars(lgObjRs("lc_amend_seq"))       & """" & vbCr
	Response.Write ".txtLCDocNo.value         = """ & ConvSPChars(lgObjRs("lc_doc_no"))          & """" & vbCr
	Response.Write ".txtLocCur.value          = """ & UCase(gCurrency)                           & """" & vbCr
				
  
		'VAT적용기준 
	If Trim(ConvSPChars(lgObjRs("vat_calc_type"))) = "1" Then
	    Response.Write ".rdoVATCalcType1.checked = True         "    & vbCr
	Else
	    Response.Write ".rdoVATCalcType2.checked = True         "    & vbCr
	End If

		'VAT포함여부 
	If TRIM (ConvSPChars(lgObjRs("vat_inc_flag"))) = "1" Then
	    Response.Write ".rdoVATIncFlag1.checked = True           "   & vbCr
	Else
	    Response.Write ".rdoVATIncFlag2.checked = True           "   & vbCr
	End If

		'매출진행상태 
	Response.Write ".txtSts.value                = """ & lgObjRs("sts")                 & """" & vbCr

		'수주, L/C참조 여부 
	Response.Write ".txtRefFlag.value            = """ & lgObjRs("ref_flag")            & """" & vbCr				
		
		'반품 여부 
	Response.Write ".txtRetItemFlag.value        = """ & lgObjRs("reverse_flag")        & """" & vbCr
				
		'VATAutoCheck 
	Response.Write ".txtchkTaxNo.value           = """ & lgObjRs("vat_auto_flag")       & """" & vbCr
		
		'약정회전일 
	Response.Write ".txtCreditRotDay.value       = """ & lgObjRs("credit_rot_day")      & """" & vbCr

	If lgObjRs("post_flag") = "Y" AND Len(Trim(lgObjRs("gl_no"))) Then
		lgArrGlFlag    =  Split(lgObjRs("gl_no"), Chr(11))      
		lgStrGlFlag = lgArrGlFlag(0) 
		
		Response.Write ".txtAcctNo.value         = """ & lgArrGlFlag(1)            & """" & vbCr		
		
		If lgArrGlFlag(0) = "G" Then	
			'회계전표번호 
			Response.Write ".txtGLNo.value       = """ & lgArrGlFlag(1)            & """" & vbCr			 
		ElseIf lgArrGlFlag(0) = "T" Then
			'결의전표번호 
			Response.Write ".txtTempGLNo.value   = """ & lgArrGlFlag(1)            & """" & vbCr		
		Else
			'Batch번호 
			Response.Write ".txtBatchNo.value    = """ & lgArrGlFlag(1)            & """" & vbCr	
		End If
	Else
		Response.Write ".txtAcctNo.value     = """""     & vbCr	
		Response.Write ".txtGLNo.value       = """""     & vbCr	
		Response.Write ".txtTempGLNo.value   = """""     & vbCr	
		Response.Write ".txtBatchNo.value    = """""     & vbCr	
				
	End If
		
		If lgObjRs("post_flag") = "Y" Then
			Response.Write ".txtRadioFlag.value = ""Y"""                & vbCr
			Response.Write ".btnPostFlag.value = ""확정취소"""      & vbCr

			if lgStrGlFlag = "G" Or lgStrGlFlag = "T" Then
			    Response.Write ".btnGLView.disabled = False"            & vbCr

			Else
			    Response.Write ".btnGLView.disabled = True"             & vbCr
			End if
		Else

			Response.Write ".txtRadioFlag.value = ""N"""                & vbCr
			Response.Write ".btnPostFlag.value = ""확정"""          & vbCr
			Response.Write ".btnGLView.disabled = True"                 & vbCr
		End If

		'선수금 현황 버튼 Enable 
		IF lgObjRs("PreRcpt_flag") = "Y" Then
			Response.Write ".btnPreRcptView.disabled = False "        & vbCr
		Else
			Response.Write ".btnPreRcptView.disabled = True "         & vbCr
		End If

	Response.Write "parent.DbQueryOk	"                             & vbCr					'☜: 조회가 성공 
    Response.Write "End With"          & vbCr
    Response.Write "</Script>"         & vbCr

	lgObjRs.Close
	lgObjConn.Close
	Set lgObjRs = Nothing
	Set lgObjConn = Nothing
										
End Sub	

'============================================
' Name : SubBizQuery
' Desc : Date data 
'============================================
Sub SubBizSave()

	Dim iS51111
	Dim lgIntFlgMode
    Dim pvCommandSent
    Dim I1_s_bill_hdr
    Dim I2_s_bill_type_config 
    Dim I3_b_biz_partner 
    Dim I4_b_sales_grp 
    Dim I5_b_sales_grp 
    Dim I6_s_bill_hdr 
    Dim I7_b_biz_partner 
    Dim I8_b_biz_partner 
    Dim I9_b_sales_org 
    Dim E2_s_bill_hdr
    
    Const S500_I6_bill_no = 0    'View Name : imp s_bill_hdr
    Const S500_I6_bill_dt = 1
    Const S500_I6_cur = 2
    Const S500_I6_xchg_rate = 3
    Const S500_I6_vat_type = 4
    Const S500_I6_vat_rate = 5
    Const S500_I6_pay_meth = 6
    Const S500_I6_pay_dur = 7
    Const S500_I6_tax_bill_no = 8
    Const S500_I6_beneficiary = 9
    Const S500_I6_applicant = 10
    Const S500_I6_post_flag = 11
    Const S500_I6_remark = 12
    Const S500_I6_vat_calc_type = 13
    Const S500_I6_tax_biz_area = 14
    Const S500_I6_pay_type = 15
    Const S500_I6_pay_terms_txt = 16
    Const S500_I6_collect_amt = 17
    Const S500_I6_collect_amt_loc = 18
    Const S500_I6_income_plan_dt = 19
    Const S500_I6_so_no = 20
    Const S500_I6_lc_no = 21
    Const S500_I6_lc_doc_no = 22
    Const S500_I6_lc_amend_seq = 23
    Const S500_I6_bl_flag = 24
    Const S500_I6_ref_flag = 25
    Const S500_I6_except_flag = 26
    Const S500_I6_reverse_flag = 27
    Const S500_I6_ext1_qty = 28
    Const S500_I6_ext2_qty = 29
    Const S500_I6_ext3_qty = 30
    Const S500_I6_ext1_amt = 31
    Const S500_I6_ext2_amt = 32
    Const S500_I6_ext3_amt = 33
    Const S500_I6_ext1_cd = 34
    Const S500_I6_ext2_cd = 35
    Const S500_I6_ext3_cd = 36
    Const S500_I6_vat_auto_flag = 37
    Const S500_I6_vat_inc_flag = 38
    Const S500_I6_deposit_amt = 39

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear     
    I1_s_bill_hdr  = ""
    I9_b_sales_org = ""
    Redim I6_s_bill_hdr(S500_I6_deposit_amt)

    '-----------------------
    'Data manipulate area
    '-----------------------

	'=============== TAB 1 ==============

	'<--매출채권번호 -->
    I6_s_bill_hdr(S500_I6_bill_no)        = UCase(Trim(Request("txtBillNo")))
	'<--매출채권일 -->
    I6_s_bill_hdr(S500_I6_bill_dt)        = UNIConvDate(Request("txtBillDt"))
	'<--수금예정일 -->
	I6_s_bill_hdr(S500_I6_income_plan_dt) = UNIConvDate(Request("txtPlanIncomeDt"))
	'<--수금처 -->
    I8_b_biz_partner = UCase(Trim(Request("txtPayerCd")))
	'<--주문처 -->
    I3_b_biz_partner = UCase(Trim(Request("txtSoldtoPartyCd")))
    '<--수금영업그룹 -->
	I4_b_sales_grp = UCase(Trim(Request("txtToBizAreaCd")))
	'<--입금유형 -->
    I6_s_bill_hdr(S500_I6_pay_type) = UCase(Trim(Request("txtPayTypeCd")))
	'<--결제방법 -->
    I6_s_bill_hdr(S500_I6_pay_meth) = UCase(Trim(Request("txtPayTermsCd")))
	'<--결제기간 -->
    I6_s_bill_hdr(S500_I6_pay_dur) =  UNIConvNum(Trim(Request("txtPayDur")),0)

	'<--환율 -->
	If Len(Trim(Request("txtXchgRate"))) Then I6_s_bill_hdr(S500_I6_xchg_rate) = UNIConvNum(Trim(Request("txtXchgRate")),0)

	'<--부가세유형 -->
	I6_s_bill_hdr(S500_I6_vat_type) = UCase(Trim(Request("txtVatType")))
	'<--부가세율 -->
    If Len(Trim(Request("txtVatRate"))) Then I6_s_bill_hdr(S500_I6_vat_rate) = UNIConvNum(Trim(Request("txtVatRate")),0)
	'<--매출채권타입 -->
	I2_s_bill_type_config = UCase(Trim(Request("txtBillTypeCd")))
	'<--발행처 -->
    I7_b_biz_partner =UCase(Trim(Request("txtBillToPartyCd")))

	If Request("txtChkSoNo") = "Y" Then
		'<--수주번호 -->
	    I6_s_bill_hdr(S500_I6_so_no) = UCase(Trim(Request("txtSoNo")))
	End If

	'<--영업그룹 -->
    I5_b_sales_grp = UCase(Trim(Request("txtSalesGrpCd")))
	'<--화폐 -->
	I6_s_bill_hdr(S500_I6_cur) = UCase(Trim(Request("txtDocCur1")))

	'<--대금결제조건 -->
    I6_s_bill_hdr(S500_I6_pay_terms_txt) = UCase(Trim(Request("txtPaytermsTxt")))
	'<--비고 -->
	I6_s_bill_hdr(S500_I6_remark) = UCase(Trim(Request("txtRemark")))

	'<--세금계산서 자동발행 -->
	I6_s_bill_hdr(S500_I6_vat_auto_flag)  = Trim(Request("txtchkTaxNo"))
	'<--부가세적용기준(개별/통합)-->
	I6_s_bill_hdr(S500_I6_vat_calc_type) = Trim(Request("txtVatCalcType"))	
	'<--부가세포함구분(별도/포함)-->
	I6_s_bill_hdr(S500_I6_vat_inc_flag) = Trim(Request("txtVatIncFlag"))	
	'<--세금신고사업장 -->
	I6_s_bill_hdr(S500_I6_tax_biz_area) = UCase(Trim(Request("txtTaxBizAreaCd")))
	'<--세금계산서번호 -->
	I6_s_bill_hdr(S500_I6_tax_bill_no)  = UCase(Trim(Request("txtTaxBillNo")))
	'=============== TAB 2 ==============

	'<--양도자 -->
	I6_s_bill_hdr(S500_I6_applicant)  = UCase(Trim(Request("txtApplicantCd")))
	'<--양수자 -->
    I6_s_bill_hdr(S500_I6_beneficiary)  = UCase(Trim(Request("txtBeneficiaryCd")))
	'<--L/C순번 -->
	If Len(Trim(Request("txtLCAmendSeq"))) Then I6_s_bill_hdr(S500_I6_lc_amend_seq)  = Trim(Request("txtLCAmendSeq"))
	'<--L/C관리번호 -->
	I6_s_bill_hdr(S500_I6_lc_doc_no) = UCase(Trim(Request("txtLCDocNo")))
	'<--L/C번호 -->
	I6_s_bill_hdr(S500_I6_lc_no) = UCase(Trim(Request("txtLCNo")))


	'=========== Common Value ===========	

	'---> 매출채권인경우 "N", 예외매출채권인경우 "Y"
    I6_s_bill_hdr(S500_I6_except_flag) = "N"
	'--->수주인경우("S"), LC인경우("L")
    I6_s_bill_hdr(S500_I6_ref_flag) = UCase(Trim(Request("txtRefFlag")))
	'---> B/L 여부인 (국내매출인 경우는 항상 N)
    I6_s_bill_hdr(S500_I6_bl_flag) = "N"

	'--->반품여부 
	lgIntFlgMode = CInt(Request("txtFlgMode"))										'☜: 저장시 Create/Update 판별  
    I6_s_bill_hdr(S500_I6_reverse_flag) = UCase(Trim(Request("txtRetItemFlag")))
    If lgIntFlgMode = OPMD_CMODE Then
		pvCommandSent = "CREATE"
    ElseIf lgIntFlgMode = OPMD_UMODE Then
		pvCommandSent = "UPDATE"
    End If
    
    Set iS51111 = Server.CreateObject("PS7G111.cSBillHdrSvr")
   
	If CheckSYSTEMError(Err,True) = True Then		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  
    
    'for 구주Tax
    '@@@@@@@@@@@
     pvCB = "F"
    '@@@@@@@@@@@
    E2_s_bill_hdr = ""
    E2_s_bill_hdr = iS51111.S_MAINT_BILL_HDR_SVR(pvCB,gStrGlobalCollection, pvCommandSent,I1_s_bill_hdr,I2_s_bill_type_config , _
              I3_b_biz_partner , I4_b_sales_grp ,I5_b_sales_grp , I6_s_bill_hdr , I7_b_biz_partner , _
              I8_b_biz_partner , I9_b_sales_org  , "" )


	If CheckSYSTEMError(Err,True) = True Then
       Call SetErrorStatus                                                           '☆: Mark that error occurs
       Set iS51111 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  
    
    Set iS51111 = Nothing

	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent"           & vbCr

	If ConvSPChars(E2_s_bill_hdr) <>"" Then
	    Response.Write ".frm1.txtConBillNo.value = """ & ConvSPChars(E2_s_bill_hdr) & """" & vbCr
	end if
    Response.Write ".DbSaveOk"                  & vbCr
	Response.Write "End With"                   & vbCr
    Response.Write "</Script>"                  & vbCr

End Sub
'============================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================
Sub SubBizDelete()
	Dim iS51111
    Dim I6_s_bill_hdr 
    Const S500_I6_bill_no = 0    'View Name : imp s_bill_hdr
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear    
	Redim I6_s_bill_hdr(S500_I6_bill_no)
	I6_s_bill_hdr(S500_I6_bill_no) = Trim(Request("txtBillNo"))

    Set iS51111 = Server.CreateObject("PS7G111.cSBillHdrSvr")

	If CheckSYSTEMError(Err,True) = True Then		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  
	
	'for 구주Tax
    '@@@@@@@@@@@
     pvCB = "F"
    '@@@@@@@@@@@
    Call iS51111.S_MAINT_BILL_HDR_SVR(pvCB,gStrGlobalCollection, "DELETE","","" , _
			                               "" , "" ,"" , I6_s_bill_hdr , "" , _
                                           "" , ""  , "" )

	If CheckSYSTEMError(Err,True) = True Then
       Call SetErrorStatus                                                           '☆: Mark that error occurs
       Set iS51111 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  
    
    Set iS51111 = Nothing
	Response.Write "<Script Language=vbscript>" & vbCr
    Response.Write "Call parent.DbDeleteOk() " & vbCr
    Response.Write "</Script>"                  & vbCr
    
End Sub

Sub SubBizSoNoHdr()                                            '☜: 현재 수주헤더관련조회를 요청받음 
                                                   
	Dim iS3G102
	Dim iCommandSent
	Dim I1_s_so_hdr
	Dim E1_s_so_hdr
    
    Dim iB5CS41
	Dim imp_biz_partner_cd
    Dim E1_b_biz_partner
    'iS3G102
    Const S308_E1_so_no = 0
    Const S308_E1_so_dt = 1
    Const S308_E1_req_dlvy_dt = 2
    Const S308_E1_cfm_flag = 3
    Const S308_E1_price_flag = 4
    Const S308_E1_cur = 5
    Const S308_E1_xchg_rate = 6
    Const S308_E1_net_amt = 7
    Const S308_E1_net_amt_loc = 8
    Const S308_E1_cust_po_no = 9
    Const S308_E1_cust_po_dt = 10
    Const S308_E1_sales_cost_center = 11
    Const S308_E1_deal_type = 12
    Const S308_E1_pay_meth = 13
    Const S308_E1_pay_dur = 14
    Const S308_E1_trans_meth = 15
    Const S308_E1_vat_inc_flag = 16
    Const S308_E1_vat_type = 17
    Const S308_E1_vat_rate = 18
    Const S308_E1_vat_amt = 19
    Const S308_E1_vat_amt_loc = 20
    Const S308_E1_origin_cd = 21
    Const S308_E1_valid_dt = 22
    Const S308_E1_contract_dt = 23
    Const S308_E1_ship_dt_txt = 24
    Const S308_E1_pack_cond = 25
    Const S308_E1_inspect_meth = 26
    Const S308_E1_incoterms = 27
    Const S308_E1_dischge_city = 28
    Const S308_E1_dischge_port_cd = 29
    Const S308_E1_loading_port_cd = 30
    Const S308_E1_beneficiary = 31
    Const S308_E1_manufacturer = 32
    Const S308_E1_agent = 33
    Const S308_E1_remark = 34
    Const S308_E1_pre_doc_no = 35
    Const S308_E1_lc_flag = 36
    Const S308_E1_rel_dn_flag = 37
    Const S308_E1_rel_bill_flag = 38
    Const S308_E1_ret_item_flag = 39
    Const S308_E1_sp_stk_flag = 40
    Const S308_E1_ci_flag = 41
    Const S308_E1_export_flag = 42
    Const S308_E1_so_sts = 43
    Const S308_E1_insrt_user_id = 44
    Const S308_E1_insrt_dt = 45
    Const S308_E1_updt_user_id = 46
    Const S308_E1_updt_dt = 47
    Const S308_E1_ext1_qty = 48
    Const S308_E1_ext2_qty = 49
    Const S308_E1_ext3_qty = 50
    Const S308_E1_ext1_amt = 51
    Const S308_E1_ext2_amt = 52
    Const S308_E1_ext3_amt = 53
    Const S308_E1_ext1_cd = 54
    Const S308_E1_maint_no = 55
    Const S308_E1_ext3_cd = 56
    Const S308_E1_pay_type = 57
    Const S308_E1_pay_terms_txt = 58
    Const S308_E1_dn_parcel_flag = 59
    Const S308_E1_to_biz_area = 60
    Const S308_E1_to_biz_grp = 61
    Const S308_E1_biz_area = 62
    Const S308_E1_to_biz_org = 63
    Const S308_E1_to_biz_cost_center = 64
    Const S308_E1_ship_dt = 65
    Const S308_E1_auto_dn_flag = 66
    Const S308_E1_ext2_cd = 67
    Const S308_E1_bank_cd = 68
    Const S308_E1_sales_grp = 69
    Const S308_E1_sales_grp_nm = 70
    Const S308_E1_so_type = 71
    Const S308_E1_so_type_nm = 72
    Const S308_E1_bill_to_party = 73
    Const S308_E1_bill_to_party_type = 74
    Const S308_E1_bill_to_party_nm = 75
    Const S308_E1_ship_to_party = 76
    Const S308_E1_ship_to_party_type = 77
    Const S308_E1_ship_to_party_nm = 78
    Const S308_E1_sold_to_party = 79
    Const S308_E1_sold_to_party_type = 80
    Const S308_E1_sold_to_party_nm = 81
    Const S308_E1_payer = 82
    Const S308_E1_payer_type = 83
    Const S308_E1_payer_nm = 84
    Const S308_E1_sales_org = 85
    Const S308_E1_sales_org_nm = 86
    Const S308_E1_bank_nm = 87
    Const S308_E1_deal_type_nm = 88
    Const S308_E1_vat_type_nm = 89
    Const S308_E1_pay_meth_nm = 90
    Const S308_E1_incoterms_nm = 91
    Const S308_E1_pack_cond_nm = 92
    Const S308_E1_inspect_meth_nm = 93
    Const S308_E1_trans_meth_nm = 94
    Const S308_E1_vat_inc_flag_nm = 95
    Const S308_E1_pay_type_nm = 96
    Const S308_E1_loading_port_nm = 97
    Const S308_E1_dischge_port_nm = 98
    Const S308_E1_origin_nm = 99
    Const S308_E1_manufacturer_nm = 100
    Const S308_E1_agent_nm = 101
    Const S308_E1_beneficiary_nm = 102
    Const S308_E1_currency_desc = 103
    Const S308_E1_biz_area_nm = 104
    Const S308_E1_to_biz_grp_nm = 105
    'iB5CS41
    Const S074_E1_credit_rot_day = 53
	On Error Resume Next
	Err.Clear               
	                   
	iCommandSent = "QUERY"
    I1_s_so_hdr = Trim(Request("txtSoNo"))

    Set iS3G102 = Server.CreateObject("PS3G102.cLookupSoHdrSvr")
    
	If CheckSYSTEMError(Err, True) = True Then
		Exit Sub
	End If
    
	Call iS3G102.S_LOOKUP_SO_HDR_SVR(gStrGlobalCollection, iCommandSent, I1_s_so_hdr, E1_s_so_hdr)
											
	If CheckSYSTEMError(Err, True) = True Then
		Set iS3G102 = Nothing		                                                 '☜: Unload Comproxy DLL
		Exit Sub
	End If
   
	Set iS3G102 = Nothing
    imp_biz_partner_cd = Trim(E1_s_so_hdr(S308_E1_sold_to_party)) 

    iCommandSent = "LOOKUP"

    Set iB5CS41 = Server.CreateObject("PB5CS41.cLookupBizPartnerSvr")    

    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If     

	Call iB5CS41.B_LOOKUP_BIZ_PARTNER_SVR(gStrGlobalCollection, iCommandSent, imp_biz_partner_cd, E1_b_biz_partner)           									 
 								 									 
    If CheckSYSTEMError(Err,True) = True Then
       Set iB5CS41 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If      

    Set iB5CS41 = Nothing   
    Response.Write "<Script Language=vbscript>" & vbCr
    Response.Write "With parent.frm1"			& vbCr
    	
		'##### Rounding Logic #####
		'항상 거래화폐가 우선 
    Response.write ".txtDocCur1.value				= """ & ConvSPChars(E1_s_so_hdr(S308_E1_cur))			       & """" & vbCr
    Response.write ".txtDocCur2.value				= """ & ConvSPChars(E1_s_so_hdr(S308_E1_cur))			       & """" & vbCr

	Response.Write "parent.CurFormatNumericOCX" & vbCr
		'##########################

	Response.write "BPayMeth                        = """ & ConvSPChars(E1_s_so_hdr(S308_E1_pay_meth))	           & """" & vbCr
		
	Response.write ".txtXchgRate.Text			    = """ & UNINumClientFormat(E1_s_so_hdr(S308_E1_xchg_rate), ggExchRate.DecPoint, 0)	& """" & vbCr
	Response.write ".txtVatType.value			    = """ & ConvSPChars(E1_s_so_hdr(S308_E1_vat_type))	           & """" & vbCr
	Response.write ".txtVatTypeNm.value			    = """ & ConvSPChars(E1_s_so_hdr(S308_E1_vat_type_nm))	       & """" & vbCr				
	Response.write ".txtVATRate.text			    = """ & UNINumClientFormat(E1_s_so_hdr(S308_E1_vat_rate), ggExchRate.DecPoint, 0)	& """" & vbCr

	Response.write ".txtBillToPartyCd.value		    = """ & ConvSPChars(E1_s_so_hdr(S308_E1_bill_to_party))	       & """" & vbCr
	Response.write ".txtBillToPartyNm.value		    = """ & ConvSPChars(E1_s_so_hdr(S308_E1_bill_to_party_nm))	   & """" & vbCr	

	Response.write ".txtSoldtoPartyCd.value		    = """ & ConvSPChars(E1_s_so_hdr(S308_E1_sold_to_party))	       & """" & vbCr
	Response.write ".txtSoldtoPartyNm.value		    = """ & ConvSPChars(E1_s_so_hdr(S308_E1_sold_to_party_nm))	   & """" & vbCr
	Response.write ".txtLocCur.value			    = """ & UCase(gCurrency)									   & """" & vbCr

	Response.write ".txtSalesGrpCd.value		    = """ & ConvSPChars(E1_s_so_hdr(S308_E1_sales_grp))	           & """" & vbCr
	Response.write ".txtSalesGrpNm.value		    = """ & ConvSPChars(E1_s_so_hdr(S308_E1_sales_grp_nm))	       & """" & vbCr
	Response.write ".txtPaytermsTxt.value		    = """ & ConvSPChars(Trim(E1_s_so_hdr(S308_E1_pay_terms_txt)))         & """" & vbCr
	Response.write ".txtPayerCd.value			    = """ & ConvSPChars(E1_s_so_hdr(S308_E1_payer))	               & """" & vbCr
	Response.write ".txtPayerNm.value			    = """ & ConvSPChars(E1_s_so_hdr(S308_E1_payer_nm))	           & """" & vbCr
	Response.write ".txtToBizAreaCd.value		    = """ & ConvSPChars(E1_s_so_hdr(S308_E1_to_biz_grp))	       & """" & vbCr
	Response.write ".txtToBizAreaNm.value		    = """ & ConvSPChars(E1_s_so_hdr(S308_E1_to_biz_grp_nm))	       & """" & vbCr
	Response.write ".txtPayTypeCd.value			    = """ & ConvSPChars(E1_s_so_hdr(S308_E1_pay_type))	           & """" & vbCr
	Response.write ".txtPayTypeNm.value			    = """ & ConvSPChars(E1_s_so_hdr(S308_E1_pay_type_nm))	       & """" & vbCr
	Response.write ".txtPayTermsCd.value		    = BPayMeth "													      & vbCr	
	Response.write ".txtPayTermsNm.value		    = """ & ConvSPChars(Trim(E1_s_so_hdr(S308_E1_pay_meth_nm)))	   & """" & vbCr			
	'VAT포함여부 
	If Trim(ConvSPChars(E1_s_so_hdr(S308_E1_vat_inc_flag))) = "2" Then
		Response.write ".rdoVATIncFlag2.checked = True "        & vbCr	
		Response.write ".txtVatIncFlag.value = ""2""  "         & vbCr	
	Else
		Response.write ".rdoVATIncFlag1.checked = True "        & vbCr
		Response.write ".txtVatIncFlag.value = ""1""   "        & vbCr	
	End If

	If E1_s_so_hdr(S308_E1_pay_dur) = 0 Then
		Response.write ".txtPayDur.Text		= """""   & vbCr	
	Else
		Response.write ".txtPayDur.Text		= """ & E1_s_so_hdr(S308_E1_pay_dur)          & """" & vbCr			 
	End If

	Response.write ".txtBillCommand.value = """""     & vbCr	

	 '반품 여부 
	Response.write ".txtRetItemFlag.value	 = """ & E1_s_so_hdr(S308_E1_ret_item_flag)  & """"   & vbCr	

		'약정회전일 
	Response.Write ".txtCreditRotDay.value   = """ & E1_b_biz_partner(S074_E1_credit_rot_day) & """" & vbCr
		
		'수금만기일 계산 
	Response.Write "parent.CalcPlanIncomeDt "         & vbCr
		'세금신고사업장 Fetch 
	Response.Write "parent.GetTaxBizArea(""*"") "     & vbCr
	Response.write "parent.SOHdrQueryOK()       "     & vbCr
	
	If UCase(ConvSPChars(E1_s_so_hdr(S308_E1_cur))) = UCase(gCurrency) Then
		Response.write "Parent.ggoOper.SetReqAttr .txtXchgRate, ""Q"""      & vbCr
	Else
		Response.Write "Parent.ggoOper.SetReqAttr .txtXchgRate, ""N"""      & vbCr
	End If

	Response.Write "End With                  "    & vbCr
    Response.Write "</Script>                 "     & vbCr 
	    
End Sub

Sub SubBizLCNoHdr()                                                                '☜: 현재 LC헤더관련조회를 요청받음 
	  
	Dim iCommandSent
    Dim iS4G119
    Dim I1_s_lc_hdr
	Dim E1_b_biz_partner
	Dim E2_b_bank
	Dim E3_b_bank
	Dim E4_b_bank
	Dim E5_b_bank
	Dim E6_b_bank
	Dim E7_b_sales_grp
	Dim E8_b_sales_org
	Dim E9_b_biz_partner
	Dim E10_b_biz_partner
	Dim E11_b_biz_partner
	Dim E12_b_biz_partner
	Dim E13_b_biz_partner
	Dim E14_b_minor
	Dim E15_b_minor
	Dim E16_b_minor
	Dim E17_b_minor
	Dim E18_b_minor
	Dim E19_b_minor
	Dim E20_b_minor
	Dim E21_b_country
	Dim E22_b_minor
	Dim E23_b_minor
	Dim E24_b_minor
	Dim E25_b_minor
	Dim E26_s_lc_hdr
	
    Dim iB5CS41
	Dim imp_biz_partner_cd
    Dim E1_b_biz_partner2
    
	Const B253_E1_std_rate = 0
	const B253_E1_multi_divide = 1
	
	Const S357_I1_lc_no = 0    
	Const S357_I1_lc_kind = 1
	
	Const S357_E1_bp_nm = 0    'exp_consignee b_biz_partner
	Const S357_E2_bank_cd = 0  'exp_issue_bank b_bank
	Const S357_E2_bank_nm = 1
	Const S357_E3_bank_cd = 0  'exp_advise_bank b_bank
	Const S357_E3_bank_nm = 1
	Const S357_E4_bank_cd = 0  'exp_renego_bank b_bank
	Const S357_E4_bank_nm = 1
	Const S357_E5_bank_cd = 0  'exp_pay_bank b_bank
	Const S357_E5_bank_nm = 1
	Const S357_E6_bank_cd = 0  'exp_confirm_bank b_bank
	Const S357_E6_bank_nm = 1
	Const S357_E7_sales_grp_nm = 0    'exp b_sales_grp
	Const S357_E7_sales_grp = 1
	Const S357_E8_sales_org_nm = 0    'exp b_sales_org
	Const S357_E8_sales_org = 1
	Const S357_E9_bp_nm = 0    'exp_beneficiary b_biz_partner
	Const S357_E9_bp_cd = 1
	Const S357_E10_bp_nm = 0   'exp_applicant b_biz_partner
	Const S357_E10_bp_cd = 1
	Const S357_E11_bp_nm = 0   'exp_agent b_biz_partner
	Const S357_E12_bp_nm = 0   'exp_manufacturer b_biz_partner
	Const S357_E13_bp_nm = 0    'View Name : exp_notify_party b_biz_partner
	Const S357_E14_minor_nm = 0    'View Name : exp_incoterms_nm b_minor
	Const S357_E15_minor_nm = 0    'View Name : exp_pay_meth_nm b_minor
	Const S357_E16_minor_nm = 0    'View Name : exp_lc_type_nm b_minor
	Const S357_E17_minor_nm = 0    'View Name : exp_loading_port_nm b_minor
	Const S357_E18_minor_nm = 0    'View Name : exp_discharge_port_nm b_minor
	Const S357_E19_minor_nm = 0    'View Name : exp_transport_nm b_minor
	Const S357_E20_minor_nm = 0    'View Name : exp_origin_nm b_minor
	Const S357_E21_country_nm = 0    'View Name : exp_origin_cntry_nm b_country
	Const S357_E22_minor_nm = 0    'View Name : exp_charge_cd_nm b_minor
	Const S357_E23_minor_nm = 0    'View Name : exp_credit_core_nm b_minor
	Const S357_E24_minor_nm = 0    'View Name : exp_freight_nm b_minor
	Const S357_E25_minor_nm = 0    'View Name : exp_llc_type_nm b_minor
	
	Const S357_E26_lc_no = 0    'View Name : exp s_lc_hdr
	Const S357_E26_lc_doc_no = 1
	Const S357_E26_lc_amend_seq = 2
	Const S357_E26_so_no = 3
	Const S357_E26_adv_no = 4
	Const S357_E26_pre_adv_ref = 5
	Const S357_E26_adv_dt = 6
	Const S357_E26_open_dt = 7
	Const S357_E26_expiry_dt = 8
	Const S357_E26_amend_dt = 9
	Const S357_E26_manufacturer = 10
	Const S357_E26_agent = 11
	Const S357_E26_cur = 12
	Const S357_E26_lc_amt = 13
	Const S357_E26_xch_rate = 14
	Const S357_E26_lc_loc_amt = 15
	Const S357_E26_bank_txt = 16
	Const S357_E26_incoterms = 17
	Const S357_E26_pay_meth = 18
	Const S357_E26_payment_txt = 19
	Const S357_E26_latest_ship_dt = 20
	Const S357_E26_shipment = 21
	Const S357_E26_doc1 = 22
	Const S357_E26_doc2 = 23
	Const S357_E26_doc3 = 24
	Const S357_E26_doc4 = 25
	Const S357_E26_doc5 = 26
	Const S357_E26_file_dt = 27
	Const S357_E26_file_dt_txt = 28
	Const S357_E26_remark = 29
	Const S357_E26_lc_kind = 30
	Const S357_E26_lc_type = 31
	Const S357_E26_delivery_plce = 32
	Const S357_E26_amt_tolerance = 33
	Const S357_E26_loading_port = 34
	Const S357_E26_dischge_port = 35
	Const S357_E26_transport = 36
	Const S357_E26_transport_comp = 37
	Const S357_E26_origin = 38
	Const S357_E26_origin_cntry = 39
	Const S357_E26_charge_txt = 40
	Const S357_E26_charge_cd = 41
	Const S357_E26_credit_core = 42
	Const S357_E26_inv_cnt = 43
	Const S357_E26_bl_awb_flg = 44
	Const S357_E26_freight = 45
	Const S357_E26_notify_party = 46
	Const S357_E26_consignee = 47
	Const S357_E26_insur_policy = 48
	Const S357_E26_pack_list = 49
	Const S357_E26_l_lc_type = 50
	Const S357_E26_open_bank_txt = 51
	Const S357_E26_o_lc_doc_no = 52
	Const S357_E26_o_lc_amend_seq = 53
	Const S357_E26_o_lc_no = 54
	Const S357_E26_o_lc_expiry_dt = 55
	Const S357_E26_o_lc_loc_amt = 56
	Const S357_E26_o_lc_type = 57
	Const S357_E26_pay_dur = 58
	Const S357_E26_partial_ship_flag = 59
	Const S357_E26_biz_area = 60
	Const S357_E26_trnshp_flag = 61
	Const S357_E26_transfer_flag = 62
	Const S357_E26_cert_origin_flag = 63
	Const S357_E26_o_lc_amd_seq = 64
	Const S357_E26_sts = 65
	Const S357_E26_nego_amt = 66
	Const S357_E26_ext1_qty = 67
	Const S357_E26_ext2_qty = 68
	Const S357_E26_ext3_qty = 69
	Const S357_E26_ext1_amt = 70
	Const S357_E26_ext2_amt = 71
	Const S357_E26_ext3_qmt = 72
	Const S357_E26_ext1_cd = 73
	Const S357_E26_ext2_cd = 74
	Const S357_E26_ext3_cd = 75
	Const S357_E26_xch_rate_op = 76    
	
    'iB5CS41
    Const S074_E1_credit_rot_day = 53
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear

	ReDim I1_s_lc_hdr(1)	
    iCommandSent = "QUERY"	
	I1_s_lc_hdr(S357_I1_lc_no) =  Trim(Request("txtLCNo"))
    I1_s_lc_hdr(S357_I1_lc_kind) = "L"
       
    Set iS4G119 = Server.CreateObject("PS4G119.cSLkLcHdrSvr")
    	
	If CheckSYSTEMError(Err, True) = True Then
		Exit Sub
	End If
    		
	Call iS4G119.S_LOOKUP_LC_HDR_SVR(gStrGlobalCollection,iCommandSent,I1_s_lc_hdr, _
	E1_b_biz_partner,E2_b_bank,E3_b_bank,E4_b_bank,E5_b_bank,E6_b_bank, _
	E7_b_sales_grp,E8_b_sales_org, _
	E9_b_biz_partner,E10_b_biz_partner,E11_b_biz_partner,E12_b_biz_partner,E13_b_biz_partner, _
	E14_b_minor,E15_b_minor,E16_b_minor,E17_b_minor,E18_b_minor,E19_b_minor,E20_b_minor,E21_b_country, _
    E22_b_minor,E23_b_minor,E24_b_minor,E25_b_minor,E26_s_lc_hdr )
   
    If CheckSYSTEMError(Err,True) = True Then
		Set PS4G119 = Nothing
		Exit Sub
	End If  
     
	Set iS4G119 = Nothing  

    imp_biz_partner_cd = E10_b_biz_partner(S357_E10_bp_cd) 
 
    iCommandSent = "LOOKUP"

    Set iB5CS41 = Server.CreateObject("PB5CS41.cLookupBizPartnerSvr")    

    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If     
    
	Call iB5CS41.B_LOOKUP_BIZ_PARTNER_SVR(gStrGlobalCollection, iCommandSent, imp_biz_partner_cd, E1_b_biz_partner2)           									 
 								 									 
    If CheckSYSTEMError(Err,True) = True Then
       Set iB5CS41 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If      
  
    Set iB5CS41 = Nothing 
    	
	  
    Response.Write "<Script Language=vbscript>" & vbCr
    Response.Write "With parent.frm1"			& vbCr
    	

		'##### Rounding Logic #####
		'항상 거래화폐가 우선 
	Response.write ".txtDocCur1.value			= """ & ConvSPChars(E26_s_lc_hdr(S357_E26_cur))           & """" & vbCr
	Response.write ".txtDocCur2.value			= """ & ConvSPChars(E26_s_lc_hdr(S357_E26_cur))           & """" & vbCr
	Response.write "parent.CurFormatNumericOCX " & vbCr
		'##########################

	Response.write ".txtXchgRate.Text			= """ & UNINumClientFormat(E26_s_lc_hdr(S357_E26_xch_rate), ggExchRate.DecPoint, 0) & """" & vbCr
	Response.write ".txtSalesGrpCd.value		= """ & ConvSPChars(E7_b_sales_grp(S357_E7_sales_grp))    & """" & vbCr
	Response.write ".txtSalesGrpNm.value		= """ & ConvSPChars(E7_b_sales_grp(S357_E7_sales_grp_nm)) & """" & vbCr
	Response.write ".txtPayTermsCd.value		= """ & ConvSPChars(E26_s_lc_hdr(S357_E26_pay_meth))      & """" & vbCr
	Response.write ".txtPayTermsNm.value		= """ & ConvSPChars(E15_b_minor(S357_E15_minor_nm))       & """" & vbCr
	Response.write ".txtPaytermsTxt.value		= """ & ConvSPChars(Trim(E26_s_lc_hdr(S357_E26_payment_txt)))    & """" & vbCr		
	Response.write ".txtRemark.value			= """ & ConvSPChars(E26_s_lc_hdr(S357_E26_remark))		  & """" & vbCr

	Response.write ".txtBeneficiaryCd.value		= """ & ConvSPChars(E9_b_biz_partner(S357_E9_bp_cd))      & """" & vbCr
	Response.write ".txtBeneficiaryNm.value		= """ & ConvSPChars(E9_b_biz_partner(S357_E9_bp_nm))      & """" & vbCr
	Response.write ".txtApplicantCd.value		= """ & ConvSPChars(E10_b_biz_partner(S357_E10_bp_cd))    & """" & vbCr
	Response.write ".txtApplicantNm.value		= """ & ConvSPChars(E10_b_biz_partner(S357_E10_bp_nm))    & """" & vbCr
	Response.write ".txtLCNo.value				= """ & ConvSPChars(E26_s_lc_hdr(S357_E26_lc_no))         & """" & vbCr
	Response.write ".txtLCAmendSeq.value		= """ & ConvSPChars(E26_s_lc_hdr(S357_E26_lc_amend_seq))  & """" & vbCr
	Response.write ".txtLCDocNo.value			= """ & ConvSPChars(E26_s_lc_hdr(S357_E26_lc_doc_no))     & """" & vbCr
	Response.write ".txtVatIncFlag.value		= ""1"""	& vbCr			
	Response.write ".rdoVATIncFlag1.checked		= True "    & vbCr

	If E26_s_lc_hdr(S357_E26_pay_dur) = 0 Then 
		Response.write ".txtPayDur.Text		= """" "   & vbCr
	Else
		Response.write ".txtPayDur.Text		= """ & ConvSPChars(E26_s_lc_hdr(S357_E26_pay_dur))        & """" & vbCr
	End If

	If Trim(ConvSPChars(E26_s_lc_hdr(S357_E26_cur))) = UCase(gCurrency) Then                             
		Response.write "Parent.ggoOper.SetReqAttr .txtXchgRate, ""Q"" " & vbCr
	Else 
		Response.write "Parent.ggoOper.SetReqAttr .txtXchgRate, ""N"" " & vbCr
	End If

	'약정회전일 
	Response.Write ".txtCreditRotDay.value   = """ & E1_b_biz_partner2(S074_E1_credit_rot_day) & """" & vbCr
		
	Response.write ".txtBillCommand.value = """" "    & vbCr
		'수금만기일 계산 
	Response.write "parent.CalcPlanIncomeDt "        & vbCr
		'세금신고사업장 Fetch 
	Response.write "parent.GetTaxBizArea(""*"") "    & vbCr
		
	Response.Write "End With                  "      & vbCr
    Response.Write "</Script>                 "      & vbCr 
End Sub

Sub SubBizPostFlag()
	Dim iS7G115
	Dim itxtBillNo
	
 	On Error Resume Next
	Err.Clear 
    Set iS7G115 = Server.CreateObject("PS7G115.cSPostOpenArSvr")
   
	If CheckSYSTEMError(Err,True) = True Then		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  
	
	itxtBillNo = Trim(Request("txtBillNo"))
	'for 구주Tax
    '@@@@@@@@@@@
     pvCB = "F"
    '@@@@@@@@@@@
    Call iS7G115.S_POST_OPEN_AR_SVR(pvCB, gStrGlobalCollection,itxtBillNo)

	If CheckSYSTEMError(Err,True) = True Then
       Call SetErrorStatus                                                           '☆: Mark that error occurs
       Set iS7G115 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  
    
    Set iS7G115 = Nothing 
	Response.Write "<Script Language=vbscript>" & vbCr
    Response.Write "parent.DbSaveOk()         " & vbCr
    Response.Write "</Script>"                  & vbCr
  
End Sub

'============================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================
Sub SubBizSaveSingleCreate()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
	
End Sub

'============================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================
Sub SubBizSaveSingleUpdate()
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    
End Sub

'============================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================
Sub SubMakeSQLStatements(pMode)
    On Error Resume Next
    
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================
Sub CommonOnTransactionAbort()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================
Sub SetErrorStatus()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub
%>

<SCRIPT LANGUAGE=VBSCRIPT RUNAT=SERVER>
'============================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================
Sub SubMakeSQLStatements()
	Dim strSelectList, strFromList
	
	strSelectList = "SELECT * "
	strFromList = " FROM dbo.ufn_s_GetBillHdrInfo ( " & FilterVar(UCase(Request("txtConBillNo")), "''", "S") & " , " & FilterVar("N", "''", "S") & " , " & FilterVar("N", "''", "S") & " , " & FilterVar("Y", "''", "S") & " , " & FilterVar("Y", "''", "S") & " , '" & Request("txtPrevNext") & "' , " & FilterVar("N", "''", "S") & " )"
	lgstrsql = strSelectList & strFromList
	'call svrmsgbox(lgstrsql, vbinformation, i_mkscript)
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub
</SCRIPT>
