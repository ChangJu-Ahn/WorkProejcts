<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 매출채권관리 
'*  3. Program ID           : S5111MB2
'*  4. Program Name         : 예외매출채권등록 
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
'*                            -2001/12/26 : VAT 개별통합 추가 
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

    Call HideStatusWnd                                                               '☜: Hide Processing message
    'for 구주Tax
	'@@@@@@@@@@@
	Dim pvCB 
	'@@@@@@@@@@@

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
        Case CStr("PostFlag")		     											'☜: 확정 요청 
		     Call SubBizPostFlag
        Case CStr("BillLookUp")														
             Call SubBizBillLookUp
        Case CStr("BillQuery")	     												
             Call SubBizBillQuery
        Case CStr("BLQuery")		     											
		     Call SubBizBLQuery
		     
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
    Dim lgArrGlFlag
    Dim lgStrGlFlag  
    Err.Clear                                                               '☜: Protect system from crashing
    
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
		
	Response.Write ".txtDocCur1.value		= """ & UCase(lgCurrency)				      & """" & vbCr
	Response.Write ".txtDocCur2.value	    = """ & UCase(lgCurrency)				      & """" & vbCr
	Response.Write " parent.CurFormatNumericOCX " & vbCr 		


		'##########################
	Response.Write ".txtBillDt.Text			= """ & UNIDateClientFormat(lgObjRs("bill_dt"))   & """" & vbCr				

	Response.Write ".txtPlanIncomeDt.Text	= """ & UNIDateClientFormat(lgObjRs("Income_plan_dt"))   & """" & vbCr				

	Response.Write ".txtConBillNo.value			= """ & ConvSPChars(lgObjRs("bill_no"))   & """" & vbCr
	Response.Write ".txtBillNo.value			= """ & ConvSPChars(lgObjRs("bill_no"))   & """" & vbCr				
	Response.Write ".txtHBillNo.value			= """ & ConvSPChars(lgObjRs("bill_no"))   & """" & vbCr				
	Response.Write ".txtRefBillNo.Value			= """ & Trim(ConvSPChars(lgObjRs("so_no")))   & """" & vbCr				
		
	If Trim(ConvSPChars(lgObjRs("so_no"))) <> "" Then Response.Write ".chkRefBillNoFlg.Checked = True " & vbCr	
		
	Response.Write ".txtTaxBillNo.value			= """ & ConvSPChars(lgObjRs("tax_bill_no"))   & """" & vbCr				
	Response.Write ".txtPayerCd.value			= """ & ConvSPChars(lgObjRs("payer"))   & """" & vbCr				
	Response.Write ".txtPayerNm.value			= """ & ConvSPChars(lgObjRs("payer_nm"))   & """" & vbCr
		'수금그룹				
	Response.Write ".txtToBizAreaCd.value		= """ & ConvSPChars(lgObjRs("to_sales_grp"))   & """" & vbCr				
	Response.Write ".txtToBizAreaNm.value		= """ & ConvSPChars(lgObjRs("to_sales_grp_nm"))   & """" & vbCr				
	Response.Write ".txtTaxBizAreaCd.value		= """ & ConvSPChars(lgObjRs("tax_biz_area"))   & """" & vbCr					
	Response.Write ".txtTaxBizAreaNm.value		= """ & ConvSPChars(lgObjRs("tax_biz_area_nm"))   & """" & vbCr		

	Response.Write ".txtIncomeAmt.Text			= """ & UNIConvNumDBToCompanyByCurrency(lgObjRs("collect_amt"), lgCurrency, ggAmtOfMoneyNo, "X" , "X")   & """" & vbCr
	Response.Write ".txtIncomeLocAmt.Text		= """ & UNIConvNumDBToCompanyByCurrency(lgObjRs("collect_amt_loc"), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")   & """" & vbCr

	Response.Write ".txtBillAmt.Text			= """ & UNIConvNumDBToCompanyByCurrency(lgObjRs("bill_amt"), lgCurrency, ggAmtOfMoneyNo, "X" , "X")   & """" & vbCr
	Response.Write ".txtBillAmtLoc.Text			= """ & UNIConvNumDBToCompanyByCurrency(lgObjRs("bill_amt_loc"),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")   & """" & vbCr
		
	Response.Write ".txtVatAmt.Text				= """ & UNIConvNumDBToCompanyByCurrency(lgObjRs("vat_amt"), lgCurrency, ggAmtOfMoneyNo, gTaxRndPolicyNo, "X")   & """" & vbCr
	Response.Write ".txtVatLocAmt.Text			= """ & UNIConvNumDBToCompanyByCurrency(lgObjRs("vat_amt_loc"), gCurrency, ggAmtOfMoneyNo, gTaxRndPolicyNo, "X")   & """" & vbCr
		
	Response.Write ".txtDepositAmt.Text			= """ & UNIConvNumDBToCompanyByCurrency(lgObjRs("deposit_amt"), lgCurrency, ggAmtOfMoneyNo, "X" , "X")   & """" & vbCr
	Response.Write ".txtDepositAmtLoc.Text		= """ & UNIConvNumDBToCompanyByCurrency(lgObjRs("deposit_amt_loc"), gCurrency, ggAmtOfMoneyNo, gLocRndPolicyNo, "X")   & """" & vbCr

	Response.Write ".txtTotBillAmt.Text			= """ & UNIConvNumDBToCompanyByCurrency(Cdbl(lgObjRs("bill_amt")) + Cdbl(lgObjRs("vat_amt")) + Cdbl(lgObjRs("deposit_amt")), lgCurrency, ggAmtOfMoneyNo, "X" , "X")   & """" & vbCr
	Response.Write ".txtTotBillAmtLoc.Text		= """ & UNIConvNumDBToCompanyByCurrency(Cdbl(lgObjRs("bill_amt_loc")) + Cdbl(lgObjRs("vat_amt_loc")) + Cdbl(lgObjRs("deposit_amt_loc")), gCurrency, ggAmtOfMoneyNo, gLocRndPolicyNo, "X")   & """" & vbCr

	Response.Write ".txtPayTypeCd.value			= """ & ConvSPChars(lgObjRs("pay_type"))   & """" & vbCr				
	Response.Write ".txtPayTypeNm.value			= """ & ConvSPChars(lgObjRs("pay_type_nm"))   & """" & vbCr					
	Response.Write ".txtPayTermsCd.value		= """ & ConvSPChars(lgObjRs("pay_meth"))   & """" & vbCr		
	Response.Write ".txtPayTermsNm.value		= """ & ConvSPChars(lgObjRs("pay_meth_nm"))   & """" & vbCr				

	If Trim(lgObjRs("pay_dur")) = "0" Then
	    Response.Write ".txtPayDur.Text		= """" " & vbCr
	Else
	    Response.Write ".txtPayDur.Text		= """ & lgObjRs("pay_dur")   & """" & vbCr
	End If

	Response.Write ".txtXchgRate.Text			= """ & UNINumClientFormat(lgObjRs("Xchg_rate"), ggExchRate.DecPoint, 0)   & """" & vbCr
	Response.Write ".txtSoldtoPartyCd.value		= """ & ConvSPChars(lgObjRs("sold_to_party"))   & """" & vbCr	
	Response.Write ".txtSoldtoPartyNm.value		= """ & ConvSPChars(lgObjRs("sold_to_party_nm"))   & """" & vbCr	
	Response.Write ".txtBillToPartyCd.value		= """ & ConvSPChars(lgObjRs("bill_to_party"))   & """" & vbCr		
	Response.Write ".txtBillToPartyNm.value		= """ & ConvSPChars(lgObjRs("bill_to_party_nm"))   & """" & vbCr				
	Response.Write ".txtVatType.value			= """ & ConvSPChars(lgObjRs("vat_type"))   & """" & vbCr				
	Response.Write ".txtVatTypeNm.value			= """ & ConvSPChars(lgObjRs("vat_type_nm"))   & """" & vbCr
		
	Response.Write ".txtVatRate.Text			= """ & UNINumClientFormat(lgObjRs("vat_rate"), ggExchRate.DecPoint, 0)   & """" & vbCr
		
	Response.Write ".txtSalesGrpCd.value		= """ & ConvSPChars(lgObjRs("sales_grp"))   & """" & vbCr				
	Response.Write ".txtSalesGrpNm.value		= """ & ConvSPChars(lgObjRs("sales_grp_nm"))   & """" & vbCr				
	Response.Write ".txtBillTypeCd.value		= """ & ConvSPChars(lgObjRs("bill_type"))   & """" & vbCr				
	Response.Write ".txtBillTypeNm.value		= """ & ConvSPChars(lgObjRs("bill_type_nm"))   & """" & vbCr				

	Response.Write ".txtLocCur.value            = """ & UCase(gCurrency)                       & """" & vbCr

		 'VATAutoCheck
	Response.Write ".txtchkTaxNo.value	= """ & lgObjRs("vat_auto_flag")   & """" & vbCr				
	If lgObjRs("vat_auto_flag")  = "Y" Then
	    Response.Write ".chkTaxNo.checked = True           " & vbCr
		Response.Write "Call parent.chkTaxNo_OnClick()     " & vbCr
	End If

		'부가세적용기준 
	If Trim(lgObjRs("vat_calc_type"))    = "1" Then
		Response.Write ".rdoVatCalcType1.Checked = True     " & vbCr
	Else
		Response.Write ".rdoVatCalcType2.Checked = True     " & vbCr
	End If
		
	Response.Write ".txtVatCalcType.value		= """ & lgObjRs("vat_calc_type")   & """" & vbCr
	Response.Write ".txtPaytermsTxt.value		= """ & ConvSPChars(lgObjRs("pay_terms_txt"))   & """" & vbCr				
	Response.Write ".txtRemark.value			= """ & ConvSPChars(lgObjRs("remark"))   & """" & vbCr				

		'VAT포함여부 
	If Trim(lgObjRs("vat_inc_flag")) = "1" Then
	    Response.Write ".rdoVATIncFlag1.checked = True   " & vbCr
	Else
		Response.Write ".rdoVATIncFlag2.checked = True   " & vbCr
	End If
		Response.Write ".txtVatIncflag.value		= """ & lgObjRs("vat_inc_flag")   & """" & vbCr

		Response.Write ".txtHRefFlag.value			= """ & lgObjRs("ref_flag")   & """" & vbCr				

		 '매출진행상태 
		Response.Write ".txtSts.value	= """ & lgObjRs("sts")   & """" & vbCr
		
		'약정회전일 
		Response.Write ".txtCreditRotDay.value		= """ & lgObjRs("credit_rot_day")   & """" & vbCr	
	
	If lgObjRs("post_flag") = "Y" AND Len(Trim(lgObjRs("gl_no"))) Then
		lgArrGlFlag    =  Split(lgObjRs("gl_no"), Chr(11))      
		lgStrGlFlag    =  lgArrGlFlag(0) 
		
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
		
	If lgObjRs("post_flag")   = "Y" Then
		Response.Write "parent.PostFlagProtect()           " & vbCr
		Response.Write ".rdoPostFlagY.checked = True       " & vbCr
		Response.Write ".btnPostFlag.value = ""확정취소"" " & vbCr
		if lgStrGlFlag= "G" Or lgStrGlFlag = "T" Then
		    Response.Write ".btnGLView.disabled = False    " & vbCr
		Else
		    Response.Write ".btnGLView.disabled = True     " & vbCr
		End If
	Else	
		Response.Write "parent.PostFlagRelease()           " & vbCr
		Response.Write ".rdoPostFlagN.checked = True	   " & vbCr	
		Response.Write ".btnPostFlag.value = ""확정""  " & vbCr
		Response.Write ".btnGLView.disabled = True         " & vbCr

		If UCase(Trim(.txtDocCur1.value)) = UCase(gCurrency) Then
			Response.Write "Parent.ggoOper.SetReqAttr .txtXchgRate, ""Q"" " & vbCr
		Else
			Response.Write "Parent.ggoOper.SetReqAttr .txtXchgRate, ""N"" " & vbCr
		End If
	End If

		 '선수금 현황 버튼 Enable 
	IF lgObjRs("PreRcpt_flag")   = "Y" Then
		Response.Write ".btnPreRcptView.disabled = False    " & vbCr
	Else
		Response.Write ".btnPreRcptView.disabled = True     " & vbCr
	End If 
	
	
	Response.Write ".txtHExportFlag.value		= """ & lgObjRs("bl_flag")   & """" & vbCr	
	
	Response.Write "parent.DbQueryOk		  " & vbCr													'☜: 조회가 성공 
	Response.Write "parent.lgBlnFlgChgValue = False         " & vbCr        				'☜: 조회가 성공 
    
    Response.Write "End With"          & vbCr
    Response.Write "</Script>"         & vbCr
	


	lgObjRs.Close
	lgObjConn.Close
	Set lgObjRs = Nothing
	Set lgObjConn = Nothing
										
End Sub	

'============================================
' Name : SubBizSave
' Desc : Date data 
'============================================
Sub SubBizSave()
	On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear 

	Dim iS7G111
	Dim lgIntFlgMode
    Dim iCommandSent
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
    
    Const S500_I6_bill_no = 0    '  View Name : imp s_bill_hdr
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

	'<--예외매출채권번호 -->pis
	I6_s_bill_hdr(S500_I6_bill_no) = UCase(Trim(Request("txtBillNo")))
	'<--이전매출채권번호 -->
	IF Trim(Request("txtRefBillNoFlg")) = "Y" Then I6_s_bill_hdr(S500_I6_so_no) = UCase(Trim(Request("txtRefBillNo")))
	'<--매출채권일 -->
	I6_s_bill_hdr(S500_I6_bill_dt) = UNIConvDate(Request("txtBillDt"))
	'<--수금예정일 -->
	I6_s_bill_hdr(S500_I6_income_plan_dt) = UNIConvDate(Request("txtPlanIncomeDt"))
	'<--선수금번호 -->
    'I6_s_bill_hdr(ImpSBillHdrPreRcptNo = TRIM(Request("txtPreRcpNo"))
	'<--수금처 -->pis
    I8_b_biz_partner = UCase(Trim(Request("txtPayerCd")))
	'<--주문처 -->pis
    I3_b_biz_partner =  UCase(Trim(Request("txtSoldtoPartyCd")))

	'<--수금영업그룹 -->
	I4_b_sales_grp = UCase(Trim(Request("txtToBizAreaCd")))
	'<--입금유형 -->
    I6_s_bill_hdr(S500_I6_pay_type) = UCase(Trim(Request("txtPayTypeCd")))
	'<--결제방법 -->
    I6_s_bill_hdr(S500_I6_pay_meth) = UCase(Trim(Request("txtPayTermsCd")))

	'<--결제기간 -->
	If Request("txtPayDur") = "" Then
	    I6_s_bill_hdr(S500_I6_pay_dur) = 0
	Else
	    I6_s_bill_hdr(S500_I6_pay_dur) = Trim(Request("txtPayDur"))
	End If    

	'<--환율 -->
	If Len(Trim(Request("txtXchgRate"))) Then I6_s_bill_hdr(S500_I6_xchg_rate) = UNIConvNum(Trim(Request("txtXchgRate")), 0)
	'<--발행처 -->
    I7_b_biz_partner = UCase(Trim(Request("txtBillToPartyCd")))
	'<--부가세유형 -->
	I6_s_bill_hdr(S500_I6_vat_type) = UCase(Trim(Request("txtVatType")))
	'<--부가세율 -->
    If Len(Trim(Request("txtVatRate"))) Then I6_s_bill_hdr(S500_I6_vat_rate) = UNIConvNum(Trim(Request("txtVatRate")), 0)
	'<--영업그룹 -->
    I5_b_sales_grp = UCase(Trim(Request("txtSalesGrpCd")))
	'<--매출채권타입 -->
	I2_s_bill_type_config = UCase(Trim(Request("txtBillTypeCd")))
	'<--화폐 -->
	I6_s_bill_hdr(S500_I6_cur) = UCase(Trim(Request("txtDocCur1")))
	'<--대금결제조건 -->
    I6_s_bill_hdr(S500_I6_pay_terms_txt) = UCase(Trim(Request("txtPaytermsTxt")))
	'<--비고 -->
	I6_s_bill_hdr(S500_I6_remark) = UCase(Trim(Request("txtRemark")))

	'<--세금계산서 자동발행 -->
	I6_s_bill_hdr(S500_I6_vat_auto_flag) = Trim(Request("txtchkTaxNo"))
	'<--부가세적용기준(개별/통합)-->
	I6_s_bill_hdr(S500_I6_vat_calc_type) = Trim(Request("txtVatCalcType"))
	'<--세금신고사업장 -->
	I6_s_bill_hdr(S500_I6_tax_biz_area) = UCase(Trim(Request("txtTaxBizAreaCd")))
	'<--세금계산서번호 -->
	I6_s_bill_hdr(S500_I6_tax_bill_no) = UCase(Trim(Request("txtTaxBillNo")))
	'<--부가세포함구분(별도/포함)-->
	I6_s_bill_hdr(S500_I6_vat_inc_flag) = Trim(Request("txtVatIncFlag"))
	
	'<--REF_FLAG값 -->
	I6_s_bill_hdr(S500_I6_ref_flag) = Trim(Request("txtHRefFlag"))	

	'=========== Common Value ===========	

	'---> 매출채권인경우 "N", 예외매출채권인경우 "Y"
    I6_s_bill_hdr(S500_I6_except_flag) = "Y"
	'---> B/L 여부인 (예외매출인 경우는 항상 N)
	
    I6_s_bill_hdr(S500_I6_bl_flag) = Trim(Request("txtHExportFlag"))
   	
	
	'--->반품여부 
	lgIntFlgMode = CInt(Request("txtFlgMode"))										'☜: 저장시 Create/Update 판별  

    If lgIntFlgMode = OPMD_CMODE Then
		'<--TransType 타입 -->
		I1_s_bill_hdr = Trim(Request("txtHRefBillNo"))
		iCommandSent = "CREATE"
    ElseIf lgIntFlgMode = OPMD_UMODE Then
		iCommandSent = "UPDATE"
    End If

    Set iS7G111 = Server.CreateObject("PS7G111.cSBillHdrSvr")

	If CheckSYSTEMError(Err,True) = True Then		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  
	
	'for 구주Tax
    '@@@@@@@@@@@
     pvCB = "F"
    '@@@@@@@@@@@
    E2_s_bill_hdr = ""
    E2_s_bill_hdr = iS7G111.S_MAINT_BILL_HDR_SVR(pvcB,gStrGlobalCollection, iCommandSent,I1_s_bill_hdr,I2_s_bill_type_config , _
              I3_b_biz_partner , I4_b_sales_grp ,I5_b_sales_grp , I6_s_bill_hdr , I7_b_biz_partner , _
              I8_b_biz_partner , I9_b_sales_org  , "" )


	If CheckSYSTEMError(Err,True) = True Then
       Call SetErrorStatus                                                           '☆: Mark that error occurs
       Set iS7G111 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  

    Set iS7G111 = Nothing


	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent"           & vbCr

	If ConvSPChars(E2_s_bill_hdr) <>"" Then
	    Response.Write ".frm1.txtConBillNo.value = """ & ConvSPChars(E2_s_bill_hdr)     & """" & vbCr
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

Sub SubBizPostFlag()                                                '☜: 확정 처리 
	Dim iS7G115
	Dim itxtConBillNo
	
 	On Error Resume Next
	Err.Clear 
    Set iS7G115 = Server.CreateObject("PS7G115.cSPostOpenArSvr")
   
	If CheckSYSTEMError(Err,True) = True Then		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  
	
	itxtConBillNo = Trim(Request("txtConBillNo"))
	
	'for 구주Tax
    '@@@@@@@@@@@
     pvCB = "F"
    '@@@@@@@@@@@
    Call iS7G115.S_POST_OPEN_AR_SVR(pvCB,gStrGlobalCollection,itxtConBillNo)

	If CheckSYSTEMError(Err,True) = True Then
       Call SetErrorStatus                                                           '☆: Mark that error occurs
       Set iS7G115 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  
    
    Set iS7G115 = Nothing 
	Response.Write "<Script Language=vbscript>" & vbCr
    Response.Write "parent.DbSaveOk()         " & vbCr
    Response.Write "</Script>"                  & vbCr
    

End  Sub


Sub SubBizBillLookUp()                                              '☜: 현재 주문처 거래 조회 요청을 받음 
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear 
    
    Dim iB5CS41
	Dim imp_biz_partner_cd
    Dim E1_b_biz_partner
	Dim iCommandSent
	
	Dim iB5GS45
	Dim I1_b_biz_partner 
	Dim E1_b_biz_partner2
	Dim E2_b_biz_partner 
	Dim E3_b_biz_partner 
	Dim E4_b_biz_partner 
	Dim E5_b_biz_partner 
	Dim E6_b_biz_partner 
	'iB5CS41
    Const S074_E1_bp_cd = 0
    Const S074_E1_bp_type = 1
    Const S074_E1_bp_rgst_no = 2
    Const S074_E1_bp_full_nm = 3
    Const S074_E1_bp_nm = 4
    Const S074_E1_bp_eng_nm = 5
    Const S074_E1_repre_nm = 6
    Const S074_E1_repre_rgst_no = 7
    Const S074_E1_fnd_dt = 8
    Const S074_E1_zip_cd = 9
    Const S074_E1_addr1 = 10
    Const S074_E1_addr1_eng = 11
    Const S074_E1_ind_type = 12
    Const S074_E1_ind_class = 13
    Const S074_E1_trade_rgst_no = 14
    Const S074_E1_contry_cd = 15
    Const S074_E1_province_cd = 16
    Const S074_E1_currency = 17
    Const S074_E1_tel_no1 = 18
    Const S074_E1_tel_no2 = 19
    Const S074_E1_fax_no = 20
    Const S074_E1_home_url = 21
    Const S074_E1_usage_flag = 22
    Const S074_E1_bp_prsn_nm = 23
    Const S074_E1_bp_contact_pt = 24
    Const S074_E1_biz_prsn = 25
    Const S074_E1_biz_grp = 26
    Const S074_E1_biz_org = 27
    Const S074_E1_deal_type = 28
    Const S074_E1_pay_meth = 29
    Const S074_E1_pay_dur = 30
    Const S074_E1_pay_day = 31
    Const S074_E1_vat_inc_flag = 32
    Const S074_E1_vat_type = 33
    Const S074_E1_vat_rate = 34
    Const S074_E1_trans_meth = 35
    Const S074_E1_trans_lt = 36
    Const S074_E1_sale_amt = 37
    Const S074_E1_capital_amt = 38
    Const S074_E1_emp_cnt = 39
    Const S074_E1_bp_grade = 40
    Const S074_E1_comm_rate = 41
    Const S074_E1_addr2 = 42
    Const S074_E1_addr2_eng = 43
    Const S074_E1_addr3_eng = 44
    Const S074_E1_pay_type = 45
    Const S074_E1_pay_terms_txt = 46
    Const S074_E1_credit_mgmt_flag = 47
    Const S074_E1_credit_grp = 48
    Const S074_E1_vat_calc_type = 49
    Const S074_E1_deposit_flag = 50
    Const S074_E1_bp_group = 51
    Const S074_E1_clearance_id = 52
    Const S074_E1_credit_rot_day = 53
    Const S074_E1_gr_insp_type = 54
    Const S074_E1_bp_alias_nm = 55
    Const S074_E1_to_org = 56
    Const S074_E1_to_grp = 57
    Const S074_E1_pay_month = 58
    Const S074_E1_expiry_dt = 59
    Const S074_E1_pur_grp = 60
    Const S074_E1_pur_org = 61
    Const S074_E1_charge_lay_flag = 62
    Const S074_E1_remark1 = 63
    Const S074_E1_remark2 = 64
    Const S074_E1_remark3 = 65
    Const S074_E1_close_day1 = 66
    Const S074_E1_close_day2 = 67
    Const S074_E1_close_day3 = 68
    Const S074_E1_tax_biz_area = 69
    Const S074_E1_cash_rate = 70
    Const S074_E1_pay_type_out = 71
    Const S074_E1_par_bank_cd1_bp = 72
    Const S074_E1_bank_acct_no1_bp = 73
    Const S074_E1_bank_cd1_bp = 74
    Const S074_E1_par_bank_cd2_bp = 75
    Const S074_E1_bank_cd2_bp = 76
    Const S074_E1_bank_acct_no2_bp = 77
    Const S074_E1_par_bank_cd3_bp = 78
    Const S074_E1_bank_cd3_bp = 79
    Const S074_E1_bank_acct_no3_bp = 80
    Const S074_E1_par_bank_cd1 = 81
    Const S074_E1_bank_cd1 = 82
    Const S074_E1_bank_acct_no1 = 83
    Const S074_E1_par_bank_cd2 = 84
    Const S074_E1_bank_cd2 = 85
    Const S074_E1_bank_acct_no2 = 86
    Const S074_E1_par_bank_cd3 = 87
    Const S074_E1_bank_cd3 = 88
    Const S074_E1_bank_acct_no3 = 89
    Const S074_E1_pay_month2 = 90
    Const S074_E1_pay_day2 = 91
    Const S074_E1_pay_month3 = 92
    Const S074_E1_pay_day3 = 93
    Const S074_E1_close_day1_sales = 94
    Const S074_E1_pay_month1_sales = 95
    Const S074_E1_pay_day1_sales = 96
    Const S074_E1_close_day2_sales = 97
    Const S074_E1_pay_month2_sales = 98
    Const S074_E1_pay_day2_sales = 99
    Const S074_E1_close_day3_sales = 100
    Const S074_E1_pay_month3_sales = 101
    Const S074_E1_pay_day3_sales = 102
    Const S074_E1_ext1_qty = 103
    Const S074_E1_ext2_qty = 104
    Const S074_E1_ext3_qty = 105
    Const S074_E1_ext1_amt = 106
    Const S074_E1_ext2_amt = 107
    Const S074_E1_ext3_amt = 108
    Const S074_E1_ext1_cd = 109
    Const S074_E1_ext2_cd = 110
    Const S074_E1_ext3_cd = 111
    Const S074_E1_in_out = 112                                '[--사내외구분]
    Const S074_E1_card_co_cd = 113                            '카드사 
    Const S074_E1_card_mem_no = 114                           '가맹점번호 
    Const S074_E1_pay_meth_pur = 115                          '결재방법(구매)
    Const S074_E1_pay_type_pur = 116                          '입출금유형(구매)
    Const S074_E1_pay_dur_pur = 117                           '결재기간(구매)
    Const S074_E1_bank_cd = 118                               '은행 
    Const S074_E1_bank_acct_no = 119                          '계좌번호 
    Const S074_E1_ind_type_nm = 120                           '[업종명]
    Const S074_E1_ind_class_nm = 121                          '[업태명]
    Const S074_E1_bp_group_nm = 122                           '[거래처분류명]
    Const S074_E1_b_country_nm = 123                          '[국가명]
    Const S074_E1_b_province_nm = 124                         '[지방명]
    Const S074_E1_trans_meth_nm = 125                         '[운송방법명]
    Const S074_E1_deal_type_nm = 126                          '[판매유형명]
    Const S074_E1_bp_grade_nm = 127                           '[업체평가등급명]
    Const S074_E1_s_credit_limit = 128                        '[여신관리그룹명]
    Const S074_E1_b_sales_grp_nm = 129                        '[영업그룹명]
    Const S074_E1_b_to_grp_nm = 130                           '[수금그룹명]
    Const S074_E1_b_pur_grp_nm = 131                          '[구매그룹명]
    Const S074_E1_vat_type_nm = 132                           '[부가세유형명]
    Const S074_E1_pay_meth_nm = 133                           '[결재방법명]
    Const S074_E1_pay_type_nm = 134                           '[입출금유형명]
    Const S074_E1_tax_area_nm = 135                           '[세금신고사업장명]
    Const S074_E1_b_zip_code = 136                            '[--우편번호]
    Const S074_E1_b_pur_org = 137                             '[--구매조직코드]
    Const S074_E1_b_pur_org_nm = 138                          '[--구매조직명]
    Const S074_E1_vat_inc_flag_nm = 139                       '[--부과세구분명]
    Const S074_E1_card_co_cd_nm = 140                         '[카드사명]
    Const S074_E1_pay_meth_pur_nm = 141                       '[결재방법명(구매)]
    Const S074_E1_pay_type_pur_nm = 142                       '[입출금유형명(구매)]
    Const S074_E1_bank_cd_nm = 143                            '[은행명]


    'iB5GS45
    Const B132_E1_bp_cd = 0    'View Name : exp_mgs b_biz_partner
    Const B132_E1_bp_nm = 1

    Const B132_E2_bp_cd = 0    'View Name : exp_mbi b_biz_partner
    Const B132_E2_bp_nm = 1

    Const B132_E3_bp_cd = 0    'View Name : exp_mpa b_biz_partner
    Const B132_E3_bp_nm = 1

    Const B132_E4_bp_cd = 0    'View Name : exp_spa b_biz_partner
    Const B132_E4_bp_nm = 1

    Const B132_E5_bp_cd = 0    'View Name : exp_sbi b_biz_partner
    Const B132_E5_bp_nm = 1

    Const B132_E6_bp_cd = 0    'View Name : exp_ssh b_biz_partner
    Const B132_E6_bp_nm = 1
	
	On Error Resume Next
	Err.Clear   

    If Request("txtSoldtoPartyCd") = "" Then								'⊙: 조회를 위한 값이 들어왔는지 체크 
	   Call ServerMesgBox("주문처값이 비어있습니다!", vbInformation, I_MKSCRIPT)              
       Exit Sub
	End If
    imp_biz_partner_cd = Trim(Request("txtSoldtoPartyCd"))
 	  
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
	
	I1_b_biz_partner= Trim(Request("txtSoldtoPartyCd"))

    Set iB5GS45 = Server.CreateObject("PB5GS45.cBListDftBpFtnSvr")    

    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If     

   call iB5GS45.B_LIST_DEFAULT_BP_FTN_SVR( gStrGlobalCollection , _
             I1_b_biz_partner ,  E1_b_biz_partner2 ,  E2_b_biz_partner ,   E3_b_biz_partner , _
             E4_b_biz_partner ,  E5_b_biz_partner ,  E6_b_biz_partner )
             								 									 
    If CheckSYSTEMError(Err,True) = True Then
       Set iB5GS45 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If      

    Set iB5GS45 = Nothing  	
		  
    Response.Write "<Script Language=vbscript> " & vbCr
    Response.Write "With parent.frm1"			 & vbCr

		'주문처 
	Response.Write ".txtSoldtoPartyCd.value	= """ & ConvSPChars(E1_b_biz_partner(S074_E1_bp_cd))      & """" & vbCr	
	Response.Write ".txtSoldtoPartyNm.value	= """ & ConvSPChars(E1_b_biz_partner(S074_E1_bp_nm))      & """" & vbCr	
		'발행처 
	Response.Write ".txtBillToPartyCd.value	= """ & ConvSPChars(E2_b_biz_partner(B132_E2_bp_cd))      & """" & vbCr	
	Response.Write ".txtBillToPartyNm.value	= """ & ConvSPChars(E2_b_biz_partner(B132_E2_bp_nm))      & """" & vbCr	
		'결제방법 
	Response.Write ".txtPayTermsCd.value	= """ & ConvSPChars(E1_b_biz_partner(S074_E1_pay_meth))      & """" & vbCr	
	Response.Write ".txtPayTermsNm.value	= """ & ConvSPChars(E1_b_biz_partner(S074_E1_pay_meth_nm))      & """" & vbCr						
		'결제기간 
	Response.Write ".txtPayDur.text = """ & E1_b_biz_partner(S074_E1_pay_dur)      & """" & vbCr	
		'부가세유형 
	Response.Write ".txtVatType.value = """ & ConvSPChars(E1_b_biz_partner(S074_E1_vat_type))      & """" & vbCr	
	Response.Write ".txtVatTypeNm.value = """ & ConvSPChars(E1_b_biz_partner(S074_E1_vat_type_nm))      & """" & vbCr	
		'부가세율 
	'Response.Write ".txtVatRate.text	= """ & UNINumClientFormat(E1_b_biz_partner(S074_E1_vat_rate), ggExchRate.DecPoint, 0)      & """" & vbCr	
	
	Response.Write " parent.txtVatType_OnChange  " & vbCr
		'화폐 
	Response.Write ".txtDocCur1.value = """ & ConvSPChars(E1_b_biz_partner(S074_E1_currency))      & """" & vbCr	
	Response.Write ".txtDocCur2.value = """ & ConvSPChars(E1_b_biz_partner(S074_E1_currency))      & """" & vbCr	
		'수금처 
	Response.Write ".txtPayerCd.value = """ & ConvSPChars(E3_b_biz_partner(B132_E3_bp_cd))      & """" & vbCr	
	Response.Write ".txtPayerNm.value = """ & ConvSPChars(E3_b_biz_partner(B132_E3_bp_nm))      & """" & vbCr	
		'부가세적용기준 
		'.txtVatCalcType.value = """ & ConvSPChars(E1_b_biz_partner(ExpBBizPartnerVatIncFlag))      & """" & vbCr	

		'약정회전일 
	Response.Write ".txtCreditRotDay.value = """ & E1_b_biz_partner(S074_E1_credit_rot_day)       & """" & vbCr	
		
		If E1_b_biz_partner(S074_E1_credit_rot_day)	 <> "0" Then
			'수금예정일 
	Response.Write ".txtPlanIncomeDt.Text = """ & UnIDateAdd("d", E1_b_biz_partner(S074_E1_credit_rot_day), Request("txtBillDt"), gDateFormat)      & """" & vbCr	
		End If

	Response.Write "parent.CurrencyOnChange    " & vbCr
		'세금신고사업장 Fetch
	Response.Write "parent.GetTaxBizArea(""*"")  " & vbCr

	Response.Write "End With                   "    & vbCr
    Response.Write "</Script>                  "     & vbCr 
	
End  Sub

Sub SubBizBillQuery()                                       '☜: 현재 주문처 거래 조회 요청을 받음 
	On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear 
    Dim iS7G119
    Dim iCommandSend
    Dim I1_s_bill_hdr
    Dim E1_a_gl 
    Dim E2_a_temp_gl  
    Dim E3_a_batch  
    Dim E4_s_bill_type_config  
    Dim E5_b_biz_partner                     
    Dim E6_b_biz_area  
    Dim E7_b_minor  
    Dim E8_b_sales_org 
    Dim E9_b_biz_partner  
    Dim E10_b_biz_partner                       
    Dim E11_b_biz_partner  
    Dim E12_b_minor 
    Dim E13_b_sales_grp 
    Dim E14_b_biz_partner 
    Dim E15_b_sales_org  
    Dim E16_b_sales_grp 
    Dim E17_s_bill_hdr 
    Dim E18_b_minor 
    
    Const S508_I1_bill_no = 0    'View Name : imp s_bill_hdr
    Const S508_I1_except_flag = 1
    
    Const S508_E4_bill_type = 0    'View Name : exp s_bill_type_config
    Const S508_E4_bill_type_nm = 1

    Const S508_E5_bp_cd = 0    'View Name : exp_sold_to_party b_biz_partner
    Const S508_E5_bp_nm = 1
    Const S508_E5_credit_rot_day = 2

    Const S508_E8_sales_org = 0    'View Name : exp_billing b_sales_org
    Const S508_E8_sales_org_nm = 1

    Const S508_E9_bp_cd = 0    'View Name : exp_bill_to_party b_biz_partner
    Const S508_E9_bp_nm = 1

    Const S508_E13_sales_grp = 0    'View Name : exp_billing b_sales_grp
    Const S508_E13_sales_grp_nm = 1

    Const S508_E14_bp_cd = 0    'View Name : exp_payer b_biz_partner
    Const S508_E14_bp_nm = 1

    Const S508_E15_sales_org = 0    'View Name : exp_income b_sales_org
    Const S508_E15_sales_org_nm = 1

    Const S508_E16_sales_grp = 0    'View Name : exp_income b_sales_grp
    Const S508_E16_sales_grp_nm = 1

    Const S508_E17_bill_no = 0    'View Name : exp s_bill_hdr
    Const S508_E17_post_flag = 1
    Const S508_E17_trans_type = 2
    Const S508_E17_bill_dt = 3
    Const S508_E17_cur = 4
    Const S508_E17_xchg_rate = 5
    Const S508_E17_xchg_rate_op = 6
    Const S508_E17_bill_amt = 7
    Const S508_E17_vat_type = 8
    Const S508_E17_vat_rate = 9
    Const S508_E17_vat_amt = 10
    Const S508_E17_pay_meth = 11
    Const S508_E17_pay_dur = 12
    Const S508_E17_tax_bill_no = 13
    Const S508_E17_tax_prt_cnt = 14
    Const S508_E17_accept_fob_amt = 15
    Const S508_E17_beneficiary = 16
    Const S508_E17_applicant = 17
    Const S508_E17_remark = 18
    Const S508_E17_bill_amt_loc = 19
    Const S508_E17_vat_calc_type = 20
    Const S508_E17_vat_amt_loc = 21
    Const S508_E17_tax_biz_area = 22
    Const S508_E17_pay_type = 23
    Const S508_E17_pay_terms_txt = 24
    Const S508_E17_collect_amt = 25
    Const S508_E17_collect_amt_loc = 26
    Const S508_E17_income_plan_dt = 27
    Const S508_E17_nego_amt = 28
    Const S508_E17_so_no = 29
    Const S508_E17_lc_no = 30
    Const S508_E17_lc_doc_no = 31
    Const S508_E17_lc_amend_seq = 32
    Const S508_E17_bl_flag = 33
    Const S508_E17_biz_area = 34
    Const S508_E17_cost_cd = 35
    Const S508_E17_to_biz_area = 36
    Const S508_E17_to_cost_cd = 37
    Const S508_E17_ref_flag = 38
    Const S508_E17_sts = 39
    Const S508_E17_except_flag = 40
    Const S508_E17_reverse_flag = 41
    Const S508_E17_ext1_qty = 42
    Const S508_E17_ext2_qty = 43
    Const S508_E17_ext3_qty = 44
    Const S508_E17_ext1_amt = 45
    Const S508_E17_ext2_amt = 46
    Const S508_E17_ext3_amt = 47
    Const S508_E17_ext1_cd = 48
    Const S508_E17_ext2_cd = 49
    Const S508_E17_ext3_cd = 50
    Const S508_E17_vat_auto_flag = 51
    Const S508_E17_vat_inc_flag = 52
    Const S508_E17_deposit_amt = 53
    Const S508_E17_deposit_amt_loc = 54

	On Error Resume Next
	Err.Clear              

    ReDim I1_s_bill_hdr(S508_I1_except_flag)	
    
    I1_s_bill_hdr(S508_I1_bill_no) = Trim(Request("txtBillNo"))
    I1_s_bill_hdr(S508_I1_except_flag) = Trim(Request("txtExceptFlg"))
    
    Set iS7G119 = Server.CreateObject("PS7G119.cSLkBillHdrSvr")    

    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If    

    Call   iS7G119.S_LOOKUP_BILL_HDR_SVR(gStrGlobalCollection ,  "QUERY"  , I1_s_bill_hdr , _
                      E1_a_gl ,  E2_a_temp_gl ,  E3_a_batch ,  E4_s_bill_type_config ,  E5_b_biz_partner , _
                      E6_b_biz_area ,  E7_b_minor ,  E8_b_sales_org , E9_b_biz_partner ,  E10_b_biz_partner , _
                      E11_b_biz_partner ,  E12_b_minor , E13_b_sales_grp , E14_b_biz_partner , E15_b_sales_org , _
                      E16_b_sales_grp , E17_s_bill_hdr , E18_b_minor )
            				 									 
    If CheckSYSTEMError(Err,True) = True Then
       Set iS7G119 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If     				

    Set iS7G119 = Nothing  
	
    Response.Write "<Script Language=vbscript>" & vbCr
    Response.Write "With parent.frm1          "	& vbCr

		'##### Rounding Logic #####
		'항상 거래화폐가 우선 
	Response.Write ".txtDocCur1.value			= """ & ConvSPChars(E17_s_bill_hdr(S508_E17_cur))  & """" & vbCr
	Response.Write ".txtDocCur2.value			= """ & ConvSPChars(E17_s_bill_hdr(S508_E17_cur))  & """" & vbCr
	Response.Write "parent.CurFormatNumericOCX "    & vbCr
		'##########################

	Response.Write ".txtPayerCd.value			= """ & ConvSPChars(E14_b_biz_partner(S508_E14_bp_cd))  & """" & vbCr				
	Response.Write ".txtPayerNm.value			= """ & ConvSPChars(E14_b_biz_partner(S508_E14_bp_nm))  & """" & vbCr				
	Response.Write ".txtToBizAreaCd.value		= """ & ConvSPChars(E16_b_sales_grp(S508_E16_sales_grp))  & """" & vbCr				
	Response.Write ".txtToBizAreaNm.value		= """ & ConvSPChars(E16_b_sales_grp(S508_E16_sales_grp_nm))  & """" & vbCr				
	Response.Write ".txtTaxBizAreaCd.value		= """ & ConvSPChars(E17_s_bill_hdr(S508_E17_tax_biz_area))  & """" & vbCr
	Response.Write ".txtTaxBizAreaNm.value		= """ & ConvSPChars(E6_b_biz_area)  & """" & vbCr

	Response.Write ".txtIncomeAmt.Text			= 0 "    & vbCr
	Response.Write ".txtBillAmt.Text			= 0 "    & vbCr
	Response.Write ".txtBillAmtLoc.Text			= 0 "    & vbCr
	Response.Write ".txtVatLocAmt.text			= 0 "    & vbCr
	Response.Write ".txtIncomeLocAmt.text		= 0 "    & vbCr
	Response.Write ".txtVatAmt.Text				= 0 "    & vbCr

	Response.Write ".txtPayTypeCd.value			= """ & ConvSPChars(E17_s_bill_hdr(S508_E17_pay_type))  & """" & vbCr				
	Response.Write ".txtPayTypeNm.value			= """ & ConvSPChars(E18_b_minor)  & """" & vbCr					
	Response.Write ".txtPayTermsCd.value		= """ & ConvSPChars(E17_s_bill_hdr(S508_E17_pay_meth))  & """" & vbCr		
	Response.Write ".txtPayTermsNm.value		= """ & ConvSPChars(E7_b_minor)  & """" & vbCr				
	Response.Write ".txtPayDur.Text				= """ & E17_s_bill_hdr(S508_E17_pay_dur)  & """" & vbCr
	Response.Write ".txtSoldtoPartyCd.value		= """ & ConvSPChars(E5_b_biz_partner(S508_E5_bp_cd))  & """" & vbCr	
	Response.Write ".txtSoldtoPartyNm.value		= """ & ConvSPChars(E5_b_biz_partner(S508_E5_bp_nm))  & """" & vbCr	
	Response.Write ".txtBillToPartyCd.value		= """ & ConvSPChars(E9_b_biz_partner(S508_E9_bp_cd))  & """" & vbCr		
	Response.Write ".txtBillToPartyNm.value		= """ & ConvSPChars(E9_b_biz_partner(S508_E9_bp_nm))  & """" & vbCr				
	Response.Write ".txtVatType.value			= """ & ConvSPChars(E17_s_bill_hdr(S508_E17_vat_type))  & """" & vbCr				
	Response.Write ".txtVatTypeNm.value			= """ & ConvSPChars(E12_b_minor)  & """" & vbCr
	Response.Write " parent.txtVatType_OnChange  " & vbCr		
'	Response.Write ".txtVatRate.Text			= """ & UNINumClientFormat(E17_s_bill_hdr(S508_E17_vat_rate), ggExchRate.DecPoint, 0)  & """" & vbCr
		
	Response.Write ".txtSalesGrpCd.value		= """ & ConvSPChars(E13_b_sales_grp(S508_E13_sales_grp))  & """" & vbCr				
	Response.Write ".txtSalesGrpNm.value		= """ & ConvSPChars(E13_b_sales_grp(S508_E13_sales_grp_nm))  & """" & vbCr				
	Response.Write ".txtDocCur1.value			= """ & ConvSPChars(E17_s_bill_hdr(S508_E17_cur))  & """" & vbCr				
	Response.Write ".txtDocCur2.value			= """ & ConvSPChars(E17_s_bill_hdr(S508_E17_cur))  & """" & vbCr				
	Response.Write ".txtVatCalcType.value		= """ & ConvSPChars(E17_s_bill_hdr(S508_E17_vat_calc_type))  & """" & vbCr
	Response.Write ".txtVatIncFlag.value		= """ & ConvSPChars(E17_s_bill_hdr(S508_E17_vat_inc_flag))  & """" & vbCr
	Response.Write ".txtPaytermsTxt.value		= """ & ConvSPChars(E17_s_bill_hdr(S508_E17_pay_terms_txt))  & """" & vbCr				
	Response.Write ".txtRemark.value			= """ & ConvSPChars(E17_s_bill_hdr(S508_E17_remark))  & """" & vbCr				

	Response.Write ".txtHExportFlag.value		= """ & ConvSPChars(E17_s_bill_hdr(S508_E17_bl_flag))  & """" & vbCr
	Response.Write ".txtHRefBillNo.value		= """ & ConvSPChars(E17_s_bill_hdr(S508_E17_bill_no))  & """" & vbCr			
	Response.Write ".txtHRefFlag.value			= ""B""		          "    & vbCr			

	If UCase(Trim(ConvSPChars(E17_s_bill_hdr(S508_E17_cur)))) = UCase(gCurrency) Then
		Response.Write "Parent.ggoOper.SetReqAttr .txtXchgRate, ""Q"" "    & vbCr
		Response.Write ".txtXchgRate.Text = ""1""  "    & vbCr
	Else
		Response.Write "Parent.ggoOper.SetReqAttr .txtXchgRate, ""N"" "    & vbCr
	End If

		'약정회전일 
	Response.Write ".txtCreditRotDay.value = """ & E5_b_biz_partner(S508_E5_credit_rot_day)  & """" & vbCr
		
	Response.Write "parent.BillQueryOk		  "    & vbCr										'☜: 조회가 성공 
		
	Response.Write "End With                  "    & vbCr
    Response.Write "</Script>                 "    & vbCr 

End  Sub

Sub SubBizBLQuery()
	On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear 
    Dim iS7G139
    Dim iCommandSent
    
    Dim I1_s_bill_hdr 
    Dim E1_a_gl 
    Dim E2_a_temp_gl 
    Dim E3_a_batch 
    Dim E4_b_biz_partner 
    Dim E5_s_bill_type_config 
    Dim E6_b_biz_area
    Dim E7_b_biz_partner 
    Dim E8_b_biz_partner 
    Dim E9_b_minor 
    Dim E10_b_minor 
    Dim E11_s_bill_hdr 
    Dim E12_b_biz_partner 
    Dim E13_b_sales_org 
    Dim E14_b_sales_grp 
    Dim E15_b_sales_org 
    Dim E16_b_sales_grp 
    Dim E17_b_biz_partner 
    Dim E18_s_bl_info 
    Dim E19_b_biz_partner 
    Dim E20_b_biz_partner 
    Dim E21_b_biz_partner 
    Dim E22_b_minor 
    Dim E23_b_minor 
    Dim E24_b_minor 
    Dim E25_b_minor 
    Dim E26_b_minor 
    Dim E27_b_minor 
    Dim E28_b_minor 
    Dim E29_b_minor 
    Dim E30_b_country 
    Dim E31_b_country 
    Dim E32_b_country   
      
    Const S528_E1_gl_no = 0    'View Name : exp a_gl
    Const S528_E2_temp_gl_no = 0    'View Name : exp a_temp_gl
    Const S528_E3_batch_no = 0    'View Name : exp a_batch
    Const S528_E4_bp_cd = 0    'View Name : exp_sold_to_party b_biz_partner
    Const S528_E4_bp_nm = 1
    Const S528_E4_credit_rot_day = 2
    Const S528_E5_bill_type = 0    'View Name : exp s_bill_type_config
    Const S528_E5_bill_type_nm = 1
    Const S528_E6_biz_area_nm = 0    'View Name : exp b_biz_area
    Const S528_E7_bp_nm = 0    'View Name : exp_apllicant_nm b_biz_partner
    Const S528_E8_bp_nm = 0    'View Name : exp_beneficiary_nm b_biz_partner
    Const S528_E9_minor_nm = 0    'View Name : exp_vat_type_nm b_minor
    Const S528_E10_minor_nm = 0    'View Name : exp_pay_meth_nm b_minor
    Const S528_E11_bill_no = 0    'View Name : exp s_bill_hdr
    Const S528_E11_post_flag = 1
    Const S528_E11_trans_type = 2
    Const S528_E11_bill_dt = 3
    Const S528_E11_cur = 4
    Const S528_E11_xchg_rate = 5
    Const S528_E11_xchg_rate_op = 6
    Const S528_E11_bill_amt = 7
    Const S528_E11_vat_type = 8
    Const S528_E11_vat_rate = 9
    Const S528_E11_vat_amt = 10
    Const S528_E11_pay_meth = 11
    Const S528_E11_pay_dur = 12
    Const S528_E11_tax_bill_no = 13
    Const S528_E11_tax_prt_cnt = 14
    Const S528_E11_accept_fob_amt = 15
    Const S528_E11_beneficiary = 16
    Const S528_E11_applicant = 17
    Const S528_E11_remark = 18
    Const S528_E11_bill_amt_loc = 19
    Const S528_E11_vat_calc_type = 20
    Const S528_E11_vat_amt_loc = 21
    Const S528_E11_tax_biz_area = 22
    Const S528_E11_pay_type = 23
    Const S528_E11_pay_terms_txt = 24
    Const S528_E11_collect_amt = 25
    Const S528_E11_collect_amt_loc = 26
    Const S528_E11_income_plan_dt = 27
    Const S528_E11_nego_amt = 28
    Const S528_E11_so_no = 29
    Const S528_E11_lc_no = 30
    Const S528_E11_lc_doc_no = 31
    Const S528_E11_lc_amend_seq = 32
    Const S528_E11_bl_flag = 33
    Const S528_E11_biz_area = 34
    Const S528_E11_cost_cd = 35
    Const S528_E11_to_biz_area = 36
    Const S528_E11_to_cost_cd = 37
    Const S528_E11_ref_flag = 38
    Const S528_E11_sts = 39
    Const S528_E11_ext1_qty = 40
    Const S528_E11_ext2_qty = 41
    Const S528_E11_ext3_qty = 42
    Const S528_E11_ext1_amt = 43
    Const S528_E11_ext2_amt = 44
    Const S528_E11_ext3_amt = 45
    Const S528_E11_ext1_cd = 46
    Const S528_E11_ext2_cd = 47
    Const S528_E11_ext3_cd = 48
    Const S528_E11_vat_auto_flag = 49
    Const S528_E11_vat_inc_flag = 50

    Const S528_E12_bp_cd = 0    'View Name : exp_payer b_biz_partner
    Const S528_E12_bp_nm = 1

    Const S528_E13_sales_org = 0    'View Name : exp_income b_sales_org
    Const S528_E13_sales_org_nm = 1

    Const S528_E14_sales_grp = 0    'View Name : exp_income b_sales_grp
    Const S528_E14_sales_grp_nm = 1

    Const S528_E15_sales_org = 0    'View Name : exp_billing b_sales_org
    Const S528_E15_sales_org_nm = 1

    Const S528_E16_sales_grp = 0    'View Name : exp_billing b_sales_grp
    Const S528_E16_sales_grp_nm = 1

    Const S528_E17_bp_cd = 0    'View Name : exp_bill_to_party b_biz_partner
    Const S528_E17_bp_nm = 1

    Const S528_E18_bl_doc_no = 0    'View Name : exp s_bl_info
    Const S528_E18_ship_no = 1
    Const S528_E18_manufacturer = 2
    Const S528_E18_agent = 3
    Const S528_E18_receipt_plce = 4
    Const S528_E18_vessel_nm = 5
    Const S528_E18_voyage_no = 6
    Const S528_E18_forwarder = 7
    Const S528_E18_vessel_cntry = 8
    Const S528_E18_loading_port = 9
    Const S528_E18_dischge_port = 10
    Const S528_E18_delivery_plce = 11
    Const S528_E18_loading_plan_dt = 12
    Const S528_E18_latest_ship_dt = 13
    Const S528_E18_dischge_plan_dt = 14
    Const S528_E18_transport = 15
    Const S528_E18_tranship_cntry = 16
    Const S528_E18_tranship_dt = 17
    Const S528_E18_final_dest = 18
    Const S528_E18_incoterms = 19
    Const S528_E18_packing_type = 20
    Const S528_E18_tot_packing_cnt = 21
    Const S528_E18_container_cnt = 22
    Const S528_E18_packing_txt = 23
    Const S528_E18_gross_weight = 24
    Const S528_E18_weight_unit = 25
    Const S528_E18_gross_volumn = 26
    Const S528_E18_volumn_unit = 27
    Const S528_E18_freight = 28
    Const S528_E18_freight_plce = 29
    Const S528_E18_trans_price = 30
    Const S528_E18_trans_currency = 31
    Const S528_E18_trans_doc_amt = 32
    Const S528_E18_trans_xch_rate = 33
    Const S528_E18_trans_loc_amt = 34
    Const S528_E18_bl_issue_cnt = 35
    Const S528_E18_bl_issue_plce = 36
    Const S528_E18_bl_issue_dt = 37
    Const S528_E18_origin = 38
    Const S528_E18_origin_cntry = 39
    Const S528_E18_loading_dt = 40
    Const S528_E18_ext1_qty = 41
    Const S528_E18_ext2_qty = 42
    Const S528_E18_ext3_qty = 43
    Const S528_E18_ext1_amt = 44
    Const S528_E18_ext2_amt = 45
    Const S528_E18_ext3_amt = 46
    Const S528_E18_ext1_cd = 47
    Const S528_E18_ext2_cd = 48
    Const S528_E18_ext3_cd = 49

    Const S528_E19_bp_nm = 0    'View Name : exp_manufacturer_nm b_biz_partner

    Const S528_E20_bp_nm = 0    'View Name : exp_agent_nm b_biz_partner

    Const S528_E21_bp_nm = 0    'View Name : exp_forwarder_nm b_biz_partner

    Const S528_E22_minor_nm = 0    'View Name : exp_loading_port_nm b_minor

    Const S528_E23_minor_nm = 0    'View Name : exp_discharge_port_nm b_minor

    Const S528_E24_minor_nm = 0    'View Name : exp_transport_nm b_minor

    Const S528_E25_minor_nm = 0    'View Name : exp_incoterms_nm b_minor

    Const S528_E26_minor_nm = 0    'View Name : exp_origin_nm b_minor

    Const S528_E27_minor_nm = 0    'View Name : exp_packing_type_nm b_minor

    Const S528_E28_minor_nm = 0    'View Name : exp_freight_nm b_minor

    Const S528_E29_minor_nm = 0    'View Name : exp_pay_type_nm b_minor

    Const S528_E30_country_nm = 0    'View Name : exp_vessel_cntry_nm b_country

    Const S528_E31_country_nm = 0    'View Name : exp_origin_cntry_nm b_country

    Const S528_E32_country_nm = 0    'View Name : exp_tranship_cntry_nm b_country

    
    I1_s_bill_hdr = Trim(Request("txtBillNo"))
    Set iS7G139 = Server.CreateObject("PS7G139.cSLkInfoSvr")    

    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If    
           		
    Call iS7G139.S_BL_INFO_SVR  ( gStrGlobalCollection ,  "QUERY" ,   I1_s_bill_hdr ,  E1_a_gl , _
             E2_a_temp_gl ,  E3_a_batch ,  E4_b_biz_partner , E5_s_bill_type_config ,  E6_b_biz_area , _
             E7_b_biz_partner , E8_b_biz_partner ,  E9_b_minor ,  E10_b_minor , E11_s_bill_hdr , _
             E12_b_biz_partner ,  E13_b_sales_org , E14_b_sales_grp ,  E15_b_sales_org ,  E16_b_sales_grp , _
             E17_b_biz_partner ,  E18_s_bl_info ,  E19_b_biz_partner , E20_b_biz_partner ,  E21_b_biz_partner , _
             E22_b_minor ,  E23_b_minor ,  E24_b_minor ,  E25_b_minor , E26_b_minor ,  E27_b_minor ,  E28_b_minor , _
             E29_b_minor ,  E30_b_country ,  E31_b_country , E32_b_country )
            				 									 
    If CheckSYSTEMError(Err,True) = True Then
       Set iS7G139 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If     				
    Set iS7G139 = Nothing  
    

	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent.frm1"           & vbCr

		'##### Rounding Logic #####
		'항상 거래화폐가 우선 
	Response.Write ".txtDocCur1.value			= """ & ConvSPChars(E11_s_bill_hdr(S528_E11_cur))      & """" & vbCr
	Response.Write ".txtDocCur2.value			= """ & ConvSPChars(E11_s_bill_hdr(S528_E11_cur))      & """" & vbCr
	Response.Write "parent.CurFormatNumericOCX   " & vbCr
		'##########################

	Response.Write ".txtPayerCd.value			= """ & ConvSPChars(E12_b_biz_partner(S528_E12_bp_cd))        & """" & vbCr				
	Response.Write ".txtPayerNm.value			= """ & ConvSPChars(E12_b_biz_partner(S528_E12_bp_nm))        & """" & vbCr				
	Response.Write ".txtToBizAreaCd.value		= """ & ConvSPChars(E14_b_sales_grp(S528_E14_sales_grp))      & """" & vbCr				
	Response.Write ".txtToBizAreaNm.value		= """ & ConvSPChars(E14_b_sales_grp(S528_E14_sales_grp_nm))   & """" & vbCr				
	Response.Write ".txtTaxBizAreaCd.value		= """ & ConvSPChars(E11_s_bill_hdr(S528_E11_tax_biz_area))    & """" & vbCr					
	Response.Write ".txtTaxBizAreaNm.value		= """ & ConvSPChars(E6_b_biz_area(S528_E6_biz_area_nm))       & """" & vbCr		

	Response.Write ".txtIncomeAmt.Text			= 0 " & vbCr				
	Response.Write ".txtBillAmt.Text			= 0 " & vbCr				
	Response.Write ".txtBillAmtLoc.Text			= 0 " & vbCr				
	Response.Write ".txtVatLocAmt.text			= 0 " & vbCr				
	Response.Write ".txtIncomeLocAmt.text		= 0 " & vbCr				

	Response.Write ".txtPayDur.Text				= """ & E11_s_bill_hdr(S528_E11_pay_dur)                     & """" & vbCr
	Response.Write ".txtSoldtoPartyCd.value		= """ & ConvSPChars(E4_b_biz_partner(S528_E4_bp_cd))         & """" & vbCr	
	Response.Write ".txtSoldtoPartyNm.value		= """ & ConvSPChars(E4_b_biz_partner(S528_E4_bp_nm))         & """" & vbCr	
	Response.Write ".txtBillToPartyCd.value		= """ & ConvSPChars(E17_b_biz_partner(S528_E17_bp_cd))       & """" & vbCr		
	Response.Write ".txtBillToPartyNm.value		= """ & ConvSPChars(E17_b_biz_partner(S528_E17_bp_nm))       & """" & vbCr				
	Response.Write ".txtVatType.value			= """ & ConvSPChars(E11_s_bill_hdr(S528_E11_vat_type))       & """" & vbCr
	Response.Write ".txtVatTypeNm.value			= """ & ConvSPChars(E9_b_minor(S528_E9_minor_nm))            & """" & vbCr
	Response.Write " parent.txtVatType_OnChange                                                                   " & vbCr
'	Response.Write ".txtVatRate.Text			= """ & UNINumClientFormat(E11_s_bill_hdr(S528_E11_vat_rate), ggExchRate.DecPoint, 0)      & """" & vbCr

		 '부가세적용기준 
	Response.Write ".txtVatCalcType.value		= """ & ConvSPChars(E11_s_bill_hdr(S528_E11_vat_calc_type))  & """" & vbCr
		
		'VAT포함여부 
	Response.Write ".txtVatIncFlag.value		= """ & ConvSPChars(E11_s_bill_hdr(S528_E11_vat_inc_flag))   & """" & vbCr
		
	Response.Write ".txtSalesGrpCd.value		= """ & ConvSPChars(E16_b_sales_grp(S528_E16_sales_grp))    & """" & vbCr
	Response.Write ".txtSalesGrpNm.value		= """ & ConvSPChars(E16_b_sales_grp(S528_E16_sales_grp_nm))  & """" & vbCr
	Response.Write ".txtDocCur1.value			= """ & ConvSPChars(E11_s_bill_hdr(S528_E11_cur))            & """" & vbCr
	Response.Write ".txtDocCur2.value			= """ & ConvSPChars(E11_s_bill_hdr(S528_E11_cur))            & """" & vbCr
	Response.Write ".txtPaytermsTxt.value		= """ & ConvSPChars(E11_s_bill_hdr(S528_E11_pay_terms_txt))  & """" & vbCr				
	Response.Write ".txtRemark.value			= """ & ConvSPChars(E11_s_bill_hdr(S528_E11_remark))         & """" & vbCr				

	Response.Write ".txtHExportFlag.value		= """ & ConvSPChars(E11_s_bill_hdr(S528_E11_bl_flag))        & """" & vbCr
	Response.Write ".txtHRefBillNo.value		= """ & ConvSPChars(E11_s_bill_hdr(S528_E11_bill_no))        & """" & vbCr		
	Response.Write ".txtHRefFlag.value			= ""B"" " & vbCr						

	If UCase(Trim(ConvSPChars(E11_s_bill_hdr(S528_E11_cur)))) = UCase(gCurrency) Then
	    Response.Write "Parent.ggoOper.SetReqAttr .txtXchgRate, ""Q"" " & vbCr
	Else
	    Response.Write "Parent.ggoOper.SetReqAttr .txtXchgRate, ""N"" " & vbCr
	End If

		'약정회전일 
	Response.Write ".txtCreditRotDay.value = """ & E4_b_biz_partner(S528_E4_credit_rot_day)      & """" & vbCr		
		
	Response.Write "parent.BillQueryOk	                                                              " & vbCr '☜: 조회가 성공 
    Response.Write "End With"          & vbCr
    Response.Write "</Script>"         & vbCr
    
End  Sub
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
'	strFromList = " FROM dbo.ufn_s_GetBillHdrInfo ('" & FilterVar(Trim(Request("txtConBillNo")),"","SNM") & "', 'Y', 'N', 'Y', 'Y', '" & Request("txtPrevNext") & "', 'N')"
	strFromList = " FROM dbo.ufn_s_GetBillHdrInfo ( " & FilterVar(Request("txtConBillNo"), "''", "S") & " , " & FilterVar("Y", "''", "S") & " , " & FilterVar("%", "''", "S") & ", " & FilterVar("Y", "''", "S") & " , " & FilterVar("Y", "''", "S") & " , '" & Request("txtPrevNext") & "', " & FilterVar("N", "''", "S") & " )"
	lgstrsql = strSelectList & strFromList
	'call svrmsgbox(lgstrsql, vbinformation, i_mkscript)
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub</SCRIPT>
