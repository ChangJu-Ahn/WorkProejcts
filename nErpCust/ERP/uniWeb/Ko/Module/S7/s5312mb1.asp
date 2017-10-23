<%@ LANGUAGE=VBSCript%>
<% Option Explicit %>

<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 매출채권관리 
'*  3. Program ID           : S5312MB1
'*  4. Program Name         : 세금계산서내역등록 
'*  5. Program Desc         :
'*  6. Comproxy List        : S53119LookupTaxBillHdrSvr, S53128ListTaxBillDtlSvr, S53121MaintTaxBillDtlSvr, S53115PostTaxBillSvr
'*  7. Modified date(First) : 2001/06/26
'*  8. Modified date(Last)  : 2001/12/19
'*  9. Modifier (First)     : Cho song hyon
'* 10. Modifier (Last)      : Cho song hyon
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2001/06/26 : 6차 화면 layout & ASP Coding
'*                            -2001/11/09 : 부가세별로 계산하는 로직 추가 
'*                            -2001/12/19 : Date 표준적용 
'**********************************************************************************************
%>

<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../ComASP/LoadInfTb19029.asp" -->
<%

Dim lgOpModeCRUD

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*","NOCOOKIE","MB") 
Call LoadBNumericFormatB("I","*","NOCOOKIE","MB")

On Error Resume Next                                                             '☜: Protect system from crashing
Err.Clear                                                                        '☜: Clear Error status

Call HideStatusWnd

lgOpModeCRUD	=	Request("txtMode")

Select Case lgOpModeCRUD
	Case CStr(UID_M0001)
		'Call SubBizQuery()
		Call SubBizQueryMulti()
	Case CStr(UID_M0002)
		'Call SubBizSave()
		Call SubBizSaveMulti()
	 Case CStr(UID_M0003)                                                         '☜: Delete
        'Call SubBizDelete()
     Case CStr("PostFlag")																'☜: 발행 요청 
		Call SubPostFlag()
End Select

'============================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================
Sub SubBizQuery()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub    
'============================================
' Name : SubBizSave
' Desc : Save Data 
'============================================
Sub SubBizSave()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub
'============================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================
Sub SubBizDelete()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub

'============================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================
Sub SubBizQueryMulti()


	Dim iS7G319												'☆ : 매출헤더 조회용 ComProxy Dll 사용 변수 
	Dim I1_s_tax_bill_no 
	Dim E1_s_tax_bill_hdr 
	Dim E2_s_tax_doc_no 
	Dim E3_b_sales_grp 
	Dim E4_b_biz_area 
	Dim E5_b_biz_partner 
	Dim E6_b_minor 	
	Dim lgCurrency													
	
	'E1_s_tax_bill_hdr 상수 정의 
	Const E1_tax_bill_no = 0    
    Const E1_tax_bill_type = 1
    Const E1_issued_dt = 2
    Const E1_vat_calc_type = 3
    Const E1_vat_io_flag = 4
    Const E1_vat_type = 5
    Const E1_vat_rate = 6
    Const E1_cur = 7
    Const E1_xch_rate_op = 8
    Const E1_xch_rate = 9
    Const E1_net_amt = 10
    Const E1_net_loc_amt = 11
    Const E1_vat_amt = 12
    Const E1_vat_loc_amt = 13
    Const E1_cost_cd = 14
    Const E1_biz_area_cd = 15
    Const E1_report_biz_area = 16
    Const E1_bill_no = 17
    Const E1_post_flag = 18
    Const E1_remarks = 19
    Const E1_ext1_qty = 20
    Const E1_ext2_qty = 21
    Const E1_ext3_qty = 22
    Const E1_ext1_amt = 23
    Const E1_ext2_amt = 24
    Const E1_ext3_amt = 25
    Const E1_ext1_cd = 26
    Const E1_ext2_cd = 27
    Const E1_ext3_cd = 28
    Const E1_vat_inc_flag = 29
	
	'E2_s_tax_doc_no 상수 정의 
    Const E2_tax_doc_no = 0    
	'E3_b_sales_grp 상수 정의 
    Const E3_sales_grp = 0    
    Const E3_sales_grp_nm = 1
	'E4_b_biz_area 상수정의 
    Const E4_biz_area_nm = 0    
	'E5_b_biz_partner 상수정의 
    Const E5_bp_cd = 0    
    Const E5_bp_nm = 1
	'E6_b_minor 상수정의 
    Const E6_minor_nm = 0    

	On Error Resume Next														
	Err.Clear                                                                '☜: Protect system from crashing

    '-----------------------
    ' 세금계산서헤더를 읽어온다.
    '-----------------------
	If Request("txtHQuery") = "T" Then
		
		'-----------------------
		'Data manipulate  area(import view match)
		'-----------------------
		I1_s_tax_bill_no = Trim(Request("txtTaxBillNo"))
		
		Set iS7G319 = Server.CreateObject("PS7G319.cSLkTaxBillHdrSvr")
		
		If CheckSYSTEMError(Err, True) = True Then
		Set iS7G319 = Nothing		                                                 '☜: Unload Comproxy DLL
		Exit Sub
		End If
	
		Call iS7G319.S_LOOKUP_TAX_BILL_HDR_SVR (gStrGlobalCollection, I1_s_tax_bill_no, _
											E1_s_tax_bill_hdr, E2_s_tax_doc_no, _
											E3_b_sales_grp, E4_b_biz_area, _
											E5_b_biz_partner, E6_b_minor)
											
		If CheckSYSTEMError(Err, True) = True Then
			Set iS7G319 = Nothing		                                                 '☜: Unload Comproxy DLL
            Response.Write "<Script Language=vbscript>"			& vbCr
            Response.Write "Parent.frm1.txtTaxBillNo.focus"		& vbCr    
            Response.Write "</Script>"							& vbCr
			Exit Sub
		End If
	
		Set iS7G319 = Nothing
		
		'-----------------------
		'Display result data
		'----------------------- 
		Response.Write "<Script Language=vbscript>" & vbCr
		Response.Write "With parent.frm1"           & vbCr
		Response.Write "parent.SetDefaultVal"		& vbcr
		Response.Write "Call parent.SetToolBar(""" & 11000000000011 & """" & ")" & vbcr
		'##### Rounding Logic #####
		'항상 거래화폐가 우선 
		lgCurrency = ConvSPChars(E1_s_tax_bill_hdr(E1_cur))
		Response.Write ".txtCurrency.Value			= """ & ConvSPChars(E1_s_tax_bill_hdr(E1_cur))						& """" & vbCr
		Response.Write "parent.CurFormatNumericOCX" & vbCr
		'##########################
        	
		'-----------------------
		' 매출헤더의 내용을 표시한다.
		'----------------------- 
		'발행처 
		Response.Write ".txtBillToParty.Value		= """ & ConvSPChars(E5_b_biz_partner(E5_bp_cd))					& """" & vbCr
		Response.Write ".txtBillToPartyNm.Value		= """ & ConvSPChars(E5_b_biz_partner(E5_bp_nm))					& """" & vbCr
		'매출채권번호 
		Response.Write ".txtBillNo.Value			= """ & ConvSPChars(E1_s_tax_bill_hdr(E1_bill_no))				& """" & vbCr
		'VAT유형 
		Response.Write ".txtVATType.Value			= """ & ConvSPChars(E1_s_tax_bill_hdr(E1_vat_type))				& """" & vbCr
		Response.Write ".txtVATTypeNm.Value			= """ & ConvSPChars(E6_b_minor(E6_minor_nm))					& """" & vbCr
		'VAT율 
		Response.Write ".txtVATRate.text			= """ & UNIConvNumDBToCompanyByCurrency(E1_s_tax_bill_hdr(E1_vat_rate),gCurrency,ggExchRateNo, "X" , "X")						& """" & vbCr
		'VAT적용기준 
		If E1_s_tax_bill_hdr(E1_vat_calc_type) = "1" Then 
		   Response.Write ".rdoVATCalcType1.checked = True "			&   vbCr
		elseif	E1_s_tax_bill_hdr(E1_vat_calc_type)   = "2" Then    
		   Response.Write ".rdoVATCalcType2.checked = True "			&   vbCr
		End If   
		'부가세포함여부 
		If E1_s_tax_bill_hdr(E1_vat_inc_flag) = "1" Then 
		   Response.Write ".rdoVATIncflag1.checked = True "				&   vbCr
		elseif	E1_s_tax_bill_hdr(E1_vat_calc_type)   = "2" Then    
		   Response.Write ".rdoVATIncflag2.checked = True "				&   vbCr
		End If   
		
		'공급가금액 
		'Response.Write ".txtSupplyAmt.text		= """ & UNINumClientFormat(E1_s_tax_bill_hdr(E1_net_amt), ggAmtOfMoney.DecPoint, 0)												& """" & vbCr
		Response.Write ".txtSupplyAmt.text		= """ & UNIConvNumDBToCompanyByCurrency(E1_s_tax_bill_hdr(E1_net_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")												& """" & vbCr
		'공급가자국금액 
		Response.Write ".txtSupplyLocAmt.text	= """ & UNIConvNumDBToCompanyByCurrency(E1_s_tax_bill_hdr(E1_net_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo  , "X")		& """" & vbCr
		Response.Write ".txtLocCur.Value		= """ & gCurrency								& """" & vbCr			
		'VAT금액 
		'Response.Write ".txtVATAmt.text			= """ & UNINumClientFormat(E1_s_tax_bill_hdr(E1_vat_amt), ggAmtOfMoney.DecPoint, 0)												& """" & vbCr
		Response.Write ".txtVATAmt.text			= """ & UNIConvNumDBToCompanyByCurrency(E1_s_tax_bill_hdr(E1_vat_amt),lgCurrency,ggAmtOfMoneyNo, gTaxRndPolicyNo  , "X")												& """" & vbCr
		'VAT자국액 
		Response.Write ".txtLocVatAmt.text		= """ & UNIConvNumDBToCompanyByCurrency(E1_s_tax_bill_hdr(E1_vat_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo  , "X")		& """" & vbCr
		'영업그룹 
		Response.Write ".HSalesGrpCd.Value		= """ & ConvSPChars(E3_b_sales_grp(E3_sales_grp))							& """" & vbCr
		Response.Write ".HSalesGrpNm.Value		= """ & ConvSPChars(E3_b_sales_grp(E3_sales_grp_nm))						& """" & vbCr
    	Response.Write ".HTaxBillNo.value		= """ & ConvSPChars(Trim(Request("txtTaxBillNo")))										& """" & vbcr
		'발행일 
		Response.Write ".txtIssueDt.value		= """ & UNIDateClientFormat(E1_s_tax_bill_hdr(E1_issued_dt))				& """" & vbCr
		
		If E1_s_tax_bill_hdr(E1_post_flag) = "Y" Then 
		   Response.Write ".HPostFlag.value		= ""Y"""				& vbCr
		else
		   Response.Write ".HPostFlag.value		= ""N"""				& vbCr
		End If   
			
		'-----------------------
		' Rounding Column Set
		'----------------------- 
		Response.Write "parent.CurFormatNumSprSheet "					& vbcr
		Response.Write "parent.DbHdrQueryOk	"							& vbcr													'☜: 조회가 성공	
		Response.Write "End With"										& vbCr
		Response.Write "</Script>"										& vbCr

	End If 		' End of Header Query
	
	'-----------------------
    ' 세금계산서내역을 읽어온다.
    '-----------------------
	Dim iS7G328												'☆ : 매출내역등록 조회용 ComProxy Dll 사용 변수 
	Dim I1_s_tax_bill_hdr
	Dim I2_s_tax_bill_dtl
	Dim E1_s_tax_bill_dtl
	Dim	EG1_exp_grp
	
	Dim iStrSvrData
	Dim iDblTotAmt, iDblTotAmtLoc		
	
	Dim iStrNextKey											' 다음 값 
	Dim lgStrPrevKey										' 이전 값 
	Dim ILngMaxRow											' 현재 그리드의 최대Row
	Dim ILngRow
	
	Const C_SHEETMAXROWS_D  = 100
	
	Const EG1_minor_nm = 0           ' exp_item b_minor
    Const EG1_item_cd = 1            ' exp_item b_item
    Const EG1_item_nm = 2
    Const EG1_tax_bill_seq = 3       ' exp_item s_tax_bill_dtl
    Const EG1_bill_dt = 4
    Const EG1_bill_qty = 5
    Const EG1_bill_unit = 6
    Const EG1_bill_price = 7
    Const EG1_bill_amt = 8
    Const EG1_bill_amt_loc = 9
    Const EG1_vat_type = 10
    Const EG1_vat_rate = 11
    Const EG1_vat_amt = 12
    Const EG1_vat_amt_loc = 13
    Const EG1_ext1_qty = 14
    Const EG1_ext2_qty = 15
    Const EG1_ext3_qty = 16
    Const EG1_ext1_amt = 17
    Const EG1_ext2_amt = 18
    Const EG1_ext3_amt = 19
    Const EG1_ext1_cd = 20
    Const EG1_ext2_cd = 21
    Const EG1_ext3_cd = 22
    Const EG1_vat_inc_flag = 23
    Const EG1_bill_seq = 24          ' exp_item s_bill_dtl
    Const EG1_bill_no = 25           ' exp_item s_bill_hdr
    Const EG1_xchg_rate = 26
    Const EG1_xchg_rate_op = 27
    Const EG1_SPEC = 28
    
    On Error Resume Next														
	Err.Clear                                                                '☜: Protect system from crashing
    
    '-----------------------
    'Data manipulate  area(import view match)
    '-----------------------
    I1_s_tax_bill_hdr = Trim(Request("txtTaxBillNo"))
    I2_s_tax_bill_dtl = UNICDbl(Request("lgStrPrevKey"), 0)

    Set iS7G328 = Server.CreateObject("PS7G328.cSListTaxBillDtlSvr")    

    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If   

    Call iS7G328.S_LIST_TAX_BILL_DTL_SVR(gStrGlobalCollection, C_SHEETMAXROWS_D, _
											I1_s_tax_bill_hdr, I2_s_tax_bill_dtl,_
											EG1_exp_grp, E1_s_tax_bill_dtl) 	
    
    If CheckSYSTEMError(Err,True) = True Then
		Set iS7G328 = Nothing
		Response.Write "<Script Language=vbscript>"			& vbCr
        Response.Write "Parent.frm1.txtTaxBillNo.focus"		& vbCr    
        Response.Write "</Script>"							& vbCr
		Exit Sub
    End If   
            
	Set iS7G328 = Nothing	
    
    Dim iArrCols, iArrRows
    Dim iLngSheetMaxRows
	
	' Set Next key
	If Ubound(EG1_exp_grp,1) = C_SHEETMAXROWS_D Then
		'세금계산서순번 
		iStrNextKey = EG1_exp_grp(C_SHEETMAXROWS_D, EG1_tax_bill_seq)
		iLngSheetMaxRows  = C_SHEETMAXROWS_D - 1
	Else
		iStrNextKey = ""
		iLngSheetMaxRows = Ubound(EG1_exp_grp,1)
	End If

	ReDim iArrCols(22)						' Column 수 
	Redim iArrRows(iLngSheetMaxRows)		' 조회된 Row 수만큼 배열 재정의 

	ILngMaxRow  = CLng(Request("txtMaxRows"))										'☜: Fetechd Count     
	
	'-----------------------
	'Result data display area
	'----------------------- 
	iArrCols(0) = ""
	For ILngRow = 0 To UBound(EG1_exp_grp, 1)
	   	iArrCols(1)  = ConvSPChars(EG1_exp_grp(ILngRow, EG1_item_cd))			' 품목코드 
	   	iArrCols(2)  = ConvSPChars(EG1_exp_grp(ILngRow, EG1_item_nm))			' 품목명 
	   	iArrCols(3)  = UNINumClientFormat(EG1_exp_grp(ILngRow, EG1_bill_qty), ggQty.DecPoint, 0)		' 수량 
	   	iArrCols(4)  = ConvSPChars(EG1_exp_grp(ILngRow, EG1_bill_unit))			' 단위 
	   	iArrCols(5)  = UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(ILngRow, EG1_bill_price), lgCurrency, ggUnitCostNo, "X", "X")		' 단가 
	   	iArrCols(6)  = UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(ILngRow, EG1_bill_amt), lgCurrency, ggAmtOfMoneyNo, "X", "X")		' 공급가액 
	   	iArrCols(7)  = ConvSPChars(EG1_exp_grp(ILngRow, EG1_vat_type))			' VAT 유형 
	   	iArrCols(8)  = ConvSPChars(EG1_exp_grp(ILngRow, EG1_minor_nm))			' VAT 유형명 
	   	iArrCols(9)  = UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(ILngRow, EG1_vat_rate), gCurrency, ggExchRateNo, "X", "X")				' VAT 율 
	   	iArrCols(10) = UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(ILngRow, EG1_vat_amt),lgCurrency,ggAmtOfMoneyNo, gTaxRndPolicyNo , "X")	' VAT 금액 
		iArrCols(11) = UNIConvNumDBToCompanyByCurrency(cdbl(EG1_exp_grp(ILngRow, EG1_bill_amt)) + cdbl(EG1_exp_grp(ILngRow, EG1_vat_amt)),lgCurrency, ggAmtOfMoneyNo, "X" , "X")	'합계금액 
	   	iArrCols(12) = UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(ILngRow, EG1_bill_amt_loc),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo  , "X")										'공급가자국금액 
	   	iArrCols(13) = UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(ILngRow, EG1_vat_amt_loc),gCurrency,ggAmtOfMoneyNo, gTaxRndPolicyNo  , "X")										'VAT자국액 
		iArrCols(14) = UNIConvNumDBToCompanyByCurrency(cdbl(EG1_exp_grp(ILngRow, EG1_bill_amt_loc)) + cdbl(EG1_exp_grp(ILngRow, EG1_vat_amt_loc)), gCurrency, ggAmtOfMoneyNo, gLocRndPolicyNo, "X")		'합계자국금액 
	   	iArrCols(15) = UNINumClientFormat(EG1_exp_grp(ILngRow, EG1_tax_bill_seq), 0, 0)		'순번 
	   	iArrCols(16) = ConvSPChars(EG1_exp_grp(ILngRow, EG1_bill_no))						'매출채권번호 
	   	iArrCols(17) = UNINumClientFormat(EG1_exp_grp(ILngRow, EG1_bill_seq), 0, 0)			'매출채권순번 
	   	iArrCols(18) = ConvSPChars(EG1_exp_grp(ILngRow, EG1_SPEC))							'품목규격 
	   	iArrCols(19) = ConvSPChars(EG1_exp_grp(ILngRow, EG1_xchg_rate_op))					'환율연산자 
	   	iArrCols(20) = UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(ILngRow, EG1_xchg_rate), gCurrency, ggExchRateNo, "X" , "X")			'환율 
	   	iArrCols(21) = ConvSPChars(EG1_exp_grp(ILngRow, EG1_vat_inc_flag))					'부가세포함여부 
	   	iArrCols(22) = ILngMaxRow + ILngRow

   		iArrRows(ILngRow) = Join(iArrCols, gColSep)
	Next
	
	Response.Write "<Script language=vbs> "											& vbCr
	Response.Write "With parent"													& vbCr   
    Response.Write " .frm1.HTaxBillNo.value = """ & ConvSPChars(Request("txtTaxBillNo")) & """" & vbCr    
    Response.Write " .ggoSpread.Source		= .frm1.vspdData" & vbCr    
    Response.Write " .frm1.vspdData.Redraw = False  " & vbCr      
    Response.Write " .ggoSpread.SSShowDataByClip   """ & Join(iArrRows, gColSep & gRowSep) & gColSep & gRowSep & """ ,""F""" & vbCr
	Response.Write ".frm1.vspdData.Redraw = True  " & vbCr
    Response.Write " .lgStrPrevKey			= """ & iStrNextKey	& """" & vbCr    
    Response.Write " .DbQueryOk " & vbCr   
    Response.Write "End With " & vbCr   
    Response.Write "</Script> "		
    
End Sub

'============================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================
Sub SubBizSaveMulti() 

	Dim iS7G321												'☆ : 매출내역등록 조회용 ComProxy Dll 사용 변수 
	Dim I1_s_tax_bill_hdr
	Dim iErrorPosition
	Dim itxtSpread
	
	On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear	
    									
    Set iS7G321 = Server.CreateObject("PS7G321.cSTaxBillDtlSvr")  
	
	If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If
    
    I1_s_tax_bill_hdr = Trim(Request("HTaxBillNo"))
    itxtSpread = Trim(Request("txtSpread"))
    
    Call iS7G321.S_MAINT_TAX_BILL_DTL_SVR(gStrGlobalCollection, I1_s_tax_bill_hdr,itxtSpread, iErrorPosition)
    
    If CheckSYSTEMError2(Err, True, iErrorPosition & "행","","","","") = True Then
       Set iS7G321 = Nothing
       Exit Sub
	End If
	
	Set iS7G321 = Nothing
	
    Response.Write "<Script Language=vbscript> "	& vbCr         
    Response.Write " Parent.DBSaveOk "				& vbCr   
    Response.Write "</Script> "           
	Response.End																				'☜: Process End

End Sub
'============================================
' Name : SubBizPostFlag
' Desc : Save Data 
'============================================
Sub SubPostFlag()
    Dim iS7G315
    Dim itxtFlgMode

	Dim I1_s_tax_bill_hdr_tax_bill_no
	Dim I2_s_wks_user_user_id

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    I1_s_tax_bill_hdr_tax_bill_no = Trim(Request("HTaxBillNo"))
    
    Set iS7G315 = Server.CreateObject("PS7G315.cSPostTaxBillSvr")
    
	If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If  
               
	call iS7G315.S_POST_TAX_BILL_SVR(gStrGlobalCollection, I1_s_tax_bill_hdr_tax_bill_no, I2_s_wks_user_user_id)
    
    If CheckSYSTEMError(Err,True) = True Then
		Set iS7G315 = Nothing
		Exit Sub
	End If     
    '-----------------------
	'Result data display area
	'----------------------- 
	Set iS7G315 = Nothing
	
	Response.Write "<Script language=vbs> " & vbCr         
    Response.Write " Parent.DbSaveOk "		& vbCr   
    Response.Write "</Script> "  
    
End Sub

%>
