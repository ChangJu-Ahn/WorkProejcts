<%@ LANGUAGE = VBSCript%>
<% Option Explicit %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 매출채권관리 
'*  3. Program ID           : S5112MB1
'*  4. Program Name         : 매출채권내역등록 
'*  5. Program Desc         :
'*  6. Comproxy List        : PS7G128.cSListBillDtlSvr,PS7G121.cSBillDtlSvr,PS7G115.cSPostOpenArSvr
'*  7. Modified date(First) : 2000/04/19
'*  8. Modified date(Last)  : 2002/11/12
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
'*							  -2002/06/24 : VB conversion
'*							  -2002/11/12 : UI성능 적용	
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

ON ERROR RESUME Next														

Call HideStatusWnd
'for 구주Tax
'@@@@@@@@@@@
Dim pvCB 
'@@@@@@@@@@@

Dim iPS7G121												'☆ : 매출내역등록입력/수정/삭제용 ComProxy Dll 사용 변수 
Dim iPS7G128												'☆ : 매출내역등록 조회용 ComProxy Dll 사용 변수 
Dim iPS7G115												'☆ : 매출내역확정용 ComProxy Dll 사용 변수 
Dim strMode		
Dim iStrNextKey											' 다음 값 
Dim lgStrPrevKey										' 이전 값 
Dim LngMaxRow											' 현재 그리드의 최대Row
Dim LngRow
Dim GroupCount															
Dim lgCurrency
Dim lgArrGlFlag
Dim lgStrGlFlag

strMode = Request("txtMode")												'☜ : 현재 상태를 받음 

Select Case strMode

Case CStr(UID_M0001)														'☜: 현재 조회/Prev/Next 요청을 받음 
	Dim iStrSvrData

    Err.Clear                                                                '☜: Protect system from crashing

    '-----------------------
    ' 매출헤더를 읽어온다.
    '-----------------------
	If Request("txtHQuery") = "T" Then
	
	    Call SubOpenDB(lgObjConn)
	    call SubMakeSQLStatements
	    
		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                 'If data not exists
		    lgObjRs.Close
		    lgObjConn.Close
		    Set lgObjRs = Nothing
		    Set lgObjConn = Nothing
			'매출채권정보가 없습니다.
			Call DisplayMsgBox("205100", vbInformation, "", "", I_MKSCRIPT)             '☜: No data is found. 
			%>
			<Script Language=vbscript>
				parent.SetDefaultVal
				Call parent.SetToolbar("11000000000011")
			</Script>
			<%
		    Response.End
		End If
	%>
	<Script Language=vbscript>
		With parent
			'-----------------------
			' 매출헤더의 내용을 표시한다.
			'----------------------- 

			'##### Rounding Logic #####
			'항상 거래화폐가 우선 
			<%
			lgCurrency = ConvSPChars(lgObjRs("Cur"))
			%>
            
            .frm1.txtCurrency.value			= "<%=lgCurrency%>"
			
			parent.CurFormatNumericOCX
			
			'##########################

			.frm1.txtHBillNo.value			= "<%=ConvSPChars(lgObjRs("bill_no"))%>"
			.frm1.txtHBillDt.value			= "<%=UNIDateClientFormat(lgObjRs("bill_dt"))%>"
			.frm1.txtSoldtoParty.value		= "<%=ConvSPChars(lgObjRs("Sold_to_Party"))%>"
			.frm1.txtSoldtoPartyNm.value	= "<%=ConvSPChars(lgObjRs("Sold_to_Party_Nm"))%>"
			.frm1.txtBillToPartyCd.value	= "<%=ConvSPChars(lgObjRs("Bill_To_Party"))%>"
			.frm1.txtBillToPartyNm.value	= "<%=ConvSPChars(lgObjRs("Bill_To_Party_Nm"))%>"
			.frm1.txtPayTermsCd.value		= "<%=ConvSPChars(lgObjRs("Pay_Meth"))%>"
			.frm1.txtPayTermsNm.value		= "<%=ConvSPChars(lgObjRs("Pay_Meth_Nm"))%>"
			.frm1.txtHSalesGrpCd.value		= "<%=ConvSPChars(lgObjRs("sales_grp"))%>"
			.frm1.txtHSalesGrpNm.value		= "<%=ConvSPChars(lgObjRs("sales_grp_nm"))%>"

			.frm1.txtOriginBillAmt.Text		= "<%=UNIConvNumDBToCompanyByCurrency(lgObjRs("bill_amt"),lgCurrency,ggAmtOfMoneyNo, "X" , "X")%>"
			.frm1.txtVatAmt.Text			= "<%=UNIConvNumDBToCompanyByCurrency(lgObjRs("vat_amt"),lgCurrency,ggAmtOfMoneyNo, gTaxRndPolicyNo, "X")%>"
			.frm1.txtTotBillAmt.Text		= "<%=UNIConvNumDBToCompanyByCurrency(Cdbl(lgObjRs("bill_amt")) + CDbl(lgObjRs("vat_amt")) + CDbl(lgObjRs("deposit_amt")),lgCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")%>"

			.frm1.txtSoNo.value				= "<%=ConvSPChars(lgObjRs("so_no"))%>"
			.frm1.txtXchgRate.Text			= "<%=UNIConvNumDBToCompanyByCurrency(lgObjRs("xchg_rate"), gCurrency, ggExchRateNo, "X" , "X")%>"
			
			<% 'vat율 화면에 숨김 %>
			.frm1.HVatRate.value			= "<%=UNIConvNumDBToCompanyByCurrency(lgObjRs("vat_rate"), gCurrency, ggExchRateNo, "X" , "X")%>"
			<% 'vat타입 화면에 숨김 %>
			.frm1.HVatType.value			= "<%=ConvSPChars(lgObjRs("vat_type"))%>"
			<% '환율연산자를 화면에 숨김 %>
			.frm1.txtXchgOp.value			= "<%=ConvSPChars(lgObjRs("xchg_rate_op"))%>"
			
			.frm1.txtBillTypeCd.value		= "<%=ConvSPChars(lgObjRs("bill_type"))%>"

			<% '수주진행상태 %>
			.frm1.txtSts.value				= "<%=lgObjRs("sts")%>"
			<% '수주, L/C참조 여부 %>
			.frm1.txtRefFlag.value			= "<%=lgObjRs("ref_flag")%>"
			<% '반품여부 %>
			.frm1.txtReverseFlag.value		= "<%=lgObjRs("reverse_flag")%>"

			'VAT적용기준 
			'VAT적용기준이 통합인 경우 VAT 유형, 유형명, 율을 Hidden 처리 
			If "<%=lgObjRs("vat_calc_type")%>" = "2" Then
				.frm1.rdoVatCalcType2.checked = True
			Else
				.frm1.rdoVatCalcType1.checked = True
			End If

			'부가세포함여부 
			If Trim("<%=lgObjRs("vat_inc_flag")%>") = "1" Then
				.frm1.rdoVatIncFlag1.checked = True
			Else
				.frm1.rdoVatIncFlag2.checked = True
			End If
            
			.frm1.HPostFlag.value = "<%=lgObjRs("post_flag")%>"

			<%If lgObjRs("post_flag") = "Y" AND Len(Trim(lgObjRs("gl_no"))) Then
				lgArrGlFlag = Split(Trim(lgObjRs("gl_no")), Chr(11)) 
				lgStrGlFlag = lgArrGlFlag(0)%>
				
				If "<%=lgArrGlFlag(0)%>" = "G" Then	
					'회계전표번호 
					.frm1.txtGLNo.value	= "<%=lgArrGlFlag(1)%>"	
				ElseIf "<%=lgArrGlFlag(0)%>" = "T" Then
					'결의전표번호 
					.frm1.txtTempGLNo.value	= "<%=lgArrGlFlag(1)%>"	
				Else
					'Batch번호 
					.frm1.txtBatchNo.value	= "<%=lgArrGlFlag(1)%>"	
				End If
			<%Else%>
					.frm1.txtGLNo.value	= ""	
					.frm1.txtTempGLNo.value	= ""	
					.frm1.txtBatchNo.value	= ""	
			<% End If %>
			
			If .frm1.HPostFlag.value = "Y" Then
				
				.frm1.btnPostFlag.value = "확정취소"
				If "<%=lgStrGlFlag%>" = "G" Or "<%=lgStrGlFlag%>" = "T" Then
					.frm1.btnGLView.disabled = False
				Else
					.frm1.btnGLView.disabled = True
				End If
			Else
				.frm1.btnPostFlag.value = "확정"
				.frm1.btnGLView.disabled = True
			End If

			<% '선수금 현황 버튼 Enable %>
			IF "<%=lgObjRs("PreRcpt_flag")%>" = "Y" Then
				.frm1.btnPreRcptView.disabled = False
			Else
				.frm1.btnPreRcptView.disabled = True
			End If

			<%
			lgObjRs.Close
			lgObjConn.Close
			Set lgObjRs = Nothing
			Set lgObjConn = Nothing
			%>
			parent.SetSpreadHidden		
			'-----------------------
			' Rounding Column Set
			'----------------------- 
			parent.CurFormatNumSprSheet
			
			.DbQueryOk														'☜: 조회가 성공 

		End With
	</Script>		
<%
	ElseIf Request("txtHQuery") = "F" Then
		lgCurrency = Request("txtCurrency")			
	End If

	'-----------------------
    ' 매출내역을 읽어온다.
    '-----------------------
    '--------------
	'Interface 정의 
	'--------------
    'View Name : imp_next s_bill_dtl
    Const S526_I1_bill_seq = 0
    'View Name : imp s_bill_hdr
    Const S526_I2_bill_no = 0
    'View Name : exp_next s_bill_dtl
    Const S526_E1_bill_seq = 0
    
    'Group Name : exp_grp
    Const S526_EG1_E1_minor_nm = 0    'View Name : exp_item b_minor
    Const S526_EG1_E2_cc_seq = 1    'View Name : exp_item s_cc_dtl
    Const S526_EG1_E3_cc_no = 2    'View Name : exp_item s_cc_hdr
    Const S526_EG1_E4_lc_seq = 3    'View Name : exp_item s_lc_dtl
    Const S526_EG1_E5_lc_no = 4    'View Name : exp_item s_lc_hdr
    Const S526_EG1_E6_bill_seq = 5    'View Name : exp_item s_bill_dtl
    Const S526_EG1_E6_bill_price = 6
    Const S526_EG1_E6_bill_amt = 7
    Const S526_EG1_E6_vat_amt = 8
    Const S526_EG1_E6_bill_qty = 9
    Const S526_EG1_E6_bill_unit = 10
    Const S526_EG1_E6_remark = 11
    Const S526_EG1_E6_item_acct = 12
    Const S526_EG1_E6_tracking_no = 13
    Const S526_EG1_E6_plant_biz_area = 14
    Const S526_EG1_E6_cost_cd = 15
    Const S526_EG1_E6_hs_no = 16
    Const S526_EG1_E6_cust_item_cd = 17
    Const S526_EG1_E6_bill_amt_loc = 18
    Const S526_EG1_E6_vat_type = 19
    Const S526_EG1_E6_vat_rate = 20
    Const S526_EG1_E6_vat_amt_loc = 21
    Const S526_EG1_E6_cust_po_no = 22
    Const S526_EG1_E6_cust_po_seq = 23
    Const S526_EG1_E6_gross_weight = 24
    Const S526_EG1_E6_net_weight = 25
    Const S526_EG1_E6_volume_size = 26
    Const S526_EG1_E6_ext1_qty = 27
    Const S526_EG1_E6_ext2_qty = 28
    Const S526_EG1_E6_ext3_qty = 29
    Const S526_EG1_E6_ext1_amt = 30
    Const S526_EG1_E6_ext2_amt = 31
    Const S526_EG1_E6_ext3_amt = 32
    Const S526_EG1_E6_ext1_cd = 33
    Const S526_EG1_E6_ext2_cd = 34
    Const S526_EG1_E6_ext3_cd = 35
    Const S526_EG1_E6_vat_inc_flag = 36
    Const S526_EG1_E6_deposit_price = 37
    Const S526_EG1_E6_deposit_amt = 38
    Const S526_EG1_E6_ret_item_flag = 39
    Const S526_EG1_E7_dn_seq = 40    'View Name : exp_item s_dn_dtl
    Const S526_EG1_E8_dn_no = 41    'View Name : exp_item s_dn_hdr
    Const S526_EG1_E9_plant_cd = 42    'View Name : exp_item b_plant
    Const S526_EG1_E10_item_cd = 43    'View Name : exp_item b_item
    Const S526_EG1_E10_item_nm = 44
    Const S526_EG1_E10_spec = 45
    Const S526_EG1_E11_so_seq = 46    'View Name : exp_item s_so_dtl
    Const S526_EG1_E12_so_no = 47    'View Name : exp_item s_so_hdr
    Const S526_EG1_E13_OUT_NO = 48
    Const S526_EG1_E13_TRANS_TIME = 49
    Const S526_EG1_E13_OUT_TYPE = 50
    Const S526_EG1_E13_CREATE_TYPE = 51
    Const S526_EG1_E13_DN_NO = 52
    Const S526_EG1_E13_DN_SEQ = 53

    Const S526_EG1_E13_pgm_name = 54
    Const S526_EG1_E13_pgm_price = 55
        
    Const C_SHEETMAXROWS_D  = 100
    
    '--------
	'View선언 
	'--------    
    Dim I2_s_bill_hdr
    Dim I1_s_bill_dtl
    Dim EG1_exp_grp
    Dim E1_s_bill_dtl 
    
     '---------------------------------------
    'Data manipulate  area(import view match)
    '----------------------------------------
    redim I2_s_bill_hdr(0)
    
    I2_s_bill_hdr(S526_I2_bill_no) = Trim(Request("txtConBillNo"))
    
    redim I1_s_bill_dtl(0)
    
    If Trim(Request("lgStrPrevKey")) = "" then
		I1_s_bill_dtl(S526_I1_bill_seq) = 0
    Else
		I1_s_bill_dtl(S526_I1_bill_seq) = cdbl(Request("lgStrPrevKey"))
	End if	
	 
    
    Set iPS7G128 = Server.CreateObject("PS7G128_KO441.cSListBillDtlSvr")    
  
    If CheckSYSTEMError(Err,True) = True Then
		Response.End		
    End If
	
	Call iPS7G128.S_LIST_BILL_DTL_SVR (gStrGlobalCollection , C_SHEETMAXROWS_D , _
	                                 I1_s_bill_dtl ,I2_s_bill_hdr, EG1_exp_grp, E1_s_bill_dtl)

	If CheckSYSTEMError(Err,True) = True Then
		Set iPS7G128 = Nothing		                                                 '☜: Unload Comproxy DLL
		Response.Write "<Script language=vbs> " & vbCr   
		Response.Write " Parent.frm1.txtConBillNo.Focus" & vbCr   
		Response.Write "</Script> " & vbCr          
		Response.End 
    End If   

    Set iPS7G128 = Nothing	
		'----------------------------
		' 매출내역의 내용을 표시한다.
		'----------------------------
	Dim iLngSheetMaxRows
	Dim iArrCols, iArrRows


	' Set Next key
	If Ubound(EG1_exp_grp,1) = C_SHEETMAXROWS_D Then
		'매출번호 
		iStrNextKey = E1_s_bill_dtl(S526_E1_bill_seq)
		iLngSheetMaxRows  = C_SHEETMAXROWS_D - 1
	Else
		iStrNextKey = ""
		iLngSheetMaxRows = Ubound(EG1_exp_grp,1)
	End If

	ReDim iArrCols(44)						' Column 수 
	Redim iArrRows(iLngSheetMaxRows)		' 조회된 Row 수만큼 배열 재정의 

	LngMaxRow = CInt(Request("txtMaxRows")) + 1

	' 변경되지 않는 값 설정 
	iArrCols(0)  = ""		' Row Header
	iArrCols(11)  = "" 		'vat유형 팝업 
	iArrCols(19) = "0"		'FOB금액 
	
	For LngRow = 0 To Ubound(EG1_exp_grp,1)
		iArrCols(1)  = ConvSPChars(EG1_exp_grp(LngRow,S526_EG1_E10_item_cd))		'품목코드 
		iArrCols(2)  = ConvSPChars(EG1_exp_grp(LngRow,S526_EG1_E10_item_nm))		'품목명 
		iArrCols(3)  = ConvSPChars(EG1_exp_grp(LngRow,S526_EG1_E6_tracking_no))		'Tracking No
		iArrCols(4)  = UNINumClientFormat(EG1_exp_grp(LngRow, S526_EG1_E6_bill_qty), ggQty.DecPoint, 0)			'수량 
		iArrCols(5)  = ConvSPChars(EG1_exp_grp(LngRow,S526_EG1_E6_bill_unit))		'단위 
		iArrCols(6) = Trim(EG1_exp_grp(LngRow,S526_EG1_E6_vat_inc_flag))			'부가세포함여부 

		'부가세포함여부명 
		If iArrCols(6) = "1" Then
			iArrCols(7) = "별도"
		Else
			iArrCols(7) = "포함"
		End If
		
		iArrCols(8)  = UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(LngRow,S526_EG1_E6_bill_price),lgCurrency,ggUnitCostNo, "X" , "X")		'단가 
		iArrCols(9)  = UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(LngRow,S526_EG1_E6_bill_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")		'금액 
		iArrCols(10)  = ConvSPChars(EG1_exp_grp(LngRow,S526_EG1_E6_vat_type))		'vat유형 
		iArrCols(12) = ConvSPChars(EG1_exp_grp(LngRow,S526_EG1_E1_minor_nm))		'vat유형명 
		iArrCols(13) = UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(LngRow,S526_EG1_E6_vat_rate), gCurrency, ggExchRateNo, "X" , "X")		'vat율 
		iArrCols(14) = UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(LngRow,S526_EG1_E6_vat_amt),lgCurrency,ggAmtOfMoneyNo, gTaxRndPolicyNo  , "X")			'부가세액 
		iArrCols(15) = UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(LngRow,S526_EG1_E6_bill_amt_loc),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo  , "X")		'원화금액 
		iArrCols(16) = UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(LngRow,S526_EG1_E6_vat_amt_loc),gCurrency,ggAmtOfMoneyNo, gTaxRndPolicyNo  , "X")		'부가세원화금액 
		iArrCols(17) = UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(LngRow,S526_EG1_E6_deposit_price ),lgCurrency,ggUnitCostNo, "X" , "X")					'적립금단가 
		iArrCols(18) = UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(LngRow,S526_EG1_E6_deposit_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")					'적립금액 
		iArrCols(20) = ConvSPChars(EG1_exp_grp(LngRow,S526_EG1_E6_remark))					'비고 
		iArrCols(21) = ConvSPChars(EG1_exp_grp(LngRow,S526_EG1_E8_dn_no))					'출하번호 
		iArrCols(22) = UNINumClientFormat(EG1_exp_grp(LngRow,S526_EG1_E7_dn_seq), 0, 0)		'출하순번 
		iArrCols(23) = ConvSPChars(EG1_exp_grp(LngRow,S526_EG1_E12_so_no))					'수주번호 
		iArrCols(24) = UNINumClientFormat(EG1_exp_grp(LngRow,S526_EG1_E11_so_seq), 0, 0)	'수주순번 
		iArrCols(25) = ConvSPChars(EG1_exp_grp(LngRow,S526_EG1_E5_lc_no))					'L/C번호 
		iArrCols(26) = UNINumClientFormat(EG1_exp_grp(LngRow,S526_EG1_E4_lc_seq), 0, 0)		'L/C순번 
		iArrCols(27) = UNINumClientFormat(EG1_exp_grp(LngRow,S526_EG1_E6_bill_seq), 0, 0)	'매출순번 
		iArrCols(28) = ConvSPChars(EG1_exp_grp(LngRow,S526_EG1_E9_plant_cd))				'공장 
		iArrCols(29) = ConvSPChars(EG1_exp_grp(LngRow,S526_EG1_E10_spec))					'품목규격 
		iArrCols(30) = ConvSPChars(EG1_exp_grp(LngRow,S526_EG1_E6_ret_item_flag))			'반품여부 
		iArrCols(31) = iArrCols(9)															'변경전금액 
		iArrCols(32) = iArrCols(6)															'변경전 VAT 포함여부 
		iArrCols(33) = iArrCols(14)															'변경전부가세액 
		iArrCols(34) = iArrCols(9)															' 변경전금액 
		iArrCols(35) = iArrCols(14)															' 변경전부가세액 

		iArrCols(36) = ConvSPChars(EG1_exp_grp(LngRow,S526_EG1_E13_OUT_NO))
		iArrCols(37) = ConvSPChars(EG1_exp_grp(LngRow,S526_EG1_E13_TRANS_TIME))
		iArrCols(38) = ConvSPChars(EG1_exp_grp(LngRow,S526_EG1_E13_OUT_TYPE))
		iArrCols(39) = ConvSPChars(EG1_exp_grp(LngRow,S526_EG1_E13_CREATE_TYPE))
		iArrCols(40) = ConvSPChars(EG1_exp_grp(LngRow,S526_EG1_E13_DN_NO))
		iArrCols(41) = ConvSPChars(EG1_exp_grp(LngRow,S526_EG1_E13_DN_SEQ))

		iArrCols(42) = ConvSPChars(EG1_exp_grp(LngRow,S526_EG1_E13_pgm_name))        '2008-06-17 1:53오전 :: hanc
		iArrCols(43) = UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(LngRow,S526_EG1_E13_pgm_price),lgCurrency,ggUnitCostNo, "X" , "X")		'단가 ConvSPChars(EG1_exp_grp(LngRow,S526_EG1_E13_pgm_price))        '2008-06-17 1:53오전 :: hanc
		iArrCols(44) = LngMaxRow + LngRow 
			
   		iArrRows(LngRow) = Join(iArrCols, gColSep)			
	Next
	
	Response.Write "<Script language=vbs> " & vbCr   
	Response.Write "With parent " & vbCr   
    Response.Write " .frm1.vspdData.Redraw = False  " & vbCr      
    Response.Write " .ggoSpread.Source = .frm1.vspdData	" & vbCr
    Response.Write " .ggoSpread.SSShowDataByClip """ & Join(iArrRows, gColSep & gRowSep) & gColSep & gRowSep & """ ,""F""" & vbCr
    
    Response.Write " .lgStrPrevKey = """ & iStrNextKey  & """" & vbCr  
    Response.Write " .DbQueryOk " & vbCr   
    
    Response.Write " If .frm1.HPostFlag.Value = """ & "Y" & """ Then " & vbCr	         
	Response.Write " .SetPostYesSpreadColor(" & LngMaxRow & ")  " & vbCr	
	Response.Write " Else " & vbCr
	Response.Write " .SetQuerySpreadColor(" & LngMaxRow & ") " & vbCr
	Response.Write " End If	" & vbCr
	Response.Write " .frm1.vspdData.Redraw = True  " & vbCr
	Response.Write "End With " & vbCr   
	Response.Write "</Script> "																				         & vbCr          

Case CStr(UID_M0002)																'☜: 저장 요청을 받음 
    Err.Clear																		'☜: Protect system from crashing
    Dim iErrorPosition
    Dim itxtSpread,itxtHBillNo
    
    Set iPS7G121 = Server.CreateObject("PS7G121_KO441.cSBillDtlSvr")  
    
    If CheckSYSTEMError(Err,True) = True Then
		Response.End		
    End If
	
	itxtSpread = Trim(Request("txtSpread"))
	itxtHBillNo = Trim(Request("txtHBillNo"))
	
	'for 구주Tax
    '@@@@@@@@@@@
     pvCB = "F"
    '@@@@@@@@@@@
	Call iPS7G121.S_MAINT_BILL_DTL_SVR (pvCB, gStrGlobalCollection ,itxtSpread , iErrorPosition, _
		                                itxtHBillNo )
	
	If CheckSYSTEMError2(Err, True, iErrorPosition & "행","","","","") = True Then
       Set iPS7G121 = Nothing
	   Response.End 
	End If
   
    Set iPS7G121 = Nothing	
	
	Response.Write "<Script language=vbs> " & vbCr         
    Response.Write " Parent.DBSaveOk "      & vbCr   
    Response.Write "</Script> "    
	    
Case CStr("PostFlag")																'☜: 확정 요청 

	Err.Clear														'☜: Protect system from crashing

	Dim itxtHBillNoPost
	
    Set iPS7G115 = Server.CreateObject("PS7G115_KO441.cSPostOpenArSvr")  
    
    If CheckSYSTEMError(Err,True) = True Then
		Response.End		
    End If
	
	itxtHBillNoPost = Trim(Request("txtHBillNo"))
	
	'for 구주Tax
    '@@@@@@@@@@@
     pvCB = "F"
    '@@@@@@@@@@@
	Call iPS7G115.S_POST_OPEN_AR_SVR(pvCB, gStrGlobalCollection,itxtHBillNoPost)
    
	If CheckSYSTEMError(Err,True) = True Then
		Set iPS7G115 = Nothing
		Response.End		
    End If
	
	Set iPS7G115 = Nothing	
	
	Response.Write "<Script language=vbs> " & vbCr         
    Response.Write " Parent.DBSaveOk "      & vbCr   
    Response.Write "</Script> "    

Case "VirtualOut"    
	Dim EG1_export_group
	Dim istrData
	Const C_OUT_NO = 0
	Const C_BP_CD = 1
	Const C_BP_NM = 2
	Const C_ITEM_CD = 3
	Const C_ITEM_NM = 4
	Const C_PLANT_CD = 5
	Const C_OUT_TYPE = 6
	Const C_OUT_TYPE_NM = 7
	Const C_OUT_TYPE_SUB = 8
	Const C_OUT_TYPE_SUB_NM = 9
	Const C_CHARGE_GB = 10
	Const C_GI_QTY = 11
	Const C_GI_UNIT = 12
	Const C_LOT_NO = 13
	Const C_LOT_SEQ = 14
	Const C_ACTUAL_GI_DT = 15
	Const C_TRANS_TIME = 16
	Const C_CREATE_TYPE = 17
	Const C_DN_NO = 18
	Const C_DN_SEQ = 19
	Const C_TRACKING_NO = 20
	Const C_VAT_INC_FLAG = 21
	Const C_VAT_RATE = 22
	Const C_VAT_TYPE = 23
	Const C_VAT_TYPE_NM = 24
	Const C_PRICE = 25
	Const C_SPEC = 26
	Const C_CUST_LOT_NO = 27

	Err.Clear														'☜: Protect system from crashing
	
  Set iPS7G115 = Server.CreateObject("PS7G116_KO441.cILstVirtualOut")  
    
  If CheckSYSTEMError(Err,True) = True Then
		Response.End		
  End If
	
	Call iPS7G115.I_LIST_VIRTUAL_OUT(gStrGlobalCollection, ""&Request("txtSpread"), EG1_export_group)
    
	If CheckSYSTEMError(Err,True) = True Then
		Set iPS7G115 = Nothing
		Response.End		
  End If
	
	Set iPS7G115 = Nothing	
	If isArray(EG1_export_group) Then 
		For LngRow = 0 To Ubound(EG1_export_group,1)
				istrData = Chr(11) &  ConvSPChars(EG1_export_group(LngRow,C_ITEM_CD))'1
				istrData = istrData & Chr(11) &  ConvSPChars(EG1_export_group(LngRow,C_ITEM_NM))'2
				istrData = istrData & Chr(11) &  ConvSPChars(EG1_export_group(LngRow,C_TRACKING_NO))'3
				istrData = istrData & Chr(11) &  UNINumClientFormat(EG1_export_group(LngRow, C_GI_QTY), ggQty.DecPoint, 0)'4
				istrData = istrData & Chr(11) &  ConvSPChars(EG1_export_group(LngRow,C_GI_UNIT))'5
				istrData = istrData & Chr(11) &  ConvSPChars(EG1_export_group(LngRow,C_VAT_INC_FLAG))'6
				
				If EG1_export_group(LngRow,C_VAT_INC_FLAG) = "1" Then
					istrData = istrData & Chr(11) &  "별도"'7
				Else
					istrData = istrData & Chr(11) &  "포함"'7
				End If
	
				istrData = istrData & Chr(11) &  UNIConvNumDBToCompanyByCurrency(EG1_export_group(LngRow,C_PRICE),lgCurrency,ggUnitCostNo, "X" , "X")'8
				istrData = istrData & Chr(11) &  "0"'9
				istrData = istrData & Chr(11) &  ConvSPChars(EG1_export_group(LngRow,C_VAT_TYPE))'10
				istrData = istrData & Chr(11) 
				istrData = istrData & Chr(11) &  ConvSPChars(EG1_export_group(LngRow,C_VAT_TYPE_NM))'12
				istrData = istrData & Chr(11) &  "0"'13
				istrData = istrData & Chr(11) &  "0"'14
				istrData = istrData & Chr(11) &  "0"'15
				istrData = istrData & Chr(11) &  "0"'16
				istrData = istrData & Chr(11) &  "0"'17
				istrData = istrData & Chr(11) &  "0"'18
				istrData = istrData & Chr(11) '19
				istrData = istrData & Chr(11) '20
				istrData = istrData & Chr(11) '21
				istrData = istrData & Chr(11) '22 
				istrData = istrData & Chr(11) '23
				istrData = istrData & Chr(11) '24
				istrData = istrData & Chr(11) '25
				istrData = istrData & Chr(11) '26
				istrData = istrData & Chr(11) '27
				istrData = istrData & Chr(11) &  ConvSPChars(EG1_export_group(LngRow,C_PLANT_CD))'28
				istrData = istrData & Chr(11) '29
				istrData = istrData & Chr(11) & "N" '30
				istrData = istrData & Chr(11) &  "0"'31
				istrData = istrData & Chr(11) &  ConvSPChars(EG1_export_group(LngRow,C_VAT_INC_FLAG))'32
				istrData = istrData & Chr(11) &  "0"'33
				istrData = istrData & Chr(11) &  "0"'34
				istrData = istrData & Chr(11) &  "0"'35
				istrData = istrData & Chr(11) &  ConvSPChars(EG1_export_group(LngRow,C_OUT_NO))'36
				istrData = istrData & Chr(11) &  ConvSPChars(EG1_export_group(LngRow,C_TRANS_TIME))'37
				istrData = istrData & Chr(11) &  ConvSPChars(EG1_export_group(LngRow,C_OUT_TYPE))'38
				istrData = istrData & Chr(11) &  ConvSPChars(EG1_export_group(LngRow,C_CREATE_TYPE))'39
				istrData = istrData & Chr(11) &  ConvSPChars(EG1_export_group(LngRow,C_LOT_NO))'40
				istrData = istrData & Chr(11) &  ConvSPChars(EG1_export_group(LngRow,C_CUST_LOT_NO))'41
				istrData = istrData & Chr(11) &  LngMaxRow + LngRow 	 & Chr(12)		
	
			Response.Write "<Script language=vbs> " & vbCr   
			Response.Write "With parent " & vbCr   
			Response.Write " .frm1.vspdData.Redraw = False  " & vbCr      
			Response.Write " .ggoSpread.Source = .frm1.vspdData	" & vbCr
			Response.Write " .ggoSpread.SSShowDataByClip """ & istrData & """,""F"""& vbCr	
			Response.Write " .frm1.vspdData.Redraw = True  " & vbCr
			Response.Write " .VirtualOutOk()"& vbCr	
			Response.Write "End With " & vbCr   
			Response.Write "</Script> "																				         & vbCr          
		
		Next
	End If
End Select

'==============================================================================
' 사용자 정의 서버 함수 
'==============================================================================
%>

<SCRIPT LANGUAGE=VBSCRIPT RUNAT=SERVER>
'============================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================
Sub SubMakeSQLStatements()
	Dim strSelectList, strFromList
	
	strSelectList = "SELECT * "
	strFromList = " FROM dbo.ufn_s_GetBillHdrInfo( " & FilterVar(Request("txtConBillNo"), "''", "S") & " , " & FilterVar("N", "''", "S") & " , " & FilterVar("N", "''", "S") & " , " & FilterVar("Y", "''", "S") & " , " & FilterVar("Y", "''", "S") & " , " & FilterVar("Q", "''", "S") & " , " & FilterVar("N", "''", "S") & " )"
	lgstrsql = strSelectList & strFromList
	'call svrmsgbox(lgstrsql, vbinformation, i_mkscript)
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub
</SCRIPT>
