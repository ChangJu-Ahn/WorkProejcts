<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 수주관리 
'*  3. Program ID           : S3171MB1
'*  4. Program Name         : STO수주정보 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2002/11/11
'*  9. Modifier (First)     : cho in kuk
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next
														
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*","NOCOOKIE","MB")
Call LoadBNumericFormatB("I","*","NOCOOKIE","MB")

Call HideStatusWnd

Dim I1_s_so_hdr	
Dim I2_s_so_hdr	
Const l2_s_so_no = 0
Const l2_s_cfm_flag = 1	
Redim I2_s_so_hdr(1)

Dim iLngRow	
Dim iLngMaxRow
Dim istrData
Dim iStrPrevKey
Dim iStrNextKey
Dim I1_command
Dim istrMode
Dim iPS3G152
Const C_SHEETMAXROWS_D  = 100	
Dim C_s_so_no
Dim C_s_so_seq

'--------------
' 수주헤더정보 
'--------------

Dim E1_s_so_hdr	
Const E1_so_no = 0
Const E1_so_dt = 1
Const E1_req_dlvy_dt = 2
Const E1_cfm_flag = 3
Const E1_price_flag = 4
Const E1_cur = 5
Const E1_xchg_rate = 6
Const E1_net_amt = 7
Const E1_net_amt_loc = 8
Const E1_cust_po_no = 9
Const E1_cust_po_dt = 10
Const E1_sales_cost_center = 11
Const E1_deal_type = 12
Const E1_pay_meth = 13
Const E1_pay_dur = 14
Const E1_trans_meth = 15
Const E1_vat_inc_flag = 16
Const E1_vat_type = 17
Const E1_vat_rate = 18
Const E1_vat_amt = 19
Const E1_vat_amt_loc = 20
Const E1_origin_cd = 21
Const E1_valid_dt = 22
Const E1_contract_dt = 23
Const E1_ship_dt_txt = 24
Const E1_pack_cond = 25
Const E1_inspect_meth = 26
Const E1_incoterms = 27
Const E1_dischge_city = 28
Const E1_dischge_port_cd = 29
Const E1_loading_port_cd = 30
Const E1_beneficiary = 31
Const E1_manufacturer = 32
Const E1_agent = 33
Const E1_remark = 34
Const E1_pre_doc_no = 35
Const E1_lc_flag = 36
Const E1_rel_dn_flag = 37
Const E1_rel_bill_flag = 38
Const E1_ret_item_flag = 39
Const E1_sp_stk_flag = 40
Const E1_ci_flag = 41
Const E1_export_flag = 42
Const E1_so_sts = 43
Const E1_insrt_user_id = 44
Const E1_insrt_dt = 45
Const E1_updt_user_id = 46
Const E1_updt_dt = 47
Const E1_ext1_qty = 48
Const E1_ext2_qty = 49
Const E1_ext3_qty = 50
Const E1_ext1_amt = 51
Const E1_ext2_amt = 52
Const E1_ext3_amt = 53
Const E1_ext1_cd = 54
Const E1_maint_no = 55
Const E1_ext3_cd = 56
Const E1_pay_type = 57
Const E1_pay_terms_txt = 58
Const E1_dn_parcel_flag = 59
Const E1_to_biz_area = 60
Const E1_to_biz_grp = 61
Const E1_biz_area = 62
Const E1_to_biz_org = 63
Const E1_to_biz_cost_center = 64
Const E1_ship_dt = 65
Const E1_auto_dn_flag = 66
Const E1_ext2_cd = 67
Const E1_bank_cd = 68
Const E1_sales_grp = 69
Const E1_sales_grp_nm = 70
Const E1_so_type = 71
Const E1_so_type_nm = 72
Const E1_bill_to_party = 73
Const E1_bill_to_party_type = 74
Const E1_bill_to_party_nm = 75
Const E1_ship_to_party = 76
Const E1_ship_to_party_type = 77
Const E1_ship_to_party_nm = 78
Const E1_sold_to_party = 79
Const E1_sold_to_party_type = 80
Const E1_sold_to_party_nm = 81
Const E1_payer = 82
Const E1_payer_type = 83
Const E1_payer_nm = 84
Const E1_sales_org = 85
Const E1_sales_org_nm = 86
Const E1_bank_nm = 87
Const E1_deal_type_nm = 88
Const E1_vat_type_nm = 89
Const E1_pay_meth_nm = 90
Const E1_incoterms_nm = 91
Const E1_pack_cond_nm = 92
Const E1_inspect_meth_nm = 93
Const E1_trans_meth_nm = 94
Const E1_vat_inc_flag_nm = 95
Const E1_pay_type_nm = 96
Const E1_loading_port_nm = 97
Const E1_dischge_port_nm = 98
Const E1_origin_nm = 99
Const E1_manufacturer_nm = 100
Const E1_agent_nm = 101
Const E1_beneficiary_nm = 102
Const E1_currency_desc = 103
Const E1_biz_area_nm = 104
Const E1_to_biz_grp_nm = 105
    

'--------------
' 수주내역정보 
'--------------
Dim EG1_exp_grp													
Const EG1_so_seq         = 0   '---수주순번 
Const EG1_hs_no          = 1   'HS부호 
Const EG1_so_price       = 2   '단가 
Const EG1_net_amt        = 3   '금액(수주순금액[거래화폐])
Const EG1_so_qty         = 4   '수량(수주량)
Const EG1_bonus_qty      = 5   '덤수량(할증수량[덤]) 
Const EG1_req_qty        = 6   '출고요청량 
Const EG1_req_bonus_qty  = 7
Const EG1_bill_qty       = 8   '---매출수량 
Const EG1_so_unit        = 9   '단위 
Const EG1_lc_qty         = 10
Const EG1_tol_more_rate  = 11  '과부족허용율(+)
Const EG1_tol_less_rate  = 12  '과부족허용율(-)
Const EG1_close_flag     = 13
Const EG1_so_status      = 14  '---수주진행상태 
Const EG1_remark         = 15  '비고 
Const EG1_cust_item_cd   = 16
Const EG1_gi_qty         = 17
Const EG1_gi_bonus_qty   = 18
Const EG1_pre_doc_seq    = 19   '---이전수주순번 
Const EG1_vat_amt        = 20   'VAT금액 
Const EG1_dlvy_dt        = 21   '납기일 
Const EG1_tracking_no    = 22   'Tracking No
Const EG1_so_base_qty    = 23   '---재고수량 
Const EG1_bonus_base_qty = 24   '---덤재고수량 
Const EG1_dn_seq         = 25   '이전출하순번 
Const EG1_cust_po_seq    = 26
Const EG1_bom_num        = 27
Const EG1_price_flag     = 28   '단가구분(가단가:N, 진단가:Y)
Const EG1_ctp_times      = 29   '---CTP Time
Const EG1_pur_qty        = 30
Const EG1_atp_flag       = 31
Const EG1_pre_doc_no     = 32   '---이전수주번호 
Const EG1_ext1_qty       = 33
Const EG1_ext2_qty       = 34
Const EG1_ext3_qty       = 35
Const EG1_ext1_amt       = 36
Const EG1_ext2_amt       = 37
Const EG1_ext3_amt       = 38
Const EG1_ext1_cd        = 39
Const EG1_ext2_cd        = 40
Const EG1_ext3_cd        = 41
Const EG1_net_amt_loc    = 42
Const EG1_vat_amt_loc    = 43
Const EG1_dn_no          = 44   '---이전출하번호 
Const EG1_lot_seq        = 45   'Lot seq
Const EG1_lot_no         = 46   'Lot no
Const EG1_ret_type       = 47   '반품유형(반품처리구분)
Const EG1_vat_type       = 48   'VAT유형 
Const EG1_vat_rate       = 49   'VAT율 
Const EG1_vat_inc_flag   = 50   '---VAT포함구분(1:별도, 2:포함)
Const EG1_sl_cd          = 51   '창고 
Const EG1_sl_nm          = 52   '---창고명 
Const EG1_plant_cd       = 53   '공장 
Const EG1_plant_nm       = 54   '---공장명 
Const EG1_item_cd        = 55   '품목 
Const EG1_item_nm        = 56   '품목명  
Const EG1_spec           = 57   '규격 
Const EG1_bp_cd          = 58   '납품처 
Const EG1_bp_nm          = 59   '---납품처명 

Const EG1_promise_dt      = 60   '출하요청일(출하예정일자)
Const EG1_vat_type_nm     = 61   'VAT유형명 
Const EG1_ret_type_nm     = 62   '반품유형명(반품처리구분명)
Const EG1_vat_inc_flag_nm = 63   'VAT포함구분명(1:별도, 2:포함)
Const EG1_aps_host        = 64   '---APSHost 
Const EG1_aps_port        = 65   '---APSPort
Const EG1_flag            = 66   '---CTPCheckFlag 

Dim EG2_exp_grp 
Const EG2_so_dt           = 0
Const EG2_req_dlvy_dt     = 1
Const EG2_cfm_flag        = 2
Const EG2_price_flag      = 3
Const EG2_cur             = 4
Const EG2_net_amt         = 5
Const EG2_cust_po_no      = 6
Const EG2_deal_type       = 7
Const EG2_pay_meth        = 8
Const EG2_vat_inc_flag    = 9
Const EG2_vat_type        = 10
Const EG2_vat_rate        = 11
Const EG2_vat_amt         = 12
Const EG2_pre_doc_no      = 13
Const EG2_ret_item_flag   = 14
Const EG2_export_flag     = 15
Const EG2_so_sts          = 16
Const EG2_maint_no        = 17
Const EG2_auto_dn_flag    = 18
Const EG2_so_type         = 19
Const EG2_bp_cd2          = 20
Const EG2_bp_cd3          = 21
Const EG2_bp_nm3          = 22
Const EG2_vat_inc_flag_nm = 23
Const EG2_ci_flag	      = 24
const EG2_dn_req_flag	  = 25


Const lsConfirm = "CONFIRM"													' 확정처리시 
istrMode = Request("txtMode")												' 현재 상태를 받음 

Select Case istrMode

Case CStr(UID_M0001)														' 현재 조회/Prev/Next 요청을 받음 

    Err.Clear     

    '--------------------------------------------------------------------------------------------------------
    ' 수주 HDR와 DTL를 읽어온다.
    '--------------------------------------------------------------------------------------------------------
    C_s_so_no = Trim(Request("txtConSoNo"))    
    
    iStrPrevKey = Trim(Request("lgStrPrevKey"))    
    
    If iStrPrevKey <> "" then
		C_s_so_seq = iStrPrevKey
    Else
		C_s_so_seq = 0
    End If
    
	Set iPS3G104 = Server.CreateObject ("PS3G104.CsListStoSalesOrder")
    
	If CheckSYSTEMError(Err, True) = True Then
		Set iPS3G104 = Nothing	
		Response.end
	End If
	
	Call iPS3G104.S_LIST_STO_SALES_ORDER(gStrGlobalCollection, C_SHEETMAXROWS_D, C_s_so_no, C_s_so_seq, _
										 E1_s_so_hdr, EG1_exp_grp, EG2_exp_grp)
											
	If CheckSYSTEMError(Err, True) = True Then
		Set iPS3G104 = Nothing		                                                 '☜: Unload Comproxy DLL
%>
		<Script Language=vbscript>
		With parent.frm1
			If .vspdData.MaxRows = 0 Then
				.btnConfirm.disabled   = True
				.btnConfirm.value      = "확정처리"
				.btnDNCheck.disabled   = True
				.btnATPCheck.disabled  = True
				.btnCTPCheck.disabled  = True
				.btnAvlStkRef.disabled = True
			End If

			If Trim(.txtHPreSONo.value) <> "" And UCase(Trim(.HRetItemFlag.value)) = "Y" Then
				parent.SetToolbar "11001001000111"
			ElseIf Trim(.txtHPreSONo.value) = "" And UCase(Trim(.HRetItemFlag.value)) = "Y" Then
				parent.SetToolbar "11001001000111"
			ElseIf UCase(Trim(.HRetItemFlag.value)) <> "Y" Then
				parent.SetToolbar "11001001000111"
			Else
				parent.SetToolbar "11001001000111"
			End If
			
            .txtConSoNo.focus
		End With
		</Script>
<%		
		Response.End          
	End If
	
	Set iPS3G104 = Nothing    
 

    '----------------------------
	' 수주헤더정보를 표시한다.
	'----------------------------
%>
<Script Language=vbscript>
	With parent		
<%
		Dim lgCurrency																			'항상 거래화폐가 우선 
		lgCurrency = ConvSPChars(E1_s_so_hdr(E1_cur))
%>
		.frm1.txtCurrency.value = "<%=lgCurrency%>"
		parent.CurFormatNumericOCX	
		
		.frm1.txtSoNo.value				=					 "<%=ConvSPChars(E1_s_so_hdr(E1_so_no))%>"
		.frm1.txtSoType.value			=					 "<%=ConvSPChars(E1_s_so_hdr(E1_so_type))%>"
		.frm1.txtSoTypeNm.value			=					 "<%=ConvSPChars(E1_s_so_hdr(E1_so_type_nm))%>"
		.frm1.txtSoDt.text				=			 "<%=UNIDateClientFormat(E1_s_so_hdr(E1_so_dt))%>"		
		.frm1.txtSoldtoparty.value		=					 "<%=ConvSPChars(E1_s_so_hdr(E1_sold_to_party))%>"     
		.frm1.txtSoldtopartyNm.value	=					 "<%=ConvSPChars(E1_s_so_hdr(E1_sold_to_party_nm))%>"  	
		.frm1.txtSalesGrp.value			=					 "<%=ConvSPChars(E1_s_so_hdr(E1_sales_grp))%>"  
		.frm1.txtSalesGrpNm.value		=					 "<%=ConvSPChars(E1_s_so_hdr(E1_sales_grp_nm))%>"  		
		.frm1.txtDealType.value			=					 "<%=ConvSPChars(E1_s_so_hdr(E1_deal_type))%>"  
		.frm1.txtDealTypeNm.value		=					 "<%=ConvSPChars(E1_s_so_hdr(E1_deal_type_nm))%>"
		.frm1.txtCustpono.value			=					 "<%=ConvSPChars(E1_s_so_hdr(E1_cust_po_no))%>"
		.frm1.txtPaymeth.value			=					 "<%=ConvSPChars(E1_s_so_hdr(E1_pay_meth))%>"
		.frm1.txtPaymethNm.value		=					 "<%=ConvSPChars(E1_s_so_hdr(E1_pay_meth_nm))%>"
		.frm1.txtNetAmt.text			=   "<%=UNINumClientFormatByCurrency(E1_s_so_hdr(E1_net_amt), lgCurrency, ggAmtOfMoneyNo)%>"
		.frm1.txtXchgRate.text			=			  "<%=UNINumClientFormat(E1_s_so_hdr(E1_xchg_rate), ggExchRate.DecPoint,0)%>"
		.frm1.txtRemark.value			=					 "<%=ConvSPChars(E1_s_so_hdr(E1_remark))%>"
		.frm1.txtHConfirmFlg.value		=					 "<%=ConvSPChars(EG2_exp_grp(EG2_cfm_flag))%>"         '---
  	
		If "<%=E1_s_so_hdr(E1_cfm_flag)%>" = "Y" Then
			.frm1.rdoCfm_flag1.checked = True	
			.frm1.txtRadioFlag.value = .frm1.rdoCfm_flag1.value		
			.frm1.RdoConfirm.value = "N"
			.frm1.btnConfirm.value = "확정취소"
		ElseIf "<%=E1_s_so_hdr(E1_cfm_flag)%>" = "N" Then
			.frm1.rdoCfm_flag2.checked = True 	
			.frm1.txtRadioFlag.value = .frm1.rdoCfm_flag2.value														    		
			.frm1.RdoConfirm.value = "Y"
			.frm1.btnConfirm.value = "확정처리"
		Else
			.frm1.RdoConfirm.value = "Y"
			.frm1.btnConfirm.value = "확정처리"
		End IF
		
		.frm1.txtHSoNo.value = "<%=ConvSPChars(Request("txtConSoNo"))%>"		
		 		
		If parent.lgIntFlgMode = parent.parent.OPMD_CMODE Then parent.CurFormatNumSprSheet		
		.lgIntFlgMode = parent.parent.OPMD_UMODE				
		parent.HideLotRetField
		
	End With
</Script>		
<%   
	'----------------------------
	' 수주내역정보를 표시한다.
	'----------------------------
%>
<Script Language=vbscript>
    Dim LngLastRow      
    Dim iLngMaxRow       
    Dim iLngRow          
    Dim strTemp
    Dim istrData
	    
	With parent
		iLngMaxRow = .frm1.vspdData.MaxRows
<%        
		For iLngRow = 0 To UBound(EG1_exp_grp,1)
		    If iLngRow < C_SHEETMAXROWS_D  Then
		    Else
		       iStrNextKey = ConvSPChars(EG1_exp_grp(iLngRow, EG1_so_seq)) 
               Exit For
            End If 	           
%>           
			istrData = istrData & Chr(11) & "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_item_cd))%>"     '품목코드			
			istrData = istrData & Chr(11)                                                             '품목코드팝업			
			istrData = istrData & Chr(11) & "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_item_nm))%>"     '품목명			
			istrData = istrData & Chr(11) & "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_spec))%>"        '규격			
			istrData = istrData & Chr(11) & "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_so_unit))%>"     '단위 
			istrData = istrData & Chr(11)                                                             '단위팝업			
			istrData = istrData & Chr(11) & "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_tracking_no))%>" 'Tracking No.
			istrData = istrData & Chr(11) & "<%=UNINumClientFormat(cdbl(EG1_exp_grp(iLngRow, EG1_so_qty)) - cdbl(EG1_exp_grp(iLngRow, EG1_req_qty)), ggQty.DecPoint, 0)%>" '수주잔량 
			istrData = istrData & Chr(11) & "<%=UNINumClientFormat(EG1_exp_grp(iLngRow, EG1_so_qty), ggQty.DecPoint, 0)%>"          '수량 
			istrData = istrData & Chr(11) & "<%=UNINumClientFormatByCurrency(EG1_exp_grp(iLngRow, EG1_so_price), lgCurrency, ggUnitCostNo)%>" '단가 
		    istrData = istrData & Chr(11) & "0"                                                       '단가체크버튼			
		    
			Select Case "<%=EG1_exp_grp(iLngRow, EG1_price_flag)%>"                                   '단가구분 (진단가/가단가)
			Case "Y"
				istrData = istrData & Chr(11) & "진단가"
			Case "N"
				istrData = istrData & Chr(11) & "가단가"
			Case Else
				istrData = istrData & Chr(11)
			End Select

        	If "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_vat_inc_flag))%>" =  "2" Then   	    	  
		     	istrData = istrData & Chr(11) & "<%=UNINumClientFormatByCurrency(cdbl(EG1_exp_grp(iLngRow, EG1_net_amt)) + cdbl(EG1_exp_grp(iLngRow, EG1_vat_amt)), lgCurrency, ggAmtOfMoneyNo)%>"
			Else
				istrData = istrData & Chr(11) & "<%=UNINumClientFormatByCurrency(EG1_exp_grp(iLngRow, EG1_net_amt), lgCurrency, ggAmtOfMoneyNo)%>"							
			End If		

			istrData = istrData & Chr(11) & "<%=UNINumClientFormatByCurrency(EG1_exp_grp(iLngRow, EG1_net_amt), lgCurrency, ggAmtOfMoneyNo)%>" '수주순금액(거래화폐)
			istrData = istrData & Chr(11) & "<%=UNINumClientFormatByTax(EG1_exp_grp(iLngRow, EG1_vat_amt),lgCurrency,ggAmtOfMoneyNo)%>"   'VAT 금액 
    		istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_plant_cd))%>"     '공장코드 
			istrData = istrData & Chr(11)                                                                               '공장팝업			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_plant_nm))%>"     '공장명									
			istrData = istrData & Chr(11) & "<%=UNIDateClientFormat(EG1_exp_grp(iLngRow, EG1_dlvy_dt))%>"      '납기일			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_bp_cd))%>"        '납품처			
			istrData = istrData & Chr(11)                                                                               '납품처팝업			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_bp_nm))%>"        '납품처명			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_hs_no))%>"        'HS Code			
			istrData = istrData & Chr(11)                                                                               'Hs Code Popup
			istrData = istrData & Chr(11) & "<%=UNINumClientFormat(EG1_exp_grp(iLngRow, EG1_tol_more_rate), ggExchRate.DecPoint, 0)%>" '과부족허용율(+)
			istrData = istrData & Chr(11) & "<%=UNINumClientFormat(EG1_exp_grp(iLngRow, EG1_tol_less_rate), ggExchRate.DecPoint, 0)%>" '과부족허용율(-)
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_vat_type))%>"        'VAT Type
			istrData = istrData & Chr(11)                                                                                  'VAT Type Popup
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_vat_type_nm))%>"     'VAT Name
			istrData = istrData & Chr(11) & "<%=UNINumClientFormat(EG1_exp_grp(iLngRow, EG1_vat_rate), ggExchRate.DecPoint, 0)%>" 'VAT Rate
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_vat_inc_flag))%>"    'VAT포함구분 			
			'istrData = istrData & Chr(11) &                  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_vat_inc_flag_nm))%>" 'VAT포함구분명 
			Select Case "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_vat_inc_flag))%>"
			Case "1"
				istrData = istrData & Chr(11) & "별도"
			Case "2"
				istrData = istrData & Chr(11) & "포함"
			Case Else
				istrData = istrData & Chr(11)
			End Select
			
     		istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_ret_type))%>"        '반품유형 
			istrData = istrData & Chr(11)                                                                                  '반품유형팝업 
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_ret_type_nm))%>"     '반품유형명			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_lot_no))%>"          'Lot No			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_lot_seq))%>"         'Lot Seq			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_dn_no))%>"           '이전출하번호			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_dn_seq))%>"          '이전출하순번			
			istrData = istrData & Chr(11) & "<%=UNIDateClientFormat(EG1_exp_grp(iLngRow, EG1_promise_dt))%>"      '출하요청일											
			istrData = istrData & Chr(11) & "<%=UNINumClientFormat(EG1_exp_grp(iLngRow, EG1_bonus_qty), ggQty.DecPoint, 0)%>" '할증수량(덤)			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_sl_cd))%>"           '창고코드			
			istrData = istrData & Chr(11)                                                                                  '창고팝업			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_sl_nm))%>"           '창고명			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_remark))%>"          '비고			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_so_status))%>"       '수주진행상태			
			istrData = istrData & Chr(11) & "<%=UNINumClientFormat(EG1_exp_grp(iLngRow, EG1_bill_qty), ggQty.DecPoint, 0)%>"      '매출수량			
			istrData = istrData & Chr(11) & "<%=UNINumClientFormat(EG1_exp_grp(iLngRow, EG1_so_base_qty), ggQty.DecPoint, 0)%>"   '재고수량			
			istrData = istrData & Chr(11) & "<%=UNINumClientFormat(EG1_exp_grp(iLngRow, EG1_bonus_base_qty), ggQty.DecPoint, 0)%>"'덤재고수량			
			istrData = istrData & Chr(11) & ""                                                                             '관리순번			
			istrData = istrData & Chr(11) & ""                                                                             '주문서순번			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_aps_host))%>"        'APSHost			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_aps_port))%>"        'APSPort			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_ctp_times))%>"       'CTPTimes			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_flag))%>"            'CTPCheckFlag			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_so_seq))%>"          '수주순번			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_pre_doc_no))%>"      '이전수주번호			
			istrData = istrData & Chr(11) &  "<%=ConvSPChars(EG1_exp_grp(iLngRow, EG1_pre_doc_seq))%>"     '이전수주순번 
			istrData = istrData & Chr(11) & iLngMaxRow + <%=iLngRow%>			
			istrData = istrData & Chr(11) & Chr(12)			
<%      
		Next
%>    

		.ggoSpread.Source = .frm1.vspdData 
		.ggoSpread.SSShowData istrData
	
		.lgStrPrevKey = "<%=iStrNextKey%>"
		
    	.frm1.txtHSoNo.value = "<%=ConvSPChars(Request("txtConSoNo"))%>"   ' Request값을 hidden input으로 넘겨줌 
					
		If "<%=EG2_exp_grp(EG2_cfm_flag)%>" = "Y" Then         '확정버튼 처리				
			If "<%=EG2_exp_grp(EG2_auto_dn_flag)%>" = "N" Or "<%=EG2_exp_grp(EG2_so_sts)%>" = 1 Then  
				.frm1.btnDNCheck.disabled = True              '출하버튼 처리 
			ElseIf "<%=EG2_exp_grp(EG2_auto_dn_flag)%>" = "Y" And "<%=EG2_exp_grp(EG2_so_sts)%>" <> 1 Then
				'임시로 막음. 2003-10-10
				'.frm1.btnDNCheck.disabled = False
				.frm1.btnDNCheck.disabled = True
			End IF								
			.frm1.btnATPCheck.disabled = True                 'ATP CHECK버튼 처리 
		Else				
			.frm1.btnDNCheck.disabled = True                  '출하버튼 처리 							
			.frm1.btnATPCheck.disabled = False                'ATP CHECK버튼 처리	
		End IF
		.DbQueryOk
	
	End With

</Script>
<%																			

Case CStr(UID_M0002)																'☜: 저장 요청을 받음 
	
	Dim orPosition
	Dim E1_s_so_no
	Dim pvCB
	
	Dim imp_s_so_hdr
	Const S103_I1_command = 0
    Const S103_I1_so_no = 1
    Const S103_I1_so_dt = 2
    Const S103_I1_cfm_flag = 3
    Const S103_I1_cur = 4
    Const S103_I1_cust_po_no = 5
    Const S103_I1_deal_type = 6
    Const S103_I1_pay_meth = 7
    Const S103_I1_remark = 8
    Const S103_I1_xchg_rate = 9
    Const S103_I1_sales_grp = 10
    Const S103_I1_sold_to_party = 11
    Const S103_I1_ship_to_party = 12
    Const S103_I1_so_type = 13
    
    ReDim imp_s_so_hdr(S103_I1_so_type) 
    								
    Err.Clear																		'☜: Protect system from crashing
    
	If Request("txtMaxRows") = "" Then
		Call ServerMesgBox("MaxRows 조건값이 비어있습니다!",vbInformation, I_MKSCRIPT)              
		Response.End 
	End If	

    imp_s_so_hdr(S103_I1_command) = "U"
    imp_s_so_hdr(S103_I1_so_no) = UCase(Trim(Request("txtSoNo")))
    imp_s_so_hdr(S103_I1_so_dt) = UNIConvDate(Trim(Request("txtSoDt")))
    imp_s_so_hdr(S103_I1_cfm_flag) = UCase(Trim(Request("txtRadioFlag")))
    imp_s_so_hdr(S103_I1_cur) = UCase(Trim(Request("txtCurrency")))
    imp_s_so_hdr(S103_I1_cust_po_no) = UCase(Trim(Request("txtCustpono")))
    imp_s_so_hdr(S103_I1_deal_type) = UCase(Trim(Request("txtDealType")))
    imp_s_so_hdr(S103_I1_pay_meth) = UCase(Trim(Request("txtPaymeth")))
    imp_s_so_hdr(S103_I1_remark) = Trim(Request("txtRemark"))    
    imp_s_so_hdr(S103_I1_xchg_rate) = UNIConvNum(Trim(Request("txtXchgRate")),0)    
    imp_s_so_hdr(S103_I1_sales_grp) = UCase(Trim(Request("txtSalesGrp")))
    imp_s_so_hdr(S103_I1_sold_to_party) = UCase(Trim(Request("txtSoldtoparty")))
    imp_s_so_hdr(S103_I1_ship_to_party) = UCase(Trim(Request("txtSoldtoparty")))
    imp_s_so_hdr(S103_I1_so_type) = UCase(Trim(Request("txtSoType")))   
    	
    Set iPS3G103 = Server.CreateObject("PS3G103.CsMaintStoSalesOrder")      
    
    If CheckSYSTEMError(Err,True) = True Then
       Response.End       
    End If    
	
	pvCB = "F"
	
    E1_s_so_no = iPS3G103.S_MAINT_STO_SALES_ORDER(pvCB, gStrGlobalCollection, imp_s_so_hdr, _
										  Trim(Request("txtSpread")), iErrorPosition)   

	If cStr(Err.Description) <> "" and iErrorPosition = "" Then     
		
		If CheckSYSTEMError(Err,True) = True Then 
		   Set iPS3G103 = Nothing		                                                 '☜: Unload Comproxy DLL
		End If  
    
	ElseIf cStr(Err.Description) <> "" and iErrorPosition <> "" Then  
	
	    If CheckSYSTEMError2(Err, True, iErrorPosition & "행","","","","") = True Then
		   Set iPS3G103 = Nothing
		   Response.End			
		End If
	
	End If
	
    Set iPS3G103 = Nothing 			

%>
<Script Language=vbscript>
	With parent																			
		.DbSaveOk
	End With
</Script>
<%																

Case "DNCheck"																'☜: 현재 출하요청처리 받음 

    Err.Clear                                                               '☜: Protect system from crashing
    Dim iPS3G117
    
    I1_s_so_hdr = Trim(Request("txtSoNo"))
    
    Set iPS3G117 = Server.CreateObject("PS3G117.cSCreateDnBySoSvr")          
    
    If CheckSYSTEMError(Err,True) = True Then
       Response.End
       'Exit Sub
    End If    

    Call iPS3G117.S_CREATE_DN_BY_SO_SVR(gStrGlobalCollection, I1_s_so_hdr)    
    
	If CheckSYSTEMError(Err,True) = True Then
       Set iPS3G117 = Nothing	
       Response.End		                                               
       'Exit Sub
    End If
	
	Set iPS3G117 = Nothing	

%>
<Script Language=vbscript>
'=	parent.FncQuery
	parent.DbSaveOk()
</Script>		
<%

Case "PRICE"																'☜: 현재 출하요청처리 받음 

    Err.Clear                                                               '☜: Protect system from crashing
    
    Dim I4_s_so_dtl
    Const S321_I4_so_unit = 0
    Const S321_I4_so_qty = 1
    ReDim I4_s_so_dtl(1)
    
    DIm E1_s_so_dt
    Const S321_E1_so_price = 0
    Const S321_E1_bonus_qty = 1
    
    Dim pS31121PR
    
    Dim I1_ief_supplied_select_char    
    Dim I2_b_item_item_cd    
    Dim I3_s_so_hdr_so_no 
    
    I1_ief_supplied_select_char = Trim(Request("lsPriceQty"))   
    
    I2_b_item_item_cd = Trim(Request("lsItemCode"))
    
    I3_s_so_hdr_so_no = Trim(Request("txtHSoNo"))
    
    I4_s_so_dtl(S321_I4_so_unit) = Trim(Request("lsSoUnit"))
    I4_s_so_dtl(S321_I4_so_qty) = UNIConvNum(Request("lsSoQty"),0)	

    Set pS31121PR = Server.CreateObject("PS3G112.cSGetSoPriceSvr")    

    If CheckSYSTEMError(Err,True) = True Then
		Response.End		
    End If
    
    E1_s_so_dtl = pS31121PR.S_GET_SO_PRICE_SVR (gStrGlobalCollection,I1_ief_supplied_select_char, I2_b_item_item_cd, _
												I3_s_so_hdr_so_no,I4_s_so_dtl)
       
	If CheckSYSTEMError(Err,True) = True Then
       Set pS31121PR = Nothing		                                                 '☜: Unload Comproxy DLL
       Response.End 
    End If   

    Set pS31121PR = Nothing	
  
%>
<Script Language=vbscript>

	With parent																'☜: 화면 처리 ASP 를 지칭함 

		.frm1.vspdData.Row = <%=Request("PRow")%>

		Select Case "<%=Trim(Request("lsPriceQty"))%>"
		Case "A"
			.frm1.vspdData.Col = .C_SoPrice
			.frm1.vspdData.Text = "<%=UNINumClientFormatByCurrency(E1_s_so_dtl(S321_E1_so_price), Trim(Request("txtCurrency")), ggUnitCostNo)%>"
			.frm1.vspdData.Col = .C_BonusQty
			.frm1.vspdData.Text = "<%=UNINumClientFormat(E1_s_so_dtl(S321_E1_bonus_qty), ggQty.DecPoint, 0)%>"
		Case "P"
			.frm1.vspdData.Col = .C_SoPrice
			.frm1.vspdData.Text = "<%=UNINumClientFormatByCurrency(E1_s_so_dtl(S321_E1_so_price), Trim(Request("txtCurrency")), ggUnitCostNo)%>"
		Case "Q"
			.frm1.vspdData.Col = .C_BonusQty
			.frm1.vspdData.Text = "<%=UNINumClientFormat(E1_s_so_dtl(S321_E1_bonus_qty), ggQty.DecPoint, 0)%>"
		End Select

	End With

</Script>		
<%

Case "btnCONFIRM"																	'☜: 확정처리 요청을 받음 
									
    Err.Clear																		'☜: Protect system from crashing
    Dim iPS3G150
    Redim I1_s_so_hdr(1)
    
	iLngMaxRow = CInt(Request("txtMaxRows"))											'☜: 최대 업데이트된 갯수 
	
	I1_s_so_hdr(0) = Trim(Request("txtHSoNo"))
	I1_s_so_hdr(1) = Trim(Request("RdoConfirm"))
   
    Set iPS3G150 = Server.CreateObject("PS3G150.cSConfirmSalesOrderSvr")      
    
    If CheckSYSTEMError(Err,True) = True Then
       Response.End
       'Exit Sub
    End If    

    Call iPS3G150.S_CONFIRM_SALES_ORDER_SVR(gStrGlobalCollection, I1_s_so_hdr)   
    
    If CheckSYSTEMError(Err,True) = True Then
       Set iPS3G121 = Nothing 
       Response.End
       'Exit Sub
    End If 

    Set iPS3G121 = Nothing 		  
%>
<Script Language=vbscript>
	parent.DbSaveOk()	
</Script>
<%					


Case "ItemByHsCode"															'☜: 품목별에 따른 HS CODE Change

	Dim iPB3C104
	
    Dim I1_b_item    
    Dim prE1_b_item
    Const prE1_item_cd = 0
    Const prE1_item_nm = 1
    Const prE1_formal_nm = 2
    Const prE1_spec = 3
    Const prE1_basic_unit = 4
    Const prE1_item_acct = 5
    Const prE1_item_class = 6
    Const prE1_phantom_flg = 7
    Const prE1_hs_cd = 8
    Const prE1_hs_unit = 9
    Const prE1_unit_weight = 10
    Const prE1_unit_of_weight = 11
    Const prE1_draw_no = 12
    Const prE1_item_image_flg = 13
    Const prE1_blanket_pur_flg = 14
    Const prE1_base_item_cd = 15
    Const prE1_proportion_rate = 16
    Const prE1_valid_flg = 17
    Const prE1_valid_from_dt = 18
    Const prE1_valid_to_dt = 19
    Const prE1_vat_type = 20
    Const prE1_vat_rate = 21
    
    
	Err.Clear
	
	I1_b_item = Trim(Request("ItemCd"))
	
    Set iPB3C104 = Server.CreateObject("PB3C104.cBLkUpItem")     
    
    If CheckSYSTEMError(Err,True) = True Then
       Response.End
       'Exit Sub
    End If
    
    Call iPB3C104.B_LOOK_UP_ITEM(gStrGlobalCollection, I1_b_item, , , , , prE1_b_item)	
	
	If CheckSYSTEMError(Err,True) = True Then
       Set iPB3C104 = Nothing	
       Response.End		                                               
       'Exit Sub
    End If	

%>

<Script Language="vbscript">
		With parent.frm1.vspdData
			.Row 	= "<%=Request("CRow")%>"
			.Col 	= parent.C_ItemName
			.text	= "<%=ConvSPChars(prE1_b_item(prE1_item_nm))%>"
			.Col 	= parent.C_ItemSpec
			.text	= "<%=ConvSPChars(prE1_b_item(prE1_spec))%>"
			.Col 	= parent.C_HsNo
			.text	= "<%=ConvSPChars(prE1_b_item(prE1_hs_cd))%>"
			.Col 	= parent.C_SoUnit
			.text	= "<%=ConvSPChars(prE1_b_item(prE1_basic_unit))%>"

			.Col	= parent.C_VatType

			If .text = "" Then
				If Len(parent.frm1.txtHVATType.value) Then
					.text	= parent.frm1.txtHVATType.value
				Else
					.text	= "<%=ConvSPChars(prE1_b_item(prE1_vat_type))%>"
				End If
			End If
			
			.Col	= parent.C_VatIncFlag

			If .text = "" Then 
				If Len(parent.frm1.txtHVATIncFlag.value) Then
					.Col	= parent.C_VatIncFlag
					.text	= parent.frm1.txtHVATIncFlag.value

					.Col	= parent.C_VatIncFlagNm
					Select Case parent.frm1.txtHVATIncFlag.value
					Case "1"
						parent.frm1.vspdData.Text = "별도"
					Case "2"
						parent.frm1.vspdData.Text = "포함"
					End Select
				End If			
			End If
			Call parent.SetVatType(<%=Request("CRow")%>)
			
			parent.lsPriceQty = "Q"
			Call parent.GetItemPrice(<%=Request("CRow")%>)
			Call parent.PricePadChange(<%=Request("CRow")%>)
		End With	
</Script>
<%
    Set iPB3C104 = Nothing

Case "CheckCreditlimit"															'☜: 여신한도 체크 

	Dim iPS3G113
	Dim BalanceAmt
	
    Dim I1_b_currency
    Const S324_I1_currency = 0
    Redim I1_b_currency(S324_I1_currency)
    
    Dim I2_ief_supplied
    Const S324_I2_total_currency = 0
    Redim I2_ief_supplied(S324_I2_total_currency)
    
    Dim I3_s_so_hdr
    Const S324_I3_so_no = 0
    Redim I3_s_so_hdr(S324_I3_so_no)
    
    Dim E1_exchange_variable
    Const S324_E1_num_value_15_2 = 0
    
    Dim E2_ief_supplied
    Const S324_E2_command = 0
    Const S324_E2_select_char = 1
    
    I3_s_so_hdr(S324_I3_so_no) = Trim(Request("txtHSONo"))
    I2_ief_supplied(S324_I2_total_currency) = UNIConvNum(Request("txtNetAmt"),0)
    I1_b_currency(S324_I1_currency) = Trim(Request("txtCurrency"))  
        
    Set iPS3G113 = Server.CreateObject("PS3G113.cChkSoCreditLimitSvr")  
    
    If CheckSYSTEMError(Err,True) = True Then
       Response.End
       'Exit Sub
    End If
 
    Call iPS3G113.S_CHK_SO_CREDIT_LIMIT_SVR(gStrGlobalCollection, I1_b_currency, I2_ief_supplied, _
                                            I3_s_so_hdr, , E1_exchange_variable, E2_ief_supplied)

    If cStr(Err.Description) = "B_MESSAGE" & Chr(11) & "201727" then    

           BalanceAmt = UNINumClientFormat(E1_exchange_variable(S324_E1_num_value_15_2), ggAmtOfMoney.DecPoint, 0)          
%>     
<Script Language=vbscript>
           Dim msgCreditlimit
		   Dim BalanceAmt		

		   BalanceAmt = FormatNumber(Parent.parent.UNICDbl("<%=BalanceAmt%>"), parent.parent.ggAmtofMoney.DecPoint, -2)	
		   
		   msgCreditlimit = parent.parent.DisplayMsgBox("201929", parent.VB_YES_NO, parent.gCurrency, BalanceAmt)		   
		   
		   If msgCreditlimit = vbYes Then parent.DbSave
		      

</Script>
<% 
    ElseIf cStr(Err.Description) = "B_MESSAGE" & Chr(11) & "201722" then
  
           BalanceAmt = UNINumClientFormat(E1_exchange_variable(S324_E1_num_value_15_2), parent.parent.ggAmtofMoney.DecPoint, 0)
%>     
<Script Language=vbscript>      
           Dim BalanceAmt	
           			
		   BalanceAmt = FormatNumber(Parent.parent.UNICDbl("<%=BalanceAmt%>"), parent.parent.ggAmtofMoney.DecPoint, -2)
		   
		   Call parent.parent.DisplayMsgBox("201722", "X", parent.parent.gCurrency, BalanceAmt)   	
</Script>
<%      
    Else
           If CheckSYSTEMError(Err,True) = True Then    
               Set iPS3G113 = Nothing	
               Response.End		                                               
               'Exit Sub	
           End if    
%>     
<Script Language=vbscript>     
           Call parent.DbSave()
</Script>
<%             
    END IF     	                                               
     
    Set iPS3G113 = Nothing  
   	
    
'자동출하생성에서 체크함. 출하에서는 여신체크안함 
Case "CheckADNCreditlimit"														

	Dim pS14115

    Set pS14115 = Server.CreateObject("S14115.S14115ChkAdnCreditLimitSvr")

    If Err.Number <> 0 Then
		Set pS14115 = Nothing																'☜: ComProxy UnLoad
		Call ServerMesgBox(Err.description, vbInformation, I_MKSCRIPT)
		Err.Clear						
		Response.End																		'☜: Process End
	End If
    
    pS14115.ImpSSoHdrSoNo = Trim(Request("txtHSONo"))
    pS14115.ServerLocation = ggServerIP

    pS14115.ComCfg = gConnectionString
    pS14115.Execute

	If Err.Number <> 0 Then
	   Call ServerMesgBox(Err.description , vbInformation, I_MKSCRIPT)                                 '⊙:
	   Set pS14115 = Nothing																	'☜: ComProxy UnLoad
	   Response.End																				'☜: Process End
	End If
	
%>
<Script Language=vbscript>
<%
		If Not (pS14115.OperationStatusMessage = MSG_OK_STR) Then
			If pS14115.OperationStatusMessage = 201727 Then
%>
				Dim msgCreditlimit
				msgCreditlimit = parent.parent.DisplayMsgBox("201727", parent.VB_YES_NO, "X", "X")
				If msgCreditlimit = vbYes Then parent.RunAutoDN
<%
			ElseIf pS14115.OperationStatusMessage = 201722 Then
%>
				Call parent.parent.DisplayMsgBox("201722", "X", "X", "X")
<%
			Else
%>
				Call parent.parent.DisplayMsgBox("<%=pS14115.OperationStatusMessage%>", "X", "X", "X")
<%
			End If
		Else
%>
			Call parent.RunAutoDN()
<%	
		End If
%>

</Script>
<%					

	Set pS14115 = Nothing
	Response.End

End Select

%>
