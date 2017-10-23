<%
'**********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : B1261MB1
'*  4. Program Name         : 거래처등록 
'*  5. Program Desc         : 거래처등록 
'*  6. Comproxy List        : PB5CS40.dll, PB5CS41.dll
'*  7. Modified date(First) : 2002/06/05
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Cho inkuk
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

<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*","NOCOOKIE","MB")
Call LoadBNumericFormatB("I","*","NOCOOKIE","MB")
Call HideStatusWnd	    												'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

On Error Resume Next														

Dim bankInfo	'거래처은행정보>>air

Dim strMode																	

Dim iS1C141
Dim imp_command
Dim imp_biz_partner_cd
Dim E1_b_biz_partner

Dim iS1C140
Dim I1_b_biz_partner
Dim imp_BpCd

Dim pvCB
Dim prICustomXML

Const E1_bp_cd               = 0
Const E1_bp_type             = 1
Const E1_bp_rgst_no          = 2
Const E1_bp_full_nm          = 3
Const E1_bp_nm               = 4
Const E1_bp_eng_nm           = 5
Const E1_repre_nm            = 6
Const E1_repre_rgst_no       = 7
Const E1_fnd_dt              = 8
Const E1_zip_cd              = 9
Const E1_addr1               = 10
Const E1_addr1_eng           = 11
Const E1_ind_type            = 12
Const E1_ind_class           = 13
Const E1_trade_rgst_no       = 14
Const E1_contry_cd           = 15
Const E1_province_cd         = 16
Const E1_currency            = 17
Const E1_tel_no1             = 18
Const E1_tel_no2             = 19
Const E1_fax_no              = 20
Const E1_home_url            = 21
Const E1_usage_flag          = 22
Const E1_bp_prsn_nm          = 23
Const E1_bp_contact_pt       = 24
Const E1_biz_prsn            = 25
Const E1_biz_grp             = 26
Const E1_biz_org             = 27
Const E1_deal_type           = 28
Const E1_pay_meth            = 29
Const E1_pay_dur             = 30
Const E1_pay_day             = 31
Const E1_vat_inc_flag        = 32
Const E1_vat_type            = 33
Const E1_vat_rate            = 34
Const E1_trans_meth          = 35
Const E1_trans_lt            = 36
Const E1_sale_amt            = 37
Const E1_capital_amt         = 38
Const E1_emp_cnt             = 39
Const E1_bp_grade            = 40
Const E1_comm_rate           = 41
Const E1_addr2               = 42
Const E1_addr2_eng           = 43
Const E1_addr3_eng           = 44
Const E1_pay_type            = 45
Const E1_pay_terms_txt       = 46
Const E1_credit_mgmt_flag    = 47
Const E1_credit_grp          = 48
Const E1_vat_calc_type       = 49
Const E1_deposit_flag        = 50
Const E1_bp_group            = 51
Const E1_clearance_id        = 52
Const E1_credit_rot_day      = 53
Const E1_gr_insp_type        = 54
Const E1_bp_alias_nm         = 55
Const E1_to_org              = 56
Const E1_to_grp              = 57
Const E1_pay_month           = 58
Const E1_expiry_dt           = 59
Const E1_pur_grp             = 60
Const E1_pur_org             = 61
Const E1_charge_lay_flag     = 62
Const E1_remark1             = 63
Const E1_remark2             = 64
Const E1_remark3             = 65
Const E1_close_day1          = 66
Const E1_close_day2          = 67
Const E1_close_day3          = 68
Const E1_tax_biz_area        = 69
Const E1_cash_rate           = 70
Const E1_pay_type_out        = 71
Const E1_par_bank_cd1_bp     = 72
Const E1_bank_acct_no1_bp    = 73
Const E1_bank_cd1_bp         = 74
Const E1_par_bank_cd2_bp     = 75
Const E1_bank_cd2_bp         = 76
Const E1_bank_acct_no2_bp    = 77
Const E1_par_bank_cd3_bp     = 78
Const E1_bank_cd3_bp         = 79
Const E1_bank_acct_no3_bp    = 80
Const E1_par_bank_cd1        = 81
Const E1_bank_cd1            = 82
Const E1_bank_acct_no1       = 83
Const E1_par_bank_cd2        = 84
Const E1_bank_cd2            = 85
Const E1_bank_acct_no2       = 86
Const E1_par_bank_cd3        = 87
Const E1_bank_cd3            = 88
Const E1_bank_acct_no3       = 89
Const E1_pay_month2          = 90
Const E1_pay_day2            = 91
Const E1_pay_month3          = 92
Const E1_pay_day3            = 93
Const E1_close_day1_sales    = 94
Const E1_pay_month1_sales    = 95
Const E1_pay_day1_sales      = 96
Const E1_close_day2_sales    = 97
Const E1_pay_month2_sales    = 98
Const E1_pay_day2_sales      = 99
Const E1_close_day3_sales    = 100
Const E1_pay_month3_sales    = 101
Const E1_pay_day3_sales      = 102
Const E1_ext1_qty            = 103
Const E1_ext2_qty            = 104
Const E1_ext3_qty            = 105
Const E1_ext1_amt            = 106
Const E1_ext2_amt            = 107
Const E1_ext3_amt            = 108
Const E1_ext1_cd             = 109
Const E1_ext2_cd             = 110
Const E1_ext3_cd             = 111
Const E1_in_out				 = 112					 '사내외구분 
'12-24 코드 추가입력사항 시작----------------------------------------------------------
Const E1_card_co_cd			 = 113					 '카드사 
Const E1_card_mem_no		 = 114					 '가맹점번호 
Const E1_pay_meth_pur		 = 115					 '결재방법(구매)
Const E1_pay_type_pur		 = 116					 '입출금유형(구매)
Const E1_pay_dur_pur		 = 117					 '결재기간(구매)
Const E1_bank_cd			 = 118					 '은행 
Const E1_bank_acct_no		 = 119					 '계좌번호 
Const E1_rgst_dt			= 120						'사업자번호 적용일 

Const E1_bank_info			= 121						'사업자번호 적용일 

'12-24 코드 추가입력사항 종료----------------------------------------------------------
Const E1_ind_type_nm         = 121                   '[업종명]
Const E1_ind_class_nm        = 122                   '[업태명]
Const E1_bp_group_nm         = 123                   '[거래처분류명]
Const E1_b_country_nm        = 124                   '[국가명]
Const E1_b_province          = 125                   '[지방명]
Const E1_trans_meth_nm       = 126                   '[운송방법명]
Const E1_deal_type_nm        = 127                   '[판매유형명]
Const E1_bp_grade_nm         = 128                   '[업체평가등급명]
Const E1_s_credit_limit      = 129                   '[여신관리그룹명]
Const E1_b_sales_grp_nm      = 130                   '[영업그룹명]
Const E1_b_to_grp_nm         = 131                   '[수금그룹명]
Const E1_b_pur_grp_nm        = 132                   '[구매그룹명]
Const E1_vat_type_nm         = 133                   '[부가세유형명]
Const E1_pay_meth_nm         = 134			         '[결재방법(영업)]
Const E1_pay_type_nm         = 135                   '[입출금유형명]
Const E1_tax_area_nm         = 136                   '[세금신고사업장명]
Const E1_b_zip_code          = 137                   '[--우편번호]
Const E1_b_pur_org           = 138                   '[--구매조직코드]                 
Const E1_b_pur_org_nm        = 139                   '[--구매조직명] 
Const E1_vat_inc_flag_nm     = 140                   '[--부과세구분명] 
'12-24 네임 추가입력사항 시작----------------------------------------------------------
Const E1_card_co_cd_nm		 = 141					 '[카드사명]
Const E1_pay_meth_pur_nm	 = 142					 '[결재방법명(구매)]
Const E1_pay_type_pur_nm	 = 143					 '[입출금유형명(구매)]
Const E1_bank_cd_nm			 = 144					 '[은행명]
'12-24 네임 추가입력사항 종료----------------------------------------------------------

strMode = Request("txtMode")	

Select Case strMode

Case CStr(UID_M0001)														'☜: 현재 조회/Prev/Next 요청을 받음 
    Err.Clear                                                               '☜: Protect system from crashing
   
    Select Case Request("txtPrevNext")
	Case "PREV"
		imp_command = "PREVQUERY"
	Case "NEXT"
		imp_command = "NEXTQUERY"
	Case Else 
		imp_command = "QUERY"
	End Select
      
    imp_biz_partner_cd = Trim(Request("txtConBp_cd"))     
    
    Set iS1C141 = Server.CreateObject("PB5CS41.cLookupBizPartnerSvr")    
    
    If CheckSYSTEMError(Err,True) = True Then
       Response.End
       'Exit Sub
    End If     
	
	Call iS1C141.B_LOOKUP_BIZ_PARTNER_SVR(gStrGlobalCollection, imp_command, imp_biz_partner_cd, E1_b_biz_partner)           									 
	
	IF cStr(Err.Description) = "B_MESSAGE" & Chr(11) & "126100" Then
       Response.Write "<Script language=vbs>  " & vbCr   
       Response.Write " With Parent           " & vbCr
       Response.Write "   .frm1.txtConBp_nm.value = """ & "" & """" & vbCr    
       Response.Write "   .frm1.chkBpTypeT.disabled = True " & vbCr    
       Response.Write " End With      " & vbCr															    	
       Response.Write "</Script>      " & vbCr    
    End If
									 									 
    If CheckSYSTEMError(Err,True) = True Then
       Set iS1C141 = Nothing		                                                 '☜: Unload Comproxy DLL
       Response.End
       'Exit Sub
    End If      
    
    Set iS1C141 = Nothing   
    imp_command = "" 
    
    '*** 거래처 은행 정보 >>air ***
	bankInfo	= fnBankInfo(E1_b_biz_partner(E1_bp_cd))
'Call ServerMesgBox(bankInfo , vbInformation, I_MKSCRIPT)  
	'-----------------------
	'Result data display area
	'----------------------- 
%>
<Script Language=vbscript>

	'-------------------------------------------------------------------------------
	' fnConvRemark(ByVal strRemark)
	'-------------------------------------------------------------------------------
	Function fnConvRemark(ByVal strRemark)
		Dim strTemp
		strTemp = Replace(strRemark, Chr(11), vbCrLf)
		fnConvRemark = strTemp
	End Function 
	
	

	With parent.frm1

		.txtConBp_cd.value		= "<%=ConvSPChars(E1_b_biz_partner(E1_bp_cd))%>"				
		.txtConBp_nm.value		= "<%=ConvSPChars(E1_b_biz_partner(E1_bp_nm))%>"				
		
		'---------------------------TAB1----------------------------------------------------------------
		.txtBp_cd.value			= "<%=ConvSPChars(E1_b_biz_partner(E1_bp_cd))%>"				
		.txtOwn_Rgst_N.value	= "<%=ConvSPChars(E1_b_biz_partner(E1_bp_rgst_no))%>"				
		.txtBp_full_nm.value	= "<%=ConvSPChars(E1_b_biz_partner(E1_bp_full_nm))%>"	
		.txtBp_Type.value		= "<%=ConvSPChars(E1_b_biz_partner(E1_bp_type))%>"		
		Call parent.chkQueryValue()
		.txtBp_nm.value			= "<%=ConvSPChars(E1_b_biz_partner(E1_bp_nm))%>"				
		.txtBp_eng_nm.value		= "<%=ConvSPChars(E1_b_biz_partner(E1_bp_eng_nm))%>"
		.txtBp_alias_nm.value	= "<%=ConvSPChars(E1_b_biz_partner(E1_bp_alias_nm))%>"								
		.txtRepre_nm.value		= "<%=ConvSPChars(E1_b_biz_partner(E1_repre_nm))%>"		
		.txtRepre_Rgst.value	= "<%=ConvSPChars(E1_b_biz_partner(E1_repre_rgst_no))%>"				
		.txtInd_Type.value		= "<%=ConvSPChars(E1_b_biz_partner(E1_ind_type))%>"				
		.txtInd_TypeNm.value	= "<%=ConvSPChars(E1_b_biz_partner(E1_ind_type_nm))%>"				
		.txtInd_Class.value		= "<%=ConvSPChars(E1_b_biz_partner(E1_ind_class))%>"				
		.txtInd_ClassNm.value	= "<%=ConvSPChars(E1_b_biz_partner(E1_ind_class_nm))%>"							
		.txtInd_AREA.value		= "<%=ConvSPChars(E1_b_biz_partner(E1_EXT2_CD))%>"				    '2008-04-16 5:35오후 :: hanc
		.txtInd_AREANM.value		= "<%=ConvSPChars(E1_b_biz_partner(E1_EXT3_CD))%>"				    '2008-04-16 5:35오후 :: hanc
		
		If "<%=E1_b_biz_partner(E1_usage_flag)%>" = .rdoUsage_flag1.value Then			
			.rdoUsage_flag1.checked = True
		ElseIf "<%=E1_b_biz_partner(E1_usage_flag)%>" = .rdoUsage_flag2.value Then
			.rdoUsage_flag2.checked = True
		End If
		
		If "<%=E1_b_biz_partner(E1_in_out)%>" = .rdoIn_out1.value Then			
			.rdoIn_out1.checked = True
		ElseIf "<%=E1_b_biz_partner(E1_in_out)%>" = .rdoIn_out2.value Then
			.rdoIn_out2.checked = True
		End If

		'---------------------------TAB2---------------------------------------------------------------
		.txtBp_Group.value		= "<%=ConvSPChars(E1_b_biz_partner(E1_bp_group))%>"
		.txtBp_Group_Nm.value	= "<%=ConvSPChars(E1_b_biz_partner(E1_bp_group_nm))%>"		
		.txtContry_cd.value		= "<%=ConvSPChars(E1_b_biz_partner(E1_contry_cd))%>"		
		.txtCountry_nm.value	= "<%=ConvSPChars(E1_b_biz_partner(E1_b_country_nm))%>"				
		.txtProvince_cd.value	= "<%=ConvSPChars(E1_b_biz_partner(E1_province_cd))%>"				
		.txtProvince_nm.value	= "<%=ConvSPChars(E1_b_biz_partner(E1_b_province))%>"				
		.txtZIP_cd.value		= "<%=ConvSPChars(E1_b_biz_partner(E1_zip_cd))%>"				
		.txtADDR1.value			= "<%=ConvSPChars(E1_b_biz_partner(E1_addr1))%>"		
		.txtADDR2.value			= "<%=ConvSPChars(E1_b_biz_partner(E1_addr2))%>"				
		.txtADDR1_Eng.value		= "<%=ConvSPChars(E1_b_biz_partner(E1_addr1_eng))%>"				
		.txtADDR2_Eng.value		= "<%=ConvSPChars(E1_b_biz_partner(E1_addr2_eng))%>"
		.txtADDR3_Eng.value		= "<%=ConvSPChars(E1_b_biz_partner(E1_addr3_eng))%>"		
		.txtTel_No1.value		= "<%=ConvSPChars(E1_b_biz_partner(E1_tel_no1))%>"				
		.txtTel_No2.value		= "<%=ConvSPChars(E1_b_biz_partner(E1_tel_no2))%>"				
		.txtFax_No.value		= "<%=ConvSPChars(E1_b_biz_partner(E1_fax_no))%>"					
		.txtHome_Url.value		= "<%=ConvSPChars(E1_b_biz_partner(E1_home_url))%>"
		.txtEmp_Cnt.Text		= "<%=UNINumClientFormat(E1_b_biz_partner(E1_emp_cnt),0,0)%>"				
		.txtSale_Amt.Text		= "<%=UNINumClientFormat(E1_b_biz_partner(E1_sale_amt),ggAmtOfMoney.DecPoint,0)%>"				
		.txtCapital_Amt.Text	= "<%=UNINumClientFormat(E1_b_biz_partner(E1_capital_amt),ggAmtOfMoney.DecPoint,0)%>"				
		.txtFnd_DT.Text		    = "<%=UNIDateClientFormat(E1_b_biz_partner(E1_fnd_dt))%>"				

		'---------------------------TAB3--------------------------------------------------------------
		.txtTrans_Meth.value	= "<%=ConvSPChars(E1_b_biz_partner(E1_trans_meth))%>"		
		.txtTrans_Meth_nm.value	= "<%=ConvSPChars(E1_b_biz_partner(E1_trans_meth_nm))%>"				
		.txtTrans_LT.Text		= "<%=UNINumClientFormat(E1_b_biz_partner(E1_trans_lt),0,0)%>"				
		.txtDeal_Type.value		= "<%=ConvSPChars(E1_b_biz_partner(E1_deal_type))%>"				
		.txtDeal_Type_nm.value	= "<%=ConvSPChars(E1_b_biz_partner(E1_deal_type_nm))%>"				
		.txtTrade_Rgst.value	= "<%=ConvSPChars(E1_b_biz_partner(E1_trade_rgst_no))%>"
		.txtClearance_ID.value	= "<%=ConvSPChars(E1_b_biz_partner(E1_clearance_id))%>"				
		.txtComm_Rate.Text		= "<%=UNINumClientFormat(E1_b_biz_partner(E1_comm_rate),ggExchRate.DecPoint,0)%>"				
		.txtBp_Grade.value		= "<%=ConvSPChars(E1_b_biz_partner(E1_bp_grade))%>"				
		.txtBp_Grade_nm.value	= "<%=ConvSPChars(E1_b_biz_partner(E1_bp_grade_nm))%>"					
		.txtBp_prsn_Nm.value	= "<%=ConvSPChars(E1_b_biz_partner(E1_bp_prsn_nm))%>"
		.txtBp_contact_Pt.value	= "<%=ConvSPChars(E1_b_biz_partner(E1_bp_contact_pt))%>"	
		.txtCredit_grp.value	= "<%=ConvSPChars(E1_b_biz_partner(E1_credit_grp))%>"				
		.txtCredit_grp_Nm.value	= "<%=ConvSPChars(E1_b_biz_partner(E1_s_credit_limit))%>"	
		.txtCreditRotDt.text	= "<%=ConvSPChars(E1_b_biz_partner(E1_credit_rot_day))%>"	
		.txtBiz_Grp.value		= "<%=ConvSPChars(E1_b_biz_partner(E1_biz_grp))%>"				
		.txtBiz_Grp_Nm.value	= "<%=ConvSPChars(E1_b_biz_partner(E1_b_sales_grp_nm))%>"	
		.txtTo_Grp.value		= "<%=ConvSPChars(E1_b_biz_partner(E1_to_grp))%>"				
		.txtTo_Grp_Nm.value		= "<%=ConvSPChars(E1_b_biz_partner(E1_b_to_grp_nm))%>"	
		.txtPur_Grp.value		= "<%=ConvSPChars(E1_b_biz_partner(E1_pur_grp))%>"				
		.txtPur_Grp_Nm.value	= "<%=ConvSPChars(E1_b_biz_partner(E1_b_pur_grp_nm))%>"			
		
		If UCase(Trim("<%=E1_b_biz_partner(E1_gr_insp_type)%>")) = "A" Then
			.rdoSoldInspectA.checked = True
		Else
			.rdoSoldInspectB.checked = True
		End If						

		If "<%=E1_b_biz_partner(E1_vat_inc_flag)%>" = .rdoVATinc_1.value Then	
			.rdoVATinc_1.checked = True														'-----별도 
		ElseIf "<%=E1_b_biz_partner(E1_vat_inc_flag)%>" = .rdoVATinc_2.value Then
			.rdoVATinc_2.checked = True														'-----포함 
		End IF

		If "<%=E1_b_biz_partner(E1_credit_mgmt_flag)%>" = "Y" Then	
			.rdoCredit_Y.checked = True										                '-----관리 
			.txtRadioCredit.value = .rdoCredit_Y.value 
		ElseIf "<%=E1_b_biz_partner(E1_credit_mgmt_flag)%>" = "N" Then
			.rdoCredit_N.checked = True										                '-----미관리 
			.txtRadioCredit.value = .rdoCredit_N.value 
		End If

		'---------------------------TAB4--------------------------------------------------------------------
		.txtCurrency.value			= "<%=ConvSPChars(E1_b_biz_partner(E1_currency))%>"	
		.txtvat_Type.value			= "<%=ConvSPChars(E1_b_biz_partner(E1_vat_type))%>"		
		.txtvat_Type_nm.value		= "<%=ConvSPChars(E1_b_biz_partner(E1_vat_type_nm))%>"	
		.txtPay_meth.value			= "<%=ConvSPChars(E1_b_biz_partner(E1_pay_meth))%>"		
		.txtPay_meth_nm.value		= "<%=ConvSPChars(E1_b_biz_partner(E1_pay_meth_nm))%>"	
		.txtPay_meth_Pur.value		= "<%=ConvSPChars(E1_b_biz_partner(E1_pay_meth_pur))%>"		
		.txtPay_meth_Pur_nm.value	= "<%=ConvSPChars(E1_b_biz_partner(E1_pay_meth_pur_nm))%>"			
		.txtvat_Rate.Text			= "<%=UNINumClientFormat(E1_b_biz_partner(E1_vat_rate),ggExchRate.DecPoint,0)%>"	
		.txtPay_dur.Text			= "<%=UNINumClientFormat(E1_b_biz_partner(E1_pay_dur),0,0)%>"
		.txtPay_dur_Pur.Text		= "<%=UNINumClientFormat(E1_b_biz_partner(E1_pay_dur_pur),0,0)%>"	
		.txtPay_type.value			= "<%=ConvSPChars(E1_b_biz_partner(E1_pay_type))%>"				
		.txtPay_type_Nm.value		= "<%=ConvSPChars(E1_b_biz_partner(E1_pay_type_nm))%>"	
		.txtPay_type_Pur.value		= "<%=ConvSPChars(E1_b_biz_partner(E1_pay_type_pur))%>"				
		.txtPay_type_Pur_Nm.value	= "<%=ConvSPChars(E1_b_biz_partner(E1_pay_type_pur_nm))%>"										
		.txtTaxBizAreaCd.value		= "<%=ConvSPChars(E1_b_biz_partner(E1_tax_biz_area))%>"		
		.txtTaxBizAreaNm.value		= "<%=ConvSPChars(E1_b_biz_partner(E1_tax_area_nm))%>"			
		.txtCash_Rate.Text			= "<%=UNINumClientFormat(E1_b_biz_partner(E1_cash_rate),ggExchRate.DecPoint,0)%>"	
		.txtPay_Month.value			= "<%=ConvSPChars(E1_b_biz_partner(E1_pay_month))%>"		
		.txtPay_day.Text			= "<%=ConvSPChars(E1_b_biz_partner(E1_pay_day))%>"
		.txtClose_day1.value		= "<%=ConvSPChars(E1_b_biz_partner(E1_close_day1))%>"
		.txtPay_terms_txt.value		= "<%=ConvSPChars(E1_b_biz_partner(E1_pay_terms_txt))%>"			
		.txtCardCoCd.value			= "<%=ConvSPChars(E1_b_biz_partner(E1_card_co_cd))%>"
		.txtCardCoCdNm.value		= "<%=ConvSPChars(E1_b_biz_partner(E1_card_co_cd_nm))%>"
		.txtCardMemNo.value			= "<%=ConvSPChars(E1_b_biz_partner(E1_card_mem_no))%>"		
		.txtBankCo.value			= "<%=ConvSPChars(E1_b_biz_partner(E1_bank_cd))%>"
		.txtBankCoNm.value			= "<%=ConvSPChars(E1_b_biz_partner(E1_bank_cd_nm))%>"
		.txtBankAcctNo.value		= "<%=ConvSPChars(E1_b_biz_partner(E1_bank_acct_no))%>"
		.txtOwn_Rgst_Dt.Text		= "<%=UNIDateClientFormat(E1_b_biz_partner(E1_rgst_dt))%>"

		If "<%=E1_b_biz_partner(E1_vat_calc_type)%>" = .rdoVATcalc_Y.value Then	            '부가세적용기준 
			.rdoVATcalc_Y.checked = True									                    '----------개별 
		ElseIf "<%=E1_b_biz_partner(E1_vat_calc_type)%>" = .rdoVATcalc_N.value Then
			.rdoVATcalc_N.checked = True									                    '----------통합 
		End IF

		If "<%=E1_b_biz_partner(E1_deposit_flag)%>" = .rdoReservePrice_Y.value Then	        '적립금적용기준 
			.rdoReservePrice_Y.checked = True									                '----------적용 
		ElseIf "<%=E1_b_biz_partner(E1_deposit_flag)%>" = .rdoReservePrice_N.value Then
			.rdoReservePrice_N.checked = True									                '----------미적용 
		End IF

        '2008-03-25 8:55오후 :: hanc
		If "<%=E1_b_biz_partner(E1_ext1_cd)%>" = .rdoReserveMove_Y.value Then	        '매출대상/이동 
			.rdoReserveMove_Y.checked = True									                
		ElseIf "<%=E1_b_biz_partner(E1_ext1_cd)%>" = .rdoReserveMove_N.value Then
			.rdoReserveMove_N.checked = True									                
		End IF
		

		
		.txtBankInfo.value			= fnConvRemark("<%=ConvSPChars(bankInfo)%>")'"<%=bankInfo%>"

		parent.DbQueryOk														

	End With
</Script>
<%

Case CStr(UID_M0002)																'☜: 저장 요청을 받음 
     
    Err.Clear																		'☜: Protect system from crashing

	lgIntFlgMode = CInt(Request("txtFlgMode"))										'☜: 저장시 Create/Update 판별 
	
	Redim I1_b_biz_partner(E1_bank_info)
		

	'---------------------------TAB1----------------------------------------------------------------
    I1_b_biz_partner(E1_bp_cd)         = UCase(Trim(Request("txtBp_cd")))
    I1_b_biz_partner(E1_bp_rgst_no)    = Trim(Request("txtOwn_Rgst_N"))	
    I1_b_biz_partner(E1_bp_full_nm)    = Trim(Request("txtBp_full_nm"))	
    I1_b_biz_partner(E1_bp_type)       = Trim(Request("txtBp_Type"))	                    
    I1_b_biz_partner(E1_usage_flag)    = Trim(Request("txtRadioFlag"))		   	    	                
    I1_b_biz_partner(E1_bp_nm)         = Trim(Request("txtBp_nm"))	                    
    I1_b_biz_partner(E1_bp_eng_nm)     = Trim(Request("txtBp_eng_nm"))
	I1_b_biz_partner(E1_repre_nm)      = Trim(Request("txtRepre_nm"))
    I1_b_biz_partner(E1_repre_rgst_no) = UCase(Trim(Request("txtRepre_Rgst")))			                
    I1_b_biz_partner(E1_ind_type)      = Trim(Request("txtInd_Type"))'업종	                    
    I1_b_biz_partner(E1_ind_class)     = Trim(Request("txtInd_Class"))'업태		   				
	I1_b_biz_partner(E1_in_out)		   = Trim(Request("txtRadioInOut"))'사내외구분	

    I1_b_biz_partner(E1_ext1_cd) = Trim(Request("txtRadioDepositMove"))    '2008-03-25 8:53오후 :: hanc
    I1_b_biz_partner(E1_ext2_cd) = Trim(Request("txtInd_AREA"))    '2008-04-16 5:31오후 :: hanc
    I1_b_biz_partner(E1_ext3_cd) = Trim(Request("txtInd_AREANM"))    '2008-04-16 5:31오후 :: hanc

    
    '반도체는 15자리 MEM은 30자리 ('2008-05-26 12:50오후 :: hanc)
    IF I1_b_biz_partner(E1_ext2_cd) = "1" THEN
        I1_b_biz_partner(E1_bp_alias_nm)   = LEFT(Trim(Request("txtBp_alias_nm")), 15)
    ELSE
        I1_b_biz_partner(E1_bp_alias_nm)   = LEFT(Trim(Request("txtBp_alias_nm")), 30)
    END IF

		 
	'---------------------------TAB2----------------------------------------------------------------
    I1_b_biz_partner(E1_bp_group)      = UCase(Trim(Request("txtBp_Group")))		    					
	I1_b_biz_partner(E1_contry_cd)     = UCase(Trim(Request("txtContry_cd")))
    I1_b_biz_partner(E1_zip_cd)        = Trim(Request("txtZIP_cd"))	                    
    I1_b_biz_partner(E1_province_cd)   = UCase(Trim(Request("txtProvince_cd")))			                
	I1_b_biz_partner(E1_addr1)         = Trim(Request("txtADDR1"))
	I1_b_biz_partner(E1_addr2)         = Trim(Request("txtADDR2"))
    I1_b_biz_partner(E1_addr1_eng)     = Trim(Request("txtADDR1_Eng"))			                
    I1_b_biz_partner(E1_addr2_eng)     = Trim(Request("txtADDR2_Eng"))
    I1_b_biz_partner(E1_addr3_eng)     = Trim(Request("txtADDR3_Eng"))    
    I1_b_biz_partner(E1_tel_no1)       = Trim(Request("txtTel_No1"))	                    
    I1_b_biz_partner(E1_tel_no2)       = Trim(Request("txtTel_No2"))							
	I1_b_biz_partner(E1_fax_no)        = Trim(Request("txtFax_No"))
    I1_b_biz_partner(E1_home_url)      = Trim(Request("txtHome_Url"))	
    I1_b_biz_partner(E1_fnd_dt)        = UNIConvDate(Request("txtFnd_DT"))	                    
	If Len(Trim(Request("txtEmp_Cnt"))) Then I1_b_biz_partner(E1_emp_cnt) = UNIConvNum(Trim(Request("txtEmp_Cnt")),0)
	If Len(Trim(Request("txtSale_Amt"))) Then I1_b_biz_partner(E1_sale_amt) = UNIConvNum(Trim(Request("txtSale_Amt")),0)
	If Len(Trim(Request("txtCapital_Amt"))) Then I1_b_biz_partner(E1_capital_amt) = UNIConvNum(Trim(Request("txtCapital_Amt")),0)

	'---------------------------TAB3----------------------------------------------------------------
    I1_b_biz_partner(E1_trans_meth) = UCase(Trim(Request("txtTrans_Meth")))	                    
    If Len(Trim(Request("txtTrans_LT"))) Then I1_b_biz_partner(E1_trans_lt) = Trim(Request("txtTrans_LT"))			                
	I1_b_biz_partner(E1_deal_type) = UCase(Trim(Request("txtDeal_Type")))
	I1_b_biz_partner(E1_vat_inc_flag) = Trim(Request("txtRadioVATinc"))
	I1_b_biz_partner(E1_trade_rgst_no) = Trim(Request("txtTrade_Rgst"))
	I1_b_biz_partner(E1_clearance_id) = Trim(Request("txtClearance_ID"))
	If Len(Trim(Request("txtComm_Rate"))) Then I1_b_biz_partner(E1_comm_rate) = UNIConvNum(Request("txtComm_Rate"),0)
	I1_b_biz_partner(E1_credit_mgmt_flag) = Trim(Request("txtRadioCredit"))
	I1_b_biz_partner(E1_credit_grp) = Trim(Request("txtCredit_grp"))
    I1_b_biz_partner(E1_bp_prsn_nm) = Trim(Request("txtBp_prsn_Nm"))	                    
    I1_b_biz_partner(E1_bp_contact_pt) = Trim(Request("txtBp_contact_Pt"))		
	If Len(Trim(Request("txtCreditRotDt"))) Then I1_b_biz_partner(E1_credit_rot_day) = Trim(Request("txtCreditRotDt"))
    I1_b_biz_partner(E1_gr_insp_type) = Trim(Request("txtRadioSoldInspect"))

    I1_b_biz_partner(E1_bp_grade) = UCase(Trim(Request("txtBp_Grade")))			                
    I1_b_biz_partner(E1_biz_grp) = UCase(Trim(Request("txtBiz_Grp")))
    I1_b_biz_partner(E1_to_grp) = UCase(Trim(Request("txtTo_Grp")))
    I1_b_biz_partner(E1_pur_grp) = UCase(Trim(Request("txtPur_Grp")))
    
	'---------------------------TAB4----------------------------------------------------------------
	I1_b_biz_partner(E1_currency) = UCase(Trim(Request("txtCurrency")))		
	I1_b_biz_partner(E1_vat_type) = UCase(Trim(Request("txtvat_Type")))	
	I1_b_biz_partner(E1_vat_rate) = UNIConvNum(Request("txtvat_Rate"),0)
	I1_b_biz_partner(E1_tax_biz_area) = Trim(Request("txtTaxBizAreaCd"))
	I1_b_biz_partner(E1_cash_rate) = UNIConvNum(Request("txtCash_Rate"),0)    
    I1_b_biz_partner(E1_vat_calc_type) = Trim(Request("txtRadioVATcalc"))  
    I1_b_biz_partner(E1_deposit_flag) = Trim(Request("txtRadioDepositPrice"))
    I1_b_biz_partner(E1_pay_type) = UCase(Trim(Request("txtPay_type")))
    I1_b_biz_partner(E1_pay_meth) = UCase(Trim(Request("txtPay_meth")))	
    If Len(Trim(Request("txtPay_dur"))) Then I1_b_biz_partner(E1_pay_dur) = UNIConvNum(Request("txtPay_dur"),0)
    If Len(Trim(Request("txtPay_day"))) Then I1_b_biz_partner(E1_pay_day) = Trim(Request("txtPay_day"))
    If Len(Trim(Request("txtPay_Month"))) Then I1_b_biz_partner(E1_pay_month) = Trim(Request("txtPay_Month"))
    If Len(Trim(Request("txtClose_day1"))) Then I1_b_biz_partner(E1_close_day1) = Trim(Request("txtClose_day1"))    
    I1_b_biz_partner(E1_pay_terms_txt) = Trim(Request("txtPay_terms_txt"))     
'12-24 추가사항 --------------------------------------------------------------------------------------------------
	I1_b_biz_partner(E1_card_co_cd) = Trim(Request("txtCardCoCd"))
	I1_b_biz_partner(E1_card_mem_no) = Trim(Request("txtCardMemNo"))
	I1_b_biz_partner(E1_pay_meth_pur) = Trim(Request("txtPay_meth_Pur"))	  
	I1_b_biz_partner(E1_pay_type_pur) = Trim(Request("txtPay_type_Pur"))
	

	If Len(Trim(Request("txtPay_dur_pur"))) Then I1_b_biz_partner(E1_pay_dur_pur) = UNIConvNum(Request("txtPay_dur_Pur"),0)
	I1_b_biz_partner(E1_bank_cd) = Trim(Request("txtBankCo"))
	I1_b_biz_partner(E1_bank_acct_no) = Trim(Request("txtBankAcctNo"))
	I1_b_biz_partner(E1_rgst_dt) = UNIConvDate(Request("txtOwn_Rgst_DT"))

	I1_b_biz_partner(E1_bank_info) = replace(Trim(Request("txtBankInfo")),chr(13)&chr(10), chr(11))

    
	
'-----------------------------------------------------------------------------------------------------------------		
    '=======  MA에서 값을 받지 않는 numeric 데이타 리스트 ======================================================
    I1_b_biz_partner(E1_close_day2) = 0
    I1_b_biz_partner(E1_close_day3) = 0   
    I1_b_biz_partner(E1_ext1_qty) = 0   
    I1_b_biz_partner(E1_ext2_qty) = 0   
    I1_b_biz_partner(E1_ext3_qty) = 0   
    I1_b_biz_partner(E1_ext1_amt) = 0   
    I1_b_biz_partner(E1_ext2_amt) = 0   
    I1_b_biz_partner(E1_ext3_amt) = 0   
    I1_b_biz_partner(E1_pay_month2) = 0  
    I1_b_biz_partner(E1_pay_month3) = 0
    I1_b_biz_partner(E1_close_day1_sales) = 0
    I1_b_biz_partner(E1_pay_month1_sales) = 0    
    I1_b_biz_partner(E1_close_day1_sales) = 0
    I1_b_biz_partner(E1_pay_month1_sales) = 0
    I1_b_biz_partner(E1_close_day2_sales) = 0
    I1_b_biz_partner(E1_pay_month2_sales) = 0    
    I1_b_biz_partner(E1_close_day3_sales) = 0
    I1_b_biz_partner(E1_pay_month3_sales) = 0
    	

    '===========================================================================================================

    
    If lgIntFlgMode = OPMD_CMODE Then
		imp_command = "CREATE"
    ElseIf lgIntFlgMode = OPMD_UMODE Then
		imp_command = "UPDATE"
    End If

	
    Set iS1C140 = Server.CreateObject("PB5CS40_KO441.cBMaintBizPartnerSvr")
    
    If CheckSYSTEMError(Err,True) = True Then
        Response.End
       'Exit Sub
    End If 
	
	pvCB = "F"

    Call iS1C140.B_MAINT_BIZ_PARTNER_SVR(pvCB, gStrGlobalCollection, imp_command, I1_b_biz_partner, prICustomXML)
            
    If CheckSYSTEMError(Err,True) = True Then
       Set iS1C141 = Nothing		                                                 '☜: Unload Comproxy DLL
       Response.End
    End If     

    '*** 거래처은행정보 저장 >>air ***
' 	bankInfo	= Trim(Request("txtBankInfo"))
 '   Call fnSetBankInfo(UCase(Trim(Request("txtBp_cd"))), bankInfo)    
	
%>
<Script Language=vbscript>
	Call parent.DbSaveOk()
</Script>
<%	

Case CStr(UID_M0003)														'☜: 삭제 요청 

    Err.Clear                                                               '☜: Protect system from crashing   
    
'2008-04-07 10:10오후 :: hanc::BEGIN
'2008-04-07 10:10오후 :: hanc::END    
    
     
    '-----------------------
    'Data manipulate area
    '-----------------------
    imp_command = "DELETE"
    imp_BpCd = Trim(Request("txtBp_cd"))
    
    Set iS1C140 = Server.CreateObject("PB5CS40.cBMaintBizPartnerSvr")

    If CheckSYSTEMError(Err,True) = True Then
       Response.End
    End If   
    
    pvCB = "F"    
    
    Call iS1C140.B_MAINT_BIZ_PARTNER_SVR(pvCB, gStrGlobalCollection, imp_command, imp_BpCd, prICustomXML)
    
    If CheckSYSTEMError(Err,True) = True Then
       Set iS1C140 = Nothing		                                                 '☜: Unload Comproxy DLL
       Response.End
    End If 
    
    Set iS1C140 = Nothing

	'-----------------------
	'Result data display area
	'----------------------- 
%>
<Script Language=vbscript>
	Call parent.DbDeleteOk()
</Script>
<%																					

End Select

'==============================================================================
' 사용자 정의 서버 함수 
'==============================================================================

'++++++++++++++++++++++++
' 거래처은행 정보 >>air
'++++++++++++++++++++++++
Function fnBankInfo(iVal_E1_b_biz_partner_bp_cd)
    
	lgStrSQL = ""
	'lgStrSQL = lgStrSQL & " SELECT	ISNULL(BANK_INFO,'') "
	lgStrSQL = lgStrSQL & " SELECT	REPLACE(ISNULL(BANK_INFO,''), CHAR(13) + CHAR(10), CHAR(11)) "
	lgStrSQL = lgStrSQL & " FROM	B_BIZ_PARTNER(NOLOCK) "
	lgStrSQL = lgStrSQL & " WHERE	BP_CD = " & FilterVar(iVal_E1_b_biz_partner_bp_cd, "''", "S")    
    
    'Call ServerMesgBox(lgStrSQL , vbInformation, I_MKSCRIPT)	
    Call SubOpenDB(lgObjConn)
	
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
		fnBankInfo = False
		Exit Function
	Else
		fnBankInfo = fnGetString(lgObjRs(0))
	End If    

	Call SubCloseDB(lgObjConn) 

End Function  

'++++++++++++++++++++++++++++
' 거래처은행 정보 저장 >>air
'++++++++++++++++++++++++++++
Sub fnSetBankInfo(iVal_E1_b_biz_partner_bp_cd, iVal_bank_info)
	Dim lgBankInfo
	
	lgErrorStatus = ""    


	lgBankInfo = fnSetString(iVal_bank_info)

    Call SubOpenDB(lgObjConn)
    
	lgStrSQL = ""
	lgStrSQL = lgStrSQL & " UPDATE B_BIZ_PARTNER SET BANK_INFO = '" & lgBankInfo & "' "
	lgStrSQL = lgStrSQL & " WHERE	BP_CD = " & FilterVar(iVal_E1_b_biz_partner_bp_cd, "''", "S")    

	'Call ServerMesgBox(lgStrSQL , vbInformation, I_MKSCRIPT)
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords

    If CheckSYSTEMError(Err,True) = True Then
       lgErrorStatus   = "YES"
       ObjectContext.SetAbort
    End If
    
    Call SubCloseDB(lgObjConn)

End Sub  


'++++++++++++++++++++++++++++++++++++
' 문자열변환
'++++++++++++++++++++++++++++++++++++
Function fnGetString(iStr)


	If not IsNull(iStr) And iStr <> "" Then
		iStr = Replace(iStr, "''", "'")
		iStr = replace(iStr,"&lt;","<")
		iStr = replace(iStr,"&gt;",">")
		iStr = replace(iStr,"&amp;","&")
		'iStr = replace(iStr,"<br>", chr(13)&chr(10))
		fnGetString = iStr
	End if
	
End Function



'++++++++++++++++++++++++++++++++++++
' 문자열변환
'++++++++++++++++++++++++++++++++++++
Function fnSetString(iStr)


	If not IsNull(iStr) And iStr <> "" Then
		iStr = Replace(iStr, "&", "&amp;")		
		iStr = Replace(iStr, "'", "''")
		iStr = Replace(iStr, "<", "&lt;")
		iStr = Replace(iStr, ">", "&gt;")
		'iStr = replace(iStr,chr(13)&chr(10), chr(11))	
					
		fnSetString = iStr
	End if
'Call ServerMesgBox(bankInfo , vbInformation, I_MKSCRIPT) 	
End Function


    
%>
<SCRIPT LANGUAGE=VBSCRIPT RUNAT=SERVER>
</SCRIPT>
