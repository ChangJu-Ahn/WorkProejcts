<%
'********************************************************************************************************
'*  1. Module Name          : 영업관리																	*
'*  2. Function Name        :																			*
'*  3. Program ID           : s3212ma2.asp																*
'*  4. Program Name         : Local L/C 내역등록														*
'*  5. Program Desc         : Local L/C 내역등록														*
'*  6. Comproxy List        :																			*
'*  7. Modified date(First) : 2000/04/20																*
'*  8. Modified date(Last)  : 2001/12/17																*
'*  9. Modifier (First)     : Kim Hyungsuk																*
'* 10. Modifier (Last)      : Son bum Yeol																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/20 : 화면 design												*
'*							  2. 2000/03/23 : Coding Start												*
'********************************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%													

On Error Resume Next

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*", "NOCOOKIE", "MB")   
Call LoadBNumericFormatB("I","*","NOCOOKIE","MB")
Call HideStatusWnd      

Dim PS4G128																	' Master L/C Detail 조회용 Object
Dim PS4G119																	' Master L/C Header 조회용 Object
Dim PS4G121
Dim strMode																		'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim LngRow
Dim intGroupCount
Dim prGroupView
Dim lgCurrency

Dim I1_s_lc_hdr
ReDim I1_s_lc_hdr(1)

Dim E7_b_sales_grp
Dim E10_b_biz_partner
Dim E14_b_minor
Dim E15_b_minor
Dim E26_s_lc_hdr

Const S357_I1_lc_no = 0    
Const S357_I1_lc_kind = 1

Const S357_E7_sales_grp_nm = 0    'exp b_sales_grp
Const S357_E7_sales_grp = 1
Const S357_E10_bp_nm = 0   'exp_applicant b_biz_partner
Const S357_E10_bp_cd = 1
Const S357_E14_minor_nm = 0    '[CONVERSION INFORMATION]  View Name : exp_incoterms_nm b_minor
Const S357_E15_minor_nm = 0    '[CONVERSION INFORMATION]  View Name : exp_pay_meth_nm b_minor

Const S357_E26_lc_no = 0    '[CONVERSION INFORMATION]  View Name : exp s_lc_hdr
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

Dim I1_s_lc_hdr_dtl

Dim I2_s_lc_dtl

Dim E1_s_lc_dtl
Dim EG1_exp_grp 
Dim E2_s_lc_dtl

Const S367_E1_lc_seq = 0    '[CONVERSION INFORMATION]  View Name : exp_next s_lc_dtl

Const S367_EG1_E1_so_qty = 0    '[CONVERSION INFORMATION]  View Name : exp_item_remain s_so_dtl
Const S367_EG1_E2_dn_seq = 1    '[CONVERSION INFORMATION]  View Name : exp_item s_dn_dtl
Const S367_EG1_E3_dn_no = 2    '[CONVERSION INFORMATION]  View Name : exp_item s_dn_hdr
Const S367_EG1_E4_item_cd = 3    '[CONVERSION INFORMATION]  View Name : exp_item b_item
Const S367_EG1_E4_item_nm = 4
Const S367_EG1_E5_so_no = 5    '[CONVERSION INFORMATION]  View Name : exp_item s_so_hdr
Const S367_EG1_E6_so_seq = 6    '[CONVERSION INFORMATION]  View Name : exp_item s_so_dtl
Const S367_EG1_E7_lc_seq = 7    '[CONVERSION INFORMATION]  View Name : exp_item s_lc_dtl
Const S367_EG1_E7_hs_cd = 8
Const S367_EG1_E7_lc_qty = 9
Const S367_EG1_E7_price = 10
Const S367_EG1_E7_lc_amt = 11
Const S367_EG1_E7_lc_unit = 12
Const S367_EG1_E7_over_tolerance = 13
Const S367_EG1_E7_under_tolerance = 14
Const S367_EG1_E7_dn_qty = 15
Const S367_EG1_E7_close_flag = 16
Const S367_EG1_E7_cc_qty = 17
Const S367_EG1_E7_bl_qty = 18
Const S367_EG1_E7_ext1_qty = 19
Const S367_EG1_E7_ext2_qty = 20
Const S367_EG1_E7_ext3_qty = 21
Const S367_EG1_E7_ext1_amt = 22
Const S367_EG1_E7_ext2_amt = 23
Const S367_EG1_E7_ext3_amt = 24
Const S367_EG1_E7_ext1_cd = 25
Const S367_EG1_E7_ext2_cd = 26
Const S367_EG1_E7_ext3_cd = 27
Const S367_EG1_E7_tracking_no		= 28
Const S367_EG1_E4_spec = 29

Const S367_E2_lc_amt = 0    '[CONVERSION INFORMATION]  View Name : exp_tot_amt s_lc_dtl

Const C_SHEETMAXROWS_D  = 100

strMode = Request("txtMode")													'☜ : 현재 상태를 받음 
lgStrPrevKey = Request("lgStrPrevKey")

Select Case strMode
	Case CStr(UID_M0001)														
		
	If Trim(Request("txtLCNo")) = "" Then											
		Call ServerMesgBox("조회 조건값이 비어있습니다!", vbInformation, I_MKSCRIPT)
		Response.End
	End If
		
	I1_s_lc_hdr(S357_I1_lc_no) = Trim(Request("txtLCNo"))
    I1_s_lc_hdr(S357_I1_lc_kind) = "L"
	
    Set PS4G119 = Server.CreateObject("PS4G119.cSLkLcHdrSvr")
		
	Call PS4G119.S_LOOKUP_LC_HDR_SVR(gStrGlobalCollection,"LOOKUP",I1_s_lc_hdr, _
	,,,,,,E7_b_sales_grp,,,E10_b_biz_partner,,,,E14_b_minor,E15_b_minor,,,,,,, _
    ,,,,E26_s_lc_hdr)
    

    If CheckSYSTEMError(Err,True) = True Then
		Set PS4G119 = Nothing
%>
	<Script Language=VBScript>
		parent.frm1.txtLcNo.focus
	</Script>	
<%
		Response.End
	End If  
	
	Set PS4G119 = Nothing  
	
	lgCurrency = ConvSPChars(E26_s_lc_hdr(S357_E26_cur))

%>
<Script Language=VBScript>
	With parent.frm1
		Dim strDt
		.txtCurrency.value = "<%=lgCurrency%>"
		.txtCurrency1.value = "<%=lgCurrency%>"
		parent.CurFormatNumericOCX

		.txtLCDocNo.value		=	"<%=ConvSPChars(E26_s_lc_hdr(S357_E26_lc_doc_no))%>"
		.txtLCAmendSeq.value	=	"<%=ConvSPChars(E26_s_lc_hdr(S357_E26_lc_amend_seq))%>"
		.txtSONo.value			=	"<%=ConvSPChars(E26_s_lc_hdr(S357_E26_so_no))%>"
		.txtApplicant.value		=	"<%=ConvSPChars(E10_b_biz_partner(S357_E10_bp_cd))%>"
		.txtApplicantNm.value	=	"<%=ConvSPChars(E10_b_biz_partner(S357_E10_bp_nm))%>"
		.txtSalesGroup.value	=	"<%=ConvSPChars(E7_b_sales_grp(S357_E7_sales_grp))%>"
		.txtSalesGroupNm.value	=	"<%=ConvSPChars(E7_b_sales_grp(S357_E7_sales_grp_nm))%>"
		
		.txtOpenDt.text = "<%=UNIDateClientFormat(E26_s_lc_hdr(S357_E26_open_dt))%>"
		
		.txtPayTerms.value		=	"<%=ConvSPChars(E26_s_lc_hdr(S357_E26_pay_meth))%>"
		.txtPayTermsNm.value	=	"<%=ConvSPChars(E15_b_minor(S357_E15_minor_nm))%>"
		.txtCurrency.value		=	"<%=ConvSPChars(E26_s_lc_hdr(S357_E26_cur))%>"
		.txtCurrency1.value		=	"<%=ConvSPChars(E26_s_lc_hdr(S357_E26_cur))%>"
	
		.txtDocAmt.text	= "<%=UNINumClientFormatByCurrency(E26_s_lc_hdr(S357_E26_lc_amt),lgCurrency,ggAmtOfMoneyNo)%>"
		
		.txtTotItemAmt.text	= "<%=UNINumClientFormatByCurrency(0,lgCurrency,ggAmtOfMoneyNo)%>"
		
		.txtHLCNo.value = "<%=ConvSPChars(E26_s_lc_hdr(S357_E26_lc_no))%>"
		.txtMaxSeq.value = 0

		If parent.lgIntFlgMode = parent.parent.OPMD_CMODE Then parent.CurFormatNumSprSheet 
		Call parent.LCHrdQueryOk()
	End With
</Script>
<%
	I1_s_lc_hdr_dtl = Trim(Request("txtLCNo"))

	If lgStrPrevKey <> "" Then
		I2_s_lc_dtl = lgStrPrevKey
	Else
		I2_s_lc_dtl = 0
	End If
	
	Set PS4G128 = Server.CreateObject("PS4G128.cSListLcDltSvr")
		
	Call PS4G128.S_LIST_LC_DTL_SVR(gStrGlobalCollection,C_SHEETMAXROWS_D,I1_s_lc_hdr_dtl,I2_s_lc_dtl, _
	E1_s_lc_dtl ,EG1_exp_grp ,E2_s_lc_dtl,prGroupView)
            
    If CheckSYSTEMError(Err,True) = True Then
		prGroupView = -1
		Set PS4G128 = Nothing
%>
	<Script Language=VBScript>
		parent.frm1.txtLcNo.focus
	</Script>	
<%		
		Response.End
	End If  
	
	Set PS4G128 = Nothing  

	intGroupCount = prGroupView

	If Cint(EG1_exp_grp(intGroupCount,S367_EG1_E7_lc_seq)) = Cint(E1_s_lc_dtl(S367_E1_lc_seq)) Then
		StrNextKey = ""
	Else
		StrNextKey = E1_s_lc_dtl(S367_E1_lc_seq)
	End If

%>
<Script Language=VBScript>
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp
    Dim strData

	With parent
		.frm1.txtTotItemAmt.text = "<%=UNINumClientFormat(E2_s_lc_dtl(S367_E2_lc_amt), ggAmtOfMoney.DecPoint, 0)%>"
		LngMaxRow = .frm1.vspdData.MaxRows					

<%      
  
	For LngRow = 0 To intGroupCount
		
%>
        strData = strData & Chr(11) & "<%=ConvSPChars(EG1_exp_grp(LngRow,S367_EG1_E7_lc_seq))%>"	
        strData = strData & Chr(11) & "<%=ConvSPChars(EG1_exp_grp(LngRow,S367_EG1_E4_item_cd))%>"
        strData = strData & Chr(11) & "<%=ConvSPChars(EG1_exp_grp(LngRow,S367_EG1_E4_item_nm))%>"
        strData = strData & Chr(11) & "<%=ConvSPChars(EG1_exp_grp(LngRow,S367_EG1_E4_spec ))%>"
        strData = strData & Chr(11) & "<%=ConvSPChars(EG1_exp_grp(LngRow,S367_EG1_E7_lc_unit))%>"
        If "<%=cdbl(EG1_exp_grp(LngRow,S367_EG1_E1_so_qty))%>" < 0 Then
			strData = strData & Chr(11) & "<%=UNINumClientFormat(0, ggQty.DecPoint, 0)%>"
		Else
		    strData = strData & Chr(11) & "<%=UNINumClientFormat(EG1_exp_grp(LngRow,S367_EG1_E1_so_qty), ggQty.DecPoint, 0)%>"
		End If
        strData = strData & Chr(11) & "<%=UNINumClientFormat(EG1_exp_grp(LngRow,S367_EG1_E7_lc_qty), ggQty.DecPoint, 0)%>"										'3
        strData = strData & Chr(11) & "<%=UNINumClientFormat(EG1_exp_grp(LngRow,S367_EG1_E7_price), ggUnitCost.DecPoint, 0)%>"		
        strData = strData & Chr(11) & "<%=UNINumClientFormatByCurrency(EG1_exp_grp(LngRow,S367_EG1_E7_lc_amt), lgCurrency, ggAmtOfMoneyNo)%>"
        strData = strData & Chr(11) & "<%=UNINumClientFormat(EG1_exp_grp(LngRow,S367_EG1_E7_over_tolerance), ggExchRate.DecPoint, 0)%>"	
        strData = strData & Chr(11) & "<%=UNINumClientFormat(EG1_exp_grp(LngRow,S367_EG1_E7_under_tolerance), ggExchRate.DecPoint, 0)%>"	
        
        If "<%=EG1_exp_grp(LngRow,S367_EG1_E7_close_flag)%>" = "Y" then
			strData = strData & Chr(11) & "1"
		Else
			strData = strData & Chr(11) & "0"					
		End if
        
        strData = strData & Chr(11) & "<%=ConvSPChars(EG1_exp_grp(LngRow,S367_EG1_E7_hs_cd))%>"																				'10
        strData = strData & Chr(11) & "<%=ConvSPChars(EG1_exp_grp(LngRow,S367_EG1_E5_so_no))%>"																				'11
        strData = strData & Chr(11) & "<%=ConvSPChars(EG1_exp_grp(LngRow,S367_EG1_E6_so_seq))%>"																			'13
        strData = strData & Chr(11) & "<%=UNINumClientFormat(EG1_exp_grp(LngRow,S367_EG1_E7_dn_qty), ggQty.DecPoint, 0)%>"									'14
        strData = strData & Chr(11) & "<%=UNINumClientFormat(EG1_exp_grp(LngRow,S367_EG1_E7_cc_qty), ggQty.DecPoint, 0)%>"									'15
        strData = strData & Chr(11) & "<%=UNINumClientFormat(EG1_exp_grp(LngRow,S367_EG1_E7_bl_qty), ggQty.DecPoint, 0)%>"													'16
        strData = strData & Chr(11) & "<%=ConvSPChars(EG1_exp_grp(LngRow,S367_EG1_E3_dn_no))%>"																				'12
        strData = strData & Chr(11) & "<%=ConvSPChars(EG1_exp_grp(LngRow,S367_EG1_E2_dn_seq))%>"
        strData = strData & Chr(11) & "<%=ConvSPChars(EG1_exp_grp(LngRow,S367_EG1_E7_tracking_no))%>"		
        strData = strData & Chr(11) & LngMaxRow + <%=LngRow%>																									'17
        strData = strData & Chr(11) & Chr(12)
<%
    Next
%>
		.ggoSpread.Source = .frm1.vspdData 
		.ggoSpread.SSShowDataByClip strData
		
		.frm1.vspdData.ReDraw = False
		.SetSpreadColor -1,-1
		.frm1.vspdData.ReDraw = True
		
		.lgStrPrevKey = "<%=StrNextKey%>"
		.frm1.txtHLCNo.value = "<%=ConvSPChars(Trim(Request("txtLCNo")))%>"
		.DbQueryOk
		
	End With
</Script>
<%
	
	Response.End																'☜: Process End
	
	Case CStr(UID_M0002)														'☜: 현재 Save 요청을 받음 
	If Trim(Request("txtLCNo")) = "" Then											
		Call ServerMesgBox("조회 조건값이 비어있습니다!", vbInformation, I_MKSCRIPT)
		Response.End
	End If
	
	ReDim I1_s_lc_hdr_dtl(S367_I1_lc_no)
    
	I1_s_lc_hdr_dtl(S367_I1_lc_no) = Trim(Request("txtLCNo"))
				
    Set PS4G121 = Server.CreateObject("PS4G121.cSLcDtlSvr")
    
    Call PS4G121.S_MAINT_LC_DTL_SVR(gStrGlobalCollection, _
            Request("txtSpread"),I1_s_lc_hdr_dtl,iErrorPosition)
       
    If CheckSYSTEMError2(Err,True,iErrorPosition & "행","","","","") = True Then
		Set PS4G121 = Nothing
		Response.End
	End If	
 	
    Set PS4G121 = Nothing                '☜: Unload Comproxy
%>
<Script Language=vbscript>
	With parent																	'☜: 화면 처리 ASP 를 지칭함 
		.DbSaveOk
	End With
</Script>
<%
		Response.End															'☜: Process End
	Case Else
		Response.End
End Select

%>