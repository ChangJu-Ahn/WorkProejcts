<% Option Explicit %>
<%
'**********************************************************************************************
'*  1. Module Name          : ���� 
'*  2. Function Name        : ���ϰ��� 
'*  3. Program ID           : S4112MA1
'*  4. Program Name         : ���ϳ������ 
'*  5. Program Desc         :
'*  6. Comproxy List        : S41111MaintDnHdrSvr, S41121MaintDnDtlSvr, S41115PostGoodsIssueSvr
'*							  S14113ChkDnCreditLimitSvr, S14114ChkGiCreditLimitSvr			
'*  7. Modified date(First) : 2000/04/20
'*  8. Modified date(Last)  : 2001/12/19
'*  9. Modifier (First)     : Cho song hyon
'* 10. Modifier (Last)      : Cho song hyon
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            -2000/04/20 : 3rd ȭ�� layout & ASP Coding
'*                            -2000/08/11 : 4th ȭ�� layout
'*                            -2001/12/19 : Date ǥ������ 
'**********************************************************************************************

'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComASP/LoadInfTb19029.asp" -->

<%	
On Error Resume Next												'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*", "NOCOOKIE", "MB")
Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")

Call HideStatusWnd

Dim pS5G118
Dim pS5G128
Dim pS5G531
Dim strMode		
Dim iStrNextKey							' ���� �� 
Dim lgStrPrevKey						' ���� �� 
Dim LngMaxRow							' ���� �׸����� �ִ�Row
Dim LngRow
Dim GroupCount															'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim BalanceAmt
'Dim CheckCnt

strMode = Request("txtMode")												'�� : ���� ���¸� ���� 

Select Case strMode

Case CStr(UID_M0001)														'��: ���� ��ȸ/Prev/Next ��û�� ���� 
    '==================== ��ȸ��� Index ==================
	Const S5G211_RS_DN_NO = 0                ' ���Ϲ�ȣ 
	Const S5G211_RS_SHIP_TO_PARTY = 1        ' ��ǰó 
	Const S5G211_RS_SHIP_TO_PARTY_NM = 2     ' ��ǰó�� 
	Const S5G211_RS_MOV_TYPE = 7             ' �������� 
	Const S5G211_RS_MOV_TYPE_NM = 8          ' �������¸� 
	Const S5G211_RS_DLVY_DT = 9				 ' ������ 
	Const S5G211_RS_PROMISE_DT = 10          ' ������� 
	Const S5G211_RS_ACTUAL_GI_DT = 11        ' ����� 
	Const S5G211_RS_TRANS_METH = 15          ' ��۹�� 
	Const S5G211_RS_TRANS_METH_NM = 16       ' ��۹���� 
	Const S5G211_RS_GOODS_MV_NO = 17         ' ���ҹ�ȣ(����ȣ)
	Const S5G211_RS_CI_FLAG = 18             ' ����ʿ俩�� 
	Const S5G211_RS_POST_FLAG = 19           ' ���ó������ 
	Const S5G211_RS_SO_TYPE = 20             ' �������� 
	Const S5G211_RS_SO_TYPE_NM = 21          ' �������¸� 
	Const S5G211_RS_SO_NO = 22               ' ���ֹ�ȣ 
	Const S5G211_RS_VAT_FLAG = 38            ' ���ݰ�꼭���� ���û������� 
	Const S5G211_RS_AR_FLAG = 39             ' �������� ���û������� 
	Const S5G211_RS_PLANT_CD = 72            ' ���� 
	Const S5G211_RS_PLANT_NM = 73            ' ����� 
	Const S5G211_RS_INV_MGR = 74             ' ����� 
	Const S5G211_RS_INV_MGR_NM = 75          ' ������ڸ� 
	Const S5G211_RS_RET_ITEM_FLAG = 77       ' ��ǰ���� 
	Const S5G211_RS_REL_BILL_FLAG = 78       ' ���⿩�� 
	Const S5G211_RS_EXPORT_FLAG = 79         ' ���⿩�� 

    '-----------------------
    ' ��������� �о�´�.
    '-----------------------
    Dim iStrDnNo
	Dim iObjS5G2111
	Dim iArrSDnHdr
	
	iStrDnNo = Trim(Request("txtConDnNo"))			' ���Ϲ�ȣ 
	
    Set iObjS5G2111 = Server.CreateObject("PS5G211.cLookUpSDnHdr")
	
	If CheckSYSTEMError(Err,True) = True Then
       Response.Write "<Script Language=vbscript>"			& vbCr
	   Response.Write "parent.frm1.txtConDnNo.focus"		& vbCr    
	   Response.Write "</Script>"							& vbCr	
       Response.End
    End If  
    
    iArrSDnHdr = iObjS5G2111.LookUp(gStrGlobalCollection, iStrDnNo, "N")

    If CheckSYSTEMError(Err,True) = True Then
		Set iObjS5G2111 = Nothing		                                                 '��: Unload Comproxy DLL
		Response.Write "<Script Language=vbscript>"			& vbCr
		Response.Write "Call parent.DbQueryNotFound"		& vbCr    
		Response.Write "</Script>"							& vbCr	
		Response.End																		'��: Process End
    End If  
    
    Set iObjS5G2111 = Nothing

	'-----------------------
	' ��������� ������ ǥ���Ѵ�.
	'----------------------- 		
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent.frm1" & vbCr
	Response.Write ".txtHDnNo.value				= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_DN_NO))				& """" & vbcr	' ���Ϲ�ȣ 

	Response.Write ".txtSoNo.value				= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_SO_NO))				& """" & vbcr	' ���ֹ�ȣ 
	Response.Write ".txtShipToParty.value		= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_SHIP_TO_PARTY))		& """" & vbcr	' ��ǰó 
	Response.Write ".txtShipToPartyNm.value		= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_SHIP_TO_PARTY_NM))		& """" & vbcr	' ��ǰó�� 
	Response.Write ".txtDnType.value			= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_MOV_TYPE))				& """" & vbcr	' ����Ÿ�� 
	Response.Write ".txtDnTypeNm.value			= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_MOV_TYPE_NM))			& """" & vbcr	' ����Ÿ�Ը� 
	Response.Write ".txtPlannedGIDt.value		= """ & UNIDateClientFormat(iArrSDnHdr(S5G211_RS_PROMISE_DT))	& """" & vbcr	' ������� 
	Response.Write ".txtDlvyDt.value			= """ & UNIDateClientFormat(iArrSDnHdr(S5G211_RS_DLVY_DT))		& """" & vbcr	' ������ 
	Response.Write ".txtGINo.value		= """ & Trim(ConvSPChars(iArrSDnHdr(S5G211_RS_GOODS_MV_NO)))	& """" & vbcr			' ����ȣ 
	'--���������--
	If Trim(iArrSDnHdr(S5G211_RS_GOODS_MV_NO)) <> "" Then
		Response.Write ".txtActualGIDt.Text			= """ & UNIDateClientFormat(iArrSDnHdr(S5G211_RS_ACTUAL_GI_DT))	& """" & vbcr
	End If

	' ����ä�� 
	If UCase(iArrSDnHdr(S5G211_RS_AR_FLAG)) = "Y" Then
		Response.Write ".chkARflag.checked = True "			& vbcr
		Response.Write "parent.lblArFlag.disabled = False " & vbcr
	Else
		Response.Write ".chkARflag.checked = False "		& vbcr
		Response.Write "parent.lblArFlag.disabled = True "	& vbcr
	End If

	' ���ݰ�꼭 
	If UCase(iArrSDnHdr(S5G211_RS_VAT_FLAG)) = "Y" Then
		Response.Write ".chkVatFlag.checked = True "		 & vbcr
		Response.Write "parent.lblVatFlag.disabled = False " & vbcr
	Else
		Response.Write ".chkVatFlag.checked = False "		& vbcr
		Response.Write "parent.lblVatFlag.disabled = True "	& vbcr
	End If

	Response.Write ".txtSoType.value	= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_SO_TYPE))			& """" & vbcr	' ����Ÿ�� 
	Response.Write ".txtSoTypeNm.value	= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_SO_TYPE_NM))		& """" & vbcr	' ����Ÿ�Ը� 
	Response.Write ".txtPlantCd.value	= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_PLANT_CD))			& """" & vbcr	' ���� 
	Response.Write ".txtPlantNm.value	= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_PLANT_NM))			& """" & vbcr	' ����� 
	Response.Write ".txtInvMgr.value	= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_INV_MGR))			& """" & vbcr	' ������� 
	Response.Write ".txtInvMgrNm.value	= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_INV_MGR_NM))		& """" & vbcr	' ������ڸ� 

	Response.Write ".txtHRetFlag.value	= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_RET_ITEM_FLAG))	& """" & vbcr	' ��ǰ���� 
	Response.Write ".txtRetBillFlag.value = """ & ConvSPChars(iArrSDnHdr(S5G211_RS_REL_BILL_FLAG))	& """" & vbcr	' ���⿩�� 
	Response.Write ".txtExportFlag.value  = """ & ConvSPChars(iArrSDnHdr(S5G211_RS_EXPORT_FLAG))	& """" & vbcr	' ���⿩�� 
	Response.Write "End With"          & vbCr
    Response.Write "</Script>"         & vbCr															'��: ��ȸ�� ���� 

	'==================== �������� ��. ======================

	'==================== ���ϳ��� ó�� ======================
	
	Dim i1_s_dn_dtl, i2_s_dn_hdr, e1_s_dn_dtl, eg1_exp_grp
	
	Const S424_I1_dn_seq = 0    '[CONVERSION INFORMATION]  View Name : imp_next s_dn_dtl

    Const S424_I2_dn_no = 0    '[CONVERSION INFORMATION]  View Name : imp s_dn_hdr
    Const S424_I2_except_dn_flag = 1
    
    Const C_SHEETMAXROWS_D  = 100
    
    Redim i1_s_dn_dtl(S424_I1_dn_seq)
    Redim i2_s_dn_hdr(S424_I2_except_dn_flag)

	' Output Index
	Const EA_b_item_item_cd2 = 0
	Const EA_b_item_item_nm2 = 1
	Const EA_b_item_item_acct2 = 2
	Const EA_b_item_spec2 = 3
	Const EA_b_plant_plant_cd2 = 4
	Const EA_b_plant_plant_nm2 = 5
	Const EA_b_storage_location_sl_cd2 = 6
	Const EA_b_storage_location_sl_nm2 = 7
	Const EA_s_dn_dtl_dn_seq2 = 8
	Const EA_s_dn_dtl_req_qty2 = 9
	Const EA_s_dn_dtl_req_bonus_qty2 = 10
	Const EA_s_dn_dtl_gi_qty2 = 11
	Const EA_s_dn_dtl_gi_bonus_qty2 = 12
	Const EA_s_dn_dtl_gi_unit2 = 13
	Const EA_s_dn_dtl_post_gi_flag2 = 14
	Const EA_s_dn_dtl_tol_more_qty2 = 15
	Const EA_s_dn_dtl_tol_less_qty2 = 16
	Const EA_s_dn_dtl_lot_no2 = 17
	Const EA_s_dn_dtl_lot_seq2 = 18
	Const EA_s_dn_dtl_cc_qty2 = 19
	Const EA_s_dn_dtl_remark2 = 20
	Const EA_s_dn_dtl_tracking_no2 = 21
	Const EA_s_dn_dtl_gi_amt_loc2 = 22
	Const EA_s_dn_dtl_qm_flag2 = 23
	Const EA_s_dn_dtl_vat_amt_loc2 = 24
	Const EA_s_dn_dtl_vat_amt2 = 25
	Const EA_s_dn_dtl_gi_amt2 = 26
	Const EA_s_dn_dtl_ext1_qty2 = 27
	Const EA_s_dn_dtl_ext2_qty2 = 28
	Const EA_s_dn_dtl_ext1_amt2 = 29
	Const EA_s_dn_dtl_ext2_amt2 = 30
	Const EA_s_dn_dtl_ext1_cd2 = 31
	Const EA_s_dn_dtl_ext2_cd2 = 32
	Const EA_s_dn_dtl_ext3_qty2 = 33
	Const EA_s_dn_dtl_ext3_amt2 = 34
	Const EA_s_dn_dtl_ext3_cd2 = 35
	Const EA_s_dn_dtl_ret_type2 = 36
	Const EA_s_dn_dtl_deposit_amt2 = 37
	Const EA_s_dn_dtl_price2 = 38
	Const EA_s_dn_dtl_vat_rate2 = 39
	Const EA_s_dn_dtl_vat_inc_flag2 = 40
	Const EA_s_dn_dtl_vat_type2 = 41
	Const EA_s_dn_dtl_dn_no2 = 42
	Const EA_s_dn_dtl_lc_no2 = 43
	Const EA_s_dn_dtl_lc_seq2 = 44
	Const EA_s_dn_dtl_so_seq2 = 45
	Const EA_s_dn_dtl_so_no2 = 46
	Const EA_s_so_hdr_lc_flag2 = 47
	Const EA_s_so_hdr_ret_item_flag2 = 48
	Const EA_s_dn_dtl_so_schd_no2 = 49
	Const EA_b_item_by_plant_lot_flg2 = 50
	Const EA_b_item_by_plant_ship_inspec_flg2 = 51
	Const EA_i_good_on_hand_qty2 = 52
	const EA_b_minor_ret_type_nm2 = 53      
	Const EA_s_dn_dtl_carton_no2 = 54
	Const EA_s_dn_dtl_rel_bill_no2 = 55
	Const EA_s_dn_dtl_rel_bill_cnt2 = 56
	Const EA_b_item_basic_unit = 57
	Const EA_s_dn_dtl_dn_req_no2 = 58
	Const EA_s_dn_dtl_dn_req_seq2 = 59
	Const EA_s_dn_dtl_OUT_NO_KO441 = 60
	Const EA_s_dn_dtl_TRANS_TIME_KO441 = 61
	Const EA_s_dn_dtl_OUT_TYPE_KO441 = 62
	Const EA_s_dn_dtl_CREATE_TYPE_KO441 = 63
	Const EA_s_dn_dtl_REF_GUBUN_KO441 = 64

    '2008-06-16 11:02���� :: hanc
	Const EA_s_dn_dtl_pgm_name_KO441 = 65
	Const EA_s_dn_dtl_pgm_price_KO441 = 66
		
	'-----------------------
    ' ���ϳ����� �о�´�.
    '-----------------------
    i2_s_dn_hdr(S424_I2_dn_no) = Trim(Request("txtConDnNo"))
    lgStrPrevKey = Trim(Request("lgStrPrevkey"))
    
    If Trim(Request("lgStrPrevKey")) = "" Then
		i1_s_dn_dtl(S424_I1_dn_seq) = 0
    Else
		i1_s_dn_dtl(S424_I1_dn_seq) = cdbl(Request("lgStrPrevKey"))
	End if	    

    Set pS5G128 = Server.CreateObject("pS5G128_KO441.cSListDnDtl")
        
    If CheckSYSTEMError(Err,True) = True Then
		Response.Write "<Script Language=vbscript>" & vbCr
		Response.Write "Call parent.DbQueryOk" & vbCr
		Response.Write "</Script>" & vbCr
		Response.End																				'��: Process End
    End If
   
	call pS5G128.S_LIST_DN_DTL(gStrGlobalCollection, C_SHEETMAXROWS_D, i1_s_dn_dtl, i2_s_dn_hdr, eg1_exp_grp)

	If CheckSYSTEMError(Err,True) = True Then
		Set pS5G128 = Nothing					
		Response.Write "<Script Language=vbscript>" & vbCr
		Response.Write "Call parent.DbQueryOk" & vbCr
		Response.Write "</Script>" & vbCr
		Response.End																				'��: Process End
	End If
	
	Set pS5G128 = Nothing

	Dim iLngLastRow, iLngSheetMaxRows
	' Client(MA)�� ���� ��ȸ�� ������ Row
	iLngLastRow = CLng(Request("txtLastRow")) + 1
	
	' Set Next key
	If Ubound(EG1_EXP_GRP,1) = C_SHEETMAXROWS_D Then
		'������ 
		iStrNextKey = EG1_EXP_GRP(C_SHEETMAXROWS_D, EA_s_dn_dtl_dn_seq2)
		iLngSheetMaxRows  = C_SHEETMAXROWS_D - 1
	Else
		iStrNextKey = ""
		iLngSheetMaxRows = Ubound(EG1_EXP_GRP,1)
	End If

	ReDim iArrCols(55)						' Column �� 
	Redim iArrRows(iLngSheetMaxRows)		' ��ȸ�� Row ����ŭ �迭 ������ 

	iArrCols(0) = ""
	iArrCols(11) = ""		' Lot Popup 
	iArrCols(23) = ""		' �˻籸�� 
	iArrCols(25) = ""		' ���� Popup 
	iArrCols(27) = ""		' â�� Popup 
	
	'-----------------------
	' ���ϳ����� ������ ǥ���Ѵ�.
	'----------------------- 
	For LngRow = 0 To UBound(EG1_EXP_GRP,1)	
   		iArrCols(1)  = ConvSPChars(EG1_EXP_GRP(LngRow,EA_b_item_item_cd2))				'ǰ���ڵ� 
   		iArrCols(2)  = ConvSPChars(EG1_EXP_GRP(LngRow,EA_b_item_item_nm2))				'ǰ��� 
   		iArrCols(3)  = ConvSPChars(EG1_EXP_GRP(LngRow,EA_b_item_spec2))					'ǰ��԰� 
   		iArrCols(4)  = ConvSPChars(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_tracking_no2))		'���� 
   		iArrCols(5)  = ConvSPChars(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_gi_unit2))			'���� 
		iArrCols(6)  = UNINumClientFormat(EG1_exp_grp(LngRow, EA_s_dn_dtl_req_qty2), ggQty.DecPoint, 0)			' ����û���� 
		iArrCols(7)  = UNINumClientFormat(EG1_exp_grp(LngRow, EA_s_dn_dtl_req_bonus_qty2), ggQty.DecPoint, 0)	' ����û������ 
		iArrCols(8)  = UNINumClientFormat(EG1_exp_grp(LngRow, EA_s_dn_dtl_gi_qty2), ggQty.DecPoint, 0)			' Picking ���� 
		iArrCols(9)  = UNINumClientFormat(EG1_exp_grp(LngRow, EA_s_dn_dtl_gi_bonus_qty2), ggQty.DecPoint, 0)	' Picking������ 
   		iArrCols(10) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_lot_no2))										'Lot��ȣ 
   		iArrCols(12) = UNINumClientFormat(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_lot_seq2), 0, 0)						'Lot���� 
		iArrCols(13) = UNINumClientFormat(EG1_exp_grp(LngRow, EA_i_good_on_hand_qty2), ggQty.DecPoint, 0)		' ��� 
   		iArrCols(14)  = ConvSPChars(EG1_EXP_GRP(LngRow,EA_b_item_basic_unit))			'������ 

   		iArrCols(15) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_carton_no2))									'carton no
   		iArrCols(16) = UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(LngRow,EA_s_dn_dtl_gi_amt2),gCurrency,ggAmtOfMoneyNo, "X" , "X")							' ���ݾ� 
   		iArrCols(17) = UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(LngRow,EA_s_dn_dtl_gi_amt_loc2),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo  , "X")		' �ڱ��ݾ� 
   		iArrCols(18) = UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(LngRow,EA_s_dn_dtl_deposit_amt2),gCurrency,ggAmtOfMoneyNo, "X" , "X")					' �����ݾ� 
   		iArrCols(19) = UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(LngRow,EA_s_dn_dtl_vat_amt2),gCurrency,ggAmtOfMoneyNo, gTaxRndPolicyNo  , "X")			' VAT �ݾ� 
   		iArrCols(20) = UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(LngRow,EA_s_dn_dtl_vat_amt_loc2),gCurrency,ggAmtOfMoneyNo, gTaxRndPolicyNo  , "X")		' VAT �ڱ��ݾ� 
   		iArrCols(21) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_b_item_by_plant_ship_inspec_flg2))						'�˻�ǰ���� 
   		iArrCols(22) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_qm_flag2))									'�˻籸�� 
   		iArrCols(24) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_b_plant_plant_cd2))									'���� 
   		iArrCols(26) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_b_storage_location_sl_cd2))							'â�� 
		iArrCols(28) = UNINumClientFormat(EG1_exp_grp(LngRow, EA_s_dn_dtl_tol_more_qty2), ggQty.DecPoint, 0)	'��������뷮(+)
		iArrCols(29) = UNINumClientFormat(EG1_exp_grp(LngRow, EA_s_dn_dtl_tol_less_qty2), ggQty.DecPoint, 0)	'��������뷮(-)
		iArrCols(30) = UNINumClientFormat(EG1_exp_grp(LngRow, EA_s_dn_dtl_cc_qty2), ggQty.DecPoint, 0)			'������� 
   		iArrCols(31) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_so_no2))										'���ֹ�ȣ 
   		iArrCols(32) = UNINumClientFormat(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_so_seq2), 0, 0)						'���ּ��� 
   		iArrCols(33) = UNINumClientFormat(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_so_schd_no2), 0, 0)					'����������ȣ 
   		iArrCols(34) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_lc_no2))										'L/C��ȣ 
   		iArrCols(35) = UNINumClientFormat(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_lc_seq2), 0, 0)						'L/C ���� 
   		iArrCols(36) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_ret_type2))									'��ǰ���� 
   		iArrCols(37) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_b_minor_ret_type_nm2))									'��ǰ������ 
   		iArrCols(38) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_remark2))										'��� 
   		iArrCols(39) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_s_so_hdr_ret_item_flag2))								'��ǰ���� 
   		iArrCols(40) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_b_item_by_plant_lot_flg2))								'Lot���� ��󿩺� 
   		iArrCols(41) = UNINumClientFormat(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_dn_seq2), 0, 0)						'���ϼ��� 
   		iArrCols(42) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_rel_bill_no2))
   		iArrCols(43) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_rel_bill_cnt2))
   		iArrCols(44) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_dn_req_no2))										'���Ͽ�û��ȣ 
   		iArrCols(45) = UNINumClientFormat(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_dn_req_seq2), 0, 0)						'���Ͽ�û���� 
   		iArrCols(46) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_ext2_cd2))						'�԰� Lot No. 
   		iArrCols(47) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_ext1_cd2))						'���� Lot No. 

   		iArrCols(48) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_OUT_NO_KO441))										 
   		iArrCols(49) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_TRANS_TIME_KO441))										 
   		iArrCols(50) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_OUT_TYPE_KO441))										 
   		iArrCols(51) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_CREATE_TYPE_KO441))										 
   		iArrCols(52) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_REF_GUBUN_KO441))										 

   		iArrCols(53) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_pgm_name_KO441))							    '2008-06-16 7:50���� :: hanc
		iArrCols(54) = UNINumClientFormat(EG1_exp_grp(LngRow, EA_s_dn_dtl_pgm_price_KO441), ggQty.DecPoint, 0)	'2008-06-16 7:50���� :: hanc

   		
		iArrCols(55) = iLngLastRow + LngRow
		
   		iArrRows(LngRow) = Join(iArrCols, gColSep)
	Next
		
	Response.Write "<Script language=vbs> " & vbCr   
	Response.Write "With parent " & vbCr   
    Response.Write " .ggoSpread.Source = .frm1.vspdData" & vbCr
    Response.Write " .frm1.vspdData.Redraw = False  "      & vbCr      
    Response.Write " .ggoSpread.SSShowDataByClip   """ & Join(iArrRows, gColSep & gRowSep) & gColSep & gRowSep & """ ,""F""" & vbCr
    Response.Write " .lgStrPrevKey = """ & iStrNextKey  & """" & vbCr  
	Response.Write ".frm1.vspdData.Redraw = True  "       & vbCr

    
    Response.Write " If .frm1.vspdData.MaxRows <= .VisibleRowCnt(.frm1.vspdData,NewTop) And .lgStrPrevKey <> """" Then	 " & vbCr	         
	Response.Write		" .DbQuery  " & vbCr	' GroupView ������� ȭ�� Row������ ������ ������ �ٽ� ������ 
	Response.Write " Else  " & vbCr
    Response.Write		" .DbQueryOk " & vbCr   
	Response.Write " End If	 " & vbCr
	Response.Write "End With " & vbCr   
	Response.Write "</Script> " & vbCr          

	Response.End																				'��: Process End
	
Case CStr(UID_M0002)																'��: ���� ��û�� ���� 

	Dim iErrorPosition
    
	Set pS5G531 = Server.CreateObject("pS5G531_KO441.cSDnDtlSvrForDnReqNo")

	If CheckSYSTEMError(Err,True) Then Response.End       


	Call pS5G531.Maintain(gStrGlobalCollection, "N", Request("txtHDnNo"), "N", Request("txtSpreadIns"), Request("txtSpreadUpd"), Request("txtSpreadDel"), iErrorPosition, "F")
    
    Set pS5G531 = Nothing
    
	If Trim(iErrorPosition) <> "" Then
		If CheckSYSTEMError2(Err, True, iErrorPosition & "��","","","","") = True Then
			Response.Write "<SCRIPT LANGUAGE=VBSCRIPT> " & vbCr   
			Response.Write " Call Parent.SubSetErrPos(" & iErrorPosition & ")" & vbCr
			Response.Write "</SCRIPT> "		
			Response.End 
		End If	
	Else
		If CheckSYSTEMError(Err,True) = True Then
			Response.End 
		End If
	End If
	
	Response.Write "<Script Language=vbscript>	" & vbcr
	Response.Write "parent.DbSaveOk	" & vbcr
	Response.Write "</Script>" & vbcr
	Response.End																				'��: Process End

' PICK,POST Logic
Case CStr("ARPOST")

    Err.Clear																		

	Dim pS5G115
	Dim pvCB, iCommand
	Dim i1_s_bill_hdr
	Dim i3_s_dn_hdr
	Dim iErrPosition
	
	Const S427_I1_bill_no = 0    '[CONVERSION INFORMATION]  View Name : imp s_bill_hdr
    
    Const S427_I3_dn_no = 0    '[CONVERSION INFORMATION]  View Name : imp s_dn_hdr
    Const S427_I3_actual_gi_dt = 1
    Const S427_I3_ar_flag = 2
    Const S427_I3_vat_flag = 3
    Const S427_I3_except_dn_flag = 4
    Const S427_I3_eu_flag = 5
    Const S427_I3_inv_mgr = 6
    
	Redim i1_s_bill_hdr(S427_I1_bill_no)
	Redim i3_s_dn_hdr(S427_I3_inv_mgr)
	
	LngMaxRow = CInt(Request("txtMaxRows"))											'��: �ִ� ������Ʈ�� ���� 

    '-----------------------
    'Data manipulate area
    '-----------------------
	
	i3_s_dn_hdr(S427_I3_dn_no) = Trim(Request("txtHDnNo"))
	i3_s_dn_hdr(S427_I3_actual_gi_dt) = UNIConvDate(Request("txtActualGIDt"))
	i3_s_dn_hdr(S427_I3_ar_flag) = Request("txtARFlag")
	i3_s_dn_hdr(S427_I3_vat_flag) = Request("txtVatFlag")
	i3_s_dn_hdr(S427_I3_eu_flag) = "ST"
	i3_s_dn_hdr(S427_I3_inv_mgr) = Request("txtInvMgr")
	
	If Trim(Request("txtGINo")) = "" Then
		iCommand = "POST"
	Else
		iCommand = "CANCEL"
	End If
	
	pvCB = "F"
	
	Set pS5G115 = Server.CreateObject("PS5G115_KO441.cSPostGISvr")
	
	If CheckSYSTEMError(Err,True) = True Then
       Response.End       
    End If
    
    Call pS5G115.S_POST_GOODS_ISSUE_SVR(pvCB, gStrGlobalCollection, iCommand, i1_s_bill_hdr, i3_s_dn_hdr, ,iErrPosition)

    Set pS5G115 = Nothing
        
	If Trim(iErrPosition) = "" Then	
		If CheckSYSTEMError(Err,True) = True Then	
		    Response.End		  
		End If 
	Else	
	
		If CheckSYSTEMError2(Err, True,  iErrPosition ,"","","","") = True Then
		     Response.End
		End If
	End If	

	Response.Write "<Script Language=vbscript>" & vbcr
	Response.Write "parent.DbSaveOk()" & vbcr
	Response.Write "</Script>" & vbcr
	Response.End																		'��: Process End

Case "ChkGiCreditLimit"															'��: �����ѵ� üũ 

    Err.Clear																		

	Dim objPS3G113	
	Dim iArrData
	Dim iDblOverLimitAmt
	
	Redim iArrData(1)
    
    iArrData(0) = "NG"
    iArrData(1) = Trim(Request("txtConDnNo"))

	
	Set objPS3G113 = Server.CreateObject("PS3G113.cChkCreditLimit")
	
	If CheckSYSTEMError(Err,True) = True Then
       Response.End       
    End If
    
    Call objPS3G113.ChkCreditLimitSvr(gStrGlobalCollection, iArrData, iDblOverLimitAmt)
    
	Set objPS3G113 = Nothing
	
	If Err.number = 0 Then
		Response.Write("<SCRIPT LANGUAGE=VBSCRIPT>" & vbCr)
		Response.Write("Call parent.BatchButton(2)" & vbCr)
		Response.Write("</SCRIPT>" & vbCr)

    Else
		' �����ѵ��� �ʰ��� ���(���ó��)
		If InStr(1, err.Description, "B_MESSAGE" & Chr(11) & "201929") > 0 Then
			Response.Write("<SCRIPT LANGUAGE=VBSCRIPT>" & vbCr)
			Response.Write("Dim iReturnVal" & vbCr)
			' �����ѵ��� %1 %2 ��ŭ �ʰ��Ͽ����ϴ�. �����Ͻðڽ��ϱ�?
			Response.Write("iReturnVal = parent.DisplayMsgBox(""201929"", parent.parent.VB_YES_NO, parent.parent.gCurrency, """ & UNINumClientFormat(iDblOverLimitAmt, ggAmtOfMoney.DecPoint, 0) & """)" & vbCr )
			Response.Write("If iReturnVal = vbYes Then Call parent.BatchButton(2)" & vbCr)
			Response.Write("</SCRIPT>" & vbCr)
		ElseIf InStr(1, err.Description, "B_MESSAGE" & Chr(11) & "201722") > 0 Then
			Response.Write("<SCRIPT LANGUAGE=VBSCRIPT>" & vbCr)
			'�����ѵ��� %1 %2 ��ŭ �ʰ��Ͽ����ϴ�.
			Response.Write("Call parent.DisplayMsgBox(""201722"", ""X"", parent.parent.gCurrency, """ & UNINumClientFormat(iDblOverLimitAmt, ggAmtOfMoney.DecPoint, 0) & """)" & vbCr)
			Response.Write("</SCRIPT>" & vbCr)
		Else
			Call CheckSYSTEMError(Err,True)
		End If
	End If
	
	Response.End
Case "GetIssueFromMES"															'��: �����ѵ� üũ 
	On Error Resume Next												'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

    Err.Clear																		

	Dim PS5G116_KO441	
	Dim EG1_export_group
	Dim istrData
	
	Set PS5G116_KO441 = Server.CreateObject("PS5G116_KO441.cILstIssueFromMES")
	
	If CheckSYSTEMError(Err,True) = True Then
       Response.End       
    End If
    
    Call PS5G116_KO441.I_LIST_ISSUE_FROM_MES(gStrGlobalCollection, ""&Trim(Request("txtSpread")), EG1_export_group)
    
	Set PS5G116_KO441 = Nothing

	If CheckSYSTEMError(Err,True) = True Then
       Response.End       
  End If
	
	Const C_OUT_NO 						= 0
	Const C_BP_CD 						= 1
	Const C_BP_NM 						= 2
	Const C_ITEM_CD 					= 3
	Const C_ITEM_NM 					= 4
	Const C_SPEC 							= 5
	Const C_PLANT_CD 					= 6
	Const C_OUT_TYPE 					= 7
	Const C_OUT_TYPE_NM 			= 8
	Const C_GOOD_ON_HAND_QTY 	= 9
	Const C_GI_QTY 						= 10
	Const C_GI_UNIT 					= 11
	Const C_LOT_NO 						= 12
	Const C_LOT_SUB_NO 				= 13
	Const C_ACTUAL_GI_DT 			= 14
	Const C_CUST_LOT_NO 			= 15
	Const C_SO_NO 						= 16
	Const C_SO_SEQ 						= 17
	Const C_SL_CD 						= 18
	Const C_SL_NM 						= 19
	Const C_TRANS_TIME  			= 20
	Const C_CREATE_TYPE				= 21

	If isArray(EG1_export_group) Then 
		For LngRow = 0 To Ubound(EG1_export_group,1)
				istrData = Chr(11) &  ConvSPChars(EG1_export_group(LngRow,C_ITEM_CD))
				istrData = istrData & Chr(11) &  ConvSPChars(EG1_export_group(LngRow,C_ITEM_NM))
				istrData = istrData & Chr(11) &  ConvSPChars(EG1_export_group(LngRow,C_SPEC))
				istrData = istrData & Chr(11) &  "*"
				istrData = istrData & Chr(11) &  ConvSPChars(EG1_export_group(LngRow,C_GI_UNIT))
				istrData = istrData & Chr(11) &  UNINumClientFormat(EG1_export_group(LngRow, C_GI_QTY), ggQty.DecPoint, 0)
				istrData = istrData & Chr(11) &  "0"
				istrData = istrData & Chr(11) &  UNINumClientFormat(EG1_export_group(LngRow, C_GI_QTY), ggQty.DecPoint, 0)
				istrData = istrData & Chr(11) &  "0"	'
				istrData = istrData & Chr(11) &  ConvSPChars(EG1_export_group(LngRow,C_LOT_NO)) 'Lot No
				istrData = istrData & Chr(11)					'Log Popup
				istrData = istrData & Chr(11) &  "0" 'Lot Seq
				istrData = istrData & Chr(11) &  UNINumClientFormat(EG1_export_group(LngRow, C_GOOD_ON_HAND_QTY), ggQty.DecPoint, 0) '��������
				istrData = istrData & Chr(11) &  ConvSPChars(EG1_export_group(LngRow,C_GI_UNIT)) '������
				istrData = istrData & Chr(11)			'Carton No
				istrData = istrData & Chr(11) &  "0" '���ݾ�
				istrData = istrData & Chr(11) &  "0" '����ڱݾ�
				istrData = istrData & Chr(11) &  "0" '�����ݾ�
				istrData = istrData & Chr(11) &  "0" '�ΰ����ݾ�
				istrData = istrData & Chr(11) &  "0" '�ΰ����ڱ��ݾ�
				istrData = istrData & Chr(11)	&  "N"		'�˻�ǰ����
				istrData = istrData & Chr(11)'�˻籸��
				istrData = istrData & Chr(11)'�˻��ȣ�˾�
				istrData = istrData & Chr(11) &  ConvSPChars(EG1_export_group(LngRow,C_PLANT_CD)) '�����ڵ�
				istrData = istrData & Chr(11)'�����ڵ��˾�
				istrData = istrData & Chr(11) &  ConvSPChars(EG1_export_group(LngRow,C_SL_CD)) 'â���ڵ�
				istrData = istrData & Chr(11)'â���ڵ��˾�
				istrData = istrData & Chr(11) &  "0" '��������뷮(+)
				istrData = istrData & Chr(11) &  "0" '��������뷮(-)
				istrData = istrData & Chr(11) &  "0" '�������
				istrData = istrData & Chr(11) &  ConvSPChars(EG1_export_group(LngRow,C_SO_NO)) '���ֹ�ȣ
				istrData = istrData & Chr(11) &  ConvSPChars(EG1_export_group(LngRow,C_SO_SEQ)) '���ּ���
				istrData = istrData & Chr(11) &  "1" '��ǰ����
				istrData = istrData & Chr(11)'L/C��ȣ
				istrData = istrData & Chr(11) &  "0" 'L/C����
				istrData = istrData & Chr(11)'��ǰ����
				istrData = istrData & Chr(11)'��ǰ������
				istrData = istrData & Chr(11)'���
				istrData = istrData & Chr(11) & "N" 'Lot��ǰ����
				istrData = istrData & Chr(11) & "Y" 'Lot��������
				istrData = istrData & Chr(11) & "0" '���ϼ���
				istrData = istrData & Chr(11) 'Ref. Bill No.
				istrData = istrData & Chr(11) & "0" 'Ref. Bill No. Seq.
				istrData = istrData & Chr(11) 'D/N Req No.
				istrData = istrData & Chr(11) & "0" 'D/N Req No. Seq.								
				istrData = istrData & Chr(11) &  ConvSPChars(EG1_export_group(LngRow,C_CUST_LOT_NO)) '�� Lot No.
				istrData = istrData & Chr(11) &  ConvSPChars(EG1_export_group(LngRow,C_OUT_NO)) 
				istrData = istrData & Chr(11) &  ConvSPChars(EG1_export_group(LngRow,C_TRANS_TIME)) 
				istrData = istrData & Chr(11) 
				istrData = istrData & Chr(11) &  ConvSPChars(EG1_export_group(LngRow,C_CREATE_TYPE)) 
				istrData = istrData & Chr(11) & "2"			
				istrData = istrData & Chr(11) &  LngMaxRow + LngRow 	 & Chr(12)		
	
			Response.Write "<Script language=vbs> " & vbCr   
			Response.Write "With parent " & vbCr   
			Response.Write " .frm1.vspdData.Redraw = False  " & vbCr      
			Response.Write " .ggoSpread.Source = .frm1.vspdData	" & vbCr
			Response.Write " .ggoSpread.SSShowDataByClip """ & istrData & """,""F"""& vbCr	
			Response.Write " .frm1.vspdData.Redraw = True  " & vbCr
			Response.Write " .GetIssueFromMESOk()"& vbCr	
			Response.Write "End With " & vbCr   
			Response.Write "</Script> "																				         & vbCr          
		
		Next
	End If

		
	Response.End
End Select

'===========================================
' ����� ���� ���� �Լ� 
'===========================================
%>
<SCRIPT LANGUAGE=VBSCRIPT RUNAT=SERVER>
</SCRIPT>
