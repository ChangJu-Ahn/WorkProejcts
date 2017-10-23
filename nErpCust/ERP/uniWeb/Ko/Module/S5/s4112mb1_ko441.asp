<% Option Explicit %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 출하관리 
'*  3. Program ID           : S4112MA1
'*  4. Program Name         : 출하내역등록 
'*  5. Program Desc         :
'*  6. Comproxy List        : S41111MaintDnHdrSvr, S41121MaintDnDtlSvr, S41115PostGoodsIssueSvr
'*							  S14113ChkDnCreditLimitSvr, S14114ChkGiCreditLimitSvr			
'*  7. Modified date(First) : 2000/04/20
'*  8. Modified date(Last)  : 2001/12/19
'*  9. Modifier (First)     : Cho song hyon
'* 10. Modifier (Last)      : Cho song hyon
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/04/20 : 3rd 화면 layout & ASP Coding
'*                            -2000/08/11 : 4th 화면 layout
'*                            -2001/12/19 : Date 표준적용 
'**********************************************************************************************

'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComASP/LoadInfTb19029.asp" -->

<%	
On Error Resume Next												'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*", "NOCOOKIE", "MB")
Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")

Call HideStatusWnd

Dim pS5G118
Dim pS5G128
Dim pS5G531
Dim strMode		
Dim iStrNextKey							' 다음 값 
Dim lgStrPrevKey						' 이전 값 
Dim LngMaxRow							' 현재 그리드의 최대Row
Dim LngRow
Dim GroupCount															'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim BalanceAmt
'Dim CheckCnt

strMode = Request("txtMode")												'☜ : 현재 상태를 받음 

Select Case strMode

Case CStr(UID_M0001)														'☜: 현재 조회/Prev/Next 요청을 받음 
    '==================== 조회결과 Index ==================
	Const S5G211_RS_DN_NO = 0                ' 출하번호 
	Const S5G211_RS_SHIP_TO_PARTY = 1        ' 납품처 
	Const S5G211_RS_SHIP_TO_PARTY_NM = 2     ' 납품처명 
	Const S5G211_RS_MOV_TYPE = 7             ' 출하형태 
	Const S5G211_RS_MOV_TYPE_NM = 8          ' 출하형태명 
	Const S5G211_RS_DLVY_DT = 9				 ' 납기일 
	Const S5G211_RS_PROMISE_DT = 10          ' 출고예정일 
	Const S5G211_RS_ACTUAL_GI_DT = 11        ' 출고일 
	Const S5G211_RS_TRANS_METH = 15          ' 운송방법 
	Const S5G211_RS_TRANS_METH_NM = 16       ' 운송방법명 
	Const S5G211_RS_GOODS_MV_NO = 17         ' 수불번호(출고번호)
	Const S5G211_RS_CI_FLAG = 18             ' 통관필요여부 
	Const S5G211_RS_POST_FLAG = 19           ' 출고처리여부 
	Const S5G211_RS_SO_TYPE = 20             ' 수주형태 
	Const S5G211_RS_SO_TYPE_NM = 21          ' 수주형태명 
	Const S5G211_RS_SO_NO = 22               ' 수주번호 
	Const S5G211_RS_VAT_FLAG = 38            ' 세금계산서정보 동시생성여부 
	Const S5G211_RS_AR_FLAG = 39             ' 매출정보 동시생성여부 
	Const S5G211_RS_PLANT_CD = 72            ' 공장 
	Const S5G211_RS_PLANT_NM = 73            ' 공장명 
	Const S5G211_RS_INV_MGR = 74             ' 재고담당 
	Const S5G211_RS_INV_MGR_NM = 75          ' 재고담당자명 
	Const S5G211_RS_RET_ITEM_FLAG = 77       ' 반품여부 
	Const S5G211_RS_REL_BILL_FLAG = 78       ' 매출여부 
	Const S5G211_RS_EXPORT_FLAG = 79         ' 수출여부 

    '-----------------------
    ' 출하헤더를 읽어온다.
    '-----------------------
    Dim iStrDnNo
	Dim iObjS5G2111
	Dim iArrSDnHdr
	
	iStrDnNo = Trim(Request("txtConDnNo"))			' 출하번호 
	
    Set iObjS5G2111 = Server.CreateObject("PS5G211.cLookUpSDnHdr")
	
	If CheckSYSTEMError(Err,True) = True Then
       Response.Write "<Script Language=vbscript>"			& vbCr
	   Response.Write "parent.frm1.txtConDnNo.focus"		& vbCr    
	   Response.Write "</Script>"							& vbCr	
       Response.End
    End If  
    
    iArrSDnHdr = iObjS5G2111.LookUp(gStrGlobalCollection, iStrDnNo, "N")

    If CheckSYSTEMError(Err,True) = True Then
		Set iObjS5G2111 = Nothing		                                                 '☜: Unload Comproxy DLL
		Response.Write "<Script Language=vbscript>"			& vbCr
		Response.Write "Call parent.DbQueryNotFound"		& vbCr    
		Response.Write "</Script>"							& vbCr	
		Response.End																		'☜: Process End
    End If  
    
    Set iObjS5G2111 = Nothing

	'-----------------------
	' 출하헤더의 내용을 표시한다.
	'----------------------- 		
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent.frm1" & vbCr
	Response.Write ".txtHDnNo.value				= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_DN_NO))				& """" & vbcr	' 출하번호 

	Response.Write ".txtSoNo.value				= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_SO_NO))				& """" & vbcr	' 수주번호 
	Response.Write ".txtShipToParty.value		= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_SHIP_TO_PARTY))		& """" & vbcr	' 납품처 
	Response.Write ".txtShipToPartyNm.value		= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_SHIP_TO_PARTY_NM))		& """" & vbcr	' 납품처명 
	Response.Write ".txtDnType.value			= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_MOV_TYPE))				& """" & vbcr	' 출하타입 
	Response.Write ".txtDnTypeNm.value			= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_MOV_TYPE_NM))			& """" & vbcr	' 출하타입명 
	Response.Write ".txtPlannedGIDt.value		= """ & UNIDateClientFormat(iArrSDnHdr(S5G211_RS_PROMISE_DT))	& """" & vbcr	' 출고예정일 
	Response.Write ".txtDlvyDt.value			= """ & UNIDateClientFormat(iArrSDnHdr(S5G211_RS_DLVY_DT))		& """" & vbcr	' 납기일 
	Response.Write ".txtGINo.value		= """ & Trim(ConvSPChars(iArrSDnHdr(S5G211_RS_GOODS_MV_NO)))	& """" & vbcr			' 출고번호 
	'--실제출고일--
	If Trim(iArrSDnHdr(S5G211_RS_GOODS_MV_NO)) <> "" Then
		Response.Write ".txtActualGIDt.Text			= """ & UNIDateClientFormat(iArrSDnHdr(S5G211_RS_ACTUAL_GI_DT))	& """" & vbcr
	End If

	' 매출채권 
	If UCase(iArrSDnHdr(S5G211_RS_AR_FLAG)) = "Y" Then
		Response.Write ".chkARflag.checked = True "			& vbcr
		Response.Write "parent.lblArFlag.disabled = False " & vbcr
	Else
		Response.Write ".chkARflag.checked = False "		& vbcr
		Response.Write "parent.lblArFlag.disabled = True "	& vbcr
	End If

	' 세금계산서 
	If UCase(iArrSDnHdr(S5G211_RS_VAT_FLAG)) = "Y" Then
		Response.Write ".chkVatFlag.checked = True "		 & vbcr
		Response.Write "parent.lblVatFlag.disabled = False " & vbcr
	Else
		Response.Write ".chkVatFlag.checked = False "		& vbcr
		Response.Write "parent.lblVatFlag.disabled = True "	& vbcr
	End If

	Response.Write ".txtSoType.value	= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_SO_TYPE))			& """" & vbcr	' 수주타입 
	Response.Write ".txtSoTypeNm.value	= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_SO_TYPE_NM))		& """" & vbcr	' 수주타입명 
	Response.Write ".txtPlantCd.value	= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_PLANT_CD))			& """" & vbcr	' 공장 
	Response.Write ".txtPlantNm.value	= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_PLANT_NM))			& """" & vbcr	' 공장명 
	Response.Write ".txtInvMgr.value	= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_INV_MGR))			& """" & vbcr	' 재고담당자 
	Response.Write ".txtInvMgrNm.value	= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_INV_MGR_NM))		& """" & vbcr	' 재고담당자명 

	Response.Write ".txtHRetFlag.value	= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_RET_ITEM_FLAG))	& """" & vbcr	' 반품여부 
	Response.Write ".txtRetBillFlag.value = """ & ConvSPChars(iArrSDnHdr(S5G211_RS_REL_BILL_FLAG))	& """" & vbcr	' 매출여부 
	Response.Write ".txtExportFlag.value  = """ & ConvSPChars(iArrSDnHdr(S5G211_RS_EXPORT_FLAG))	& """" & vbcr	' 수출여부 
	Response.Write "End With"          & vbCr
    Response.Write "</Script>"         & vbCr															'☜: 조회가 성공 

	'==================== 출하정보 끝. ======================

	'==================== 출하내역 처리 ======================
	
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

    '2008-06-16 11:02오후 :: hanc
	Const EA_s_dn_dtl_pgm_name_KO441 = 65
	Const EA_s_dn_dtl_pgm_price_KO441 = 66
		
	'-----------------------
    ' 출하내역을 읽어온다.
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
		Response.End																				'☜: Process End
    End If
   
	call pS5G128.S_LIST_DN_DTL(gStrGlobalCollection, C_SHEETMAXROWS_D, i1_s_dn_dtl, i2_s_dn_hdr, eg1_exp_grp)

	If CheckSYSTEMError(Err,True) = True Then
		Set pS5G128 = Nothing					
		Response.Write "<Script Language=vbscript>" & vbCr
		Response.Write "Call parent.DbQueryOk" & vbCr
		Response.Write "</Script>" & vbCr
		Response.End																				'☜: Process End
	End If
	
	Set pS5G128 = Nothing

	Dim iLngLastRow, iLngSheetMaxRows
	' Client(MA)의 현재 조회된 마직막 Row
	iLngLastRow = CLng(Request("txtLastRow")) + 1
	
	' Set Next key
	If Ubound(EG1_EXP_GRP,1) = C_SHEETMAXROWS_D Then
		'출고순번 
		iStrNextKey = EG1_EXP_GRP(C_SHEETMAXROWS_D, EA_s_dn_dtl_dn_seq2)
		iLngSheetMaxRows  = C_SHEETMAXROWS_D - 1
	Else
		iStrNextKey = ""
		iLngSheetMaxRows = Ubound(EG1_EXP_GRP,1)
	End If

	ReDim iArrCols(55)						' Column 수 
	Redim iArrRows(iLngSheetMaxRows)		' 조회된 Row 수만큼 배열 재정의 

	iArrCols(0) = ""
	iArrCols(11) = ""		' Lot Popup 
	iArrCols(23) = ""		' 검사구분 
	iArrCols(25) = ""		' 공장 Popup 
	iArrCols(27) = ""		' 창고 Popup 
	
	'-----------------------
	' 출하내역의 내용을 표시한다.
	'----------------------- 
	For LngRow = 0 To UBound(EG1_EXP_GRP,1)	
   		iArrCols(1)  = ConvSPChars(EG1_EXP_GRP(LngRow,EA_b_item_item_cd2))				'품목코드 
   		iArrCols(2)  = ConvSPChars(EG1_EXP_GRP(LngRow,EA_b_item_item_nm2))				'품목명 
   		iArrCols(3)  = ConvSPChars(EG1_EXP_GRP(LngRow,EA_b_item_spec2))					'품목규격 
   		iArrCols(4)  = ConvSPChars(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_tracking_no2))		'제번 
   		iArrCols(5)  = ConvSPChars(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_gi_unit2))			'단위 
		iArrCols(6)  = UNINumClientFormat(EG1_exp_grp(LngRow, EA_s_dn_dtl_req_qty2), ggQty.DecPoint, 0)			' 출고요청수량 
		iArrCols(7)  = UNINumClientFormat(EG1_exp_grp(LngRow, EA_s_dn_dtl_req_bonus_qty2), ggQty.DecPoint, 0)	' 출고요청덤수량 
		iArrCols(8)  = UNINumClientFormat(EG1_exp_grp(LngRow, EA_s_dn_dtl_gi_qty2), ggQty.DecPoint, 0)			' Picking 수량 
		iArrCols(9)  = UNINumClientFormat(EG1_exp_grp(LngRow, EA_s_dn_dtl_gi_bonus_qty2), ggQty.DecPoint, 0)	' Picking덤수량 
   		iArrCols(10) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_lot_no2))										'Lot번호 
   		iArrCols(12) = UNINumClientFormat(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_lot_seq2), 0, 0)						'Lot순번 
		iArrCols(13) = UNINumClientFormat(EG1_exp_grp(LngRow, EA_i_good_on_hand_qty2), ggQty.DecPoint, 0)		' 재고량 
   		iArrCols(14)  = ConvSPChars(EG1_EXP_GRP(LngRow,EA_b_item_basic_unit))			'재고단위 

   		iArrCols(15) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_carton_no2))									'carton no
   		iArrCols(16) = UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(LngRow,EA_s_dn_dtl_gi_amt2),gCurrency,ggAmtOfMoneyNo, "X" , "X")							' 출고금액 
   		iArrCols(17) = UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(LngRow,EA_s_dn_dtl_gi_amt_loc2),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo  , "X")		' 자국금액 
   		iArrCols(18) = UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(LngRow,EA_s_dn_dtl_deposit_amt2),gCurrency,ggAmtOfMoneyNo, "X" , "X")					' 적립금액 
   		iArrCols(19) = UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(LngRow,EA_s_dn_dtl_vat_amt2),gCurrency,ggAmtOfMoneyNo, gTaxRndPolicyNo  , "X")			' VAT 금액 
   		iArrCols(20) = UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(LngRow,EA_s_dn_dtl_vat_amt_loc2),gCurrency,ggAmtOfMoneyNo, gTaxRndPolicyNo  , "X")		' VAT 자국금액 
   		iArrCols(21) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_b_item_by_plant_ship_inspec_flg2))						'검사품여부 
   		iArrCols(22) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_qm_flag2))									'검사구분 
   		iArrCols(24) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_b_plant_plant_cd2))									'공장 
   		iArrCols(26) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_b_storage_location_sl_cd2))							'창고 
		iArrCols(28) = UNINumClientFormat(EG1_exp_grp(LngRow, EA_s_dn_dtl_tol_more_qty2), ggQty.DecPoint, 0)	'과부족허용량(+)
		iArrCols(29) = UNINumClientFormat(EG1_exp_grp(LngRow, EA_s_dn_dtl_tol_less_qty2), ggQty.DecPoint, 0)	'과부족허용량(-)
		iArrCols(30) = UNINumClientFormat(EG1_exp_grp(LngRow, EA_s_dn_dtl_cc_qty2), ggQty.DecPoint, 0)			'통관수량 
   		iArrCols(31) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_so_no2))										'수주번호 
   		iArrCols(32) = UNINumClientFormat(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_so_seq2), 0, 0)						'수주순번 
   		iArrCols(33) = UNINumClientFormat(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_so_schd_no2), 0, 0)					'수주일정번호 
   		iArrCols(34) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_lc_no2))										'L/C번호 
   		iArrCols(35) = UNINumClientFormat(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_lc_seq2), 0, 0)						'L/C 순번 
   		iArrCols(36) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_ret_type2))									'반품유형 
   		iArrCols(37) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_b_minor_ret_type_nm2))									'반품유형명 
   		iArrCols(38) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_remark2))										'비고 
   		iArrCols(39) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_s_so_hdr_ret_item_flag2))								'반품여부 
   		iArrCols(40) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_b_item_by_plant_lot_flg2))								'Lot관리 대상여부 
   		iArrCols(41) = UNINumClientFormat(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_dn_seq2), 0, 0)						'출하순번 
   		iArrCols(42) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_rel_bill_no2))
   		iArrCols(43) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_rel_bill_cnt2))
   		iArrCols(44) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_dn_req_no2))										'출하요청번호 
   		iArrCols(45) = UNINumClientFormat(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_dn_req_seq2), 0, 0)						'출하요청순번 
   		iArrCols(46) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_ext2_cd2))						'입고 Lot No. 
   		iArrCols(47) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_ext1_cd2))						'고객사 Lot No. 

   		iArrCols(48) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_OUT_NO_KO441))										 
   		iArrCols(49) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_TRANS_TIME_KO441))										 
   		iArrCols(50) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_OUT_TYPE_KO441))										 
   		iArrCols(51) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_CREATE_TYPE_KO441))										 
   		iArrCols(52) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_REF_GUBUN_KO441))										 

   		iArrCols(53) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_pgm_name_KO441))							    '2008-06-16 7:50오후 :: hanc
		iArrCols(54) = UNINumClientFormat(EG1_exp_grp(LngRow, EA_s_dn_dtl_pgm_price_KO441), ggQty.DecPoint, 0)	'2008-06-16 7:50오후 :: hanc

   		
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
	Response.Write		" .DbQuery  " & vbCr	' GroupView 사이즈로 화면 Row수보다 쿼리가 작으면 다시 쿼리함 
	Response.Write " Else  " & vbCr
    Response.Write		" .DbQueryOk " & vbCr   
	Response.Write " End If	 " & vbCr
	Response.Write "End With " & vbCr   
	Response.Write "</Script> " & vbCr          

	Response.End																				'☜: Process End
	
Case CStr(UID_M0002)																'☜: 저장 요청을 받음 

	Dim iErrorPosition
    
	Set pS5G531 = Server.CreateObject("pS5G531_KO441.cSDnDtlSvrForDnReqNo")

	If CheckSYSTEMError(Err,True) Then Response.End       


	Call pS5G531.Maintain(gStrGlobalCollection, "N", Request("txtHDnNo"), "N", Request("txtSpreadIns"), Request("txtSpreadUpd"), Request("txtSpreadDel"), iErrorPosition, "F")
    
    Set pS5G531 = Nothing
    
	If Trim(iErrorPosition) <> "" Then
		If CheckSYSTEMError2(Err, True, iErrorPosition & "행","","","","") = True Then
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
	Response.End																				'☜: Process End

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
	
	LngMaxRow = CInt(Request("txtMaxRows"))											'☜: 최대 업데이트된 갯수 

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
	Response.End																		'☜: Process End

Case "ChkGiCreditLimit"															'☜: 여신한도 체크 

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
		' 여신한도가 초과된 경우(경고처리)
		If InStr(1, err.Description, "B_MESSAGE" & Chr(11) & "201929") > 0 Then
			Response.Write("<SCRIPT LANGUAGE=VBSCRIPT>" & vbCr)
			Response.Write("Dim iReturnVal" & vbCr)
			' 여신한도를 %1 %2 만큼 초과하였습니다. 저장하시겠습니까?
			Response.Write("iReturnVal = parent.DisplayMsgBox(""201929"", parent.parent.VB_YES_NO, parent.parent.gCurrency, """ & UNINumClientFormat(iDblOverLimitAmt, ggAmtOfMoney.DecPoint, 0) & """)" & vbCr )
			Response.Write("If iReturnVal = vbYes Then Call parent.BatchButton(2)" & vbCr)
			Response.Write("</SCRIPT>" & vbCr)
		ElseIf InStr(1, err.Description, "B_MESSAGE" & Chr(11) & "201722") > 0 Then
			Response.Write("<SCRIPT LANGUAGE=VBSCRIPT>" & vbCr)
			'여신한도를 %1 %2 만큼 초과하였습니다.
			Response.Write("Call parent.DisplayMsgBox(""201722"", ""X"", parent.parent.gCurrency, """ & UNINumClientFormat(iDblOverLimitAmt, ggAmtOfMoney.DecPoint, 0) & """)" & vbCr)
			Response.Write("</SCRIPT>" & vbCr)
		Else
			Call CheckSYSTEMError(Err,True)
		End If
	End If
	
	Response.End
Case "GetIssueFromMES"															'☜: 여신한도 체크 
	On Error Resume Next												'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

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
				istrData = istrData & Chr(11) &  UNINumClientFormat(EG1_export_group(LngRow, C_GOOD_ON_HAND_QTY), ggQty.DecPoint, 0) '현재고수량
				istrData = istrData & Chr(11) &  ConvSPChars(EG1_export_group(LngRow,C_GI_UNIT)) '재고단위
				istrData = istrData & Chr(11)			'Carton No
				istrData = istrData & Chr(11) &  "0" '출고금액
				istrData = istrData & Chr(11) &  "0" '출고자금액
				istrData = istrData & Chr(11) &  "0" '적립금액
				istrData = istrData & Chr(11) &  "0" '부가세금액
				istrData = istrData & Chr(11) &  "0" '부가세자국금액
				istrData = istrData & Chr(11)	&  "N"		'검사품여부
				istrData = istrData & Chr(11)'검사구분
				istrData = istrData & Chr(11)'검사번호팝업
				istrData = istrData & Chr(11) &  ConvSPChars(EG1_export_group(LngRow,C_PLANT_CD)) '공장코드
				istrData = istrData & Chr(11)'공장코드팝업
				istrData = istrData & Chr(11) &  ConvSPChars(EG1_export_group(LngRow,C_SL_CD)) '창고코드
				istrData = istrData & Chr(11)'창고코드팝업
				istrData = istrData & Chr(11) &  "0" '과부족허용량(+)
				istrData = istrData & Chr(11) &  "0" '과부족허용량(-)
				istrData = istrData & Chr(11) &  "0" '통관수량
				istrData = istrData & Chr(11) &  ConvSPChars(EG1_export_group(LngRow,C_SO_NO)) '수주번호
				istrData = istrData & Chr(11) &  ConvSPChars(EG1_export_group(LngRow,C_SO_SEQ)) '수주순번
				istrData = istrData & Chr(11) &  "1" '납품순번
				istrData = istrData & Chr(11)'L/C번호
				istrData = istrData & Chr(11) &  "0" 'L/C순번
				istrData = istrData & Chr(11)'반품유형
				istrData = istrData & Chr(11)'반품유형명
				istrData = istrData & Chr(11)'비고
				istrData = istrData & Chr(11) & "N" 'Lot반품여부
				istrData = istrData & Chr(11) & "Y" 'Lot관리여부
				istrData = istrData & Chr(11) & "0" '출하순번
				istrData = istrData & Chr(11) 'Ref. Bill No.
				istrData = istrData & Chr(11) & "0" 'Ref. Bill No. Seq.
				istrData = istrData & Chr(11) 'D/N Req No.
				istrData = istrData & Chr(11) & "0" 'D/N Req No. Seq.								
				istrData = istrData & Chr(11) &  ConvSPChars(EG1_export_group(LngRow,C_CUST_LOT_NO)) '고객 Lot No.
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
' 사용자 정의 서버 함수 
'===========================================
%>
<SCRIPT LANGUAGE=VBSCRIPT RUNAT=SERVER>
</SCRIPT>
