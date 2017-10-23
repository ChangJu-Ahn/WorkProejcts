<% Option Explicit %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 수주관리 
'*  3. Program ID           : S3112MA1
'*  4. Program Name         : 수주내역등록 
'*  5. Program Desc         : 
'*  6. Comproxy List        : S31121MaintSoDtlSvr, S31119LookupSoHdrSvr, S31124CreateDnBySoSvr,
'*							  S14112ChkSoCreditLimitSvr, S31122GetSoPriceSvr, B1b019LookUpItem
'*  7. Modified date(First) : 2000/04/20
'*  8. Modified date(Last)  : 2003/03/25
'*  9. Modifier (First)     : Cho song hyon
'* 10. Modifier (Last)      : Hwangseongbae
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/04/20 : 3rd 화면 layout & ASP Coding
'*                            -2000/08/11 : 4th 화면 layout
'*                            -2001/12/18 : Date 표준적용 
'**********************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComASP/LoadInfTb19029.asp" -->

<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*", "NOCOOKIE", "MB")
Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")


On Error Resume Next														'☜: 

Call HideStatusWnd

Dim iObjS5G211, iObjS5G128, iObjS5G121
Dim iStrMode		
Dim iStrNextKey							' 다음 값 
Dim lgStrPrevKey						' 이전 값 
Dim LngRow
Dim iCommand

Dim iStrDnNo, iStrCUDFlag, pvCB
Dim iArrSDnHdr

iStrMode = Request("txtMode")												'☜ : 현재 상태를 받음 

Select Case iStrMode
Case CStr(UID_M0001)														'☜: 현재 조회/Prev/Next 요청을 받음 
    '-----------------------
    ' 출하헤더를 읽어온다.
    '-----------------------
	iStrDnNo = Trim(Request("txtConDN_no"))			' 출하번호 
    lgStrPrevKey = Trim(Request("lgStrPrevkey"))
	
	' 처음 조회시만 Header 정보를 조회한다.
	If lgStrPrevKey = "" Then 
		Set iObjS5G211 = Server.CreateObject("PS5G211.cLookUpSDnHdr")
	
		If CheckSYSTEMError(Err,True) = True Then
		   Response.Write "<Script Language=vbscript>"					& vbCr
		   Response.Write "parent.frm1.txtConDN_no.focus"				& vbCr    
		   Response.Write "Call parent.SetToolbar(""11101111000011"")"	& vbCr    
		   Response.Write "</Script>"									& vbCr	
		   Response.End																		'☜: Process End
		End If  
    
		iArrSDnHdr = iObjS5G211.LookUp(gStrGlobalCollection, iStrDnNo, "Y")
    
		If CheckSYSTEMError(Err,True) = True Then
		   Set iObjS5G211 = Nothing		                                                 '☜: Unload Comproxy DLL
		   Response.Write "<Script Language=vbscript>"					& vbCr
		   Response.Write "parent.frm1.txtConDN_no.focus"				& vbCr    
		   Response.Write "Call parent.SetToolbar(""11101111000011"")"	& vbCr    
		   Response.Write "</Script>"									& vbCr	
		   Response.End																		'☜: Process End
		End If  
    
		Set iObjS5G211 = Nothing

		'==================== 조회결과 Index ==================
		Const S5G211_RS_DN_NO = 0                ' 출하번호 
		Const S5G211_RS_SHIP_TO_PARTY = 1        ' 납품처 
		Const S5G211_RS_SHIP_TO_PARTY_NM = 2     ' 납품처명 
		Const S5G211_RS_SALES_GRP = 3            ' 영업그룹 
		Const S5G211_RS_SALES_GRP_NM = 4         ' 영업그룹명 
		Const S5G211_RS_SALES_ORG = 5            ' 영업조직 
		Const S5G211_RS_SALES_ORG_NM = 6         ' 영업조직명 
		Const S5G211_RS_MOV_TYPE = 7             ' 출하형태 
		Const S5G211_RS_MOV_TYPE_NM = 8          ' 출하형태명 
		Const S5G211_RS_DLVY_DT = 9             ' 납기일 
		Const S5G211_RS_PROMISE_DT = 10          ' 출고예정일 
		Const S5G211_RS_ACTUAL_GI_DT = 11        ' 출고일 
		Const S5G211_RS_COST_CD = 12             ' Cost Center
		Const S5G211_RS_BIZ_AREA = 13            ' 사업장 
		Const S5G211_RS_BIZ_AREA_NM = 14         ' 사업장명 
		Const S5G211_RS_TRANS_METH = 15          ' 운송방법 
		Const S5G211_RS_TRANS_METH_NM = 16       ' 운송방법명 
		Const S5G211_RS_GOODS_MV_NO = 17         ' 수불번호(출고번호)
		Const S5G211_RS_CI_FLAG = 18             ' 통관필요여부 
		Const S5G211_RS_POST_FLAG = 19           ' 출고처리여부 
		Const S5G211_RS_SO_TYPE = 20             ' 수주형태 
		Const S5G211_RS_SO_TYPE_NM = 21          ' 수주형태명 
		Const S5G211_RS_SO_NO = 22               ' 수주번호 
		Const S5G211_RS_SHIP_TO_PLCE = 23        ' 납품장송 
		Const S5G211_RS_INSRT_USER_ID = 24       ' 등록자 
		Const S5G211_RS_INSRT_DT = 25            ' 등록일 
		Const S5G211_RS_UPDT_USER_ID = 26        ' 변경자 
		Const S5G211_RS_UPDT_DT = 27             ' 변경일 
		Const S5G211_RS_EXT1_QTY = 28            ' 여유필드(수량)
		Const S5G211_RS_EXT2_QTY = 29            ' 여유필드(수량)
		Const S5G211_RS_EXT3_QTY = 30            ' 여유필드(수량)
		Const S5G211_RS_EXT1_AMT = 31            ' 여유필드(금액)
		Const S5G211_RS_EXT2_AMT = 32            ' 여유필드(금액)
		Const S5G211_RS_EXT3_AMT = 33            ' 여유필드(금액)
		Const S5G211_RS_EXT1_CD = 34             ' 여유필드(Text)
		Const S5G211_RS_EXT2_CD = 35             ' 여유필드(Text)
		Const S5G211_RS_EXT3_CD = 36             ' 여유필드(Text)
		Const S5G211_RS_TEMP_SO_NO = 37          ' 수주번호 
		Const S5G211_RS_VAT_FLAG = 38            ' 세금계산서정보 동시생성여부 
		Const S5G211_RS_AR_FLAG = 39             ' 매출정보 동시생성여부 
		Const S5G211_RS_CUR = 40                 ' 화폐단위 
		Const S5G211_RS_XCHG_RATE = 41           ' 환율 
		Const S5G211_RS_XCHG_RATE_OP = 42        ' 환율연산자 
		Const S5G211_RS_NET_AMT = 43             ' 출고금액 
		Const S5G211_RS_NET_AMT_LOC = 44         ' 출고자국금액 
		Const S5G211_RS_VAT_AMT = 45             ' VAT 금액 
		Const S5G211_RS_VAT_AMT_LOC = 46         ' VAT 자국금액 
		Const S5G211_RS_EXCEPT_DN_FLAG = 47      ' 예외출고여부 
		Const S5G211_RS_REMARK = 48              ' 비고 
		Const S5G211_RS_ARRIVAL_DT = 49          ' 실제납품일 
		Const S5G211_RS_ARRIVAL_TIME = 50        ' 납품시간 
		Const S5G211_RS_STP_INFO_NO = 51         ' 납품처상세정보 번호 
		Const S5G211_RS_ZIP_CD = 52              ' 납품처 우편번호 
		Const S5G211_RS_ADDR1 = 53               ' 납품처 주소 
		Const S5G211_RS_ADDR2 = 54               ' 납품처 주소 
		Const S5G211_RS_ADDR3 = 55               ' 납품처 주소 
		Const S5G211_RS_RECEIVER = 56            ' 인수자명 
		Const S5G211_RS_TEL_NO1 = 57             ' 전화번호1
		Const S5G211_RS_TEL_NO2 = 58             ' 전화번호2
		Const S5G211_RS_TRANS_INFO_NO = 59       ' 운송정보번호 
		Const S5G211_RS_TRANS_CO = 60            ' 운송회사 
		Const S5G211_RS_DRIVER = 61              ' 운전자명 
		Const S5G211_RS_VEHICLE_NO = 62          ' 차량번호 
		Const S5G211_RS_SENDER = 63              ' 인계자명 
		Const S5G211_RS_STO_FLAG = 64            ' STO Flag
		Const S5G211_RS_CASH_DC_AMT = 65         ' 현금할인액 
		Const S5G211_RS_TAX_DC_AMT = 66          ' 세금할인액 
		Const S5G211_RS_TAX_BASE_AMT = 67        ' 세금계산기초금액 
		Const S5G211_RS_CASH_DC_AMT_LOC = 68     ' 현금할인액(자국)
		Const S5G211_RS_TAX_DC_AMT_LOC = 69      ' 세금할인액(자국)
		Const S5G211_RS_TAX_BASE_AMT_LOC = 70    ' 세금계산기초금액(자국)
		Const S5G211_RS_SO_AUTO_FLAG = 71        ' 수주로부터 자동생성여 여부 
		Const S5G211_RS_PLANT_CD = 72            ' 공장 
		Const S5G211_RS_PLANT_NM = 73            ' 공장명 
		Const S5G211_RS_INV_MGR = 74             ' 재고담당 
		Const S5G211_RS_INV_MGR_NM = 75          ' 재고담당자명 
		Const S5G211_RS_CONTRY_CD = 76			 ' 납품처 국가코드 
		Const S5G211_RS_RET_ITEM_FLAG = 77       ' 반품여부 
		Const S5G211_RS_REL_BILL_FLAG = 78       ' 매출여부 

		' 예외출고 
		Const S5G211_RS_ORD_DT = 79
		Const S5G211_RS_CUST_PO_NO = 80
		Const S5G211_RS_CUST_PO_DT = 81
		Const S5G211_RS_DEAL_TYPE = 82
		Const S5G211_RS_DEAL_TYPE_NM = 83
		Const S5G211_RS_PAY_TYPE = 84
		Const S5G211_RS_PAY_TYPE_NM = 85
		Const S5G211_RS_PAY_METH = 86
		Const S5G211_RS_PAY_METH_NM = 87
		Const S5G211_RS_PAY_DUR = 88
		Const S5G211_RS_VAT_INC_FLAG = 89
		Const S5G211_RS_VAT_CALC_TYPE = 90
		Const S5G211_RS_VAT_TYPE = 91
		Const S5G211_RS_VAT_TYPE_NM = 92
		Const S5G211_RS_VAT_RATE = 93
		Const S5G211_RS_TAX_BIZ_AREA = 94
		Const S5G211_RS_TAX_BIZ_AREA_NM = 95
		Const S5G211_RS_SP_STK_FLAG = 96
		Const S5G211_RS_PAY_TERMS_TXT = 97
		Const S5G211_RS_COLLECT_TYPE = 98
		Const S5G211_RS_COLLECT_TYPE_NM = 99
		Const S5G211_RS_COLLECT_DOC_AMT = 100
		Const S5G211_RS_COLLECT_LOC_AMT = 101
		Const S5G211_RS_SL_CD = 102
		Const S5G211_RS_SL_NM = 103
		Const S5G211_RS_SOLD_TO_PARTY = 104
		Const S5G211_RS_SOLD_TO_PARTY_NM = 105
		Const S5G211_RS_PAY_TERM = 106
		Const S5G211_RS_CASH_DC_RATE = 107
		Const S5G211_RS_TAX_CALC_TYPE = 108
		Const S5G211_RS_CASH_DC_TYPE = 109
		Const S5G211_RS_EVIDENCE_TYPE = 110

		Response.Write "<Script Language=vbscript>" & vbCr
		Response.Write "With parent.frm1" & vbCr
		Response.Write ".txtDnNo.value				= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_DN_NO))			& """" & vbcr
		Response.Write ".txtDn_Type.value			= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_MOV_TYPE))			& """" & vbcr
		Response.Write ".txtDn_TypeNm.value			= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_MOV_TYPE_NM))		& """" & vbcr
		Response.Write ".txtShip_to_party.value		= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_SHIP_TO_PARTY))	& """" & vbcr
		Response.Write ".txtShip_to_partyNm.value	= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_SHIP_TO_PARTY_NM))	& """" & vbcr
		Response.Write ".txtSold_to_party.value		= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_SOLD_TO_PARTY))	& """" & vbcr
		Response.Write ".txtSold_to_partyNm.value	= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_SOLD_TO_PARTY_NM))	& """" & vbcr
		Response.Write "Call parent.SetTransLT()" & vbCr
		Response.Write ".txtDeal_Type.value			= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_DEAL_TYPE))		& """" & vbcr
		Response.Write ".txtDeal_Type_nm.value		= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_DEAL_TYPE_NM))		& """" & vbcr
		Response.Write ".txtDlvyDt.Text				= """ & UNIDateClientFormat(iArrSDnHdr(S5G211_RS_DLVY_DT))	& """" & vbcr
		Response.Write ".txtPlannedGIDt.Text		= """ & UNIDateClientFormat(iArrSDnHdr(S5G211_RS_PROMISE_DT))	& """" & vbcr
		Response.Write ".txtSales_Grp.value			= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_SALES_GRP))		& """" & vbcr
		Response.Write ".txtSales_GrpNm.value		= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_SALES_GRP_NM))		& """" & vbcr
		' 결재방법(주의:DB 필드명과 asp의 필드명이 다름)
		Response.Write ".txtPay_terms.value			= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_PAY_METH))			& """" & vbcr
		Response.Write ".txtPay_terms_nm.value		= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_PAY_METH_NM))		& """" & vbcr
		Response.Write ".txtTaxBizAreaCd.value		= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_TAX_BIZ_AREA))		& """" & vbcr
		Response.Write ".txtTaxBizAreaNm.value		= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_TAX_BIZ_AREA_NM))	& """" & vbcr
		Response.Write ".txt_Payterms_txt.value		= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_PAY_TERMS_TXT))	& """" & vbcr
		Response.Write ".txtVat_Type.value			= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_VAT_TYPE))			& """" & vbcr
		Response.Write ".txtVatTypeNm.value			= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_VAT_TYPE_NM))		& """" & vbcr
		Response.Write ".txtVat_rate.Text			= """ & UNINumClientFormat(iArrSDnHdr(S5G211_RS_VAT_RATE),ggExchRate.DecPoint, 0)	& """" & vbcr
	
		' VAT 포함여부 
		If iArrSDnHdr(S5G211_RS_VAT_INC_FLAG) = "1" Then
			Response.Write ".rdoVat_Inc_flag1.checked = True" & vbCr
		Else
			Response.Write ".rdoVat_Inc_flag2.checked = True" & vbCr
		End If
	
		' VAT 적용기준 
		If iArrSDnHdr(S5G211_RS_VAT_CALC_TYPE) = "1" Then
			Response.Write ".rdoVat_Calc_Type1.checked = True" & vbCr
		Else
			Response.Write ".rdoVat_Calc_Type2.checked = True" & vbCr
		End If

		Response.Write ".txtCurrency.value	= """ & iArrSDnHdr(S5G211_RS_CUR)		& """" & vbcr
		Response.Write ".txtNet_Amt.Text	= """ & UNIConvNumDBToCompanyByCurrency(iArrSDnHdr(S5G211_RS_NET_AMT_LOC), gCurrency, ggAmtOfMoneyNo, gLocRndPolicyNo, "X") & """" & vbcr
		Response.Write ".txtVAT_Amt.Text	= """ & UNIConvNumDBToCompanyByCurrency(iArrSDnHdr(S5G211_RS_VAT_AMT_LOC), gCurrency, ggAmtOfMoneyNo, gLocRndPolicyNo, "X") & """" & vbcr
		Response.Write ".txtTot_amt.Text	= """ & UNIConvNumDBToCompanyByCurrency(CDbl(iArrSDnHdr(S5G211_RS_NET_AMT_LOC)) + CDbl(iArrSDnHdr(S5G211_RS_VAT_AMT_LOC)), gCurrency, ggAmtOfMoneyNo, gLocRndPolicyNo, "X") & """" & vbcr
		Response.Write ".txtTotal_Amt.Text	= """ & UNIConvNumDBToCompanyByCurrency(CDbl(iArrSDnHdr(S5G211_RS_NET_AMT_LOC)) + CDbl(iArrSDnHdr(S5G211_RS_VAT_AMT_LOC)), gCurrency, ggAmtOfMoneyNo, gLocRndPolicyNo, "X") & """" & vbcr
		Response.Write ".txtHCntryCd.value	= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_CONTRY_CD))	& """" & vbcr
		Response.Write ".txtRemark.value	= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_REMARK))		& """" & vbcr
		Response.Write ".txtArriv_dt.Text	= """ & UNIDateClientFormat(iArrSDnHdr(S5G211_RS_ARRIVAL_DT))	& """" & vbcr
		Response.Write ".txtArriv_tm.value	= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_ARRIVAL_TIME))	& """" & vbcr
		Response.Write ".txtInvMgr.value	= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_INV_MGR))	& """" & vbcr
		Response.Write ".txtInvMgrNm.value	= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_INV_MGR_NM))	& """" & vbcr

		' 납품처상세정보번호 
		Response.Write ".txtSTP_Inf_No.value	= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_STP_INFO_NO))		& """" & vbcr		' 납품처상세정보번호 
		Response.Write ".txtTrnsp_Inf_No.value	= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_TRANS_INFO_NO))	& """" & vbcr
		Response.Write ".txtZIP_cd.value	= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_ZIP_CD))			& """" & vbcr
		Response.Write ".txtADDR1_Dlv.value	= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_ADDR1))			& """" & vbcr
		Response.Write ".txtADDR2_Dlv.value	= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_ADDR2))			& """" & vbcr
		Response.Write ".txtADDR3_Dlv.value	= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_ADDR3))			& """" & vbcr
		Response.Write ".txtReceiver.value	= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_RECEIVER))			& """" & vbcr
		Response.Write ".txtTel_No1.value	= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_TEL_NO1))			& """" & vbcr
		Response.Write ".txtTel_No2.value	= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_TEL_NO2))			& """" & vbcr
		Response.Write ".txtTransCo.value	= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_TRANS_CO))			& """" & vbcr
		Response.Write ".txtDriver.value	= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_DRIVER))			& """" & vbcr
		Response.Write ".txtVehicleNo.value	= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_VEHICLE_NO))		& """" & vbcr
		Response.Write ".txtSender.value	= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_SENDER))			& """" & vbcr
	
		Response.Write ".txtTrans_Meth.value	= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_TRANS_METH))	& """" & vbcr
		Response.Write ".txtTrans_Meth_nm.value	= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_TRANS_METH_NM))	& """" & vbcr
		Response.Write ".txtDlvyPlace.value		= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_SHIP_TO_PLCE))	& """" & vbcr
		Response.Write ".txtCol_Type.value		= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_COLLECT_TYPE))	& """" & vbcr
		Response.Write ".txtCol_Type_nm.value	= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_COLLECT_TYPE_NM))	& """" & vbcr
		Response.Write ".txtCol_amt.Text		= """ & UNIConvNumDBToCompanyByCurrency(iArrSDnHdr(S5G211_RS_COLLECT_LOC_AMT), gCurrency, ggAmtOfMoneyNo, gLocRndPolicyNo, "X") & """" & vbcr
		Response.Write ".txtPlant.value		= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_PLANT_CD))	& """" & vbcr
		Response.Write ".txtPlantNm.value	= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_PLANT_NM))	& """" & vbcr
		Response.Write ".txtSlCd.value	= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_SL_CD))		& """" & vbcr
		Response.Write ".txtSlNm.value	= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_SL_NM))		& """" & vbcr
		Response.Write ".txtGINo.value	= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_GOODS_MV_NO))	& """" & vbcr
	
		' 출고처리 된 경우 
		If Trim(iArrSDnHdr(S5G211_RS_GOODS_MV_NO)) <> "" Then 
			Response.Write ".txtGI_Dt.Text		= """ & UNIDateClientFormat(iArrSDnHdr(S5G211_RS_ACTUAL_GI_DT))	& """" & vbcr
		End if
		Response.Write ".txtRetItemFlag.value	= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_RET_ITEM_FLAG))	& """" & vbcr
		Response.Write ".txtRetBillFlag.value	= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_REL_BILL_FLAG))		& """" & vbcr

		If UCase(iArrSDnHdr(S5G211_RS_AR_FLAG)) = "Y" Or Trim(iArrSDnHdr(S5G211_RS_COLLECT_TYPE)) <> "" Then
			Response.Write ".chkARflag.checked = True" & vbcr
			Response.Write ".txtARFlag.value = ""Y""" & vbcr
			Response.Write "parent.lblArFlag.disabled = False" & vbcr
		Else
			Response.Write ".chkARflag.checked = False" & vbcr
			Response.Write ".txtARFlag.value = ""N""" & vbcr
			Response.Write "parent.lblArFlag.disabled = True" & vbcr
		End If

		If UCase(iArrSDnHdr(S5G211_RS_VAT_FLAG)) = "Y" Then
			Response.Write ".chkVatFlag.checked = True" & vbcr
			Response.Write ".txtVatFlag.value = ""Y""" & vbcr
			Response.Write "parent.lblVatFlag.disabled = False" & vbcr
		Else
			Response.Write ".chkVatFlag.checked = False" & vbcr
			Response.Write ".txtVatFlag.value = ""N""" & vbcr
			Response.Write "parent.lblVatFlag.disabled = True" & vbcr
		End If
	
		If UCase(iArrSDnHdr(S5G211_RS_RET_ITEM_FLAG)) = "Y" Then
			Response.Write ".btnPosting.value = ""입고처리""" & vbcr
			Response.Write ".btnPostCancel.value = ""입고처리취소""" & vbcr
		Else
			Response.Write ".btnPosting.value = ""출고처리""" & vbcr
			Response.Write ".btnPostCancel.value = ""출고처리취소""" & vbcr
		End If

		Response.Write ".txtHDnNo.value	= """ & ConvSPChars(iArrSDnHdr(S5G211_RS_DN_NO)) & """" & vbcr
		Response.Write "Call parent.SetToolbar(""11111111000111"")" & vbcr
		Response.Write "End With"          & vbCr
		Response.Write "</Script>"         & vbCr															'☜: 조회가 성공 
		
	End If ' 출하 Header 정보 조회 끝.
	
	'-----------------------
    ' 출하내역을 읽어온다.
    '-----------------------
    
	Dim i1_s_dn_dtl, i2_s_dn_hdr, e1_s_dn_dtl, eg1_exp_grp

	Const S424_I1_dn_seq = 0    '[CONVERSION INFORMATION]  View Name : imp_next s_dn_dtl

    Const S424_I2_dn_no = 0    '[CONVERSION INFORMATION]  View Name : imp s_dn_hdr
    Const S424_I2_except_dn_flag = 1
    
    Const C_SHEETMAXROWS_D  = 100
    
    Redim i1_s_dn_dtl(S424_I1_dn_seq)
    Redim i2_s_dn_hdr(S424_I2_except_dn_flag)
    
    i2_s_dn_hdr(S424_I2_dn_no) = iStrDnNo
    i2_s_dn_hdr(S424_I2_except_dn_flag) = "Y"
    
    If lgStrPrevKey = "" Then
		i1_s_dn_dtl(S424_I1_dn_seq) = 0
    Else
		i1_s_dn_dtl(S424_I1_dn_seq) = CLng(lgStrPrevKey)
	End if	    

    Set iObjS5G128 = Server.CreateObject("pS5G128.cSListDnDtl")
        
    If CheckSYSTEMError(Err,True) = True Then
		Response.Write "<Script Language=vbscript>" & vbCr
'		Response.Write "Call parent.DbQueryOk" & vbCr
		Response.Write "</Script>" & vbCr
		Response.End																				'☜: Process End
    End If
   
	call iObjS5G128.S_LIST_DN_DTL(gStrGlobalCollection, C_SHEETMAXROWS_D, i1_s_dn_dtl, i2_s_dn_hdr, eg1_exp_grp)

	If CheckSYSTEMError(Err,True) = True Then
		Set iObjS5G128 = Nothing					
		Response.Write "<Script Language=vbscript>" & vbCr
		Response.Write "Call parent.DbHdrQueryOk" & vbCr
		Response.Write "</Script>" & vbCr
		Response.End																				'☜: Process End
	End If
	
	Set iObjS5G128 = Nothing
	
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

	' 조회결과 Index
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
	Const EA_b_item_by_plant_lot_flg2 = 45
	Const EA_b_item_by_plant_ship_inspec_flg2 = 46
	const EA_b_minor_ret_type_nm2 = 47 
	const EA_b_minor_vat_type_nm2 = 48
	Const EA_s_dn_dtl_carton_no2 = 49

	'-----------------------
	' 출하내역의 내용을 표시한다.
	'----------------------- 
	Dim iArrCols, iArrRow
	ReDim iArrCols(35)						' Column 수 
	Redim iArrRows(iLngSheetMaxRows)		' 조회된 Row 수만큼 배열 재정의 

	iArrCols(0) = ""
   	iArrCols(2)  = ""	' 창고 Popup
   	iArrCols(4)  = ""	' 품목 Popup
   	iArrCols(8)  = ""	' 단위 Popup
   	iArrCols(14)  = ""	' Lot Popup
   	iArrCols(20)  = ""	' VAT Popup
   	iArrCols(25)  = ""	' 반품유형 Popup
   	
	For LngRow = 0 To UBound(EG1_EXP_GRP,1)	
   		iArrCols(1)  = ConvSPChars(EG1_EXP_GRP(LngRow,EA_b_storage_location_sl_cd2))	'창고 
   		iArrCols(3)  = ConvSPChars(EG1_EXP_GRP(LngRow,EA_b_item_item_cd2))				'품목코드 
   		iArrCols(5)  = ConvSPChars(EG1_EXP_GRP(LngRow,EA_b_item_item_nm2))				'품목명 
   		iArrCols(6)  = ConvSPChars(EG1_EXP_GRP(LngRow,EA_b_item_spec2))					'품목규격 
   		iArrCols(7)  = ConvSPChars(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_gi_unit2))			'단위 

		iArrCols(9)  = UNINumClientFormat(EG1_exp_grp(LngRow, EA_s_dn_dtl_gi_qty2), ggQty.DecPoint, 0)					' 출고수량 
		iArrCols(10)  = UNINumClientFormatByCurrency(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_price2), gCurrency, ggUnitCostNo)	' 단가 

		If EG1_EXP_GRP(LngRow,EA_s_dn_dtl_vat_inc_flag2) = "2" Then
			iArrCols(11)  = UNIConvNumDBToCompanyByCurrency(CDbl(EG1_EXP_GRP(LngRow, EA_s_dn_dtl_gi_amt_loc2)) + CDbl(EG1_EXP_GRP(LngRow, EA_s_dn_dtl_vat_amt_loc2)), gCurrency, ggAmtOfMoneyNo, gLocRndPolicyNo, "X")	' 총금액 
		Else
			iArrCols(11)  = UNIConvNumDBToCompanyByCurrency(EG1_EXP_GRP(LngRow, EA_s_dn_dtl_gi_amt_loc2), gCurrency, ggAmtOfMoneyNo, gLocRndPolicyNo, "X")	' 총금액 
		End If

		iArrCols(12)  = UNIConvNumDBToCompanyByCurrency(EG1_EXP_GRP(LngRow, EA_s_dn_dtl_gi_amt_loc2), gCurrency, ggAmtOfMoneyNo, gLocRndPolicyNo, "X")		' 출고금액 
   		iArrCols(13) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_lot_no2))						'Lot번호 
   		iArrCols(15) = UNINumClientFormat(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_lot_seq2), 0, 0)		'Lot순번 
   		iArrCols(16) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_carton_no2))					'carton no
   		iArrCols(17) = EG1_EXP_GRP(LngRow,EA_s_dn_dtl_vat_inc_flag2)							'VAT 포함여부 

		If EG1_EXP_GRP(LngRow,EA_s_dn_dtl_vat_inc_flag2) = "1" Then
	   		iArrCols(18) = "별도"
	   	Else
	   		iArrCols(18) = "포함"
	   	End If

   		iArrCols(19) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_vat_type2))									'VAT유형 
   		iArrCols(21) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_b_minor_vat_type_nm2))									'VAT 명 
   		iArrCols(22) = UNINumClientFormat(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_vat_rate2), ggExchRate.DecPoint, 0)	'VAT율 
		iArrCols(23) = UNIConvNumDBToCompanyByCurrency(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_vat_amt_loc2), gCurrency, ggAmtOfMoneyNo, gLocRndPolicyNo, "X")	' VAT금액 
   		iArrCols(24) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_ret_type2))									'반품유형 
   		iArrCols(26) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_b_minor_ret_type_nm2))									'반품유형명 
		iArrCols(27) = UNIConvNumDBToCompanyByCurrency(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_deposit_amt2), gCurrency, ggAmtOfMoneyNo, gLocRndPolicyNo, "X")			' Picking 수량 
   		iArrCols(28) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_remark2))										'비고 
   		iArrCols(29) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_b_item_by_plant_lot_flg2))								'Lot 관리여부 
   		iArrCols(30) = ConvSPChars(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_dn_seq2))										'출고순번 

		iArrCols(31) = UNIConvNumDBToCompanyByCurrency(EG1_EXP_GRP(LngRow, EA_s_dn_dtl_gi_amt_loc2), gCurrency, ggAmtOfMoneyNo, gLocRndPolicyNo, "X")	' 출고금액 
		iArrCols(32) = UNIConvNumDBToCompanyByCurrency(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_vat_amt_loc2), gCurrency, ggAmtOfMoneyNo, gLocRndPolicyNo, "X")	' VAT금액 
		iArrCols(33) = UNIConvNumDBToCompanyByCurrency(EG1_EXP_GRP(LngRow, EA_s_dn_dtl_gi_amt_loc2), gCurrency, ggAmtOfMoneyNo, gLocRndPolicyNo, "X")	' 출고금액 
		iArrCols(34) = UNIConvNumDBToCompanyByCurrency(EG1_EXP_GRP(LngRow,EA_s_dn_dtl_vat_amt_loc2), gCurrency, ggAmtOfMoneyNo, gLocRndPolicyNo, "X")	' VAT금액 

		iArrCols(35) = iLngLastRow + LngRow

   		iArrRows(LngRow) = Join(iArrCols, gColSep)
	Next
	
	Response.Write "<Script language=vbs> " & vbCr   
	Response.Write "With parent " & vbCr   
    Response.Write " .ggoSpread.Source = .frm1.vspdData" & vbCr
    Response.Write " .frm1.vspdData.Redraw = False  "      & vbCr      
    Response.Write " .ggoSpread.SSShowDataByClip   """ & Join(iArrRows, gColSep & gRowSep) & gColSep & gRowSep & """ ,""F""" & vbCr
    Response.Write " .lgStrPrevKey = """ & iStrNextKey  & """" & vbCr  
    
    Response.Write " If Trim(parent.frm1.txtGINo.value) = """" Then" & vbCr  
    Response.Write		" If Trim(lgStrPrevKey) = """" Then" & vbCr  
    Response.Write			" .SetSpreadColor 1, .frm1.vspdData.MaxRows " & vbCr  
    Response.Write		" Else" & vbCr  
    Response.Write			" .SetSpreadColor " & iLngLastRow + 1 & ", .frm1.vspdData.MaxRows " & vbCr  
    Response.Write		" End If" & vbCr  
    Response.Write " Else" & vbCr  
    Response.Write " .SetSpreadColorConfirmed -1 " & vbCr  
    Response.Write " End If" & vbCr  
	Response.Write ".frm1.vspdData.Redraw = True  "       & vbCr

    
    Response.Write " If .frm1.vspdData.MaxRows <= .VisibleRowCnt(.frm1.vspdData,NewTop) And .lgStrPrevKey <> """" Then	 " & vbCr	         
	Response.Write		" .DbQuery  " & vbCr	' GroupView 사이즈로 화면 Row수보다 쿼리가 작으면 다시 쿼리함 
	Response.Write " Else  " & vbCr
    Response.Write		" .DbQueryOk " & vbCr   
	Response.Write " End If	 " & vbCr
	Response.Write "End With " & vbCr   
	Response.Write "</Script> " & vbCr          
	Response.End																				'☜: Process End

Case CStr(UID_M0002)			'☜: 저장 요청을 받음 
	Dim iErrorPosition
	
	' ========== Header 정보가 변경된 경우 Header 정보 Insert / Update 처리 ==========
	If Trim(Request("txtHdrStateFlg")) = 2 Then 
		Dim iStrSTPInfoNo, iStrTransInfoNo
		Dim iStrCrSTPFlag, iStrCrTransFlag, lgIntFlgMode

		lgIntFlgMode = CInt(Request("txtFlgMode"))									'☜: 저장시 Create/Update 판별 

		If lgIntFlgMode = OPMD_CMODE Then
			iStrCUDFlag = "C"
		ElseIf lgIntFlgMode = OPMD_UMODE Then
			iStrCUDFlag = "U"
		Else
			Call ServerMesgBox("TXTFLGMODE 조건값이 비어있습니다!", vbInformation, I_MKSCRIPT)              
			Response.End 
		End If
    
		' 출하정보 
		Const S5G211_DnHdr_DN_NO = 0           '(O)출하번호 
		Const S5G211_DnHdr_EXCEPT_DN_FLAG = 1  '(M)예외출고여부(Y/N)	
		Const S5G211_DnHdr_SHIP_TO_PARTY = 2   '(M)납품처 
		Const S5G211_DnHdr_SALES_GRP = 3       '(M)영업그룹 
		Const S5G211_DnHdr_MOV_TYPE = 4        '(M)출하형태 
		Const S5G211_DnHdr_DLVY_DT = 5         '(M)납기일 
		Const S5G211_DnHdr_PROMISE_DT = 6      '(M)출고예정일 
		Const S5G211_DnHdr_TRANS_METH = 7      '(O)운송방법 
		Const S5G211_DnHdr_SO_TYPE = 8         '(M)수주형태 
		Const S5G211_DnHdr_SO_NO = 9           '(O)S/O번호 
		Const S5G211_DnHdr_SHIP_TO_PLCE = 10    '(O)납품장소 
		Const S5G211_DnHdr_CUR = 21            '(O)화폐단위 
		Const S5G211_DnHdr_XCHG_RATE = 22      '(0)환율 
		Const S5G211_DnHdr_XCHG_RATE_OP = 23   '(O)환율연산자 
		Const S5G211_DnHdr_REMARK = 24         '(O)비고 
		Const S5G211_DnHdr_ARRIVAL_DT = 25     '(O)실제납품일 
		Const S5G211_DnHdr_ARRIVAL_TIME = 26   '(O)납품시간 
		Const S5G211_DnHdr_SO_AUTO_FLAG = 27   '(O)납품시간 
		Const S5G211_DnHdr_PLANT_CD = 28       '(M)공장 
		Const S5G211_DnHdr_INV_MGR = 29        '(O)재고담당 

		Redim iArrSDnHdr(S5G211_DnHdr_INV_MGR)

		' 예외출고 
		Dim iArrSDnSales
		 
	    Const S5G211_DnSales_DN_NO = 0               '(O)출하번호 
	    Const S5G211_DnSales_ORD_DT = 1              '(O)주문일 
	    Const S5G211_DnSales_CUST_PO_NO = 2          '(O)고객주문번호 
	    Const S5G211_DnSales_CUST_PO_DT = 3          '(O)고객주문일 
	    Const S5G211_DnSales_DEAL_TYPE = 4           '(M)판매유형 
	    Const S5G211_DnSales_PAY_TYPE = 5            '(O)입금유형 
	    Const S5G211_DnSales_PAY_METH = 6            '(M)결재방법 
	    Const S5G211_DnSales_PAY_DUR = 7             '(0)결재기간 
	    Const S5G211_DnSales_VAT_INC_FLAG = 8        '(M)VAT 포함구분(1-별도, 2-포함)
	    Const S5G211_DnSales_VAT_CALC_TYPE = 9       '(M)VAT 적용기준(1-개벌, 2-통합)
	    Const S5G211_DnSales_VAT_TYPE = 10           '(O)VAT 유형(통합일때 필수)
	    Const S5G211_DnSales_VAT_RATE = 11           '(O)VAT 율 
	    Const S5G211_DnSales_TAX_BIZ_AREA = 12       '(M)세금신고사업장 
	    Const S5G211_DnSales_SP_STK_FLAG = 13        '(O)수탁재고여부-현재 사용하지 않음 
	    Const S5G211_DnSales_PAY_TERMS_TXT = 14      '(O)대금결제참조 
	    Const S5G211_DnSales_REMARK = 15             '(O)비고 
	    Const S5G211_DnSales_COLLECT_TYPE = 16       '(O)입금유형 
	    Const S5G211_DnSales_COLLECT_DOC_AMT = 17    '(O)수금액 
	    Const S5G211_DnSales_COLLECT_LOC_AMT = 18    '(O)수금자국금액 
	    Const S5G211_DnSales_SL_CD = 19              '(O)창고 
	    Const S5G211_DnSales_SOLD_TO_PARTY = 29      '(M)주문처 
		' 유럽버전에서만 사용 
	    Const S5G211_DnSales_PAY_TERM = 30           '(O)입금유형 
	    Const S5G211_DnSales_CASH_DC_RATE = 31       '(O)현금할인율 
	    Const S5G211_DnSales_TAX_CALC_TYPE = 32      '(O)세금할인유형 
	    Const S5G211_DnSales_CASH_DC_TYPE = 33       '(O)현금할인유형 
	    Const S5G211_DnSales_EVIDENCE_TYPE = 34      '(O)증빙유형 

		Redim iArrSDnSales(S5G211_DnSales_EVIDENCE_TYPE)

		' 납품처 상세정보 
		Dim iArrSTPInfo
    
		Const S5G211_STPInfo_STP_INFO_NO = 0
		Const S5G211_STPInfo_SHIP_TO_PARTY = 1
		Const S5G211_STPInfo_ZIP_CD = 2
		Const S5G211_STPInfo_ADDR1 = 3
		Const S5G211_STPInfo_ADDR2 = 4
		Const S5G211_STPInfo_ADDR3 = 5
		Const S5G211_STPInfo_RECEIVER = 6
		Const S5G211_STPInfo_TEL_NO1 = 7
		Const S5G211_STPInfo_TEL_NO2 = 8

		Redim iArrSTPInfo(S5G211_STPInfo_TEL_NO2)

		' 운송상세정보 
		Dim iArrTransInfo
    
		Const S5G211_TransInfo_TRANS_INFO_NO = 0
		Const S5G211_TransInfo_TRANS_CO = 1
		Const S5G211_TransInfo_DRIVER = 2
		Const S5G211_TransInfo_VEHICLE_NO = 3
		Const S5G211_TransInfo_SENDER = 4

		Redim iArrTransInfo(S5G211_TransInfo_SENDER)

		'-----------------------
		'Data manipulate area
		'-----------------------
		iArrSDnHdr(S5G211_DnHdr_DN_NO) = UCase(Trim(Request("txtDnNo")))					'(O)출하번호 
		iArrSDnHdr(S5G211_DnHdr_EXCEPT_DN_FLAG) = "Y"										'(M)예외출고여부 
		iArrSDnHdr(S5G211_DnHdr_SHIP_TO_PARTY) = UCase(Trim(Request("txtShip_to_party")))	'(M)납품처 
		iArrSDnHdr(S5G211_DnHdr_MOV_TYPE) = UCase(Trim(Request("txtDn_Type")))				'(M)출하형태 
		iArrSDnHdr(S5G211_DnHdr_SALES_GRP) = UCase(Trim(Request("txtSales_Grp")))			'(M)영업그룹 
		iArrSDnHdr(S5G211_DnHdr_DLVY_DT) = UNIConvDate(Request("txtDlvyDt"))				'(M)납기일 
		iArrSDnHdr(S5G211_DnHdr_PROMISE_DT) = UNIConvDate(Request("txtPlannedGIDt"))		'(M)출고예정일 
		iArrSDnHdr(S5G211_DnHdr_TRANS_METH) = UCase(Trim(Request("txtTrans_Meth")))			'(O)운송방법 
		iArrSDnHdr(S5G211_DnHdr_SHIP_TO_PLCE) = Trim(Request("txtDlvyPlace"))				'(O)납품장소 
		iArrSDnHdr(S5G211_DnHdr_SO_TYPE) = UCase(Trim(Request("txtSO_TYPE")))				'(M)수주형태 
		iArrSDnHdr(S5G211_DnHdr_CUR) = UCase(Trim(Request("txtCurrency")))					'(O)화폐단위 
		iArrSDnHdr(S5G211_DnHdr_XCHG_RATE) = "1"      										'(0)환율 
		iArrSDnHdr(S5G211_DnHdr_XCHG_RATE_OP) = "*"   										'(O)환율연산자 
		iArrSDnHdr(S5G211_DnHdr_REMARK) = UCase(Trim(Request("txtRemark")))					'(O)비고 
		iArrSDnHdr(S5G211_DnHdr_ARRIVAL_DT) = UNIConvDate(Request("txtArriv_dt"))			'(O)실제납품일 
		iArrSDnHdr(S5G211_DnHdr_ARRIVAL_TIME) = Trim(Request("txtArriv_Tm"))				'(O)납품시간 
		iArrSDnHdr(S5G211_DnHdr_SO_AUTO_FLAG) = "N"											'(O)납품시간 
		iArrSDnHdr(S5G211_DnHdr_PLANT_CD) = UCase(Trim(Request("txtPlant")))				'(M)공장 
		iArrSDnHdr(S5G211_DnHdr_INV_MGR) = UCase(Trim(Request("txtInvMgr")))				'(O)재고담당 

		' 예외출고정보 
		iArrSDnSales(S5G211_DnSales_DEAL_TYPE) = UCase(Trim(Request("txtDeal_Type")))
		iArrSDnSales(S5G211_DnSales_PAY_METH) = UCase(Trim(Request("txtPay_terms")))
		iArrSDnSales(S5G211_DnSales_TAX_BIZ_AREA) = UCase(Trim(Request("txtTaxBizAreaCd")))
		iArrSDnSales(S5G211_DnSales_PAY_TERMS_TXT) = Trim(Request("txt_Payterms_txt"))
		iArrSDnSales(S5G211_DnSales_VAT_TYPE) = UCase(Trim(Request("txtVat_Type")))
		iArrSDnSales(S5G211_DnSales_VAT_RATE) = UNIConvNum(Request("txtVat_rate"),0)
		iArrSDnSales(S5G211_DnSales_VAT_INC_FLAG) = Trim(Request("rdoVat_Inc_flag"))
		iArrSDnSales(S5G211_DnSales_VAT_CALC_TYPE) = Trim(Request("rdoVat_Calc_Type"))
		iArrSDnSales(S5G211_DnSales_COLLECT_TYPE) = Trim(Request("txtCol_Type"))
		iArrSDnSales(S5G211_DnSales_COLLECT_DOC_AMT) = UNIConvNum(Request("txtCol_amt"), 0)
		iArrSDnSales(S5G211_DnSales_COLLECT_LOC_AMT) = UNIConvNum(Request("txtCol_amt"), 0)
		iArrSDnSales(S5G211_DnSales_SL_CD) = UCase(Trim(Request("txtSlCd")))
		iArrSDnSales(S5G211_DnSales_SOLD_TO_PARTY) = UCase(Trim(Request("txtSold_to_party")))

		' 납품처 상세정보 생성여부 
		If Request("txtlgBlnChgValue1") = "True" Then
			iStrCrSTPFlag = "C"		' 생성 
		Else
			iStrCrSTPFlag = "N"
		End If
	
		iStrSTPInfoNo = UCase(Trim(Request("txtSTP_Inf_No")))			'납품처상세정보번호 
		IF iStrSTPInfoNo = "" Then
			iArrSTPInfo(S5G211_STPInfo_STP_INFO_NO) = ""	
			iArrSTPInfo(S5G211_STPInfo_ZIP_CD) = UCase(Trim(Request("txtZIP_cd")))
			iArrSTPInfo(S5G211_STPInfo_ADDR1) = Trim(Request("txtADDR1_Dlv"))
			iArrSTPInfo(S5G211_STPInfo_ADDR2) = Trim(Request("txtADDR2_Dlv"))
			iArrSTPInfo(S5G211_STPInfo_ADDR3) = Trim(Request("txtADDR3_Dlv"))
			iArrSTPInfo(S5G211_STPInfo_RECEIVER) = Trim(Request("txtReceiver"))
			iArrSTPInfo(S5G211_STPInfo_TEL_NO1) = UCase(Trim(Request("txtTel_No1")))
			iArrSTPInfo(S5G211_STPInfo_TEL_NO2) = UCase(Trim(Request("txtTel_No2")))

			' 2003.09.20 - By Hwang Seongbae
			If Trim(Join(iArrSTPInfo, "")) <> "" Then
				iArrSTPInfo(S5G211_STPInfo_SHIP_TO_PARTY) = UCase(Trim(Request("txtShip_to_party")))
			End If
		Else
			iArrSTPInfo(S5G211_STPInfo_STP_INFO_NO) = iStrSTPInfoNo
		End If
	
		' 운송상세정보 변경여부 
		If Request("txtlgBlnChgValue2") = "True" Then
			iStrCrTransFlag = "C"
		Else
			iStrCrTransFlag = "N"
		End If
	
		iStrTransInfoNo = UCase(Trim(Request("txtTrnsp_Inf_No")))		'운송정보번호 
		IF iStrTransInfoNo = "" Then
			iArrTransInfo(S5G211_TransInfo_TRANS_INFO_NO) = ""
			iArrTransInfo(S5G211_TransInfo_TRANS_CO) = UCase(Trim(Request("txtTransCo")))
			iArrTransInfo(S5G211_TransInfo_DRIVER) = Trim(Request("txtDriver"))
			iArrTransInfo(S5G211_TransInfo_VEHICLE_NO) = UCase(Trim(Request("txtVehicleNo")))
			iArrTransInfo(S5G211_TransInfo_SENDER) = Trim(Request("txtSender"))
		Else
			iArrTransInfo(S5G211_TransInfo_TRANS_INFO_NO) = iStrTransInfoNo
		End If
		'###################
		'2003.01.02 SMJ	
		pvCB = "F"
		'###################

		Set iObjS5G211 = Server.CreateObject("PS5G211.cSDnHdrSvr")

		If CheckSYSTEMError(Err,True) = True Then
			Response.Write "<Script language=vbs> " & vbCr   
			Response.Write "Call parent.RemovedivTextArea " & vbCr   
			Response.Write "</Script> "																				         & vbCr          
			Response.End       
		End If
		
		Call iObjS5G211.Maintain(pvCB, gStrGlobalCollection, iStrCUDFlag, iArrSDnHdr, iArrSDnSales , "Y", , _
								 iStrCrSTPFlag, iArrSTPInfo, iStrCrTransFlag, iArrTransInfo, iStrDnNo)
    
		If CheckSYSTEMError(Err,True) = True Then
			Set iObjS5G211 = Nothing
			Response.Write "<Script language=vbs> " & vbCr   
			Response.Write "Call parent.RemovedivTextArea " & vbCr   
			Response.Write "</Script> "																				         & vbCr          
			Response.End
		End If
    
		Set iObjS5G211 = Nothing

	End If	' Header 변경 사항 적용 끝.

	'===================================================
	'				내역 저장 
	'===================================================
    Dim itxtSpreadIns, itxtSpreadUpd, itxtSpreadDel
    Dim iIntIndex
    Dim itxtSpreadArr
    Dim iCCount, iUCount, iDCount

	' 변경시 
	If iStrDnNo = "" Then
		iStrDnNo = Trim(Request("txtDnNo"))
	End If
	
    iCCount = Request.Form("txtCSpread").Count
    iUCount = Request.Form("txtUSpread").Count
    iDCount  = Request.Form("txtDSpread").Count

	' 삭제 
    ReDim itxtSpreadArr(iDCount)
    For iIntIndex = 1 To iDCount
        itxtSpreadArr(iIntIndex) = Request.Form("txtDSpread")(iIntIndex)
    Next
    itxtSpreadDel = Join(itxtSpreadArr,"")
    
    ' 수정 
    ReDim itxtSpreadArr(iUCount)
    For iIntIndex = 1 To iUCount
        itxtSpreadArr(iIntIndex) = Request.Form("txtUSpread")(iIntIndex)
    Next
    itxtSpreadUpd = Join(itxtSpreadArr,"")
    
    ' 추가 
    ReDim itxtSpreadArr(iCCount)
    For iIntIndex = 1 To iCCount
        itxtSpreadArr(iIntIndex) = Request.Form("txtCSpread")(iIntIndex)
    Next
    itxtSpreadIns = Join(itxtSpreadArr,"")
    
    If Trim(itxtSpreadIns + itxtSpreadUpd + itxtSpreadDel) <> "" Then
		Set iObjS5G121 = Server.CreateObject("pS5G121.cSDnDtlSvr")

		If CheckSYSTEMError(Err,True) Then
			Response.Write "<Script language=vbs> " & vbCr   
			Response.Write "Call parent.RemovedivTextArea " & vbCr   
			Response.Write "</Script> "	& vbCr          
			Response.End
		End If

		' Modified by Hwangseongbae (2003.03.25)
		Call iObjS5G121.Maintain(gStrGlobalCollection, "Y", iStrDnNo, "N", itxtSpreadIns, itxtSpreadUpd, itxtSpreadDel, iErrorPosition, "F")
    
		Set iObjS5G121 = Nothing
    
		If CheckSYSTEMError2(Err, True, iErrorPosition & "행","","","","") = True Then
			Response.Write "<SCRIPT LANGUAGE=VBSCRIPT> " & vbCr   
			Response.Write "Call parent.RemovedivTextArea " & vbCr   
			If iErrorPosition > 0 Then
				Response.Write " Call parent.ChangeTabs(parent.TAB2)" & vbCr
				Response.Write " Call Parent.SubSetErrPos(" & iErrorPosition & ")" & vbCr
			End If
			Response.Write "</SCRIPT> "		
			Response.End 
		End If
	End If
	
	Response.Write "<Script Language=vbscript>	" & vbcr
	Response.Write "parent.frm1.txtConDn_no.value = """ & ConvSPChars(iStrDnNo) & """" & vbcr
	Response.Write "parent.DbSaveOk	" & vbcr
	Response.Write "</Script>" & vbcr
	Response.End																				'☜: Process End

Case CStr(UID_M0003)														'☜: 삭제 요청 	    
	   
    If Request("txtDnNo") = "" Then										'⊙: 조회를 위한 값이 들어왔는지 체크 
		Call ServerMesgBox("삭제 조건값이 비어있습니다!", vbInformation, I_MKSCRIPT)
		Response.End 
	End If

	'================ 출하내역 삭제 ================
	Set iObjS5G121 = Server.CreateObject("PS5G121.cSDnDtlSvr")  
		
	If CheckSYSTEMError(Err,True) = True Then
	   Response.End       
	End If
	
	' Modified by Hwangseongbae (2003.03.25)
	Call iObjS5G121.Maintain(gStrGlobalCollection, "Y", Trim(Request("txtDnNo")), "Y")

    Set iObjS5G121 = Nothing
    
	If CheckSYSTEMError(Err,True) = True Then
	   Response.End       
	End If

	'================ 출하정보 삭제 ================
	iStrCUDFlag = "D"
	pvCB = "F"
	
	Redim iArrSDnHdr(1)
	iArrSDnHdr(0) = Trim(Request("txtDnNo"))
	iArrSDnHdr(1) = "Y"								' 예외출고여부 

    Set iObjS5G211 = Server.CreateObject("PS5G211.cSDnHdrSvr")
    
    If CheckSYSTEMError(Err,True) = True Then
	   Response.End       
    End If  
    
    Call iObjS5G211.Maintain(pvCB, gStrGlobalCollection, iStrCUDFlag, iArrSDnHdr)
	
	Set iObjS5G211 = Nothing
							 
	If CheckSYSTEMError(Err,True) = True Then
       Set iObjS5G211 = Nothing		                                                 '☜: Unload Comproxy DLL
	   Response.End       
    End If 

	'-----------------------
	'Result data display area
	'----------------------- 
	Response.Write "<Script Language=vbscript>"							& vbCr
	Response.Write "Call parent.DbDeleteOk()"							& vbCr
	Response.Write "</Script>"											& vbCr
	Response.End		 

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
	
    '-----------------------
    'Data manipulate area
    '-----------------------
	I1_b_item = Trim(Request("ItemCd"))	
	
    Set iPB3C104 = Server.CreateObject("PB3C104.cBLkUpItem")     
    
    If CheckSYSTEMError(Err,True) = True Then
       Response.End       
    End If
    
    Call iPB3C104.B_LOOK_UP_ITEM(gStrGlobalCollection, I1_b_item, , , , , prE1_b_item)	
	
	If CheckSYSTEMError(Err,True) = True Then
       Set iPB3C104 = Nothing	
       Response.End
    End If 
    Set iPB3C104 = Nothing	
%>

<SCRIPT Language="vbscript">
		With parent.frm1.vspdData
			.Row 	= "<%=Request("CRow")%>"
			.Col 	= parent.C_ItemNm
			.text	= "<%=ConvSPChars(prE1_b_item(prE1_item_nm))%>" 
			.Col 	= parent.C_Spec
			.text	= "<%=ConvSPChars(prE1_b_item(prE1_spec))%>"
			.Col 	= parent.C_DnUnit
			.text	= "<%=ConvSPChars(prE1_b_item(prE1_basic_unit))%>"

			.Col	= parent.C_VatType

			If Len(Trim(parent.frm1.txtVAT_Type.value)) Then
				.text	= parent.frm1.txtVAT_Type.value
			Else
				.text	= "<%=ConvSPChars(prE1_b_item(prE1_vat_type))%>"
			End If

			If parent.frm1.rdoVat_Inc_flag1.checked = True Then
				.Col	= parent.C_VatIncFlag
				.text	= parent.frm1.rdoVat_Inc_flag1.value
				.Col	= parent.C_VatIncFlagNm
				parent.frm1.vspdData.Text = "별도"
			Else
				.Col	= parent.C_VatIncFlag
				.text	= parent.frm1.rdoVat_Inc_flag2.value
				.Col	= parent.C_VatIncFlagNm
				parent.frm1.vspdData.Text = "포함"
			End If			
			
			Call parent.SetVatTypeForSpread(<%=Request("CRow")%>)
			
			Call parent.GetItemPrice(<%=Request("CRow")%>)
			Call parent.CalcAmtLoc(<%=Request("CRow")%>)
		End With			
</SCRIPT>

<%		
	Response.End

Case "LookUp"																'☜: 현재 주문처 거래 조회 요청을 받음 

    Err.Clear                                                               

	Dim pB5CS41
	Dim PB5GS45
    Dim I2_b_biz_partner
    Dim E1_Look_biz_partner
    Dim I1_b_biz_partner
    Dim E1_b_biz_partner
    Dim E2_b_biz_partner, E3_b_biz_partner, E4_b_biz_partner, E5_b_biz_partner, E6_b_biz_partner
	Dim I1_command
	                                  
    Const B132_E2_bp_cd = 0   'E2_b_biz_partner
    Const B132_E2_bp_nm = 1
    
    
    Const S074_E1_bp_cd = 0		'E1_Look_biz_partner
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
'12-24 코드 추가입력사항 종료----------------------------------------------------------
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
'12-24 네임 추가입력사항 시작----------------------------------------------------------
    Const S074_E1_card_co_cd_nm = 140                         '[카드사명]
    Const S074_E1_pay_meth_pur_nm = 141                       '[결재방법명(구매)]
    Const S074_E1_pay_type_pur_nm = 142                       '[입출금유형명(구매)]
    Const S074_E1_bank_cd_nm = 143                            '[은행 
    
    If Trim(Request("txtSold_to_party")) = "" Then								'⊙: 조회를 위한 값이 들어왔는지 체크 
		Call ServerMesgBox("주문처값이 비어있습니다!", vbInformation, I_MKSCRIPT)              
		Response.End 
	End If
    
    I1_command = "LOOKUP"
    I2_b_biz_partner = Trim(Request("txtSold_to_party"))
    
    Set pB5CS41 = Server.CreateObject("PB5CS41.cLookupBizPartnerSvr")
    
	If CheckSYSTEMError(Err,True) = True Then
       Response.End       
    End If 
   
    Call pB5CS41.B_LOOKUP_BIZ_PARTNER_SVR(gStrGlobalCollection, I1_command, I2_b_biz_partner, _
                              E1_Look_biz_partner)
    
    If CheckSYSTEMError(Err,True) = True Then
		Set pB5CS41 = Nothing
		Response.End
    End If

    Set pB5CS41 = Nothing
   
   	I1_b_biz_partner = Trim(Request("txtSold_to_party"))
	Set PB5GS45 = Server.CreateObject("PB5GS45.cBListDftBpFtnSvr")    

    If CheckSYSTEMError(Err,True) = True Then
       Response.End       
    End If 
   
    Call PB5GS45.B_LIST_DEFAULT_BP_FTN_SVR(gStrGlobalCollection, I1_b_biz_partner, E1_b_biz_partner, _
                              E2_b_biz_partner, E3_b_biz_partner, E4_b_biz_partner, _
                              E5_b_biz_partner, E6_b_biz_partner)
    
    If CheckSYSTEMError(Err,True) = True Then
		Set PB5GS45 = Nothing
		Response.End
    End If

    Set PB5GS45 = Nothing

	'-----------------------
	'Result data display area
	'----------------------- 
%>
<SCRIPT Language=vbscript>
	With parent.frm1

		'주문처 
		.txtSold_to_party.value		= "<%=ConvSPChars(E1_Look_biz_partner(S074_E1_bp_cd))%>"
		.txtSold_to_partyNm.value	= "<%=ConvSPChars(E1_Look_biz_partner(S074_E1_bp_nm))%>"
		'거래유형 
		.txtDeal_Type.value			= "<%=ConvSPChars(E1_Look_biz_partner(S074_E1_deal_type))%>"
		.txtDeal_Type_nm.value		= "<%=ConvSPChars(E1_Look_biz_partner(S074_E1_deal_type_nm))%>"				
		'결제방법 
		.txtPay_terms.value			= "<%=ConvSPChars(E1_Look_biz_partner(S074_E1_pay_meth))%>"
		.txtPay_terms_nm.value		= "<%=ConvSPChars(E1_Look_biz_partner(S074_E1_pay_meth_nm))%>"
		'부가세유형 
		.txtVat_Type.value			= "<%=ConvSPChars(E1_Look_biz_partner(S074_E1_vat_type))%>"
		.txtVatTypeNm.value			= "<%=ConvSPChars(E1_Look_biz_partner(S074_E1_vat_type_nm))%>"
		'부가세율 
		.txtVat_rate.text			= "<%=UNINumClientFormat(E1_Look_biz_partner(S074_E1_vat_rate),ggExchRate.DecPoint,0)%>"
		'운송방법 
		.txtTrans_Meth.value		= "<%=ConvSPChars(E1_Look_biz_partner(S074_E1_trans_meth))%>"
		.txtTrans_Meth_nm.value		= "<%=ConvSPChars(E1_Look_biz_partner(S074_E1_trans_meth_nm))%>"
		'영업그룹 
		.txtSales_Grp.value			= "<%=ConvSPChars(E1_Look_biz_partner(S074_E1_biz_grp))%>"
		.txtSales_GrpNm.value		= "<%=ConvSPChars(E1_Look_biz_partner(S074_E1_b_sales_grp_nm))%>"
		'부가세구분 
		If "<%=ConvSPChars(E1_Look_biz_partner(S074_E1_vat_inc_flag))%>" = "1" Then
			.rdoVat_Inc_flag1.checked = True
		Else
			.rdoVat_Inc_flag2.checked = True
		End If
		
		.txtShip_to_party.value		= "<%=ConvSPChars(E1_b_biz_partner(B132_E2_bp_cd))%>"
		.txtShip_to_partyNm.value	= "<%=ConvSPChars(E1_b_biz_partner(B132_E2_bp_nm))%>"
		<% '납품처 국가코드 %>
		If Len(Trim("<%=ConvSPChars(E1_b_biz_partner(B132_E2_bp_cd))%>")) Then
			Call parent.GetContryCd
		End If 
	End With	
	
	If Len(Trim("<%=E1_Look_biz_partner(S074_E1_vat_type)%>")) Then
		Call parent.SetVatTypeForHdr
	End If	
	<% '세금신고사업장 Fetch %>
	Call parent.GetTaxBizArea("BP")										

	parent.lgBlnFlgChgValue = true

	Call Parent.SoldToPartyLookUpOK ' 박정순 추가 (2006-05-26) 
</SCRIPT>
<% 
    
	Response.End																				'☜: Process End

Case CStr("ARPOST")

    Err.Clear																		

	Dim pS5G115
	Dim i1_s_bill_hdr
	Dim i3_s_dn_hdr
		
	Const S427_I1_bill_no = 0    '[CONVERSION INFORMATION]  View Name : imp s_bill_hdr
    
    Const S427_I3_dn_no = 0    '[CONVERSION INFORMATION]  View Name : imp s_dn_hdr
    Const S427_I3_actual_gi_dt = 1
    Const S427_I3_ar_flag = 2
    Const S427_I3_vat_flag = 3
    Const S427_I3_except_dn_flag = 4
    
	Redim i1_s_bill_hdr(S427_I1_bill_no)
	Redim i3_s_dn_hdr(S427_I3_except_dn_flag)
	
    i3_s_dn_hdr(S427_I3_dn_no) = Trim(Request("txtHDnNo"))
	i3_s_dn_hdr(S427_I3_actual_gi_dt) = UNIConvDate(Request("txtActualGIDt"))	
	i3_s_dn_hdr(S427_I3_ar_flag) = Request("txtARFlag")
	i3_s_dn_hdr(S427_I3_vat_flag) = Request("txtVatFlag")
	i3_s_dn_hdr(S427_I3_except_dn_flag)  = "Y"	
	
	If Trim(Request("txtGINo")) = "" Then
		iCommand = "POST"
	Else
		iCommand = "CANCEL"
	End If
	
	pvCB = "F"
	
	Set pS5G115 = Server.CreateObject("PS5G115.cSPostGISvr")
	
	If CheckSYSTEMError(Err,True) = True Then
       Response.End       
    End If
    
    Call pS5G115.S_POST_GOODS_ISSUE_SVR(pvCB, gStrGlobalCollection, iCommand, i1_s_bill_hdr, i3_s_dn_hdr)

    If CheckSYSTEMError(Err,True) = True Then
		Set pS41115 = Nothing                              '☜: Unload Comproxy
		Response.End 
	End If
	
	Set pS41115 = Nothing                              '☜: Unload Comproxy
	
	Response.Write "<SCRIPT Language=vbscript>" & vbcr
	Response.Write "parent.DbSaveOk()" & vbcr
	Response.Write "</SCRIPT>" & vbcr
	Response.End																		'☜: Process End

Case "ChkGiCreditLimit"															'☜: 여신한도 체크 

    Err.Clear																		

	Dim objPS3G113	
	Dim iArrData
	Dim iDblOverLimitAmt
	
	Redim iArrData(1)
    
    iArrData(0) = "EG"
    iArrData(1) = Trim(Request("txtHDnNo"))

	
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
End Select
'===========================================
' 사용자 정의 서버 함수 
'===========================================
%>
