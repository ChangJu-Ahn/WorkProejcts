<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<%
'**********************************************************************************************
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : 고정자산관리
'*  3. Program ID           : a7102mb2
'*  4. Program Name         : 고정자산취득내역등록
'*  5. Program Desc         : 고정자산별 취득내역을 등록,수정
'*  6. Comproxy List        : +As0021ManageSvr
'                             +B1a028ListMinorCode
'*  7. Modified date(First) : 2001/05/24
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 김희정
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
Response.Expires = -1								'☜ : ASP가 캐쉬되지 않도록 한다.
Response.Buffer = True								'☜ : ASP가 버퍼에 저장되어 마지막에 바로 Client에 내려간다.

'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다
On Error Resume Next														'☜: 
Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide


    Call LoadBasisGlobalInf()
	'Dim lgCurrency, lgStrPrevKey_i, lgBlnFlgChgValue, plgStrPrevKey_i
	Call LoadInfTB19029B("I","*", "NOCOOKIE", "MB")
	'Call LoadBNumericFormat("I", "*")
'	Dim gChangeOrgId 'iChangeOrgId
	gChangeOrgId = request("hOrgChangeId")
	
	
'-------------------------
' 변수, 상수 선언 
'-------------------------
	Dim iPAAG010																'☆ : 저장용 ComProxy Dll 사용 변수
	Dim lgIntFlgMode

	lgIntFlgMode = CInt(Request("txtFlgMode"))									'☜: 저장시 Create/Update 판별
		
	'Dim IntRows
	'Dim IntCols
	'Dim vbIntRet
	'Dim lEndRow
	'Dim boolCheck
	'Dim lgIntFlgMode
	'Dim LngMaxRow_m
	'Dim LngMaxRow_i

    '[Import 변수]
	Dim I1_a_acct_trans_type	'거래유형
	Dim I2_b_currency			'거래통화
	Dim I3_a_asset_acq			'Control Data
	Dim I4_ief_supplied			'구분키(C,U)
	Dim IG1_import_mst_grp		'Master Data
	Dim IG2_import_itm_grp		'출금내역
	Dim I5_b_biz_partner		'거래처
	Dim I6_b_acct_dept			'부서정보
	Dim I7_a_asset_acq			'채무만기일자(미지급금)
	Dim I8_a_batch				'전표일자
	Dim E3_a_asset_acq

	'[Import 상수]
    Const A504_I1_trans_type = 0	'import a_acct_trans_type

    Const A504_I2_currency = 0		'import b_currency

    Const A504_I3_acq_no = 0		'import a_asset_acq
    Const A504_I3_acq_dt = 1
    Const A504_I3_acq_fg = 2
    Const A504_I3_doc_cur = 3
    Const A504_I3_xch_rate = 4
    Const A504_I3_tot_acq_amt = 5
    Const A504_I3_tot_acq_loc_amt = 6
    Const A504_I3_extra_acq_amt = 7
    Const A504_I3_extra_acq_loc_amt = 8
    Const A504_I3_vat_type = 9
    Const A504_I3_vat_make_fg = 10
    Const A504_I3_vat_amt = 11
    Const A504_I3_vat_loc_amt = 12
    Const A504_I3_ref_no = 13
    '20030301	미지급금계정 추가
    Const A504_I3_ap_acct_cd = 14
    Const A504_I3_ap_due_dt = 15
    Const A504_I3_ap_amt = 16
    Const A504_I3_ap_loc_amt = 17
    Const A504_I3_acq_desc = 18
    Const A504_I3_ap_no = 19
    Const A504_I3_gl_no = 20
    Const A504_I3_temp_gl_no = 21
    Const A504_I3_internal_cd = 22
    Const A504_I3_insrt_user_id = 23
    Const A504_I3_insrt_dt = 24
    Const A504_I3_updt_user_id = 25
    Const A504_I3_updt_dt = 26
    Const A504_I3_vat_io_fg = 27
    Const A504_I3_vat_rate = 28
    Const A504_I3_issued_dt = 29
    Const A504_I3_tax_biz_area_cd = 30
    Const A504_I3_Credit_card_No = 31															'신용카드 번호 추가
   

    Const A504_I4_select_char = 0	'import_mode_fg ief_supplied

    Const A504_I5_bp_cd = 0			'import b_biz_partner

    Const A504_I6_org_change_id = 0	'import b_acct_dept
    Const A504_I6_dept_cd = 1

    Const A504_I7_ap_due_dt = 0		'import_null_dt a_asset_acq

    Const A504_I8_gl_dt = 0			'import_a_batch a_batch

    '[IMPORTS Group 상수]	'20080305 주석처리 air
    'Group Name : import_mst_grp
    'Const A504_IG1_I1_acct_cd = 0        'View Name : import_mst_itm a_acct
    'Const A504_IG1_I2_select_char = 1    'View Name : import_mst_itm ief_supplied
    'Const A504_IG1_I3_org_change_id = 2  'View Name : import_mst_itm b_acct_dept
    'Const A504_IG1_I3_dept_cd = 3
    'Const A504_IG1_I4_asst_no = 4        'View Name : import_mst_itm a_asset_master
    'Const A504_IG1_I4_asst_nm = 5
    'Const A504_IG1_I4_reg_dt = 6
    'Const A504_IG1_I4_spec = 7
    'Const A504_IG1_I4_doc_cur = 8
    'Const A504_IG1_I4_xch_rate = 9
    'Const A504_IG1_I4_acq_amt = 10
    'Const A504_IG1_I4_acq_loc_amt = 11
    'Const A504_IG1_I4_acq_qty = 12
    'Const A504_IG1_I4_inv_qty = 13
    'Const A504_IG1_I4_asset_desc = 14
    'Const A504_IG1_I4_ref_no = 15
    'Const A504_IG1_I4_tax_dur_yrs = 16
    'Const A504_IG1_I4_cas_dur_yrs = 17
    'Const A504_IG1_I4_tax_end_l_term_cpt_tot_amt = 18
    'Const A504_IG1_I4_cas_end_l_term_cpt_tot_amt = 19
    'Const A504_IG1_I4_tax_end_l_term_depr_tot_amt = 20
    'Const A504_IG1_I4_cas_end_l_term_depr_tot_amt = 21
    'Const A504_IG1_I4_tax_end_l_term_bal_amt = 22
    'Const A504_IG1_I4_cas_end_l_term_bal_amt = 23
    'Const A504_IG1_I4_tax_depr_sts = 24
    'Const A504_IG1_I4_cas_depr_sts = 25
    'Const A504_IG1_I4_tax_depr_end_yyyymm = 26
    'Const A504_IG1_I4_cas_depr_end_yyyymm = 27
    'Const A504_IG1_I4_start_depr_yymm = 28
    'Const A504_IG1_I4_tax_dur_mnth = 29
    'Const A504_IG1_I4_cas_dur_mnth = 30
    '
    ''Group Name : import_itm_grp
    'Const A504_IG2_I1_select_char = 0    'View Name : import_itm_item ief_supplied
    'Const A504_IG2_I2_acq_seq = 1        'View Name : import_itm_item a_asset_acq_item
    'Const A504_IG2_I2_paym_type = 2
    'Const A504_IG2_I2_paym_amt = 3
    'Const A504_IG2_I2_paym_loc_amt = 4
    'Const A504_IG2_I2_note_no = 5
    'Const A504_IG2_I3_bank_acct_no = 6   'View Name : import_itm_item b_bank_acct
	
	'Export
	Const A073_E3_acq_no = 0             'View Name : export a_asset_acq
'-------------------------   
' Data Matching
'------------------------- 
	Redim I1_a_acct_trans_type(0)
	'Redim I2_b_currency(0)
	'20030301	미지급금계정 추가	
	Redim I3_a_asset_acq(31)
	Redim I4_ief_supplied(0)
	Redim I5_b_biz_partner(0)
	Redim I6_b_acct_dept(1)	
	Redim I7_a_asset_acq(0)
	Redim I8_a_batch(0)
	Redim E3_a_asset_acq(0)


	' -- 권한관리추가
	Const A504_I9_a_data_auth_data_BizAreaCd = 0
	Const A504_I9_a_data_auth_data_internal_cd = 1
	Const A504_I9_a_data_auth_data_sub_internal_cd = 2
	Const A504_I9_a_data_auth_data_auth_usr_id = 3

	Dim I9_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동

  	Redim I9_a_data_auth(3)
	I9_a_data_auth(A504_I9_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I9_a_data_auth(A504_I9_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I9_a_data_auth(A504_I9_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I9_a_data_auth(A504_I9_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))

	
	I1_a_acct_trans_type(A504_I1_trans_type)  = "AS001"		'거래유형
	
	I2_b_currency = gCurrency	'거래통화	'(A504_I2_currency)

	'Control Data
    I3_a_asset_acq(A504_I3_acq_no)			  = Request("txtAcqNo")						'00 취득번호
    I3_a_asset_acq(A504_I3_acq_dt)			  = UNIConvDate(Request("txtAcqDt"))		'01 취득일자
    I3_a_asset_acq(A504_I3_acq_fg)			  = Request("cboAcqFg")						'02 취득구분
    I3_a_asset_acq(A504_I3_doc_cur)			  = UCase(Trim(Request("txtDocCur")))				'03 거래통화
    I3_a_asset_acq(A504_I3_xch_rate)		  = UNIConvNum(Request("txtXchRate"),0)		'04 환율
    I3_a_asset_acq(A504_I3_tot_acq_amt)       = UNIConvNum(Request("txtAcqAmt"),0)		'05 총취득금액
    I3_a_asset_acq(A504_I3_tot_acq_loc_amt)   = UNIConvNum(Request("txtAcqLocAmt"),0)	'06 총취득금액(자국)
   'I3_a_asset_acq(A504_I3_extra_acq_amt)     											'07 부대비용
   'I3_a_asset_acq(A504_I3_extra_acq_loc_amt) 											'08 부대비용(자국)
    I3_a_asset_acq(A504_I3_vat_type)		  = UCase(Request("txtVatType"))					'09 부가세유형
   'I3_a_asset_acq(A504_I3_vat_make_fg)													'10 부가세 생성여부
    I3_a_asset_acq(A504_I3_vat_amt)			  = UNIConvNum(Request("txtVatAmt"),0)		'11 부가세금액
    I3_a_asset_acq(A504_I3_vat_loc_amt)		  = UNIConvNum(Request("txtVatLocAmt"),0)	'12 부가세금액(자국)
   'I3_a_asset_acq(A504_I3_ref_no)														'13 참조번호(Master Spread)
   '20030301	미지급금계정 추가
    I3_a_asset_acq(A504_I3_ap_acct_cd)		  = Trim(UCase(Request("txtApAcctCd")))		'14 미지급금 계정
    I3_a_asset_acq(A504_I3_ap_due_dt)		  = UNIConvDate(Request("txtApDueDt"))		'14 미지급금 만기일자
    I3_a_asset_acq(A504_I3_ap_amt)			  = UNIConvNum(Request("txtApAmt"),0)		'15 미지급금액
    I3_a_asset_acq(A504_I3_ap_loc_amt)		  = UNIConvNum(Request("txtApLocAmt"),0)	'16 미지급금액(자국)
    I3_a_asset_acq(A504_I3_acq_desc)		  = Trim(Request("txtDesc"))					'17 적요(Master Spread)
    I3_a_asset_acq(A504_I3_ap_no)			  = Trim(Request("txtApNo"))				'18 미지급금 번호
    I3_a_asset_acq(A504_I3_gl_no)			  = Trim(Request("txtGLNo"))				'19 회계전표번호
    I3_a_asset_acq(A504_I3_temp_gl_no)		  = Trim(Request("txtTempGLNo"))			'20 결의전표번호
   'I3_a_asset_acq(A504_I3_internal_cd)		  											'21 내부부서코드
   'I3_a_asset_acq(A504_I3_insrt_user_id)	  											'22 입력자
   'I3_a_asset_acq(A504_I3_insrt_dt)		  											'23 입력일
   'I3_a_asset_acq(A504_I3_updt_user_id)	  											'24 수정자
   'I3_a_asset_acq(A504_I3_updt_dt)			  											'25 수정일
    I3_a_asset_acq(A504_I3_vat_io_fg)		 = "I" 											'26 부가세 매입/매출 구분
    I3_a_asset_acq(A504_I3_vat_rate)		  = UNIConvNum(Request("txtVatRate"),0)		'28 부가세율
	
	If Trim(Request("txtIssuedDt")) <>"" Then
		I3_a_asset_acq(A504_I3_issued_dt)	=	UNIConvDate(Request("txtIssuedDt"))     ' 10월 정기 패치 추가
	End If																				'29 계산서 발행일
	
	I3_a_asset_acq(A504_I3_tax_biz_area_cd)	=	Trim(Request("txtReportAreaCd"))		'30 신고사업장
	I3_a_asset_acq(A504_I3_Credit_card_No)	=	Trim(Request("txtCardNo"))										'31 신용카드 번호 (수정)

    '구분키(C,U)
	If lgIntFlgMode = OPMD_CMODE Then
		I4_ief_supplied(A504_I4_select_char) = "C"
	ElseIf lgIntFlgMode = OPMD_UMODE Then
		I4_ief_supplied(A504_I4_select_char) = "U"
	End If
    
	IG1_import_mst_grp = Request("txtSpread_m")		'취득상세내역 Spread
	IG2_import_itm_grp = Request("txtSpread_i")		'출금내역 Spread

	I5_b_biz_partner(A504_I5_bp_cd) = Trim(Request("txtBpCd"))	'거래처
	
	'부서정보
	I6_b_acct_dept(A504_I6_org_change_id) = gChangeOrgId
	I6_b_acct_dept(A504_I6_dept_cd) = Trim(Request("txtDeptCd"))	'취득부서
	
	'채무만기일자(미지급금)
	I7_a_asset_acq(A504_I7_ap_due_dt) = UNIConvDate(Request("txtApDueDt"))
	
	'전표일자
	I8_a_batch(A504_I8_gl_dt) = UNIConvDate(Request("txtGLDt"))				

'-------------------------   
' 업무 처리 
'-------------------------    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear 

    Set iPAAG010 = Server.CreateObject("PAAG010_KO441.cAAcqMngSvr")	'AIR :: PAAG010 -> PAAG010_KO441


    If CheckSYSTEMError(Err, True) = True Then					
       Response.End
    End If  
    

	call iPAAG010.AS0021_ACQ_MANAGE_SVR(gStrGlobalCollection, _
										I1_a_acct_trans_type, _
										I2_b_currency, _
										I3_a_asset_acq, _							
										I4_ief_supplied, _
										IG1_import_mst_grp, _
										IG2_import_itm_grp, _
										I5_b_biz_partner, _	
										I6_b_acct_dept, _
										I7_a_asset_acq, _		
										I8_a_batch, _
										E1_a_asset_master, _
										E3_a_asset_acq, _
										I9_a_data_auth)

    If CheckSYSTEMError(Err, True) = True Then					
       Set iPAAG010 = Nothing
       Response.End
    End If    

    Set iPAAG010 = Nothing
    
	Response.Write " <Script Language=vbscript> " & vbCr
	Response.Write "	parent.frm1.txtAcqNo.value = """ & E3_a_asset_acq(A073_E3_acq_no) & """" & vbCr '자산취득번호        
    Response.Write "	parent.DbSaveOk()												   	   " & vbCr
    Response.Write " </Script>					" & vbCr
%>
