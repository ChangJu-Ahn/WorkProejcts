<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%
'**********************************************************************************************
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : 
'*  3. Program ID           : a7103mb1
'*  4. Program Name         : 고정자산 MASTER 수정
'*  5. Program Desc         : 고정자산별 MASTER를 수정, 조회
'*  6. Comproxy List        : +As0041ManageSvr
'                             +As0049LookupSvr
'                             +B1a028ListMinorCode
'*  7. Modified date(First) : 2000/03/29
'*  8. Modified date(Last)  : 2000/09/14
'*  9. Modifier (First)     : 조익성
'* 10. Modifier (Last)      : hersheys
'* 11. Comment              :
'**********************************************************************************************
Response.Expires = -1								'☜ : ASP가 캐쉬되지 않도록 한다.
Response.Buffer = True								'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd       
    Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I","*","NOCOOKIE","MB")
	Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")                                                        '☜: Hide Processing message

	gChangeOrgId = GetGlobalInf("gChangeOrgId")

	Dim lgOpModeCRUD
    '---------------------------------------Common-----------------------------------------------------------
'    lgErrorStatus     = "NO"
'    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)

	'------ Developer Coding part (Start ) ------------------------------------------------------------------


	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSave()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select

	Response.End 
    'Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()

	'------------------
	' 변수, 상수 선언 
	'------------------
	Dim iPAAG015
	Dim I1_a_asset_master_asst_no
	Dim I2_b_acct_dept_org_change_id
	Dim E1_a_batch							
	Dim E2_a_asset_master
	Dim E3_a_mcs
	Dim E4_a_acct		
	Dim E5_a_asset_acct
	Dim E6_a_asset_depr_rate	
	Dim E7_a_asset_depr_rate	
	Dim E8_a_mcs
	Dim E9_b_biz_partner
	Dim E10_b_acct_dept
	Dim E11_a_asset_acq
	Dim E12_b_cost_center	'코스트센터(반환) : air

    'export a_batch						[전표일괄생성]
    Const A512_E1_gl_dt = 0		'회계일자
    
    'export a_asset_master				[자산마스터]
    Const A512_E2_asst_no = 0			'자산번호
    Const A512_E2_asst_nm = 1			'자산명
    Const A512_E2_ref_no = 2			'품목그룹
    Const A512_E2_reg_dt = 3			'자산등록일자
    Const A512_E2_acq_amt = 4			'취득금액
    Const A512_E2_doc_cur = 5			'거래통화
    Const A512_E2_xch_rate = 6			'환율
    Const A512_E2_acq_loc_amt = 7		'취득금액(자국)
    Const A512_E2_acq_qty = 8			'취득수량
    Const A512_E2_inv_qty = 9			'재고수량
    Const A512_E2_tax_dur_yrs = 10		'세법기준내용년수
    Const A512_E2_cas_dur_yrs = 11		'기업회계기준내용년수
    Const A512_E2_tax_end_l_term_cpt_tot_amt = 12	'세법기준전기말자본적지출금액
    Const A512_E2_cas_end_l_term_cpt_tot_amt = 13	'기업회계기준전기말자본적지출금액
    Const A512_E2_tax_end_l_term_depr_tot_amt = 14	'세법기준전기말감가상각누계금액
    Const A512_E2_cas_end_l_term_depr_tot_amt = 15  '기업회계기준전기말감가상각누계금액
    Const A512_E2_tax_end_l_term_bal_amt = 16		'세법기준미상각잔액
    Const A512_E2_cas_end_l_term_bal_amt = 17		'기업회계기준미상각잔액
    Const A512_E2_tax_depr_end_yyyymm = 18			'세법기준상각완료년월
    Const A512_E2_cas_depr_end_yyyymm = 19			'기업회계기준상각완료년월
    Const A512_E2_tax_depr_sts = 20		'세법기준상각상태
    Const A512_E2_cas_depr_sts = 21		'기업회계기준상각상태
    Const A512_E2_spec = 22				'용도/규격
    Const A512_E2_asset_desc = 23		'적요
    Const A512_E2_start_depr_yymm = 24	'감가상각시작년월
    Const A512_E2_gl_no = 25			'전표번호
    Const A512_E2_temp_gl_no = 26		'결의전표번호
    Const A512_E2_temp_fg1 = 27			'잔존율표시
    Const A512_E2_disuse_fg = 28		'매각/폐기완료
    Const A512_E2_disuse_yymm = 29		'매각/폐기완료년월
    Const A512_E2_vat_rate = 30			'부가세율
    Const A512_E2_net_amt = 31			'공급가액
    Const A512_E2_net_loc_amt = 32		'공급가액(자국)
    Const A512_E2_tax_dur_mnth = 33		'세법기준내용월수
    Const A512_E2_cas_dur_mnth = 34		'기업회계기준기준내용월수
	
	'export_start_yymm a_mcs		[a_mcs]
	Const A512_E3_txt_from_dt = 0   'txt_from_dt
	
    'export a_acct					[계정코드]
    Const A512_E4_acct_cd = 0		'계정코드    
    Const A512_E4_acct_nm = 1		'계정단명

    'export a_asset_acct			[자산계정코드]
    Const A512_E5_depr_mthd = 0		'상각방법
    Const A512_E5_dur_yrs = 1		'내용년수

    'export_tax a_asset_depr_rate   [감가상각률정보]
    Const A512_E6_depr_rate = 0     '감가상각률(세법기준)

    'export_cas a_asset_depr_rate	[감가상각률정보]
    Const A512_E7_depr_rate = 0     '감가상각률(기업회계기준)

    'export a_mcs
    Const A512_E8_txt_from_dt = 0    
    Const A512_E8_txt_to_dt = 1

    'export b_biz_partner			[거래처]
    Const A512_E9_bp_cd = 0			'거래처코드          
    Const A512_E9_bp_nm = 1			'거래처(약명)

    'export b_acct_dept				[회계부서정보]
    Const A512_E10_dept_cd = 0		'부서코드
    Const A512_E10_dept_nm = 1		'부서약명

    'export a_asset_acq				[자산취득]
    Const A512_E11_acq_fg = 0       '취득구분
    Const A512_E11_acq_no = 1		'자산취득번호
    Const A512_E11_gl_no = 2		'전표번호
    Const A512_E11_temp_gl_no = 3	'결의전표번호

    'export B_COST_CENTER			[코스트센터]
    Const A512_E12_cost_cd = 0       'cost code
    Const A512_E12_cost_nm = 1       'cost name


	' -- 권한관리추가
	Const A512_I3_a_data_auth_data_BizAreaCd = 0
	Const A512_I3_a_data_auth_data_internal_cd = 1
	Const A512_I3_a_data_auth_data_sub_internal_cd = 2
	Const A512_I3_a_data_auth_data_auth_usr_id = 3

	Dim I3_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동

  	Redim I3_a_data_auth(3)
	I3_a_data_auth(A512_I3_a_data_auth_data_BizAreaCd)       = Trim(Request("lgAuthBizAreaCd"))
	I3_a_data_auth(A512_I3_a_data_auth_data_internal_cd)     = Trim(Request("lgInternalCd"))
	I3_a_data_auth(A512_I3_a_data_auth_data_sub_internal_cd) = Trim(Request("lgSubInternalCd"))
	I3_a_data_auth(A512_I3_a_data_auth_data_auth_usr_id)     = Trim(Request("lgAuthUsrID"))

	'------------------
	' Data Matching
	'------------------
	'ReDim E2_a_asset_master(34)
	'ReDim E4_a_acct(1)
	'ReDim E5_a_asset_acct(1)
	'ReDim E8_a_mcs(1)
	'ReDim E9_b_biz_partner(1)
	'ReDim E10_b_acct_dept(1)
	'ReDim E11_a_asset_acq(3)
	
	'*** Import Data ***
	I1_a_asset_master_asst_no = Request("txtCondAsstNo")	'자산번호
	I2_b_acct_dept_org_change_id = gChangeOrgId				'조직변경ID

	'------------------
	' 요청 변수 처리 
	'------------------
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	Set iPAAG015 = Server.CreateObject("PAAG015_KO441.cAAS0049LkUpSvr")

    If CheckSYSTEMError(Err, True) = True Then					
       Exit Sub
    End If    

	Call iPAAG015.AS0049_LOOKUP_SVR(gStrGloBalCollection, _ 
									I1_a_asset_master_asst_no, _
									I2_b_acct_dept_org_change_id, _
									E1_a_batch, _
									E2_a_asset_master, _
									E3_a_mcs, _
									E4_a_acct, _
									E5_a_asset_acct, _
									E6_a_asset_depr_rate, _
									E7_a_asset_depr_rate, _
									E8_a_mcs, _
									E9_b_biz_partner, _
									E10_b_acct_dept, _
									E11_a_asset_acq, _
									E12_b_cost_center, _
									I3_a_data_auth)

    If CheckSYSTEMError(Err, True) = True Then					
       Set iPAAG015 = Nothing
       'Response.End
       Exit Sub
    End If    
    Set iPAAG015 = Nothing

	'---------------------
	' HTML 및 결과 생성부 
	'---------------------
	Response.Write " <Script Language=vbscript>				" & vbCr
	Response.Write " With parent						    " & vbCr

    '*** 기본정보 ***
    Response.Write "	.frm1.txtAsstNm.value    = """ & ConvSPChars(E2_a_asset_master(A512_E2_asst_nm))		& """" & vbCr '자산명
    Response.Write "	.frm1.txtRefNo.value     = """ & ConvSPChars(E2_a_asset_master(A512_E2_ref_no))			& """" & vbCr '품목그룹            
    Response.Write "	.frm1.txtDeptCd.value    = """ & ConvSPChars(E10_b_acct_dept(A512_E10_dept_cd))			& """" & vbCr '관리부서코드            
    Response.Write "	.frm1.txtDeptNm.value    = """ & ConvSPChars(E10_b_acct_dept(A512_E10_dept_nm))			& """" & vbCr '관리부서명
    
    Response.Write "	.frm1.txtCostCd.value    = """ & ConvSPChars(E12_b_cost_center(A512_E12_cost_cd))			& """" & vbCr '코스트센터코드            
    Response.Write "	.frm1.txtCostNm.value    = """ & ConvSPChars(E12_b_cost_center(A512_E12_cost_nm))			& """" & vbCr '코스트센터명
    
    Response.Write "	.frm1.txtRegDt.text     = """ & UNIDateClientFormat(E2_a_asset_master(A512_E2_reg_dt))	& """" & vbCr '취득일자(자산등록일자)
    Response.Write "	.frm1.txtDocCur.value    = """ & ConvSPChars(E2_a_asset_master(A512_E2_doc_cur))		& """" & vbCr '거래통화
    Response.Write "	.frm1.txtXchRate.value   = """ & UNINumClientFormat(E2_a_asset_master(A512_E2_xch_rate),   ggExchRate.DecPoint, 0)		& """" & vbCr '환율            
    Response.Write "	.frm1.txtAcqAmt.value    = """ & UNINumClientFormat(E2_a_asset_master(A512_E2_acq_amt),    ggAmtOfMoney.DecPoint, 0)	& """" & vbCr '취득금액            
    Response.Write "	.frm1.txtAcqLocAmt.value = """ & UNINumClientFormat(E2_a_asset_master(A512_E2_acq_loc_amt), ggAmtOfMoney.DecPoint, 0)	& """" & vbCr '취득금액(자국)            
    Response.Write "	.frm1.txtAcqQty.value    = """ & UNINumClientFormat(E2_a_asset_master(A512_E2_acq_qty),    ggQty.DecPoint, 0)			& """" & vbCr '취득수량            
    Response.Write "	.frm1.txtInvQty.value    = """ & UNINumClientFormat(E2_a_asset_master(A512_E2_inv_qty),    ggQty.DecPoint, 0)			& """" & vbCr '재고수량            
    Response.Write "	.frm1.txtAcctCd.value    = """ & ConvSPChars(E4_a_acct(A512_E4_acct_cd))				& """" & vbCr '계정코드            
    Response.Write "	.frm1.txtAcctNm.value    = """ & ConvSPChars(E4_a_acct(A512_E4_acct_nm))				& """" & vbCr '계정코드명            
    Response.Write "	.frm1.txtBpCd.value      = """ & ConvSPChars(E9_b_biz_partner(A512_E9_bp_cd))			& """" & vbCr '구입거래처코드            
    Response.Write "	.frm1.txtBpNm.value      = """ & ConvSPChars(E9_b_biz_partner(A512_E9_bp_nm))			& """" & vbCr '구입거래처명            
    Response.Write "	.frm1.cboAcqFg.value     = """ & E11_a_asset_acq(A512_E11_acq_fg)						& """" & vbCr '취득구분            
    Response.Write "	.frm1.txtSpec.value      = """ & Trim(ConvSPChars(E2_a_asset_master(A512_E2_spec)))			& """" & vbCr '구조/용도/크기            
    Response.Write "	.frm1.txtDesc.value      = """ & Trim(ConvSPChars(E2_a_asset_master(A512_E2_asset_desc)))	& """" & vbCr '적요            
    Response.Write "	.frm1.txtDeprFrdt.text  = """ & UNIMonthClientFormat(E2_a_asset_master(A512_E2_start_depr_yymm))						& """" & vbCr '감가상각시작년월            
    
    '*** 전기말 상각내역:세법기준(자국) ***
	'Response.Write "	.frm1.txtTaxDurYrs.value     = """ & E2_a_asset_master(A512_E2_tax_dur_yrs)				& """" & vbCr '내용연수            
	Response.Write "	.frm1.txtTaxDurYrs.value     = """ & E2_a_asset_master(A512_E2_tax_dur_mnth)			& """" & vbCr '내용월수	>>air            
	Response.Write "	.frm1.txtTaxDeprRate.value   = """ & UNINumClientFormat(E6_a_asset_depr_rate(A512_E6_depr_rate), ggExchRate.DecPoint, 0)				  & """" & vbCr '상각율            
	Response.Write "	.frm1.txtTaxDeprTotAmt.value = """ & UNINumClientFormat(E2_a_asset_master(A512_E2_tax_end_l_term_depr_tot_amt), ggAmtOfMoney.DecPoint, 0) & """" & vbCr '상각누계
	Response.Write "	.frm1.txtTaxCptTotAmt.value  = """ & UNINumClientFormat(E2_a_asset_master(A512_E2_tax_end_l_term_cpt_tot_amt), ggAmtOfMoney.DecPoint, 0)  & """" & vbCr '자본적지출누계            
	Response.Write "	.frm1.txtTaxBalAmt.value     = """ & UNINumClientFormat(E2_a_asset_master(A512_E2_tax_end_l_term_bal_amt), ggAmtOfMoney.DecPoint, 0)	  & """" & vbCr '미상각잔액            
	Response.Write "	.frm1.cboTaxDeprSts.value    = """ & ConvSPChars(E2_a_asset_master(A512_E2_tax_depr_sts))								& """" & vbCr '상각상태            
	Response.Write "	.frm1.txtTaxDeprEnd.text    = """ & UNIMonthClientFormat(E2_a_asset_master(A512_E2_tax_depr_end_yyyymm))				& """" & vbCr '상각완료년월            
	
    '*** 전기말 상각내역:기업회계기준(자국) ***
	'Response.Write "	.frm1.txtCasDurYrs.value     = """ & E2_a_asset_master(A512_E2_cas_dur_yrs)				& """" & vbCr '내용연수            
	Response.Write "	.frm1.txtCasDurYrs.value     = """ & E2_a_asset_master(A512_E2_cas_dur_mnth)			& """" & vbCr '내용월수 >>air           
	Response.Write "	.frm1.txtCasDeprRate.value   = """ & UNINumClientFormat(E7_a_asset_depr_rate(A512_E7_depr_rate), ggExchRate.DecPoint, 0)				  & """" & vbCr '상각율            
	Response.Write "	.frm1.txtCasDeprTotAmt.value = """ & UNINumClientFormat(E2_a_asset_master(A512_E2_cas_end_l_term_depr_tot_amt), ggAmtOfMoney.DecPoint, 0) & """" & vbCr '상각누계
	Response.Write "	.frm1.txtCasCptTotAmt.value  = """ & UNINumClientFormat(E2_a_asset_master(A512_E2_cas_end_l_term_cpt_tot_amt), ggAmtOfMoney.DecPoint, 0)  & """" & vbCr '자본적지출누계            
	Response.Write "	.frm1.txtCasBalAmt.value     = """ & UNINumClientFormat(E2_a_asset_master(A512_E2_cas_end_l_term_bal_amt), ggAmtOfMoney.DecPoint, 0)	  & """" & vbCr '미상각잔액            
	Response.Write "	.frm1.cboCasDeprSts.value    = """ & ConvSPChars(E2_a_asset_master(A512_E2_cas_depr_sts))								& """" & vbCr '상각상태            
	Response.Write "	.frm1.txtCasDeprEnd.text    = """ & UNIMonthClientFormat(E2_a_asset_master(A512_E2_cas_depr_end_yyyymm))				& """" & vbCr '상각완료년월            
    
    Response.Write "	.DbQueryOk " & vbCr
    Response.Write " End With   " & vbCr
    Response.Write " </Script>  " & vbCr
    'Response.End

End Sub	
'============================================================================================================
' Name : SubBizSave
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()
'Call ServerMesgBox("a", vbInformation, I_MKSCRIPT)	
	'------------------
	' 변수, 상수 선언 
	'------------------
    Dim iPAAG015
	Dim I1_a_asset_master
	
    'import a_asset_master
    Const A510_I1_asst_no = 0						'자산번호
    Const A510_I1_ref_no = 1						'품목그룹
    Const A510_I1_tax_dur_yrs = 2					'세법기준내용년수
    Const A510_I1_cas_dur_yrs = 3					'기업회계기준내용년수
    Const A510_I1_tax_end_l_term_cpt_tot_amt = 4	'세법기준전기말자본적지출금액
    Const A510_I1_cas_end_l_term_cpt_tot_amt = 5	'기업회계기준전기말자본적지출금액
    Const A510_I1_tax_end_l_term_depr_tot_amt = 6	'세법기준전기말감가상각누계금액
    Const A510_I1_cas_end_l_term_depr_tot_amt = 7	'기업회계기준전기말감가상각누계금액
    Const A510_I1_tax_end_l_term_bal_amt = 8		'세법기준미상각잔액
    Const A510_I1_cas_end_l_term_bal_amt = 9		'기업회계기준미상각잔액
    Const A510_I1_tax_depr_sts = 10					'세법기준상각상태
    Const A510_I1_cas_depr_sts = 11					'기업회계기준상각상태
    Const A510_I1_tax_depr_end_yyyymm = 12			'세법기준상각완료년월
    Const A510_I1_cas_depr_end_yyyymm = 13			'기업회계기준상각완료년월
    Const A510_I1_updt_user_id = 14					'User ID
    Const A510_I1_start_depr_yymm = 15				'감가상각시작년월
    Const A510_I1_vat_rate = 16						'부가세율
    Const A510_I1_net_amt = 17						'공급가액
    Const A510_I1_net_loc_amt = 18					'공급가액(자국)
    Const A510_I1_temp_fg1 = 19						'잔존율표시
    Const A510_I1_asst_nm = 20						'자산명
    Const A510_I1_spec = 21							'용도/규격
    Const A510_I1_asset_desc = 22					'적요
    Const A510_I1_tax_dur_mnth = 23					'세법기준내용월수
    Const A510_I1_cas_dur_mnth = 24					'기업회계기준기준내용월수
    Const A510_I1_COST_CD = 25						'코스트센터 >>AIR

	' -- 권한관리추가
	Const A512_I2_a_data_auth_data_BizAreaCd = 0
	Const A512_I2_a_data_auth_data_internal_cd = 1
	Const A512_I2_a_data_auth_data_sub_internal_cd = 2
	Const A512_I2_a_data_auth_data_auth_usr_id = 3

	Dim I2_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동

  	Redim I2_a_data_auth(3)
	I2_a_data_auth(A512_I2_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I2_a_data_auth(A512_I2_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I2_a_data_auth(A512_I2_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I2_a_data_auth(A512_I2_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))

	'------------------
	' Data Matching
	'------------------
	'*** Import Data ***
	Redim I1_a_asset_master(25)
	
	I1_a_asset_master(A510_I1_asst_no)						= Trim(Request("txtCondAsstNo")) 			   '자산번호                    
    I1_a_asset_master(A510_I1_ref_no)						= Trim(Request("txtRefNo"))					   '품목그룹                    
    I1_a_asset_master(A510_I1_tax_dur_yrs)					= UNIConvNum(Request("txtTaxDurYrs"), 0)       '세법기준내용년수            
    I1_a_asset_master(A510_I1_cas_dur_yrs)					= UNIConvNum(Request("txtCasDurYrs"), 0)       '기업회계기준내용년수        
    I1_a_asset_master(A510_I1_tax_end_l_term_cpt_tot_amt)	= UNIConvNum(Request("txtTaxCptTotAmt"), 0)   '세법기준전기말자본적지출금액
    I1_a_asset_master(A510_I1_cas_end_l_term_cpt_tot_amt)	= UNIConvNum(Request("txtCasCptTotAmt"), 0)   '기업회계기준전기말자본적지출
    I1_a_asset_master(A510_I1_tax_end_l_term_depr_tot_amt)	= UNIConvNum(Request("txtTaxDeprTotAmt"), 0)   '세법기준전기말감가상각누계금
    I1_a_asset_master(A510_I1_cas_end_l_term_depr_tot_amt)	= UNIConvNum(Request("txtCasDeprTotAmt"), 0)   '기업회계기준전기말감가상각누
    I1_a_asset_master(A510_I1_tax_end_l_term_bal_amt)		= UNIConvNum(Request("txtTaxBalAmt"), 0)       '세법기준미상각잔액          
    I1_a_asset_master(A510_I1_cas_end_l_term_bal_amt)		= UNIConvNum(Request("txtCasBalAmt"), 0)       '기업회계기준미상각잔액      
    I1_a_asset_master(A510_I1_tax_depr_sts)					= Request("cboTaxDeprSts")      '세법기준상각상태            
    I1_a_asset_master(A510_I1_cas_depr_sts)					= Request("cboCasDeprSts")      '기업회계기준상각상태        
    I1_a_asset_master(A510_I1_tax_depr_end_yyyymm)			= Request("txtTaxDeprEnd")      '세법기준상각완료년월        
    I1_a_asset_master(A510_I1_cas_depr_end_yyyymm)			= Request("txtCasDeprEnd")      '기업회계기준상각완료년월    
    I1_a_asset_master(A510_I1_updt_user_id)					= gUsrID						'User ID                     
    I1_a_asset_master(A510_I1_start_depr_yymm)				= Request("txtDeprFrdt")        '감가상각시작년월            
    I1_a_asset_master(A510_I1_vat_rate)						= 0								'? 부가세율                    
    I1_a_asset_master(A510_I1_net_amt)						= 0								'? 공급가액                    
    I1_a_asset_master(A510_I1_net_loc_amt)					= 0								'? 공급가액(자국)              
   'I1_a_asset_master(A510_I1_temp_fg1)						= ?						  	    '잔존율표시                  
    I1_a_asset_master(A510_I1_asst_nm)						= Request("txtAsstNm")          '자산명                      
    I1_a_asset_master(A510_I1_spec)							= Request("txtSpec")            '용도/규격                   
    I1_a_asset_master(A510_I1_asset_desc)					= Request("txtDesc")            '적요                        
   'I1_a_asset_master(A510_I1_tax_dur_mnth)					= ?                             '세법기준내용월수            
   'I1_a_asset_master(A510_I1_cas_dur_mnth)					= ?                             '기업회계기준기준내용월수    
	I1_a_asset_master(A510_I1_COST_CD)						= Trim(Request("txtCostCd"))	'코스트센터 >>AIR
	'------------------
'Call ServerMesgBox(Request("txtCostCd"), vbInformation, I_MKSCRIPT)	
	' 요청 변수 처리 
	'------------------
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    Set iPAAG015 = Server.CreateObject("PAAG015_KO441.cAAS0041MngSvr")

    If CheckSYSTEMError(Err, True) = True Then					
       Exit Sub
    End If    
    
    Call iPAAG015.AS0041_MANAGE_SVR(gStrGloBalCollection, I1_a_asset_master, I2_a_data_auth)

    If CheckSYSTEMError(Err, True) = True Then					
       Set iPAAG015 = Nothing
       'Response.End
       Exit Sub
    End If    
    
    Set iPAAG015 = Nothing
	
	'---------------------
	' HTML 및 결과 생성부 
	'---------------------
	Response.Write " <Script Language=vbscript> " & vbCr
	Response.Write " parent.DbSaveOk            " & vbCr
    Response.Write "</Script>                   " & vbCr    

End Sub	

%>
