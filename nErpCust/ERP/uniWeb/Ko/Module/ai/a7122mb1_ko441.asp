<%@ LANGUAGE=VBSCript%>
<%Option Explicit%>
<% 
'**********************************************************************************************
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : 
'*  3. Program ID           : a7103mb1
'*  4. Program Name         : 고정자산취득내역등록 
'*  5. Program Desc         : 고정자산취득내역을 조회 
'*  6. Comproxy List        : +As0029LookupSvr
'                             +B1a028ListMinorCode
'*  7. Modified date(First) : 2000/03/30
'*  8. Modified date(Last)  : 2001/05/25
'*  9. Modifier (First)     : 김희정 
'* 10. Modifier (Last)      : 김희정 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************

Response.Expires = -1								'☜ : ASP가 캐쉬되지 않도록 한다.
Response.Buffer = True								'☜ : ASP가 버퍼에 저장되어 마지막에 바로 Client에 내려간다.

'===================================================================
'						1. Include
'===================================================================
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%
'===================================================================
'						2. 조건부 
'===================================================================
Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
On Error Resume Next														'☜: 

    Call LoadBasisGlobalInf()
	Dim lgCurrency, lgStrPrevKey_i, lgBlnFlgChgValue, plgStrPrevKey_i
	Call LoadInfTB19029B("I","*","NOCOOKIE","MB")
	Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")

Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
strMode = Request("txtMode")												'☜ : 현재 상태를 받음 

'-------------------------  
' 2.1 조건 체크 
'-------------------------  
If strMode = "" Then
	Response.End 
ElseIf strMode <> CStr(UID_M0001) Then											'☜: 조회 전용 Biz 이므로 다른값은 그냥 종료함 
	Call ServerMesgBox("700118", vbInformation, I_MKSCRIPT)		'⊙: 조회 전용인데 다른 상태로 요청이 왔을 경우, 필요없으면 빼도 됨, 메세지는 ID값으로 사용해야 함 
	Response.End 
ElseIf Trim(Request("txtAcqNo")) = "" Then												'⊙: 조회를 위한 값이 들어왔는지 체크 
	Call ServerMesgBox("700112", vbInformation, I_MKSCRIPT)						'⊙:
	Response.End 
End If

'===================================================================
'						2. 업무 처리 수행부 
'===================================================================
'-------------------------  
' 2.1. 변수, 상수 선언 
'-------------------------  

    Dim iPAAG010
    
    'Import 변수 
    Dim I1_ief_supplied		'구분자 
    Dim I2_a_asset_acq_item	'자산취득순번 
    Dim I3_a_asset_master	'자산번호 
    Dim I4_a_asset_acq		'자산취득번호 

    'Export 변수 
    Dim E1_b_minor			'부가세유형명 
    Dim E2_a_batch			'전표일자 
    Dim E3_a_asset_acq_item	'자산취득순번 
    Dim E4_a_asset_master	'자산번호 
    Dim EG1_exp_group
    Dim E5_b_biz_partner
    Dim E6_b_acct_dept
    Dim E7_a_asset_acq
    Dim EG2_export_itm_grp
     
    Dim iLngRow,iLngCol
    Dim idate, StrDate
    Dim iStrData
    Dim iStrData2
    Dim iStsNm	'상각상태명 
    Dim strYear,strMonth,strDay
    
'    Dim iStrPrevKey

    '[[EXPORTS 상수]]
    Const A505_I1_select_char = 0		'import_fg ief_supplied - 구분자 
    Const A505_I2_acq_seq = 0			'import_next a_asset_acq_item - 자산취득순번 
    Const A505_I3_asst_no = 0			'import_next a_asset_master - 자산번호 
    Const A505_I4_acq_no = 0			'import a_asset_acq - 자산취득번호 
    
    Const A311_E1_minor_nm = 0			'export b_minor - 부가세유형명 
    Const A311_E2_gl_dt = 0				'export a_batch - 전표일자 
    Const A311_E3_asst_no = 0			'export_next a_asset_master - 자산취득순번 
	Const A311_E4_acq_seq = 0			'export_next a_asset_acq_item - 자산번호 

    'export b_acct_dept
    Const A311_E6_org_change_id = 0		'조직변경ID
    Const A311_E6_dept_cd = 1			'부서코드 
    Const A311_E6_dept_nm = 2			'부서약명 

    'export b_biz_partner
    Const A311_E7_bp_cd = 0				'거래처코드 
    Const A311_E7_bp_type = 1			'거래처Type
    Const A311_E7_bp_nm = 2				'거래처명(약명)

    'export a_asset_acq >>> 자산취득내역 
    Const A311_E5_acq_no = 0			'자산취득번호 
    Const A311_E5_acq_dt = 1			'자산취득일자 
    Const A311_E5_doc_cur = 2			'거래통화 
    Const A311_E5_xch_rate = 3			'환율 
    Const A311_E5_acq_fg = 4			'취득구분 
    Const A311_E5_tot_acq_amt = 5		'총취득금액 
    Const A311_E5_tot_acq_loc_amt = 6	'총취득금액(자국)
    Const A311_E5_extra_acq_amt = 7		'부대비용 
    Const A311_E5_extra_acq_loc_amt = 8	'부대비용(자국)
    Const A311_E5_vat_make_fg = 9	
    Const A311_E5_vat_no = 10			'부가세번호 
    Const A311_E5_vat_amt = 11			'부가세금액 
    Const A311_E5_vat_loc_amt = 12		'부가세금액(자국)
    Const A311_E5_ref_no = 13			'참조번호 
    Const A311_E5_ap_acct_cd = 14		'미지급금 계정 
    Const A311_E5_ap_no = 15			'채무번호 
    Const A311_E5_ap_due_dt = 16		'채무만료일자 
    Const A311_E5_ap_amt = 17			'채무금액 
    Const A311_E5_ap_loc_amt = 18		'채무금액(자국)
    Const A311_E5_gl_no = 19			'전표번호 
    Const A311_E5_temp_gl_no = 20		'결의전표번호 
    Const A311_E5_acq_desc = 21			'적요 
    Const A311_E5_internal_cd = 22		'내부부서코드 
    Const A311_E5_vat_type = 23			'부가세유형 
    Const A311_E5_vat_rate = 24			'부가세율 
    
    '[[EXPORTS Group 상수]]
    'Group Name : export_group		(old)
'    Const A505_EG1_E1_dept_cd = 0		'부서코드	'b_acct_dept
'    Const A505_EG1_E1_dept_nm = 1		'부서약명 
'    Const A505_EG1_E2_acct_cd = 2		'계정코드	'a_acct
'    Const A505_EG1_E2_acct_nm = 3		'계정단명 
'    Const A505_EG1_E3_asst_no = 4		'자산번호	'a_asset_master
 '   Const A505_EG1_E3_asst_nm = 5		'자산명 
 '   Const A505_EG1_E3_reg_dt = 7		'자산등록일자 
 '   Const A505_EG1_E3_spec = 8			'용도/규격 
 '   Const A505_EG1_E3_doc_cur = 9		'거래통화 
 '   Const A505_EG1_E3_xch_rate = 10		'환율 
 '   Const A505_EG1_E3_acq_amt = 11		'취득금액 
 '   Const A505_EG1_E3_ref_no = 5		'참조번호 
 '   Const A505_EG1_E3_acq_loc_amt = 12	'취득금액(자국)
 '   Const A505_EG1_E3_acq_qty = 13		'취득수량 
 '   Const A505_EG1_E3_inv_qty = 14		'재고수량 
 '   Const A505_EG1_E3_tax_dur_yrs = 15	'세법기준내용년수 
 '   Const A505_EG1_E3_cas_dur_yrs = 16	'기업회계기준내용년수 
 '   Const A505_EG1_E3_asset_desc = 17	'적요 
 '   Const A505_EG1_E3_gl_no = 18		'전표번호 
 '   Const A505_EG1_E3_temp_gl_no = 19	'결의전표번호 
 '   Const A505_EG1_E3_cas_end_l_term_depr_tot_amt = 20	'기업회계기준전기말감가상각누계금액 
 '   Const A505_EG1_E3_start_depr_yymm = 21				'감가상각시작년월 
 '   Const A505_EG1_E3_cas_depr_sts = 22	'기업회계기준상각상태 
 
'    Const A505_EG1_E1_dept_cd = 0		'부서코드		(new)
'    Const A505_EG1_E1_dept_nm = 1		'부서약명 
'    Const A505_EG1_E2_acct_cd = 2		'계정코드 
'    Const A505_EG1_E2_acct_nm = 3		'계정단명 
'    Const A505_EG1_E3_asst_no = 4		'자산번호 
'    Const A505_EG1_E3_asst_nm = 5		'자산명 
'    Const A505_EG1_E3_acq_amt = 6		'취득금액 
'    Const A505_EG1_E3_acq_loc_amt = 7	'취득금액(자국)
'    Const A505_EG1_E3_acq_qty = 8		'취득수량 
'    Const A505_EG1_E3_res_amt = 9		'잔존가액(자국)
'    Const A505_EG1_E3_ref_no = 10		'참조번호 
'    Const A505_EG1_E3_asset_desc = 11	'적요 
'    Const A505_EG1_E3_reg_dt = 12		'자산등록일자 
'    Const A505_EG1_E3_spec = 13			'용도/규격 
'    Const A505_EG1_E3_doc_cur = 14		'거래통화 
'    Const A505_EG1_E3_xch_rate = 15		'환율 
'    Const A505_EG1_E3_inv_qty = 16		'재고수량 
'    Const A505_EG1_E3_tax_dur_yrs = 17	'세법기준내용년수 
'    Const A505_EG1_E3_cas_dur_yrs = 18	'기업회계기준내용년수 
'    Const A505_EG1_E3_gl_no = 19		'전표번호 
'    Const A505_EG1_E3_temp_gl_no = 20	'결의전표번호 
'    Const A505_EG1_E3_cas_end_l_term_depr_tot_amt = 21		'상각누계 
'    Const A505_EG1_E3_start_depr_yymm = 22					'감가상각시작년월 
'    Const A505_EG1_E3_cas_depr_sts = 23					'상태 
'    Const A505_EG1_E3_cas_dur_mnth = 24	

     
    'Group Name : EG1_exp_group >>> 출금내역 
'	Const A505_EG2_E1_bank_acct_no = 0    '예적금코드 
'	Const A505_EG2_E2_acq_seq = 1         '순번 
'	Const A505_EG2_E2_paym_type = 2		  '출금유형 
'	Const A505_EG2_E2_paym_amt = 3		  '금액 
'	Const A505_EG2_E2_paym_loc_amt = 4	  '금액(자국)
'	Const A505_EG2_E2_note_no = 5		  '어음번호 
	
	
	Const C_AcqDt		= 1
    Const C_Deptcd		= 2
	Const C_DeptPop		= 3
	Const C_DeptNm		= 4
	Const C_AcctCd		= 5
	Const C_AcctPop		= 6
	Const C_AcctNm		= 7
	Const C_AsstNo		= 8
	Const C_AsstNm		= 9
    Const C_AcqAmt		= 10
	Const C_AcqLocAmt	= 11
	Const C_DeprLocAmt	= 12
	Const C_InvQty		= 13
	Const C_ResAmt		= 14
	Const C_DeprFrDt	= 15
	Const C_DurYrs		= 16
	Const C_DeprstsCd	= 17
	Const C_DeprstsPop	= 18
	Const C_Deprsts		= 19
	Const C_RefNo		= 20
	Const C_Desc		= 21



'-------------------------   
' 2.2. 요청 변수 처리 
'-------------------------
	
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	Dim intMaxRows_i
	
	plgStrPrevKey_i = Request("lgStrPrevKey_i")    
	intMaxRows_i = Request("txtMaxRows_i") 
	
	'-----------------------
	'Data manipulate  area(import view match) (자산취득순번)
	'-----------------------
	If plgStrPrevKey_i = "" Then
		I2_a_asset_acq_item = 0
	Else
		I2_a_asset_acq_item = plgStrPrevKey_i
	End If
	
	Redim I4_a_asset_acq(0)
    
    I4_a_asset_acq(A505_I4_acq_no) = Request("txtAcqNo")	'자산취득번호 
    
	Set iPAAG010 = Server.CreateObject("PAAG010_KO441.cAAcqLkUpSvr")

    If CheckSYSTEMError(Err, True) = True Then					
       Response.End
    End If    

	Call iPAAG010.AS0029_ACQ_LOOKUP_SVR(gStrGloBalCollection, _
										I1_ief_supplied, _
										I2_a_asset_acq_item, _
										I3_a_asset_master, _
										I4_a_asset_acq, _
										E1_b_minor, _
										E2_a_batch, _
										E3_a_asset_acq_item, _
										E4_a_asset_master, _
										EG1_exp_group, _
										E5_b_biz_partner, _
										E6_b_acct_dept, _
										E7_a_asset_acq, _
										EG2_export_itm_grp)

    If CheckSYSTEMError(Err, True) = True Then					
       Set iPAAG010 = Nothing
       Response.End
    End If    

    Set iPAAG010 = Nothing

'-------------------------
' 2.4. HTML 및 결과 생성부 
'-------------------------

    '자산취득일자, 전표일자 
    idate = Trim(replace(E7_a_asset_acq(A311_E5_acq_dt),"-",""))
    idate = right(idate,4) & "-" & left(idate,2) & "-" & mid(idate,3,2)
    E7_a_asset_acq(A311_E5_acq_dt) = idate

    idate = Trim(replace(E2_a_batch(A311_E2_gl_dt),"-",""))
    idate = right(idate,4) & "-" & left(idate,2) & "-" & mid(idate,3,2)
    E2_a_batch(A311_E2_gl_dt) = idate

    iStrData = ""
    For iLngRow = 0 To UBound(EG1_exp_group,1) - 1
		Call ExtractDateFrom(EG1_exp_group(iLngRow,24),"YYYYMM","",strYear,strMonth,strDay)
		StrDate = UniConvYYYYMMDDtoDate(gAPDateFormat,strYear,strMonth,"01")

		iStrData = iStrData & Chr(11) & UNIDateClientFormat(EG1_exp_group(iLngRow,14))		'1  취득일자 
		iStrData = iStrData & Chr(11) & Trim(ConvSPChars(EG1_exp_group(iLngRow,0)))			'2  부서코드 
		iStrData = iStrData & Chr(11) & ""													'3  부서pop
		iStrData = iStrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,1))				'4	부서약명 
		'iStrData = iStrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,2))				'	코스트센터
		iStrData = iStrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,3))				'5  계정코드 
		iStrData = iStrData & Chr(11) & ""													'6  계정pop
		iStrData = iStrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,4))				'7	계정단명 
		iStrData = iStrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,5))				'8	자산번호 
		iStrData = iStrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,6))				'9	자산명 
		iStrData = iStrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_exp_group(iLngRow,7),lgCurrency,ggAmtOfMoneyNo, "X" , "X")	'10	취득금액 	
		iStrData = iStrData & Chr(11) &	UNIConvNumDBToCompanyByCurrency(EG1_exp_group(iLngRow,8),gCurrency,ggAmtOfMoneyNo,gLocRndPolicyNo,"X")'11	취득금액(자국)
		iStrData = iStrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,23))				'12	상각누계 
		iStrData = iStrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,18))				'13	재고수량 
		iStrData = iStrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_exp_group(iLngRow,11),lgCurrency,ggAmtOfMoneyNo, "X" , "X")'14	잔존가액 
		iStrData = iStrData & Chr(11) & UNIMonthClientFormat(ConvSPChars(strDate))			'15 감가상각시작년월 
		iStrData = iStrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,26))				'16 내용월수 ??
		iStrData = iStrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,25))				'17 상각상태코드 
		iStrData = iStrData & Chr(11) & ""													'18 상각상태코드팝업 
		iStrData = iStrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,27))				'19	상각상태명 
		iStrData = iStrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,12))				'20 참조번호 
		iStrData = iStrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow,13))				'21 적요 
		iStrData = iStrData & Chr(11) & intMaxRows_i + iLngRow + 1
		iStrData = iStrData & Chr(11) & Chr(12)
    Next

	plgStrPrevKey_i = ""

	Response.Write " <Script Language=vbscript>				" & vbCr
	Response.Write " With parent						    " & vbCr

	Response.Write "    if """ & E7_a_asset_acq(A311_E5_acq_fg) & """ <> ""03"" then " & vbCr          
	Response.Write "       	IntRetCD = .DisplayMsgBox(""117214"",""X"",""X"",""X"") " & vbCr ''취득구분 체크.
	Response.Write "       	.lgBlnFlgChgValue = False " & vbCr          
	Response.Write "       	Call .fncnew()" & vbCr          
	Response.Write "	else	" & vbCr

    '*** Master ***
    Response.Write "	.ggoSpread.Source = .frm1.vspdData  " & vbCr 			 
    Response.Write "	.ggoSpread.SSShowData """ & iStrData  & """" & vbCr
    
    '*** Control ***
    Response.Write "	.frm1.txtDeptCd.value    = """ & Trim(ConvSPChars(E6_b_acct_dept(A311_E6_dept_cd)))  & """" & vbCr '부서코드 
    Response.Write "	.frm1.txtDeptNm.value    = """ & ConvSPChars(E6_b_acct_dept(A311_E6_dept_nm))		 & """" & vbCr '부서약명 
    Response.Write "	.frm1.txtGLDt.text       = """ & UNIDateClientFormat(E2_a_batch(A311_E2_gl_dt))		 & """" & vbCr '전표일자 
    Response.Write "	.frm1.txtBpCd.value      = """ & ConvSPChars(E5_b_biz_partner(A311_E7_bp_cd))		 & """" & vbCr '거래처코드 
    Response.Write "	.frm1.txtBpNm.value      = """ & ConvSPChars(E5_b_biz_partner(A311_E7_bp_nm))		 & """" & vbCr '거래처명(약명)
    
    Response.Write "	.frm1.txtAcqNo.value     = """ & ConvSPChars(E7_a_asset_acq(A311_E5_acq_no))		 & """" & vbCr '자산취득번호 
    Response.Write "	.frm1.txtAcqDt.text      = """ & UNIDateClientFormat(E7_a_asset_acq(A311_E5_acq_dt)) & """" & vbCr '자산취득일자 
    Response.Write "	.frm1.txtDocCur.value    = """ & E7_a_asset_acq(A311_E5_doc_cur)					 & """" & vbCr '거래통화 
    Response.Write "	.frm1.txtXchRate.value   = """ & UNINumClientFormat(E7_a_asset_acq(A311_E5_xch_rate), ggExchRate.DecPoint, 0)											 & """" & vbCr '환율 
    Response.Write "	.frm1.cboAcqFg.value     = """ & E7_a_asset_acq(A311_E5_acq_fg)						 & """" & vbCr '취득구분 
    Response.Write "	.frm1.txtAcqAmt.text    = """ & UNIConvNumDBToCompanyByCurrency(E7_a_asset_acq(A311_E5_tot_acq_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")				 & """" & vbCr '총취득금액 
    Response.Write "	.frm1.txtAcqLocAmt.value = """ & UNIConvNumDBToCompanyByCurrency(E7_a_asset_acq(A311_E5_tot_acq_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X") & """" & vbCr '총취득금액(자국)
    Response.Write "	.frm1.txtGLNo.value      = """ & ConvSPChars(E7_a_asset_acq(A311_E5_gl_no))			 & """" & vbCr '전표번호 
    Response.Write "	.frm1.txtTempGLNo.value  = """ & ConvSPChars(E7_a_asset_acq(A311_E5_temp_gl_no))	 & """" & vbCr '결의전표번호 
    Response.Write "	.frm1.txtDesc.value      = """ & Trim(ConvSPChars(E7_a_asset_acq(A311_E5_acq_desc))) & """" & vbCr '적요 

    Response.Write "	.lgStrPrevKey = """ & plgStrPrevKey_i & """" & vbCr
    Response.Write "	.DbQueryOk " & vbCr

	Response.Write "    end if	" & vbCr

    Response.Write " End With   " & vbCr
    Response.Write " </Script>  " & vbCr
    Response.End

%>
