<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

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
'					1. Include
'===================================================================
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%
'===================================================================
'					2. 조건부 
'===================================================================
	On Error Resume Next
	Err.Clear  

	Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide


    '---------------------------------------Common-----------------------------------------------------------
    Call LoadBasisGlobalInf()
	Dim lgCurrency, lgStrPrevKey_i, lgBlnFlgChgValue, plgStrPrevKey_i
	Call LoadInfTB19029B("I","*", "NOCOOKIE", "MB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")
'    lgErrorStatus     = "NO"
'    lgErrorPos        = ""                                                           '☜: Set to space
'    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)

	Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
	strMode = Request("txtMode")												'☜ : 현재 상태를 받음 

'===================================================================
'					2.1 조건 체크 
'===================================================================

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
'					2. 업무 처리 수행부 
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
    Dim E8_b_tax_biz_area	'세금신고사업장
	'20030301	미지급금계정추가
	Dim E9_a_acct			'미지급금계정
	'20050512	신용카드번호 추가
	DIM E10_CREDIT_CARD_NO	'신용카드 번호
  
    Dim iLngRow,iLngCol
    Dim idate
    Dim iStrData
    Dim iStrData2
    Dim intMaxRows_i
    Dim intMaxRows_m
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
    Const A311_E5_ref_no = 13			'품목그룹
    '20030301 미지급금계정 추가 -> A311_E9_ap_acct_cd 로 대체.
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
    
'추가내역
    Const A311_E5_issued_dt = 25
    Const A311_E5_tax_biz_area_cd = 26

	'[CONVERSION INFORMATION]  EXPORTS View 상수
	'[CONVERSION INFORMATION]  View Name : exp_fr b_tax_biz_area
	Const A311_E8_tax_biz_area_cd = 0
	Const A311_E8_tax_biz_area_nm = 1
'추가내역
 
    '20030301   EXPORTS View 상수
    Const A311_E9_ap_acct_cd = 0
    Const A311_E9_ap_acct_nm = 1
    
    '20050512	EXPORTS View 상수
    Const A311_E10_CREDIT_CARD_NO = 0
    Const A311_E10_CREDIT_CARD_NM = 1
    

    '[[EXPORTS Group 상수]]
    'Group Name : export_group(b_acct_dept, a_acct,a_asset_master) >>> 자산Master
    Const A311_EG2_E1_dept_cd = 0		'부서코드
    Const A311_EG2_E1_dept_nm = 1		'부서약명
    Const A311_EG2_E3_cost_cd = 3		'Cost Center : air   
    Const A311_EG2_E2_acct_cd = 4		'계정코드
    Const A311_EG2_E2_acct_nm = 5		'계정단명
    Const A311_EG2_E3_asst_no = 6		'자산번호
    Const A311_EG2_E3_asst_nm = 7		'자산명
    Const A311_EG2_E3_acq_amt = 8		'취득금액
    Const A311_EG2_E3_acq_loc_amt = 9	'취득금액(자국)
    Const A311_EG2_E3_acq_qty = 10		'취득수량
    Const A311_EG2_E3_dur_yrs = 11		'내용연수    : air 
    Const A311_EG2_E3_res_amt = 12		'잔존가액(자국)
    Const A311_EG2_E3_ref_no = 12		'품목그룹	 : air
    Const A311_EG2_E3_asset_desc = 13	'적요

    Const A311_EG2_E3_reg_dt = 14		'자산등록일자
    Const A311_EG2_E3_spec = 15			'용도/규격
    Const A311_EG2_E3_doc_cur = 16		'거래통화
    Const A311_EG2_E3_xch_rate = 17		'환율
    Const A311_EG2_E3_inv_qty = 18		'재고수량
    Const A311_EG2_E3_tax_dur_yrs = 19	'세법기준내용년수
    Const A311_EG2_E3_cas_dur_yrs = 20	'기업회계기준내용년수
    Const A311_EG2_E3_gl_no = 21		'전표번호
    Const A311_EG2_E3_temp_gl_no = 22	'결의전표번호
    
    'Group Name : EG1_exp_group >>> 출금내역
	Const A505_EG2_E1_bank_acct_no = 0    '예적금코드
	Const A505_EG2_E2_acq_seq = 1         '순번
	Const A505_EG2_E2_paym_type = 2		  '출금유형
	Const A505_EG2_E2_paym_amt = 3		  '금액
	Const A505_EG2_E2_paym_loc_amt = 4	  '금액(자국)
	Const A505_EG2_E2_note_no = 5		  '어음번호
	Const A505_EG2_E2_b_minor_nm = 6	  '예적금명
	Const A505_EG2_E2_bulid_asst_no = 7	  '건설중인자산번호 '>>air


	' -- 권한관리추가
	Const I5_a_data_auth_data_BizAreaCd = 0
	Const I5_a_data_auth_data_internal_cd = 1
	Const I5_a_data_auth_data_sub_internal_cd = 2
	Const I5_a_data_auth_data_auth_usr_id = 3

	Dim I5_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동
	
'-------------------------   
' 2.2. 요청 변수 처리 
'-------------------------
	
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	plgStrPrevKey_i = Request("lgStrPrevKey_i")     
	intMaxRows_i = Request("txtMaxRows_i")
	intMaxRows_m = Request("txtMaxRows_m")	

  	Redim I5_a_data_auth(3)
	I5_a_data_auth(I5_a_data_auth_data_BizAreaCd)       = Trim(Request("lgAuthBizAreaCd"))
	I5_a_data_auth(I5_a_data_auth_data_internal_cd)     = Trim(Request("lgInternalCd"))
	I5_a_data_auth(I5_a_data_auth_data_sub_internal_cd) = Trim(Request("lgSubInternalCd"))
	I5_a_data_auth(I5_a_data_auth_data_auth_usr_id)     = Trim(Request("lgAuthUsrID"))
	
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
    
	Set iPAAG010 = Server.CreateObject("PAAG010_KO441.cAAcqLkUpSvr")	'AIR :: PAAG010 -> PAAG010_KO441

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
										EG2_export_itm_grp, _
										E8_b_tax_biz_area, _
										E9_a_acct, _
										E10_CREDIT_CARD_NO, _
										I5_a_data_auth)

    If CheckSYSTEMError(Err, True) = True Then			
		
       Set iPAAG010 = Nothing
       Response.End
    End If    

    Set iPAAG010 = Nothing

'-------------------------
' 2.4. HTML 및 결과 생성부 
'-------------------------

    '자산취득일자, 미지급금 만기일자, 전표일자
    idate = Trim(replace(E7_a_asset_acq(A311_E5_ap_due_dt),"-",""))
    idate = right(idate,4) & "-" & left(idate,2) & "-" & mid(idate,3,2)
    E7_a_asset_acq(A311_E5_ap_due_dt) = idate

    idate = Trim(replace(E7_a_asset_acq(A311_E5_acq_dt),"-",""))
    idate = right(idate,4) & "-" & left(idate,2) & "-" & mid(idate,3,2)
    E7_a_asset_acq(A311_E5_acq_dt) = idate

    idate = Trim(replace(E2_a_batch(A311_E2_gl_dt),"-",""))
    idate = right(idate,4) & "-" & left(idate,2) & "-" & mid(idate,3,2)
    E2_a_batch(A311_E2_gl_dt) = idate

	iStrData = ""
    For iLngRow = 0 To UBound(EG1_exp_group,1) - 1
        For iLngCol = 0 To UBound(EG1_exp_group,2)
        	if iLngCol = A311_EG2_E1_dept_cd then			'부서코드 0		
        		iStrData = iStrData & gColSep & ConvSPChars(EG1_exp_group(iLngRow,iLngCol)) & gColSep
        	elseif iLngCol = A311_EG2_E1_dept_nm then       '부서약명 1		
                	iStrData = iStrData & gColSep & ConvSPChars(EG1_exp_group(iLngRow,iLngCol))

 	
        	elseif iLngCol = A311_EG2_E3_cost_cd then       '코스트센터	2	: air
        		iStrData = iStrData & gColSep & ConvSPChars(EG1_exp_group(iLngRow,iLngCol)) & gColSep
        		
        	elseif iLngCol = A311_EG2_E2_acct_cd then		'계정코드 3		
        		iStrData = iStrData & gColSep & ConvSPChars(EG1_exp_group(iLngRow,iLngCol)) & gColSep


        	elseif iLngCol = A311_EG2_E2_acct_nm then       '계정단명 4		
        		iStrData = iStrData & gColSep & ConvSPChars(EG1_exp_group(iLngRow,iLngCol))
        	elseif iLngCol = A311_EG2_E3_asst_no then       '자산번호 5		
        		iStrData = iStrData & gColSep & ConvSPChars(EG1_exp_group(iLngRow,iLngCol))
        	elseif iLngCol = A311_EG2_E3_asst_nm then       '자산명   6		
        		iStrData = iStrData & gColSep & ConvSPChars(EG1_exp_group(iLngRow,iLngCol))
        	elseif iLngCol = A311_EG2_E3_acq_amt then       '취득금액 7		
        		iStrData = iStrData & gColSep & UNIConvNumDBToCompanyByCurrency(EG1_exp_group(iLngRow,iLngCol),lgCurrency,ggAmtOfMoneyNo, "X" , "X")
		elseif iLngCol = A311_EG2_E3_acq_loc_amt then	'취득금액(자국) 8  
				iStrData = iStrData & gColSep & UNIConvNumDBToCompanyByCurrency(EG1_exp_group(iLngRow,iLngCol),gCurrency,ggAmtOfMoneyNo,gLocRndPolicyNo,"X")
		elseif iLngCol = A311_EG2_E3_acq_qty then       '취득수량 9       
			iStrData = iStrData & gColSep & ConvSPChars(EG1_exp_group(iLngRow,iLngCol))
		
		elseif iLngCol = A311_EG2_E3_dur_yrs then		'내용연수 10	: air 	
			iStrData = iStrData & gColSep & UNIConvNumDBToCompanyByCurrency(EG1_exp_group(iLngRow,iLngCol),gCurrency,ggAmtOfMoneyNo,gLocRndPolicyNo,"X")
		
		elseif iLngCol = A311_EG2_E3_res_amt then       '잔존가액(자국) 11 
			iStrData = iStrData & gColSep & UNIConvNumDBToCompanyByCurrency(EG1_exp_group(iLngRow,iLngCol),lgCurrency,ggAmtOfMoneyNo, "X" , "X")			
		elseif iLngCol = A311_EG2_E3_ref_no then        '품목그룹 12	: air
		       iStrData = iStrData & gColSep & ConvSPChars(EG1_exp_group(iLngRow,iLngCol)) & gColSep
		elseif iLngCol = A311_EG2_E3_asset_desc then    '적요 13		
		     iStrData = iStrData & gColSep & ConvSPChars(EG1_exp_group(iLngRow,iLngCol))



        	end if
        	'iLngCol & "-" & 
		next
		
		iStrData = iStrData & gColSep & intMaxRows_i + iLngRow + 1
		iStrData = iStrData & gColSep & gRowSep
    Next
	'--------------------------------------------------
	'Spread Column
	'Const C_Seq		   = 1			'순번
	'Const C_RcptType	   = 2			'출금유형
	'Const C_RcptTypeNm	   = 3			'출금유형명
	'Const C_Amt		   = 4			'금액
	'Const C_LocAmt		   = 5			'금액(자국)
	'Const C_BankAcct	   = 6			'예적금코드
	'Const C_BankAcctPopup = 7
	'Const C_NoteNo		   = 8			'어음번호
	'Const C_NoteNoPopup   = 9
'    Public Const A505_EG2_E1_bank_acct_no = 0    'View Name : exp_itm_item b_bank_acct
'    Public Const A505_EG2_E2_acq_seq = 1         'View Name : exp_itm_item a_asset_acq_item
'    Public Const A505_EG2_E2_paym_type = 2
'    Public Const A505_EG2_E2_paym_amt = 3
'    Public Const A505_EG2_E2_paym_loc_amt = 4
'    Public Const A505_EG2_E2_note_no = 5
	'--------------------------------------------------

	if isarray(EG2_export_itm_grp) then
		iStrData2 = ""
		For iLngRow = 0 To UBound(EG2_export_itm_grp,1) - 1

			iStrData2 = iStrData2 & gColSep & ConvSPChars(EG2_export_itm_grp(iLngRow,A505_EG2_E2_acq_seq))
		    iStrData2 = iStrData2 & gColSep & ConvSPChars(EG2_export_itm_grp(iLngRow,A505_EG2_E2_paym_type))
		    iStrData2 = iStrData2 & gColSep & ""
		    iStrData2 = iStrData2 & gColSep & ConvSPChars(EG2_export_itm_grp(iLngRow,A505_EG2_E2_b_minor_nm))
		    iStrData2 = iStrData2 & gColSep & UNIConvNumDBToCompanyByCurrency(EG2_export_itm_grp(iLngRow,A505_EG2_E2_paym_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")
		    iStrData2 = iStrData2 & gColSep & UNIConvNumDBToCompanyByCurrency(EG2_export_itm_grp(iLngRow,A505_EG2_E2_paym_loc_amt),gCurrency,ggAmtOfMoneyNo,gLocRndPolicyNo,"X")
		    iStrData2 = iStrData2 & gColSep & ConvSPChars(EG2_export_itm_grp(iLngRow,A505_EG2_E1_bank_acct_no))
		    iStrData2 = iStrData2 & gColSep & ""
		    iStrData2 = iStrData2 & gColSep & ConvSPChars(EG2_export_itm_grp(iLngRow,A505_EG2_E2_note_no))
		    iStrData2 = iStrData2 & gColSep & ""
		    iStrData2 = iStrData2 & gColSep & ConvSPChars(EG2_export_itm_grp(iLngRow,A505_EG2_E2_bulid_asst_no))
		    iStrData2 = iStrData2 & gColSep & ""		    
			iStrData2 = iStrData2 & gColSep & intMaxRows_m + iLngRow + 1
			iStrData2 = iStrData2 & gColSep & gRowSep
		Next
	end if

	plgStrPrevKey_i = ""

	Response.Write " <Script Language=vbscript>				" & vbCr
	Response.Write " With parent						    " & vbCr

	Response.Write "    if """ & E7_a_asset_acq(A311_E5_acq_fg) & """ = ""03"" then " & vbCr          
	Response.Write "       	IntRetCD = .DisplayMsgBox(""117217"",""X"",""X"",""X"") " & vbCr ''취득구분 체크.
	Response.Write "       	.lgBlnFlgChgValue = False " & vbCr          
	Response.Write "       	Call .fncnew()" & vbCr          
	Response.Write "	else	" & vbCr
    
    '*** Master ***
    Response.Write "	.ggoSpread.Source = .frm1.vspdData  " & vbCr 			 
    Response.Write "	.ggoSpread.SSShowData """ & iStrData  & """" & vbCr
    
    '*** 출금내역 ***
	Response.Write "	.ggoSpread.Source = .frm1.vspdData2 " & vbCr 			 
	Response.Write "	.ggoSpread.SSShowData """ & iStrData2 & """" & vbCr

    '*** Control ***
    Response.Write "	.frm1.txtDeptCd.value    = """ & ConvSPChars(E6_b_acct_dept(A311_E6_dept_cd))			 & """" & vbCr '부서코드            
    Response.Write "	.frm1.horgchangeID.value    = """ & ConvSPChars(E6_b_acct_dept(A311_E6_org_change_id))			 & """" & vbCr '부서코드            

    
    Response.Write "	.frm1.txtDeptNm.value    = """ & ConvSPChars(E6_b_acct_dept(A311_E6_dept_nm))			 & """" & vbCr '부서약명        
    Response.Write "	.frm1.txtGLDt.text       = """ & UNIDateClientFormat(E2_a_batch(A311_E2_gl_dt))			 & """" & vbCr '전표일자            
    Response.Write "	.frm1.txtBpCd.value      = """ & ConvSPChars(E5_b_biz_partner(A311_E7_bp_cd))			 & """" & vbCr '거래처코드 
    Response.Write "	.frm1.txtBpNm.value      = """ & ConvSPChars(E5_b_biz_partner(A311_E7_bp_nm))			 & """" & vbCr '거래처명(약명)           
    Response.Write "	.frm1.txtVatTypeNm.value = """ & ConvSPChars(E1_b_minor(A311_E1_minor_nm))		         & """" & vbCr '부가세유형명        
    
    Response.Write "	.frm1.txtAcqNo.value     = """ & ConvSPChars(E7_a_asset_acq(A311_E5_acq_no))			 & """" & vbCr '자산취득번호        
    Response.Write "	.frm1.txtAcqDt.text      = """ & UNIDateClientFormat(E7_a_asset_acq(A311_E5_acq_dt))	 & """" & vbCr '자산취득일자        
    Response.Write "	.frm1.txtDocCur.value    = """ & E7_a_asset_acq(A311_E5_doc_cur)						 & """" & vbCr '거래통화            
    Response.Write "	.frm1.txtXchRate.value   = """ & UNINumClientFormat(E7_a_asset_acq(A311_E5_xch_rate), ggExchRate.DecPoint, 0)											 & """" & vbCr '환율                
    Response.Write "	.frm1.cboAcqFg.value     = """ & E7_a_asset_acq(A311_E5_acq_fg)							 & """" & vbCr '취득구분            
    Response.Write "	.frm1.txtAcqAmt.text	= """ & UNIConvNumDBToCompanyByCurrency(E7_a_asset_acq(A311_E5_tot_acq_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")				 & """" & vbCr '총취득금액          
    Response.Write "	.frm1.txtAcqLocAmt.value = """ & UNIConvNumDBToCompanyByCurrency(E7_a_asset_acq(A311_E5_tot_acq_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X") & """" & vbCr '총취득금액(자국)    
    Response.Write "	.frm1.txtVatAmt.value    = """ & UNIConvNumDBToCompanyByCurrency(E7_a_asset_acq(A311_E5_vat_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")					 & """" & vbCr '부가세금액          
    Response.Write "	.frm1.txtVatLocAmt.value = """ & UNIConvNumDBToCompanyByCurrency(E7_a_asset_acq(A311_E5_vat_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")	 & """" & vbCr '부가세금액(자국)    
    '20030301 미지급금계정 추가
    Response.Write "	.frm1.txtApAcctCd.value  = """ & ConvSPChars(E9_a_acct(A311_E9_ap_acct_cd))				 & """" & vbCr '미지급금계정            
    Response.Write "	.frm1.txtApAcctNm.value  = """ & ConvSPChars(E9_a_acct(A311_E9_ap_acct_nm))				 & """" & vbCr '미지급금계정 
    '20050512 신용카드 번호 추가     
    Response.Write "	.frm1.txtCardNo.value  = """ & ConvSPChars(E10_CREDIT_CARD_NO(A311_E10_CREDIT_CARD_NO))				 & """" & vbCr '신용카드번호            
    Response.Write "	.frm1.txtCardNm.value  = """ & ConvSPChars(E10_CREDIT_CARD_NO(A311_E10_CREDIT_CARD_NM))				 & """" & vbCr '신용카드명 
          
    Response.Write "	.frm1.txtApNo.value      = """ & ConvSPChars(E7_a_asset_acq(A311_E5_ap_no))				 & """" & vbCr '채무번호            
    Response.Write "	.frm1.txtApDueDt.text    = """ & UNIDateClientFormat(E7_a_asset_acq(A311_E5_ap_due_dt))	 & """" & vbCr '채무만료일자        
    Response.Write "	.frm1.txtApAmt.text     = """ & UNIConvNumDBToCompanyByCurrency(E7_a_asset_acq(A311_E5_ap_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")					 & """" & vbCr '채무금액            
    Response.Write "	.frm1.txtApLocAmt.value  = """ & UNIConvNumDBToCompanyByCurrency(E7_a_asset_acq(A311_E5_ap_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")		 & """" & vbCr '채무금액(자국)      
    Response.Write "	.frm1.txtGLNo.value      = """ & ConvSPChars(E7_a_asset_acq(A311_E5_gl_no))				 & """" & vbCr '전표번호            
    Response.Write "	.frm1.txtTempGLNo.value  = """ & ConvSPChars(E7_a_asset_acq(A311_E5_temp_gl_no))		 & """" & vbCr '결의전표번호        
    Response.Write "	.frm1.txtDesc.value      = """ & ConvSPChars(E7_a_asset_acq(A311_E5_acq_desc))			 & """" & vbCr '적요                
    Response.Write "	.frm1.txtVatType.value   = """ & Trim(ConvSPChars(E7_a_asset_acq(A311_E5_vat_type)))			 & """" & vbCr '부가세유형          
    Response.Write "	.frm1.txtVatRate.value   = """ & UNINumClientFormat(E7_a_asset_acq(A311_E5_vat_rate), ggExchRate.DecPoint, 0)											 & """" & vbCr '부가세율            



'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'10 월 정기 패치 추가 사항
	Response.Write "	.frm1.txtReportAreaCd.value        = """ & ConvSPChars(E8_b_tax_biz_area(A311_E8_tax_biz_area_cd)) &				"""" & vbCr
	Response.Write "	.frm1.txtReportAreaNm.value        = """ & ConvSPChars(E8_b_tax_biz_area(A311_E8_tax_biz_area_nm)) &				"""" & vbCr    		 	    
	Response.Write "	.frm1.fpDateTime4.text				= """ & UNIDateClientFormat(E7_a_asset_acq(A311_E5_issued_dt)) &	"""" & vbCr       'AP 만기일자       '변동일자        
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

    Response.Write "	.lgStrPrevKey = """ & plgStrPrevKey_i & """" & vbCr
    Response.Write "	.DbQueryOk " & vbCr
	
	
	Response.Write "    end if	" & vbCr
	
	
    Response.Write " End With   " & vbCr
    Response.Write " </Script>  " & vbCr
    Response.End

%>
