<%@ LANGUAGE=VBSCript %>
<% Option Explicit %>
<%
'**********************************************************************************************
'*  1. Module Name          : Account
'*  2. Function Name        : 
'*  3. Program ID           : a4106mb1.asp
'*  4. Program Name         : 채무반제(선금급) 조회 logic
'*  5. Program Desc         :
'*  6. Complus List         : 
'*  7. Modified date(First) : 2000/03/30
'*  8. Modified date(Last)  : 2000/03/30
'*  9. Modifier (First)     : You So Eun
'* 10. Modifier (Last)      : You So Eun
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
															'☜ : ASP가 캐쉬되지 않도록 한다.
															'☜ : ASP가 버퍼에 저장되어 마지막에 바로 Client에 내려간다.

'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%
    
Dim lgOpModeCRUD
    
On Error Resume Next															'☜: Protect system from crashing
Err.Clear																		'☜: Clear Error status

Call HideStatusWnd																'☜: Hide Processing message
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*", "NOCOOKIE", "MB")
Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")
    
Call SubBizQueryMulti()															'☜: Multi  --> Query

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                        '☜: Protect system from crashing
    Err.Clear                                                                   '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub    
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	
    Dim iIntPrevKeyIndex
    Dim iStrKeyStream
    Dim iLngMaxRow, iLngMaxRow1
    Dim iLngRow, iStrData, iStrData1
    Dim iStrPrevKey
    Dim lgCurrency
    
    Dim I1_a_paym_dc, I2_a_open_ap, I3_a_allc_paym
    Dim E1_b_biz_area, E2_a_paym_dc, E3_a_open_ap
    Dim E4_b_acct_dept, E5_a_acct, E6_b_biz_partner
    Dim E7_b_bank, E8_b_bank_acct, E9_a_allc_paym
    Dim E10_f_prpaym, E11_a_gl
    Dim EG1_export_group, EG2_export_group_dc
    
    Dim iPAPG030
    
    Const A292_E1_biz_area_cd = 0
    Const A292_E1_biz_area_nm = 1

    Const A292_E4_dept_cd = 0
    Const A292_E4_dept_nm = 1

    Const A292_E5_acct_cd = 0
    Const A292_E5_acct_nm = 1

    Const A292_E6_bp_cd = 0
    Const A292_E6_bp_nm = 1

    Const A292_E7_bank_cd = 0
    Const A292_E7_bank_nm = 1

    Const A292_E9_paym_no = 0
    Const A292_E9_paym_dt = 1
    Const A292_E9_allc_type = 2
    Const A292_E9_paym_amt = 3
    Const A292_E9_paym_loc_amt = 4
    Const A292_E9_ref_no = 5
    Const A292_E9_xch_rate = 6
    Const A292_E9_paym_type = 7
    Const A292_E9_note_no = 8
    Const A292_E9_dc_amt = 9
    Const A292_E9_dc_loc_amt = 10
    Const A292_E9_doc_cur = 11
    Const A292_E5_diff_kind_cur_amt = 12
    Const A292_E5_diff_kind_cur_loc_amt = 13
    
    Const A292_E5_paym_desc = 14
    Const A292_E9_temp_gl_no = 15

    Const A292_E10_prpaym_no = 0
    Const A292_E10_prpaym_dt = 1
    Const A292_E10_prpaym_amt = 2
    Const A292_E10_prpaym_loc_amt = 3
    Const A292_E10_bal_amt = 4
    Const A292_E10_bal_loc_amt = 5

    Const A292_EG1_E1_b_biz_area_biz_area_cd = 0
    Const A292_EG1_E1_b_biz_area_biz_area_nm = 1
    Const A292_EG1_E2_b_acct_dept_dept_cd = 2
    Const A292_EG1_E2_b_acct_dept_dept_nm = 3
    Const A292_EG1_E3_b_biz_partner_bp_cd = 4
    Const A292_EG1_E3_b_biz_partner_bp_nm = 5
    Const A292_EG1_E4_a_acct_acct_cd = 6
    Const A292_EG1_E4_a_acct_acct_nm = 7
    Const A292_EG1_E5_a_cls_ap_cls_dt = 8
    Const A292_EG1_E5_a_cls_ap_doc_cur = 9
    Const A292_EG1_E5_a_cls_ap_xch_rate = 10
    Const A292_EG1_E5_a_cls_ap_cls_amt = 11
    Const A292_EG1_E5_a_cls_ap_cls_loc_amt = 12
    Const A292_EG1_E5_a_cls_ap_dc_amt = 13
    Const A292_EG1_E5_a_cls_ap_dc_loc_amt = 14
    Const A292_EG1_E6_a_open_ap_ap_no = 15
    Const A292_EG1_E6_a_open_ap_ap_dt = 16
    Const A292_EG1_E6_a_open_ap_doc_cur = 17
    Const A292_EG1_E6_a_open_ap_xch_rate = 18
    Const A292_EG1_E6_a_open_ap_ap_due_dt = 19
    Const A292_EG1_E6_a_open_ap_ap_amt = 20
    Const A292_EG1_E6_a_open_ap_ap_loc_amt = 21
    Const A292_EG1_E6_a_open_ap_bal_amt = 22
    Const A292_EG1_E6_a_open_ap_bal_loc_amt = 23
    Const A292_EG1_E6_a_open_ap_inv_doc_no = 24
    Const A292_EG1_E6_a_open_ap_ref_no = 25
    Const A292_EG1_E6_a_open_ap_cls_amt = 26
    Const A292_EG1_E6_a_open_ap_cls_loc_amt = 27
    Const A292_EG1_E6_a_open_ap_paym_type = 28
    Const A292_EG1_E6_a_open_ap_paym_terms = 29
    Const A292_EG1_E5_a_cls_ap_diff_kind_cur = 30
    Const A292_EG1_E5_a_cls_ap_diff_kind_cur_amt = 31
    Const A292_EG1_E5_a_cls_ap_diff_kind_cur_loc_amt = 32
    Const A292_EG1_E5_a_cls_ap_cls_ap_desc = 33
    Const A292_EG1_E6_a_open_ap_ap_desc = 34
    Const A292_EG1_E6_a_open_ap_adjust_amt = 35
    Const A292_EG1_E6_a_open_ap_adjust_loc_amt = 36

    Const A292_EG2_E1_a_acct_acct_cd = 0
    Const A292_EG2_E1_a_acct_acct_nm = 1
    Const A292_EG2_E2_a_paym_dc_seq = 2
    Const A292_EG2_E2_a_paym_dc_dc_amt = 3
    Const A292_EG2_E2_a_paym_dc_dc_loc_amt = 4

	' 권한관리 추가 
	Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
	Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
	Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
	Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

	Const A292_I3_paym_no = 0
        
    On Error Resume Next																	'☜: Protect system from crashing
    Err.Clear																				'☜: Clear Error status
	
'	iIntPrevKeyIndex = UNICInt(Trim(Request("txtPrevKeyIndex")),0)							'☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	
	iLngMaxRow  = CLng(Request("txtMaxRows"))												'☜: Fetechd Count      
	iLngMaxRow1  = CLng(Request("txtMaxRows1"))
	
'	iStrPrevKey = Request("lgStrPrevKey")

	lgAuthBizAreaCd	= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd	= Trim(Request("lgInternalCd"))
	lgSubInternalCd	= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID		= Trim(Request("lgAuthUsrID"))

    Redim I3_a_allc_paym(A292_I3_paym_no+4)
    I3_a_allc_paym(A292_I3_paym_no)   = Trim(Request("txtAllcNo"))
	I3_a_allc_paym(A292_I3_paym_no+1) = lgAuthBizAreaCd
	I3_a_allc_paym(A292_I3_paym_no+2) = lgInternalCd
	I3_a_allc_paym(A292_I3_paym_no+3) = lgSubInternalCd
	I3_a_allc_paym(A292_I3_paym_no+4) = lgAuthUsrID	
	
'    I3_a_allc_paym = Trim(Request("txtAllcNo"))
	
	Set iPAPG030 = Server.CreateObject("PAPG030.cALkUpAllcPpSvr")	
	
	If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If						       
	
	Call iPAPG030.A_LOOKUP_ALLC_PREPAYM_SVR (gStrGlobalCollection,I1_a_paym_dc,I2_a_open_ap,I3_a_allc_paym, E1_b_biz_area, E2_a_paym_dc, _
										E3_a_open_ap, E4_b_acct_dept, E5_a_acct, E6_b_biz_partner, E7_b_bank, E8_b_bank_acct, E9_a_allc_paym, _
										 E10_f_prpaym, E11_a_gl, EG1_export_group, EG2_export_group_dc)
	  		
	If CheckSYSTEMError(Err,True) = True Then
       Set iPAPG030 = Nothing																'☜: Err.Raise 일경우 Nothing
       Exit Sub
    End If   

    Set iPAPG030 = Nothing																	'☜: Unload Comproxy DLL
    
    lgCurrency = ConvSPChars(E9_a_allc_paym(A292_E9_doc_cur))
    
    Response.Write "<Script Language=VBScript> " & vbCr
    Response.Write " With parent.frm1 "          & vbCr
    Response.Write " .hApDocCur.value  = """ & ConvSPChars(EG1_export_group(0, A292_EG1_E6_a_open_ap_doc_cur)) & """" & vbCr

    Response.Write ".txtPPDt.TEXT	       = """ & UNIDateClientFormat(E10_f_prpaym(A292_E10_prpaym_dt)) & """" & vbCr 
	Response.Write ".txtPPNo.Value		   = """ & ConvSPChars(E10_f_prpaym(A292_E10_prpaym_no))         & """" & vbCr
	Response.Write ".txtAllcDt.TEXT	       = """ & UNIDateClientFormat(E9_a_allc_paym(A292_E9_paym_dt))  & """" & vbCr
	Response.Write ".txtDeptCd.Value	   = """ & ConvSPChars(E4_b_acct_dept(A292_E4_dept_cd))          & """" & vbCr
	Response.Write ".txtDeptNm.Value	   = """ & ConvSPChars(E4_b_acct_dept(A292_E4_dept_nm))          & """" & vbCr
	Response.Write ".txtBizCd.Value	       = """ & ConvSPChars(E1_b_biz_area(A292_E1_biz_area_cd))       & """" & vbCr
	Response.Write ".txtBizNm.Value		   = """ & ConvSPChars(E1_b_biz_area(A292_E1_biz_area_nm))       & """" & vbCr
	Response.Write ".txtBpCd.value	       = """ & ConvSPChars(E6_b_biz_partner(A292_E6_bp_cd))          & """" & vbCr
	Response.Write ".txtBpNm.Value	       = """ & ConvSPChars(E6_b_biz_partner(A292_E6_bp_nm))          & """" & vbCr
	Response.Write ".txtBankAcct.Value	   = """ & ConvSPChars(E8_b_bank_acct)                           & """" & vbCr
	Response.Write ".txtCheckNo.Value	   = """ & ConvSPChars(E9_a_allc_paym(A292_E9_note_no))          & """" & vbCr
	
	Response.Write ".txtDocCur.Value	   = """ & ConvSPChars(E9_a_allc_paym(A292_E9_doc_cur))          & """" & vbCr
	Response.Write ".txtGlNo.value		   = """ & ConvSPChars(E11_a_gl)								 & """" & vbCr
	Response.Write ".txtTempGlNo.value	   = """ & ConvSPChars(E9_a_allc_paym(A292_E9_temp_gl_no))       & """" & vbCr
	Response.Write ".txtAllcDesc.value	   = """ & ConvSPChars(E9_a_allc_paym(A292_E5_paym_desc))        & """" & vbCr
	
	Response.Write ".txtBalAmt.Text	       = """ & UNIConvNumDBToCompanyByCurrency(E10_f_prpaym(A292_E10_bal_amt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")                 & """" & vbCr
	Response.Write ".txtBalLocAmt.Text	   = """ & UNIConvNumDBToCompanyByCurrency(E10_f_prpaym(A292_E10_bal_loc_amt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")   & """" & vbCr
	Response.Write ".txtClsAmt.Text	       = """ & UNIConvNumDBToCompanyByCurrency(E9_a_allc_paym(A292_E9_paym_amt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")               & """" & vbCr
	Response.Write ".txtClsLocAmt.Text	   = """ & UNIConvNumDBToCompanyByCurrency(E9_a_allc_paym(A292_E9_paym_loc_amt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X") & """" & vbCr
	Response.Write ".txtDcAmt.Text	       = """ & UNIConvNumDBToCompanyByCurrency(E9_a_allc_paym(A292_E9_dc_amt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")                 & """" & vbCr
	Response.Write ".txtDcLocAmt.Text	   = """ & UNIConvNumDBToCompanyByCurrency(E9_a_allc_paym(A292_E9_dc_loc_amt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")   & """" & vbCr
	Response.Write ".txtDcAmt2.Text	       = """ & UNIConvNumDBToCompanyByCurrency(E9_a_allc_paym(A292_E9_dc_amt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")                 & """" & vbCr
	Response.Write ".txtDcLocAmt2.Text	   = """ & UNIConvNumDBToCompanyByCurrency(E9_a_allc_paym(A292_E9_dc_loc_amt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")   & """" & vbCr
	
	Response.Write " End With "                  & vbCr
    Response.Write "</Script>"	
   
    iStrData = ""
    
    lgCurrency = ConvSPChars(EG1_export_group(0, A292_EG1_E6_a_open_ap_doc_cur))
 
	If isArray(EG1_export_group) and isEmpty(EG1_export_group) = False Then
		For iLngRow = 0 To UBound(EG1_export_group)														
			iStrData = iStrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, A292_EG1_E6_a_open_ap_ap_no))
		    iStrData = iStrData & Chr(11) & UNIDateClientFormat(EG1_export_group(iLngRow, A292_EG1_E6_a_open_ap_ap_dt))
		    iStrData = iStrData & Chr(11) & UNIDateClientFormat(EG1_export_group(iLngRow, A292_EG1_E6_a_open_ap_ap_due_dt))
		    iStrData = iStrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, A292_EG1_E6_a_open_ap_doc_cur))		    
		    iStrData = iStrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group(iLngRow, A292_EG1_E6_a_open_ap_ap_amt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")
		    iStrData = iStrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group(iLngRow, A292_EG1_E6_a_open_ap_bal_amt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")
		    iStrData = iStrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group(iLngRow, A292_EG1_E5_a_cls_ap_cls_amt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")
		    iStrData = iStrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group(iLngRow, A292_EG1_E5_a_cls_ap_cls_loc_amt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
		    iStrData = iStrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group(iLngRow, A292_EG1_E5_a_cls_ap_dc_amt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")
		    iStrData = iStrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group(iLngRow, A292_EG1_E5_a_cls_ap_dc_loc_amt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
		    iStrData = iStrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, A292_EG1_E5_a_cls_ap_cls_ap_desc))
		    
		    iStrData = iStrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, A292_EG1_E4_a_acct_acct_cd))
		    iStrData = iStrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, A292_EG1_E4_a_acct_acct_nm))
		    iStrData = iStrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, A292_EG1_E1_b_biz_area_biz_area_cd))
		    iStrData = iStrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, A292_EG1_E1_b_biz_area_biz_area_nm)) 
	      
		    iStrData = iStrData & Chr(11) & iLngMaxRow1 + iLngRow + 1
		    iStrData = iStrData & Chr(11) & Chr(12)
		Next  
    End If
    
    If isArray(EG2_export_group_dc) And  isEmpty(EG2_export_group_dc) = False Then
		iStrData1 = ""
		For iLngRow = 0 To UBound(EG2_export_group_dc)														
			iStrData1 = iStrData1 & Chr(11) & ConvSPChars(EG2_export_group_dc(iLngRow, A292_EG2_E2_a_paym_dc_seq))
		    iStrData1 = iStrData1 & Chr(11) & ConvSPChars(EG2_export_group_dc(iLngRow, A292_EG2_E1_a_acct_acct_cd))
		    iStrData1 = iStrData1 & Chr(11) & ""
		    
		    iStrData1 = iStrData1 & Chr(11) & ConvSPChars(EG2_export_group_dc(iLngRow, A292_EG2_E1_a_acct_acct_nm))
		    
		    iStrData1 = iStrData1 & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG2_export_group_dc(iLngRow, A292_EG2_E2_a_paym_dc_dc_amt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")
		    iStrData1 = iStrData1 & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG2_export_group_dc(iLngRow, A292_EG2_E2_a_paym_dc_dc_loc_amt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
		    
		    iStrData1 = iStrData1 & Chr(11) & iLngMaxRow + iLngRow + 1
		    iStrData1 = iStrData1 & Chr(11) & Chr(12)
		Next  
    End if
   
    Response.Write "<Script Language=VBScript> "                                                          & vbCr  
    Response.Write " With parent "                                                                        & vbCr
    Response.Write " .ggoSpread.Source          = .frm1.vspdData1 "								  & vbCr
    Response.Write " .ggoSpread.SSShowData        """ & iStrData							   & """" & vbCr 
    Response.Write " .ggoSpread.Source          = .frm1.vspdData "						          & vbCr
    Response.Write " .ggoSpread.SSShowData        """ & iStrData1						   & """" & vbCr
    Response.Write " .DbQueryOk "																		  & vbCr   
    Response.Write " End With "                                                                           & vbCr
    Response.Write "</Script>"  																		  & vbCr          
End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()   		                                                                    
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub    

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    Call SetErrorStatus()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

%>
