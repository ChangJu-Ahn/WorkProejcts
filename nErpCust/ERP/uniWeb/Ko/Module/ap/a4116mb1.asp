
<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name          : Account 
'*  2. Function Name        : 
'*  3. Program ID           : a4116mb1.adp
'*  4. Program Name         : (-)채권/출금반제 조회 Logic
'*  5. Program Desc         :
'*  6. Complus List         : 
'*  7. Modified date(First) : 2000/03/30
'*  8. Modified date(Last)  : 2000/03/30
'*  9. Modifier (First)     : YOU SO EUN
'* 10. Modifier (Last)      : YOU SO EUN
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
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
   
On Error Resume Next															'☜: Protect system from crashing
Err.Clear																		'☜: Clear Error status

Call HideStatusWnd																'☜: Hide Processing message
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*","NOCOOKIE","MB")
Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")

Call SubBizQueryMulti()															'☜: Multi  --> Query

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
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

    On Error Resume Next																	'☜: Protect system from crashing
    Err.Clear																				'☜: Clear Error status
	
    Dim iIntPrevKeyIndex
    Dim iStrKeyStream
    Dim iLngMaxRow, iLngMaxRow1
    Dim iLngRow, iStrData
    Dim lgStrPrevKey, lgStrPrevKey1
    Dim lgCurrency
    
    Dim I1_a_open_ar, I2_a_allc_paym
    Dim E1_b_biz_area, E2_b_minor, E3_a_open_ar
    Dim E4_a_allc_paym, E5_b_bank, E6_a_acct
    Dim E7_b_biz_partner, E8_a_gl, E9_b_acct_dept
    Dim E10_b_bank_acct
    Dim EG1_export_group
    
    Dim iPAPG080
    
    Const C_SHEETMAXROWS_D  = 100
    Const ConDate = "1899/12/30"
    
    Const A289_E1_biz_area_cd = 0    
    Const A289_E1_biz_area_nm = 1

    Const A289_E4_paym_no = 0    
    Const A289_E4_paym_dt = 1
    Const A289_E4_allc_type = 2
    Const A289_E4_paym_amt = 3
    Const A289_E4_paym_loc_amt = 4
    Const A289_E4_ref_no = 5
    Const A289_E4_xch_rate = 6
    Const A289_E4_paym_type = 7
    Const A289_E4_note_no = 8
    
    Const A289_E4_diff_kind_cur_amt = 9
    Const A289_E4_diff_kind_cur_loc_amt = 10
    Const A289_E4_dc_amt = 11
    Const A289_E4_dc_loc_amt = 12
    Const A289_E4_doc_cur = 13
    Const A289_E4_diff_kind_cur = 14
    Const A289_E4_paym_desc = 15
    Const A289_E4_temp_gl_no = 16

    Const A289_E5_bank_cd = 0    
    Const A289_E5_bank_nm = 1

    Const A289_E6_acct_cd = 0    
    Const A289_E6_acct_nm = 1

    Const A289_E7_bp_cd = 0    
    Const A289_E7_bp_nm = 1

    Const A289_E9_dept_cd = 0    
    Const A289_E9_dept_nm = 1

    Const A289_EG1_E1_biz_area_cd = 0    
    Const A289_EG1_E1_biz_area_nm = 1
    Const A289_EG1_E2_dept_cd = 2    
    Const A289_EG1_E2_dept_nm = 3
    Const A289_EG1_E3_cls_dt = 4   
    Const A289_EG1_E3_ar_due_dt = 5
    Const A289_EG1_E3_cls_amt = 6
    Const A289_EG1_E3_cls_loc_amt = 7
    Const A289_EG1_E3_dc_amt = 8
    Const A289_EG1_E3_dc_loc_amt = 9
    Const A289_EG1_E3_cls_ar_no = 10
	Const A288_EG1_E3_cls_ar_desc = 11    
    Const A289_EG1_E4_acct_cd = 12   
    Const A289_EG1_E4_acct_nm = 13
    Const A289_EG1_E5_ar_no = 14    
    Const A289_EG1_E5_ar_dt = 15
    Const A289_EG1_E5_ar_amt = 16
    Const A289_EG1_E5_ar_loc_amt = 17
    Const A289_EG1_E5_ar_due_dt = 18
    Const A289_EG1_E5_bal_amt = 19
    Const A289_EG1_E5_bal_loc_amt = 20

	' 권한관리 추가 
	Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
	Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
	Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
	Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

	Const A289_I2_paym_no = 0

	' 권한관리 추가 
	lgAuthBizAreaCd	= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd	= Trim(Request("lgInternalCd"))
	lgSubInternalCd	= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID		= Trim(Request("lgAuthUsrID"))

    Redim I2_a_allc_paym(A289_I2_paym_no+4)
    I2_a_allc_paym(A289_I2_paym_no)   = Trim(Request("txtAllcNo"))
	I2_a_allc_paym(A289_I2_paym_no+1) = lgAuthBizAreaCd
	I2_a_allc_paym(A289_I2_paym_no+2) = lgInternalCd
	I2_a_allc_paym(A289_I2_paym_no+3) = lgSubInternalCd
	I2_a_allc_paym(A289_I2_paym_no+4) = lgAuthUsrID	
    
'	iIntPrevKeyIndex = UNICInt(Trim(Request("txtPrevKeyIndex")),0)							'☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	
	iLngMaxRow  = CLng(Request("txtMaxRows"))												'☜: Fetechd Count      
	
    lgStrPrevKey = Request("lgStrPrevKey")
    lgStrPrevKey1 = Request("lgStrPrevKey1")

'    I2_a_allc_paym = Trim(Request("txtAllcNo"))
	I1_a_open_ar = lgStrPrevKey
	
	Set iPAPG080 = Server.CreateObject("PAPG080.cALkUpAllcPayByArSvr")	
	
	If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
    End If						       
	
	Call iPAPG080.A_LOOKUP_ALLC_PAYM_BY_AR_SVR (gStrGlobalCollection,C_SHEETMAXROWS_D, I1_a_open_ar,I2_a_allc_paym,E1_b_biz_area, E2_b_minor, E3_a_open_ar, _
										E4_a_allc_paym, E5_b_bank, E6_a_acct, E7_b_biz_partner, E8_a_gl, E9_b_acct_dept, E10_b_bank_acct, _
										 EG1_export_group)
	  		
	If CheckSYSTEMError(Err,True) = True Then
		Set iPAPG080 = Nothing																'☜: Err.Raise 일경우 Nothing
		Exit Sub
    End If   

    Set iPAPG080 = Nothing																	'☜: Unload Comproxy DLL
    
    lgCurrency = ConvSPChars(E4_a_allc_paym(A289_E4_doc_cur))
    
    Response.Write "<Script Language=VBScript> "																					  & vbCr
    Response.Write " With parent.frm1 "																								  & vbCr
    Response.Write ".txtAllcDt.text	      = """ & UNIDateClientFormat(E4_a_allc_paym(A289_E4_paym_dt))                         & """" & vbCr 
	Response.Write ".txtDeptCd.Value	  = """ & ConvSPChars(E9_b_acct_dept(A289_E9_dept_cd))                                 & """" & vbCr
	Response.Write ".txtDeptNm.Value	  = """ & ConvSPChars(E9_b_acct_dept(A289_E9_dept_nm))                                 & """" & vbCr
	Response.Write ".txtBankCd.Value	  = """ & ConvSPChars(E5_b_bank(A289_E5_bank_cd))                                      & """" & vbCr
	Response.Write ".txtBankNm.Value	  = """ & ConvSPChars(E5_b_bank(A289_E5_bank_nm))                                      & """" & vbCr
	Response.Write ".txtBpCd.Value		  = """ & ConvSPChars(E7_b_biz_partner(A289_E7_bp_cd))                                 & """" & vbCr
	Response.Write ".txtBpNm.Value		  = """ & ConvSPChars(E7_b_biz_partner(A289_E7_bp_nm))                                 & """" & vbCr
	Response.Write ".txtBankAcct.Value	  = """ & ConvSPChars(E10_b_bank_acct)                                                 & """" & vbCr
	Response.Write ".txtInputType.Value	  = """ & ConvSPChars(E4_a_allc_paym(A289_E4_paym_type))                               & """" & vbCr
	Response.Write ".txtInputTypeNm.Value = """ & ConvSPChars(E2_b_minor)													   & """" & vbCr
	Response.Write ".txtCheckCd.Value	  = """ & ConvSPChars(E4_a_allc_paym(A289_E4_note_no))                                 & """" & vbCr
	Response.Write ".txtDocCur.value	  = """ & ConvSPChars(E4_a_allc_paym(A289_E4_doc_cur))                                 & """" & vbCr
	Response.Write ".txtXchRate.Text	  = """ & UNINumClientFormat(E4_a_allc_paym(A289_E4_xch_rate), ggExchRate.DecPoint, 0) & """" & vbCr
	Response.Write ".txtAcctCd.Value      = """ & ConvSPChars(E6_a_acct(A289_E6_acct_cd))									   & """" & vbCr	
	Response.Write ".txtGlNo.value	      = """ & ConvSPChars(E8_a_gl)														   & """" & vbCr
	Response.Write ".txtTempGlNo.value	  = """ & ConvSPChars(E4_a_allc_paym(A289_E4_temp_gl_no))                              & """" & vbCr
	Response.Write ".txtAllcDesc.value	  = """ & ConvSPChars(E4_a_allc_paym(A289_E4_paym_desc))                               & """" & vbCr
	Response.Write ".txtPaymAmt.Text	  = """ & UNIConvNumDBToCompanyByCurrency(E4_a_allc_paym(A289_E4_paym_amt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")               & """" & vbCr
	Response.Write ".txtPaymLocAmt.Text	  = """ & UNIConvNumDBToCompanyByCurrency(E4_a_allc_paym(A289_E4_paym_loc_amt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X") & """" & vbCr
	Response.Write " End With "																								          & vbCr
    Response.Write "</Script>"	
    
    iStrData = ""
 
	If isArray(EG1_export_group)  And isEmpty(EG1_export_group) = False Then
		For iLngRow = 0 To UBound(EG1_export_group)														
			iStrData = iStrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, A289_EG1_E5_ar_no))
		    iStrData = iStrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, A289_EG1_E4_acct_cd))
		    iStrData = iStrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, A289_EG1_E4_acct_nm))
		    iStrData = iStrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, A289_EG1_E1_biz_area_cd))
		    iStrData = iStrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, A289_EG1_E1_biz_area_nm))
		    iStrData = iStrData & Chr(11) & UNIDateClientFormat(EG1_export_group(iLngRow, A289_EG1_E5_ar_dt))
			iStrData = iStrData & Chr(11) & UNIDateClientFormat(EG1_export_group(iLngRow, A289_EG1_E3_ar_due_dt))	    

		    iStrData = iStrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group(iLngRow, A289_EG1_E5_ar_amt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")
		    iStrData = iStrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group(iLngRow, A289_EG1_E5_bal_amt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")
		    iStrData = iStrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group(iLngRow, A289_EG1_E3_cls_amt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")
		    iStrData = iStrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_export_group(iLngRow, A289_EG1_E3_cls_loc_amt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
		    iStrData = iStrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, A288_EG1_E3_cls_ar_desc))	    

		    iStrData = iStrData & Chr(11) & iLngMaxRow + iLngRow + 1
		    iStrData = iStrData & Chr(11) & Chr(12)
		Next  
	End if	
		
    Response.Write "<Script Language=VBScript>						   " & vbCr  
    Response.Write " With parent								  	   " & vbCr 
    Response.Write " .ggoSpread.Source    = .frm1.vspdData      " & vbCr
    Response.Write " .ggoSpread.SSShowData  """ & iStrData & """" & vbCr
    Response.Write " .DbQueryOk										   " & vbCr   
    Response.Write " End With										   " & vbCr
    Response.Write "</Script>                                          " & vbCr          
 
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
