

<%'**********************************************************************************************
'*  1. Module명          : 회계 
'*  2. Function명        : Prepayment
'*  3. Program ID        : f6102mb1
'*  4. Program 이름      : 입금 청산 
'*  5. Program 설명      : 입금 청산 List, Create, Delete, Update
'*  6. Comproxy 리스트   : Ar0071, ar0071
'*  7. 최초 작성년월일   : 2000/10/07
'*  8. 최종 수정년월일   : 2002/06/28
'*  9. 최초 작성자       : 송봉훈 
'* 10. 최종 작성자       : 정승기 
'* 11. 전체 comment      :
'* 12. 공통 Coding Guide : this mark(☜) means that "Do not change"
'*                         this mark(⊙) Means that "may  change"
'*                         this mark(☆) Means that "must change"
'* 13. History           :
'**********************************************************************************************
								'☜ : ASP가 캐쉬되지 않도록 한다.
								'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

Call HideStatusWnd

On Error Resume Next								'☜: 

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*","NOCOOKIE","MB")
Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")

Dim lgOpModeCRUD
Dim lgIntFlgMode

lgIntFlgMode = Request("txtFlgMode")
lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)

Select Case lgOpModeCRUD
    Case CStr(UID_M0001)                                                         '☜: Query
         Call SubBizQueryMulti()
    Case CStr(UID_M0002)                                                         '☜: Save,Update
         Call SubBizSaveMulti()
    Case CStr(UID_M0003)                                                         '☜: Delete
         Call SubBizDelete()
End Select

Sub SubBizQueryMulti() 	'☜: 현재 조회/Prev/Next 요청을 받음 

	On Error Resume Next

	Dim pPARG111 					' 조회용 ComProxy Dll 사용 변수			... 일반 
	Dim lgStrPrevKeyOne_Seq
	Dim iIntQueryCount
	Dim txtAdjust
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp
    Dim strData
	Dim lgCurrency


    Const A340_E2_acct_cd = 1    ' : export a_acct
    Const A340_E2_acct_nm = 2
    Const A340_E3_rcpt_no = 3
    Const A340_E3_adjust_no = 4    ' : export a_rcpt_adjust
    Const A340_E3_adjust_dt = 5
    Const A340_E3_ref_no = 6
    Const A340_E3_doc_dur = 7
    Const A340_E3_xch_rate = 8
    Const A340_E3_adjust_amt = 9
    Const A340_E3_adjust_loc_amt = 10
    Const A340_E3_allc_amt = 11
    Const A340_E3_allc_loc_amt = 12
    Const A340_E3_prrcpt_no = 13
    Const A340_E3_adjust_desc = 14
    Const A340_E3_temp_gl_no = 15
    Const A340_E3_gl_no = 16

    Const A340_E2_dept_cd = 0    ' : export b_acct_dept
    Const A340_E2_dept_nm = 1

    Const A340_E3_bp_cd = 0    ' : export b_biz_partner
    Const A340_E3_bp_nm = 1
 
    Const A340_E5_rcpt_no = 0    ' : export a_rcpt
    Const A340_E5_rcpt_dt = 1
    Const A340_E5_doc_cur = 2
    Const A340_E5_xch_rate = 3
    Const A340_E5_bnk_chg_amt = 4
    Const A340_E5_bnk_chg_loc_amt = 5
    Const A340_E5_rcpt_amt = 6
    Const A340_E5_rcpt_loc_amt = 7
    Const A340_E5_allc_amt = 8
    Const A340_E5_allc_loc_amt = 9
    Const A340_E5_adjust_amt = 10
    Const A340_E5_adjust_loc_amt = 11
    Const A340_E5_bal_amt = 12
    Const A340_E5_bal_loc_amt = 13
    Const A340_E5_rcpt_type = 14
    Const A340_E5_rcpt_desc = 15
    Const A340_E5_bp_cd = 16
    Const A340_E5_gl_no = 17
    
	' -- 조회용 
	' -- 권한관리추가 
	Const A750_I1_a_data_auth_data_BizAreaCd = 0
	Const A750_I1_a_data_auth_data_internal_cd = 1
	Const A750_I1_a_data_auth_data_sub_internal_cd = 2
	Const A750_I1_a_data_auth_data_auth_usr_id = 3

	Dim I1_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 

  	Redim I1_a_data_auth(3)
	I1_a_data_auth(A750_I1_a_data_auth_data_BizAreaCd)       = Trim(Request("lgAuthBizAreaCd"))
	I1_a_data_auth(A750_I1_a_data_auth_data_internal_cd)     = Trim(Request("lgInternalCd"))
	I1_a_data_auth(A750_I1_a_data_auth_data_sub_internal_cd) = Trim(Request("lgSubInternalCd"))
	I1_a_data_auth(A750_I1_a_data_auth_data_auth_usr_id)     = Trim(Request("lgAuthUsrID"))
	    	
    '-----------------------------------------
    'Data manipulate  area(import view match)
    '-----------------------------------------
	txtAdjust = Request("txtAdjustNo")
	
	Set pPARG111  = Server.CreateObject("PARG095.cALkUpRcAdjSvr")

    If CheckSYSTEMError(Err, True) = True Then					
       Exit Sub
    End If    

    Call pPARG111.A_LOOKUP_RCPT_ADJUST_SVR( gStrGlobalCollection, _
											Request("txtAdjustNo"), _
											E1_a_rcpt ,	_
											E2_a_gl , _
											E3_b_biz_partner , _
											E4_b_acct_dept, _
											E5_a_rcpt_adjust , _
											I1_a_data_auth)
											
    If CheckSYSTEMError(Err, True) = True Then					
		Set pPARG111  = Nothing
		Exit Sub
    End If    
		
    Set pPARG111  = Nothing

	lgCurrency = ConvSPChars(E1_a_rcpt(A340_E5_doc_cur))
		
	Response.Write " <Script Language=vbscript> " & vbCr
	Response.Write " With parent.frm1			" & vbCr
	
	If Not IsEmpty(E4_b_acct_dept) Then
		Response.Write " 	.txtDeptCd.value	 	=	""" & ConvSPChars(E4_b_acct_dept(A340_E2_dept_cd))		& """" & vbCr 'BAcctDeptDeptCd
		Response.Write " 	.txtDeptNm.value	 	=	""" & ConvSPChars(E4_b_acct_dept(A340_E2_dept_nm))		& """" & vbCr 'BAcctDeptDeptNm
	End If
	If Not IsEmpty(E3_b_biz_partner) Then
		Response.Write " 	.txtBpCd.value		 	=	""" & ConvSPChars(E3_b_biz_partner(A340_E3_bp_cd))   & """" & vbCr 'BBizPartnerBpCd
		Response.Write " 	.txtBpNm.value		 	=	""" & ConvSPChars(E3_b_biz_partner(A340_E3_bp_nm))   & """" & vbCr 'BBizPartnerBpNm
	End If	
	
	If Not IsEmpty(E1_a_rcpt) Then
		Response.Write " 	.txtRcptNo.value	 		=	""" & ConvSPChars(E1_a_rcpt(A340_E5_rcpt_no))				& """" & vbCr 'ARcptRcptNo=0
		Response.Write " 	.txtRcptDt.text		 	=	""" & UNIDateClientFormat(E1_a_rcpt(A340_E5_rcpt_dt))		& """" & vbCr 'ARcptRcptDt
		Response.Write " 	.txtDocCur.value	 	=	""" & ConvSPChars(E1_a_rcpt(A340_E5_doc_cur))				& """" & vbCr 'ARcptDocCur
		Response.Write " 	.txtXchRate.text		=   """ & UNINumClientFormat(E1_a_rcpt(A340_E5_xch_rate),	ggExchRate.DecPoint,	0) & """" & vbcr
		Response.Write " 	.txtRcptAmt.text	 	=	""" & UNIConvNumDBToCompanyByCurrency(E1_a_rcpt(A340_E5_rcpt_amt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")				& """" & vbCr 'ARcptRcptAmt
		Response.Write " 	.txtRcptLocAmt.text 	=	""" & UNIConvNumDBToCompanyByCurrency(E1_a_rcpt(A340_E5_rcpt_loc_amt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")  & """" & vbCr 'ARcptRcptLocAmt
		Response.Write " 	.txtBalAmt.text	 		=	""" & UNIConvNumDBToCompanyByCurrency(E1_a_rcpt(A340_E5_bal_amt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")				& """" & vbCr 'ARcptBalAmt
		Response.Write " 	.txtBalLocAmt.text	 	=	""" & UNIConvNumDBToCompanyByCurrency(E1_a_rcpt(A340_E5_bal_loc_amt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")   & """" & vbCr 'ARcptBalLocAmt
		Response.Write " 	.txtRcptDesc.value   	=	""" & ConvSPChars(E1_a_rcpt(A340_E5_rcpt_desc))			& """" & vbCr 'ARcptRcptDesc
		Response.Write " 	.txtRefNo.value		 	=	""" & ConvSPChars(E1_a_rcpt(A340_E5_rcpt_no))								& """" & vbCr 'ARcptRcptNo
	End If

    If Not IsEmpty(E5_a_rcpt_adjust) Then
		Response.Write " 	.txtAdjustDt.text	 	=	""" & UNIDateClientFormat(E5_a_rcpt_adjust(A340_E3_adjust_dt))		& """" & vbCr					'2 C_AdjustDt
		Response.Write " 	.txtAcctCd.value	 	=	""" & ConvSPChars(E5_a_rcpt_adjust(A340_E2_acct_cd))					& """" & vbCr					'3  C_AcctCd
		Response.Write " 	.txtAcctNm.value	 	=	""" & ConvSPChars(E5_a_rcpt_adjust(A340_E2_acct_nm))  				& """" & vbCr					'5  C_AcctNm  	
		Response.Write " 	.txtAdjustAmt.value	 	=	""" & UNINumClientFormat(E5_a_rcpt_adjust(A340_E3_allc_amt),	ggAmtOfMoney.DecPoint	,0) & """" & vbCr		' AArAdjustAdjustAmt 
		Response.Write " 	.txtAdjustLocAmt.value	 	=	""" & UNINumClientFormat(E5_a_rcpt_adjust(A340_E3_allc_loc_amt),	ggAmtOfMoney.DecPoint	,0) & """" & vbCr ' AArAdjustAdjustLocAmt 
'		Response.Write " 	.txtAdDocCur.value	 	=	""" &  ConvSPChars(E5_a_rcpt_adjust(A340_E3_doc_dur))			& """" & vbCr					'8  C_DocCur     AArAdjustDocDur 	
'		Response.Write " 	.txtAdXchRate.text		=   """ &  UNINumClientFormat(E5_a_rcpt_adjust(A340_E3_xch_rate),	ggExchRate.DecPoint,	0) & """" &  vbcr
		Response.Write " 	.txtAdDesc.value	 	=	""" & ConvSPChars(E5_a_rcpt_adjust(A340_E3_adjust_desc))		& """" & vbCr					'10  AdjustDesc  AArAdjustAdjustDesc 	
		Response.Write " 	.txtTEMPGlNo.value	 	=	""" & ConvSPChars(E5_a_rcpt_adjust(A340_E3_temp_gl_no))		& """" & vbCr					'TempGlNo        AArAdjustTempGlNo 
		Response.Write " 	.txtGlNo.value	 	=	""" & ConvSPChars(E5_a_rcpt_adjust(A340_E3_gl_no))			& """" & vbCr					'10 GlNo         AdjustAGlGlNo 	
		Response.Write " 	.txtAdJustNo.value	 	=	""" & ConvSPChars(E5_a_rcpt_adjust(A340_E3_adjust_no))		& """" & vbCr					'11 AdjustNo     AArAdjustAdjustNo 	
	End IF
                 
                 
	Response.Write " End With				" & vbCr
	Response.Write "	parent.DbQueryOk	" & vbCr		
	Response.Write " </Script>  " & vbCr       

		
 End Sub
'--------------------------------------------------------------------------------------------------------
'                                   SAVE
'--------------------------------------------------------------------------------------------------------
Sub SubBizSaveMulti() 	
									'☜: 저장 요청을 받음 
	On Error Resume Next
	Err.Clear																		'☜: Protect system from crashing
	
	Dim pPARG111 					' 조회용 ComProxy Dll 사용 변수			... 일반 
	Dim AAcctTransTypeTransType		
	Dim iCommandSent
	Dim ARcptRcptNo
	Dim ARcptAdjustAdjustNo
	Dim ARcptAdjustAdjustDt
	Dim ARcptAdjustAcctCd
	Dim ARcptAdjustDocCur
	Dim ARcptAdjustAdjustAmt
	Dim ARcptAdjustAdjustLocAmt
	Dim ARcptAdjustAdjustDesc
    Dim E1_a_rcpt_adjust_adjust_no

	' -- 저장용 
	' -- 권한관리추가 
	Const A750_I1_a_data_auth_data_BizAreaCd = 0
	Const A750_I1_a_data_auth_data_internal_cd = 1
	Const A750_I1_a_data_auth_data_sub_internal_cd = 2
	Const A750_I1_a_data_auth_data_auth_usr_id = 3

	Dim I1_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 

  	Redim I1_a_data_auth(3)
	I1_a_data_auth(A750_I1_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I1_a_data_auth(A750_I1_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I1_a_data_auth(A750_I1_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I1_a_data_auth(A750_I1_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))
	
	If Trim(lgIntFlgMode) = Trim(OPMD_CMODE) Then
		iCommandSent = "CREATE"
	ElseIf Trim(lgIntFlgMode) = Trim(OPMD_UMODE) Then
		iCommandSent = "UPDATE"
	End If
   										'☜: 최대 업데이트된 갯수    
    LngMaxRow3 = CInt(Request("txtMaxRows3"))

    ' GL HEADER 저장 
	ARcptRcptNo				= Trim(Request("txtRcptNo"))	
	AAcctTransTypeTransType = "AR009"
	ARcptAdjustAdjustNo		=	Trim(Request("txtAdjustNo"))
	ARcptAdjustAdjustDt		=	uniConvDate(Request("txtAdjustDt"))
	ARcptAdjustAcctCd		=	Trim(Request("txtAcctCd"))
	ARcptAdjustDocCur		=	Trim(Request("txtDocCur"))
	ARcptAdjustAdjustAmt	=	uniConvNum(Request("txtAdjustAmt"),0)
	ARcptAdjustAdjustLocAmt	=	uniConvNum(Request("txtAdjustLocAmt"),0)
	ARcptAdjustAdjustDesc	=	Trim(Request("txtAdDesc"))
	
	
    
    Set pPARG111 = Server.CreateObject("PARG095.cAMngRcAdjSvr") 
       
    If CheckSYSTEMError(Err, True) = True Then					
       Exit Sub
    End If    

    Call pPARG111.A_MANAGE_RCPT_ADJUST_SVR(gStrGlobalCollection, _
											iCommandSent, _
											AAcctTransTypeTransType, _
											ARcptRcptNo, _
											ARcptAdjustAdjustNo, _
											ARcptAdjustAdjustDt, _
											ARcptAdjustDocCur, _
											ARcptAdjustAdjustAmt, _
											ARcptAdjustAdjustLocAmt, _
											ARcptAdjustAdjustDesc, _
											ARcptAdjustAcctCd, _
											Request("txtSpread2"), _
											E1_a_rcpt_adjust_adjust_no, _
											I1_a_data_auth)

    If CheckSYSTEMError(Err, True) = True Then					
		Set pPARG111 = Nothing
		Exit Sub
    End If    
    
    Set pPARG111 = Nothing
	'Call ServerMesgBox(Trim(E1_a_rcpt_adjust_adjust_no) , vbInformation, I_MKSCRIPT)
	Response.Write " <Script Language=vbscript> " & vbCr
	Response.Write "	parent.frm1.txtAdjustNo.value = """ & Trim(E1_a_rcpt_adjust_adjust_no)  & """" & vbcr
	Response.Write " parent.DbSaveOk            " & vbCr
    Response.Write "</Script>                   " & vbCr

End Sub


'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
	On Error Resume Next
	Err.Clear																		'☜: Protect system from crashing
	
	Dim pPARG111 					' 조회용 ComProxy Dll 사용 변수			... 일반 
	Dim AAcctTransTypeTransType		
	Dim iCommandSent
	Dim ARcptRcptNo
	Dim ARcptAdjustAdjustNo

	' -- 권한관리추가 
	Const A750_I1_a_data_auth_data_BizAreaCd = 0
	Const A750_I1_a_data_auth_data_internal_cd = 1
	Const A750_I1_a_data_auth_data_sub_internal_cd = 2
	Const A750_I1_a_data_auth_data_auth_usr_id = 3

	Dim I1_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 

  	Redim I1_a_data_auth(3)
	I1_a_data_auth(A750_I1_a_data_auth_data_BizAreaCd)       = Trim(Request("lgAuthBizAreaCd"))
	I1_a_data_auth(A750_I1_a_data_auth_data_internal_cd)     = Trim(Request("lgInternalCd"))
	I1_a_data_auth(A750_I1_a_data_auth_data_sub_internal_cd) = Trim(Request("lgSubInternalCd"))
	I1_a_data_auth(A750_I1_a_data_auth_data_auth_usr_id)     = Trim(Request("lgAuthUsrID"))
	
    'Call ServerMesgBox("iCommandSent" , vbInformation, I_MKSCRIPT)
	
	iCommandSent = "DELETE"
	
	ARcptRcptNo   = Trim(Request("txtRcptNo"))	
	AAcctTransTypeTransType = "AR009"
    ARcptAdjustAdjustNo		=	Trim(Request("txtAdjustNo"))
    
    Set pPARG111 = Server.CreateObject("PARG095.cAMngRcAdjSvr") 
       
    If CheckSYSTEMError(Err, True) = True Then					
       Exit Sub
    End If    

    Call pPARG111.A_MANAGE_RCPT_ADJUST_SVR(gStrGlobalCollection, _
											iCommandSent, _
											AAcctTransTypeTransType, _
											ARcptRcptNo, _
											ARcptAdjustAdjustNo, _
											, _
											, _
											, _
											, _
											, _
											, _
											, _
											, _
											I1_a_data_auth)

    If CheckSYSTEMError(Err, True) = True Then					
		Set pPARG111 = Nothing
		Exit Sub
    End If    
    
    Set pPARG111 = Nothing
	
	Response.Write " <Script Language=vbscript> " & vbCr
	Response.Write " parent.DbDeleteOk            " & vbCr
    Response.Write "</Script>                   " & vbCr
    
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub
  
%>
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                             
