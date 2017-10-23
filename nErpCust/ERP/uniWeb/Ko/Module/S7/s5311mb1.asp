<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 매출채권관리 
'*  3. Program ID           : S5311MB1
'*  4. Program Name         : 세금계산서등록 
'*  5. Program Desc         : 
'*  6. Comproxy List        : S53119LookupTaxBillHdrSvr, S53111MaintTaxBillHdrSvr, S53115PostTaxBillSvr
'*  7. Modified date(First) : 2000/03/27
'*  8. Modified date(Last)  : 2001/12/19
'*  9. Modifier (First)     : Kim Hyungsuk
'* 10. Modifier (Last)      : Kim Hyungsuk
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            2000/03/27
'*                            2001/12/19	Date표준적용 
'**********************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../ComASP/LoadInfTb19029.asp" -->

<%
    Dim lgOpModeCRUD

	Const lsPOST	= "POST"									

	Call LoadBasisGlobalInf()
	Call loadInfTB19029B("I", "*","NOCOOKIE","MB") 
	Call LoadBNumericFormatB("I","*","NOCOOKIE","MB") 
	
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------

	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 
 
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)
             Call SubBizSave()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
        Case CStr(lsPOST)                                                         '☜: Delete
             Call SubBizPost()
    End Select
'============================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================
Sub SubBizQuery()
	Dim lgCurrency													
    Dim pS7G319
    
    Dim I1_s_tax_bill_hdr_tax_bill_no 
    Dim E1_s_tax_bill_hdr 
    Dim E2_s_tax_doc_no 
    Dim E3_b_sales_grp 
    Dim E4_b_biz_area 
    Dim E5_b_biz_partner 
    Dim E6_b_minor 

    'Const S542_I1_tax_bill_no = 0    'View Name : imp s_tax_bill_hdr

    Const S542_E1_tax_bill_no = 0    'View Name : exp s_tax_bill_hdr
    Const S542_E1_tax_bill_type = 1
    Const S542_E1_issued_dt = 2
    Const S542_E1_vat_calc_type = 3
    Const S542_E1_vat_io_flag = 4
    Const S542_E1_vat_type = 5
    Const S542_E1_vat_rate = 6
    Const S542_E1_cur = 7
    Const S542_E1_xch_rate_op = 8
    Const S542_E1_xch_rate = 9
    Const S542_E1_net_amt = 10
    Const S542_E1_net_loc_amt = 11
    Const S542_E1_vat_amt = 12
    Const S542_E1_vat_loc_amt = 13
    Const S542_E1_cost_cd = 14
    Const S542_E1_biz_area_cd = 15
    Const S542_E1_report_biz_area = 16
    Const S542_E1_bill_no = 17
    Const S542_E1_post_flag = 18
    Const S542_E1_remarks = 19
    Const S542_E1_ext1_qty = 20
    Const S542_E1_ext2_qty = 21
    Const S542_E1_ext3_qty = 22
    Const S542_E1_ext1_amt = 23
    Const S542_E1_ext2_amt = 24
    Const S542_E1_ext3_amt = 25
    Const S542_E1_ext1_cd = 26
    Const S542_E1_ext2_cd = 27
    Const S542_E1_ext3_cd = 28
    Const S542_E1_vat_inc_flag = 29

    Const S542_E2_tax_doc_no = 0    'View Name : exp s_tax_doc_no

    Const S542_E3_sales_grp = 0    'View Name : exp b_sales_grp
    Const S542_E3_sales_grp_nm = 1

    Const S542_E4_biz_area_nm = 0    'View Name : exp_report b_biz_area

    Const S542_E5_bp_cd = 0    'View Name : exp b_biz_partner
    Const S542_E5_bp_nm = 1

    Const S542_E6_minor_nm = 0    'View Name : exp_vat_type_nm b_minor

    On Error Resume Next
    Err.Clear 

    If Request("txtTaxbillNo") = "" Then										'⊙: 조회를 위한 값이 들어왔는지 체크 
		Call ServerMesgBox("조회 조건값이 비어있습니다!", vbInformation, I_MKSCRIPT)              
        Exit Sub
	End If
		
	I1_s_tax_bill_hdr_tax_bill_no  = Trim(Request("txtTaxbillNo"))
	
    Set pS7G319 = Server.CreateObject("PS7G319.cSLkTaxBillHdrSvr")    

	If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If  
    
    Call pS7G319.S_LOOKUP_TAX_BILL_HDR_SVR(gStrGlobalCollection, I1_s_tax_bill_hdr_tax_bill_no, E1_s_tax_bill_hdr, _
							 E2_s_tax_doc_no, E3_b_sales_grp, E4_b_biz_area, E5_b_biz_partner, E6_b_minor )
      
	If CheckSYSTEMError(Err,True) = True Then
       Set pS7G319 = Nothing
		Response.Write "<Script language=vbs>  " & vbCr   
		Response.Write " Parent.frm1.txtTaxbillNo.focus  " & vbCr   		
		Response.Write "</Script>      " & vbCr      
       Exit Sub
    End If  
    
    Set pS7G319 = Nothing


	'-----------------------
	'Display result data
	'----------------------- 
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent.frm1"           & vbCr
	'항상 거래화폐가 우선 
	lgCurrency = ConvSPChars(E1_s_tax_bill_hdr(E1_cur))	
	Response.Write ".txtCurrency.Value			= """ & ConvSPChars(E1_s_tax_bill_hdr(S542_E1_cur))       & """" & vbCr
	Response.Write "parent.CurFormatNumericOCX" & vbCr

	Response.Write ".txtTaxbillNo.Value			= """ & ConvSPChars(E1_s_tax_bill_hdr(S542_E1_tax_bill_no))     & """" & vbCr
	Response.Write ".txtHTaxbillNo.Value		= """ & ConvSPChars(E1_s_tax_bill_hdr(S542_E1_tax_bill_no))     & """" & vbCr
    Response.Write ".txtTaxbillNo1.Value		= """ & ConvSPChars(E1_s_tax_bill_hdr(S542_E1_tax_bill_no))		& """" & vbCr
	Response.Write ".txtBillNo.Value			= """ & ConvSPChars(E1_s_tax_bill_hdr(S542_E1_bill_no))			& """" & vbCr
    Response.Write ".txtTaxbillDocNo.Value		= """ & ConvSPChars(E2_s_tax_doc_no(S542_E2_tax_doc_no))		& """" & vbCr
	Response.Write ".txtBillToParty.Value		= """ & ConvSPChars(E5_b_biz_partner(S542_E5_bp_cd))			& """" & vbCr
	Response.Write ".txtBillToPartyNm.Value		= """ & ConvSPChars(E5_b_biz_partner(S542_E5_bp_nm))			& """" & vbCr

	'If Len(Trim(.txtBillNo.value)) Then
	If Len(Trim(ConvSPChars(E1_s_tax_bill_hdr(S542_E1_bill_no)))) Then
	   Response.Write ".chkBillNoFlg.checked = True " & vbCr
	End If   

' 형변환 R,D		   
	If Trim(Cstr(E1_s_tax_bill_hdr(S542_E1_tax_bill_type)))   = "R" Then 
	   Response.Write ".rdoTaxBillType1.checked = True " & vbCr
	else
	   Response.Write ".rdoTaxBillType2.checked = True " & vbCr
	End If   

	If E1_s_tax_bill_hdr(S542_E1_post_flag)   = "Y" Then 
	   Response.Write ".rdoPostFlg1.checked = True " & vbCr
	else
	   Response.Write ".rdoPostFlg2.checked = True " & vbCr
	End If   

	Response.Write ".txtIssueDt.text		= """ & UNIDateClientFormat(E1_s_tax_bill_hdr(S542_E1_issued_dt))		& """" & vbCr

    Response.Write ".txtBillAmt.Text		= """ & UNIConvNumDBToCompanyByCurrency(E1_s_tax_bill_hdr(S542_E1_net_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")	& """" & vbCr	'##### Rounding Logic #####
	Response.Write ".txtVATAmt.Text			= """ & UNIConvNumDBToCompanyByCurrency(E1_s_tax_bill_hdr(S542_E1_vat_amt),lgCurrency,ggAmtOfMoneyNo, gTaxRndPolicyNo  , "X")	& """" & vbCr	'##### Rounding Logic #####
    
    Response.Write ".txtTaxBizAreaCd.Value	= """ & ConvSPChars(E1_s_tax_bill_hdr(S542_E1_report_biz_area))			& """" & vbCr
	Response.Write ".txtTaxBizAreaNm.Value	= """ & ConvSPChars(E4_b_biz_area(S542_E4_biz_area_nm))					& """" & vbCr

	If E1_s_tax_bill_hdr(S542_E1_vat_calc_type)   = "1" Then 
	   Response.Write ".rdoVATCalcType1.checked = True " & vbCr
	elseif	E1_s_tax_bill_hdr(S542_E1_vat_calc_type)   = "2" Then    
	   Response.Write ".rdoVATCalcType2.checked = True " & vbCr
	End If   

	If E1_s_tax_bill_hdr(S542_E1_vat_inc_flag)   = "1" Then 
	   Response.Write ".rdoVATIncflag1.checked = True " & vbCr
	elseif	E1_s_tax_bill_hdr(S542_E1_vat_inc_flag)   = "2" Then    
	   Response.Write ".rdoVATIncflag2.checked = True " & vbCr
	End If   

	Response.Write ".txtBillLocAmt.text		= """ & UNIConvNumDBToCompanyByCurrency(E1_s_tax_bill_hdr(S542_E1_net_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo  , "X")		& """" & vbCr
	Response.Write ".txtVATLocAmt.text		= """ & UNIConvNumDBToCompanyByCurrency(E1_s_tax_bill_hdr(S542_E1_vat_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo  , "X")		& """" & vbCr
	Response.Write ".txtVATRate.text		= """ & UNIConvNumDBToCompanyByCurrency(E1_s_tax_bill_hdr(S542_E1_vat_rate),gCurrency,ggExchRateNo, "X" , "X")				& """" & vbCr

    Response.Write ".txtVATType.Value		= """ & ConvSPChars(E1_s_tax_bill_hdr(S542_E1_vat_type))			& """" & vbCr
	Response.Write ".txtVATTypeNm.Value		= """ & ConvSPChars(E6_b_minor(S542_E6_minor_nm))					& """" & vbCr
    Response.Write ".txtSalesGroup.Value	= """ & ConvSPChars(E3_b_sales_grp(S542_E3_sales_grp))				& """" & vbCr
	Response.Write ".txtSalesGroupNm.Value	= """ & ConvSPChars(E3_b_sales_grp(S542_E3_sales_grp_nm))			& """" & vbCr
    Response.Write ".txtRemark.Value		= """ & ConvSPChars(E1_s_tax_bill_hdr(S542_E1_remarks))				& """" & vbCr

	If Cdbl(E1_s_tax_bill_hdr(S542_E1_net_amt))		<> 0 Then 
	   Response.Write ".btnPosting.disabled = False " & vbCr
	else
	   Response.Write ".btnPosting.disabled = True " & vbCr
	End If   

	If E1_s_tax_bill_hdr(S542_E1_post_flag)   = "Y" Then 
	   Response.Write ".btnPosting.value = ""발행취소""" & vbCr
	else
	   Response.Write ".btnPosting.value = ""발행""" & vbCr
	End If   
		
	Response.Write "parent.DbQueryOk" & vbCr
	Response.Write "End With"          & vbCr
    Response.Write "</Script>"         & vbCr
	
End Sub
'============================================
' Name : SubBizSave
' Desc : Save Data 
'============================================
Sub SubBizSave()
    Dim pS7G311
    Dim iCommandSent
    Dim itxtFlgMode

	Dim I1_s_tax_bill_hdr
	Dim I2_b_biz_partner_bp_cd
	Dim I3_b_sales_grp_sales_grp
	Dim I4_s_wks_user_user_id
	Dim I5_s_BillNoFlg
	
	Dim E2_s_tax_bill_hdr
	
    Const S537_I1_tax_bill_no = 0    'View Name : imp s_tax_bill_hdr
    Const S537_I1_tax_bill_type = 1
    Const S537_I1_issued_dt = 2
    Const S537_I1_vat_calc_type = 3
    Const S537_I1_vat_type = 4
    Const S537_I1_vat_rate = 5
    Const S537_I1_cur = 6
    Const S537_I1_report_biz_area = 7
    Const S537_I1_bill_no = 8
    Const S537_I1_remarks = 9
    Const S537_I1_ext1_qty = 10
    Const S537_I1_ext2_qty = 11
    Const S537_I1_ext3_qty = 12
    Const S537_I1_ext1_amt = 13
    Const S537_I1_ext2_amt = 14
    Const S537_I1_ext3_amt = 15
    Const S537_I1_ext1_cd = 16
    Const S537_I1_ext2_cd = 17
    Const S537_I1_ext3_cd = 18
    Const S537_I1_created_meth = 19
    Const S537_I1_history_flag = 20
    Const S537_I1_tax_doc_no = 21
    Const S537_I1_vat_inc_flag = 22

    'Const S537_I2_bp_cd = 0    'View Name : imp b_biz_partner

    'Const S537_I3_sales_grp = 0    'View Name : imp b_sales_grp

    'Const S537_I4_user_id = 0    'View Name : imp s_wks_user

    Const S537_E2_tax_bill_no = 0    'View Name : exp s_tax_bill_hdr

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status        
    
	ReDim I1_s_tax_bill_hdr(S537_I1_vat_inc_flag)
	'ReDim I2_b_biz_partner(S537_I2_bp_cd)
	'ReDim I3_b_sales_grp(S537_I3_sales_grp)
	'ReDim I4_s_wks_user(S537_I4_user_id)
	ReDim E2_s_tax_bill_hdr(S537_E2_tax_bill_no)

	I1_s_tax_bill_hdr(S537_I1_tax_bill_no)   = Trim(Request("txtTaxbillNo1"))
	
	'If Trim(Request("txtBillNoFlg")) = "Y" Then
		I1_s_tax_bill_hdr(S537_I1_bill_no) = Trim(Request("txtBillNo"))
	'Else 
	'	I1_s_tax_bill_hdr(S537_I1_bill_no) = ""
	'End If		

	
	'세금계산서번호 
	I1_s_tax_bill_hdr(S537_I1_tax_doc_no) = Trim(Request("txtTaxbillDocNo"))
	
	'세금계산서 번호 생성방법 
	I1_s_tax_bill_hdr(S537_I1_created_meth) = Trim(Request("txtMinor_cd"))
	
	'세금계산서번호의 History관리 여부 
	if Trim(Request("txtReference")) = "3" then
		I1_s_tax_bill_hdr(S537_I1_history_flag) = "Y"
	else 
		I1_s_tax_bill_hdr(S537_I1_history_flag) = "N"
	end if

	I2_b_biz_partner_bp_cd = Trim(Request("txtBillToParty"))
    I1_s_tax_bill_hdr(S537_I1_tax_bill_type) = Trim(Request("rdoTaxBillType"))
	I1_s_tax_bill_hdr(S537_I1_cur) = Trim(Request("txtCurrency"))
	
	I1_s_tax_bill_hdr(S537_I1_issued_dt) = UNIConvDate(Request("txtIssueDt"))

    I1_s_tax_bill_hdr(S537_I1_report_biz_area) = Trim(Request("txtTaxBizAreaCd"))

    I1_s_tax_bill_hdr(S537_I1_vat_calc_type) = Trim(Request("txtVatCalcType"))
    
	I1_s_tax_bill_hdr(S537_I1_vat_type) = Trim(Request("txtVATType"))
    I3_b_sales_grp_sales_grp = Trim(Request("txtSalesGroup"))
	I4_s_wks_user_user_id = Trim(Request("txtInsrtUserId"))
	'PIS
	I1_s_tax_bill_hdr(S537_I1_remarks) = Trim(Request("txtRemark"))
	I1_s_tax_bill_hdr(S537_I1_vat_rate) = UNIConvNum(Trim(Request("txtVATRate")),0)
	I1_s_tax_bill_hdr(S537_I1_vat_inc_flag) = Trim(Request("txtVATIncFlag"))
	
	I5_s_BillNoFlg = Trim(Request("txtBillNoFlg"))									' 매출채권번호지정 flag

	itxtFlgMode = CInt(Request("txtFlgMode"))										'☜: 저장시 Create/Update 판별 

    If itxtFlgMode = OPMD_CMODE Then
		iCommandSent = "CREATE"
    ElseIf itxtFlgMode = OPMD_UMODE Then
    	iCommandSent = "UPDATE"
    End If

    Set pS7G311 = Server.CreateObject("PS7G311.cSTaxBillHdrSvr")

	If CheckSYSTEMError(Err,True) = True Then
       Set pS7G311 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  

    E2_s_tax_bill_hdr(S537_E2_tax_bill_no) =  pS7G311.S_MAINT_TAX_BILL_HDR_SVR(gStrGlobalCollection, iCommandSent, I1_s_tax_bill_hdr, _
											I2_b_biz_partner_bp_cd, I3_b_sales_grp_sales_grp, I4_s_wks_user_user_id, I5_s_BillNoFlg)
    
	If CheckSYSTEMError(Err,True) = True Then
       Set pS7G311 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  
    
    Set pS7G311 = Nothing	

 			
	Response.Write "<Script language=vbs> " & vbCr         
	Response.Write "With Parent "               & vbCr
	If E2_s_tax_bill_hdr(S537_E2_tax_bill_no) <> "" Then
	'iStrCode = Replace(Trim(frm1.txtSoldToParty.value), "'", "''")
		'Response.Write "   .frm1.txtTaxbillNo.value = """   & Replace(ConvSPChars(E2_s_tax_bill_hdr(S537_E2_tax_bill_no)), "''", "'")    & """" & vbCr 
		Response.Write "   .frm1.txtTaxbillNo.value = """   & ConvSPChars(E2_s_tax_bill_hdr(S537_E2_tax_bill_no))    & """" & vbCr 
	Else
		Response.Write "   .frm1.txtTaxbillNo.value = .frm1.txtTaxbillNo1.value " & vbCr    
	End If
    Response.Write " .DbSaveOk " & vbCr
    Response.Write "End With"     & vbCr      
    Response.Write "</Script> "  
    
End Sub
'============================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================
Sub SubBizDelete()
    Dim pS7G311
    Dim iCommandSent
    Dim itxtFlgMode

	Dim I1_s_tax_bill_hdr
	Dim I2_b_biz_partner_bp_cd
	Dim I3_b_sales_grp_sales_grp
	Dim I4_s_wks_user_user_id
	Dim I5_s_BillNoFlg
	
    Const S537_I1_tax_bill_no = 0    'View Name : imp s_tax_bill_hdr
    Const S537_I1_vat_inc_flag = 22

    'Const S537_I2_bp_cd = 0    'View Name : imp b_biz_partner

    'Const S537_I3_sales_grp = 0    'View Name : imp b_sales_grp

    'Const S537_I4_user_id = 0    'View Name : imp s_wks_user

                
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
   
    ReDim I1_s_tax_bill_hdr(S537_I1_vat_inc_flag)
	'ReDim I2_b_biz_partner(S537_I2_bp_cd)
	'ReDim I3_b_sales_grp(S537_I3_sales_grp)
	'ReDim I4_s_wks_user(S537_I4_user_id)

    If Request("txtTaxbillNo") = "" Then										'⊙: 조회를 위한 값이 들어왔는지 체크 
	    Call ServerMesgBox("삭제 조건값이 비어있습니다!", vbInformation, I_MKSCRIPT)              
        Exit Sub
	End If

	I1_s_tax_bill_hdr(S537_I1_tax_bill_no) = Request("txtTaxbillNo")

    iCommandSent = "DELETE"
    
    Set pS7G311 = Server.CreateObject("PS7G311.cSTaxBillHdrSvr")
    
	If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If  
               
	call pS7G311.S_MAINT_TAX_BILL_HDR_SVR(gStrGlobalCollection, iCommandSent, I1_s_tax_bill_hdr, _
											I2_b_biz_partner_bp_cd, I3_b_sales_grp_sales_grp, I4_s_wks_user_user_id, I5_s_BillNoFlg)
    
    If CheckSYSTEMError(Err,True) = True Then
		Set pS7G311 = Nothing
		Exit Sub
	End If     
    '-----------------------
	'Result data display area
	'----------------------- 
	Set pS7G311 = Nothing
	
	Response.Write "<Script language=vbs> " & vbCr         
    Response.Write " Parent.DbDeleteOk "    & vbCr   
    Response.Write "</Script> "  
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub
'============================================
' Name : SubBizPost
' Desc : Post
'============================================
Sub SubBizPost()
    Dim pS7G315
    Dim itxtFlgMode

	Dim I1_s_tax_bill_hdr_tax_bill_no
	Dim I2_s_wks_user_user_id

    'Public Const S552_I1_tax_bill_no = 0    'View Name : imp s_tax_bill_hdr

    'Public Const S552_I2_user_id = 0    'View Name : imp s_wks_user
                
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
   
    If Request("txtTaxbillNo") = "" Then										'⊙: 조회를 위한 값이 들어왔는지 체크 
	    Call ServerMesgBox("발행에 필요한 조건값이 비어있습니다!", vbInformation, I_MKSCRIPT)              
        Exit Sub
	End If

	I1_s_tax_bill_hdr_tax_bill_no = Request("txtTaxbillNo")   

    
    Set pS7G315 = Server.CreateObject("PS7G315.cSPostTaxBillSvr")
    
	If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If  
               
	call pS7G315.S_POST_TAX_BILL_SVR(gStrGlobalCollection, I1_s_tax_bill_hdr_tax_bill_no, I2_s_wks_user_user_id)
    
    If CheckSYSTEMError(Err,True) = True Then
		Set pS7G315 = Nothing
		Exit Sub
	End If     
    '-----------------------
	'Result data display area
	'----------------------- 
	Set pS7G315 = Nothing
	
	Response.Write "<Script language=vbs> " & vbCr         
    Response.Write " Parent.PostOk "		& vbCr   
    Response.Write "</Script> "  
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub
'============================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================
Sub SubBizQueryMulti()
    
End Sub    

'============================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================
Sub SubBizSaveMulti()        
    
End Sub    

'============================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================
Sub SetErrorStatus()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================
Sub CommonOnTransactionAbort()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

%>

