<%@ LANGUAGE=VBSCript%>
<% Option Explicit %>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->

<%
    'Dim lgOpModeCRUD
    
    On Error Resume Next                                                              
    
    Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*", "NOCOOKIE", "MB")   
	Call LoadBNumericFormatB("I","*","NOCOOKIE","MB")
	Call HideStatusWnd          
    '---------------------------------------Common-----------------------------------------------------------
	
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
            'Call SubBizQuery()
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
            'Call SubBizSave()
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
            'Call SubBizDelete()
    End Select


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    'On Error Resume Next                                                             
    Err.Clear                                                                         

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub    
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             
    Err.Clear                                                                        

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             
    Err.Clear                                                                        

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    
	Dim iLngRow	
	Dim iLngMaxRow
	Dim istrData
	Dim iStrPrevKey
    Dim iStrNextKey  	
	
	Dim iS4G219
	Dim I1_s_lc_amend_hdr
    Dim E1_b_minor  
    Dim E2_b_minor  
    Dim E3_b_minor
    Dim E4_b_minor  
    Dim E5_b_minor 
    Dim E6_b_minor 
    Dim E7_b_minor  
    Dim E8_b_minor  
    Dim E9_s_lc_hdr 
    Dim E10_b_sales_org 
    Dim E11_b_sales_grp 
    Dim E12_s_lc_amend_hdr
    Dim E13_b_biz_partner 
    Dim E14_b_biz_partner 
    Dim E15_b_biz_partner
    Dim E16_b_bank 
    Dim E17_b_bank  
    Dim E18_b_biz_partner 
    Dim E19_s_lc_dtl 
	
    Dim iS4G228   

    Const C_SHEETMAXROWS_D  = 100
    Dim   I1_s_lc_amend_hdr2 
    Dim   I2_s_lc_amend_dtl
    Dim   EG1_exp_grp 
 	    
    Const S375_I1_lc_amd_no = 0    '[CONVERSION INFORMATION]  View Name : imp s_lc_amend_hdr
    Const S375_I1_lc_kind = 1

    '[CONVERSION INFORMATION]  EXPORTS View 상수 
    
    Const S375_E1_minor_nm = 0    '[CONVERSION INFORMATION]  View Name : exp_pay_meth_nm b_minor

    Const S375_E2_minor_nm = 0    '[CONVERSION INFORMATION]  View Name : exp_incoterms_nm b_minor
    
    Const S375_E3_minor_nm = 0    '[CONVERSION INFORMATION]  View Name : exp_be_transport_nm b_minor
    
    Const S375_E4_minor_nm = 0    '[CONVERSION INFORMATION]  View Name : exp_be_discharge_port_nm b_minor
    
    Const S375_E5_minor_nm = 0    '[CONVERSION INFORMATION]  View Name : exp_be_loading_port_nm b_minor
    
    Const S375_E6_minor_nm = 0    '[CONVERSION INFORMATION]  View Name : exp_at_transport_nm b_minor

    Const S375_E7_minor_nm = 0    '[CONVERSION INFORMATION]  View Name : exp_at_discharge_port_nm b_minor

    Const S375_E8_minor_nm = 0    '[CONVERSION INFORMATION]  View Name : exp_at_loading_port_nm b_minor

    Const S375_E9_lc_no = 0    '[CONVERSION INFORMATION]  View Name : exp s_lc_hdr
    Const S375_E9_so_no = 1
    Const S375_E9_incoterms = 2
    Const S375_E9_pay_meth = 3

    Const S375_E10_sales_org = 0    '[CONVERSION INFORMATION]  View Name : exp b_sales_org
    Const S375_E10_sales_org_nm = 1

    Const S375_E11_sales_grp = 0    '[CONVERSION INFORMATION]  View Name : exp b_sales_grp
    Const S375_E11_sales_grp_nm = 1

    Const S375_E12_lc_amd_no = 0    '[CONVERSION INFORMATION]  View Name : exp s_lc_amend_hdr
    Const S375_E12_lc_doc_no = 1
    Const S375_E12_lc_amend_seq = 2
    Const S375_E12_adv_no = 3
    Const S375_E12_pre_adv_ref = 4
    Const S375_E12_open_dt = 5
    Const S375_E12_be_expiry_dt = 6
    Const S375_E12_at_expiry_dt = 7
    Const S375_E12_manufacturer = 8
    Const S375_E12_agent = 9
    Const S375_E12_amend_dt = 10
    Const S375_E12_amend_req_dt = 11
    Const S375_E12_cur = 12
    Const S375_E12_be_lc_amt = 13
    Const S375_E12_at_lc_amt = 14
    Const S375_E12_at_xch_rate = 15
    Const S375_E12_inc_amt = 16
    Const S375_E12_dec_amt = 17
    Const S375_E12_be_loc_amt = 18
    Const S375_E12_at_loc_amt = 19
    Const S375_E12_be_latest_ship_dt = 20
    Const S375_E12_at_latest_ship_dt = 21
    Const S375_E12_be_xch_rate = 22
    Const S375_E12_remark = 23
    Const S375_E12_be_loading_port = 24
    Const S375_E12_at_loading_port = 25
    Const S375_E12_be_dischge_port = 26
    Const S375_E12_at_dischge_port = 27
    Const S375_E12_be_transport = 28
    Const S375_E12_at_transport = 29
    Const S375_E12_remark2 = 30
    Const S375_E12_be_partial_ship_flag = 31
    Const S375_E12_at_partial_ship_flag = 32
    Const S375_E12_lc_kind = 33
    Const S375_E12_be_trnshp_flag = 34
    Const S375_E12_at_trnshp_flag = 35
    Const S375_E12_be_transfer_flag = 36
    Const S375_E12_at_transfer_flag = 37
    Const S375_E12_advise_bank = 38
    Const S375_E12_ext1_qty = 39
    Const S375_E12_ext2_qty = 40
    Const S375_E12_ext3_qty = 41
    Const S375_E12_ext1_amt = 42
    Const S375_E12_ext2_amt = 43
    Const S375_E12_ext3_amt = 44
    Const S375_E12_ext1_cd = 45
    Const S375_E12_ext2_cd = 46
    Const S375_E12_ext3_cd = 47
    Const S375_E12_xch_rate_op = 48

    Const S375_E13_bp_nm = 0    '[CONVERSION INFORMATION]  View Name : exp_agent b_biz_partner
    
    Const S375_E14_bp_nm = 0    '[CONVERSION INFORMATION]  View Name : exp_manufacturer b_biz_partner
    
    Const S375_E15_bp_nm = 0    '[CONVERSION INFORMATION]  View Name : exp_applicant b_biz_partner
    Const S375_E15_bp_cd = 1
    
    Const S375_E16_bank_nm = 0    '[CONVERSION INFORMATION]  View Name : exp_issue b_bank
    Const S375_E16_bank_cd = 1

    Const S375_E17_bank_nm = 0    '[CONVERSION INFORMATION]  View Name : exp_advise b_bank
    
    Const S375_E18_bp_nm = 0    '[CONVERSION INFORMATION]  View Name : exp_beneficiary b_biz_partner
    Const S375_E18_bp_cd = 1



    Const S375_E19_lc_amt = 0    '[CONVERSION INFORMATION]  View Name : exp_tot s_lc_dtl
    Const S375_E19_lc_loc_amt = 1


'   iS4G228
'    Const S383_I1_lc_amd_no = 0    '[CONVERSION INFORMATION]  View Name : imp s_lc_amend_hdr

 '   Const S383_I2_lc_amd_seq = 0    '[CONVERSION INFORMATION]  View Name : imp_next s_lc_amend_dtl

    '[CONVERSION INFORMATION]  Group Name : exp_grp
    Const S383_EG1_E1_lc_seq = 0    '[CONVERSION INFORMATION]  View Name : exp_item s_lc_dtl
    Const S383_EG1_E2_lc_no = 1    '[CONVERSION INFORMATION]  View Name : exp_item s_lc_hdr
    Const S383_EG1_E3_so_no = 2    '[CONVERSION INFORMATION]  View Name : exp_item s_so_hdr
    Const S383_EG1_E4_plant_cd = 3    '[CONVERSION INFORMATION]  View Name : exp_item b_plant
    Const S383_EG1_E4_plant_nm = 4
    Const S383_EG1_E5_lc_amd_seq = 5    '[CONVERSION INFORMATION]  View Name : exp_item s_lc_amend_dtl
    Const S383_EG1_E5_hs_cd = 6
    Const S383_EG1_E5_be_qty = 7
    Const S383_EG1_E5_at_qty = 8
    Const S383_EG1_E5_be_price = 9
    Const S383_EG1_E5_at_price = 10
    Const S383_EG1_E5_be_doc_amt = 11
    Const S383_EG1_E5_at_doc_amt = 12
    Const S383_EG1_E5_be_loc_amt = 13
    Const S383_EG1_E5_at_loc_amt = 14
    Const S383_EG1_E5_lc_unit = 15
    Const S383_EG1_E5_over_tolerance = 16
    Const S383_EG1_E5_under_tolerance = 17
    Const S383_EG1_E5_lc_kind = 18
    Const S383_EG1_E5_amend_flag = 19
    Const S383_EG1_E5_ext1_qty = 20
    Const S383_EG1_E5_ext2_qty = 21
    Const S383_EG1_E5_ext3_qty = 22
    Const S383_EG1_E5_ext1_amt = 23
    Const S383_EG1_E5_ext2_amt = 24
    Const S383_EG1_E5_ext3_amt = 25
    Const S383_EG1_E5_ext1_cd = 26
    Const S383_EG1_E5_ext2_cd = 27
    Const S383_EG1_E5_ext3_cd = 28
    Const S383_EG1_E6_item_nm = 29    '[CONVERSION INFORMATION]  View Name : exp_item b_item
    Const S383_EG1_E6_item_cd = 30
    Const S383_EG1_E7_so_seq = 31    '[CONVERSION INFORMATION]  View Name : exp_item s_so_dtl
    Const S383_EG1_E5_tracking_no = 32
    Const S383_EG1_E6_spec = 33

    '[CONVERSION INFORMATION] ===========================================================================
'    Const S383_E1_lc_amd_seq = 0    '[CONVERSION INFORMATION]  View Name : exp_next s_lc_amend_dtl

    On Error Resume Next                                                             
    Err.Clear                                                                        
	
	
    'iS4G219	
	If Request("txtLCAmdNo") = "" Then											
	    Call ServerMesgBox("조회 조건값이 비어있습니다!", vbInformation, I_MKSCRIPT)
		Exit sub   
	End If

	If Request("txtMaxRows") = 0 Then
		Redim I1_s_lc_amend_hdr(S375_I1_lc_kind)
	    I1_s_lc_amend_hdr(S375_I1_lc_amd_no) = Trim(Request("txtLCAmdNo"))
	    I1_s_lc_amend_hdr(S375_I1_lc_kind) = "M"
	

	    Set iS4G219 = Server.CreateObject("PS4G219.cSLkLcAmendHdrSvr")	

	    If CheckSYSTEMError(Err,True) = True Then
            Exit Sub
        End If
   
        Call iS4G219.S_LOOKUP_LC_AMEND_HDR_SVR(gStrGlobalCollection , "LOOKUP" , I1_s_lc_amend_hdr, _
                 E1_b_minor  ,  E2_b_minor  ,   E3_b_minor  ,   E4_b_minor  ,   E5_b_minor , _
                 E6_b_minor  ,  E7_b_minor  ,   E8_b_minor  ,   E9_s_lc_hdr ,   E10_b_sales_org  , _
                 E11_b_sales_grp  ,   E12_s_lc_amend_hdr , E13_b_biz_partner , E14_b_biz_partner , E15_b_biz_partner, _
                 E16_b_bank ,   E17_b_bank  ,   E18_b_biz_partner , E19_s_lc_dtl )

        If CheckSYSTEMError(Err,True) = True Then
           Set iS4G219 = Nothing		                                                 
           
           Response.Write "<Script language=vbs>  " & vbCr   
		   Response.Write " With Parent	       " & vbCr
	       Response.Write "   .frm1.txtLCAmdNo.focus  " & vbCr
	       Response.Write " End With      " & vbCr															    	
		   Response.Write "</Script>      " & vbCr     
           
           Exit Sub
        End If   

        Set iS4G219 = Nothing		


		'-----------------------
		'Result data display area
		'-----------------------
		Dim lgCurrency
			lgCurrency = ConvSPChars(E12_s_lc_amend_hdr(S375_E12_cur))
				

        Response.Write "<Script Language=VBScript>  " & vbCr   
        Response.Write " With parent.frm1	        " & vbCr
		Dim strDt

		'##### Rounding Logic #####
		'항상 거래화폐가 우선 
		
		Response.Write ".txtCurrency.value  =  """ & lgCurrency & """" & vbCr
		Response.Write ".txtCurrency1.value =  """ & lgCurrency & """" & vbCr

		Response.Write "parent.CurFormatNumericOCX " & vbCr
		'##########################
		
		Response.Write ".txtLCDocNo.value     =  """ & ConvSPChars(E12_s_lc_amend_hdr(S375_E12_lc_doc_no)) & """" & vbCr
		Response.Write ".txtLCAmendSeq.value  =  """ & ConvSPChars(E12_s_lc_amend_hdr(S375_E12_lc_amend_seq)) & """" & vbCr
		Response.Write ".txtLCNo.value        =  """ & ConvSPChars(E9_s_lc_hdr(S375_E9_lc_no)) & """" & vbCr
		Response.Write ".txtApplicant.value   =  """ & ConvSPChars(E15_b_biz_partner(S375_E15_bp_cd)) & """" & vbCr
		Response.Write ".txtApplicantNm.value =  """ & ConvSPChars(E15_b_biz_partner(S375_E15_bp_nm)) & """" & vbCr

		Response.Write ".txtAmendDt.text      =  """ & UNIDateClientFormat(E12_s_lc_amend_hdr(S375_E12_amend_dt)) & """" & vbCr

		Response.Write ".txtCurrency.value    =  """ & ConvSPChars(E12_s_lc_amend_hdr(S375_E12_cur)) & """" & vbCr
		Response.Write ".txtCurrency1.value   =  """ & ConvSPChars(E12_s_lc_amend_hdr(S375_E12_cur)) & """" & vbCr
		
		'##### Rounding Logic #####
		'.txtDocAmt.text =  """ & UNINumClientFormat(E12_s_lc_amend_hdr(ExpSLcAmendHdrAtLcAmt, ggAmtOfMoney.DecPoint, 0)) & """" & vbCr
		Response.Write ".txtDocAmt.text	      =  """ & UNINumClientFormatByCurrency(E12_s_lc_amend_hdr(S375_E12_at_lc_amt),lgCurrency,ggAmtOfMoneyNo) & """" & vbCr

		'.txtTotItemAmt.text =  """ & UNINumClientFormat(S32219.ExpTotSLcDtlLcAmt, ggAmtOfMoney.DecPoint, 0)) & """" & vbCr
		Response.Write ".txtTotItemAmt.text	  =  """ & UNINumClientFormatByCurrency(E19_s_lc_dtl(S375_E19_lc_amt),lgCurrency,ggAmtOfMoneyNo) & """" & vbCr

		'##########################
		Response.Write ".txtHBeDocAmt.value	  =  """ & UNINumClientFormatByCurrency(E19_s_lc_dtl(S375_E19_lc_amt),lgCurrency,ggAmtOfMoneyNo) & """" & vbCr
		Response.Write ".txtMaxSeq.value      = 0  " & vbCr
		Response.Write ".txtHLCAmdNo.value    =  """ & ConvSPChars(Request("txtLCAmdNo")) & """" & vbCr
		Response.Write ".txtHLCNo.value       =  """ & ConvSPChars(E9_s_lc_hdr(S375_E9_lc_no)) & """" & vbCr
		Response.Write ".txtHSONo.value       =  """ & ConvSPChars(E9_s_lc_hdr(S375_E9_so_no)) & """" & vbCr
		Response.Write ".txtHSalesGroup.value =  """ & ConvSPChars(E11_b_sales_grp(S375_E11_sales_grp)) & """" & vbCr
		Response.Write ".txtHSalesGroupNm.value =  """ & ConvSPChars(E11_b_sales_grp(S375_E11_sales_grp_nm)) & """" & vbCr
		Response.Write ".txtHPayTerms.value   =  """ & ConvSPChars(E9_s_lc_hdr(S375_E9_pay_meth)) & """" & vbCr
		Response.Write ".txtHPayTermsNm.value =  """ & ConvSPChars(E1_b_minor(S375_E1_minor_nm)) & """" & vbCr
		Response.Write ".txtHIncoTerms.value  =  """ & ConvSPChars(E9_s_lc_hdr(S375_E9_incoterms)) & """" & vbCr
		Response.Write ".txtHIncoTermsNm.value =  """ & ConvSPChars(E2_b_minor(S375_E2_minor_nm)) & """" & vbCr
		
		
		'-----------------------
		' Rounding Column Set
		'----------------------- 
		'Response.Write "parent.CurFormatNumSprSheet   " & vbCr
		Response.Write "If parent.lgIntFlgMode = parent.parent.OPMD_CMODE Then parent.CurFormatNumSprSheet " & vbCr	
		Response.Write "Call parent.LCAmendQueryOk() " & vbCr
		'Response.Write "Call parent.parent.MASetToolBar(" & """11101011000011""" & ") " & vbCr
	    Response.Write "End With                      " & vbCr
        Response.Write "</Script>	                  " & vbCr

	
	End if



'iS4G228
	I1_s_lc_amend_hdr2 = Trim(Request("txtLCAmdNo"))
'	I2_s_lc_amend_dtl =  "M"
    iStrPrevKey      = Trim(Request("lgStrPrevKey"))	'☜: Next Key				
	Dim iarrValue
	If iStrPrevKey <> "" then					
		iarrValue = Split(iStrPrevKey, gColSep)
		I2_s_lc_amend_dtl = Trim(iarrValue(0))
	else			
		I2_s_lc_amend_dtl = ""	
	End If		        

	Set iS4G228 = Server.CreateObject("PS4G228.cSLtLcAmendDtlSvr")	

	If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If
 
    Call iS4G228.S_LIST_LC_AMEND_DTL_SVR( gStrGlobalCollection, C_SHEETMAXROWS_D , I1_s_lc_amend_hdr2 , I2_s_lc_amend_dtl, _
                                                  EG1_exp_grp )	
    
    If CheckSYSTEMError(Err,True) = True Then
       Set iS4G228 = Nothing		                                                 
       
       Response.Write "<Script language=vbs>  " & vbCr   
	   Response.Write " With Parent	       " & vbCr
	   Response.Write "   .frm1.txtLCAmdNo.focus  " & vbCr
	   Response.Write " End With      " & vbCr															    	
	   Response.Write "</Script>      " & vbCr     
       
       Exit Sub
    End If   

    Set iS4G228 = Nothing	
   
    istrData = ""
    iLngMaxRow  = CInt(Request("txtMaxRows"))										 '☜: Fetechd Count      
	For iLngRow = 0 To UBound(EG1_exp_grp,1)
		If  iLngRow < C_SHEETMAXROWS_D  Then
		Else
		   iStrNextKey = ConvSPChars(EG1_exp_grp(iLngRow,S383_EG1_E1_lc_seq)) 
           Exit For
        End If 

        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow,S383_EG1_E5_amend_flag))
               Select Case Trim(ConvSPChars(EG1_exp_grp(iLngRow,S383_EG1_E5_amend_flag)))
        Case "C"
			istrData = istrData & Chr(11) & "품목추가"
        Case "U"
			istrData = istrData & Chr(11) & "내용변경"
        Case "D"
			istrData = istrData & Chr(11) & "품목삭제"
        Case Else
			istrData = istrData & Chr(11) & ""
        End Select
			
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow,S383_EG1_E6_item_cd))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow,S383_EG1_E6_item_nm))	
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow,S383_EG1_E6_spec ))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow,S383_EG1_E5_lc_unit))
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow,S383_EG1_E5_be_qty), ggQty.DecPoint, 0)										'3
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow,S383_EG1_E5_at_qty), ggQty.DecPoint, 0)										'3
        
        '##### Rounding Logic #####
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow,S383_EG1_E5_be_price), ggUnitCost.DecPoint, 0)		'6
        'istrData = istrData & Chr(11) & UNINumClientFormatByCurrency(EG1_exp_grp(iLngRow,ExpItemSLcAmendDtlBePrice(iLngRow), lgCurrency, ggUnitCostNo)

        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow,S383_EG1_E5_at_price), ggUnitCost.DecPoint, 0)		'6
		'istrData = istrData & Chr(11) & UNINumClientFormatByCurrency(EG1_exp_grp(iLngRow,ExpItemSLcAmendDtlAtPrice(iLngRow), lgCurrency, ggUnitCostNo)

		'istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow,ExpItemSLcAmendDtlBeDocAmt(iLngRow), ggAmtOfMoney.DecPoint, 0)
		istrData = istrData & Chr(11) & UNINumClientFormatByCurrency(EG1_exp_grp(iLngRow,S383_EG1_E5_be_doc_amt), lgCurrency, ggAmtOfMoneyNo)

		'istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow,ExpItemSLcAmendDtlAtDocAmt(iLngRow), ggAmtOfMoney.DecPoint, 0)
        istrData = istrData & Chr(11) & UNINumClientFormatByCurrency(EG1_exp_grp(iLngRow,S383_EG1_E5_at_doc_amt), lgCurrency, ggAmtOfMoneyNo)

        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow,S383_EG1_E5_hs_cd))														
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow,S383_EG1_E5_over_tolerance), ggExchRate.DecPoint, 0)	'10
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow,S383_EG1_E5_under_tolerance), ggExchRate.DecPoint, 0)	'11
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow,S383_EG1_E1_lc_seq))	
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow,S383_EG1_E3_so_no))															
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow,S383_EG1_E7_so_seq))															
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow,S383_EG1_E5_lc_amd_seq))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow,S383_EG1_E5_tracking_no))	
        '2002-12-24 CInt(iLngMaxRow)
        istrData = istrData & Chr(11) & iLngMaxRow + iLngRow
        istrData = istrData & Chr(11) & Chr(12)
    
    Next    

       
    Response.Write "<Script language=vbs>  " & vbCr   
    Response.Write " With Parent	       " & vbCr
	Response.Write "   .ggoSpread.Source = .frm1.vspdData         " & vbCr
	Response.Write "   .ggoSpread.SSShowDataByClip      """ & istrData    & """" & vbCr
	Response.Write "   .frm1.vspdData.ReDraw = False  "   & vbCr
	Response.Write "   .SetSpreadColor """ & iLngMaxRow + 1 & """" & vbCr
	Response.Write "    Call parent.SumItemVal()    " & vbCr
	Response.Write "   .lgStrPrevKey  =           """ & iStrNextKey & """" & vbCr  
	Response.Write "   .DbQueryOk                         " & vbCr
	Response.Write "   .frm1.vspdData.ReDraw = True " & vbCr
	Response.Write "   .frm1.txtHLCAmdNo.value =  """ & ConvSPChars(Request("txtLCAmdNo")) & """" & vbCr
    Response.Write " End With      " & vbCr															    	
    Response.Write "</Script>      " & vbCr      
           	
End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()   
		                                                                    
	Dim iS4G221
	Dim iErrorPosition
	Dim I1_s_lc_hdr_no
	Dim I2_s_lc_amend_hdr_amd_no
    On Error Resume Next                                                                 
    Err.Clear																			                                                             
	If Request("txtLCAmdNo") = "" Then										
		Call ServerMesgBox("조회 조건값이 비어있습니다!", vbInformation, I_MKSCRIPT)
		Exit sub
	End If

	I1_s_lc_hdr_no = UCase(Trim(Request("txtLCNo")))
    I2_s_lc_amend_hdr_amd_no =UCase(Trim(Request("txtHLCAmdNo")))
   	Set iS4G221 = CreateObject("PS4G221.cSLcAmendDtlSvr")
    
    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If
'    Response.Write "ee"
    Call iS4G221.S_MAINT_LC_AMEND_DTL_SVR(gStrGlobalCollection,I1_s_lc_hdr_no,I2_s_lc_amend_hdr_amd_no , _
             Request("txtSpread"),"",iErrorPosition)
    'Response.Write "ff"	& iErrorPosition				      
    If CheckSYSTEMError2(Err, True, iErrorPosition & "행","","","","") = True Then
       Set iS4G221 = Nothing
       Exit Sub
	End If
	
    Set iS4G221 = Nothing
                                                       
    Response.Write "<Script language=vbs> " & vbCr         
    Response.Write " Parent.DBSaveOk "      & vbCr   
    Response.Write "</Script> "             & vbCr   
              
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
