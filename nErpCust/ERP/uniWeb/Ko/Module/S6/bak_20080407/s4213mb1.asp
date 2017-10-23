<%
'********************************************************************************************************
'*  1. Module Name          : 영업관리																	*
'*  2. Function Name        :																			*
'*  3. Program ID           : s4213mb1.asp																*
'*  4. Program Name         : 통관란등록																*
'*  5. Program Desc         : 통관란등록																*
'*  6. Comproxy List        :																			*
'*  7. Modified date(First) : 2000/04/11																*
'*  8. Modified date(Last)  : 2001/12/18																*
'*  9. Modifier (First)     : Kim Hyungsuk																*
'* 10. Modifier (Last)      : Kim Hyungsuk																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/04/11 : 화면 design												*
'*							  2. 2000/05/04 : Coding Start												*
'********************************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%

On Error Resume Next                                                             
Err.Clear                                                                        

Call LoadBasisGlobalInf()
Call LoadInfTB19029B( "I", "*", "NOCOOKIE", "MB")
Call LoadBNumericFormatB("I","*","NOCOOKIE","MB")
Call HideStatusWnd                                                               
'---------------------------------------Common-----------------------------------------------------------
	
lgOpModeCRUD = Request("txtMode")                                           

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query           
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update            
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete            
    End Select
'============================================================================================================
Sub SubBizQuery()                                                          
    Err.Clear                                                                    

    
End Sub    
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             
    Err.Clear                                                                        
End Sub
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             
    Err.Clear                                                                        
End Sub

'============================================================================================================
Sub SubBizQueryMulti()
    
	Dim iLngRow	
	Dim iLngMaxRow
	Dim istrData
	Dim iStrPrevKey

	Dim iS42119
    Dim iS42138

    On Error Resume Next                                                             
    Err.Clear                                                                        
    Dim i_CommandSent
    Dim i_s_cc_hdr_cc_no
    Dim exp_grp_s_cc_hdr   
	Dim lgCurrency
	
	Const c_exp_iv_dt = 2	
	Const c_exp_E1_cur = 38
	Const c_exp_E1_doc_amt = 39
	Const c_exp_fob_doc_amt = 40
	Const c_exp_E1_usd_xch_rate = 67
	Const c_exp_E1_loc_amt = 42
	Const c_exp_fob_loc_amt = 43
    Const c_exp_E1_xch_rate_op = 81
    Const c_exp_E2_bp_cd = 82
	Const c_exp_E2_bp_nm = 83
    
    'Update 2005-01-07 LSW
    'i_s_cc_hdr_cc_no = FilterVar(Trim(Request("txtCCNo")),"","SNM")
    i_s_cc_hdr_cc_no = Trim(Request("txtCCNo"))
	
	Select Case Request("txtPrevNext")

		Case "PREV"
			i_CommandSent = "PREV"
		Case "NEXT"
			i_CommandSent = "NEXT"
		Case Else 
			i_CommandSent = "LOOKUP"
			
	End Select    

	If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If   

    set iS42119 = CREATEOBJECT("PS6G219.cSLkExportCcHdrSvr")
    
	' Call the Dll   		
    Call iS42119.S_LOOKUP_EXPORT_CC_HDR_SVR ( gStrGlobalCollection, _
                           cstr(i_CommandSent), cstr(i_s_cc_hdr_cc_no), _
                           exp_grp_s_cc_hdr)

	If CheckSYSTEMError(Err,TRUE) = True Then
       Set iS42119 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If   

    lgCurrency = exp_grp_s_cc_hdr(c_exp_E1_cur)
                           
    Response.Write "<Script language=vbs> " & vbCr    
    Response.Write " Parent.frm1.txtCCCurrency.value  = """ & lgCurrency & """" & vbCr
    Response.Write " parent.CurFormatNumericOCX " & vbCr
    Response.Write " Parent.frm1.txtApplicant.value  = """ & ConvSPChars(exp_grp_s_cc_hdr(c_exp_E2_bp_cd)) & """" & vbCr
    Response.Write " Parent.frm1.txtApplicantNm.value  = """ & ConvSPChars(exp_grp_s_cc_hdr(c_exp_E2_bp_nm)) & """" & vbCr
    Response.Write " Parent.frm1.txtUsdXchRate.text  = """ & UNINumClientFormat(exp_grp_s_cc_hdr(c_exp_E1_usd_xch_rate), ggExchRate.DecPoint, 0) & """" & vbCr
    Response.Write " Parent.frm1.txtHExchRateOp.value  = """ & ConvSPChars(exp_grp_s_cc_hdr(c_exp_E1_xch_rate_op)) & """" & vbCr

	Response.Write "Parent.frm1.txtDocAmt.Text	= """ & UNIConvNumDBToCompanyByCurrency(exp_grp_s_cc_hdr(c_exp_E1_doc_amt), lgCurrency, ggAmtOfMoneyNo, "X" , "X")  & """" & vbCr
	Response.Write "Parent.frm1.txtLocAmt.text	= """ & UNIConvNumDBToCompanyByCurrency(exp_grp_s_cc_hdr(c_exp_E1_loc_amt), gCurrency, ggAmtOfMoneyNo, gLocRndPolicyNo , "X")  & """" & vbCr

    Response.Write " Parent.frm1.txtFOBDocAmt.text  = """ & UNIConvNumDBToCompanyByCurrency(exp_grp_s_cc_hdr(c_exp_fob_doc_amt), "USD", ggAmtOfMoneyNo, "X" , "X")  & """" & vbCr
    Response.Write " Parent.frm1.txtFOBLocAmt.Text  = """ & UNIConvNumDBToCompanyByCurrency(exp_grp_s_cc_hdr(c_exp_fob_loc_amt), gCurrency, ggAmtOfMoneyNo, gLocRndPolicyNo , "X")  & """" & vbCr

    Response.Write " Parent.frm1.txtHOpenDt.value  = """ & ConvSPChars(exp_grp_s_cc_hdr(c_exp_iv_dt)) & """" & vbCr
    Response.Write " Parent.frm1.txtHCCNo.value  = """ & ConvSPChars(Trim(Request("txtCCNo"))) & """" & vbCr
    Response.Write " Parent.frm1.txtMaxSeq.value = 0 "	& vbCr
    Response.Write " If parent.lgIntFlgMode = parent.parent.OPMD_CMODE Then parent.CurFormatNumSprSheet "	& vbCr              
    Response.Write "</Script> "	& vbCr      

    Set iS42119 = Nothing
    Dim i_s_cc_hdr1
    Dim i_s_cc_lan1
    Dim i_s_cc_hdr
    Dim i_s_cc_lan
    Dim imp_next_s_cc_lan    
    Dim exp_grp   
    
    Dim intGroupCount
    Dim StrNextKey  	
    Dim arrValue
    
    Const C_SHEETMAXROWS_D  = 100
    Const c_exp_b_hs_code_hs_cd = 0
    Const c_exp_b_hs_code_hs_nm = 1
    Const c_exp_s_cc_lan_lan_no = 2
    Const c_exp_s_cc_lan_qty = 3
    Const c_exp_s_cc_lan_doc_amt = 4
    Const c_exp_s_cc_lan_fob_doc_amt = 5
    Const c_exp_s_cc_lan_loc_amt = 6
    Const c_exp_s_cc_lan_fob_loc_amt = 7
    Const c_exp_s_cc_lan_unit = 8
    Const c_exp_s_cc_lan_net_weight = 9
    Const c_exp_s_cc_lan_packing_cnt = 10
    Const c_exp_s_cc_lan_ext1_qty = 11
    Const c_exp_s_cc_lan_ext2_qty = 12
    Const c_exp_s_cc_lan_ext3_qty = 13
    Const c_exp_s_cc_lan_ext1_amt = 14
    Const c_exp_s_cc_lan_ext2_amt = 15
    Const c_exp_s_cc_lan_ext3_amt = 16
    Const c_exp_s_cc_lan_ext1_cd = 17
    Const c_exp_s_cc_lan_ext2_cd = 18
    Const c_exp_s_cc_lan_ext3_cd = 19
    Const c_exp_s_cc_lan_cc_no = 20    
    
	Const c_exp_s_cc_lan_gross_weight = 21
	Const c_exp_s_cc_lan_measurement = 22
    'Update 2005-01-07 LSW
    'i_s_cc_hdr1 = FilterVar(Trim(Request("txtCCNo")),"","SNM")
    i_s_cc_hdr1 = Trim(Request("txtCCNo"))
    i_s_cc_lan1 = UNIConvNum(Request("lgStrPrevKey"),0)

	iStrPrevKey = Trim(Request("lgStrPrevKey"))                                        '☜: Next Key	

    i_s_cc_hdr = i_s_cc_hdr1
    i_s_cc_lan = i_s_cc_lan1
        
	set iS42138 = CREATEOBJECT("PS6G238.cSListExportCcLanSvr")

	If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If   

    Call iS42138.S_LIST_EXPORT_CC_LAN_SVR ( gStrGlobalCollection, _
                           C_SHEETMAXROWS_D, _
                           cstr(i_s_cc_hdr), cstr(i_s_cc_lan), imp_next_s_cc_lan, _
                           exp_grp)

	If CheckSYSTEMError(Err,TRUE) = True Then
       Set iS42138 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If   
    
    Set iS42138 = Nothing	
        
    iLngMaxRow  = CLng(Request("txtMaxRows"))										'☜: Fetechd Count      
        
    For iLngRow = 0 To UBound(exp_grp,1)
    		
        If  iLngRow < C_SHEETMAXROWS_D  Then
        
        Else

            StrNextKey = ConvSPChars(exp_grp(iLngRow, c_exp_s_cc_lan_lan_no))                
            Exit For
        End If 
       
        istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, c_exp_s_cc_lan_lan_no))
        istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, c_exp_b_hs_code_hs_cd))
        istrdata = istrdata & Chr(11) & UNINumClientFormat(exp_grp(iLngRow, c_exp_s_cc_lan_fob_doc_amt), ggAmtOfMoney.DecPoint, 0)
        istrdata = istrdata & Chr(11) & UNINumClientFormat(exp_grp(iLngRow, c_exp_s_cc_lan_fob_loc_amt), ggAmtOfMoney.DecPoint, 0)
        istrdata = istrdata & Chr(11) & UNINumClientFormat(exp_grp(iLngRow, c_exp_s_cc_lan_packing_cnt), ggAmtOfMoney.DecPoint, 0)
        istrdata = istrdata & Chr(11) & UNINumClientFormat(exp_grp(iLngRow, c_exp_s_cc_lan_gross_weight), ggAmtOfMoney.DecPoint, 0)
        istrdata = istrdata & Chr(11) & UNINumClientFormat(exp_grp(iLngRow, c_exp_s_cc_lan_measurement), ggAmtOfMoney.DecPoint, 0)        
        istrdata = istrdata & Chr(11) & UNINumClientFormat(exp_grp(iLngRow, c_exp_s_cc_lan_net_weight), ggAmtOfMoney.DecPoint, 0)
        istrdata = istrdata & Chr(11) & UNINumClientFormat(exp_grp(iLngRow, c_exp_s_cc_lan_qty), ggAmtOfMoney.DecPoint, 0)
        istrdata = istrdata & Chr(11) & UNINumClientFormat(exp_grp(iLngRow, c_exp_s_cc_lan_doc_amt), ggAmtOfMoney.DecPoint, 0)
        
        ' 공통 구문 
        istrData = istrData & Chr(11) & iLngMaxRow + iLngRow
        istrData = istrData & Chr(11) & Chr(12)       
        
        StrNextKey = ConvSPChars(exp_grp(iLngRow, c_exp_s_cc_lan_lan_no))
        
    Next 
    
    Response.Write "<Script language=vbs> " & vbCr
    Response.Write " Parent.SetSpreadColor -1 " & vbCr
    Response.Write " Parent.frm1.txtHCCNo.value  = """ & ConvSPChars(Trim(Request("txtCCNo"))) & """" & vbCr
    Response.Write " Parent.ggoSpread.Source     = Parent.frm1.vspdData	" & vbCr
    Response.Write " Parent.ggoSpread.SSShowData """ & istrData	& """" & vbCr
    Response.Write " Parent.lgStrPrevKey              = """ & StrNextKey & """" & vbCr  
    Response.Write " Parent.DbQueryOk "	& vbCr   
    Response.Write "</Script> "	& vbCr      
                   	
End Sub    
'============================================================================================================
Sub SubBizSaveMulti()   

Dim S42131	
Dim iErrorPosition
	
On Error Resume Next                                                                 
Err.Clear																			                                                             
     
	Set S42131 = Server.CreateObject("PS6G231.cSExportCcLanSvr")
	
    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If
    
    'Update 2005-01-07 LSW 
    'Call S42131.S_MAINT_EXPORT_CC_LAN_SVR  (gStrGlobalCollection, cstr(FilterVar(Trim(Request("txtHCCNo")),"","SNM")), _
	'											cstr(Trim(Request("txtSpread"))), iErrorPosition)
	Call S42131.S_MAINT_EXPORT_CC_LAN_SVR  (gStrGlobalCollection, cstr(Trim(Request("txtHCCNo"))), _
												cstr(Trim(Request("txtSpread"))), iErrorPosition)   

    If CheckSYSTEMError2(Err, True, iErrorPosition & "행","","","","") = True Then
       Set S42131 = Nothing
       Exit Sub
	End If
	
    Set S42131 = Nothing
                                                       
    Response.Write "<Script language=vbs> " & vbCr         
    Response.Write " Parent.DBSaveOk "      & vbCr   
    Response.Write "</Script> "             & vbCr   
                  
End Sub    
'============================================================================================================
Sub CommonOnTransactionAbort()
    Call SetErrorStatus()
End Sub

%>
