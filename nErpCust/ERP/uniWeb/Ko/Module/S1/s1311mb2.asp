<%@ LANGUAGE=VBSCript %>
<%
'**********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : S1311MB2
'*  4. Program Name         : 품목할증등록 
'*  5. Program Desc         :
'*  6. Comproxy List        : PS1G107.dll, PS1G108.dll
'*  7. Modified date(First) : 2000/04/09
'*  8. Modified date(Last)  : 2001/12/18
'*  9. Modifier (First)     : Mr Cho
'* 10. Modifier (Last)      : son bumyeol
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              : 
'**********************************************************************************************
%>
<%Option Explicit%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../ComASP/LoadInfTb19029.asp" -->

<%
    Dim lgOpModeCRUD
    
	Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I","*","NOCOOKIE","MB")
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
    End Select

'============================================================================================================
Sub SubBizQueryMulti()

	Dim iLngRow	
	Dim iLngMaxRow
	Dim istrData
	Dim iStrPrevKey
	Dim pS1G108		'구 pS13118    
	Dim arrValue
	
    Dim I1_b_item 'imp_next_b_item
    Dim I2_s_item_dc 'imp_next_s_item_dc 
    Dim I3_s_item_dc 'imp_s_item_dc
    Dim I4_b_item 'imp_b_item
    Dim E1_b_item 'exp_b_item
    Dim E2_b_item 'exp_next_b_item
    Dim E3_s_item_dc 'exp_next_s_item_dc                    
    Dim EG1_exp_grp 'exp_grp
    Dim E4_b_minor 'exp_pay_meth_b_minor   
    Dim E5_b_unit_of_measure 'exp_b_unit_of_measure       

    Dim intGroupCount
    Dim StrNextKey	
    
    Const C_SHEETMAXROWS_D  = 100

    '[CONVERSION INFORMATION]  IMPORTS View 상수 
    Const S023_I2_valid_from_dt = 0    '[CONVERSION INFORMATION]  View Name : imp_next s_item_dc
    Const S023_I2_pay_meth = 1
    Const S023_I2_sales_unit = 2
    Const S023_I2_dc_bas_qty = 3

    Const S023_I3_valid_from_dt = 0    '[CONVERSION INFORMATION]  View Name : imp s_item_dc
    Const S023_I3_pay_meth = 1
    Const S023_I3_sales_unit = 2

    '[CONVERSION INFORMATION]  EXPORTS View 상수 
    Const S023_E1_item_cd = 0    '[CONVERSION INFORMATION]  View Name : exp b_item
    Const S023_E1_item_nm = 1

	Const S023_E2_item_cd = 0    '[CONVERSION INFORMATION]  View Name : exp_next b_item

    Const S023_E3_valid_from_dt = 0    '[CONVERSION INFORMATION]  View Name : exp_next s_item_dc
    Const S023_E3_pay_meth = 1
    Const S023_E3_sales_unit = 2
    Const S023_E3_dc_bas_qty = 3

    '[CONVERSION INFORMATION]  EXPORTS Group View 상수 
    '[CONVERSION INFORMATION] ===========================================================================
    '[CONVERSION INFORMATION]  Group Name : exp_grp
    Const S023_EG1_E1_item_cd = 0   '[CONVERSION INFORMATION]  View Name : exp_item b_item
    Const S023_EG1_E1_item_nm = 1
    Const S023_EG1_E2_pay_meth = 2  '[CONVERSION INFORMATION]  View Name : exp_item s_item_dc
    Const S023_EG1_E2_sales_unit = 3
    Const S023_EG1_E2_dc_bas_qty = 4
    Const S023_EG1_E2_valid_from_dt = 5
    Const S023_EG1_E2_dc_kind = 6
    Const S023_EG1_E2_dc_rate = 7
    Const S023_EG1_E2_round_type = 8
    Const S023_EG1_E2_ext1_qty = 9
    Const S023_EG1_E2_ext2_qty = 10
    Const S023_EG1_E2_ext1_amt = 11
    Const S023_EG1_E2_ext2_amt = 12
    Const S023_EG1_E2_ext1_cd = 13
    Const S023_EG1_E2_ext2_cd = 14
    Const S023_EG1_E3_minor_nm = 15 '[CONVERSION INFORMATION]  View Name : exp_item_dc_type b_minor
    Const S023_EG1_E4_minor_nm = 16 '[CONVERSION INFORMATION]  View Name : exp_item_pay_meth b_minor
    Const S023_EG1_E5_minor_nm = 17 '[CONVERSION INFORMATION]  View Name : exp_item_round_type b_minor
    Const S023_EG1_E5_spec	= 18

    '[CONVERSION INFORMATION]  EXPORTS View 상수 
    Const S023_E4_minor_cd = 0    '[CONVERSION INFORMATION]  View Name : exp_pay_meth b_minor
    Const S023_E4_minor_nm = 1
    
    Const S023_E5_unit = 0    '[CONVERSION INFORMATION]  View Name : exp b_unit_of_measure
    Const S023_E5_unit_nm = 1
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    I4_b_item = Trim(Request("txtconItem_cd"))
        
    ReDim I3_s_item_dc(S023_I3_sales_unit)

	I3_s_item_dc(S023_I3_valid_from_dt) = UNIConvDate(Request("txtconValid_from_dt"))
	I3_s_item_dc(S023_I3_pay_meth) = Trim(Request("txtconPay_terms"))
	I3_s_item_dc(S023_I3_sales_unit) = Trim(Request("txtconSales_unit"))

    ReDim I2_s_item_dc(S023_I2_dc_bas_qty)

	iStrPrevKey      = Request("lgStrPrevKey")                                      '☜: Next Key

	If iStrPrevKey <> "" then
		arrValue = Split(iStrPrevKey, gColSep)

		I1_b_item = Trim(arrValue(0))

		I2_s_item_dc(0) = UNIConvDate(Trim(arrValue(1)))
		I2_s_item_dc(1) = Trim(arrValue(2))
		I2_s_item_dc(2) = Trim(arrValue(3))
		I2_s_item_dc(3) = UNIConvNum(Trim(arrValue(4)),0)
	else
 		I1_b_item = ""

 		I2_s_item_dc(0) = ""
		I2_s_item_dc(1) = ""
		I2_s_item_dc(2) = ""
		I2_s_item_dc(3) = ""
	End If	


	Set pS1G108 = Server.CreateObject("PS1G108.CsListItemDcSvr")	 

	If CheckSYSTEMError(Err,True) = True Then		
       Exit Sub
    End If   

	Call pS1G108.S_LIST_ITEM_DC_SVR(gStrGlobalCollection, C_SHEETMAXROWS_D, I1_b_item, I2_s_item_dc, _
									I3_s_item_dc, I4_b_item, E1_b_item, E2_b_item, E3_s_item_dc, _
									EG1_exp_grp, E4_b_minor, E5_b_unit_of_measure)	
																					   
	If CheckSYSTEMError(Err,True) = True Then
	    Response.Write "<Script Language=vbscript>" & vbCr
		Response.Write "With Parent "               & vbCr
		Response.Write " .frm1.txtconItem_nm.value        = """ & ConvSPChars(E1_b_item(S023_E1_item_nm))                       & """" & vbCr    
		Response.Write " .frm1.txtconPay_terms_nm.value   = """ & ConvSPChars(E4_b_minor(S023_E4_minor_nm))  & """" & vbCr    
		Response.Write " .frm1.txtconSales_unit_nm.value  = """ & ConvSPChars(E5_b_unit_of_measure(S023_E5_unit_nm)) & """" & vbCr    
        Response.Write " .frm1.txtconItem_cd.focus " & vbCr    
		Response.Write "End With"          & vbCr
		Response.Write "</Script>"         & vbCr

       Set pS1G108 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If   

    Set pS1G108 = Nothing	     

    iLngMaxRow  = CLng(Request("txtMaxRows"))                                  '☜: Fetechd Count    

    'ReDim E2_b_item(S023_E2_item_cd)

	For iLngRow = 0 To UBound(EG1_exp_grp,1)
        If  iLngRow < C_SHEETMAXROWS_D  Then
        
        Else        
		    StrNextKey = ConvSPChars(EG1_exp_grp(iLngRow, S023_EG1_E1_item_cd))    
            StrNextKey = StrNextKey & gColSep &	UNIDateClientFormat(EG1_exp_grp(iLngRow, S023_EG1_E2_valid_from_dt))
            StrNextKey = StrNextKey & gColSep & ConvSPChars(EG1_exp_grp(iLngRow, S023_EG1_E2_pay_meth))
            StrNextKey = StrNextKey & gColSep & ConvSPChars(EG1_exp_grp(iLngRow, S023_EG1_E2_sales_unit))
            StrNextKey = StrNextKey & gColSep & UNINumClientFormat(EG1_exp_grp(iLngRow, S023_EG1_E2_dc_bas_qty), ggQty.DecPoint, 0)
           Exit For
        End If
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, S023_EG1_E1_item_cd))
		istrData = istrData & Chr(11)
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, S023_EG1_E1_item_nm))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, S023_EG1_E5_spec))        
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, S023_EG1_E2_pay_meth))
		istrData = istrData & Chr(11)
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, S023_EG1_E4_minor_nm))
        istrData = istrData & Chr(11) & UNIDateClientFormat(EG1_exp_grp(iLngRow, S023_EG1_E2_valid_from_dt))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, S023_EG1_E2_sales_unit))
		istrData = istrData & Chr(11)        
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow, S023_EG1_E2_dc_bas_qty), ggQty.DecPoint, 0)
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow, S023_EG1_E2_dc_rate), ggQty.DecPoint, 0)
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, S023_EG1_E2_dc_kind))
		istrData = istrData & Chr(11)        
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, S023_EG1_E3_minor_nm))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, S023_EG1_E2_round_type))
		istrData = istrData & Chr(11)        
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow, S023_EG1_E5_minor_nm))                        
        istrData = istrData & Chr(11) & iLngMaxRow + iLngRow
        istrData = istrData & Chr(11) & Chr(12)        

    Next            

	'ReDim E1_b_item(S023_E1_item_nm)

  
    '------ Developer Coding part (Start ) ------------------------------------------------------------------
    Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With Parent "               & vbCr
    Response.Write " .frm1.txtconItem_cd.value        = """ & ConvSPChars(E1_b_item(S023_E1_item_cd))                       & """" & vbCr    
    Response.Write " .frm1.txtconItem_nm.value        = """ & ConvSPChars(E1_b_item(S023_E1_item_nm))                       & """" & vbCr    
    Response.Write " .frm1.txtconPay_terms.value      = """ & ConvSPChars(E4_b_minor(S023_E4_minor_cd))  & """" & vbCr    
    Response.Write " .frm1.txtconPay_terms_nm.value   = """ & ConvSPChars(E4_b_minor(S023_E4_minor_nm))  & """" & vbCr    
    Response.Write " .frm1.txtconSales_unit.value     = """ & ConvSPChars(E5_b_unit_of_measure(S023_E5_unit))    & """" & vbCr    
    Response.Write " .frm1.txtconSales_unit_nm.value  = """ & ConvSPChars(E5_b_unit_of_measure(S023_E5_unit_nm)) & """" & vbCr    
    Response.Write " .SetSpreadColor1 -1 " & vbCr
    Response.Write " .frm1.txtHconItem_cd.value       = """ & I4_b_item                                   & """" & vbCr
    Response.Write " .frm1.txtHconPay_terms.value     = """ & I3_s_item_dc(S023_I3_pay_meth)      & """" & vbCr    
    Response.Write " .frm1.txtHconValid_from_dt.value = """ & I3_s_item_dc(S023_I3_valid_from_dt) & """" & vbCr    
    Response.Write " .frm1.txtHconSales_unit.value    = """ & I3_s_item_dc(S023_I3_sales_unit)    & """" & vbCr    

	Response.Write " .ggoSpread.Source = .frm1.vspdData"                                  & vbCr
	Response.Write " .ggoSpread.SSShowDataByClip     """ & istrData                      & """" & vbCr
	Response.Write " .lgStrPrevKey           = """ & StrNextKey                    & """" & vbCr
    Response.Write " .DbQueryOk " & vbCr
    Response.Write "End With"     & vbCr
    Response.Write "</Script>"    & vbCr    
	'------ Developer Coding part (End   ) ------------------------------------------------------------------     
      	
End Sub    

'============================================================================================================
Sub SubBizSaveMulti()        
	                                                                    
	Dim pS1G107		'구 pS13111
	Dim iErrorPosition	
	
    On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear																			 '☜: Clear Error status

	Set pS1G107 = Server.CreateObject("PS1G107.CsItemDcMultiSvr")

	If CheckSYSTEMError(Err,True) = true then 		
		Exit Sub
	End If

	Dim reqtxtSpread
	reqtxtSpread = Request("txtSpread")
	Call pS1G107.S_MAINT_ITEM_DC_MULTI_SVR(gStrGlobalCollection, Trim(reqtxtSpread), iErrorPosition)   

    If CheckSYSTEMError2(Err, True, iErrorPosition & "행","","","","") = True Then
       Set pS1G107 = Nothing
       Exit Sub
	End If

    Set pS1G107 = Nothing    

    Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "Parent.DBSaveOK "           & vbCr
    Response.Write "</Script>"                  & vbCr 

End Sub    

%>

