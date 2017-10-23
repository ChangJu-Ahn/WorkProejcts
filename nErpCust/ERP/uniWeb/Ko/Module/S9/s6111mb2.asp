<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<%
'********************************************************************************************************
'*  1. Module Name          : Sales																		*
'*  2. Function Name        : 판매경비관리																*
'*  3. Program ID           : S6111MA2																	*
'*  4. Program Name         : 판매경비일괄처리															*
'*  5. Program Desc         :																			*
'*  6. Comproxy List        : PS9G115.dll, PS9G241.dll
'*  7. Modified date(First) : 2000/04/26																*
'*  8. Modified date(Last)  : 2001/12/19																*
'*  9. Modifier (First)     : An Chang Hwan																*
'* 10. Modifier (Last)      : Cho Sung Hyun																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/04/26 : 화면 design												*
'*							  2. 2000/09/22 : 4th Coding Start											*
'*							  3. 2001/12/19 : Date 표준적용												*
'********************************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../ComASP/LoadInfTB19029.asp" -->
<%

On Error Resume Next                                                             '☜: Protect system from crashing
Err.Clear                                                                        '☜: Clear Error status

Call LoadBasisGlobalInf()
Call loadInfTB19029B("I", "*","NOCOOKIE","MB") 
Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")
Call HideStatusWnd                                                                 '☜: Hide Processing message

lgOpModeCRUD	=	Request("txtMode")

Select Case lgOpModeCRUD
Case CStr(UID_M0001)
	Call SubBizQueryMulti()
Case CStr(UID_M0002)
	Call SubBizSaveMulti()
Case CStr(UID_M0003)                                                         '☜: Delete

End Select

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

	Dim iS9G115
	Dim StrNextKey							' 다음 값 
	Dim lgStrPrevKey						' 이전 값 
	Dim ILngMaxRow							' 현재 그리드의 최대Row
	Dim ILngRow
	Dim istrData
	
	Dim I1_s_wks_date
    Dim I2_s_sales_charge
    Dim I3_b_biz_partner
    Dim I4_b_sales_grp
    Dim I5_s_sales_charge
    Dim EG1_exp_grp
    Dim E1_b_biz_partner
    Dim E2_b_sales_grp
    Dim E3_a_jnl_item
    Dim E4_b_minor
    Dim E5_s_sales_charge
    
    Const C_SHEETMAXROWS_D = 100
    Const I1_from_date = 0          'I1_s_wks_date
    Const I1_to_date = 1

    Const I2_process_step = 0       'I5_s_sales_charge
    Const I2_charge_cd = 1
    Const I2_posting_flag = 2

    Const EG1_bp_cd = 0				'EG1_exp_grp
    Const EG1_bp_nm = 1
    Const EG1_sales_grp = 2
    Const EG1_sales_grp_nm = 3
    Const EG1_process_step = 4
    Const EG1_process_step_nm = 5
    Const EG1_charge_cd = 6
    Const EG1_charge_nm = 7
    Const EG1_posting_flag = 8
    Const EG1_charge_no = 9
    Const EG1_cost_flag = 10
    Const EG1_bas_no = 11
    Const EG1_bas_doc_no = 12
    Const EG1_acct_trans_type = 13
    Const EG1_cur = 14
    Const EG1_charge_doc_amt = 15
    Const EG1_charge_loc_amt = 16
    Const EG1_xch_rate = 17
    Const EG1_charge_dt = 18
    Const EG1_vat_amt = 19
    Const EG1_vat_loc_amt = 20
    Const EG1_vat_type = 21
    Const EG1_vat_type_nm = 22
    Const EG1_vat_rate = 23
    Const EG1_pay_type = 24
    Const EG1_pay_type_nm = 25
    Const EG1_cost_cd = 26
    Const EG1_biz_area = 27
    Const EG1_note_no = 28
    Const EG1_remark = 29
    Const EG1_bank_cd = 30
    Const EG1_bank_nm = 31
    Const EG1_bank_acct_no = 32
    Const EG1_ext1_qty = 33
    Const EG1_ext2_qty = 34
    Const EG1_ext3_qty = 35
    Const EG1_ext1_amt = 36
    Const EG1_ext2_amt = 37
    Const EG1_ext3_amt = 38
    Const EG1_ext1_cd = 39
    Const EG1_ext2_cd = 40
    Const EG1_ext3_cd = 41
    
    On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear	
    
    ReDim I1_s_wks_date(1)
    ReDim I2_s_sales_charge(2)
    
    '-----------------------
    ' 판매경비일괄처리 내용을 읽어온다.
    '-----------------------
    Set iS9G115 = Server.CreateObject("PS9G115.cSLtSalesChargeSvr")

    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If   

    I1_s_wks_date(I1_from_date) = UNIConvDate(Request("txtFromDt"))
	I1_s_wks_date(I1_to_date) = UNIConvDate(Request("txtToDt"))

    I2_s_sales_charge(I2_charge_cd) = UCase(Trim(Request("txtCharge")))
	I2_s_sales_charge(I2_process_step) = UCase(Trim(Request("txtProcessStep")))

	If Len(Request("txtRadio")) Then
		I2_s_sales_charge(I2_posting_flag) = Request("txtRadio")    
	End If

    I3_b_biz_partner = UCase(Trim(Request("txtPayCharge")))    
    I4_b_sales_grp = UCase(Trim(Request("txtSalesGrp")))
	I5_s_sales_charge = Request("lgStrPrevKey")

    Call iS9G115.S_LIST_SALES_CHARGE_SVR(gStrGlobalCollection, C_SHEETMAXROWS_D, _
										I1_s_wks_date, I2_s_sales_charge, _
										I3_b_biz_partner, I4_b_sales_grp, _
										I5_s_sales_charge, EG1_exp_grp, _
										E1_b_biz_partner, E2_b_sales_grp, _
										E3_a_jnl_item, E4_b_minor, _
										E5_s_sales_charge) 
	
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent.frm1"			& vbCr
	Response.Write ".txtPayChargeNm.value		= """ & ConvSPChars(E1_b_biz_partner(0))	& """" & vbCr
	Response.Write ".txtSalesGrpNm.value		= """ & ConvSPChars(E2_b_sales_grp(0))		& """" & vbCr
	Response.Write ".txtChargeNm.value			= """ & ConvSPChars(E3_a_jnl_item(1))		& """" & vbCr
	Response.Write ".txtProcessStepNm.value		= """ & ConvSPChars(E4_b_minor(0))			& """" & vbCr
	Response.Write "End With"					& vbCr
	Response.Write "</Script>"					& vbCr    
    
    If CheckSYSTEMError(Err,True) = True Then
		Set iS9G115 = Nothing
		Response.Write "<Script Language=vbscript>"			& vbCr
	    Response.Write "parent.frm1.txtCharge.focus"		& vbCr    
	    Response.Write "</Script>"							& vbCr	
        Exit Sub
    End If   
            
	Set iS9G115 = Nothing	
    
    ILngMaxRow  = CLng(Request("txtMaxRows"))
    
	'-----------------------
	'Result data display area
	'----------------------- 
	For ILngRow = 0 To UBound(EG1_exp_grp, 1)
		If  ILngRow < C_SHEETMAXROWS_D  Then
		Else
		   StrNextKey = ConvSPChars(EG1_exp_grp(ILngRow, EG1_charge_no))
		   Exit For
        End If  
	
		istrData = istrData & Chr(11) & "0"   
			
		If ConvSPChars(EG1_exp_grp(ILngRow, EG1_posting_flag)) = "Y" Then
			istrData = istrData & Chr(11) & "1"
		Else
			istrData = istrData & Chr(11) & "0"
		End If

		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(ILngRow, EG1_charge_no))
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(ILngRow, EG1_charge_cd)) 					
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(ILngRow, EG1_process_step))						
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(ILngRow, EG1_process_step_nm))						
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(ILngRow, EG1_bas_no))						
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(ILngRow, EG1_sales_grp))						

		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(ILngRow, EG1_bp_nm))  
			
		istrData = istrData & Chr(11) & UNIDateClientFormat(EG1_exp_grp(ILngRow, EG1_charge_dt))
			
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(ILngRow, EG1_vat_type)) 
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(ILngRow, EG1_cur)) 
		istrData = istrData & Chr(11) & UniConvNumDBToCompanyWithOutChange(EG1_exp_grp(ILngRow, EG1_charge_doc_amt), 0)	
		istrData = istrData & Chr(11) & UniConvNumDBToCompanyWithOutChange(EG1_exp_grp(ILngRow, EG1_xch_rate), 0)		
		istrData = istrData & Chr(11) & UniConvNumDBToCompanyWithOutChange(EG1_exp_grp(ILngRow, EG1_charge_loc_amt), 0)	
		istrData = istrData & Chr(11) & UniConvNumDBToCompanyWithOutChange(EG1_exp_grp(ILngRow, EG1_vat_rate), 0)		
		istrData = istrData & Chr(11) & UniConvNumDBToCompanyWithOutChange(EG1_exp_grp(ILngRow, EG1_vat_amt), 0)		
		istrData = istrData & Chr(11) & UniConvNumDBToCompanyWithOutChange(EG1_exp_grp(ILngRow, EG1_vat_loc_amt), 0)	

		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(ILngRow, EG1_cost_flag)) 
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(ILngRow, EG1_pay_type)) 
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(ILngRow, EG1_note_no)) 
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(ILngRow, EG1_bank_acct_no))
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(ILngRow, EG1_bank_nm)) 
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(ILngRow, EG1_remark))  
			
		istrData = istrData & Chr(11) & ILngMaxRow + ILngRow                                
		istrData = istrData & Chr(11) & Chr(12)
			
	Next
		
		Response.Write "<Script language=vbs> "												& vbCr
		Response.Write "With parent"														& vbCr   
		Response.Write " .frm1.txtHCharge.value			= """ & ConvSPChars(Request("txtCharge"))		& """" & vbCr
		Response.Write " .frm1.txtHSalesGrp.value		= """ & ConvSPChars(Request("txtSalesGrp"))		& """" & vbCr    
		Response.Write " .frm1.txtHPayCharge.value		= """ & ConvSPChars(Request("txtHPayCharge"))	& """" & vbCr
		Response.Write " .frm1.txtHFromDt.value			= """ & Request("txtFromDt")					& """" & vbCr        
		Response.Write " .frm1.txtHToDt.value			= """ & Request("txtToDt")						& """" & vbCr        
		Response.Write " .frm1.txtHProcessStep.value	= """ & ConvSPChars(Request("txtProcessStep"))	& """" & vbCr
		Response.Write " .frm1.txtHRadio.value			= """ & ConvSPChars(Request("txtRadio"))		& """" & vbCr                               
		Response.Write " .ggoSpread.Source				= .frm1.vspdData"					& vbCr
		Response.Write " .frm1.vspdData.Redraw = False   "                     & vbCr      
		Response.Write " .ggoSpread.SSShowDataByClip	""" & istrData & """ ,""F""" & vbCr
		Response.Write " Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData," & ILngMaxRow + 1 & "," & ILngMaxRow + UBound(EG1_exp_grp,1) + 1  & ",.C_Cur,.C_ChargeDocAmt,""A"" ,""Q"",""X"",""X"")" & vbCr
		Response.Write " Call .ReFormatSpreadCellByCellByCurrency2(.Frm1.vspdData," & ILngMaxRow + 1 & "," & ILngMaxRow + UBound(EG1_exp_grp,1) + 1  & ",.Parent.gCurrency,.C_ChargeLocAmt,""A"" ,""Q"",""X"",""X"")" & vbCr
		Response.Write " Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData," & ILngMaxRow + 1 & "," & ILngMaxRow + UBound(EG1_exp_grp,1) + 1  & ",.C_Cur,.C_VATAmt,""A"" ,""Q"",""X"",""X"")" & vbCr
		Response.Write " Call .ReFormatSpreadCellByCellByCurrency2(.Frm1.vspdData," & ILngMaxRow + 1 & "," & ILngMaxRow + UBound(EG1_exp_grp,1) + 1  & ",.Parent.gCurrency,.C_VatLocAmt,""A"" ,""Q"",""X"",""X"")" & vbCr

		Response.Write " .lgStrPrevKey					= """ & StrNextKey					& """" & vbCr    
		Response.Write " .DbQueryOk "														& vbCr   
		Response.Write " .frm1.vspdData.Redraw = True   "                      & vbCr              
		Response.Write "End With "															& vbCr   
		Response.Write "</Script> "		
		
End Sub

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti() 

	Dim iS9G241																	'☆ : 입력/수정용 ComProxy Dll 사용 변수 
	Dim iErrorPosition
	Dim strSpread
	
	On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear	
    
    strSpread = Trim(Request("txtSpread"))
    								
    Set iS9G241 = Server.CreateObject("PS9G241.cSPostBatChargeSvr")  
	
	If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If
    
    Call iS9G241.S_POST_BATCH_CHARGE_SVR(gStrGlobalCollection, strSpread, iErrorPosition)
    
    If CheckSYSTEMError2(Err, True, iErrorPosition & "행","","","","") = True Then
       Set iS9G241 = Nothing
       Exit Sub
	End If
	
	Set iS9G241 = Nothing
	
    Response.Write "<Script Language=vbscript> "	& vbCr         
    Response.Write " Parent.DBSaveOk "				& vbCr   
    Response.Write "</Script> "           
	Response.End																				'☜: Process End
    
End Sub

%>
