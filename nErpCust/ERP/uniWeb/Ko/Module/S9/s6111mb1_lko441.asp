<%
'********************************************************************************************************
'*  1. Module Name          : Sales																		*
'*  2. Function Name        : 판매경비관리																*
'*  3. Program ID           : S6111MB1																	*
'*  4. Program Name         : 판매경비등록																*
'*  5. Program Desc         :																			*
'*  6. Comproxy List        : PS9G111.dll, PS9G118.dll, PB0C003.dll, PB0C004.dll
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

Call loadInfTB19029B("I", "*","NOCOOKIE","MB") 
Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")
Call LoadBasisGlobalInf()

Call HideStatusWnd                                                                 '☜: Hide Processing message
'---------------------------------------Common-----------------------------------------------------------
	
lgOpModeCRUD = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete

        Case CStr("ConXchRate")                                                      '☜: 환율 요청을 받음 
			Call SubConXchRate()
        Case CStr("ConVatType")														 '☜: 계산서종류 요청을 받음 
            Call SubConVatType
    End Select


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    
	Dim iLngRow	
	Dim iLngMaxRow
	Dim istrData
	Dim iStrPrevKey
	
    Dim iS61118    
    '-----------------------------------------------
    ' Declare User Variable
    '-----------------------------------------------
    ' 조회 조건 
    ' 진행구분 / 영업그룹 / 발생근거번호 
    Dim i_b_sales_grp1 
    Dim i_s_sales_charge1 
    ReDim i_s_sales_charge1(1)
    
    ' Next Page Variable
    Dim imp_next_s_sales_charge1   

    ' Reruen Call Variable
    Dim i_b_sales_grp
    Dim i_s_sales_charge 
    Dim imp_next_s_sales_charge

    ' Export Variables
    Dim exp_b_minor 
    Dim exp_b_sales_grp 
    Dim exp_s_sales_charge
    Dim exp_s_sales_charge_bas_no
    Dim exp_grp
    
    Dim intGroupCount
    Dim StrNextKey  	
    Dim arrValue
    
    Const C_SHEETMAXROWS_D  = 100
    
    ' exp_grp 저장 
    Const c_exp_E1_biz_area_nm = 0    '[CONVERSION INFORMATION]  View Name : exp_item b_biz_area
    Const c_exp_E2_jnl_nm = 1    '[CONVERSION INFORMATION]  View Name : exp_item a_jnl_item
    Const c_exp_E3_minor_nm = 2    '[CONVERSION INFORMATION]  View Name : exp_item_vat_type_nm b_minor
    Const c_exp_E4_minor_nm = 3    '[CONVERSION INFORMATION]  View Name : exp_item_pay_type_nm b_minor
    Const c_exp_E5_bank_acct_no = 4    '[CONVERSION INFORMATION]  View Name : exp_item b_bank_acct
    Const c_exp_E6_bank_cd = 5    '[CONVERSION INFORMATION]  View Name : exp_item b_bank
    Const c_exp_E6_bank_nm = 6
    Const c_exp_E7_bp_nm = 7    '[CONVERSION INFORMATION]  View Name : exp_item b_biz_partner
    Const c_exp_E7_bp_cd = 8
    Const c_exp_E8_charge_no = 9    '[CONVERSION INFORMATION]  View Name : exp_item s_sales_charge
    Const c_exp_E8_bas_doc_no = 10
    Const c_exp_E8_charge_cd = 11
    Const c_exp_E8_cur = 12
    Const c_exp_E8_charge_doc_amt = 13
    Const c_exp_E8_charge_loc_amt = 14
    Const c_exp_E8_xch_rate = 15
    Const c_exp_E8_charge_dt = 16
    Const c_exp_E8_vat_loc_amt = 17
    Const c_exp_E8_vat_type = 18
    Const c_exp_E8_pay_type = 19
    Const c_exp_E8_remark = 20
    Const c_exp_E8_posting_flag = 21
    Const c_exp_E8_process_step = 22
    Const c_exp_E8_bas_no = 23
    Const c_exp_E8_vat_rate = 24
    Const c_exp_E8_cost_flag = 25
    Const c_exp_E8_note_no = 26
    Const c_exp_E8_vat_amt = 27
    Const c_exp_E8_tax_biz_area = 28
    Const c_exp_E8_pay_doc_amt = 29
    Const c_exp_E8_pay_loc_amt = 30
    Const c_exp_E8_pay_due_dt = 31
    Const c_exp_E8_xch_rate_op = 32
    Const c_exp_E8_ext1_qty = 33
    Const c_exp_E8_ext2_qty = 34
    Const c_exp_E8_ext3_qty = 35
    Const c_exp_E8_ext1_amt = 36
    Const c_exp_E8_ext2_amt = 37
    Const c_exp_E8_ext3_amt = 38
    Const c_exp_E8_ext1_cd = 39
    Const c_exp_E8_ext2_cd = 40
    Const c_exp_E8_ext3_cd = 41

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                      '☜: Clear Error status
    '-----------------------------------------------
    ' 진행구분 
    '-----------------------------------------------
    i_s_sales_charge1(0) = Trim(Request("txtProcessStepCd"))
    '-----------------------------------------------
    ' 영업그룹 
    '-----------------------------------------------
    i_b_sales_grp1 = Trim(Request("txtSalesGrp"))
    '-----------------------------------------------
    ' 발생근거번호 
    '-----------------------------------------------
    i_s_sales_charge1(1) = Trim(Request("txtBasNo"))
   
	iStrPrevKey = Trim(Request("lgStrPrevKey"))                                        '☜: Next Key	
	
	If iStrPrevKey <> "" then
		arrValue = Split(iStrPrevKey, gColSep)
        imp_next_s_sales_charge1 = arrValue(0)
	else			
        imp_next_s_sales_charge1 = ""
	End If

    i_b_sales_grp = i_b_sales_grp1
    i_s_sales_charge = i_s_sales_charge1
    imp_next_s_sales_charge = imp_next_s_sales_charge1
    
	Set iS61118 = Server.CreateObject("PS9G118.cSLtSalesChargeSvr")

	If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If

	' Call the Dll   		
    Call iS61118.S_LIST_SALES_CHARGE_SVR ( gStrGlobalCollection, _
                           C_SHEETMAXROWS_D, _
                           CSTR(imp_next_s_sales_charge), i_s_sales_charge, CSTR(i_b_sales_grp), _
                           exp_b_minor, exp_b_sales_grp, exp_s_sales_charge, _
                           exp_grp)

	If CheckSYSTEMError(Err,TRUE) = True Then
       Set iS61118 = Nothing		                                                 '☜: Unload Comproxy DLL
		Response.Write "<Script language=vbs>  " & vbCr   
		Response.Write " Parent.frm1.txtConProcessStepCd.focus  " & vbCr   		
		Response.Write "</Script>      " & vbCr      
       Exit Sub
    End If   
    
    Set iS61118 = Nothing	
        
    iLngMaxRow  = CLng(Request("txtMaxRows"))										'☜: Fetechd Count      

    ' test Mode
        
    For iLngRow = 0 To UBound(exp_grp,1)
    		
    		
        If  iLngRow < C_SHEETMAXROWS_D  Then
        Else

            StrNextKey = ConvSPChars(exp_grp(iLngRow, c_exp_E8_charge_no))
            
            Exit For
        End If 
        
        exp_s_sales_charge_bas_no = ConvSPChars(exp_grp(iLngRow, c_exp_E8_bas_doc_no))
        
		'--확정여부 
		If ConvSPChars(exp_grp(iLngRow, c_exp_E8_posting_flag)) = "Y" Then
			istrdata = istrdata & Chr(11) & "1"
		Else
			istrdata = istrdata & Chr(11) & "0"
		End If        
        
		'--경비관리번호 
		istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, c_exp_E8_charge_no))					
		 '--경비항목 
		istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, c_exp_E8_charge_cd))
		istrdata = istrdata & Chr(11) & "" 
		 '--경비항목명 
		istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, c_exp_E2_jnl_nm))
		 '--거래처 
		istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, c_exp_E7_bp_cd))
		istrdata = istrdata & Chr(11) & "" 
		 '--거래처명 
		istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, c_exp_E7_bp_nm))
		 '--'2008-04-21 6:25오후 :: hanc거래처 
		istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, c_exp_E8_ext1_cd))
		istrdata = istrdata & Chr(11) & "" 
		 '--'2008-04-21 6:25오후 :: hanc거래처명 
		istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, c_exp_E8_ext2_cd))
		 '--발생일 
		istrdata = istrdata & Chr(11) & UNIDateClientFormat(exp_grp(iLngRow,c_exp_E8_charge_dt))
		 '--계산서종류 
		istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, c_exp_E8_vat_type))
		istrdata = istrdata & Chr(11) & "" 
		 '--계산서종류명 
		istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, c_exp_E3_minor_nm))
		 '--화페단위 
		istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, c_exp_E8_cur))
		istrdata = istrdata & Chr(11) & "" 

		 '--발생금액 
		istrdata = istrdata & Chr(11) & UniConvNumDBToCompanyWithOutChange(exp_grp(iLngRow, c_exp_E8_charge_doc_amt), 0)

		 '--환율연산자 
		istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, c_exp_E8_xch_rate_op))

		 '--환율 
		istrdata = istrdata & Chr(11) & UNINumClientFormat(exp_grp(iLngRow, c_exp_E8_xch_rate), ggExchRate.DecPoint, 0)
		 '--발생자국금액 
		istrdata = istrdata & Chr(11) & UniConvNumDBToCompanyWithOutChange(exp_grp(iLngRow, c_exp_E8_charge_loc_amt), 0)
		 '--부가세율 
		istrdata = istrdata & Chr(11) & UNINumClientFormat(exp_grp(iLngRow, c_exp_E8_vat_rate), ggExchRate.DecPoint, 0)
		 '--부가세발생금액 
		istrdata = istrdata & Chr(11) & UniConvNumDBToCompanyWithOutChange(exp_grp(iLngRow, c_exp_E8_vat_amt), 0)
		 '--부가세자국금액 
		istrdata = istrdata & Chr(11) & UniConvNumDBToCompanyWithOutChange(exp_grp(iLngRow, c_exp_E8_vat_loc_amt), 0)

		 '--세금신고사업장 
		istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, c_exp_E8_tax_biz_area))
		istrdata = istrdata & Chr(11) & "" 
		 '--세금신고사업장명 
		istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, c_exp_E1_biz_area_nm))

		 '--지급만기일 
		istrdata = istrdata & Chr(11) & UNIDateClientFormat(exp_grp(iLngRow, c_exp_E8_pay_due_dt))
		 '--지급유형 
		istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, c_exp_E8_pay_type))
		istrdata = istrdata & Chr(11) & "" 
		 '--지급유형명 
		istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, c_exp_E4_minor_nm))

		 '--지급금액 
		istrdata = istrdata & Chr(11) & UniConvNumDBToCompanyWithOutChange(exp_grp(iLngRow, c_exp_E8_pay_doc_amt), 0)
		 '--지급자국금액 
		istrdata = istrdata & Chr(11) & UniConvNumDBToCompanyWithOutChange(exp_grp(iLngRow, c_exp_E8_pay_loc_amt), 0)

		 '--어음번호 
		istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, c_exp_E8_note_no))
		istrdata = istrdata & Chr(11) & "" 
		 '--출금은행코드 
		istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, c_exp_E6_bank_cd))
		istrdata = istrdata & Chr(11) & "" 
		 '--출금은행 
		istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, c_exp_E6_bank_nm))
		 '--출금계좌 
		istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, c_exp_E5_bank_acct_no))
		istrdata = istrdata & Chr(11) & "" 

		 '--기타참조사항 
		istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, c_exp_E8_remark))

        istrdata = istrData & Chr(11) & iLngMaxRow + iLngRow
        istrData = istrData & Chr(11) & Chr(12)       
            
    Next 
    
        Response.Write "<Script language=vbs> " & vbCr 

        Response.Write " Parent.frm1.txtConProcessStepNm.value	= """ & ConvSPChars(exp_b_minor) & """" & vbCr    
        Response.Write " Parent.frm1.txtConSalesGrpNm.value		= """ & ConvSPChars(exp_b_sales_grp(1)) & """" & vbCr    
        'Response.Write " Parent.frm1.txtConBasNo.value			= """ & ConvSPChars(exp_s_sales_charge(0)) & """" & vbCr   
        Response.Write " Parent.HdrQueryOk " & vbCr   
        
        Response.Write " Parent.frm1.txtProcessStepCd.value		= """ & ConvSPChars(Trim(Request("txtProcessStepCd"))) & """" & vbCr
        Response.Write " Parent.frm1.txtProcessStepNm.value		= """ & ConvSPChars(exp_b_minor) & """" & vbCr
        Response.Write " Parent.frm1.txtSalesGrp.value			= """ & ConvSPChars(Trim(Request("txtSalesGrp")))  & """" & vbCr
        Response.Write " Parent.frm1.txtSalesGrpNm.value		= """ & ConvSPChars(exp_b_sales_grp(1)) & """" & vbCr
        Response.Write " Parent.frm1.txtBasNo.value				= """ & ConvSPChars(Trim(Request("txtBasNo"))) & """" & vbCr
        Response.Write " Parent.frm1.txtBasDocNo.value			= """ & ConvSPChars(exp_s_sales_charge_bas_no) & """" & vbCr

        Response.Write " Parent.ggoSpread.Source				= Parent.frm1.vspdData " & vbCr

		Response.Write " Parent.frm1.vspdData.Redraw = False   "                     & vbCr      
		Response.Write " Parent.ggoSpread.SSShowDataByClip   """ & istrData & """ ,""F""" & vbCr
		Response.Write " Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData," & iLngMaxRow + 1 & "," & iLngMaxRow + UBound(exp_grp,1) + 1  & ",Parent.C_Curr,Parent.C_ChargeDocAmt,""A"" ,""I"",""X"",""X"")" & vbCr
		Response.Write " Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.Frm1.vspdData," & iLngMaxRow + 1 & "," & iLngMaxRow + UBound(exp_grp,1) + 1  & ",Parent.Parent.gCurrency,Parent.C_ChargeLocAmt,""A"" ,""I"",""X"",""X"")" & vbCr
		Response.Write " Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData," & iLngMaxRow + 1 & "," & iLngMaxRow + UBound(exp_grp,1) + 1  & ",Parent.C_Curr,Parent.C_VatDocAmt,""A"" ,""I"",""X"",""X"")" & vbCr
		Response.Write " Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.Frm1.vspdData," & iLngMaxRow + 1 & "," & iLngMaxRow + UBound(exp_grp,1) + 1  & ",Parent.Parent.gCurrency,Parent.C_VatLocAmt,""A"" ,""I"",""X"",""X"")" & vbCr
		Response.Write " Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData," & iLngMaxRow + 1 & "," & iLngMaxRow + UBound(exp_grp,1) + 1  & ",Parent.C_Curr,Parent.C_PayDocAmt,""A"" ,""I"",""X"",""X"")" & vbCr		
		Response.Write " Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.Frm1.vspdData," & iLngMaxRow + 1 & "," & iLngMaxRow + UBound(exp_grp,1) + 1  & ",Parent.Parent.gCurrency,Parent.C_PayLocAmt,""A"" ,""I"",""X"",""X"")" & vbCr										

        Response.Write " Parent.lgStrPrevKey					= """ & StrNextKey & """" & vbCr

        Response.Write " Parent.frm1.txtHConProcessStepCd.value = """ & ConvSPChars(Trim(Request("txtProcessStepCd"))) & """" & vbCr
        Response.Write " Parent.frm1.txtHConBasNo.value			= """ & ConvSPChars(Trim(Request("txtBasNo"))) & """" & vbCr
        Response.Write " Parent.frm1.txtHConSalesGrp.value		= """ & ConvSPChars(Trim(Request("txtSalesGrp"))) & """" & vbCr

        Response.Write " Parent.DbQueryOK " & vbCr   
		Response.Write " Parent.frm1.vspdData.Redraw = True   "                      & vbCr              
        Response.Write "</Script> "		
                  	
End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()   

	Dim iS61111	
	Dim iErrorPosition
		
	Dim I1_s_sales_charge1
	ReDim I1_s_sales_charge1(1)

	Dim I1_s_sales_charge
	Dim I2_b_sales_grp
	Dim strSpread
	Dim pvCB
		
	On Error Resume Next                                                                 '☜: Protect system from crashing
	Err.Clear																			 '☜: Clear Error status                                                            
    
    pvCB = "F"
    I1_s_sales_charge1(0) = UCase(Trim(Request("txtProcessStepCd")))
    I1_s_sales_charge1(1) = UCase(Trim(Request("txtBasNo")))

    I2_b_sales_grp = UCase(Trim(Request("txtSalesGrp")))
    
    I1_s_sales_charge = I1_s_sales_charge1

	strSpread = Trim(Request("txtSpread"))
	
	Set iS61111 = Server.CreateObject("PS9G111.cSSalesChargeSvr")
	
    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If
    
    Call iS61111.S_MAINT_SALES_CHARGE_SVR  (pvCB, gStrGlobalCollection, _
												strSpread, _
												I1_s_sales_charge, cstr(I2_b_sales_grp), iErrorPosition)    
												      
    If CheckSYSTEMError2(Err, True, iErrorPosition & "행","","","","") = True Then
       Set iS61111 = Nothing
       Exit Sub
	End If
	

	
    Set iS61111 = Nothing
                                                       
    Response.Write "<Script language=vbs> " & vbCr 
    Response.Write " Parent.frm1.txtConProcessStepCd.value	= """ & ConvSPChars(Trim(Request("txtProcessStepCd"))) & """" & vbCr   
    Response.Write " Parent.frm1.txtConBasNo.value			= """ & ConvSPChars(Trim(Request("txtBasNo"))) & """" & vbCr   
    Response.Write " Parent.frm1.txtConSalesGrp.value		= """ & ConvSPChars(Trim(Request("txtSalesGrp"))) & """" & vbCr           
    Response.Write " Parent.DBSaveOk "      & vbCr   
    Response.Write "</Script> "             & vbCr   
          
End Sub 

'============================================================================================================
' Name : SubConXchRate
' Desc : 환율 요청을 받음 
'============================================================================================================
Sub SubConXchRate()

Dim iB17013	
Dim iErrorPosition

Dim arrXchVal, arrXchTemp								'☜: Spread Sheet 의 값을 받을 Array 변수 
	
On Error Resume Next                                                                 '☜: Protect system from crashing
Err.Clear																			 '☜: Clear Error status                                                            

    '-----------------------
    'Data manipulate area
    '-----------------------
    arrXchTemp = Split(Request("txtSpread"), gRowSep)    
    arrXchVal = Split(arrXchTemp(0), gColSep)

    Dim I1_currency 
    Dim I2_currency
    Dim I3_b_daily_exchange_rate
    Dim I4_num_value_15_2
    Dim I5_data_type
	 	    
    Dim I3_b_daily_exchange_rate1
    Redim I3_b_daily_exchange_rate1(1)
	
    Dim E1_b_daily_exchange_rate
    Redim E1_b_daily_exchange_rate(1)
    Dim E2_exchange_variable_num_value_15_2
    

    I1_currency = Trim(arrXchVal(1))
    I2_currency	= Trim(arrXchVal(2))
    I3_b_daily_exchange_rate1(0) = UNIConvDate(arrXchVal(0))
    I3_b_daily_exchange_rate1(1) =  Trim(arrXchVal(4))
    I4_num_value_15_2 = UNIConvNum(arrXchVal(3),0)
    I5_data_type = Trim("2")
        
    I3_b_daily_exchange_rate = I3_b_daily_exchange_rate1
        
    Set iB17013 = Server.CreateObject("PB0C004.CB0C004")
	
    If CheckSYSTEMError(Err,True) = True Then
        Exit Sub
    End If
    
    Call iB17013.B_TRANS_EXCH_RATE_BY_USER (gStrGlobalCollection, _
												cstr(I1_currency), cstr(I2_currency), I3_b_daily_exchange_rate, _
												CSTR(I4_num_value_15_2), CSTR(I5_data_type), _
												E1_b_daily_exchange_rate, E2_exchange_variable_num_value_15_2)
	
    If CheckSYSTEMError(Err,TRUE) = True Then
        Set iB17013 = Nothing		                                                 '☜: Unload Comproxy DLL
        Exit Sub
    End If   

    Set iB17013 = Nothing
%>
<Script Language=vbscript>
	With parent																			
		'--원화금액 
		.frm1.txtSpread.value = "<%=UNINumClientFormat(E2_exchange_variable_num_value_15_2, ggAmtOfMoney.DecPoint, 0)%>"
	End With
</Script>
<%

    Call SubConDataType()
    
End Sub

Sub SubConDataType()

'====================================================
'=		화폐에 따른 환율연산자						=
'====================================================

Dim iB17014
Dim iErrorPosition

Dim arrXchVal, arrXchTemp								'☜: Spread Sheet 의 값을 받을 Array 변수 
	
On Error Resume Next                                                                 '☜: Protect system from crashing
Err.Clear																			 '☜: Clear Error status                                                            

    '-----------------------
    'Data manipulate area
    '-----------------------
    arrXchTemp = Split(Request("txtSpread"), gRowSep)    
    arrXchVal = Split(arrXchTemp(0), gColSep)

    '-----------------------
    'Data manipulate area
    '-----------------------
    Dim I1_currency
    Dim I2_currency
    Dim I3_apprl_dt
    
    Dim Exp_Result
    Redim Exp_Result(1)
    
    I1_currency = Trim(arrXchVal(1))
    I2_currency = Trim(arrXchVal(2))
    I3_apprl_dt = UNIConvDate(arrXchVal(0))
    
    Set iB17014 = Server.CreateObject("PB0C004.CB0C004")
	
    If CheckSYSTEMError(Err,True) = True Then
        Exit Sub
    End If
    
    Exp_Result(0) = iB17014.B_SELECT_EXCHANGE_RATE (gStrGlobalCollection, _
								cstr(I1_currency), cstr(I2_currency), I3_apprl_dt )(0)
								
    Exp_Result(1) = iB17014.B_SELECT_EXCHANGE_RATE (gStrGlobalCollection, _
								cstr(I1_currency), cstr(I2_currency), I3_apprl_dt )(1)
								

    If CheckSYSTEMError(Err,TRUE) = True Then
        Set iB17014 = Nothing		                                                 '☜: Unload Comproxy DLL
        Exit Sub
    End If   
	
    Set iB17014 = Nothing
    
%>
<Script Language=vbscript>
	With parent
		.frm1.vspdData.Row = <%=arrXchVal(5)%>
		.frm1.vspdData.Col = .C_XchRate
		.frm1.vspdData.Text = "<%=UNINumClientFormat(Exp_Result(0), ggExchRate.DecPoint, 0)%>"
		.frm1.vspdData.Col = .C_XchCalop
		.frm1.vspdData.Text = "<%=Exp_Result(1)%>"
		.ChangeXchRateOk(<%=arrXchVal(5)%>)		
	End With
</Script>
<%

End Sub
'============================================================================================================
' Name : SubConVatType
' Desc : 계산서종류 요청을 받음 
'============================================================================================================
Sub SubConVatType

Dim iB1A059	
Dim iErrorPosition

Dim arrVatVal, arrVatTemp																'☜: Spread Sheet 의 값을 받을 Array 변수 
	
On Error Resume Next                                                                 '☜: Protect system from crashing
Err.Clear																			 '☜: Clear Error status                                                            
    
    '-----------------------
    'Data manipulate area
    '-----------------------
    
    arrVatTemp = Split(Request("txtSpread"), gRowSep)
    
	arrVatVal = Split(arrVatTemp(0), gColSep)

	Dim I1_Major_cd 
	Dim I2_Minor_cd
	Dim I3_Seq_No 	
	
	Dim E1_b_minor
	Dim E2_b_configuration

    I1_Major_cd = Trim(arrVatVal(1))
	I2_Minor_cd	= Trim(arrVatVal(2))
	I3_Seq_No = Cint(arrVatVal(0))

	Set iB1A059 = Server.CreateObject("PB0C003.CB0C003")
	
    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If
    
    Call iB1A059.B_SELECT_CONFIGURATION (gStrGlobalCollection, _
												cstr(I1_Major_cd), cstr(I2_Minor_cd), cint(I3_Seq_No), _
												E1_b_minor, E2_b_configuration)
												      
	If CheckSYSTEMError(Err,TRUE) = True Then
       Set iB1A059 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If   
	
    Set iB1A059 = Nothing
   
%>
<Script Language=vbscript>
	With parent																			
		'--부가세율 
		.frm1.vspdData.Row = <%=arrVatVal(3)%>
		.frm1.vspdData.Col = 12
		.frm1.vspdData.Text = "<%=E1_b_minor(1)%>"
		.frm1.txtSpread.value = "<%=UNINumClientFormat(E2_b_configuration(1), ggExchRate.DecPoint, 0)%>"
		.ChangeVatTypeOk(<%=arrVatVal(3)%>)
	End With
</Script>
<%                                                      

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
