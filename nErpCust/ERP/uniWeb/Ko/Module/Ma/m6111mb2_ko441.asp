<%@ LANGUAGE=VBSCript%>
<%Option explicit%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : 
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : iPMAG111(Maint)
'*							  iPMAG118(List)
'*  7. Modified date(First) : 2000/08/28
'*  8. Modified date(Last)  : 2003/06/04
'*  9. Modifier (First)     : Yoon Ji young
'* 10. Modifier (Last)      : Kim Jin Ha
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*							  -2000/03/27 :	
'*							  -2000/04/18 :	
'*							  -2000/05/18 :	표준적용 
'* 14. Meno					: 구매경비등록등록- Biz Logic
'**********************************************************************************************
Dim lgOpModeCRUD
Dim pvCB
 
On Error Resume Next                                                             '☜: Protect system from crashing
Err.Clear                                                                        '☜: Clear Error status

    
Call HideStatusWnd                                                               '☜: Hide Processing message
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*", "NOCOOKIE", "MB")
Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")
    
lgOpModeCRUD  = Request("txtMode") 

    
Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)
             Call SubBizSaveMulti()
        Case "LookUpSupplier"
			 Call SubLookUpSupplier()
		Case "LookupDailyExRt"
			 Call SubLookupDailyExRt()
		Case "LookupVatType"
			 Call SubLookupVatType()
		Case Else
End Select

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	On Error Resume Next
    Err.Clear																		'☜: Protect system from crashing

	Dim iPMAG118																	'☆ : 조회용 ComProxy Dll 사용 변수 

	Dim istrData
	Dim istrDt
	Dim StrNextKey		' 다음 값 
	Dim lgStrPrevKey	' 이전 값 
	Dim iLngMaxRow		' 현재 그리드의 최대Row
	Dim iLngRow
	Dim GroupCount          
	Const C_SHEETMAXROWS_D  = 100
	
	Dim TmpBuffer
	Dim iMax
	Dim iIntLoopCount
	Dim iTotalStr
	
    Dim I1_m_purchase_charge 
    Dim I2_m_purchase_charge_charge_no 
    Dim I3_b_pur_grp_pur_grp 
    Dim EG1_exp_group 
    Dim E1_b_minor 
    Dim E2_b_pur_grp 
    Dim E3_m_purchase_charge_charge_no		'next key
	
	Const M621_I1_bas_no = 0    'View Name : imp m_purchase_charge
	Const M621_I1_process_step = 1
	Const M621_I1_from_charge_dt = 2
	Const M621_I1_to_charge_dt = 3

	Redim I1_m_purchase_charge(M621_I1_to_charge_dt)

	'Group Name : exp_group
	Const M621_EG1_E1_reference = 0    '  View Name : export_item_pay_type_cd b_configuration
	Const M621_EG1_E2_biz_area_nm = 1    '  View Name : export_item b_biz_area
	Const M621_EG1_E3_jnl_nm = 2    '  View Name : export_item a_jnl_item
	Const M621_EG1_E4_bank_nm = 3    '  View Name : export_item b_bank
	Const M621_EG1_E5_minor_nm = 4    '  View Name : export_item_pay_type b_minor
	Const M621_EG1_E6_minor_nm = 5    '  View Name : export_item_vat b_minor
	Const M621_EG1_E7_bp_nm = 6    '  View Name : export_item_payee b_biz_partner
	Const M621_EG1_E8_bp_nm = 7    '  View Name : export_item_build b_biz_partner
	Const M621_EG1_E9_charge_no = 8    '  View Name : export_item m_purchase_charge
	Const M621_EG1_E9_bas_no = 9
	Const M621_EG1_E9_charge_acct = 10
	Const M621_EG1_E9_charge_type = 11
	Const M621_EG1_E9_currency = 12
	Const M621_EG1_E9_charge_doc_amt = 13
	Const M621_EG1_E9_charge_loc_amt = 14
	Const M621_EG1_E9_xch_rate = 15
	Const M621_EG1_E9_charge_dt = 16
	Const M621_EG1_E9_vat_loc_amt = 17
	Const M621_EG1_E9_vat_type = 18
	Const M621_EG1_E9_bp_cd = 19
	Const M621_EG1_E9_bank_cd = 20
	Const M621_EG1_E9_bank_acct = 21
	Const M621_EG1_E9_pay_type = 22
	Const M621_EG1_E9_pay_doc_amt = 23
	Const M621_EG1_E9_pay_loc_amt = 24
	Const M621_EG1_E9_pay_due_dt = 25
	Const M621_EG1_E9_insrt_user_id = 26
	Const M621_EG1_E9_insrt_dt = 27
	Const M621_EG1_E9_updt_user_id = 28
	Const M621_EG1_E9_updt_dt = 29
	Const M621_EG1_E9_cost_flg = 30
	Const M621_EG1_E9_remark = 31
	Const M621_EG1_E9_posting_flg = 32
	Const M621_EG1_E9_process_step = 33
	Const M621_EG1_E9_charge_rate = 34
	Const M621_EG1_E9_distribution_flag = 35
	Const M621_EG1_E9_vat_doc_amt = 36
	Const M621_EG1_E9_vat_rate = 37
	Const M621_EG1_E9_domestic_flag = 38
	Const M621_EG1_E9_bas_doc_no = 39
	Const M621_EG1_E9_biz_area = 40 
	Const M621_EG1_E9_tax_biz_area = 41
	Const M621_EG1_E9_cost_cd = 42
	Const M621_EG1_E9_trans_type = 43
	Const M621_EG1_E9_note_no = 44
	Const M621_EG1_E9_ext1_qty = 45
	Const M621_EG1_E9_ext1_amt = 46
	Const M621_EG1_E9_ext1_rt = 47
	Const M621_EG1_E9_ext1_dt = 48
	Const M621_EG1_E9_ext2_cd = 49
	Const M621_EG1_E9_ext2_qty = 50
	Const M621_EG1_E9_ext2_amt = 51
	Const M621_EG1_E9_ext2_rt = 52
	Const M621_EG1_E9_ext2_dt = 53
	Const M621_EG1_E9_ext3_cd = 54
	Const M621_EG1_E9_ext3_qty = 55
	Const M621_EG1_E9_ext3_amt = 56
	Const M621_EG1_E9_ext3_rt = 57
	Const M621_EG1_E9_ext3_dt = 58
	Const M621_EG1_E9_io_flg = 59
	Const M621_EG1_E9_pre_pay_no = 60
	Const M621_EG1_E9_exp_item_flg = 61
	Const M621_EG1_E9_ext1_cd = 62
	Const M621_EG1_E9_payee_cd = 63
	Const M621_EG1_E9_build_cd = 64
	Const M621_EG1_E9_pp_xch_rt = 65
	Const M621_EG1_E9_xch_rate_op = 66
	Const M621_EG1_E9_comment = 67      '2008-03-25 3:57오후 :: hanc

	Const M621_E1_minor_cd = 0    'View Name : exp_cond_step b_minor
	Const M621_E1_minor_nm = 1

	Const M621_E2_pur_grp = 0    'View Name : exp_cond b_pur_grp
	Const M621_E2_pur_grp_nm = 1

    
	lgStrPrevKey = Request("lgStrPrevKey")

    Set iPMAG118 = Server.CreateObject("PMAG118_ko441.cMLsPurchaseChargeS")     '2008-03-25 3:56오후 :: hanc

    If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
	End If

    I1_m_purchase_charge(M621_I1_process_step)  = Request("txtprocess_step") 
    I1_m_purchase_charge(M621_I1_bas_no)		= Request("txtbas_no") 
    If Trim(Request("txtChargeFrDt"))<>"" Then I1_m_purchase_charge(M621_I1_from_charge_dt)= UNIConvDate(Request("txtChargeFrDt"))
    If Trim(Request("txtChargeToDt"))<>"" Then I1_m_purchase_charge(M621_I1_to_charge_dt)	= UNIConvDate(Request("txtChargeToDt"))
    I3_b_pur_grp_pur_grp 						= Request("txtpur_grp") 
    I2_m_purchase_charge_charge_no				= Request("lgStrPrevKey")		'next key
  
    
    Call iPMAG118.M_LIST_PURCHASE_CHARGE_SVR(gStrGlobalCollection, C_SHEETMAXROWS_D, _
												I1_m_purchase_charge, I2_m_purchase_charge_charge_no, _
												I3_b_pur_grp_pur_grp, EG1_exp_group, E1_b_minor, _
												E2_b_pur_grp, E3_m_purchase_charge_charge_no)	

	
	If CheckSYSTEMError(Err,True) = True Then
		Set iPMAG118 = Nothing
		Response.Write "<Script Language=VBScript>" & vbCr
		Response.Write "With parent" & vbCr
			
		Response.Write ".frm1.txtProcess_step1.Value	= """ & ConvSPChars(UCase(Request("txtprocess_step")))	& """" & vbCr
		Response.Write ".frm1.txtprocess_stepNm1.Value 	= """ & ConvSPChars(E1_b_minor(M621_E1_minor_nm))	& """" & vbCr
		Response.Write ".frm1.txtpur_grp1.Value	    	= """ & ConvSPChars(UCase(Request("txtpur_grp")))		& """" & vbCr
		Response.Write ".frm1.txtpur_grpNm1.Value 		= """ & ConvSPChars(E2_b_pur_grp(M621_E2_pur_grp_nm))		& """" & vbCr
		Response.Write ".frm1.txtbas_no1.Value	    	= """ & ConvSPChars(UCase(Request("txtbas_no")))		& """" & vbCr
		
		Response.Write ".frm1.txtprocess_stepNm.Value 	= """ & ConvSPChars(E1_b_minor(M621_E1_minor_nm))	& """" & vbCr
		Response.Write ".frm1.txtpur_grpNm.Value 		= """ & ConvSPChars(E2_b_pur_grp(M621_E2_pur_grp_nm))		& """" & vbCr
					
		Response.Write ".frm1.hprocecc_step.value		= """ & ConvSPChars(Request("txtprocess_step"))			& """" & vbCr
		Response.Write ".frm1.hbas_no.value		        = """ & ConvSPChars(Request("txtbase_no"))				& """" & vbCr
		Response.Write ".frm1.hpur_grp.value	      	= """ & ConvSPChars(Request("txtpur_grp"))				& """" & vbCr
		Response.write ".frm1.vspdData.MaxRows = 0 " & vbCr
		Response.Write ".dbQueryOk " & vbCr
		Response.Write "End With" & vbCr
		Response.Write "</Script>" & vbCr
		Exit Sub
	End If

	Response.Write "<Script Language=VBScript>" & vbCr
	Response.Write "With parent" & vbCr
	Response.Write "  .frm1.txtprocess_stepNm.Value = """ & ConvSPChars(E1_b_minor(M621_E1_minor_nm))      & """" & vbCr
	Response.Write "  .frm1.txtpur_grpNm.Value 		= """ & ConvSPChars(E2_b_pur_grp(M621_E2_pur_grp_nm))  & """" & vbCr
	Response.Write "End With"				& vbCr
	Response.Write "</Script>"					& vbCr

	If lgStrPrevKey = StrNextKey And UBound(EG1_exp_group,1) < 0 Then
		Set iPMAG118 = Nothing												'☜: ComProxy Unload
		Exit Sub														'☜: 비지니스 로직 처리를 종료함 
	End If

	iLngMaxRow = CInt(Request("txtMaxRows"))
	GroupCount = UBound(EG1_exp_group,1)
    
	If EG1_exp_group(GroupCount, M621_EG1_E9_charge_no) = E3_m_purchase_charge_charge_no Then
		StrNextKey  = ""
	Else
		StrNextKey  = E3_m_purchase_charge_charge_no	'next key
	End If
		
	Const strDefDate = "1899-12-30"
	
	iIntLoopCount = 0
	iMax = UBound(EG1_exp_group,1)
	ReDim TmpBuffer(iMax)
	
	For iLngRow = 0 To iMax
	
		If  iLngRow < C_SHEETMAXROWS_D  Then
		Else
		   StrNextKey = E3_m_purchase_charge_charge_no
           Exit For
        End If  
		
		istrData = ""
		
		If ConvSPChars(EG1_exp_group(iLngRow, M621_EG1_E9_posting_flg)) = "Y" then
        	istrData = istrData & Chr(11) & "1"
        Else
        	istrData = istrData & Chr(11) & "0"
        End if

		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M621_EG1_E9_charge_no))
        istrData = istrData & Chr(11) & ""
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M621_EG1_E9_charge_type))
        istrData = istrData & Chr(11) & ""
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M621_EG1_E3_jnl_nm))
        
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M621_EG1_E9_payee_cd))
        istrData = istrData & Chr(11) & ""
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M621_EG1_E7_bp_nm))
        
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M621_EG1_E9_build_cd)) '계산서발행처 
        istrData = istrData & Chr(11) & ""
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M621_EG1_E8_bp_nm)) '계산서발행처명 
        
        istrData = istrData & Chr(11) & UNIDateClientFormat(EG1_exp_group(iLngRow, M621_EG1_E9_charge_dt))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M621_EG1_E9_vat_type))
        istrData = istrData & Chr(11) & ""
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M621_EG1_E6_minor_nm))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M621_EG1_E9_tax_biz_area))
        istrData = istrData & Chr(11) & ""
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M621_EG1_E2_biz_area_nm))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M621_EG1_E9_ext2_cd))         '2008-03-25 4:05오후 :: hanc
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M621_EG1_E9_ext3_cd))         '2008-03-25 4:05오후 :: hanc
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M621_EG1_E9_ext1_cd))         '2008-03-25 4:05오후 :: hanc

        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M621_EG1_E9_currency))
        istrData = istrData & Chr(11) & ""
        
        istrData = istrData & Chr(11) & UniConvNumDBToCompanyWithOutChange(EG1_exp_group(iLngRow, M621_EG1_E9_charge_doc_amt),0) '금액 
		istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_group(iLngRow, M621_EG1_E9_charge_loc_amt),ggAmtOfMoney.DecPoint,0)					 '자국금액 
        
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_group(iLngRow, M621_EG1_E9_xch_rate),ggExchRate.DecPoint,0)
        istrData = istrData & Chr(11) & UniConvNumDBToCompanyWithOutChange(EG1_exp_group(iLngRow, M621_EG1_E9_vat_rate),0)						 'vat율 
        istrData = istrData & Chr(11) & UniConvNumDBToCompanyWithOutChange(EG1_exp_group(iLngRow, M621_EG1_E9_vat_doc_amt),0)
        istrData = istrData & Chr(11) & UNINumClientFormatByTax(EG1_exp_group(iLngRow, M621_EG1_E9_vat_loc_amt),gCurrency,ggAmtOfMoneyNo)
       
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M621_EG1_E9_pay_type))
        istrData = istrData & Chr(11) & ""
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M621_EG1_E5_minor_nm))
		istrData = istrData & Chr(11) & UniConvNumDBToCompanyWithOutChange(EG1_exp_group(iLngRow, M621_EG1_E9_pay_doc_amt),0)
		'지급자국금액 추가(2003.08.14)
		istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_group(iLngRow, M621_EG1_E9_pay_loc_amt),ggAmtOfMoney.DecPoint,0)					 '자국금액 
		istrDt = UNIDateClientFormat(EG1_exp_group(iLngRow, M621_EG1_E9_pay_due_dt))
		
		If istrDt <> strDefDate Then
			istrData = istrData & Chr(11) & UNIDateClientFormat(EG1_exp_group(iLngRow, M621_EG1_E9_pay_due_dt))
		Else
			istrData = istrData & Chr(11) & ""
		End If
        
        istrData = istrData & Chr(11) & UniConvNumDBToCompanyWithOutChange(EG1_exp_group(iLngRow, M621_EG1_E9_charge_rate),0)
	
		If ConvSPChars(EG1_exp_group(iLngRow, M621_EG1_E9_cost_flg)) = "M" Then
			istrData = istrData & Chr(11) & "1"
		Else
			istrData = istrData & Chr(11) & "0"
		End If
		
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M621_EG1_E9_bank_cd))
        istrData = istrData & Chr(11) & ""
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M621_EG1_E4_bank_nm))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M621_EG1_E9_bank_acct))
        istrData = istrData & Chr(11) & ""
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M621_EG1_E9_note_no))
		istrData = istrData & Chr(11) & ""
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M621_EG1_E9_pre_pay_no))
		istrData = istrData & Chr(11) & ""
		istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_group(iLngRow, M621_EG1_E9_pp_xch_rt),ggExchRate.DecPoint,0)
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M621_EG1_E9_remark)) 
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M621_EG1_E9_bas_no))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M621_EG1_E9_xch_rate_op)) '* or / 가 들어간다 
		istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_group(iLngRow, M621_EG1_E1_reference))
        
        istrData = istrData & Chr(11) & ""
        istrData = istrData & Chr(11) & ""
        
        if ConvSPChars(EG1_exp_group(iLngRow, M621_EG1_E9_posting_flg))= "Y" then
        	istrData = istrData & Chr(11) & "1"
        else
        	istrData = istrData & Chr(11) & "0"
        end if
  
              
        istrData = istrData & Chr(11) & iLngMaxRow + iLngRow
        istrData = istrData & Chr(11) & Chr(12)
		
		TmpBuffer(iIntLoopCount) = istrData
        iIntLoopCount = iIntLoopCount + 1
    Next

	iTotalStr = Join(TmpBuffer, "")    

	Response.Write "<Script Language=VBScript>" & vbCr
	Response.Write "With parent" & vbCr
		
	Response.Write "	.frm1.txtProcess_step1.Value	= """ & ConvSPChars(UCase(Request("txtprocess_step")))	& """" & vbCr
	Response.Write "	.frm1.txtprocess_stepNm1.Value 	= """ & ConvSPChars(E1_b_minor(M621_E1_minor_nm))		& """" & vbCr
	Response.Write "	.frm1.txtpur_grp1.Value	    	= """ & ConvSPChars(UCase(Request("txtpur_grp")))		& """" & vbCr
	Response.Write "	.frm1.txtpur_grpNm1.Value 		= """ & ConvSPChars(E2_b_pur_grp(M621_E2_pur_grp_nm))	& """" & vbCr
	Response.Write "	.frm1.txtbas_no1.Value	    	= """ & ConvSPChars(UCase(Request("txtbas_no")))		& """" & vbCr
	Response.Write "	.frm1.hprocecc_step.value		= """ & ConvSPChars(Request("txtprocess_step"))			& """" & vbCr
	Response.Write "	.frm1.hbas_no.value		        = """ & ConvSPChars(Request("txtbase_no"))				& """" & vbCr
	Response.Write "	.frm1.hpur_grp.value	      	= """ & ConvSPChars(Request("txtpur_grp"))				& """" & vbCr
	Response.Write "	.ggoSpread.Source          =  .frm1.vspdData "         & vbCr
    Response.Write "    .frm1.vspdData.Redraw = False   "                     & vbCr      
    Response.Write "	.ggoSpread.SSShowData        """ & iTotalStr	    & """ , ""F""" & vbCr	
    
    Response.Write  "    Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData," & iLngMaxRow + 1 & "," & iLngMaxRow + iLngRow & ",.C_currency,.C_charge_doc_amt,""A"" ,""I"",""X"",""X"")" & vbCr
    Response.Write  "    Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData," & iLngMaxRow + 1 & "," & iLngMaxRow + iLngRow & ",.C_currency,.C_Vat_rate,""D"" ,""I"",""X"",""X"")" & vbCr
    Response.Write  "    Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData," & iLngMaxRow + 1 & "," & iLngMaxRow + iLngRow & ",.C_currency,.C_vat_doc_amt,""A"" ,""I"",""X"",""X"")" & vbCr
    Response.Write  "    Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData," & iLngMaxRow + 1 & "," & iLngMaxRow + iLngRow & ",.C_currency,.C_pay_doc_amt,""A"" ,""I"",""X"",""X"")" & vbCr
    Response.Write  "    Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData," & iLngMaxRow + 1 & "," & iLngMaxRow + iLngRow & ",.C_currency,.C_charge_rate,""D"" ,""I"",""X"",""X"")" & vbCr   
    
    Response.Write "	.lgStrPrevKey              = """ & StrNextKey   & """" & vbCr 
	Response.Write "	.DbQueryOk "		    	  & vbCr 
    Response.Write "    .frm1.vspdData.Redraw = True   "                      & vbCr   
	Response.Write "End With				" & vbCr
	Response.Write "</Script>					" & vbCr
   
    Set iPMAG118 = Nothing
		
End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

	On Error Resume Next
    Err.Clear																		'☜: Protect system from crashing

	Dim iPMAG111																	'☆ : 입력/수정용 ComProxy Dll 사용 변수 
	Dim iErrorPosition
	
	Dim itxtSpread
	Dim itxtSpreadArr
    Dim itxtSpreadArrCount
    Dim iCUCount
    Dim iDCount
    Dim ii
    
	Dim I1_b_pur_grp_pur_grp 
    Dim I2_b_company 
    Dim I3_m_user_id 
    Dim I4_m_purchase_charge 
	Dim iStrHdnAccount
	Dim iStrSpread
    
	Const M624_I2_loc_cur = 0    'View Name : import b_company
	Const M624_I2_cur_org_change_id = 1
	Redim I2_b_company(M624_I2_cur_org_change_id)

	Const M624_I4_process_step = 0    'View Name : imp_hdr m_purchase_charge
	Const M624_I4_bas_no = 1
	Redim I4_m_purchase_charge(M624_I4_bas_no)
	

	itxtSpread = ""
             
    iCUCount = Request.Form("txtCUSpread").Count
    iDCount  = Request.Form("txtDSpread").Count
    
    itxtSpreadArrCount = -1
             
    ReDim itxtSpreadArr(iCUCount + iDCount)
             
    For ii = 1 To iDCount
        itxtSpreadArrCount = itxtSpreadArrCount + 1
        itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtDSpread")(ii)
    Next
    
    For ii = 1 To iCUCount
        itxtSpreadArrCount = itxtSpreadArrCount + 1
        itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(ii)
    Next
   
    itxtSpread = Join(itxtSpreadArr,"")
	
	I4_m_purchase_charge(M624_I4_process_step)		= UCase(Request("txtProcess_step1")) 
	I4_m_purchase_charge(M624_I4_bas_no)			= UCase(Request("txtbas_no1")) 
	I1_b_pur_grp_pur_grp 				            = Trim(Request("txtpur_grp1")) 
	I2_b_company(M624_I2_loc_cur)				    = gCurrency
	I2_b_company(M624_I2_cur_org_change_id)	        = gChangeOrgId
	I3_m_user_id				     				= gUsrID
	iStrHdnAccount									= Trim(Request("hdninterface_Account"))
	pvCB = "F"
	
	Call RemovedivTextArea()
	
	Set iPMAG111 = Server.CreateObject("PMAG111_KO441.cMMaintPurChargeS")       '2008-03-25 4:47오후 :: hanc

    If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
	End If

	Call iPMAG111.M_MAINT_PURCHASE_CHARGE_SVR(pvCB, gStrGlobalCollection, iStrHdnAccount, I1_b_pur_grp_pur_grp, _
												I2_b_company, I3_m_user_id, I4_m_purchase_charge, itxtSpread, iErrorPosition)

	If CheckSYSTEMError2(Err, True, iErrorPosition & "행","","","","") = True Then
	    Set iPMAG111 = Nothing
	    Exit Sub
	End If

	Response.Write "<Script Language=VBScript>" & vbCr
	Response.Write "With parent" & vbCr
	Response.Write ".frm1.txtProcess_step.Value		= """ & ConvSPChars(UCase(Request("txtprocess_step1"))) & """	" & vbCr
	Response.Write ".frm1.txtbas_no.Value	    	= """ & ConvSPChars(UCase(Request("txtbas_no1"))) & """			" & vbCr
	Response.Write ".frm1.txtpur_grp.Value	    	= """ & ConvSPChars(UCase(Request("txtpur_grp1"))) & """		" & vbCr
	Response.Write ".DbSaveOk		"   & vbCr
	Response.Write "End With"			& vbCr
	Response.Write "</Script>"			& vbCr

	Set iPMAG111 = Nothing

        
End Sub  
'============================================================================================================
' Name : SubLookUpSupplier
' Desc :
'============================================================================================================
Sub SubLookUpSupplier()

	On Error Resume Next
	Err.Clear
	
	Dim iPB5CS41
	Dim iPB5GS45
	Dim iCommandSent
	Dim I1_b_biz_partner
	Dim E1_b_biz_partner
	
	Dim E1_b_biz_partner2		
    Dim E2_b_biz_partner		
    Dim E3_b_biz_partner		
    Dim E4_b_biz_partner		
    Dim E5_b_biz_partner		
    Dim E6_b_biz_partner		
	
	Const S074_E1_bp_cd = 0			
    Const S074_E1_bp_nm = 4 
    Const S074_E1_currency = 17
    Const S074_E1_vat_type = 33
    Const S074_E1_vat_rate = 34
    Const S074_E1_pay_type = 45 
    '인자 순서가 바뀜(2003.08.27) - Lee Eun Hee
    Const S074_E1_vat_type_nm = 132               '[부가세유형명]
    Const S074_E1_pay_type_nm = 134               '[입출금유형명]
    Const S074_E1_pay_type_pur = 116				'입출금유형(구매)
    Const S074_E1_pay_type_pur_nm = 142				'[입출금유형명(구매)]
        
    Const B132_E5_bp_cd = 0    
    Const B132_E5_bp_nm = 1
		
    Set iPB5CS41 = Server.CreateObject("PB5CS41.cLookupBizPartnerSvr")    
    Set iPB5GS45 = Server.CreateObject("PB5GS45.cBListDftBpFtnSvr")
    
    If CheckSYSTEMError(Err,True) = True Then
		Set iPB5CS41 = Nothing
		Set iPB5GS45 = Nothing
		Exit Sub
	End If
	
	I1_b_biz_partner 	= Request("txtBpCd") 
	iCommandSent = "QUERY"
	
	Call iPB5CS41.B_LOOKUP_BIZ_PARTNER_SVR(gStrGlobalCollection, iCommandSent, _
												I1_b_biz_partner, _
												E1_b_biz_partner)
												
	'Change Event시 에러메시지 뿌려주지 않음(2003.08.27)
	'If CheckSYSTEMError(Err,True) = True Then
	'	Set iPB5CS41 = Nothing
	'	Exit Sub
	'End If
	
	
    I1_b_biz_partner = Request("txtBpCd") 


	Call iPB5GS45.B_LIST_DEFAULT_BP_FTN_SVR(gStrGlobalCollection, I1_b_biz_partner, _
											E1_b_biz_partner2, E2_b_biz_partner, _
											E3_b_biz_partner, E4_b_biz_partner, _
											E5_b_biz_partner, E6_b_biz_partner)
	
	Set iPB5CS41 = Nothing
	Set iPB5GS45 = Nothing
	
	If Err.number <> 0 Then
		Response.Write "<Script Language=VBScript>" & vbCr
		'**수정(2003.08.27)
		Response.Write "Dim Row " & vbCr
		Response.Write "With parent.frm1" & vbCr
		Response.Write "Row = .vspdData.Row "						& vbCr	
		Response.Write ".vspdData.Col  = parent.C_bp_cd_Nm										" & vbCr			'지급처명 
		Response.Write ".vspdData.Text = """"" & vbCr
		Response.Write ".vspdData.Col  = parent.C_BuildCd										" & vbCr		    '계산서발행처 
		Response.Write ".vspdData.Text = """"" & vbCr
		Response.Write ".vspdData.Col  = parent.C_Build_Nm										" & vbCr			'계산서발행처명 
		Response.Write ".vspdData.Text = """"" & vbCr
		Response.Write ".vspdData.Col  = parent.C_charge_doc_amt								" & vbCr			
		Response.Write ".vspdData.Text = ""0"" " & vbCr
		Response.Write ".vspdData.Col  = parent.C_charge_loc_amt								" & vbCr			
		Response.Write ".vspdData.Text = ""0"" " & vbCr
		Response.Write ".vspdData.Col  = parent.C_pay_doc_amt									" & vbCr			
		Response.Write ".vspdData.Text = ""0"" " & vbCr
		Response.Write ".vspdData.Col  = parent.C_pay_loc_amt									" & vbCr			
		Response.Write ".vspdData.Text = ""0"" " & vbCr
		Response.Write ".vspdData.Col = parent.C_currency										" & vbCr				 '화폐 
		Response.Write ".vspdData.Text = """"" & vbCr
		Response.Write ".vspdData.Col = parent.C_vat_type										" & vbCr				 'VAT
		Response.Write ".vspdData.Text = """"" & vbCr
		Response.Write ".vspdData.Col = parent.C_vat_type_Nm									" & vbCr				 'VAT명 
		Response.Write ".vspdData.Text = """"" & vbCr
		Response.Write ".vspdData.Col = parent.C_pay_type										" & vbCr				  '지급유형 
		Response.Write ".vspdData.Text = """"" & vbCr
		Response.Write ".vspdData.Col = parent.C_pay_type_Nm									" & vbCr				  '지급유형명 
		Response.Write ".vspdData.Text = """"" & vbCr
		Response.Write ".vspdData.Col = parent.C_Vat_rate										" & vbCr			      'VAT율 
		Response.Write ".vspdData.Text = ""0"" " & vbCr
		Response.Write ".vspdData.Col = parent.C_xch_rate" & vbCr
		Response.Write ".vspdData.Text = """"" & vbCr
		Response.Write ".vspdData.Col = parent.C_calcd" & vbCr
		Response.Write ".vspdData.Text = """"" & vbCr
		Response.Write ".vspdData.Col = parent.C_tax_biz_area" & vbCr
		Response.Write ".vspdData.Text = """"" & vbCr
		Response.Write ".vspdData.Col = parent.C_tax_biz_area_nm" & vbCr
		Response.Write ".vspdData.Text = """"" & vbCr
		Response.Write " Call parent.SetSpreadColor(Row, Row) " & vbCr
		Response.Write " Call parent.vspdData_Change(C_pay_type ,Row) " & vbCr
		
		Response.Write "End With"				& vbCr
		Response.Write "</Script>"				& vbCr
	
		Exit Sub
	End If
	'==============================================================
	
	Response.Write "<Script Language=VBScript>" & vbCr
	'**수정(2003.08.27)
	Response.Write "Dim Row " & vbCr
	Response.Write "With parent.frm1" & vbCr
	Response.Write "Row = .vspdData.Row "						& vbCr	
	Response.Write ".vspdData.Col  = parent.C_bp_cd_Nm										" & vbCr			'지급처명 
	Response.Write ".vspdData.Text = """ & ConvSPChars(E1_b_biz_partner(S074_E1_bp_nm)) & """" & vbCr
		
	Response.Write ".vspdData.Col  = parent.C_BuildCd										" & vbCr		    '계산서발행처 
	Response.Write ".vspdData.Text = """ & ConvSPChars(E5_b_biz_partner(B132_E5_bp_cd)) & """" & vbCr
	Response.Write ".vspdData.Col  = parent.C_Build_Nm										" & vbCr			'계산서발행처명 
	Response.Write ".vspdData.Text = """ & ConvSPChars(E5_b_biz_partner(B132_E5_bp_nm)) & """" & vbCr
		
	Response.Write ".vspdData.Col  = parent.C_charge_doc_amt								" & vbCr			
	Response.Write ".vspdData.Text = ""0"" " & vbCr
	Response.Write ".vspdData.Col  = parent.C_pay_doc_amt									" & vbCr			
	Response.Write ".vspdData.Text = ""0"" " & vbCr
	
	Response.Write ".vspdData.Col = parent.C_currency										" & vbCr				 '화폐 
	Response.Write ".vspdData.Text				= """ & ConvSPChars(E1_b_biz_partner(S074_E1_currency)) & """" & vbCr
	Response.Write ".vspdData.Col = parent.C_vat_type										" & vbCr				 'VAT
	Response.Write ".vspdData.Text				= """ & ConvSPChars(E1_b_biz_partner(S074_E1_vat_type)) & """" & vbCr
	Response.Write ".vspdData.Col = parent.C_vat_type_Nm									" & vbCr				 'VAT명 
	Response.Write ".vspdData.Text				= """ & ConvSPChars(E1_b_biz_partner(S074_E1_vat_type_nm)) & """" & vbCr
	Response.Write ".vspdData.Col = parent.C_pay_type										" & vbCr				  '지급유형 
	Response.Write ".vspdData.Text				= """ & ConvSPChars(E1_b_biz_partner(S074_E1_pay_type_pur)) & """" & vbCr
	Response.Write ".vspdData.Col = parent.C_pay_type_Nm									" & vbCr				  '지급유형명 
	Response.Write ".vspdData.Text				= """ & ConvSPChars(E1_b_biz_partner(S074_E1_pay_type_pur_nm)) & """" & vbCr
	Response.Write ".vspdData.Col = parent.C_Vat_rate										" & vbCr			      'VAT율 
	Response.Write ".vspdData.Text				= """ & UNINumClientFormat(E1_b_biz_partner(S074_E1_vat_rate),ggExchRate.DecPoint,0) & """" & vbCr
	
	Response.Write  "    Call parent.ReFormatSpreadCellByCellByCurrency(.vspdData,Row,Row,parent.C_currency,parent.C_charge_doc_amt,""A"" ,""I"",""X"",""X"")" & vbCr
    Response.Write  "    Call parent.ReFormatSpreadCellByCellByCurrency(.vspdData,Row,Row,parent.C_currency,parent.C_Vat_rate,""D"" ,""I"",""X"",""X"")" & vbCr
    Response.Write  "    Call parent.ReFormatSpreadCellByCellByCurrency(.vspdData,Row,Row,parent.C_currency,parent.C_vat_doc_amt,""A"" ,""I"",""X"",""X"")" & vbCr
    Response.Write  "    Call parent.ReFormatSpreadCellByCellByCurrency(.vspdData,Row,Row,parent.C_currency,parent.C_pay_doc_amt,""A"" ,""I"",""X"",""X"")" & vbCr
    Response.Write  "    Call parent.ReFormatSpreadCellByCellByCurrency(.vspdData,Row,Row,parent.C_currency,parent.C_charge_rate,""D"" ,""I"",""X"",""X"")" & vbCr   
	
   	'GetTaxBizArea함수에 Row 인자 추가함.(2003.08.27)
   	Response.Write "Call parent.GetTaxBizArea(""*"", Row) "			& vbCr
	Response.Write "Call parent.dbquerysupplierok(Row) "				& vbCr
	
	Response.Write "End With"				& vbCr
	Response.Write "</Script>"				& vbCr

End Sub
'============================================================================================================
' Name : SubLookupDailyExRt
' Desc :
'============================================================================================================
Sub SubLookupDailyExRt()

	On Error Resume Next
	Err.Clear                                                               '☜: Protect system from crashing

    Dim iPB0C004
    
    Dim I1_currency
    Dim I2_currency 
    Dim I3_apprl_dt
    
    Dim E1_b_daily_exchange_rate
    Const B253_E1_std_rate = 0
    Const B253_E1_multi_divide = 1
        
    Set iPB0C004 = CreateObject("PB0C004.CB0C004")

    If CheckSYSTEMError(Err,True) = True Then
		Set iPB0C004 = Nothing
		Exit Sub
	End If
    
    I1_currency		= Request("Currency")
    I2_currency		= gCurrency
    I3_apprl_dt		= UNIConvDate(Request("ChargeDt"))    
        

     E1_b_daily_exchange_rate = iPB0C004.B_SELECT_EXCHANGE_RATE(gStrGlobalCollection, I1_currency, _
												I2_currency, I3_apprl_dt)  
     
	'Change Event시 에러메시지 뿌려주지 않음(2003.08.27)
	'If CheckSYSTEMError(Err,True) = True Then
	'	Set iPB0C004 = Nothing
	'	Exit Sub
	'End If
	If Err.number <> 0 Then
		Set iPB0C004 = Nothing
		Response.Write "<Script Language=VBScript>" & vbCr
		'**수정(2003.08.27)
		Response.Write "Dim Row " & vbCr
		Response.Write "With parent.frm1" & vbCr
		Response.Write "Row = .vspdData.Row "						& vbCr	
		Response.Write ".vspdData.Col  = parent.C_charge_loc_amt "	& vbCr			
		Response.Write ".vspdData.Text = ""0"" "					& vbCr
		Response.Write ".vspdData.Col = parent.C_vat_loc_amt"		& vbCr
		Response.Write ".vspdData.Text = ""0"""						& vbCr
		Response.Write ".vspdData.Col  = parent.C_pay_loc_amt "		& vbCr			
		Response.Write ".vspdData.Text = ""0"" "					& vbCr
		Response.Write ".vspdData.Col = parent.C_xch_rate"			& vbCr
		Response.Write ".vspdData.Text = """""						& vbCr
		Response.Write ".vspdData.Col = parent.C_calcd"				& vbCr
		Response.Write ".vspdData.Text = """""						& vbCr
	
		Response.Write "End With"				& vbCr
		Response.Write "</Script>"				& vbCr
		Exit Sub
	End If

	Response.Write "<Script Language=VBScript>" & vbCr
	'**수정(2003.08.27)
	Response.Write "Dim Row " & vbCr
	Response.Write "With parent.frm1.vspdData" & vbCr
	Response.Write "Row = .Row "						& vbCr
	
	Response.Write "If parent.gChangeOpt <> ""XCH"" then " & vbCr
	Response.Write ".Col = parent.C_xch_rate" & vbCr
	Response.Write ".Text = """ & UNINumClientFormat(E1_b_daily_exchange_rate(B253_E1_std_rate),ggExchRate.DecPoint,0) & """" & vbCr
	Response.Write "End If" & vbCr
	Response.Write ".Col = parent.C_calcd" & vbCr
	Response.Write ".Text = """ & ConvSPChars(E1_b_daily_exchange_rate(B253_E1_multi_divide)) & """" & vbCr
	Response.Write "End With"				& vbCr

	Response.Write "Call parent.ChangeCurOrDtOk(Row)"			& vbCr
	Response.Write "</Script>"					& vbCr

	        
    Set iPB0C004 = Nothing                                                   '☜: Unload Comproxy

End Sub																		'☜: Process End

'============================================================================================================
' Name : SubLookupVatType
' Desc : 
'============================================================================================================
Sub SubLookupVatType()									  'vat타입 변동시 변동값 setting(타입명과vatsetting)

    On Error Resume Next
    Err.Clear                                                               '☜: Protect system from crashing

    Dim iPB0C003
    
    Dim I1_b_major_major_cd
	Dim I2_b_minor_minor_cd
	Dim I3_b_configuration_seq_no
	Dim E1_b_minor
	Dim E2_b_configuration
	
	Const B249_E1_minor_cd = 0
    Const B249_E1_minor_nm = 1

    Const B249_E2_seq_no = 0
    Const B249_E2_reference = 1
    Const B249_E2_ref_type = 2
    
    Set iPB0C003 = CreateObject("PB0C003.CB0C003")

    If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
	End If

    I3_b_configuration_seq_no	= "1"
	I1_b_major_major_cd			= "b9001"
    I2_b_minor_minor_cd			= Request("VatType")
	

    Call iPB0C003.B_SELECT_CONFIGURATION(gStrGlobalCollection, I1_b_major_major_cd, _
										I2_b_minor_minor_cd, I3_b_configuration_seq_no, _
										E1_b_minor, E2_b_configuration)  
   
	
	'Change Event시 에러메시지 뿌려주지 않음(2003.08.27)
	'If CheckSYSTEMError(Err,True) = True Then
	'	Set iPB0C003 = Nothing
	'	Exit Sub
	'End If
	If Err.number <> 0 Then
		Set iPB0C003 = Nothing
		Response.Write "<Script Language=VBScript>" & vbCr
		Response.Write "With parent.frm1" & vbCr
		Response.Write ".vspdData.Col = parent.C_vat_type_Nm"		& vbCr
		Response.Write ".vspdData.Text = """""						& vbCr
		Response.Write ".vspdData.Col = parent.C_Vat_rate"			& vbCr
		Response.Write ".vspdData.Text = ""0"""						& vbCr
		Response.Write ".vspdData.Col = parent.C_vat_doc_amt"		& vbCr
		Response.Write ".vspdData.Text = ""0"""						& vbCr
		Response.Write ".vspdData.Col = parent.C_vat_loc_amt"		& vbCr
		Response.Write ".vspdData.Text = ""0"""						& vbCr
	
		Response.Write "End With"				& vbCr
		Response.Write "</Script>"				& vbCr
		Exit Sub
	End If

	'-----------------------
	'Result data display area
	'----------------------- 
	Response.Write "<Script Language=VBScript>" & vbCr
	'**수정(2003.08.27)
	Response.Write "Dim Row "					& vbCr
	Response.Write "With parent.frm1.vspdData"	& vbCr
	Response.Write "Row = .Row "				& vbCr

	Response.Write ".Col = parent.C_vat_type_Nm " & vbCr
	Response.Write ".Text = """ & ConvSPChars(E1_b_minor(B249_E1_minor_nm)) & """" & vbCr
	Response.Write ".Col = parent.C_Vat_rate" & vbCr
	Response.Write ".Text = """ & E2_b_configuration(B249_E2_reference) & """" & vbCr
	Response.Write "End With"				& vbCr
	'**부가세금액수정(2003.08.14)
	Response.Write "Call parent.ChangeVatAmt(Row) " & vbCr
	Response.Write "</Script>"					& vbCr

    Set iPB0C003 = Nothing                                                   '☜: Unload Comproxy

End Sub
'============================================================================================================
' Name : RemovedivTextArea
' Desc : 
'============================================================================================================
Sub RemovedivTextArea()
    On Error Resume Next                                                             
    Err.Clear                                                                        
	
	Response.Write "<Script language=vbs> " & vbCr   
    Response.Write "Parent.RemovedivTextArea "      & vbCr   
    Response.Write "</Script> "      & vbCr   
End Sub
%>
