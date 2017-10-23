<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : m3111mb4
'*  4. Program Name         : �ܰ�Ȯ�� 
'*  5. Program Desc         : �ܰ�Ȯ�� 
'*  6. Component List       : PM3G1F8.cMListPoDtlFixPrcS / PM3G1FP.cMFixPurOrdPriceS / PM3G1P9.cMLookupPriceS
'*  7. Modified date(First) : 2000/05/11
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Shin Jin Hyun
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<%	
call LoadBasisGlobalInf()
call LoadInfTB19029B("I", "*","NOCOOKIE","MB") 
call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")

    Dim lgOpModeCRUD
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    Call HideStatusWnd                                                               '��: Hide Processing message

    lgOpModeCRUD  = Request("txtMode") 
										                                              '��: Read Operation Mode (CRUD)
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '��: Query
             Call  SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '��: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '��: Delete
		Case "lookupPrice"
			 Call SublookupPrice()
		Case "lookupPriceForSelection"			
			 Call lookupPriceForSelection()
    End Select
    
    

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    On Error Resume Next                                                             '��: Protect system from crashing	
	Dim iPM3G1F8																	'�� : ��ȸ�� ComProxy Dll ��� ���� 
	Dim iMax
	Dim PvArr

	Dim StrNextKey		' ���� �� 
	Dim lgStrPrevKey	' ���� �� 
	Dim iLngMaxRow		' ���� �׸����� �ִ�Row
	Dim iLngRow
	Dim GroupCount   
	Dim arrValue       
	Dim istrData
	Dim lgCurrency
	
	Const C_SHEETMAXROWS_D  = 100

    Dim I1_m_po_dtl_po_seq
    Dim l2_m_po_hdr_po_no
    
    Dim I3_to_po_dt
    Dim I4_m_pur_ord_hdr
    Const M098_I4_po_no = 0
    Const M098_I4_po_dt = 1
    ReDim  I4_m_pur_ord_hdr(M098_I4_po_dt)

    Dim I5_b_biz_partner
    Dim I6_b_pur_grp
    Dim l7_b_cfm_flg
    Dim I8_b_item_cd
    
    Dim E1_m_po_dtl_next 
	Const M098_E1_po_no = 0

	Dim E2_m_pur_ord_dtl
	Const M098_E2_po_seq_no = 0
   
    Dim EG1_exp_grp
    Const M098_EG1_E1_reference = 0   
    Const M098_EG1_E2_minor_nm = 1    
    Const M098_EG1_E3_po_no = 2    
    Const M098_EG1_E3_po_dt = 3
    Const M098_EG1_E3_po_cur = 4
    Const M098_EG1_E4_plant_cd = 5 
    Const M098_EG1_E4_plant_nm = 6
    Const M098_EG1_E5_item_cd = 7  
    Const M098_EG1_E5_item_nm = 8
    Const M098_EG1_E6_po_seq_no = 9
    Const M098_EG1_E6_dlvy_dt = 10
    Const M098_EG1_E6_po_qty = 11
    Const M098_EG1_E6_po_unit = 12
    Const M098_EG1_E6_po_base_qty = 13
    Const M098_EG1_E6_po_base_unit = 14
    Const M098_EG1_E6_fr_trans_coef = 15
    Const M098_EG1_E6_to_trans_coef = 16
    Const M098_EG1_E6_po_prc = 17
    Const M098_EG1_E6_po_prc_flg = 18
    Const M098_EG1_E6_po_doc_amt = 19
    Const M098_EG1_E6_po_loc_amt = 20
    Const M098_EG1_E6_rcpt_qty = 21
    Const M098_EG1_E6_iv_qty = 22
    Const M098_EG1_E6_lc_qty = 23
    Const M098_EG1_E6_bl_qty = 24
    Const M098_EG1_E6_cc_qty = 25
    Const M098_EG1_E6_po_sts = 26
    Const M098_EG1_E6_cls_flg = 27
    Const M098_EG1_E6_tracking_no = 28
    Const M098_EG1_E6_so_no = 29
    Const M098_EG1_E6_so_seq_no = 30
    Const M098_EG1_E6_sl_cd = 31
    Const M098_EG1_E6_rcpt_biz_area = 32
    Const M098_EG1_E6_ref_po_no = 33
    Const M098_EG1_E6_ref_po_seq_no = 34
    Const M098_EG1_E6_hs_cd = 35
    Const M098_EG1_E6_over_tol = 36
    Const M098_EG1_E6_under_tol = 37
    Const M098_EG1_E6_vat_type = 38
    Const M098_EG1_E6_vat_inc_flag = 39
    Const M098_EG1_E6_vat_rate = 40
    Const M098_EG1_E6_vat_doc_amt = 41
    Const M098_EG1_E6_vat_loc_amt = 42

    Const M098_EG1_E7_bp_cd = 43    '[CONVERSION INFORMATION]  View Name : exp_item b_biz_partner
    Const M098_EG1_E7_bp_nm = 44
    Const M098_EG1_E5_spec = 45
    
    '�̼��� �߰�(�ܰ���å ����)
    Const M098_EG1_E6_prc_type_cd =46
    

    Dim E3_b_biz_partner
    Const M098_E3_bp_cd = 0
    Const M098_E3_bp_nm = 1

    Dim E4_b_pur_grp
    Const M098_E4_pur_grp = 0
    Const M098_E4_pur_grp_nm = 1
    
    Dim E5_b_item
    Const M098_E5_item_cd = 0
    Const M098_E5_item_nm = 1
   
    If Len(Trim(Request("txtFrDt"))) Then
		If UNIConvDate(Request("txtFrDt")) = "" Then
		    Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		    Call LoadTab("parent.frm1.txtFrDt", 0, I_MKSCRIPT)
		    Exit Sub
		End If
	End If
	
	If Len(Trim(Request("txtToDt"))) Then
		If UNIConvDate(Request("txtToDt")) = "" Then
		    Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		    Call LoadTab("parent.frm1.txtToDt", 0, I_MKSCRIPT)
		    Exit Sub
		End If
	End If

	lgStrPrevKey = Request("lgStrPrevKey")
    If lgStrPrevKey <> "" then	
        arrValue = Split(lgStrPrevKey, gColSep)		
		l2_m_po_hdr_po_no = arrValue(0)
		I1_m_po_dtl_po_seq = arrValue(1)
	else			
		l2_m_po_hdr_po_no = ""
		I1_m_po_dtl_po_seq = ""
	End If			

   Set iPM3G1F8 = Server.CreateObject("PM3G1F8.cMListPoDtlFixPrcS")    

     '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
   If CheckSYSTEMError(Err,True) = true Then 		
		Set iPM3G1F8 = Nothing												'��: ComProxy Unload
		Exit Sub														'��: �����Ͻ� ���� ó���� ������ 
	End if

    I4_m_pur_ord_hdr(M098_I4_po_no)		= Trim(Request("txtPoNo"))		'���ֹ�ȣ 
    I5_b_biz_partner		= Trim(Request("txtSupplierCd"))			'����ó 
    I6_b_pur_grp			= Trim(Request("txtGroupCd"))				'���ű׷� 
    l7_b_cfm_flg			= Trim(Request("txtCfmFlg"))
    I8_b_item_cd			= Trim(Request("txtItemCd"))				'ǰ�� 
    
    if Request("txtFrDt") = "" then
		I4_m_pur_ord_hdr(M098_I4_po_dt)	= "1900-01-01"
	else
		I4_m_pur_ord_hdr(M098_I4_po_dt)	= uniConvDate(Request("txtFrDt"))				'Fr������ 
	End if
	if Request("txtToDt") = "" then
		I3_to_po_dt	= "2999-12-31"
	else
		I3_to_po_dt	= uniConvDate(Request("txtToDt"))				'To������ 
    End if
    Call iPM3G1F8.M_LIST_PO_DTL_FIX_PRC_SVR (gStrGlobalCollection, _
											C_SHEETMAXROWS_D, _
											I1_m_po_dtl_po_seq, _
											l2_m_po_hdr_po_no, _
											I3_to_po_dt, _
											I4_m_pur_ord_hdr, _
											I5_b_biz_partner, _
											I6_b_pur_grp, _
											l7_b_cfm_flg, _
											I8_b_item_cd, _
											E1_m_po_dtl_next, _
											E2_m_pur_ord_dtl, _
											EG1_exp_grp, _
											E3_b_biz_partner, _
											E4_b_pur_grp, _
											E5_b_item)			

	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	 If CheckSYSTEMError(Err,True) = true Then 		
		Set iPM3G1F8 = Nothing												'��: ComProxy Unload
		Response.Write "<Script Language=vbscript>" & vbCr
		Response.Write "with parent" & vbCr
		Response.Write "	.frm1.txtSupplierNm.value = """ & ConvSPChars(E3_b_biz_partner(M098_E3_bp_nm))      & """" & vbCr
		Response.Write "	.frm1.txtGroupNm.value    = """ & ConvSPChars(E4_b_pur_grp(M098_E4_pur_grp_nm))     & """" & vbCr
		Response.Write "	.frm1.txtGroupCd.focus"   & vbCr
		Response.Write "End With "   & vbCr
		Response.Write "</Script>"                  & vbCr
		Exit Sub														'��: �����Ͻ� ���� ó���� ������ 
	 End if
	 
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "with parent" & vbCr
	Response.Write "	.frm1.txtSupplierNm.value = """ & ConvSPChars(E3_b_biz_partner(M098_E3_bp_nm))      & """" & vbCr
	Response.Write "	.frm1.txtGroupNm.value    = """ & ConvSPChars(E4_b_pur_grp(M098_E4_pur_grp_nm))     & """" & vbCr
	Response.Write "	.frm1.txtItemNm.value    = """ & ConvSPChars(E5_b_item(M098_E5_item_nm))     		& """" & vbCr
	Response.Write "End With "   & vbCr
    Response.Write "</Script>"                  & vbCr
    
	 If CheckSYSTEMError(Err,True) = true Then 		
		Set iPM3G1F8 = Nothing												'��: ComProxy Unload
		Exit Sub														'��: �����Ͻ� ���� ó���� ������ 
	 End if

	iLngMaxRow = Request("txtMaxRows")											'Save previous Maxrow                                                
    GroupCount = UBound(EG1_exp_grp,1)
	iMax = GroupCount
	ReDim PvArr(iMax)
	'-----------------------
	'Result data display area
	'----------------------- 
	For iLngRow = 0 To GroupCount
		If  iLngRow = C_SHEETMAXROWS_D  Then
		   StrNextKey = ConvSPChars(E1_m_po_dtl_next(M098_E1_po_no)) & gColSep & ConvSPChars(E2_m_pur_ord_dtl(M098_E2_po_seq_no))
           Exit For
        End If  
        lgCurrency = ConvSPChars(EG1_exp_grp(iLngRow,M098_EG1_E3_po_cur))
        '�̼��� �߰�(�ܰ���å ����)
        istrData = istrData & Chr(11) & "0"
        if ConvSPChars(EG1_exp_grp(iLngRow,M098_EG1_E6_prc_type_cd )) = "T" Then		'�ܰ������� Description
			istrData = istrData & Chr(11) & "���ܰ�"
		Else
			istrData = istrData & Chr(11) & "���ܰ�"
		End If			
        
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow,M098_EG1_E3_po_no ))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow,M098_EG1_E6_po_seq_no ))		'���ּ��� 
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow,M098_EG1_E4_plant_cd))		'���� 
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow,M098_EG1_E4_plant_nm))		'����� 
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow,M098_EG1_E5_item_cd))			'ǰ�� 
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow,M098_EG1_E5_item_nm))			'ǰ��� 
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow,M098_EG1_E5_spec))			'�԰� 
        
        istrData = istrData & Chr(11) & UNIDateClientFormat(EG1_exp_grp(iLngRow,M098_EG1_E3_po_dt))		
        
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow,M098_EG1_E6_po_qty), ggQty.DecPoint,0)		'���ּ��� 
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow,M098_EG1_E6_po_unit))	
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow,M098_EG1_E6_po_prc),ggUnitCost.DecPoint,0) '���ܰ� 
        istrData = istrData & Chr(11) & UNINumClientFormat(0,ggUnitCost.DecPoint,0)			'���ܰ�		
        istrData = istrData & Chr(11) & "0"

        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow,M098_EG1_E3_po_cur))		'ȭ��	
'���� ���ۺκ�(�ݾ� ���� vat������ �����ϼ���...)
     
        If ConvSPChars(EG1_exp_grp(iLngRow,M098_EG1_E6_vat_inc_flag)) = "1" Then
            istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(iLngRow,M098_EG1_E6_po_doc_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")	'�ݾ� 
        Else 
            istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(cdbl(EG1_exp_grp(iLngRow,M098_EG1_E6_po_doc_amt))+cdbl(EG1_exp_grp(iLngRow,M098_EG1_E6_vat_doc_amt)),lgCurrency,ggAmtOfMoneyNo,"X","X")	'�ݾ� 
        End If
        
        istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(iLngRow,M098_EG1_E6_po_doc_amt),lgCurrency,ggAmtOfMoneyNo,"X","X")	'���ݾ� 
        istrData = istrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(iLngRow,M098_EG1_E6_vat_doc_amt),lgCurrency,ggAmtOfMoneyNo,gTaxRndPolicyNo,"X")	                'vat �ݾ� 
        
        If ConvSPChars(EG1_exp_grp(iLngRow,M098_EG1_E6_vat_inc_flag)) = "2" Then   'vat ���Կ��� 
            istrData = istrData & Chr(11) & "����"
        Else
            istrData = istrData & Chr(11) & "����"
        End If
        
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow,M098_EG1_E6_vat_inc_flag))        'vat flg
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow,M098_EG1_E6_vat_type))       'vat type
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow,M098_EG1_E2_minor_nm))          'vat �� 
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow,M098_EG1_E6_vat_rate),ggExchRate.DecPoint,0) 'vat �� 
 '������      
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow,M098_EG1_E7_bp_cd))	
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow,M098_EG1_E7_bp_nm))												    
        ' 2005-10-21 ���Լ��� > 0 -> Message
        istrData = istrData & Chr(11) & ConvSPChars(EG1_exp_grp(iLngRow,M098_EG1_E6_iv_qty))												    

        istrData = istrData & Chr(11) & iLngMaxRow + iLngRow                             
        istrData = istrData & Chr(11) & Chr(12)       

		PvArr(iLngRow) = istrData
		istrData=""
    Next  
    istrData = Join(PvArr, "")

    Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent" & vbCr												'��: ȭ�� ó�� ASP �� ��Ī�� 
    Response.Write "	.ggoSpread.Source          =  .frm1.vspdData "         & vbCr
    Response.Write "	.ggoSpread.SSShowData        """ & istrData	    & """" & ",""F""" & vbCr
    Response.Write "	.lgStrPrevKey              = """ & StrNextKey   & """" & vbCr 

    Response.Write  "    Call Parent.ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,-1,-1,.C_PoCurrency,.C_PoPrice1,""C"" ,""I"",""X"",""X"")" & vbCr
    Response.Write  "    Call Parent.ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,-1,-1,.C_PoCurrency,.C_PoPrice2,""C"" ,""I"",""X"",""X"")" & vbCr
    Response.Write  "    Call Parent.ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,-1,-1,.C_PoCurrency,.C_PoAmt,""A"" ,""I"",""X"",""X"")" & vbCr
    Response.Write  "    Call Parent.ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,-1,-1,.C_PoCurrency,.C_NetPoAmt,""A"" ,""I"",""X"",""X"")" & vbCr
    Response.Write  "    Call Parent.ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,-1,-1,.C_PoCurrency,.C_VatAmt,""A"" ,""I"",""X"",""X"")" & vbCr
     
	Response.Write " .frm1.hdnSupplier.value = """ & ConvSPChars(Request("txtSupplierCd")) & """" & vbCr
	Response.Write " .frm1.hdnGroup.value    = """ & ConvSPChars(Request("txtGroupCd"))    & """" & vbCr
	Response.Write " .frm1.hdnFrDt.value     = """ & ConvSPChars(Request("txtFrDt"))       & """" & vbCr
	Response.Write " .frm1.hdnToDt.value     = """ & ConvSPChars(Request("txtToDt"))       & """" & vbCr
	
	Response.Write " .frm1.hdnPoNo.value   = """ & ConvSPChars(Request("txtPoNo"))   & """" & vbCr
	
    Response.Write " .DbQueryOk "		    	  & vbCr 
    Response.Write " .frm1.vspdData.focus "		  & vbCr 
    Response.Write "End With" & vbCr
    Response.Write "</Script>" & vbCr

    Set iPM3G1F8 = Nothing    
End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
	Dim iPM3G1FP																	'�� : �Է�/������ ComProxy Dll ��� ���� 
    Dim iErrorPosition
	Dim itxtSpread
    Dim itxtSpreadArr
    Dim itxtSpreadArrCount

    Dim iCUCount
    Dim ii
             
    itxtSpread = ""
             
    iCUCount = Request.Form("txtCUSpread").Count
             
    itxtSpreadArrCount = -1
             
    ReDim itxtSpreadArr(iCUCount)
             
    For ii = 1 To iCUCount
        itxtSpreadArrCount = itxtSpreadArrCount + 1
        itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(ii)
    Next
    itxtSpread = Join(itxtSpreadArr,"")

    Response.Write "<Script language=vbs> " & vbCr   
    Response.Write "Parent.RemovedivTextArea "      & vbCr   
    Response.Write "</Script> "      & vbCr   
    
	On Error Resume Next
    Err.Clear																		'��: Protect system from crashing

	If Len(Trim(Request("txtStampDt"))) Then
		If UNIConvDate(Request("txtStampDt")) = "" Then
		    Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		    Call LoadTab("parent.frm1.txtStampDt", 0, I_MKSCRIPT)
		    Exit Sub	
		End If
	End If

    Set iPM3G1FP = Server.CreateObject("PM3G1FP.cMFixPurOrdPriceS")    

    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
   If CheckSYSTEMError(Err,True) = true Then 		
		Set iPM3G1FP = Nothing												'��: ComProxy Unload
		Exit Sub														'��: �����Ͻ� ���� ó���� ������ 
	
	End if

	Call iPM3G1FP.M_FIX_PUR_ORD_PRICE_SVR(gStrGlobalCollection, itxtSpread, iErrorPosition) 
                   
    If CheckSYSTEMError2(Err,True,iErrorPosition & "��:","","","","") = True Then
	 	Set iPM3G1FP = Nothing
	 	Call SheetFocus(iErrorPosition, 1, I_MKSCRIPT)
	 	Exit Sub
	 End If                  
             
	Response.Write "<Script language=vbs> " & vbCr         
    Response.Write " Parent.DbSaveOk "      & vbCr							'��: ȭ�� ó�� ASP �� ��Ī�� 
    Response.Write "</Script> "				& vbCr      

    Set iPM3G1FP = Nothing	  
        
End Sub  
'============================================================================================================
' Name : SublookupPrice
' Desc : 
'============================================================================================================
Sub SublookupPrice()
	
	On Error Resume Next
    Err.Clear																		'��: Protect system from crashing
	
	
	Dim iPM3G1P9
	
	Dim I1_m_supplier_item_price 
	Dim I2_b_biz_partner_bp_cd 
	Dim I3_b_item_item_cd 
	Dim I4_b_plant_plant_cd 
	'�̼����߰�(�ܰ������Ģ)
	Dim l5_b_price_type_cd
	
	Dim E1_m_supplier_item_price 
	Dim E2_b_item 
	Dim E3_b_plant 
	Dim E4_b_storage_location
	Dim E5_b_hs_code 
	Dim E6_m_supplier_item_by_plant 
	Dim E7_b_minor 
	
	Const M106_I1_pur_unit = 0    '  View Name : imp m_supplier_item_price
    Const M106_I1_pur_cur = 1
    Const M106_I1_valid_fr_dt = 2
    ReDim I1_m_supplier_item_price(M106_I1_valid_fr_dt)

    Const M106_E1_pur_prc = 0    '  View Name : exp m_supplier_item_price

    Const M106_E2_item_cd = 0    '  View Name : exp b_item
    Const M106_E2_item_nm = 1
    Const M106_E2_spec = 2
    Const M106_E2_vat_type = 3
    Const M106_E2_vat_rate = 4
    ReDim E2_b_item(M106_E2_vat_rate)

    Const M106_E3_plant_cd = 0    '  View Name : exp b_plant
    Const M106_E3_plant_nm = 1
    ReDim E3_b_plant(M106_E3_plant_nm)
    
    Const M106_E4_sl_cd = 0    '  View Name : exp b_storage_location
    Const M106_E4_sl_nm = 1
    ReDim E4_b_storage_location(M106_E4_sl_nm)

    Const M106_E5_hs_cd = 0    '  View Name : exp b_hs_code
    Const M106_E5_hs_nm = 1
    ReDim E5_b_hs_code(M106_E5_hs_nm)

    Const M106_E6_pur_priority = 0    '  View Name : exp m_supplier_item_by_plant
    Const M106_E6_pur_unit = 1
    Const M106_E6_sppl_item_cd = 2
    Const M106_E6_sppl_item_nm = 3
    Const M106_E6_sppl_item_spec = 4
    Const M106_E6_maker_nm = 5
    Const M106_E6_valid_fr_dt = 6
    Const M106_E6_valid_to_dt = 7
    Const M106_E6_usage_flg = 8
    Const M106_E6_sppl_sales_prsn = 9
    Const M106_E6_sppl_tel_no = 10
    Const M106_E6_sppl_dlvy_lt = 11
    Const M106_E6_under_tol = 12
    Const M106_E6_over_tol = 13
    Const M106_E6_min_qty = 14
    Const M106_E6_def_flg = 15
    Const M106_E6_max_qty = 16
    ReDim E6_m_supplier_item_by_plant(M106_E6_max_qty)

    Const M106_E7_minor_nm = 0    '  View Name : exp_vat_nm b_minor
    Const M106_E7_minor_cd = 1
    ReDim E7_b_minor(M106_E7_minor_cd)

	

	
    Set iPM3G1P9 = Server.CreateObject("PM3G1P9.cMLookupPriceS")    
    
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err,True) = true Then 		
		Set iPM3G1P9 = Nothing												'��: ComProxy Unload
		Exit Sub														'��: �����Ͻ� ���� ó���� ������ 
	End if
	
	I1_m_supplier_item_price(M106_I1_valid_fr_dt)	= UNIConvDate(Trim(Request("txtStampDt")))
	I2_b_biz_partner_bp_cd							= Trim(Request("txtBpCd"))
	I3_b_item_item_cd								= Trim(Request("txtItemCd"))
	I4_b_plant_plant_cd								= Trim(Request("txtPlantCd"))
	'�̼��� �߰� 
	l5_b_price_type_cd								= Trim(Request("txtPrcType"))
	
	I1_m_supplier_item_price(M106_I1_pur_unit)		= Trim(Request("txtUnit"))
	I1_m_supplier_item_price(M106_I1_pur_cur)		= Trim(Request("txtCurrency"))

	Call iPM3G1P9.M_LOOKUP_PRICE_SVR(gStrGlobalCollection, _
									I1_m_supplier_item_price, _
									I2_b_biz_partner_bp_cd, _
									I3_b_item_item_cd, _
									I4_b_plant_plant_cd, _
									l5_b_price_type_cd, _
									E1_m_supplier_item_price, _
									E2_b_item, _
									E3_b_plant, _
									E4_b_storage_location, _
									E5_b_hs_code, _
									E6_m_supplier_item_by_plant, _
									E7_b_minor)

	If CheckSYSTEMError(Err,True) = true Then 		
		Set iPM3G1P9 = Nothing												'��: ComProxy Unload
		Response.Write "<Script Language=vbscript>" & vbCr		
        Response.Write "Dim PoPrice1              " & vbCr
        Response.Write "parent.frm1.vspdData.Row  = """ & ConvSPChars(Trim(Request("txtRow"))) & """" & vbCr 
        Response.Write "parent.frm1.vspdData.Col  = parent.C_PoPrice1 " & vbCr
        Response.Write "PoPrice1 = Parent.frm1.vspdData.Text " & vbCr
        Response.Write "Parent.frm1.vspdData.Col  = Parent.C_PoPrice2 " & vbCr
        '�̼��� ���� 
		Response.Write "Parent.frm1.vspdData.Text = PoPrice1 " & vbCr
		'Response.Write "Parent.frm1.vspdData.Text = """ & UNINumClientFormat(0,ggUnitCost.DecPoint,0)   & """" & vbCr
        Response.Write "Parent.vspdData_Change Parent.C_PoPrice2 , """ & Trim(Request("txtRow")) & """" & vbCr
        Response.Write "</Script> " & vbCr
		Exit Sub														'��: �����Ͻ� ���� ó���� ������ 

	End if
	
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "parent.frm1.vspdData.Row  = """ & ConvSPChars(Trim(Request("txtRow"))) & """" & vbCr
	Response.Write "Parent.frm1.vspdData.Col  = Parent.C_PoPrice2 " & vbCr
	Response.Write "Parent.frm1.vspdData.Text = """ & UNINumClientFormat(E1_m_supplier_item_price(M106_E1_pur_prc),ggUnitCost.DecPoint,0)   & """" & vbCr
    Response.Write "Parent.vspdData_Change Parent.C_PoPrice2 , """ & ConvSPChars(Trim(Request("txtRow"))) & """" & vbCr
    Response.Write "</Script>"                  & vbCr

	Set iPM3G1P9 = Nothing

End Sub

'============================================================================================================
' Name : lookupPriceForSelection
' Desc :
'============================================================================================================
Sub lookupPriceForSelection()

	On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear          

    Dim iPM3G1P9
    Dim iLngMaxRow
    Dim iLngRow
  
    Dim I1_m_supplier_item_price 
	Dim I2_b_biz_partner_bp_cd 
	Dim I3_b_item_item_cd 
	Dim I4_b_plant_plant_cd 
	'�̼����߰�(�ܰ������Ģ)
	Dim l5_b_price_type_cd
		
	Dim E1_m_supplier_item_price 
	Dim E2_b_item 
	Dim E3_b_plant 
	Dim E4_b_storage_location
	Dim E5_b_hs_code 
	Dim E6_m_supplier_item_by_plant 
	Dim E7_b_minor 
	
	Dim lgPriceMsg
	
	Const M106_I1_pur_unit = 0    '  View Name : imp m_supplier_item_price
    Const M106_I1_pur_cur = 1
    Const M106_I1_valid_fr_dt = 2
    '�̼��� �߰� 
    Const M106_l1_po_price = 3
    ReDim I1_m_supplier_item_price(M106_l1_po_price)

    Const M106_E1_pur_prc = 0    '  View Name : exp m_supplier_item_price

    Const M106_E2_item_cd = 0    '  View Name : exp b_item
    Const M106_E2_item_nm = 1
    Const M106_E2_spec = 2
    Const M106_E2_vat_type = 3
    Const M106_E2_vat_rate = 4
    ReDim E2_b_item(M106_E2_vat_rate)

    Const M106_E3_plant_cd = 0    '  View Name : exp b_plant
    Const M106_E3_plant_nm = 1
    ReDim E3_b_plant(M106_E3_plant_nm)
    
    Const M106_E4_sl_cd = 0    '  View Name : exp b_storage_location
    Const M106_E4_sl_nm = 1
    ReDim E4_b_storage_location(M106_E4_sl_nm)

    Const M106_E5_hs_cd = 0    '  View Name : exp b_hs_code
    Const M106_E5_hs_nm = 1
    ReDim E5_b_hs_code(M106_E5_hs_nm)

    Const M106_E6_pur_priority = 0    '  View Name : exp m_supplier_item_by_plant
    Const M106_E6_pur_unit = 1
    Const M106_E6_sppl_item_cd = 2
    Const M106_E6_sppl_item_nm = 3
    Const M106_E6_sppl_item_spec = 4
    Const M106_E6_maker_nm = 5
    Const M106_E6_valid_fr_dt = 6
    Const M106_E6_valid_to_dt = 7
    Const M106_E6_usage_flg = 8
    Const M106_E6_sppl_sales_prsn = 9
    Const M106_E6_sppl_tel_no = 10
    Const M106_E6_sppl_dlvy_lt = 11
    Const M106_E6_under_tol = 12
    Const M106_E6_over_tol = 13
    Const M106_E6_min_qty = 14
    Const M106_E6_def_flg = 15
    Const M106_E6_max_qty = 16
    ReDim E6_m_supplier_item_by_plant(M106_E6_max_qty)

    Const M106_E7_minor_nm = 0    '  View Name : exp_vat_nm b_minor
    Const M106_E7_minor_cd = 1
    ReDim E7_b_minor(M106_E7_minor_cd)
    
    If Len(Trim(Request("txtStampDt"))) Then
		If UNIConvDate(Request("txtStampDt")) = "" Then
		    Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		    Call LoadTab("parent.frm1.txtStampDt", 0, I_MKSCRIPT)
		    Exit Sub	
		End If
	End If

    Set iPM3G1P9 = Server.CreateObject("PM3G1P9.cMLookupPriceS")      
    
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
   If CheckSYSTEMError(Err,True) = true Then 		
		Set iPM3G1P9 = Nothing												'��: ComProxy Unload
		Exit Sub														'��: �����Ͻ� ���� ó���� ������ 
	
	End if

	
	iLngMaxRow = CInt(Request("txtMaxRows"))											'��: �ִ� ������Ʈ�� ���� 

	Dim arrVal, arrTemp																'��: Spread Sheet �� ���� ���� Array ���� 
	Dim strStatus																	'��: Sheet �� ���� Row�� ���� (Create/Update/Delete)
	Dim	lGrpCnt
	Dim strVal	
	Dim priceTemp																	'��: Group Count

	arrTemp = Split(Request("txtSpread"), gRowSep)									'��: Spread Sheet ������ ��� �ִ� Element�� 
	
    '-----------------------
    'Data manipulate area
    '-----------------------
    
    lGrpCnt = 0
    lgPriceMsg = ""
    
	ReDim returnValue(iLngMaxRow)
	'�̼��� �߰� 
	l5_b_price_type_cd								= Trim(Request("txtPrcType"))		
	

	
    For iLngRow = 1 To iLngMaxRow
    
		lGrpCnt = lGrpCnt +1														'��: Group Count
		 
		arrVal = Split(arrTemp(iLngRow-1), gColSep)

		I1_m_supplier_item_price(M106_I1_valid_fr_dt) = UNIConvDate(Trim(Request("txtStampDt")))
		I2_b_biz_partner_bp_cd						  = Trim(arrVal(0))
		I3_b_item_item_cd							  = Trim(arrVal(1))
		I4_b_plant_plant_cd							  = Trim(arrVal(2))
		I1_m_supplier_item_price(M106_I1_pur_unit)	  = Trim(arrVal(3))
		I1_m_supplier_item_price(M106_I1_pur_cur)	  = Trim(arrVal(4))
		priceTemp									  = Trim(arrVal(5))

		Call iPM3G1P9.M_LOOKUP_PRICE_SVR(gStrGlobalCollection, _
										I1_m_supplier_item_price, _
										I2_b_biz_partner_bp_cd, _
										I3_b_item_item_cd, _
										I4_b_plant_plant_cd, _
										l5_b_price_type_cd, _
										E1_m_supplier_item_price, _
										E2_b_item, _
										E3_b_plant, _
										E4_b_storage_location, _
										E5_b_hs_code, _
										E6_m_supplier_item_by_plant, _
										E7_b_minor)
										
		If E1_m_supplier_item_price(M106_E1_pur_prc) = "" then
			E1_m_supplier_item_price(M106_E1_pur_prc) = priceTemp
			lgPriceMsg = lgPriceMsg & "[" & Trim(arrVal(6)) & "]"
		End If						
		    
		returnValue(iLngRow) = UNINumClientFormat(E1_m_supplier_item_price(M106_E1_pur_prc),ggUnitCost.DecPoint,0)
		strval = strval & returnValue(iLngRow) & gColSep
		strval = strval & Trim(arrVal(6)) & gRowSep
			
    Next
    
	Set iPM3G1P9 = Nothing
	
	'�̼��� �߰��κ� 
	Dim rowindex, rowCount, resultindex
	Dim arrSpread
	
	rowCount = Request("hdnMaxRows")
	resultindex = 0
	arrTemp = Split(strval, gRowSep)
	
	For rowindex = 0 to UBOUND(arrTemp,1)-1
	
		arrSpread = Split(arrTemp(rowindex) , gColSep)
	
		Response.Write "<script language=vbscript>" & vbCr
		Response.Write " Parent.frm1.vspdData.Row  = """ & arrSpread(1) & """" & vbCr
		Response.Write " Parent.frm1.vspdData.Col  = Parent.C_PoPrice2 "   & vbCr
		Response.Write " Parent.frm1.vspdData.Text = """ & arrSpread(0) & """" & vbCr
		Response.Write "</script>" & vbCr	
	
	Next
	
' === 2005.07.06 �ܰ� �ϰ� �ҷ����� ���� ���� ===========================================	
		Response.Write "<script language=vbscript>" & vbCr	
		Response.Write " Call parent.btnCallPrice_Ok() " & vbCr
		Response.Write "</script>" & vbCr	
' === 2005.07.06 �ܰ� �ϰ� �ҷ����� ���� ���� ===========================================	
		
	if lgPriceMsg <> "" Then
		Call DisplayMsgBox("173221", vbInformation, lgPriceMsg , "", I_MKSCRIPT)
	End if

End Sub  

'==============================================================================
' Function : SheetFocus
' Description : �����߻��� Spread Sheet�� ��Ŀ���� 
'==============================================================================
Function SheetFocus(Byval lRow, Byval lCol, Byval iLoc)
	
	If Trim(lRow) = "" Then Exit Function
	If iLoc = I_INSCRIPT Then
		strHTML = "parent.frm1.vspdData.focus" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Row = " & lRow & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Col = " & lCol & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Action = 0" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelStart = 0 " & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
		Response.Write strHTML
	ElseIf iLoc = I_MKSCRIPT Then
		strHTML = "<" & "Script LANGUAGE=VBScript" & ">" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.focus" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Row = " & lRow & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Col = " & lCol & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Action = 0" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelStart = 0 " & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
		strHTML = strHTML & "</" & "Script" & ">" & vbCrLf
		Response.Write strHTML
	End If
End Function

%>
