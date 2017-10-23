<%
'********************************************************************************************************
'*  1. Module Name          : Sales & Distribution														*
'*  2. Function Name        :																			*
'*  3. Program ID           : S5212MB1    																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : 수출 B/L 내역등록															*
'*  6. Comproxy List        : PS7G128.cSListBillDtlSvr,PS7G121.cSBillDtlSvr,PS7G115.cSPostOpenArSvr     * 
'*  7. Modified date(First) : 2000/04/24																*
'*  8. Modified date(Last)  : 2002/11/15																*
'*  9. Modifier (First)     : Kim Hyungsuk																*
'* 10. Modifier (Last)      : AHN TAE HEE																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/04/24 : Coding Start												*
'*                            2. 2002/07/04 : VB CONVERSION												*
'*							  3. 2002/11/15 : UI성능 적용												*				
'********************************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%
Call LoadBasisGlobalInf()
Call loadInfTB19029B("I", "*","NOCOOKIE","MB") 
Call LoadBNumericFormatB("I","*","NOCOOKIE","MB") 
																	'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
On Error Resume Next

Dim pvCB 
Call HideStatusWnd

Dim strMode																		'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim LngRow
Dim intGroupCount

'------------------  
Dim iLngLastRow  
Dim iStrNextKey
Dim iLngSheetMaxRows
Dim iArrCols
Dim iArrRows
Dim iLngRow
'------------------

strMode = Request("txtMode")													'☜ : 현재 상태를 받음 

Select Case strMode
	Case CStr(UID_M0001)														'☜: 현재 조회/Prev/Next 요청을 받음 

		Err.Clear																'☜: Protect system from crashing
		
		Dim iPS7G128
		Dim iPS7G121
		Dim iPS7G115
		Dim lgCurrency
		Dim lgArrGlFlag
		Dim lgStrGlFlag

	'---------------------------------- B/L Header Data Query ----------------------------------
		Call SubOpenDB(lgObjConn)
		call SubMakeSQLStatements

		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                 'If data not exists
		    lgObjRs.Close
		    lgObjConn.Close
		    Set lgObjRs = Nothing
		    Set lgObjConn = Nothing
			'B/L정보가 없습니다.
			Call DisplayMsgBox("205300", vbInformation, "", "", I_MKSCRIPT)             '☜: No data is found. 
			%>
			<Script Language=vbscript>
				parent.SetDefaultVal
				Call Parent.SetToolbar("11000000000011")
			</Script>
			<%
		    Response.End
		End If

		'-----------------------
		'Result data display area
		'-----------------------
		lgCurrency = ConvSPChars(lgObjRs("Cur"))

%>
	<Script Language=VBScript>
	With parent.frm1
		Dim strDt

		.txtCurrency.value		= "<%=lgCurrency%>"

		Call parent.CurFormatNumericOCX

		.txtApplicant.value		= "<%=ConvSPChars(lgObjRs("applicant"))%>"
		.txtApplicantNm.value	= "<%=ConvSPChars(lgObjRs("applicant_nm"))%>"
		.txtSONo.value			= "<%=ConvSPChars(Trim(lgObjRs("so_no")))%>"
		.txtSalesGroup.value	= "<%=ConvSPChars(lgObjRs("sales_grp"))%>"
		.txtSalesGroupNm.value	= "<%=ConvSPChars(lgObjRs("sales_grp_nm"))%>"
		.txtBillType.value		= "<%=ConvSPChars(lgObjRs("bill_type"))%>"
		.txtBillTypeNm.value	= "<%=ConvSPChars(lgObjRs("bill_type_nm"))%>"
		.txtWeightUnit.value	= "<%=ConvSPChars(lgObjRs("weight_unit"))%>"
		.txtVolumnUnit.value	= "<%=ConvSPChars(lgObjRs("volumn_unit"))%>"
		.txtPayTerms.value		= "<%=ConvSPChars(lgObjRs("pay_meth"))%>"
		.txtPayTermsNm.value	= "<%=ConvSPChars(lgObjRs("pay_meth_nm"))%>"
		.txtIncoterms.value		= "<%=ConvSPChars(lgObjRs("incoterms"))%>"
		.txtIncotermsNm.value	= "<%=ConvSPChars(lgObjRs("incoterms_nm"))%>"
		.txtXchgRate.value		= "<%=UNINumClientFormat(lgObjRs("Xchg_rate"), ggExchRate.DecPoint, 0)%>"
		.txtXchgRateOp.value	= "<%=ConvSPChars(lgObjRs("Xchg_rate_op"))%>"
		.txtDocAmt.Text			= "<%=UNINumClientFormatByCurrency(lgObjRs("bill_amt"),lgCurrency,ggAmtOfMoneyNo)%>"
		.txtLocAmt.Text			= "<%=UNINumClientFormatByCurrency(lgObjRs("bill_amt_loc"),gCurrency,ggAmtOfMoneyNo)%>"
		.txtLocCurrency.Value	= "<%=UCase(gCurrency)%>"
		.txtRefFlg.value		= "<%=ConvSPChars(lgObjRs("ref_flag"))%>"
		.txtStatusFlg.value		= "<%=ConvSPChars(lgObjRs("sts"))%>"		
		.txtHBLNo.value			= "<%=ConvSPChars(Request("txtBLNo"))%>"
		.txtHBLIssueDT.value	= "<%=UNIDateClientFormat(lgObjRs("bl_issue_dt"))%>"
		.txtMaxSeq.value		= 0
		
		<% If lgObjRs("post_flag") = "Y" AND Len(Trim(lgObjRs("gl_no"))) Then %>
		<% lgArrGlFlag = Split(lgObjRs("gl_no"), Chr(11)) %>
		<% lgStrGlFlag = lgArrGlFlag(0) %>
		If "<%=lgArrGlFlag(0)%>" = "G" Then	
			'회계전표번호 
			.txtGLNo.value	= "<%=lgArrGlFlag(1)%>"	
		ElseIf "<%=lgArrGlFlag(0)%>" = "T" Then
			'결의전표번호 
			.txtTempGLNo.value	= "<%=lgArrGlFlag(1)%>"	
		Else
			'Batch번호 
			.txtBatchNo.value	= "<%=lgArrGlFlag(1)%>"	
		End If
		<%Else%>
		.txtGLNo.value	= ""	
		.txtTempGLNo.value = ""	
		.txtBatchNo.value	= ""	
		<% End If %>
		
		If "<%=lgObjRs("post_flag")%>" = "Y" Then
			.rdoPostingflg1.Checked = True
			.btnPosting.value = "확정취소"

			if "<%=lgStrGlFlag%>" = "G" Or "<%=lgStrGlFlag%>" = "T" Then
				.btnGLView.disabled = False
			Else
				.btnGLView.disabled = True
			End if
		Else
			.rdoPostingflg2.Checked = True
			.btnPosting.value = "확정"
			.btnGLView.disabled = True
		End If

		If "<%=Trim(lgObjRs("ref_flag"))%>" = "M" then
			.btnPosting.disabled = true
		Else
			.btnPosting.disabled = False
		End If
		
		<% '선수금 현황 버튼 Enable %>
		IF "<%=lgObjRs("PreRcpt_flag")%>" = "Y" Then
			.btnPreRcptView.disabled = False
		Else
			.btnPreRcptView.disabled = True
		End If

		Call parent.CurFormatNumSprSheet
		Call parent.BLHdrQueryOk()

	End With
	</Script>
	<%
		lgObjRs.Close
		lgObjConn.Close
		Set lgObjRs = Nothing
		Set lgObjConn = Nothing

	'---------------------------------- B/L Detail Data Query ----------------------------------
	'--------------
	'Interface 정의 
	'--------------
    'View Name : imp_next s_bill_dtl
    Const S526_I1_bill_seq = 0
    'View Name : imp s_bill_hdr
    Const S526_I2_bill_no = 0
    'View Name : exp_next s_bill_dtl
    Const S526_E1_bill_seq = 0

    'Group Name : exp_grp
    Const S526_EG1_E1_minor_nm = 0 
    Const S526_EG1_E2_cc_seq = 1   
    Const S526_EG1_E3_cc_no = 2    
    Const S526_EG1_E4_lc_seq = 3   
    Const S526_EG1_E5_lc_no = 4    
    Const S526_EG1_E6_bill_seq = 5 
    Const S526_EG1_E6_bill_price = 6
    Const S526_EG1_E6_bill_amt = 7
    Const S526_EG1_E6_vat_amt = 8
    Const S526_EG1_E6_bill_qty = 9
    Const S526_EG1_E6_bill_unit = 10
    Const S526_EG1_E6_remark = 11
    Const S526_EG1_E6_item_acct = 12
    Const S526_EG1_E6_tracking_no = 13
    Const S526_EG1_E6_plant_biz_area = 14
    Const S526_EG1_E6_cost_cd = 15
    Const S526_EG1_E6_hs_no = 16
    Const S526_EG1_E6_cust_item_cd = 17
    Const S526_EG1_E6_bill_amt_loc = 18
    Const S526_EG1_E6_vat_type = 19
    Const S526_EG1_E6_vat_rate = 20
    Const S526_EG1_E6_vat_amt_loc = 21
    Const S526_EG1_E6_cust_po_no = 22
    Const S526_EG1_E6_cust_po_seq = 23
    Const S526_EG1_E6_gross_weight = 24
    Const S526_EG1_E6_net_weight = 25
    Const S526_EG1_E6_volume_size = 26
    Const S526_EG1_E6_ext1_qty = 27
    Const S526_EG1_E6_ext2_qty = 28
    Const S526_EG1_E6_ext3_qty = 29
    Const S526_EG1_E6_ext1_amt = 30
    Const S526_EG1_E6_ext2_amt = 31
    Const S526_EG1_E6_ext3_amt = 32
    Const S526_EG1_E6_ext1_cd = 33
    Const S526_EG1_E6_ext2_cd = 34
    Const S526_EG1_E6_ext3_cd = 35
    Const S526_EG1_E6_vat_inc_flag = 36
    Const S526_EG1_E6_deposit_price = 37
    Const S526_EG1_E6_deposit_amt = 38
    Const S526_EG1_E6_ret_item_flag = 39
    Const S526_EG1_E7_dn_seq = 40  
    Const S526_EG1_E8_dn_no = 41   
    Const S526_EG1_E9_plant_cd = 42
    Const S526_EG1_E10_item_cd = 43
    Const S526_EG1_E10_item_nm = 44
    Const S526_EG1_E10_spec = 45
    Const S526_EG1_E11_so_seq = 46 
    Const S526_EG1_E12_so_no = 47  
    
    Const C_SHEETMAXROWS_D  = 1000
    
    '--------
	'View선언 
	'--------    
    Dim I2_s_bill_hdr
    Dim I1_s_bill_dtl
    Dim EG1_exp_grp
    Dim E1_s_bill_dtl 
    
     '---------------------------------------
    'Data manipulate  area(import view match)
    '----------------------------------------
    redim I2_s_bill_hdr(0)
    
    I2_s_bill_hdr(S526_I2_bill_no) = Trim(Request("txtBLNo"))
    
    redim I1_s_bill_dtl(0)
    
    If Trim(Request("lgStrPrevKey")) = "" then
		I1_s_bill_dtl(S526_I1_bill_seq) = 0
    Else
		I1_s_bill_dtl(S526_I1_bill_seq) = cdbl(Request("lgStrPrevKey"))
	End if	
	 
    
    Set iPS7G128 = Server.CreateObject("PS7G128.cSListBillDtlSvr")    
  
    If CheckSYSTEMError(Err,True) = True Then
		Response.End		
    End If
	
	Call iPS7G128.S_LIST_BILL_DTL_SVR(gStrGlobalCollection , C_SHEETMAXROWS_D , _
	                                 I1_s_bill_dtl ,I2_s_bill_hdr, EG1_exp_grp, E1_s_bill_dtl)

	If CheckSYSTEMError(Err,True) = True Then
       Set iPS7G128 = Nothing		                                                 '☜: Unload Comproxy DLL
       Response.End 
    End If   

    Set iPS7G128 = Nothing	
   
			
	' Client(MA)의 현재 조회된 마직막 Row
	iLngLastRow = CLng(Request("txtMaxRows")) 

	' Set Next key
	If Ubound(EG1_exp_grp,1) = C_SHEETMAXROWS_D Then
		iStrNextKey = E1_s_bill_dtl(S526_E1_bill_seq) 'b/l 순번 
		iLngSheetMaxRows  = C_SHEETMAXROWS_D - 1
	Else
		iStrNextKey = ""
		iLngSheetMaxRows = Ubound(EG1_exp_grp,1)
	End If

	ReDim iArrCols(29)						' Column 수 
	Redim iArrRows(iLngSheetMaxRows)		' 조회된 Row 수만큼 배열 재정의 

   	For iLngRow = 0 To iLngSheetMaxRows
					'품목코드 
		iArrCols(1) = ConvSPChars(EG1_exp_grp(iLngRow,S526_EG1_E10_item_cd))
					'품목명 
		iArrCols(2) = ConvSPChars(EG1_exp_grp(iLngRow,S526_EG1_E10_item_nm))
					'단위 
		iArrCols(3) = ConvSPChars(EG1_exp_grp(iLngRow,S526_EG1_E6_bill_unit))
					'수량 
		iArrCols(4) = UNINumClientFormat(EG1_exp_grp(iLngRow, S526_EG1_E6_bill_qty), ggQty.DecPoint, 0)
					'단가 
		iArrCols(5) = UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(iLngRow,S526_EG1_E6_bill_price),lgCurrency,ggUnitCostNo, "X" , "X")
					'금액 
		iArrCols(6) = UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(iLngRow,S526_EG1_E6_bill_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")
					'부가세포함여부 
		iArrCols(7) = ConvSPChars(EG1_exp_grp(iLngRow,S526_EG1_E6_vat_inc_flag))
					'vat유형 
		iArrCols(8) = ConvSPChars(EG1_exp_grp(iLngRow,S526_EG1_E6_vat_type))
					'vat율 
		iArrCols(9) = UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(iLngRow,S526_EG1_E6_vat_rate), gCurrency, ggExchRateNo, "X" , "X")
					'부가세액 
		iArrCols(10) = UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(iLngRow,S526_EG1_E6_vat_amt),lgCurrency,ggAmtOfMoneyNo, gTaxRndPolicyNo  , "X")
					'원화금액 
		iArrCols(11) = UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(iLngRow,S526_EG1_E6_bill_amt_loc),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo  , "X")
					'부가세원화금액 
		iArrCols(12) = UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(iLngRow,S526_EG1_E6_vat_amt_loc),gCurrency,ggAmtOfMoneyNo, gTaxRndPolicyNo  , "X")
					'총중량 
		iArrCols(13) = UNINumClientFormat(EG1_exp_grp(iLngRow,S526_EG1_E6_gross_weight), ggQty.DecPoint, 0)
					'용적 
		iArrCols(14) = UNINumClientFormat(EG1_exp_grp(iLngRow,S526_EG1_E6_volume_size), ggQty.DecPoint, 0)
					'순중량 
		iArrCols(15) = UNINumClientFormat(EG1_exp_grp(iLngRow,S526_EG1_E6_net_weight), ggQty.DecPoint, 0)
					'Tracking No
		iArrCols(16) = ConvSPChars(EG1_exp_grp(iLngRow,S526_EG1_E6_tracking_no))
					'공장 
		iArrCols(17) = ConvSPChars(EG1_exp_grp(iLngRow,S526_EG1_E9_plant_cd))
					'HS번호 
		iArrCols(18) = ConvSPChars(EG1_exp_grp(iLngRow,S526_EG1_E6_hs_no))
					'C/C번호 
		iArrCols(19) = ConvSPChars(EG1_exp_grp(iLngRow,S526_EG1_E3_cc_no))
					'C/C순번 
		iArrCols(20) = UNINumClientFormat(EG1_exp_grp(iLngRow,S526_EG1_E2_cc_seq), 0, 0)
					'수주번호 
		iArrCols(21) = ConvSPChars(EG1_exp_grp(iLngRow,S526_EG1_E12_so_no))
					'수주순번 
		iArrCols(22) = UNINumClientFormat(EG1_exp_grp(iLngRow,S526_EG1_E11_so_seq), 0, 0)
					'L/C번호 
		iArrCols(23) = ConvSPChars(EG1_exp_grp(iLngRow,S526_EG1_E5_lc_no))
					'L/C순번 
		iArrCols(24) = UNINumClientFormat(EG1_exp_grp(iLngRow,S526_EG1_E4_lc_seq), 0, 0)
					'출하번호 
		iArrCols(25) = ConvSPChars(EG1_exp_grp(iLngRow,S526_EG1_E8_dn_no))
					'출하순번 
		iArrCols(26) = UNINumClientFormat(EG1_exp_grp(iLngRow,S526_EG1_E7_dn_seq), 0, 0)
					'B/L순번 
		iArrCols(27) = UNINumClientFormat(EG1_exp_grp(iLngRow,S526_EG1_E6_bill_seq), 0, 0)
					'품목규격 
		iArrCols(28) = ConvSPChars(EG1_exp_grp(iLngRow,S526_EG1_E10_spec))

   		iArrCols(29) = iLngLastRow + iLngRow 
   		
   		iArrRows(iLngRow) = Join(iArrCols, gColSep)

	Next

	Response.Write "<Script language=vbs> "	& vbCr   
    Response.Write " With Parent	       " & vbCr
        
    Response.Write "	.ggoSpread.Source     = .frm1.vspdData	" & vbCr
    Response.Write "	.ggoSpread.SSShowDataByClip   """ & Join(iArrRows, gColSep & gRowSep) & gColSep & gRowSep & """" & vbCr
 
    Response.Write "	.frm1.txtHBLNo.value  = """ & ConvSPChars(Request("txtBLNo"))   	  & """" & vbCr
        
    Response.Write "	.frm1.vspdData.ReDraw = False    		" & vbCr
    Response.Write "	.SetSpreadColor -1,-1					" & vbCr
  	Response.Write "	.frm1.vspdData.ReDraw = True			" & vbCr
	
    Response.Write "	.lgStrPrevKey = """ & iStrNextKey & """" & vbCr  
    
    Response.Write "	.DbQueryOk                             " & vbCr   
    Response.Write "	.HideNonCCGrid()                       " & vbCr   

    Response.Write " End With       " & vbCr
    Response.Write "</Script> "	& vbCr          

	Case CStr(UID_M0002)
		
		Err.Clear	
		Dim iErrorPosition
		Dim itxtSpread,itxtHBLNo
	
		Set iPS7G121 = Server.CreateObject("PS7G121.cSBillDtlSvr")  
    
		If CheckSYSTEMError(Err,True) = True Then
			Response.End		
		End If
		
	    itxtSpread = Trim(Request("txtSpread"))
	    itxtHBLNo = Trim(Request("txtHBLNo"))
		
		 pvCB = "F"
		Call iPS7G121.S_MAINT_BILL_DTL_SVR (pvCB, gStrGlobalCollection ,itxtSpread, iErrorPosition, _
		                                itxtHBLNo)
    
		If CheckSYSTEMError2(Err, True, iErrorPosition & "행","","","","") = True Then
			Set iPS7G121 = Nothing
			Response.End 
		End If

		Set iPS7G121 = Nothing	
    
		Response.Write "<Script language=vbs> " & vbCr         
		Response.Write " Parent.DBSaveOk "      & vbCr   
		Response.Write "</Script> " 		

	Case CStr("PostFlag")																'☜: 확정 요청 

		Err.Clear												'☜: Protect system from crashing
		
		Dim itxtHBillNo
	    
	    Set iPS7G115 = Server.CreateObject("PS7G115.cSPostOpenArSvr")  
	    
	    If CheckSYSTEMError(Err,True) = True Then
			Response.End		
	    End If
		
		itxtHBillNo = Trim(Request("txtHBillNo"))
		'for 구주Tax
		pvCB = "F"
		 
		Call iPS7G115.S_POST_OPEN_AR_SVR (pvCB, gStrGlobalCollection ,itxtHBillNo )
	    
		If CheckSYSTEMError(Err,True) = True Then
		Set iPS7G115 = Nothing
		Response.End		
		End If
		
		Set iPS7G115 = Nothing	
		
		Response.Write "<Script language=vbs> " & vbCr         
	    Response.Write " Call parent.PostingOk()	 "      & vbCr   
	    Response.Write "</Script> "   				

	Case Else
		Response.End
End Select
'==============================================================================
%>
<SCRIPT LANGUAGE=VBSCRIPT RUNAT=SERVER>
'============================================================================================================
Sub SubMakeSQLStatements()
	Dim strSelectList, strFromList
	
	strSelectList = "SELECT * "
	strFromList = " FROM dbo.ufn_s_GetBLInfo ( " & FilterVar(Request("txtBlNo"), "''", "S") & " , " & FilterVar("N", "''", "S") & " , " & FilterVar("Y", "''", "S") & " , " & FilterVar("Y", "''", "S") & " , " & FilterVar("Y", "''", "S") & " , " & FilterVar("Q", "''", "S") & " )"
	lgstrsql = strSelectList & strFromList
	
	
	'call svrmsgbox(lgstrsql, vbinformation, i_mkscript)
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub
</SCRIPT>
