<%@ LANGUAGE = VBSCript%>
<% Option Explicit %>
<%
'**********************************************************************************************
'*  1. Module Name          : ���� 
'*  2. Function Name        : ����ä�ǰ��� 
'*  3. Program ID           : S5112MB1
'*  4. Program Name         : ����ä�ǳ������ 
'*  5. Program Desc         :
'*  6. Comproxy List        : PS7G128.cSListBillDtlSvr,PS7G121.cSBillDtlSvr,PS7G115.cSPostOpenArSvr
'*  7. Modified date(First) : 2000/04/19
'*  8. Modified date(Last)  : 2002/11/12
'*  9. Modifier (First)     : Cho song hyon
'* 10. Modifier (Last)      : Ahn Tae Hee
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            -2000/04/19 : 3rd ȭ�� Layout & ASP Coding
'*                            -2000/08/11 : 4th ȭ�� Layout
'*                            -2001/12/18 : Date ǥ������ 
'*							  -2002/06/24 : VB conversion
'*							  -2002/11/12 : UI���� ����	
'**********************************************************************************************
%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%Call LoadBasisGlobalInf()
Call loadInfTB19029B("I", "*","NOCOOKIE","MB") 
Call LoadBNumericFormatB("I","*","NOCOOKIE","MB") %>

<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<%													

ON ERROR RESUME Next														

Call HideStatusWnd
'for ����Tax
'@@@@@@@@@@@
Dim pvCB 
'@@@@@@@@@@@

Dim iPS7G121												'�� : ���⳻������Է�/����/������ ComProxy Dll ��� ���� 
Dim iPS7G128												'�� : ���⳻����� ��ȸ�� ComProxy Dll ��� ���� 
Dim iPS7G115												'�� : ���⳻��Ȯ���� ComProxy Dll ��� ���� 
Dim strMode		
Dim iStrNextKey											' ���� �� 
Dim lgStrPrevKey										' ���� �� 
Dim LngMaxRow											' ���� �׸����� �ִ�Row
Dim LngRow
Dim GroupCount															
Dim lgCurrency
Dim lgArrGlFlag
Dim lgStrGlFlag

strMode = Request("txtMode")												'�� : ���� ���¸� ���� 

Select Case strMode

Case CStr(UID_M0001)														'��: ���� ��ȸ/Prev/Next ��û�� ���� 
	Dim iStrSvrData

    Err.Clear                                                                '��: Protect system from crashing

    '-----------------------
    ' ��������� �о�´�.
    '-----------------------
	If Request("txtHQuery") = "T" Then
	
	    Call SubOpenDB(lgObjConn)
	    call SubMakeSQLStatements
	    
		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                 'If data not exists
		    lgObjRs.Close
		    lgObjConn.Close
		    Set lgObjRs = Nothing
		    Set lgObjConn = Nothing
			'����ä�������� �����ϴ�.
			Call DisplayMsgBox("205100", vbInformation, "", "", I_MKSCRIPT)             '��: No data is found. 
			%>
			<Script Language=vbscript>
				parent.SetDefaultVal
				Call parent.SetToolbar("11000000000011")
			</Script>
			<%
		    Response.End
		End If
	%>
	<Script Language=vbscript>
		With parent
			'-----------------------
			' ��������� ������ ǥ���Ѵ�.
			'----------------------- 

			'##### Rounding Logic #####
			'�׻� �ŷ�ȭ�� �켱 
			<%
			lgCurrency = ConvSPChars(lgObjRs("Cur"))
			%>
            
            .frm1.txtCurrency.value			= "<%=lgCurrency%>"
			
			parent.CurFormatNumericOCX
			
			'##########################

			.frm1.txtHBillNo.value			= "<%=ConvSPChars(lgObjRs("bill_no"))%>"
			.frm1.txtHBillDt.value			= "<%=UNIDateClientFormat(lgObjRs("bill_dt"))%>"
			.frm1.txtSoldtoParty.value		= "<%=ConvSPChars(lgObjRs("Sold_to_Party"))%>"
			.frm1.txtSoldtoPartyNm.value	= "<%=ConvSPChars(lgObjRs("Sold_to_Party_Nm"))%>"
			.frm1.txtBillToPartyCd.value	= "<%=ConvSPChars(lgObjRs("Bill_To_Party"))%>"
			.frm1.txtBillToPartyNm.value	= "<%=ConvSPChars(lgObjRs("Bill_To_Party_Nm"))%>"
			.frm1.txtPayTermsCd.value		= "<%=ConvSPChars(lgObjRs("Pay_Meth"))%>"
			.frm1.txtPayTermsNm.value		= "<%=ConvSPChars(lgObjRs("Pay_Meth_Nm"))%>"
			.frm1.txtHSalesGrpCd.value		= "<%=ConvSPChars(lgObjRs("sales_grp"))%>"
			.frm1.txtHSalesGrpNm.value		= "<%=ConvSPChars(lgObjRs("sales_grp_nm"))%>"

			.frm1.txtOriginBillAmt.Text		= "<%=UNIConvNumDBToCompanyByCurrency(lgObjRs("bill_amt"),lgCurrency,ggAmtOfMoneyNo, "X" , "X")%>"
			.frm1.txtVatAmt.Text			= "<%=UNIConvNumDBToCompanyByCurrency(lgObjRs("vat_amt"),lgCurrency,ggAmtOfMoneyNo, gTaxRndPolicyNo, "X")%>"
			.frm1.txtTotBillAmt.Text		= "<%=UNIConvNumDBToCompanyByCurrency(Cdbl(lgObjRs("bill_amt")) + CDbl(lgObjRs("vat_amt")) + CDbl(lgObjRs("deposit_amt")),lgCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")%>"

			.frm1.txtSoNo.value				= "<%=ConvSPChars(lgObjRs("so_no"))%>"
			.frm1.txtXchgRate.Text			= "<%=UNIConvNumDBToCompanyByCurrency(lgObjRs("xchg_rate"), gCurrency, ggExchRateNo, "X" , "X")%>"
			
			<% 'vat�� ȭ�鿡 ���� %>
			.frm1.HVatRate.value			= "<%=UNIConvNumDBToCompanyByCurrency(lgObjRs("vat_rate"), gCurrency, ggExchRateNo, "X" , "X")%>"
			<% 'vatŸ�� ȭ�鿡 ���� %>
			.frm1.HVatType.value			= "<%=ConvSPChars(lgObjRs("vat_type"))%>"
			<% 'ȯ�������ڸ� ȭ�鿡 ���� %>
			.frm1.txtXchgOp.value			= "<%=ConvSPChars(lgObjRs("xchg_rate_op"))%>"
			
			.frm1.txtBillTypeCd.value		= "<%=ConvSPChars(lgObjRs("bill_type"))%>"

			<% '����������� %>
			.frm1.txtSts.value				= "<%=lgObjRs("sts")%>"
			<% '����, L/C���� ���� %>
			.frm1.txtRefFlag.value			= "<%=lgObjRs("ref_flag")%>"
			<% '��ǰ���� %>
			.frm1.txtReverseFlag.value		= "<%=lgObjRs("reverse_flag")%>"

			'VAT������� 
			'VAT��������� ������ ��� VAT ����, ������, ���� Hidden ó�� 
			If "<%=lgObjRs("vat_calc_type")%>" = "2" Then
				.frm1.rdoVatCalcType2.checked = True
			Else
				.frm1.rdoVatCalcType1.checked = True
			End If

			'�ΰ������Կ��� 
			If Trim("<%=lgObjRs("vat_inc_flag")%>") = "1" Then
				.frm1.rdoVatIncFlag1.checked = True
			Else
				.frm1.rdoVatIncFlag2.checked = True
			End If
            
			.frm1.HPostFlag.value = "<%=lgObjRs("post_flag")%>"

			<%If lgObjRs("post_flag") = "Y" AND Len(Trim(lgObjRs("gl_no"))) Then
				lgArrGlFlag = Split(Trim(lgObjRs("gl_no")), Chr(11)) 
				lgStrGlFlag = lgArrGlFlag(0)%>
				
				If "<%=lgArrGlFlag(0)%>" = "G" Then	
					'ȸ����ǥ��ȣ 
					.frm1.txtGLNo.value	= "<%=lgArrGlFlag(1)%>"	
				ElseIf "<%=lgArrGlFlag(0)%>" = "T" Then
					'������ǥ��ȣ 
					.frm1.txtTempGLNo.value	= "<%=lgArrGlFlag(1)%>"	
				Else
					'Batch��ȣ 
					.frm1.txtBatchNo.value	= "<%=lgArrGlFlag(1)%>"	
				End If
			<%Else%>
					.frm1.txtGLNo.value	= ""	
					.frm1.txtTempGLNo.value	= ""	
					.frm1.txtBatchNo.value	= ""	
			<% End If %>
			
			If .frm1.HPostFlag.value = "Y" Then
				
				.frm1.btnPostFlag.value = "Ȯ�����"
				If "<%=lgStrGlFlag%>" = "G" Or "<%=lgStrGlFlag%>" = "T" Then
					.frm1.btnGLView.disabled = False
				Else
					.frm1.btnGLView.disabled = True
				End If
			Else
				.frm1.btnPostFlag.value = "Ȯ��"
				.frm1.btnGLView.disabled = True
			End If

			<% '������ ��Ȳ ��ư Enable %>
			IF "<%=lgObjRs("PreRcpt_flag")%>" = "Y" Then
				.frm1.btnPreRcptView.disabled = False
			Else
				.frm1.btnPreRcptView.disabled = True
			End If

			<%
			lgObjRs.Close
			lgObjConn.Close
			Set lgObjRs = Nothing
			Set lgObjConn = Nothing
			%>
			parent.SetSpreadHidden		
			'-----------------------
			' Rounding Column Set
			'----------------------- 
			parent.CurFormatNumSprSheet
			
			.DbQueryOk														'��: ��ȸ�� ���� 

		End With
	</Script>		
<%
	ElseIf Request("txtHQuery") = "F" Then
		lgCurrency = Request("txtCurrency")			
	End If

	'-----------------------
    ' ���⳻���� �о�´�.
    '-----------------------
    '--------------
	'Interface ���� 
	'--------------
    'View Name : imp_next s_bill_dtl
    Const S526_I1_bill_seq = 0
    'View Name : imp s_bill_hdr
    Const S526_I2_bill_no = 0
    'View Name : exp_next s_bill_dtl
    Const S526_E1_bill_seq = 0
    
    'Group Name : exp_grp
    Const S526_EG1_E1_minor_nm = 0    'View Name : exp_item b_minor
    Const S526_EG1_E2_cc_seq = 1    'View Name : exp_item s_cc_dtl
    Const S526_EG1_E3_cc_no = 2    'View Name : exp_item s_cc_hdr
    Const S526_EG1_E4_lc_seq = 3    'View Name : exp_item s_lc_dtl
    Const S526_EG1_E5_lc_no = 4    'View Name : exp_item s_lc_hdr
    Const S526_EG1_E6_bill_seq = 5    'View Name : exp_item s_bill_dtl
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
    Const S526_EG1_E7_dn_seq = 40    'View Name : exp_item s_dn_dtl
    Const S526_EG1_E8_dn_no = 41    'View Name : exp_item s_dn_hdr
    Const S526_EG1_E9_plant_cd = 42    'View Name : exp_item b_plant
    Const S526_EG1_E10_item_cd = 43    'View Name : exp_item b_item
    Const S526_EG1_E10_item_nm = 44
    Const S526_EG1_E10_spec = 45
    Const S526_EG1_E11_so_seq = 46    'View Name : exp_item s_so_dtl
    Const S526_EG1_E12_so_no = 47    'View Name : exp_item s_so_hdr
        
    Const C_SHEETMAXROWS_D  = 100
    
    '--------
	'View���� 
	'--------    
    Dim I2_s_bill_hdr
    Dim I1_s_bill_dtl
    Dim EG1_exp_grp
    Dim E1_s_bill_dtl 
    
     '---------------------------------------
    'Data manipulate  area(import view match)
    '----------------------------------------
    redim I2_s_bill_hdr(0)
    
    I2_s_bill_hdr(S526_I2_bill_no) = Trim(Request("txtConBillNo"))
    
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
	
	Call iPS7G128.S_LIST_BILL_DTL_SVR (gStrGlobalCollection , C_SHEETMAXROWS_D , _
	                                 I1_s_bill_dtl ,I2_s_bill_hdr, EG1_exp_grp, E1_s_bill_dtl)

	If CheckSYSTEMError(Err,True) = True Then
		Set iPS7G128 = Nothing		                                                 '��: Unload Comproxy DLL
		Response.Write "<Script language=vbs> " & vbCr   
		Response.Write " Parent.frm1.txtConBillNo.Focus" & vbCr   
		Response.Write "</Script> " & vbCr          
		Response.End 
    End If   

    Set iPS7G128 = Nothing	
		'----------------------------
		' ���⳻���� ������ ǥ���Ѵ�.
		'----------------------------
	Dim iLngSheetMaxRows
	Dim iArrCols, iArrRows


	' Set Next key
	If Ubound(EG1_exp_grp,1) = C_SHEETMAXROWS_D Then
		'�����ȣ 
		iStrNextKey = E1_s_bill_dtl(S526_E1_bill_seq)
		iLngSheetMaxRows  = C_SHEETMAXROWS_D - 1
	Else
		iStrNextKey = ""
		iLngSheetMaxRows = Ubound(EG1_exp_grp,1)
	End If

	ReDim iArrCols(36)						' Column �� 
	Redim iArrRows(iLngSheetMaxRows)		' ��ȸ�� Row ����ŭ �迭 ������ 

	LngMaxRow = CInt(Request("txtMaxRows")) + 1

	' ������� �ʴ� �� ���� 
	iArrCols(0)  = ""		' Row Header
	iArrCols(11)  = "" 		'vat���� �˾� 
	iArrCols(19) = "0"		'FOB�ݾ� 
	
	For LngRow = 0 To Ubound(EG1_exp_grp,1)
		iArrCols(1)  = ConvSPChars(EG1_exp_grp(LngRow,S526_EG1_E10_item_cd))		'ǰ���ڵ� 
		iArrCols(2)  = ConvSPChars(EG1_exp_grp(LngRow,S526_EG1_E10_item_nm))		'ǰ��� 
		iArrCols(3)  = ConvSPChars(EG1_exp_grp(LngRow,S526_EG1_E6_tracking_no))		'Tracking No
		iArrCols(4)  = UNINumClientFormat(EG1_exp_grp(LngRow, S526_EG1_E6_bill_qty), ggQty.DecPoint, 0)			'���� 
		iArrCols(5)  = ConvSPChars(EG1_exp_grp(LngRow,S526_EG1_E6_bill_unit))		'���� 
		iArrCols(6) = Trim(EG1_exp_grp(LngRow,S526_EG1_E6_vat_inc_flag))			'�ΰ������Կ��� 

		'�ΰ������Կ��θ� 
		If iArrCols(6) = "1" Then
			iArrCols(7) = "����"
		Else
			iArrCols(7) = "����"
		End If
		
		iArrCols(8)  = UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(LngRow,S526_EG1_E6_bill_price),lgCurrency,ggUnitCostNo, "X" , "X")		'�ܰ� 
		iArrCols(9)  = UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(LngRow,S526_EG1_E6_bill_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")		'�ݾ� 
		iArrCols(10)  = ConvSPChars(EG1_exp_grp(LngRow,S526_EG1_E6_vat_type))		'vat���� 
		iArrCols(12) = ConvSPChars(EG1_exp_grp(LngRow,S526_EG1_E1_minor_nm))		'vat������ 
		iArrCols(13) = UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(LngRow,S526_EG1_E6_vat_rate), gCurrency, ggExchRateNo, "X" , "X")		'vat�� 
		iArrCols(14) = UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(LngRow,S526_EG1_E6_vat_amt),lgCurrency,ggAmtOfMoneyNo, gTaxRndPolicyNo  , "X")			'�ΰ����� 
		iArrCols(15) = UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(LngRow,S526_EG1_E6_bill_amt_loc),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo  , "X")		'��ȭ�ݾ� 
		iArrCols(16) = UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(LngRow,S526_EG1_E6_vat_amt_loc),gCurrency,ggAmtOfMoneyNo, gTaxRndPolicyNo  , "X")		'�ΰ�����ȭ�ݾ� 
		iArrCols(17) = UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(LngRow,S526_EG1_E6_deposit_price ),lgCurrency,ggUnitCostNo, "X" , "X")					'�����ݴܰ� 
		iArrCols(18) = UNIConvNumDBToCompanyByCurrency(EG1_exp_grp(LngRow,S526_EG1_E6_deposit_amt),lgCurrency,ggAmtOfMoneyNo, "X" , "X")					'�����ݾ� 
		iArrCols(20) = ConvSPChars(EG1_exp_grp(LngRow,S526_EG1_E6_remark))					'��� 
		iArrCols(21) = ConvSPChars(EG1_exp_grp(LngRow,S526_EG1_E8_dn_no))					'���Ϲ�ȣ 
		iArrCols(22) = UNINumClientFormat(EG1_exp_grp(LngRow,S526_EG1_E7_dn_seq), 0, 0)		'���ϼ��� 
		iArrCols(23) = ConvSPChars(EG1_exp_grp(LngRow,S526_EG1_E12_so_no))					'���ֹ�ȣ 
		iArrCols(24) = UNINumClientFormat(EG1_exp_grp(LngRow,S526_EG1_E11_so_seq), 0, 0)	'���ּ��� 
		iArrCols(25) = ConvSPChars(EG1_exp_grp(LngRow,S526_EG1_E5_lc_no))					'L/C��ȣ 
		iArrCols(26) = UNINumClientFormat(EG1_exp_grp(LngRow,S526_EG1_E4_lc_seq), 0, 0)		'L/C���� 
		iArrCols(27) = UNINumClientFormat(EG1_exp_grp(LngRow,S526_EG1_E6_bill_seq), 0, 0)	'������� 
		iArrCols(28) = ConvSPChars(EG1_exp_grp(LngRow,S526_EG1_E9_plant_cd))				'���� 
		iArrCols(29) = ConvSPChars(EG1_exp_grp(LngRow,S526_EG1_E10_spec))					'ǰ��԰� 
		iArrCols(30) = ConvSPChars(EG1_exp_grp(LngRow,S526_EG1_E6_ret_item_flag))			'��ǰ���� 
		iArrCols(31) = iArrCols(9)															'�������ݾ� 
		iArrCols(32) = iArrCols(6)															'������ VAT ���Կ��� 
		iArrCols(33) = iArrCols(14)															'�������ΰ����� 
		iArrCols(34) = iArrCols(9)															' �������ݾ� 
		iArrCols(35) = iArrCols(14)															' �������ΰ����� 
		iArrCols(36) = LngMaxRow + LngRow 
		
   		iArrRows(LngRow) = Join(iArrCols, gColSep)			
	Next
	
	Response.Write "<Script language=vbs> " & vbCr   
	Response.Write "With parent " & vbCr   
    Response.Write " .frm1.vspdData.Redraw = False  " & vbCr      
    Response.Write " .ggoSpread.Source = .frm1.vspdData	" & vbCr
    Response.Write " .ggoSpread.SSShowDataByClip """ & Join(iArrRows, gColSep & gRowSep) & gColSep & gRowSep & """ ,""F""" & vbCr
    
    Response.Write " .lgStrPrevKey = """ & iStrNextKey  & """" & vbCr  
    Response.Write " .DbQueryOk " & vbCr   
    
    Response.Write " If .frm1.HPostFlag.Value = """ & "Y" & """ Then " & vbCr	         
	Response.Write " .SetPostYesSpreadColor(" & LngMaxRow & ")  " & vbCr	
	Response.Write " Else " & vbCr
	Response.Write " .SetQuerySpreadColor(" & LngMaxRow & ") " & vbCr
	Response.Write " End If	" & vbCr
	Response.Write " .frm1.vspdData.Redraw = True  " & vbCr
	Response.Write "End With " & vbCr   
	Response.Write "</Script> "																				         & vbCr          

Case CStr(UID_M0002)																'��: ���� ��û�� ���� 
									
    Err.Clear																		'��: Protect system from crashing
    Dim iErrorPosition
    Dim itxtSpread,itxtHBillNo
    
    Set iPS7G121 = Server.CreateObject("PS7G121.cSBillDtlSvr")  
    
    If CheckSYSTEMError(Err,True) = True Then
		Response.End		
    End If
	
	itxtSpread = Trim(Request("txtSpread"))
	itxtHBillNo = Trim(Request("txtHBillNo"))
	
	'for ����Tax
    '@@@@@@@@@@@
     pvCB = "F"
    '@@@@@@@@@@@
	Call iPS7G121.S_MAINT_BILL_DTL_SVR (pvCB, gStrGlobalCollection ,itxtSpread , iErrorPosition, _
		                                itxtHBillNo )
	
	If CheckSYSTEMError2(Err, True, iErrorPosition & "��","","","","") = True Then
       Set iPS7G121 = Nothing
	   Response.End 
	End If
   
    Set iPS7G121 = Nothing	
	
	Response.Write "<Script language=vbs> " & vbCr         
    Response.Write " Parent.DBSaveOk "      & vbCr   
    Response.Write "</Script> "    
	    
Case CStr("PostFlag")																'��: Ȯ�� ��û 

	Err.Clear														'��: Protect system from crashing

	Dim itxtHBillNoPost
	
    Set iPS7G115 = Server.CreateObject("PS7G115.cSPostOpenArSvr")  
    
    If CheckSYSTEMError(Err,True) = True Then
		Response.End		
    End If
	
	itxtHBillNoPost = Trim(Request("txtHBillNo"))
	
	'for ����Tax
    '@@@@@@@@@@@
     pvCB = "F"
    '@@@@@@@@@@@
	Call iPS7G115.S_POST_OPEN_AR_SVR(pvCB, gStrGlobalCollection,itxtHBillNoPost)
    
	If CheckSYSTEMError(Err,True) = True Then
		Set iPS7G115 = Nothing
		Response.End		
    End If
	
	Set iPS7G115 = Nothing	
	
	Response.Write "<Script language=vbs> " & vbCr         
    Response.Write " Parent.DBSaveOk "      & vbCr   
    Response.Write "</Script> "    
    
End Select

'==============================================================================
' ����� ���� ���� �Լ� 
'==============================================================================
%>

<SCRIPT LANGUAGE=VBSCRIPT RUNAT=SERVER>
'============================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================
Sub SubMakeSQLStatements()
	Dim strSelectList, strFromList
	
	strSelectList = "SELECT * "
	strFromList = " FROM dbo.ufn_s_GetBillHdrInfo( " & FilterVar(Request("txtConBillNo"), "''", "S") & " , " & FilterVar("N", "''", "S") & " , " & FilterVar("N", "''", "S") & " , " & FilterVar("Y", "''", "S") & " , " & FilterVar("Y", "''", "S") & " , " & FilterVar("Q", "''", "S") & " , " & FilterVar("N", "''", "S") & " )"
	lgstrsql = strSelectList & strFromList
	'call svrmsgbox(lgstrsql, vbinformation, i_mkscript)
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub
</SCRIPT>
