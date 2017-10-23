
<%
'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : Capital Expense
'*  3. Program ID           : a7126mb1(�ں������⳻�����)
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : +B19029LookupNumericFormat
'                             +B25011ManagePlant
'                             +B25011ManagePlant
'                             +B25018ListPlant
'                             +B25019LookUpPlant
'*  7. Modified date(First) : 2002/11/14
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Seo Hyo Seok
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                          
'**********************************************************************************************
Response.Expires = -1								'�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
Response.Buffer = True								'�� : ASP�� ���ۿ� ������� �ʰ� �ٷ� Client�� ��������.

%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->

<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

	On Error Resume Next														'��: 
    Call LoadBasisGlobalInf()
	Dim  lgBlnFlgChgValue
	Call LoadInfTB19029B("I","*", "NOCOOKIE", "MB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")

	Dim strMode																	'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
	DIM strApDueDt
	Dim amt1 
	Dim amt2 
	Dim amt3 
	Dim amt4 
	DIm strChgFg
	Dim LngMaxRow
	Dim lgCurrency
	Dim lgCurrencyAcq
	Dim lgChangeOrgId

    Call HideStatusWnd                                                               '��: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '��: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '��: Read Operation Mode (CRUD)
    lgLngMaxRow       = Request("txtMaxRows")                                        '��: Read MAX (CRUD)
    lgChangeOrgId	  = Request("hOrgChangeId")
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '��: Query
			call SubBizQuery() 
        Case CStr(UID_M0002)                                                         '��: Save,Update
             Call SubBizSave()
        Case CStr(UID_M0003)                                                         '��: Delete
             Call SubBizDelete()
    End Select
    
    Response.End       
	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
	
    Dim iPAAG075

    Dim I1_a_asset_chg
    Dim E1_a_asset_master
    Dim E2_a_asset_chg
 	Dim EG1_a_asset_chg
	
	Const I1_asst_cd = 0
	Const I1_cap_exp_no = 1
	Const I1_org_change_id = 2

	Const E1_asst_no		= 0
	Const E1_asst_nm		= 1
	Const E1_dept_nm		= 2
	Const E1_reg_dt			= 3
	Const E1_acq_qty		= 4
	Const E1_inv_qty		= 5
	
	Const E2_cap_exp_no		= 0
	Const E2_from_dept_cd	= 1
	Const E2_from_dept_nm	= 2
	Const E2_from_org_change_id  = 3
	Const E2_chg_dt			= 4
	Const E2_doc_cur		= 5
	Const E2_xch_rate		= 6
	Const E2_chg_tot_amt	= 7
	Const E2_chg_tot_loc_amt	= 8
	Const E2_gl_no			= 9
	Const E2_temp_gl_no		= 10
	
	Const EG1_chg_no				= 0
	Const EG1_asset_chg_desc		= 1
	Const EG1_bp_cd					= 2
	Const EG1_bp_Popup				= 3
	Const EG1_bp_nm					= 4
	Const EG1_chg_amt				= 5
	Const EG1_chg_loc_amt			= 6
	Const EG1_tax_type_cd			= 7
	Const EG1_tax_type_popup		= 8
	Const EG1_tax_type_nm			= 9
    Const EG1_tax_rate				= 10
    Const EG1_tax_amt				= 11
    Const EG1_tax_loc_amt			= 12
    Const EG1_tax_biz_area_cd		= 13
    Const EG1_tax_biz_area_popup	= 14
    Const EG1_tax_biz_area_nm		= 15
    Const EG1_Issued_dt				= 16
    Const EG1_paym_type_cd			= 17
    Const EG1_paym_type_nm			= 18
    Const EG1_paym_type_desc		= 19
    
    '�����ޱݰ����߰� 
    Const EG1_ar_ap_acct_cd			= 20
    Const EG1_ar_ap_acct_popup		= 21
    Const EG1_ar_ap_acct_nm			= 22
    
    Const EG1_ar_ap_due_dt			= 23
    
    Redim I1_a_asset_chg(I1_org_change_id)
    
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
  '********************************************************  
  '                        Query
  '********************************************************  

    Set iPAAG075 = Server.CreateObject("PAAG075.cAListAsChg01Svr")    
     
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
 	If CheckSYSTEMError(Err,True) = True Then
 		Set iPAAG075 =nothing	
       Exit Sub
    End If 
 
   
    '-----------------------
    'Data manipulate  area(import view match)
    '-----------------------
    I1_a_asset_chg(I1_asst_cd) = Trim(Request("txtAsstNo"))
    I1_a_asset_chg(I1_cap_exp_no) = Trim(Request("txtCapExpNo"))
	I1_a_asset_chg(I1_org_change_id) = lgChangeOrgId
	
    '-----------------------
    'Com action area
    '-----------------------
    Call iPAAG075.A_LIST_CAPITAL_EXPENSE_SVR(gStrGlobalCollection , I1_a_asset_chg , _
											 E1_a_asset_master, E2_a_asset_chg , EG1_a_asset_chg )
											 
	'----------------------------------------------
	'Com action result check area
	'----------------------------------------------
 	If CheckSYSTEMError(Err,True) = True Then
 		Set iPAAG075 =nothing
       Exit Sub
    End If
    
	'-----------------------
	'Result data display area
	'-----------------------
	Response.Write "<Script Language=vbscript>				 " & vbCr
	Response.Write "	With parent.frm1	" & vbCr

	''''''''''''''''''''''''''''''''
	'  The Part for Asset master
	''''''''''''''''''''''''''''''''
	Response.Write "		.txtAsstNo.value	= """ & ConvSPChars(E1_a_asset_master(E1_asst_no)) &			"""" & vbCr		'������ȣ 
	Response.Write "		.txtAsstNm.value	= """ & ConvSPChars(E1_a_asset_master(E1_asst_nm)) &			"""" & vbCr		'�ڻ�� 
	Response.Write "		.txtCapExpNo.value	= """ & ConvSPChars(E2_a_asset_chg(E2_cap_exp_no)) &			"""" & vbCr		'�ڻ��ȣ 

	Response.Write "		.txtAsstNo1.value	= """ & ConvSPChars(E1_a_asset_master(E1_asst_no)) &			"""" & vbCr		'������ȣ 
	Response.Write "		.txtAsstNm1.value	= """ & ConvSPChars(E1_a_asset_master(E1_asst_nm)) &			"""" & vbCr		'�ڻ�� 
	Response.Write "		.txtAcctDeptNm.value	= """ & ConvSPChars(E1_a_asset_master(E1_dept_nm)) &		"""" & vbCr		'�ڻ�����μ� 
	Response.Write "		.fpDateTime1.text	= """ & UNIDateClientFormat(E1_a_asset_master(E1_reg_dt)) &		"""" & vbCr		'�ڻ�������� 
	Response.Write "		.txtAcqQty.text     = """ & E1_a_asset_master(E1_acq_qty) &							"""" & vbCr		'������ 
	Response.Write "		.txtInvQty.text     = """ & E1_a_asset_master(E1_inv_qty) &							"""" & vbCr		'������ 
	
	Response.Write "		.txtCapExpNo1.value	= """ & ConvSPChars(E2_a_asset_chg(E2_cap_exp_no)) &			"""" & vbCr		'�ڻ��ȣ 
	Response.Write "		.txtDeptCd.Value	= """ & ConvSPChars(E2_a_asset_chg(E2_from_dept_cd)) &			"""" & vbCr		'ȸ��μ�        
	Response.Write "		.txtDeptNm.Value	= """ & ConvSPChars(E2_a_asset_chg(E2_from_dept_nm)) &			"""" & vbCr		'ȸ��μ���	    	    	    
	Response.Write "		.fpDateTime2.text	= """ & UNIDateClientFormat(E2_a_asset_chg(E2_chg_dt)) &		"""" & vbCr		'�ں����������� 
	
	Response.Write "		.txtDocCur.value		= """ & ConvSPChars(E2_a_asset_chg(E2_doc_cur)) &				"""" & vbCr		'�ŷ���ȭ 
	if gIsShowLocal <> "N" then
		Response.Write "	.txtXchRate.text		= """ & E2_a_asset_chg(E2_xch_rate) &							"""" & vbCr		'�ŷ�ȯ�� 
	else
		Response.Write "	.txtXchRate.value		= """ & E2_a_asset_chg(E2_xch_rate) &							"""" & vbCr		'�ŷ�ȯ�� 
	end if
	
	Response.Write "		.txtTotalAmt.text	= """ & E2_a_asset_chg(E2_chg_tot_amt) &						"""" & vbCr		'�����ݾ� 
	if gIsShowLocal <> "N" then
		Response.Write "	.txtTotalLocAmt.text= """ & E2_a_asset_chg(E2_chg_tot_loc_amt) &					"""" & vbCr		'�����ݾ�(�ڱ�)
	else
		Response.Write "	.txtTotalLocAmt.text= """ & E2_a_asset_chg(E2_chg_tot_loc_amt) &					"""" & vbCr		'�����ݾ�(�ڱ�)
	end if

 	Response.Write "		.txtTempGLNo.Value	= """ & ConvSPChars(E2_a_asset_chg(E2_temp_gl_no)) &			"""" & vbCr		'TempGL No
	Response.Write "		.txtGLNo.Value		= """ & ConvSPChars(E2_a_asset_chg(E2_gl_no)) &					"""" & vbCr		'GL No
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	Response.Write "		.hOrgChangeId.Value	= """ & ConvSPChars(E2_a_asset_chg(E2_from_org_change_id)) &					"""" & vbCr		'GL No
    Response.Write "	End With	" & vbCr  

	iStrData = ""
	iIntLoopCount = 0	
	For iLngRow = 0 To UBound(EG1_a_asset_chg, 1) 		    
		
				iStrData = iStrData & Chr(11)  & ConvSPChars(EG1_a_asset_chg(iLngRow,EG1_chg_no))				'���⳻������ 
				iStrData = iStrData & Chr(11)  & ConvSPChars(EG1_a_asset_chg(iLngRow,EG1_asset_chg_desc))					'���⳻�� 
				iStrData = iStrData & Chr(11)  & ConvSPChars(EG1_a_asset_chg(iLngRow,EG1_bp_cd))				'�ŷ�ó�ڵ� 
				iStrData = iStrData & Chr(11)  & ""																'�ŷ�ó�˾� 
				iStrData = iStrData & Chr(11)  & ConvSPChars(EG1_a_asset_chg(iLngRow,EG1_bp_nm))				'�ŷ�ó�� 
				iStrData = iStrData & Chr(11)  & EG1_a_asset_chg(iLngRow,EG1_chg_amt)							'����ݾ� 
				iStrData = iStrData & Chr(11)  & EG1_a_asset_chg(iLngRow,EG1_chg_loc_amt)						'����ݾ�(�ڱ�)
				iStrData = iStrData & Chr(11)  & ConvSPChars(EG1_a_asset_chg(iLngRow,EG1_tax_type_cd))			'�ΰ������� 
				iStrData = iStrData & Chr(11)  & ""																'�ΰ��������˾� 
				iStrData = iStrData & Chr(11)  & ConvSPChars(EG1_a_asset_chg(iLngRow,EG1_tax_type_nm))			'�ΰ��������� 
				iStrData = iStrData & Chr(11)  & EG1_a_asset_chg(iLngRow,EG1_tax_rate)							'�ΰ����� 
				iStrData = iStrData & Chr(11)  & EG1_a_asset_chg(iLngRow,EG1_tax_amt)							'�ΰ����ݾ� 
				iStrData = iStrData & Chr(11)  & EG1_a_asset_chg(iLngRow,EG1_tax_loc_amt)						'�ΰ����ݾ�(�ڱ�)
				iStrData = iStrData & Chr(11)  & ConvSPChars(EG1_a_asset_chg(iLngRow,EG1_tax_biz_area_cd))		'�Ű����� 
				iStrData = iStrData & Chr(11)  & ""																'�Ű������˾� 
				iStrData = iStrData & Chr(11)  & ConvSPChars(EG1_a_asset_chg(iLngRow,EG1_tax_biz_area_nm))		'�Ű������ 
				iStrData = iStrData & Chr(11)  & UNIDateClientFormat(EG1_a_asset_chg(iLngRow,EG1_Issued_dt))	'��꼭������ 
				iStrData = iStrData & Chr(11)  & ConvSPChars(EG1_a_asset_chg(iLngRow,EG1_paym_type_cd))			'�Ű������ 
				iStrData = iStrData & Chr(11)  & ConvSPChars(EG1_a_asset_chg(iLngRow,EG1_paym_type_nm))			'�Ű������ 
				iStrData = iStrData & Chr(11)  & ConvSPChars(EG1_a_asset_chg(iLngRow,EG1_paym_type_desc))		'�Ű������ 
				iStrData = iStrData & Chr(11)  & ""
				
				'�����ޱ��ڵ��߰� 
				iStrData = iStrData & Chr(11)  & ConvSPChars(EG1_a_asset_chg(iLngRow,EG1_ar_ap_acct_cd))		'�����ޱݰ����ڵ� 
				iStrData = iStrData & Chr(11)  & ""																'�����ޱݰ����˾� 
				iStrData = iStrData & Chr(11)  & ConvSPChars(EG1_a_asset_chg(iLngRow,EG1_ar_ap_acct_nm))		'�����ޱݰ����� 
				
				iStrData = iStrData & Chr(11)  & UNIDateClientFormat(EG1_a_asset_chg(iLngRow,EG1_ar_ap_due_dt))	'��꼭������ 
				iStrData = iStrData & Chr(11)  & Chr(12)
	Next

	Response.Write "	With parent	" & vbCr

    Response.Write "	.ggoSpread.Source = .frm1.vspdData  " & vbCr 			 
    Response.Write "	.ggoSpread.SSShowData """ & iStrData  & """" & vbCr
	Response.Write "	.DBQueryOk()	" & vbCr

	Response.Write "	End With	" & vbCr

	Response.Write "</Script>		" & vbCr  
	 
    Set iPAAG075 = Nothing															    '��: Unload 

	Response.End																		'��: Process End


end sub
'============================================================================================================
' Name : SubBizSave
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()

	Dim iPAAG075
	
	Dim I1_a_asset_chg 
	Dim IG1_asset_chg
    Dim E1_a_asset_chg
	
	Const iCommandSent		= 0
	Const I1_asst_cd				= 1
	Const I1_cap_exp_no		= 2
	Const I1_from_dept_cd		= 3
	Const I1_org_change_id	= 4
	Const I1_chg_dt				= 5
	Const I1_loc_cur				= 6
	Const I1_doc_cur				= 7
	Const I1_xch_rate			= 8
	Const I1_user_id				= 9
	Const I1_gl_no					= 10
	Const I1_temp_gl_no		= 11
	
	Const E1_asst_cd			= 0
	Const E1_cap_exp_no	= 1
	
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
    
	redim I1_a_asset_chg(I1_temp_gl_no)

	'***************************************************************
	'                              SAVE
	'***************************************************************									
	
    Set iPAAG075 = Server.CreateObject("PAAG075.cAMngAsChg01Svr") 
    
	If CheckSYSTEMError(Err,True) = True Then
 		Set iPAAG075 =nothing	
       Exit Sub
    End If 

    '-----------------------
    'Data manipulate area
    '-----------------------
	lgIntFlgMode = CInt(Request("txtFlgMode"))									'��: ����� Create/Update �Ǻ� 

    If lgIntFlgMode = OPMD_CMODE Then
		I1_a_asset_chg(iCommandSent) = "CREATE"
    ElseIf lgIntFlgMode = OPMD_UMODE Then
		I1_a_asset_chg(iCommandSent) = "UPDATE"
    End If

	I1_a_asset_chg(I1_asst_cd) = Trim(UCase(Request("txtAsstNo1")))
	I1_a_asset_chg(I1_cap_exp_no) = Trim(UCase(Request("txtCapExpNo1")))
	I1_a_asset_chg(I1_from_dept_cd) = Trim(UCase(Request("txtDeptCd")))
	I1_a_asset_chg(I1_org_change_id) = lgChangeOrgId
	I1_a_asset_chg(I1_chg_dt) = UNIConvDate(Request("txtChgDt"))
	I1_a_asset_chg(I1_loc_cur) = gCurrency
	I1_a_asset_chg(I1_doc_cur) = Trim(UCase(Request("txtDocCur")))
	I1_a_asset_chg(I1_xch_rate) = UNIConvNum(Request("txtXchRate"),0) 
	I1_a_asset_chg(I1_user_id) = gUsrID
	I1_a_asset_chg(I1_gl_no) = Trim(UCase(Request("txtGlNo")))    
	I1_a_asset_chg(I1_temp_gl_no) = Trim(UCase(Request("txtTempGlNo")))             

	'-----------------------
	'Com Action Area
	'-----------------------
	IF ERR.number <> 0 THEN
		Response.Write "XXX" & ERR.CODE & " :: " & ERR.Description 
		Response.End
	END IF

	Call iPAAG075.A_MAN_CAPITAL_EXPENSE_SVR(gStrGlobalCollection, I1_a_asset_chg, Request("txtSpread"), E1_a_asset_chg)

 	If CheckSYSTEMError(Err,True) = True Then
 		Set iPAAG075 =nothing	
		Exit Sub
    End If 
    
   Set iPAAG075 = Nothing                                                  '��: Unload Comproxy

	Response.Write "<Script Language=vbscript>				 " & vbCr
	Response.Write "With parent						" & vbCr
	Response.Write "  .frm1.txtAsstNo.Value=  """ & ConvSPChars(E1_a_asset_chg(E1_asst_cd)) & 				"""" & vbCr
	Response.Write "  .frm1.txtCapExpNo.Value=  """ & ConvSPChars(E1_a_asset_chg(E1_cap_exp_no)) & 				"""" & vbCr
	Response.Write "	.DbSaveOk " & vbCr  	' ���� Ű �� �Ѱ��� , ���� ComProxy�� ����� �ȵǾ� ���� 															'��: ��ȭ�� ���� 
	Response.Write "	End With		" & vbCr  
	Response.Write "</Script>		" & vbCr  

	Response.End    													   '��: Process End  	  

End Sub	
	    
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
	Dim iPAAG075
	
	Dim I1_a_asset_chg 
	Dim IG1_asset_chg
    Dim E1_a_asset_chg
	
	Const iCommandSent		= 0
	Const I1_asst_cd				= 1
	Const I1_cap_exp_no		= 2
	Const I1_from_dept_cd		= 3
	Const I1_org_change_id	= 4
	Const I1_chg_dt				= 5
	Const I1_loc_cur				= 6
	Const I1_doc_cur				= 7
	Const I1_xch_rate			= 8
	Const I1_user_id				= 9
	Const I1_gl_no					= 10
	Const I1_temp_gl_no		= 11
	
	Const E1_asst_cd		= 0
	Const E1_cap_exp_no		= 1
	
	redim I1_a_asset_chg(I1_temp_gl_no)

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
	'***************************************************************
	'                              DELETE
	'***************************************************************
    Err.Clear                                                                        '��: Clear Error status
    On Error Resume Next                                                             '��: Protect system from crashing

    If Request("txtAsstNo") = "" or Request("txtCapExpNo") = "" Then    	'��: ������ ���� ���� ���Դ��� üũ 
		Call ServerMesgBox("700114", vbInformation, I_MKSCRIPT)			'���� ���ǰ��� ����ֽ��ϴ�!
		Response.End 
	End If

    Set iPAAG075 = Server.CreateObject("PAAG075.cAMngAsChg01Svr") 
    
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
 	If CheckSYSTEMError(Err,True) = True Then
 		Set iPAAG075 =nothing	
       Exit Sub
    End If 
    
    '-----------------------
    'Data manipulate  area(import view match)
    '-----------------------
    I1_a_asset_chg(ICommandSent) = "DELETE"
	I1_a_asset_chg(I1_asst_cd) = Trim(UCase(Request("txtAsstNo")))
	I1_a_asset_chg(I1_cap_exp_no) = Trim(UCase(Request("txtCapExpNo")))
	I1_a_asset_chg(I1_gl_no) = Trim(UCase(Request("txtGlNo")))
	I1_a_asset_chg(I1_temp_gl_no) = Trim(UCase(Request("txtTempGlNo")))
            
	Call iPAAG075.A_MAN_CAPITAL_EXPENSE_SVR(gStrGlobalCollection, I1_a_asset_chg, , E1_a_asset_chg)
    
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    
 	If CheckSYSTEMError(Err,True) = True Then
 		Set iPAAG075 =nothing	
       Exit Sub
    End If 
    
    Set iPAAG075 = Nothing                                                   '��: Unload Comproxy
    
	Response.Write "<Script Language=vbscript>				 " & vbCr
	Response.Write "	Call parent.DbDeleteOk()		" & vbCr
	Response.Write "</Script>		" & vbCr 
	
	Response.End    													   '��: Process End  	  
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub

%>

