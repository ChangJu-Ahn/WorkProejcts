
<%
'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : Capital Expense
'*  3. Program ID           : a7126mb1(자본적지출내역등록)
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
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                          
'**********************************************************************************************
Response.Expires = -1								'☜ : ASP가 캐쉬되지 않도록 한다.
Response.Buffer = True								'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.

%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->

<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

	On Error Resume Next														'☜: 
    Call LoadBasisGlobalInf()
	Dim  lgBlnFlgChgValue
	Call LoadInfTB19029B("I","*", "NOCOOKIE", "MB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")

	Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
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

    Call HideStatusWnd                                                               '☜: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read MAX (CRUD)
    lgChangeOrgId	  = Request("hOrgChangeId")
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
			call SubBizQuery() 
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSave()
        Case CStr(UID_M0003)                                                         '☜: Delete
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
    
    '미지급금계정추가 
    Const EG1_ar_ap_acct_cd			= 20
    Const EG1_ar_ap_acct_popup		= 21
    Const EG1_ar_ap_acct_nm			= 22
    
    Const EG1_ar_ap_due_dt			= 23
    
    Redim I1_a_asset_chg(I1_org_change_id)
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
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
	Response.Write "		.txtAsstNo.value	= """ & ConvSPChars(E1_a_asset_master(E1_asst_no)) &			"""" & vbCr		'변동번호 
	Response.Write "		.txtAsstNm.value	= """ & ConvSPChars(E1_a_asset_master(E1_asst_nm)) &			"""" & vbCr		'자산명 
	Response.Write "		.txtCapExpNo.value	= """ & ConvSPChars(E2_a_asset_chg(E2_cap_exp_no)) &			"""" & vbCr		'자산번호 

	Response.Write "		.txtAsstNo1.value	= """ & ConvSPChars(E1_a_asset_master(E1_asst_no)) &			"""" & vbCr		'변동번호 
	Response.Write "		.txtAsstNm1.value	= """ & ConvSPChars(E1_a_asset_master(E1_asst_nm)) &			"""" & vbCr		'자산명 
	Response.Write "		.txtAcctDeptNm.value	= """ & ConvSPChars(E1_a_asset_master(E1_dept_nm)) &		"""" & vbCr		'자산관리부서 
	Response.Write "		.fpDateTime1.text	= """ & UNIDateClientFormat(E1_a_asset_master(E1_reg_dt)) &		"""" & vbCr		'자산취득일자 
	Response.Write "		.txtAcqQty.text     = """ & E1_a_asset_master(E1_acq_qty) &							"""" & vbCr		'취득수량 
	Response.Write "		.txtInvQty.text     = """ & E1_a_asset_master(E1_inv_qty) &							"""" & vbCr		'재고수량 
	
	Response.Write "		.txtCapExpNo1.value	= """ & ConvSPChars(E2_a_asset_chg(E2_cap_exp_no)) &			"""" & vbCr		'자산번호 
	Response.Write "		.txtDeptCd.Value	= """ & ConvSPChars(E2_a_asset_chg(E2_from_dept_cd)) &			"""" & vbCr		'회계부서        
	Response.Write "		.txtDeptNm.Value	= """ & ConvSPChars(E2_a_asset_chg(E2_from_dept_nm)) &			"""" & vbCr		'회계부서명	    	    	    
	Response.Write "		.fpDateTime2.text	= """ & UNIDateClientFormat(E2_a_asset_chg(E2_chg_dt)) &		"""" & vbCr		'자본적지출일자 
	
	Response.Write "		.txtDocCur.value		= """ & ConvSPChars(E2_a_asset_chg(E2_doc_cur)) &				"""" & vbCr		'거래통화 
	if gIsShowLocal <> "N" then
		Response.Write "	.txtXchRate.text		= """ & E2_a_asset_chg(E2_xch_rate) &							"""" & vbCr		'거래환율 
	else
		Response.Write "	.txtXchRate.value		= """ & E2_a_asset_chg(E2_xch_rate) &							"""" & vbCr		'거래환율 
	end if
	
	Response.Write "		.txtTotalAmt.text	= """ & E2_a_asset_chg(E2_chg_tot_amt) &						"""" & vbCr		'총취득금액 
	if gIsShowLocal <> "N" then
		Response.Write "	.txtTotalLocAmt.text= """ & E2_a_asset_chg(E2_chg_tot_loc_amt) &					"""" & vbCr		'총취득금액(자국)
	else
		Response.Write "	.txtTotalLocAmt.text= """ & E2_a_asset_chg(E2_chg_tot_loc_amt) &					"""" & vbCr		'총취득금액(자국)
	end if

 	Response.Write "		.txtTempGLNo.Value	= """ & ConvSPChars(E2_a_asset_chg(E2_temp_gl_no)) &			"""" & vbCr		'TempGL No
	Response.Write "		.txtGLNo.Value		= """ & ConvSPChars(E2_a_asset_chg(E2_gl_no)) &					"""" & vbCr		'GL No
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	Response.Write "		.hOrgChangeId.Value	= """ & ConvSPChars(E2_a_asset_chg(E2_from_org_change_id)) &					"""" & vbCr		'GL No
    Response.Write "	End With	" & vbCr  

	iStrData = ""
	iIntLoopCount = 0	
	For iLngRow = 0 To UBound(EG1_a_asset_chg, 1) 		    
		
				iStrData = iStrData & Chr(11)  & ConvSPChars(EG1_a_asset_chg(iLngRow,EG1_chg_no))				'지출내역순번 
				iStrData = iStrData & Chr(11)  & ConvSPChars(EG1_a_asset_chg(iLngRow,EG1_asset_chg_desc))					'지출내역 
				iStrData = iStrData & Chr(11)  & ConvSPChars(EG1_a_asset_chg(iLngRow,EG1_bp_cd))				'거래처코드 
				iStrData = iStrData & Chr(11)  & ""																'거래처팝업 
				iStrData = iStrData & Chr(11)  & ConvSPChars(EG1_a_asset_chg(iLngRow,EG1_bp_nm))				'거래처명 
				iStrData = iStrData & Chr(11)  & EG1_a_asset_chg(iLngRow,EG1_chg_amt)							'지출금액 
				iStrData = iStrData & Chr(11)  & EG1_a_asset_chg(iLngRow,EG1_chg_loc_amt)						'지출금액(자국)
				iStrData = iStrData & Chr(11)  & ConvSPChars(EG1_a_asset_chg(iLngRow,EG1_tax_type_cd))			'부가세유형 
				iStrData = iStrData & Chr(11)  & ""																'부가세유형팝업 
				iStrData = iStrData & Chr(11)  & ConvSPChars(EG1_a_asset_chg(iLngRow,EG1_tax_type_nm))			'부가세유형명 
				iStrData = iStrData & Chr(11)  & EG1_a_asset_chg(iLngRow,EG1_tax_rate)							'부가세율 
				iStrData = iStrData & Chr(11)  & EG1_a_asset_chg(iLngRow,EG1_tax_amt)							'부가세금액 
				iStrData = iStrData & Chr(11)  & EG1_a_asset_chg(iLngRow,EG1_tax_loc_amt)						'부가세금액(자국)
				iStrData = iStrData & Chr(11)  & ConvSPChars(EG1_a_asset_chg(iLngRow,EG1_tax_biz_area_cd))		'신고사업장 
				iStrData = iStrData & Chr(11)  & ""																'신고사업장팝업 
				iStrData = iStrData & Chr(11)  & ConvSPChars(EG1_a_asset_chg(iLngRow,EG1_tax_biz_area_nm))		'신고사업장명 
				iStrData = iStrData & Chr(11)  & UNIDateClientFormat(EG1_a_asset_chg(iLngRow,EG1_Issued_dt))	'계산서발행일 
				iStrData = iStrData & Chr(11)  & ConvSPChars(EG1_a_asset_chg(iLngRow,EG1_paym_type_cd))			'신고사업장명 
				iStrData = iStrData & Chr(11)  & ConvSPChars(EG1_a_asset_chg(iLngRow,EG1_paym_type_nm))			'신고사업장명 
				iStrData = iStrData & Chr(11)  & ConvSPChars(EG1_a_asset_chg(iLngRow,EG1_paym_type_desc))		'신고사업장명 
				iStrData = iStrData & Chr(11)  & ""
				
				'미지급금코드추가 
				iStrData = iStrData & Chr(11)  & ConvSPChars(EG1_a_asset_chg(iLngRow,EG1_ar_ap_acct_cd))		'미지급금계정코드 
				iStrData = iStrData & Chr(11)  & ""																'미지급금계정팝업 
				iStrData = iStrData & Chr(11)  & ConvSPChars(EG1_a_asset_chg(iLngRow,EG1_ar_ap_acct_nm))		'미지급금계정명 
				
				iStrData = iStrData & Chr(11)  & UNIDateClientFormat(EG1_a_asset_chg(iLngRow,EG1_ar_ap_due_dt))	'계산서발행일 
				iStrData = iStrData & Chr(11)  & Chr(12)
	Next

	Response.Write "	With parent	" & vbCr

    Response.Write "	.ggoSpread.Source = .frm1.vspdData  " & vbCr 			 
    Response.Write "	.ggoSpread.SSShowData """ & iStrData  & """" & vbCr
	Response.Write "	.DBQueryOk()	" & vbCr

	Response.Write "	End With	" & vbCr

	Response.Write "</Script>		" & vbCr  
	 
    Set iPAAG075 = Nothing															    '☜: Unload 

	Response.End																		'☜: Process End


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
	
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
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
	lgIntFlgMode = CInt(Request("txtFlgMode"))									'☜: 저장시 Create/Update 판별 

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
    
   Set iPAAG075 = Nothing                                                  '☜: Unload Comproxy

	Response.Write "<Script Language=vbscript>				 " & vbCr
	Response.Write "With parent						" & vbCr
	Response.Write "  .frm1.txtAsstNo.Value=  """ & ConvSPChars(E1_a_asset_chg(E1_asst_cd)) & 				"""" & vbCr
	Response.Write "  .frm1.txtCapExpNo.Value=  """ & ConvSPChars(E1_a_asset_chg(E1_cap_exp_no)) & 				"""" & vbCr
	Response.Write "	.DbSaveOk " & vbCr  	' 이전 키 값 넘겨줌 , 현재 ComProxy가 제대로 안되어 있음 															'☜: 조화가 성공 
	Response.Write "	End With		" & vbCr  
	Response.Write "</Script>		" & vbCr  

	Response.End    													   '☜: Process End  	  

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
    Err.Clear                                                                        '☜: Clear Error status
    On Error Resume Next                                                             '☜: Protect system from crashing

    If Request("txtAsstNo") = "" or Request("txtCapExpNo") = "" Then    	'⊙: 삭제를 위한 값이 들어왔는지 체크 
		Call ServerMesgBox("700114", vbInformation, I_MKSCRIPT)			'삭제 조건값이 비어있습니다!
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
    
    Set iPAAG075 = Nothing                                                   '☜: Unload Comproxy
    
	Response.Write "<Script Language=vbscript>				 " & vbCr
	Response.Write "	Call parent.DbDeleteOk()		" & vbCr
	Response.Write "</Script>		" & vbCr 
	
	Response.End    													   '☜: Process End  	  
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub

%>

