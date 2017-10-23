<%@ LANGUAGE=VBSCript%>
<%Option Explicit%>
<%
Response.Expires = -1								'☜ : ASP가 캐쉬되지 않도록 한다.
Response.Buffer = True								'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.
%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<!-- #Include file="../ag/incAcctMBFunc.asp"  -->
<%
	Dim strYYYYMM
	Dim strVerCd
	Dim iStrData

    On Error Resume Next																'☜: Protect system from crashing
    Err.Clear  

	' 권한관리 추가 
	Dim lgBizAreaAuthYn, lgAuthBizAreaCd, lgAuthBizAreaNm								' 사업장 
	Dim lgInternalAuthYn, lgInternalCd													' 내부부서 
	Dim lgSubInternalAuthYn, lgSubInternalCd											' 내부부서(하위포함)
	Dim lgAuthUsrIDAuthYn, lgAuthUsrID	
    
	Call LoadBasisGlobalInf() 
	Call LoadInfTB19029B("I","*", "NOCOOKIE", "MB")										'☜: Clear Error status
	Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")	
    Call HideStatusWnd													'☜: Hide Processing message
    
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID			= Trim(Request("lgAuthUsrID"))
	    
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""																'☜: Set to space
    lgOpModeCRUD      = Request("txtMode")												'☜: Read Operation Mode (CRUD)

    'Multi SpreadSheet
    lgLngMaxRow       = Request("txtMaxRows")											'☜: Read Operation Mode (CRUD)
    lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)					'☜: "0"(First),"1"(Second),"2"(Third),"3"(...)

	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
	
	Const C_SHEETMAXROWS_D  = 100        
	lgMaxCount = CInt(C_SHEETMAXROWS_D)													'☜: Max fetched data at a time

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)															'☜: Query
             Call SubBizQuery()
    End Select


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                           '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    Call SubBizQueryMulti()
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	Dim iPACG060																	'☆ : 조회용 ComProxy Dll 사용 변수 

	Dim I1_yyyymm  
	Dim l2_module_cd
	Dim l3_biz_area_cd
	Dim EG2_exchange_result_info
	Dim EG1_exchange_result_gl
	
	Dim iLngRow
	Dim iStrCurrency
	
	On Error Resume Next 
	
	Const A65_EG2_exchange_result_info_module_cd = 0
	Const A65_EG2_exchange_result_info_module_nm = 1
    Const A65_EG2_exchange_result_info_ref_no = 2
    Const A65_EG2_exchange_result_info_tran_date = 3
	Const A65_EG2_exchange_result_info_acct_no = 4
	Const A65_EG2_exchange_result_info_acct_nm = 5
	Const A65_EG2_exchange_result_info_doc_cur = 6
	Const A65_EG2_exchange_result_info_xch_rate = 7
	Const A65_EG2_exchange_result_info_item_amt = 8
	Const A65_EG2_exchange_result_info_item_loc_amt = 9
	Const A65_EG2_exchange_result_info_eval_xch_rate = 10
	Const A65_EG2_exchange_result_info_eval_loc_amt = 11
	Const A65_EG2_exchange_result_info_eval_loss_amt = 12
	Const A65_EG2_exchange_result_info_eval_profit_amt = 13
	Const A65_EG2_exchange_result_info_eval_Conf_flg = 14

    Const A65_EG1_exchange_result_temp_Gl_no = 0
    Const A65_EG1_exchange_result_gl_No = 1
    Const A65_EG1_exchange_result_Rev_temp_Gl_no = 2
    Const A65_EG1_exchange_result_Rev_gl_No = 3
    Const A65_EG1_exchange_result_dept_cd = 4
    Const A65_EG1_exchange_result_Gl_dt = 5
    Const A65_EG1_exchange_result_org_change_id = 6
    Const A65_EG1_exchange_result_Dept_nm = 7
    Const A65_EG1_exchange_result_biz_area_nm = 8
    Const A65_EG1_exchange_result_module_cd_nm = 9    

	'#########################################################################################################
	'												2.2. 요청 변수 처리 
	'##########################################################################################################

'	LngMaxRow = Cint(Request("txtMaxRows"))
	I1_yyyymm = Trim(Request("txtYYYYMM"))
	l2_module_cd = Trim(Request("txtModuleCd"))
	l3_biz_area_cd = Trim(Request("txtBizAreaCd"))
	'#########################################################################################################
	'												2.3. 업무 처리 
	'##########################################################################################################

	Set iPACG060 = Server.CreateObject("PACG060.cALkupExchangeResultSvr")

	If CheckSYSTEMError(Err,True) = True Then
		Response.End 
	End If

	Call iPACG060.A_LIST_EXCHANGE_RESULT_SVR (gStrGlobalCollection, I1_yyyymm,l2_module_cd,l3_biz_area_cd,EG2_exchange_result_info, EG1_exchange_result_gl)

	If CheckSYSTEMError(Err,True) = True Then
		Set iPACG060 = Nothing		
		Response.End 
	End If
    
	Set iPACG060 = Nothing 

	'#########################################################################################################
	'												2.4. HTML 결과 생성부 
	'##########################################################################################################
	Response.Write "<Script Language=vbscript>										" & vbcr
	Response.Write " With parent.frm1                                               " & vbcr 
	Response.Write "  .txtGLDt.Text		    = """ & UNIDateClientFormat(EG1_exchange_result_gl(A65_EG1_exchange_result_Gl_dt)) & """" & vbcr
	Response.Write "  .txtDeptCd.Value		= """ & ConvSPChars(EG1_exchange_result_gl(A65_EG1_exchange_result_dept_cd)) & """" & vbcr
	Response.Write "  .txtDeptNm.Value		= """ & ConvSPChars(EG1_exchange_result_gl(A65_EG1_exchange_result_Dept_nm)) & """" & vbcr
	Response.Write "  .hOrgChangeId.Value	= """ & ConvSPChars(EG1_exchange_result_gl(A65_EG1_exchange_result_org_change_id)) & """" & vbcr
	Response.Write "  .txtTempGLNo.Value	= """ & ConvSPChars(EG1_exchange_result_gl(A65_EG1_exchange_result_temp_Gl_no)) & """" & vbcr
	Response.Write "  .txtGLNo.Value		= """ & ConvSPChars(EG1_exchange_result_gl(A65_EG1_exchange_result_gl_No)) & """" & vbcr
	Response.Write "  .txtRevTempGLNo.Value	= """ & ConvSPChars(EG1_exchange_result_gl(A65_EG1_exchange_result_Rev_temp_Gl_no)) & """" & vbcr
	Response.Write "  .txtRevGLNo.Value		= """ & ConvSPChars(EG1_exchange_result_gl(A65_EG1_exchange_result_Rev_gl_No)) & """" & vbcr
	Response.Write "  .htxtBizAreaCd.Value	= """ & ConvSPChars(l3_biz_area_cd) & """" & vbcr
	Response.Write "  .txtBizAreaNm.Value	= """ & ConvSPChars(EG1_exchange_result_gl(A65_EG1_exchange_result_biz_area_nm)) & """" & vbcr
	Response.Write "  .txtModuleName.Value	= """ & ConvSPChars(EG1_exchange_result_gl(A65_EG1_exchange_result_module_cd_nm)) & """" & vbcr
	Response.Write " End With														" & vbcr		    
	Response.write "</Script>														" & vbcr  

	iStrData = ""

	For iLngRow = 0 To UBound(EG2_exchange_result_info,1)
		iStrCurrency = Trim(EG2_exchange_result_info(iLngRow, A65_EG2_exchange_result_info_doc_cur))

		iStrData = iStrData & Chr(11) & ConvSPChars(EG2_exchange_result_info(iLngRow, A65_EG2_exchange_result_info_module_cd))
		iStrData = iStrData & Chr(11) & ConvSPChars(EG2_exchange_result_info(iLngRow, A65_EG2_exchange_result_info_module_nm))
		iStrData = iStrData & Chr(11) & ConvSPChars(EG2_exchange_result_info(iLngRow, A65_EG2_exchange_result_info_ref_no))
		iStrData = iStrData & Chr(11) & EG2_exchange_result_info(iLngRow, A65_EG2_exchange_result_info_tran_date)
		iStrData = iStrData & Chr(11) & ConvSPChars(EG2_exchange_result_info(iLngRow, A65_EG2_exchange_result_info_acct_no))
		iStrData = iStrData & Chr(11) & ConvSPChars(EG2_exchange_result_info(iLngRow, A65_EG2_exchange_result_info_acct_nm))
		iStrData = iStrData & Chr(11) & ConvSPChars(EG2_exchange_result_info(iLngRow, A65_EG2_exchange_result_info_doc_cur))
		iStrData = iStrData & Chr(11) & UNINumClientFormat(EG2_exchange_result_info(iLngRow, A65_EG2_exchange_result_info_xch_rate), ggExchRate.DecPoint, 0)
		iStrData = iStrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG2_exchange_result_info(iLngRow, A65_EG2_exchange_result_info_item_amt),iStrCurrency,ggAmtOfMoneyNo, "X" , "X")	
		iStrData = iStrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG2_exchange_result_info(iLngRow, A65_EG2_exchange_result_info_item_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
		iStrData = iStrData & Chr(11) & UNINumClientFormat(EG2_exchange_result_info(iLngRow, A65_EG2_exchange_result_info_eval_xch_rate), ggExchRate.DecPoint, 0)
		iStrData = iStrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG2_exchange_result_info(iLngRow, A65_EG2_exchange_result_info_eval_loc_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
		iStrData = iStrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG2_exchange_result_info(iLngRow, A65_EG2_exchange_result_info_eval_loss_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
		iStrData = iStrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(EG2_exchange_result_info(iLngRow, A65_EG2_exchange_result_info_eval_profit_amt),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
		iStrData = iStrData & Chr(11) & iLngRow + 1 
		iStrData = iStrData & Chr(11) & Chr(12)
	Next                                                         '☜: Release RecordSSet



	Response.Write "<Script Language=vbscript>" & vbcr
	Response.Write " With Parent																																					 " & vbCr
	Response.Write " 	.ggoSpread.Source = .frm1.vspdData																															 " & vbCr
	Response.Write "	.ggoSpread.SSShowData """ & iStrData   & """ ,""F""																											 " & vbCr
	Response.Write  "    Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData," & 1 & "," & iLngRow  & ",.C_DOC_CUR,.C_XCH_RATE, ""D"" ,""Q"",""X"",""X"")						 " & vbCr
	Response.Write  "    Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData," & 1 & "," & iLngRow  & ",.C_DOC_CUR,.C_ITEM_AMT, ""A"" ,""Q"",""X"",""X"")						 " & vbCr
	Response.Write  "    Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData," & 1 & "," & iLngRow  & ",.C_DOC_CUR,.C_EVAL_XCH_RATE, ""D"" ,""Q"",""X"",""X"")					 " & vbCr	
	Response.Write " 	.DbQueryOk																																					 " & vbCr
	Response.Write " End With																																						 " & vbCr
	Response.Write "</Script>																																						 " & vbCr 	
	
End Sub	
%>

